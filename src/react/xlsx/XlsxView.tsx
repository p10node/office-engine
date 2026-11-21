import { useEffect, useLayoutEffect, useMemo, useRef, useState, useSyncExternalStore, type CSSProperties, type MutableRefObject, type ReactNode } from 'react'
import type { BookOp, CellDetail, ClipCell, CellInput, CellOut, FindMatch, FindOptions, FmtPatch, Merge, Range, SheetOp, SheetProgress, TextRun, WorkbookMeta, XlsxClip, XlsxEditResult } from '../../core/xlsx/types'
import { parseTsv, shiftedClip, tsvCell } from '../../core/xlsx/clip'
import { colName, MAX_ROWS } from '../../core/xlsx/workbook'
import { detectSeparator, parseCsv } from '../../core/convert'
import { closeFormula, formulaContext, isFormula, rangeText, type FormulaContext } from '../../core/xlsx/formula/assist'
import { translateFormula } from '../../core/xlsx/formula'
import { argSlot, functionInfo, suggestFunctions, type FnInfo } from '../../core/xlsx/formula/catalog'
import { MAX_COL, MAX_ROW } from '../../core/xlsx/formula/types'
import type { DocumentClient } from '../../core/worker/client'
import { getPeers, publishLocal, subscribePeers, getLocalHistory, subscribeLocalHistory } from '../../collab/hosts'
import { claimFocus, useForceUpdate, useScroll, useSize } from '../common/hooks'
import { useFileOps, useViewer, useSelectAll } from '../common/Slots'
import { useViewCommands, type TextFormat } from '../controller'
import { cellFormat, cellPatch, marksAfter } from '../edit/format'
import { Status } from '../common/Status'
import { ContextMenu, MOD, type MenuItem } from '../common/ContextMenu'
import { Toast } from '../common/Dialog'
import { FindBar } from '../docx/FindBar'
import { XlsxToolbar, type MergeMode } from './XlsxToolbar'
import { Icon } from '../edit/Toolbar'
import { useLocale } from '../common/locale'
import { numberText } from '../../core/number-input'
import { Axis } from './axis'
import { drawingBox, drawingEl, drawingExtent, useSheetDrawings } from './drawings'
import { fitNumber, fontPx, forgetWidths } from './fit'
import { useDocFonts } from '../common/fonts'
import { htmlToTsv } from './html'
import { fontOf, rotatedHeight, textWidth, wrapHeight } from './measure'
import { sheetPrintable } from './print'
import { useCommentsCtl, useCommentsState, type CommentsCtl, type PlacedThread } from '../comments/Comments'
import { COMMENT_KEYS } from '../comments/CommentsPanel'
import type { CommentThread } from '../../core/comments'
import { cssColor, cssPaint } from '../../core/color'
import { cssFont } from '../../core/fonts'
import { dict } from '../../core/dict'
import { formatCount, useStrings } from '../strings'

const BLOCK = 128
const HH = 22
const MAX_SCROLL_PX = 12_000_000

interface CachedBlock {
  rows: Map<number, CellOut[]>
  /** the format (style index) of the rows that have one of their own */
  own: Map<number, number>
  complete: boolean
  rowsAt: number
}

class SheetState {
  rows = 0
  cols = 0
  done = false
  error?: string
  header = false
  rowAxis = new Axis(20)
  colAxis = new Axis(64)
  merges: Merge[] = []
  /** the columns that have a format of their own: [first, last, style index] */
  colStyles: [number, number, number][] = []
  frozen = { rows: 0, cols: 0 }
  gridLines = true
  blocks = new Map<number, CachedBlock>()
  inflight = new Set<number>()
  /** view state kept while another sheet is shown */
  sel?: Selection
  scrollPos?: { x: number; y: number }
  /** a cell to select and scroll to (a comment shown from the panel); `n` changes every time */
  goto?: { r: number; c: number; n: number }
}

/** the columns of a sheet that have a format of their own (`<col style>`) */
function colStylesOf(cols: SheetProgress['colInfo']): [number, number, number][] {
  const out: [number, number, number][] = []
  for (const ci of cols ?? []) {
    const st = Number(ci.x?.style ?? 0)
    if (st > 0) out.push([ci.min, ci.max, st])
  }
  return out
}

/** whether a format shows in a cell that holds nothing: a fill, a border */
function shows(css: unknown): boolean {
  if (!css || typeof css !== 'object') return false
  for (const k in css as object) if (k.startsWith('background') || k.startsWith('border')) return true
  return false
}

/**
 * How many of the `n` frozen rows / columns are shown as such: those that fit in most of the view. A pane frozen
 * past it leaves nothing to scroll (the rest of the sheet could not be reached), and a file may freeze a million
 * rows: every one of them would be read and drawn.
 */
function frozenShown(axis: Axis, n: number, view: number): number {
  if (!(n > 0)) return 0
  const room = view * 0.75
  let k = 0
  while (k < n && k < 4096 && axis.pos(k + 1) <= room) k++
  return k
}

/** the threads on cells of sheet `path` (not resolved, or the current one), by cell key */
function commentCells(cm: CommentsCtl | null, path: string | undefined, active: string | null): Map<number, string> {
  const out = new Map<number, string>()
  if (!cm || !path) return out
  for (const t of cm.threads()) {
    const a = t.root.at
    if (!a || !('sheet' in a) || a.sheet !== path || (t.root.done && t.root.id !== active)) continue
    if (!out.has(a.r * 16384 + a.c)) out.set(a.r * 16384 + a.c, t.root.id)
  }
  return out
}

class XlsxStore {
  sheets = new Map<number, SheetState>()
  /** focus the grid when the next sheet grid mounts */
  focusGrid = false
  /** edit results applied (here or a collaborator's): charts read their cells again */
  edits = 0
  /** edits that undo / redo can still take back (counted here: the worker's history says nothing until asked) */
  private standaloneHistory = { undo: 0, redo: 0 }
  get history() { return getLocalHistory(this.client) ?? this.standaloneHistory }
  set history(history: { undo: number; redo: number }) {
    if (!getLocalHistory(this.client)) this.standaloneHistory = history
  }
  private listeners = new Set<() => void>()
  private editListeners = new Set<() => void>()
  private off: () => void

  constructor(private client: DocumentClient) {
    const offs = [
      subscribeLocalHistory(client, () => {
        this.standaloneHistory = { undo: 0, redo: 0 }
        this.emit()
      }),
      client.on('xlsx:progress', (p) => this.onProgress(p)),
      // edits of collaborators (src/collab)
      client.on('collab:xlsx', (res) => this.applied(res, client.info.xlsx as WorkbookMeta)),
    ]
    this.off = () => offs.forEach((f) => f())
  }

  subscribe(fn: () => void) {
    this.listeners.add(fn)
    return () => {
      this.listeners.delete(fn)
    }
  }

  onEdit(fn: () => void) {
    this.editListeners.add(fn)
    return () => { this.editListeners.delete(fn) }
  }

  edited() {
    for (const fn of this.editListeners) fn()
    this.emit()
  }

  emit() {
    for (const fn of this.listeners) fn()
  }

  sheet(i: number): SheetState {
    let s = this.sheets.get(i)
    if (!s) {
      s = new SheetState()
      this.sheets.set(i, s)
      this.client.call('xlsx.load', i).catch((e) => {
        s!.error = (e as Error).message
        s!.done = true
        this.emit()
      })
    }
    return s
  }

  private onProgress(p: SheetProgress) {
    const s = this.sheet(p.sheet)
    if (p.defaultRowH !== undefined && !s.header) {
      s.header = true
      s.rowAxis.def = p.defaultRowH || 20
      s.colAxis.def = p.defaultColW || 64
      for (const ci of p.colInfo ?? []) s.colAxis.setRange(ci.min, ci.max, ci.w)
      s.colStyles = colStylesOf(p.colInfo)
      s.frozen = p.frozen ?? s.frozen
      s.gridLines = p.gridLines ?? true
    }
    for (const [r, px] of p.rowSizes) s.rowAxis.set(r, px)
    // values computed after the sheet was read: its rows are fetched again, and the charts read their cells again
    if (p.recalc) {
      s.blocks.clear()
      this.edits++
    }
    s.rows = Math.max(s.rows, p.rows)
    s.cols = Math.max(s.cols, p.cols)
    if (p.merges) s.merges = p.merges
    if (p.error) s.error = p.error
    s.done = p.done
    this.emit()
  }

  ensure(sheet: number, r0: number, r1: number) {
    const s = this.sheet(sheet)
    if (!s.header && !s.done) return
    for (let b = Math.floor(r0 / BLOCK); b <= Math.floor(r1 / BLOCK); b++) {
      const cached = s.blocks.get(b)
      if (s.inflight.has(b)) continue
      if (cached && (cached.complete || (cached.rowsAt === s.rows && !s.done))) continue
      if (b * BLOCK > s.rows && !s.done && cached) continue
      s.inflight.add(b)
      const rowsAt = s.rows
      this.client
        .call('xlsx.block', sheet, b * BLOCK, (b + 1) * BLOCK)
        .then((res) => {
          const rows = new Map<number, CellOut[]>()
          const own = new Map<number, number>()
          for (const r of res.rows) {
            rows.set(r.r, r.cells)
            if (r.s) own.set(r.r, r.s)
          }
          s.blocks.set(b, { rows, own, complete: res.complete, rowsAt })
          if (s.blocks.size > 400) {
            const firstKey = s.blocks.keys().next().value as number
            s.blocks.delete(firstKey)
          }
        })
        .catch(() => {})
        .finally(() => {
          s.inflight.delete(b)
          this.emit()
        })
    }
  }

  /** drop cached rows so they are fetched again (after edits) */
  invalidate(sheet: number, r0: number, r1: number) {
    const s = this.sheet(sheet)
    for (let b = Math.floor(r0 / BLOCK); b <= Math.floor(r1 / BLOCK); b++) s.blocks.delete(b)
    s.rows = Math.max(s.rows, r1 + 1)
    this.emit()
  }

  /** applies an edit result: new styles, layout, sheet list + refetch of touched rows */
  applied(res: XlsxEditResult | null, meta: WorkbookMeta) {
    if (!res) return
    this.edits++
    for (const st of res.styles ?? []) {
      meta.styles[st.index] = st.css
      meta.formats[st.index] = st.format
    }
    for (const i of res.recalc ?? []) this.sheets.get(i)?.blocks.clear()
    if (res.sheets) {
      meta.sheets = res.sheets
      this.sheets.clear()
      this.emit()
      return
    }
    const s = this.sheet(res.sheet)
    if (res.layout) {
      const l = res.layout
      s.rowAxis = new Axis(l.defaultRowH || s.rowAxis.def)
      s.colAxis = new Axis(l.defaultColW || s.colAxis.def)
      for (const ci of l.colInfo ?? []) s.colAxis.setRange(ci.min, ci.max, ci.w)
      s.colStyles = colStylesOf(l.colInfo)
      for (const [r, px] of l.rowSizes) s.rowAxis.set(r, px)
      s.merges = l.merges ?? []
      s.frozen = l.frozen ?? s.frozen
      if (l.gridLines !== undefined) s.gridLines = l.gridLines
      s.rows = l.rows
      s.cols = l.cols
      s.blocks.clear()
      this.emit()
      return
    }
    // (a formula filled cells past the last row / column: the grid reaches them)
    if (res.grown) {
      s.rows = Math.max(s.rows, res.grown.rows)
      s.cols = Math.max(s.cols, res.grown.cols)
    }
    // (a range of whole rows reaches the last column of the sheet, not of what there is)
    if (res.range.c1 < 16383) s.cols = Math.max(s.cols, res.range.c1 + 1)
    this.invalidate(res.sheet, res.range.r0, res.range.r1)
  }

  row(sheet: number, r: number): CellOut[] | undefined {
    return this.sheets.get(sheet)?.blocks.get(Math.floor(r / BLOCK))?.rows.get(r)
  }

  /**
   * The format (style index) a place without a cell shows: its row's own, else its column's (0: none). What a
   * whole row / column / sheet was formatted with, without a cell for every place.
   */
  placeStyle(sheet: number, r: number, c: number): number {
    const s = this.sheets.get(sheet)
    if (!s) return 0
    const own = s.blocks.get(Math.floor(r / BLOCK))?.own.get(r)
    if (own) return own
    for (const [min, max, st] of s.colStyles) if (min <= c && c <= max) return st
    return 0
  }

  /** whether the rows r0..r1 are here as the worker has them now (not being read, none missing) */
  loaded(sheet: number, r0: number, r1: number): boolean {
    const s = this.sheets.get(sheet)
    if (!s || !s.done) return false
    for (let b = Math.floor(r0 / BLOCK); b <= Math.floor(r1 / BLOCK); b++) if (s.inflight.has(b) || !s.blocks.get(b)?.complete) return false
    return true
  }

  /**
   * The cells copied or cut last (from any sheet of the workbook): `text` is what went to the clipboard. Pasting
   * that same text pastes them (values, formulas, formats) instead of the text.
   */
  clip: { text: string; data: XlsxClip; cut: boolean } | null = null

  dispose() {
    this.off()
    this.listeners.clear()
  }
}

const stores = new WeakMap<DocumentClient, XlsxStore>()
function storeFor(client: DocumentClient) {
  let s = stores.get(client)
  if (!s) stores.set(client, (s = new XlsxStore(client)))
  return s
}

export interface XlsxViewProps {
  client: DocumentClient
  onStatus?: (s: string) => void
  initialSheet?: number
}

export function XlsxView({ client, onStatus, initialSheet }: XlsxViewProps) {
  const meta = client.info.xlsx as WorkbookMeta
  const visible = meta.sheets.filter((s) => !s.hidden)
  const [active, setActive] = useState(initialSheet ?? visible[0]?.index ?? 0)
  const store = useMemo(() => storeFor(client), [client])
  const [, bump] = useForceUpdate()
  useEffect(() => store.subscribe(bump), [store, bump])
  const state = store.sheet(active)
  const str = useStrings()
  const n = (v: number) => formatCount(str, v)

  const status = state.error
    ? str.common.error(state.error)
    : `${str.xlsx.status(state.rows, state.cols, str.locale)}${state.done ? '' : str.xlsx.statusLoading}${state.rows > MAX_ROWS ? str.xlsx.overLimit(MAX_ROWS, str.locale) : ''}`

  const viewer = useViewer()
  // The workbook owns completion publication even when the grid that began the edit changed sheets.
  // Layout cleanup removes this listener before a replacement document can receive late dirty state.
  useLayoutEffect(() => store.onEdit(() => viewer.setDirty(true)), [store, viewer.setDirty])
  useFileOps({
    printable: () => sheetPrintable(client, cur, store.sheet(cur), meta, str),
    text: (format) => client.call('xlsx.text', cur, format === 'tsv' ? '\t' : ','),
    stats: () => {
      const s = store.sheet(cur)
      return [
        [str.xlsx.sheets, `${n(meta.sheets.length)}${meta.sheets.some((x) => x.hidden) ? str.xlsx.hiddenSheets(meta.sheets.filter((x) => x.hidden).length, str.locale) : ''}`],
        [str.xlsx.currentSheet, meta.sheets[cur]?.name ?? ''],
        [str.xlsx.usedRange, `${str.xlsx.usedRangeValue(s.rows, s.cols, str.locale)}${s.done ? '' : str.xlsx.usedRangeLoading}`],
      ]
    },
    importText: async (file, name) => {
      const sep = /\.(tsv|tab)$/i.test((file as File).name ?? '') ? '\t' : undefined
      // a shared session: the rows reach the others as cell edits
      if (client.interceptor) return importRows(file, name, sep)
      // else the worker streams the file into the new sheet, like a sheet's XML (a million rows is fine)
      const res = await client.call('xlsx.import', cur, file, { name, sep })
      store.applied(res, meta)
      historyCleared()
      viewer.setDirty(true)
      setActive(cur + 1)
    },
  })
  // a sheet added, deleted, moved or renamed clears the worker's undo history
  const historyCleared = () => {
    store.history = { undo: 0, redo: 0 }
    viewer.ctl?.update({ canUndo: false, canRedo: false })
  }
  /** the rows of a text file as cell edits of a new sheet (parsed here, sent through the session) */
  const importRows = async (file: Blob, name: string, sep?: string) => {
    const text = await file.text()
    const rows = parseCsv(text, sep ?? detectSeparator(text))
    const res = await client.call('xlsx.bookOp', { op: 'add', name, after: cur })
    store.applied(res, meta)
    historyCleared()
    viewer.setDirty(true)
    const idx = cur + 1
    setActive(idx)
    // cells can be written once the new sheet is loaded
    await new Promise<void>((resolve, reject) => {
      let off = () => {}
      const check = () => {
        const s = store.sheet(idx)
        if (s.error) reject(new Error(s.error))
        else if (s.done) resolve()
        else return
        off()
      }
      off = store.subscribe(check)
      check()
    })
    const cells: CellInput[] = []
    rows.forEach((row, r) => row.forEach((input, c) => input !== '' && cells.push({ r, c, input })))
    if (cells.length) store.applied(await client.call('xlsx.set', idx, cells), meta)
    bump()
  }
  const [renaming, setRenaming] = useState<number | null>(null)
  /** a message about a sheet action (a name refused or changed) */
  const [note, setNote] = useState<string | null>(null)
  const [tabMenu, setTabMenu] = useState<{ index: number; x: number; y: number } | null>(null)
  const bookOp = async (op: BookOp, goTo?: number) => {
    try {
      const res = await client.call('xlsx.bookOp', op)
      store.applied(res, meta)
      historyCleared()
      viewer.setDirty(true)
      if (goTo !== undefined) setActive(Math.max(0, Math.min(goTo, meta.sheets.length - 1)))
      bump()
    } catch (e) {
      alertError(e)
      setNote((e as Error).message)
    }
  }
  /** renames sheet `index` to what was typed: an empty or taken name is refused with the reason, a name that is cut or changed says so */
  const renameSheet = (index: number, raw: string) => {
    const old = meta.sheets[index]
    if (!old || raw === old.name) return
    const name = cleanSheetName(raw)
    if (!name) return setNote(str.xlsx.sheetNameEmpty)
    if (meta.sheets.some((x) => x.index !== index && x.name.toLowerCase() === name.toLowerCase())) return setNote(str.xlsx.sheetNameTaken(name))
    const said: string[] = []
    if (raw.trim().length > 31) said.push(str.xlsx.sheetNameShortened(name))
    if (/[\\/?*[\]:]/.test(raw)) said.push(str.xlsx.sheetNameInvalidChars)
    if (name !== old.name) void bookOp({ op: 'rename', index, name: raw })
    if (said.length) setNote(said.join(' '))
  }
  const cur = Math.min(active, meta.sheets.length - 1)
  const curRef = useRef(cur)
  curRef.current = cur

  // ---- comments: a thread points at a cell ----
  const cm = useCommentsCtl()
  useEffect(() => {
    if (!cm) return
    let n = 0
    const sheetOf = (t: CommentThread) => {
      const a = t.root.at
      const i = a && 'sheet' in a ? meta.sheets.findIndex((s) => s.path === a.sheet) : -1
      return i >= 0 ? { i, a: a as { sheet: string; r: number; c: number } } : null
    }
    cm.setHost({
      kind: 'xlsx',
      hint: () => cm.t.xlsx.commentHint,
      pending: () => {
        const i = curRef.current
        const sh = meta.sheets[i]
        if (!sh) return null
        const sel = store.sheet(i).sel ?? { ar: 0, ac: 0 }
        // one thread per cell: commenting on a cell that has one answers it
        const had = cm.threads().find((t) => {
          const x = sheetOf(t)
          return x?.i === i && x.a.r === sel.ar && x.a.c === sel.ac
        })
        return { quote: `${sh.name}!${colName(sel.ac)}${sel.ar + 1}`, at: { sheet: sh.path, r: sel.ar, c: sel.ac }, existing: had?.root.id }
      },
      placed: (threads) => {
        const out: (PlacedThread & { k: number[] })[] = []
        for (const t of threads) {
          const x = sheetOf(t)
          if (x) out.push({ ...t, quote: `${meta.sheets[x.i].name}!${colName(x.a.c)}${x.a.r + 1}`, k: [x.i, x.a.r, x.a.c] })
        }
        return out.sort((a, b) => a.k[0] - b.k[0] || a.k[1] - b.k[1] || a.k[2] - b.k[2])
      },
      reveal: (id) => {
        const t = cm.threads().find((x) => x.root.id === id)
        const x = t && sheetOf(t)
        if (!x) return
        const s = store.sheet(x.i)
        s.sel = { ar: x.a.r, ac: x.a.c, fr: x.a.r, fc: x.a.c }
        s.goto = { r: x.a.r, c: x.a.c, n: ++n }
        setActive(x.i)
        store.emit()
      },
    })
    // sheet names / order show in the panel
    const off = store.subscribe(() => cm.touch())
    return () => {
      off()
      cm.setHost(null)
    }
  }, [cm, store, meta])

  const [edit, setEdit] = useState<EditState | null>(null)
  /** set while switching sheets during a formula edit: blur must not commit */
  const hold = useRef(false)
  const keepEditing = !!edit && isFormula(edit.value) && canPoint(edit, formulaContext(edit.value, edit.caret))

  // ---- controller: the sheets as the outline, sheet switching, undo ----
  const { ctl, headless } = viewer
  const sheetItems = meta.sheets.filter((s) => !s.hidden).map((s) => ({ id: String(s.index), title: s.name, level: 0, current: s.index === cur }))
  useEffect(() => {
    const h = store.history
    ctl?.update({ outline: { kind: 'sheets', items: sheetItems }, canUndo: viewer.editable && h.undo > 0, canRedo: viewer.editable && h.redo > 0, pointing: keepEditing })
  })
  const gotoSheet = (id: string) => {
    const i = Number(id)
    if (!meta.sheets[i] || meta.sheets[i].hidden) return
    // a formula being typed stays open on the other sheet, for cross-sheet references, as with the built-in tabs
    if (keepEditing && i !== cur) hold.current = true
    setActive(i)
  }

  return (
    <div className="de-xlsx">
      <Status text={status} onStatus={onStatus} />
      <SheetGrid
        key={`${cur}:${meta.sheets[cur]?.path}`}
        store={store}
        sheet={cur}
        meta={meta}
        client={client}
        edit={edit}
        setEdit={setEdit}
        hold={hold}
        onSheet={setActive}
        gotoSheet={gotoSheet}
      />
      <div className="de-tabs" role="tablist" hidden={headless}>
        {viewer.editable && (
          <button type="button" className="de-tab de-tab-add" title={str.xlsx.addSheet} onClick={() => void bookOp({ op: 'add', after: cur }, cur + 1)}>
            <Icon name="sheetAdd" size={16} />
          </button>
        )}
        {meta.sheets.filter((s) => !s.hidden).map((s) =>
          renaming === s.index ? (
            <input
              key={s.index}
              className="de-tab-input"
              autoFocus
              defaultValue={s.name}
              onBlur={(e) => {
                setRenaming(null)
                renameSheet(s.index, e.target.value)
              }}
              onKeyDown={(e) => {
                if (e.key === 'Enter') (e.target as HTMLInputElement).blur()
                if (e.key === 'Escape') setRenaming(null)
              }}
            />
          ) : (
            <button
              key={s.index}
              type="button"
              role="tab"
              className="de-tab"
              aria-selected={s.index === cur}
              onMouseDown={(e) => {
                // formula being typed: clicking another sheet keeps editing (cross-sheet references)
                if (!keepEditing) return
                e.preventDefault()
                if (s.index !== cur) hold.current = true
              }}
              onClick={() => setActive(s.index)}
              onDoubleClick={() => viewer.editable && setRenaming(s.index)}
              onContextMenu={(e) => {
                if (!viewer.editable) return
                e.preventDefault()
                setTabMenu({ index: s.index, x: e.clientX, y: e.clientY })
              }}
            >
              {s.name}
            </button>
          ),
        )}
      </div>
      {tabMenu && (
        <ContextMenu
          x={tabMenu.x}
          y={tabMenu.y}
          onClose={() => setTabMenu(null)}
          items={[
            { label: str.xlsx.renameSheet, run: () => setRenaming(tabMenu.index) },
            { label: str.xlsx.duplicateSheet, run: () => void bookOp({ op: 'duplicate', index: tabMenu.index }, tabMenu.index + 1) },
            { label: str.xlsx.moveLeft, disabled: tabMenu.index === 0, run: () => void bookOp({ op: 'move', from: tabMenu.index, to: tabMenu.index - 1 }, tabMenu.index - 1) },
            { label: str.xlsx.moveRight, disabled: tabMenu.index >= meta.sheets.length - 1, run: () => void bookOp({ op: 'move', from: tabMenu.index, to: tabMenu.index + 1 }, tabMenu.index + 1) },
            { label: str.xlsx.deleteSheet, disabled: meta.sheets.length <= 1, run: () => void bookOp({ op: 'delete', index: tabMenu.index }, Math.max(0, tabMenu.index - 1)) },
          ]}
        />
      )}
      {note && <Toast text={note} onDone={() => setNote(null)} />}
    </div>
  )
}

function alertError(e: unknown) {
  console.error('[office-engine]', e)
}

/** the name a sheet gets for `raw`: what the worker makes of it (31 characters, none of \ / ? * [ ] :, no apostrophe at the ends) */
const cleanSheetName = (raw: string) => raw.trim().slice(0, 31).replace(/[\\/?*[\]:]/g, '_').replace(/^'|'$/g, '_').trim()

export { ContextMenu, type MenuItem } from '../common/ContextMenu'

interface GridProps {
  store: XlsxStore
  sheet: number
  meta: WorkbookMeta
  client: DocumentClient
  edit: EditState | null
  setEdit: React.Dispatch<React.SetStateAction<EditState | null>>
  hold: MutableRefObject<boolean>
  onSheet: (i: number) => void
  gotoSheet: (id: string) => void
}

/** text rotation (OOXML degrees, + = counter-clockwise, 255 = stacked) */
function rotated(rot: number): CSSProperties {
  if (rot === 255) return { writingMode: 'vertical-rl', textOrientation: 'upright', whiteSpace: 'nowrap' }
  if (rot === 90) return { writingMode: 'vertical-rl', transform: 'rotate(180deg)', whiteSpace: 'nowrap' }
  if (rot === -90) return { writingMode: 'vertical-rl', whiteSpace: 'nowrap' }
  return { display: 'inline-block', transform: `rotate(${-rot}deg)`, whiteSpace: 'nowrap' }
}

/**
 * Tilted text, as Excel draws it: from the bottom left corner of its cell upwards (from the top left one downwards),
 * out over the cells to the right. `line` is the height of a line of it.
 */
function tilted(rot: number, line: number): CSSProperties {
  const up = rot > 0
  return {
    display: 'inline-block', flex: 'none', whiteSpace: 'nowrap', transform: `rotate(${-rot}deg)`, transformOrigin: up ? 'left bottom' : 'left top',
    alignSelf: up ? 'flex-end' : 'flex-start', marginLeft: Math.ceil(line * Math.sin((Math.abs(rot) * Math.PI) / 180)),
  }
}

/** the font of a part of a rich text: what its run properties say, and nothing of the cell's font where they are silent */
function runCss(r: TextRun): CSSProperties | undefined {
  if (!r.own) return undefined
  const deco = [r.u ? 'underline' : '', r.strike ? 'line-through' : ''].filter(Boolean).join(' ')
  const css: CSSProperties = { fontWeight: r.b ? 700 : 400, fontStyle: r.i ? 'italic' : 'normal' }
  if (deco) css.textDecoration = deco
  const color = r.color && cssColor(r.color)
  if (color) css.color = color
  if (r.sz) css.fontSize = `${r.va ? r.sz * 0.7 : r.sz}pt`
  else if (r.va) css.fontSize = '0.7em'
  if (r.font) css.fontFamily = cssFont(r.font)
  if (r.va) css.verticalAlign = r.va === 'superscript' ? 'super' : 'sub'
  return css
}

/** the tallest a row grows to show its text (Excel's limit: 409 pt) */
const MAX_FIT_ROW = 545

const JUSTIFY: Record<string, string> = dict({ center: 'center', right: 'flex-end', left: 'flex-start', justify: 'flex-start' })

interface Selection {
  ar: number
  ac: number
  fr: number
  fc: number
  /** selected as whole columns / rows / the sheet (a header, Ctrl+Space…), not as the cells there are: a format goes to the columns / rows themselves */
  whole?: 'col' | 'row' | 'all'
}

interface EditState {
  /** sheet of the edited cell (may differ from the shown sheet while pointing) */
  sheet: number
  r: number
  c: number
  value: string
  initial: string
  /** editing in the formula bar */
  bar: boolean
  /** 'enter': started by typing (arrows point / commit), 'edit': F2, double click, formula bar */
  mode: 'enter' | 'edit'
  caret: number
  caretEnd: number
  /** bump to push caret/focus into the input */
  sync: number
  /** reference inserted by pointing: span in `value` + pointed block */
  point?: { s: number; e: number; sheet: number; ar: number; ac: number; fr: number; fc: number }
  /** autocomplete dismissed (Escape) for the word starting here */
  noSuggest?: number
}

/** pointing a reference is possible at the caret */
function canPoint(e: EditState, fc: FormulaContext | null): boolean {
  if (!fc) return false
  if (e.point && e.point.e === e.caret && e.caret === e.caretEnd) return true
  if (e.caret !== e.caretEnd) return fc.refs.some((r) => r.s === e.caret && r.e === e.caretEnd)
  return fc.pointable
}

const REF_COLORS = ['#1a73e8', '#e8710a', '#9334e6', '#188038', '#d93025', '#12a4b8', '#e52592', '#b08800']
const refKey = (t: string) => t.replace(/\$/g, '').toUpperCase()

const ARROWS: Record<string, [number, number]> = dict({ ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] })

export const rectOf = (s: Selection): Range => ({ r0: Math.min(s.ar, s.fr), c0: Math.min(s.ac, s.fc), r1: Math.max(s.ar, s.fr), c1: Math.max(s.ac, s.fc) })

/** grows a range until it holds every merged cell it touches (like Excel / Sheets selections) */
export function withMerges(r: Range, merges: Merge[]): Range {
  let out = r
  for (let changed = true; changed; ) {
    changed = false
    for (const m of merges) {
      if (m.r1 < out.r0 || m.r0 > out.r1 || m.c1 < out.c0 || m.c0 > out.c1) continue
      if (m.r0 >= out.r0 && m.r1 <= out.r1 && m.c0 >= out.c0 && m.c1 <= out.c1) continue
      out = { r0: Math.min(out.r0, m.r0), c0: Math.min(out.c0, m.c0), r1: Math.max(out.r1, m.r1), c1: Math.max(out.c1, m.c1) }
      changed = true
    }
  }
  return out
}

/** the hidden input field of a grid at rest: inside the view, so that nothing scrolls to it */
const SINK_STYLE: CSSProperties = { left: 4, top: 4, width: 1, height: 1 }

function SheetGrid({ store, sheet, meta, client, edit, setEdit, hold, onSheet, gotoSheet }: GridProps) {
  const viewer = useViewer()
  const str = useStrings()
  const editable = viewer.editable
  // numbers typed are read in it, and a number edited is shown in it (1,5 in de-DE: what is shown is read back the same)
  const locale = useLocale()
  const typed = locale ? { locale } : undefined
  const s = store.sheet(sheet)
  const [scrollRef, size] = useSize<HTMLDivElement>()
  const scroll = useScroll(scrollRef)
  const [sel, setSelState] = useState<Selection>(() => s.sel ?? { ar: 0, ac: 0, fr: 0, fc: 0 })
  s.sel = sel
  const [detailState, setDetail] = useState<{ r: number; c: number; d?: CellDetail } | null>(null)
  const detail = detailState && detailState.r === sel.ar && detailState.c === sel.ac ? detailState.d : undefined
  const [detailTick, setDetailTick] = useState(0)
  const dragging = useRef(false)
  const editRef = useRef<HTMLTextAreaElement>(null)
  const hlRef = useRef<HTMLDivElement>(null)
  const fxRef = useRef<HTMLTextAreaElement>(null)
  const [sugIdx, setSugIdx] = useState(0)
  const [ctx, setCtx] = useState<{ x: number; y: number } | null>(null)
  /** paint format: style index copied from the active cell, applied to the next selection */
  const [painter, setPainter] = useState<number | null>(null)
  const afterDrag = useRef<(() => void) | null>(null)
  /** the fill handle being dragged: the cells it started from, and the cell it is over */
  const fillRef = useRef<{ src: Range; to?: { r: number; c: number } } | null>(null)
  const [fillTo, setFillTo] = useState<{ r: number; c: number } | null>(null)
  /** the column a row of entries began in: Enter after Tab goes back to it */
  const tabStart = useRef<number | null>(null)
  /** the last edit sent to the worker (a save waits for it) */
  const lastRun = useRef<Promise<void>>(Promise.resolve())
  const [find, setFind] = useState<{ open: boolean; q: string; idx: number; replace: boolean; tick: number }>({ open: false, q: '', idx: 0, replace: false, tick: 0 })
  const [findOpts, setFindOpts] = useState<FindOptions>({})
  const [matches, setMatches] = useState<FindMatch[]>([])
  const [toast, setToast] = useState<string | null>(null)
  const cm = useCommentsCtl()
  const cmState = useCommentsState(cm)
  const cmCells = useMemo(() => commentCells(cm, meta.sheets[sheet]?.path, cmState.active), [cm, cmState, meta.sheets, sheet])
  const cmAt = (r: number, c: number) => {
    const m = mergeAt(r, c)
    return cmCells.get(m ? m.r0 * 16384 + m.c0 : r * 16384 + c)
  }
  const [resizing, setResizing] = useState<{ axis: 'col' | 'row'; i: number; start: number; size: number } | null>(null)

  const { rowAxis, colAxis } = s
  const drawings = useSheetDrawings(client, sheet, store.edits)
  // the sheet reaches as far as its drawings, past its last filled cell
  const drawn = drawingExtent(drawings, colAxis, rowAxis)
  const HW = Math.max(40, 12 + 7 * String(Math.max(s.rows, 1000)).length)
  const viewW = Math.max(0, size.w - HW)
  const viewH = Math.max(0, size.h - HH)
  // the rows / columns frozen, as far as they are shown: see `frozenShown`
  const frozen = { rows: frozenShown(rowAxis, s.frozen.rows, viewH), cols: frozenShown(colAxis, s.frozen.cols, viewW) }
  rowAxis.count = Math.max(s.rows + (s.done ? 20 : 1), Math.ceil(viewH / rowAxis.def) + 1, frozen.rows + 1, sel.fr + 2, drawn.rows + 1)
  colAxis.count = Math.max(s.cols + 3, 26, Math.ceil(viewW / Math.max(1, colAxis.def)) + 1, frozen.cols + 1, sel.fc + 2, drawn.cols + 1)
  const frozenW = colAxis.pos(frozen.cols)
  const frozenH = rowAxis.pos(frozen.rows)
  const totalW = colAxis.total()
  const totalH = rowAxis.total()

  // virtual scroll with clamped spacer height (browsers cap element size)
  const spacerW = HW + totalW
  const spacerH = Math.min(HH + totalH, MAX_SCROLL_PX)
  const maxLogicalY = Math.max(0, HH + totalH - size.h)
  const maxScrollY = Math.max(0, spacerH - size.h)
  const scrollY = maxScrollY > 0 ? Math.min(maxLogicalY, (scroll.y / maxScrollY) * maxLogicalY) : 0
  const scrollX = scroll.x

  const mainOx = frozenW + scrollX
  const mainOy = frozenH + scrollY
  const mainW = Math.max(0, viewW - frozenW)
  const mainH = Math.max(0, viewH - frozenH)
  const r0 = rowAxis.index(mainOy)
  const r1 = Math.min(rowAxis.count - 1, rowAxis.index(mainOy + mainH))
  const c0 = colAxis.index(mainOx)
  const c1 = Math.min(colAxis.count - 1, colAxis.index(mainOx + mainW))

  // merges touching the visible area (+ frozen panes)
  const visMerges = useMemo(() => {
    const inRange = (m: Merge) =>
      (m.r1 >= r0 && m.r0 <= r1) || m.r0 < frozen.rows ? (m.c1 >= c0 && m.c0 <= c1) || m.c0 < frozen.cols : false
    return s.merges.filter(inRange)
  }, [s.merges, r0, r1, c0, c1, frozen.rows, frozen.cols])

  useEffect(() => {
    if (frozen.rows) store.ensure(sheet, 0, frozen.rows - 1)
    store.ensure(sheet, Math.max(0, r0 - 20), r1 + 40)
    for (const m of visMerges) if (m.r0 < r0) store.ensure(sheet, m.r0, m.r0)
  })

  useEffect(() => {
    let live = true
    const { ar, ac } = sel
    client.call('xlsx.cell', sheet, ar, ac).then((d) => live && setDetail({ r: ar, c: ac, d }), () => {})
    return () => {
      live = false
    }
  }, [client, sheet, sel.ar, sel.ac, s.done, detailTick])

  const styles = meta.styles
  const basePx = fontPx(meta.baseFont.fontSize, 14.67)
  // the host's fonts the workbook is set in are loaded: numbers fitted to their columns before one arrived (or before
  // the page's own web font did) are fitted again
  const fonts = useDocFonts()
  // eslint-disable-next-line react-hooks/exhaustive-deps
  const stacks = useMemo(() => new Set([meta.baseFont, ...styles].map((x) => x?.fontFamily).filter((f): f is string => typeof f === 'string')), [styles.length])
  const fontsKey = `${fonts.ready(stacks)}|${fonts.gen}`
  const lastFonts = useRef(fontsKey)
  if (lastFonts.current !== fontsKey) {
    lastFonts.current = fontsKey
    forgetWidths()
  }
  const mergeAt = (r: number, c: number) => s.merges.find((m) => r >= m.r0 && r <= m.r1 && c >= m.c0 && c <= m.c1)

  // A row that says no height shows its turned text whole: Excel fits such rows when it opens the file. In the view
  // only: nothing is written, and the file still leaves the height to who opens it.
  useLayoutEffect(() => {
    let grew = false
    for (const [a, b] of [[0, frozen.rows - 1], [r0, r1]]) {
      for (let r = a; r <= b; r++) {
        if (rowAxis.size(r) !== rowAxis.def) continue
        let h = 0
        for (const [c, text, si] of store.row(sheet, r) ?? []) {
          const css = styles[si] as Record<string, string | number> | undefined
          if (!text || !css?.textRotate) continue
          // (a cell merged over several rows has them all to show its text)
          const m = s.merges.length ? mergeAt(r, c) : undefined
          if (!m || m.r0 === m.r1) h = Math.max(h, Math.min(MAX_FIT_ROW, rotatedHeight(text, css, meta.baseFont as Record<string, string>, basePx, Number(css.textRotate))))
        }
        if (h > rowAxis.def) {
          rowAxis.set(r, h)
          grew = true
        }
      }
    }
    if (grew) store.emit()
  })
  const rect = withMerges(rectOf(sel), s.merges)
  /** one cell, or exactly one merged cell */
  const oneCell = (rect.r0 === rect.r1 && rect.c0 === rect.c1) || s.merges.some((m) => m.r0 === rect.r0 && m.c0 === rect.c0 && m.r1 === rect.r1 && m.c1 === rect.c1)
  const selRef = useRef(rect)
  selRef.current = rect
  // collaboration: share the selection, show the collaborators' ones
  useEffect(() => publishLocal(client, { sheet, range: rect, ar: sel.ar, ac: sel.ac }), [client, sheet, rect.r0, rect.c0, rect.r1, rect.c1, sel.ar, sel.ac]) // eslint-disable-line react-hooks/exhaustive-deps
  const peers = useSyncExternalStore(
    (fn) => subscribePeers(client, fn),
    () => getPeers(client),
  )
  const sheetPath = meta.sheets[sheet]?.path
  const peerSels = peers.flatMap((p) => {
    const x = p.presence as { k?: string; sheet?: string; range?: Range; ar?: number; ac?: number } | undefined
    return x?.k === 'xlsx' && x.sheet === sheetPath && x.range ? [{ ...p, range: x.range, ar: x.ar ?? x.range.r0, ac: x.ac ?? x.range.c0 }] : []
  })

  useLayoutEffect(() => {
    hold.current = false
    const el = scrollRef.current
    if (el && s.scrollPos) {
      el.scrollLeft = s.scrollPos.x
      el.scrollTop = s.scrollPos.y
    }
    // the grid's keyboard input is a hidden text field (see `de-cell-sink`): a view that cannot edit has none
    const target = (editable ? editRef.current : null) ?? el ?? null
    if (store.focusGrid) {
      store.focusGrid = false
      target?.focus({ preventScroll: true })
    } else claimFocus(target) // keyboard first: a focused host page control keeps the focus
    return () => {
      if (el?.isConnected) s.scrollPos = { x: el.scrollLeft, y: el.scrollTop }
    }
  }, []) // eslint-disable-line react-hooks/exhaustive-deps

  const setSel = (next: Selection) => {
    const m = mergeAt(next.ar, next.ac)
    if (m) next = { ...next, ar: m.r0, ac: m.c0 }
    setSelState(next)
  }

  // ---------------------------------------------------------------- rendering helpers

  function region(key: string, rr0: number, rr1: number, cc0: number, cc1: number, ox: number, oy: number, left: number, top: number, w: number, h: number) {
    if (w <= 0 || h <= 0 || rr1 < rr0 || cc1 < cc0) return null
    const out: ReactNode[] = []
    if (s.gridLines) {
      for (let r = rr0; r <= rr1; r++) out.push(<div key={'h' + r} className="de-gl-h" style={{ top: rowAxis.pos(r + 1) - oy - 1 }} />)
      for (let c = cc0; c <= cc1; c++) out.push(<div key={'v' + c} className="de-gl-v" style={{ left: colAxis.pos(c + 1) - ox - 1 }} />)
    }
    const covered = new Set<number>()
    const anchors: Merge[] = []
    for (const m of visMerges) {
      if (m.r1 < rr0 || m.r0 > rr1 || m.c1 < cc0 || m.c0 > cc1) continue
      anchors.push(m)
      for (let r = Math.max(m.r0, rr0); r <= Math.min(m.r1, rr1); r++)
        for (let c = Math.max(m.c0, cc0); c <= Math.min(m.c1, cc1); c++) covered.add(r * 16384 + c)
    }
    const cellEl = (r: number, cell: CellOut, wOverride?: number, hOverride?: number, filled?: Set<number>, merged = false) => {
      const [c, raw, si, kind, color, runs] = cell
      const cw = wOverride ?? colAxis.size(c)
      const ch = hOverride ?? rowAxis.size(r)
      if (!cw || !ch) return null
      const css = (styles[si] ?? {}) as Record<string, string | number>
      let text = raw
      if (kind === 1 && raw) text = fitNumber(raw, cw - 6, fontPx(css.fontSize, basePx), fontOf(css, meta.baseFont as Record<string, string>, basePx))
      const { textAlign, textRotate, alignItems, whiteSpace, shrinkToFit, ...rest } = css
      // the fill is a color or a gradient of colors, never more: `background` would load a url() in it
      if (rest.background !== undefined) rest.background = cssPaint(rest.background) as string
      const rot = Number(textRotate ?? 0)
      const justify = (textAlign && JUSTIFY[textAlign as string]) || (kind === 1 ? 'flex-end' : kind === 2 ? 'center' : 'flex-start')
      let ov = 0
      if (text && !rot && !whiteSpace && !shrinkToFit && justify === 'flex-start' && !merged && filled) {
        let next = c + 1
        while (next <= cc1 + 8 && !filled.has(next) && !covered.has(r * 16384 + next)) next++
        ov = colAxis.pos(next) - colAxis.pos(c)
      }
      const style: CSSProperties = {
        ...(rest as CSSProperties),
        left: colAxis.pos(c) - ox,
        top: rowAxis.pos(r) - oy,
        width: cw,
        height: ch,
        justifyContent: justify,
        alignItems: (alignItems as string) ?? 'flex-end',
        textAlign: (textAlign as CSSProperties['textAlign']) ?? undefined,
      }
      if (color) style.color = cssColor(color)
      if (whiteSpace) style.whiteSpace = whiteSpace as CSSProperties['whiteSpace']
      // a merged cell hides the gridlines inside it (its own right / bottom edge lines stay visible)
      if (merged && !rest.background) style.background = 'linear-gradient(#fff, #fff) 0 0 / calc(100% - 1px) calc(100% - 1px) no-repeat'
      const overflow = ov > cw
      // tilted text goes out over the cells to the right, not over the rows above and below
      const tilt = !!text && rot !== 0 && Math.abs(rot) < 90 && !whiteSpace && justify === 'flex-start'
      if (tilt) Object.assign(style, { overflow: 'visible', clipPath: 'inset(0 -4000px 0 0)' })
      // shrink to fit: the text scaled down until it fits the cell
      let shrink: CSSProperties | undefined
      if (shrinkToFit && text && !rot && !whiteSpace) {
        const w = textWidth(text, css, meta.baseFont as Record<string, string>, basePx)
        if (w > cw - 6) shrink = { transform: `scale(${Math.max(0.2, (cw - 6) / w)})`, transformOrigin: justify === 'flex-end' ? 'right center' : justify === 'center' ? 'center' : 'left center', whiteSpace: 'nowrap', flex: 'none' }
      }
      return (
        <div key={`${r}:${c}`} className={overflow ? 'de-cell de-ov' : 'de-cell'} style={style}>
          {text && <span style={tilt ? tilted(rot, fontPx(css.fontSize, basePx) * 1.2) : rot ? rotated(rot) : overflow ? { maxWidth: ov - 6, overflow: 'hidden', flex: 'none' } : shrink}>{runs ? runs.map((x, i) => <span key={i} style={runCss(x)}>{x.t}</span>) : text}</span>}
        </div>
      )
    }
    for (const m of anchors) {
      const row = store.row(sheet, m.r0)
      const cell = row?.find((x) => x[0] === m.c0) ?? ([m.c0, '', 0, 0] as CellOut)
      out.push(cellEl(m.r0, cell, colAxis.pos(m.c1 + 1) - colAxis.pos(m.c0), rowAxis.pos(m.r1 + 1) - rowAxis.pos(m.r0), undefined, true))
    }
    // the places without a cell in rows / columns that have a format of their own show it (a fill, borders)
    const block = (r: number) => s.blocks.get(Math.floor(r / BLOCK))
    if (s.colStyles.length || [...new Set([block(rr0), block(rr1)])].some((b) => b?.own.size)) {
      for (let r = rr0; r <= rr1; r++) {
        const own = block(r)?.own.get(r)
        if (!own && !s.colStyles.length) continue
        const row = store.row(sheet, r)
        const has = row?.length ? new Set(row.map((x) => x[0])) : undefined
        for (let c = cc0; c <= cc1; c++) {
          if (has?.has(c) || covered.has(r * 16384 + c)) continue
          const st = own || store.placeStyle(sheet, r, c)
          if (st && shows(styles[st])) out.push(cellEl(r, [c, '', st, 0]))
        }
      }
    }
    for (let r = rr0; r <= rr1; r++) {
      const row = store.row(sheet, r)
      if (!row) continue
      let filled: Set<number> | undefined
      for (const cell of row) {
        const c = cell[0]
        if (c < cc0 - 8 || c > cc1) continue
        if (covered.has(r * 16384 + c)) continue
        if (c < cc0 && (!cell[1] || cell[3] !== 0)) continue
        if (anchors.some((m) => m.r0 === r && m.c0 === c)) continue
        if (!filled && cell[3] === 0 && cell[1]) filled = new Set(row.filter((x) => x[1] !== '').map((x) => x[0]))
        out.push(cellEl(r, cell, undefined, undefined, filled))
      }
    }
    // pictures and charts over the cells, where they reach into this pane
    const px0 = colAxis.pos(cc0)
    const px1 = colAxis.pos(cc1 + 1)
    const py0 = rowAxis.pos(rr0)
    const py1 = rowAxis.pos(rr1 + 1)
    drawings.forEach((d, i) => {
      const b = drawingBox(d, colAxis, rowAxis)
      if (!b || b.x > px1 || b.x + b.w < px0 || b.y > py1 || b.y + b.h < py0) return
      out.push(drawingEl('dr' + i, d, { left: b.x - ox, top: b.y - oy, width: b.w, height: b.h }, str))
    })
    // the fill handle: the corner of the selection, dragged to copy it down or to the right
    if (editable && !edit && !fillRef.current && rect.r1 >= rr0 && rect.r1 <= rr1 && rect.c1 >= cc0 && rect.c1 <= cc1 && rect.r1 - rect.r0 < 5000 && rect.c1 - rect.c0 < 200) {
      out.push(<div key="fillh" data-fill className="de-fill-handle" style={{ left: colAxis.pos(rect.c1 + 1) - ox - 4, top: rowAxis.pos(rect.r1 + 1) - oy - 4 }} />)
    }
    if (fillTo) {
      const f = fillExtent(fillTo)
      if (f) {
        out.push(
          <div key="fillp" className="de-fill-preview" style={{ left: colAxis.pos(f.c0) - ox - 1, top: rowAxis.pos(f.r0) - oy - 1, width: colAxis.pos(f.c1 + 1) - colAxis.pos(f.c0) + 1, height: rowAxis.pos(f.r1 + 1) - rowAxis.pos(f.r0) + 1 }} />,
        )
      }
    }
    // selection range + active cell
    if (rect.r1 >= rr0 && rect.r0 <= rr1 && rect.c1 >= cc0 && rect.c0 <= cc1 && !oneCell) {
      out.push(
        <div
          key="range"
          className="de-range"
          style={{
            left: colAxis.pos(rect.c0) - ox,
            top: rowAxis.pos(rect.r0) - oy,
            width: colAxis.pos(rect.c1 + 1) - colAxis.pos(rect.c0),
            height: rowAxis.pos(rect.r1 + 1) - rowAxis.pos(rect.r0),
          }}
        />,
      )
    }
    if (sel.ar >= rr0 - 1 && sel.ar <= rr1 && sel.ac >= cc0 - 1 && sel.ac <= cc1) {
      const m = mergeAt(sel.ar, sel.ac)
      const sw = m ? colAxis.pos(m.c1 + 1) - colAxis.pos(m.c0) : colAxis.size(sel.ac)
      const sh = m ? rowAxis.pos(m.r1 + 1) - rowAxis.pos(m.r0) : rowAxis.size(sel.ar)
      out.push(<div key="sel" className="de-sel" style={{ left: colAxis.pos(sel.ac) - ox - 1, top: rowAxis.pos(sel.ar) - oy - 1, width: sw + 1, height: sh + 1 }} />)
    }
    // cells with comments: a mark in the top right corner
    for (const [k, id] of cmCells) {
      const r = Math.floor(k / 16384)
      const c = k % 16384
      if (r < rr0 || r > rr1 || c < cc0 || c > cc1) continue
      const m = mergeAt(r, c)
      const right = colAxis.pos((m ? m.c1 : c) + 1)
      out.push(<div key={'cm' + id} className={'de-cm-mark' + (cmState.active === id ? ' on' : '')} style={{ left: right - ox - 8, top: rowAxis.pos(r) - oy }} />)
    }
    // collaborators' selections
    for (const p of peerSels) {
      const b = p.range
      if (b.r1 < rr0 || b.r0 > rr1 || b.c1 < cc0 || b.c0 > cc1) continue
      const x0 = colAxis.pos(b.c0)
      const y0 = rowAxis.pos(b.r0)
      const x1 = colAxis.pos(b.c1 + 1)
      const y1 = rowAxis.pos(b.r1 + 1)
      // a peer's color is what the peer chose: used only when it is a color
      const pc = cssColor(p.color)
      out.push(
        <div key={'peer' + p.site} className="de-peer-range" style={{ left: x0 - ox - 1, top: y0 - oy - 1, width: x1 - x0 + 1, height: y1 - y0 + 1, borderColor: pc, background: pc && pc + '12' }}>
          {p.ar >= rr0 && p.ar <= rr1 && p.ac >= cc0 && p.ac <= cc1 && (
            <span className="de-peer-tag" style={{ background: pc }}>
              {p.name}
            </span>
          )}
        </div>,
      )
    }
    // find matches
    findBoxes.forEach((m, i) => {
      if (m.r < rr0 || m.r > rr1 || m.c < cc0 || m.c > cc1) return
      out.push(
        <div
          key={'fm' + i}
          className={i === find.idx ? 'de-findcell on' : 'de-findcell'}
          style={{ left: colAxis.pos(m.c) - ox, top: rowAxis.pos(m.r) - oy, width: colAxis.size(m.c), height: rowAxis.size(m.r) }}
        />,
      )
    })
    // references of the formula being edited
    refBoxes.forEach((b, i) => {
      if (b.r1 < rr0 || b.r0 > rr1 || b.c1 < cc0 || b.c0 > cc1) return
      const y0 = rowAxis.pos(Math.max(b.r0, rr0 - 1))
      const y1 = rowAxis.pos(Math.min(b.r1, rr1 + 1) + 1)
      const x0 = colAxis.pos(Math.max(b.c0, cc0 - 1))
      const x1 = colAxis.pos(Math.min(b.c1, cc1 + 1) + 1)
      out.push(
        <div
          key={'fr' + i}
          className={b.point ? 'de-fref de-fref-pt' : 'de-fref'}
          style={{ left: x0 - ox - 1, top: y0 - oy - 1, width: x1 - x0 + 1, height: y1 - y0 + 1, borderColor: b.color, background: b.color + '14' }}
        />,
      )
    })
    return (
      <div key={key} className="de-region" style={{ left, top, width: w, height: h }}>
        {out}
      </div>
    )
  }

  function colHeaders(key: string, cc0: number, cc1: number, ox: number, left: number, w: number) {
    if (w <= 0) return null
    const out: ReactNode[] = []
    for (let c = cc0; c <= cc1; c++) {
      const cw = colAxis.size(c)
      if (!cw) continue
      out.push(
        <div key={c} data-col={c} className={c >= rect.c0 && c <= rect.c1 ? 'de-head-c de-head-on' : 'de-head-c'} style={{ left: colAxis.pos(c) - ox, top: 0, width: cw, height: HH }}>
          {colName(c)}
        </div>,
      )
      if (editable) out.push(<div key={'g' + c} data-grip-col={c} className="de-grip-c" style={{ left: colAxis.pos(c + 1) - ox - 3, height: HH }} />)
    }
    return (
      <div key={key} className="de-region" style={{ left, top: 0, width: w, height: HH }}>
        {out}
      </div>
    )
  }

  function rowHeaders(key: string, rr0: number, rr1: number, oy: number, top: number, h: number) {
    if (h <= 0) return null
    const out: ReactNode[] = []
    for (let r = rr0; r <= rr1; r++) {
      const rh = rowAxis.size(r)
      if (!rh) continue
      out.push(
        <div key={r} data-row={r} className={r >= rect.r0 && r <= rect.r1 ? 'de-head-r de-head-on' : 'de-head-r'} style={{ left: 0, top: rowAxis.pos(r) - oy, width: HW, height: rh }}>
          {r + 1}
        </div>,
      )
      if (editable) out.push(<div key={'g' + r} data-grip-row={r} className="de-grip-r" style={{ top: rowAxis.pos(r + 1) - oy - 3, width: HW }} />)
    }
    return (
      <div key={key} className="de-region" style={{ left: 0, top, width: HW, height: h }}>
        {out}
      </div>
    )
  }

  const fr = frozen.rows
  const fc = frozen.cols
  const fw = Math.min(frozenW, viewW)
  const fh = Math.min(frozenH, viewH)

  /** cell rectangle in grid-view coordinates (frozen-pane aware) */
  const cellBox = (r: number, c: number) => {
    const m = mergeAt(r, c)
    const x = HW + colAxis.pos(c) - (c >= fc ? scrollX : 0)
    const y = HH + rowAxis.pos(r) - (r >= fr ? scrollY : 0)
    return {
      x,
      y,
      w: m ? colAxis.pos(m.c1 + 1) - colAxis.pos(m.c0) : colAxis.size(c),
      h: m ? rowAxis.pos(m.r1 + 1) - rowAxis.pos(m.r0) : rowAxis.size(r),
    }
  }

  const scrollIntoView = (r: number, c: number) => {
    const el = scrollRef.current
    if (!el) return
    const x = colAxis.pos(c)
    const y = rowAxis.pos(r)
    // a cell of a frozen pane is always in view: going there brings the sheet back to its top / left, as Excel does
    if (c < fc && scrollX > 0) el.scrollLeft = 0
    if (r < fr && scrollY > 0) el.scrollTop = 0
    if (c >= fc) {
      if (x - frozenW < scrollX) el.scrollLeft = x - frozenW
      else if (x + colAxis.size(c) > mainOx + mainW) el.scrollLeft = x + colAxis.size(c) - frozenW - mainW
    }
    if (r >= fr && maxLogicalY > 0) {
      let ty = scrollY
      if (y - frozenH < scrollY) ty = y - frozenH
      else if (y + rowAxis.size(r) > mainOy + mainH) ty = y + rowAxis.size(r) - frozenH - mainH
      if (ty !== scrollY) el.scrollTop = (ty / maxLogicalY) * maxScrollY
    }
  }

  // a comment shown from the panel: its cell selected and in view
  const gotoN = s.goto?.n
  useEffect(() => {
    const g = s.goto
    if (!g) return
    setSel({ ar: g.r, ac: g.c, fr: g.r, fc: g.c })
    scrollIntoView(g.r, g.c)
  }, [gotoN]) // eslint-disable-line react-hooks/exhaustive-deps

  const cellAt = (clientX: number, clientY: number, el: HTMLElement) => {
    const b = el.getBoundingClientRect()
    const x = clientX - b.left - HW
    const y = clientY - b.top - HH
    const lx = x < frozenW ? x : x + scrollX
    const ly = y < frozenH ? y : y + scrollY
    return { r: y < 0 ? -1 : rowAxis.index(Math.max(0, ly)), c: x < 0 ? -1 : colAxis.index(Math.max(0, lx)) }
  }

  // ---------------------------------------------------------------- editing

  const focusGrid = () => ((editable ? editRef.current : null) ?? scrollRef.current)?.focus({ preventScroll: true })
  const selectSheet = () => setSel({ ar: 0, ac: 0, fr: Math.max(0, s.rows - 1), fc: Math.max(0, s.cols - 1), whole: 'all' })
  // Ctrl/⌘+A from the viewer chrome / the page; a cell being edited keeps its own select-all
  useSelectAll(() => {
    if (edit) return
    focusGrid()
    selectSheet()
  })
  const styleOf = (r: number, c: number) => store.row(sheet, r)?.find((x) => x[0] === c)?.[2] ?? store.placeStyle(sheet, r, c)

  const run = async (p: Promise<XlsxEditResult | null>, kind: 'edit' | 'undo' | 'redo' = 'edit') => {
    try {
      const res = await p
      store.applied(res, meta)
      const h = store.history
      // the worker's own count where it gives one: an edit that changed nothing records no step
      if (res?.history) store.history = res.history
      else if (res) {
        if (kind === 'edit') store.history = { undo: h.undo + 1, redo: 0 }
        else if (kind === 'undo') store.history = { undo: Math.max(0, h.undo - 1), redo: h.redo + 1 }
        else store.history = { undo: h.undo + 1, redo: Math.max(0, h.redo - 1) }
      }
      // nothing to undo or redo: the worker's history is shorter than counted (it keeps the last 200 edits)
      else if (kind === 'undo') store.history = { ...h, undo: 0 }
      else if (kind === 'redo') store.history = { ...h, redo: 0 }
      // The current workbook view publishes dirty/history; a detached store has no listeners.
      if (res) store.edited()
      else store.emit()
      setDetailTick((t) => t + 1)
    } catch (e) {
      console.error('[office-engine]', e)
      setToast((e as Error).message)
    }
  }

  const shown = (d: CellDetail | undefined, r: number, c: number) => {
    if (!d) return ''
    if (d.formula) return d.formula
    // dates & times: show the formatted value rather than the serial number
    if (d.type === 'number' && /[dmyhs]/i.test((meta.formats[styleOf(r, c)] ?? '').replace(/"[^"]*"|\[[^\]]*\]/g, ''))) {
      return store.row(sheet, r)?.find((x) => x[0] === c)?.[1] ?? d.value
    }
    return d.type === 'number' && locale && d.value !== '' && Number.isFinite(Number(d.value)) ? numberText(Number(d.value), locale) : d.value
  }

  const startEdit = async (initial?: string, bar = false) => {
    if (!editable || !s.done || edit) return
    const { ar, ac } = sel
    const begin = (value: string, initialText: string, mode: EditState['mode']) =>
      setEdit({ sheet, r: ar, c: ac, value, initial: initialText, bar, mode, caret: value.length, caretEnd: value.length, sync: bar ? 0 : 1 })
    // typing replaces the content: no need to know the current value
    if (initial !== undefined) return begin(initial, '\u0000', 'enter')
    const d = detailState && detailState.r === ar && detailState.c === ac ? detailState.d : await client.call('xlsx.cell', sheet, ar, ac).catch(() => undefined)
    const current = shown(d, ar, ac)
    begin(current, current, 'edit')
  }

  // Enter / Escape and the editor's blur all run before React re-renders: the first
  // one closes the edit, the blur that follows must neither write nor write twice
  const closed = useRef<EditState | null>(null)
  const commit = (move?: [number, number]) => {
    const e = edit
    if (!e || closed.current === e) return
    closed.current = e
    setEdit(null)
    const value = closeFormula(e.value)
    if (value !== e.initial) lastRun.current = run(client.call('xlsx.set', e.sheet, [{ r: e.r, c: e.c, input: value }], typed)).then(() => widen({ r0: e.r, c0: e.c, r1: e.r, c1: e.c }, e.sheet))
    if (e.sheet !== sheet) return backTo(e, move)
    if (move) moveBy(move[0], move[1], false)
    focusGrid()
  }

  const cancel = () => {
    const e = edit
    if (!e || closed.current === e) return
    closed.current = e
    setEdit(null)
    if (e.sheet !== sheet) return backTo(e)
    focusGrid()
  }

  /** Ctrl+Enter: the entry goes into every cell of the selection (a formula keeps its relative references) */
  const commitFill = () => {
    const e = edit
    if (!e || closed.current === e) return
    const R = e.sheet === sheet ? rect : { r0: e.r, c0: e.c, r1: e.r, c1: e.c }
    if ((R.r1 - R.r0 + 1) * (R.c1 - R.c0 + 1) > 100_000) return setToast(str.xlsx.fillLimit(100_000, str.locale))
    closed.current = e
    setEdit(null)
    const value = closeFormula(e.value)
    const cells: CellInput[] = []
    for (let r = R.r0; r <= R.r1; r++) {
      for (let c = R.c0; c <= R.c1; c++) {
        const m = e.sheet === sheet ? mergeAt(r, c) : undefined
        if (m && (m.r0 !== r || m.c0 !== c)) continue
        let input = value
        if (isFormula(value) && (r !== e.r || c !== e.c)) {
          try {
            input = '=' + translateFormula(value.slice(1), r - e.r, c - e.c)
          } catch {
            input = value
          }
        }
        cells.push({ r, c, input })
      }
    }
    lastRun.current = run(client.call('xlsx.set', e.sheet, cells, typed)).then(() => widen(R, e.sheet))
    if (e.sheet !== sheet) return backTo(e)
    focusGrid()
  }

  /** text arrived in the hidden field of the selected cell (an IME, a virtual keyboard, an emoji picker, a key): the edit starts with it */
  const sinkStart = (el: HTMLTextAreaElement) => {
    const v = el.value
    if (!v || !editable || !s.done) return
    scrollIntoView(sel.ar, sel.ac)
    setEdit({ sheet, r: sel.ar, c: sel.ac, value: v, initial: '\u0000', bar: false, mode: 'enter', caret: el.selectionStart ?? v.length, caretEnd: el.selectionEnd ?? v.length, sync: 0 })
  }

  /** F4: the reference at the caret goes relative → $A$1 → A$1 → $A1 → relative, both ends of a range alike */
  const cycleAbsolute = () => {
    if (!edit || !fctx) return
    const f = fctx.refs.find((x) => edit.caret >= x.s && edit.caretEnd <= x.e)
    if (!f) return
    const bang = f.text.lastIndexOf('!')
    const parts = f.text.slice(bang + 1).split(':')
    const m = /^(\$?)([A-Za-z]*)(\$?)(\d*)$/.exec(parts[0])
    if (!m) return
    const [, col, , row] = m
    const state = `${col ? 1 : 0}${row ? 1 : 0}`
    const next = state === '00' ? '11' : state === '11' ? '01' : state === '01' ? '10' : '00'
    const text = f.text.slice(0, bang + 1) + parts.map((x) => x.replace(/^(\$?)([A-Za-z]*)(\$?)(\d*)$/, (_, _c: string, letters: string, _r: string, digits: string) => (letters && next[0] === '1' ? '$' : '') + letters + (digits && next[1] === '1' ? '$' : '') + digits)).join(':')
    setEdit((p) => (p ? { ...p, value: p.value.slice(0, f.s) + text + p.value.slice(f.e), caret: f.s + text.length, caretEnd: f.s + text.length, sync: p.sync + 1, point: undefined } : p))
  }

  /** formula edited from another sheet: return to the edited cell's sheet */
  const backTo = (e: EditState, move?: [number, number]) => {
    const r = Math.max(0, e.r + (move?.[0] ?? 0))
    const c = Math.max(0, e.c + (move?.[1] ?? 0))
    store.sheet(e.sheet).sel = { ar: r, ac: c, fr: r, fc: c }
    store.focusGrid = true
    onSheet(e.sheet)
  }

  // ---------------------------------------------------------------- formula assist

  const own = !!edit && edit.sheet === sheet
  const inCell = own && !edit?.bar
  const fctx = edit && isFormula(edit.value) ? formulaContext(edit.value, edit.caret) : null
  const pointOk = !!edit && canPoint(edit, fctx)
  const pointActive = !!edit?.point && edit.point.e === edit.caret && edit.caret === edit.caretEnd

  // references of the formula being edited: colored spans + boxes on this sheet
  const refSpans: { s: number; e: number; color: string }[] = []
  const refBoxes: { r0: number; c0: number; r1: number; c1: number; color: string; point: boolean }[] = []
  if (fctx && edit) {
    const colors = new Map<string, string>()
    for (const f of fctx.refs) {
      const k = refKey(f.text)
      let color = colors.get(k)
      if (!color) colors.set(k, (color = REF_COLORS[colors.size % REF_COLORS.length]))
      refSpans.push({ s: f.s, e: f.e, color })
      const sh = f.ref.sheet
      const target = sh === undefined ? edit.sheet : meta.sheets.findIndex((x) => x.name.toLowerCase() === sh.toLowerCase())
      if (target === sheet) refBoxes.push({ r0: f.ref.r0, c0: f.ref.c0, r1: f.ref.r1, c1: f.ref.c1, color, point: pointActive && f.s === edit.point!.s })
    }
  }

  const word = fctx?.word && edit?.noSuggest !== fctx.word.s && edit?.caret === edit?.caretEnd ? fctx.word : undefined
  const sugg: FnInfo[] = word ? suggestFunctions(word.text, 8) : []
  const hint = !sugg.length && fctx?.call ? functionInfo(fctx.call.name) : undefined
  const sugKey = word ? `${word.s}:${word.text}` : ''
  useEffect(() => setSugIdx(0), [sugKey])

  const inputEl = (): HTMLTextAreaElement | null => (inCell ? editRef.current : fxRef.current)

  // push focus + caret into the active input (start, pointing, autocomplete, sheet switch)
  useLayoutEffect(() => {
    if (!edit) return
    const el = inputEl()
    if (!el) return
    const root = el.getRootNode() as Document | ShadowRoot
    if (root.activeElement !== el) el.focus({ preventScroll: true })
    else if (edit.sync && el === editRef.current && typeof matchMedia === 'function' && matchMedia('(pointer: coarse)').matches) {
      // the field had the focus without an on-screen keyboard (`inputMode` none): focus it again, now that the edit opens it
      hold.current = true
      el.blur()
      el.focus({ preventScroll: true })
      hold.current = false
    }
    if (edit.sync) el.setSelectionRange(edit.caret, edit.caretEnd)
  }, [edit?.sync, inCell, edit !== null]) // eslint-disable-line react-hooks/exhaustive-deps

  const readCaret = (el: HTMLTextAreaElement) => {
    const a = el.selectionStart ?? el.value.length
    const b = el.selectionEnd ?? a
    setEdit((p) => (p && (p.caret !== a || p.caretEnd !== b) ? { ...p, caret: a, caretEnd: b } : p))
  }

  const onInput = (el: HTMLTextAreaElement) => {
    const v = el.value
    const a = el.selectionStart ?? v.length
    const b = el.selectionEnd ?? a
    setEdit((p) => (p ? { ...p, value: v, caret: a, caretEnd: b, point: undefined } : p))
  }

  /** inserts / replaces the pointed reference at the caret */
  const pointAt = (ar: number, ac: number, fr: number, fc: number) => {
    setEdit((p) => {
      if (!p) return p
      const text = rangeText(ar, ac, fr, fc, sheet !== p.sheet ? meta.sheets[sheet]?.name : undefined)
      const live = p.point && p.point.e === p.caret && p.caret === p.caretEnd
      const s0 = live ? p.point!.s : p.caret
      const e0 = live ? p.point!.e : p.caretEnd
      if (live && p.value.slice(s0, e0) === text) return p
      const end = s0 + text.length
      return { ...p, value: p.value.slice(0, s0) + text + p.value.slice(e0), caret: end, caretEnd: end, sync: p.sync + 1, point: { s: s0, e: end, sheet, ar, ac, fr, fc } }
    })
  }

  const pointDrag = useRef<{ ar: number; ac: number; axis: 'cell' | 'col' | 'row' } | null>(null)

  /** keyboard pointing (Enter mode): arrows move the reference, Shift extends it */
  const movePoint = (dr: number, dc: number, extend: boolean) => {
    const p = edit?.point && pointActive && edit.point.sheet === sheet ? edit.point : null
    const baseR = p ? p.ar : own ? edit!.r : sel.ar
    const baseC = p ? p.ac : own ? edit!.c : sel.ac
    const clampR = (r: number) => Math.max(0, Math.min(rowAxis.count - 1, r))
    const clampC = (c: number) => Math.max(0, Math.min(colAxis.count - 1, c))
    if (extend) {
      const fr = clampR((p ? p.fr : baseR) + dr)
      const fc = clampC((p ? p.fc : baseC) + dc)
      pointAt(baseR, baseC, fr, fc)
      scrollIntoView(fr, fc)
    } else {
      const r = clampR(baseR + dr)
      const c = clampC(baseC + dc)
      pointAt(r, c, r, c)
      scrollIntoView(r, c)
    }
  }

  const accept = (f: FnInfo) => {
    const w = word
    if (!w) return
    setEdit((p) => {
      if (!p) return p
      const noArgs = !f.args.length && !f.variadic
      const ins = f.name + (w.call ? '' : noArgs ? '()' : '(')
      const caret = w.s + f.name.length + (noArgs && !w.call ? 2 : 1)
      return { ...p, value: p.value.slice(0, w.s) + ins + p.value.slice(w.e), caret, caretEnd: caret, sync: p.sync + 1, point: undefined }
    })
  }

  const moveBy = (dr: number, dc: number, extend: boolean) => {
    const baseR = extend ? sel.fr : sel.ar
    const baseC = extend ? sel.fc : sel.ac
    const m = !extend ? mergeAt(sel.ar, sel.ac) : undefined
    let r = baseR + dr
    let c = baseC + dc
    if (m && dr > 0) r = m.r1 + dr
    if (m && dc > 0) c = m.c1 + dc
    r = Math.max(0, Math.min(rowAxis.count - 1, r))
    c = Math.max(0, Math.min(colAxis.count - 1, c))
    while (r > 0 && rowAxis.size(r) === 0 && dr < 0) r--
    while (rowAxis.size(r) === 0 && r < rowAxis.count - 1) r++
    while (c > 0 && colAxis.size(c) === 0 && dc < 0) c--
    while (colAxis.size(c) === 0 && c < colAxis.count - 1) c++
    if (extend) setSelState({ ...sel, fr: r, fc: c, whole: undefined })
    else setSel({ ar: r, ac: c, fr: r, fc: c })
    scrollIntoView(r, c)
  }

  /** the rows r0..r1 as the worker has them (the cached blocks where they are whole, else read) */
  const rowsOf = async (r0: number, r1: number): Promise<Map<number, CellOut[]>> => {
    const out = new Map<number, CellOut[]>()
    for (let b = Math.floor(r0 / BLOCK); b <= Math.floor(r1 / BLOCK); b++) {
      const cached = s.blocks.get(b)
      let rows = cached?.complete ? cached.rows : undefined
      if (!rows) rows = new Map((await client.call('xlsx.block', sheet, b * BLOCK, (b + 1) * BLOCK)).rows.map((x) => [x.r, x.cells]))
      for (const [r, cells] of rows) if (r >= r0 && r <= r1) out.set(r, cells)
    }
    return out
  }

  const gotoCell = (r: number, c: number, extend: boolean) => {
    if (extend) setSelState((p) => ({ ...p, fr: r, fc: c, whole: undefined }))
    else setSel({ ar: r, ac: c, fr: r, fc: c })
    scrollIntoView(r, c)
  }

  /**
   * Ctrl+Arrow, as Excel has it: from a filled cell with a filled neighbour to the last cell of that block; else to
   * the next filled cell, or to the end of the data when there is none.
   */
  const jump = async (dr: number, dc: number, extend: boolean) => {
    const lastR = Math.max(0, s.rows - 1)
    const lastC = Math.max(0, s.cols - 1)
    let r = extend ? sel.fr : sel.ar
    let c = extend ? sel.fc : sel.ac
    const inside = (a: number, b: number) => a >= 0 && b >= 0 && a <= lastR && b <= lastC
    const seen = new Map<number, CellOut[]>()
    let reads = 0
    const filled = async (a: number, b: number) => {
      if (!seen.has(a)) {
        const from = Math.floor(a / BLOCK) * BLOCK
        reads++
        const rows = await rowsOf(from, from + BLOCK - 1)
        for (let k = from; k < from + BLOCK; k++) seen.set(k, rows.get(k) ?? [])
      }
      return seen.get(a)!.some((x) => x[0] === b && x[1] !== '')
    }
    if (!inside(r + dr, c + dc)) return
    if ((await filled(r, c)) && (await filled(r + dr, c + dc))) {
      while (inside(r + dr, c + dc) && (await filled(r + dr, c + dc))) {
        r += dr
        c += dc
      }
    } else {
      do {
        r += dr
        c += dc
      } while (inside(r, c) && reads < 300 && !(await filled(r, c)))
      // nothing filled ahead (or too far to look): the end of the data
      if (!inside(r, c) || reads >= 300) {
        r = dr > 0 ? lastR : dr < 0 ? 0 : Math.max(0, Math.min(lastR, r))
        c = dc > 0 ? lastC : dc < 0 ? 0 : Math.max(0, Math.min(lastC, c))
      }
    }
    gotoCell(r, c, extend)
  }

  /** Home / End / Ctrl+Home / Ctrl+End */
  const homeEnd = async (end: boolean, mod: boolean, extend: boolean) => {
    const row = extend ? sel.fr : sel.ar
    if (mod) return gotoCell(end ? Math.max(0, s.rows - 1) : 0, end ? Math.max(0, s.cols - 1) : 0, extend)
    if (!end) return gotoCell(row, 0, extend)
    // End: the last filled cell of the row (the last column of the sheet when the row is empty)
    const cells = (await rowsOf(row, row)).get(row) ?? []
    const last = cells.reduce((m, x) => (x[1] !== '' ? Math.max(m, x[0]) : m), -1)
    gotoCell(row, last >= 0 ? last : Math.max(0, s.cols - 1), extend)
  }

  /** Tab and Enter in a row of entries: Enter goes back to the column the Tabs began in (and a Shift+Enter or an arrow forgets it) */
  const tabMove = (back: boolean, col: number): [number, number] => {
    if (tabStart.current === null) tabStart.current = col
    return [0, back ? -1 : 1]
  }
  const enterMove = (back: boolean, col: number): [number, number] => {
    const start = tabStart.current
    tabStart.current = null
    return [back ? -1 : 1, start !== null && !back ? start - col : 0]
  }

  /** columns that show #### for the numbers written in `range` are widened, as Excel does (a width set by hand stays) */
  const widen = async (range: Range, at = sheet) => {
    // it runs after the edit's result: the document may have been closed (another one opened) meanwhile
    if (!editable || at !== sheet || client.closed) return
    const live = store.sheet(sheet)
    const need = new Map<number, number>()
    const rows = await rowsOf(range.r0, Math.min(range.r1, range.r0 + 999)).catch((e: Error) => {
      if (client.closed) return null
      throw e
    })
    if (!rows) return
    for (const [, cells] of rows) {
      for (const [c, text, si, kind] of cells) {
        if (c < range.c0 || c > range.c1 || kind !== 1 || !text || (meta.formats[si] ?? 'General') === 'General' || live.colAxis.size(c) !== live.colAxis.def) continue
        const css = (styles[si] ?? {}) as Record<string, string | number>
        const px = fontPx(css.fontSize, basePx)
        const font = fontOf(css, meta.baseFont as Record<string, string>, basePx)
        if (fitNumber(text, live.colAxis.size(c) - 6, px, font) === text) continue
        // as wide as the text, and as wide as the grid needs to show it whole (it decides, by its own measure, between the number and ####)
        let w = Math.ceil(textWidth(text, css, meta.baseFont as Record<string, string>, basePx)) + 12
        while (w < 400 && fitNumber(text, w - 6, px, font) !== text) w += 6
        need.set(c, Math.max(need.get(c) ?? 0, w))
      }
    }
    for (const [c, px] of need) {
      if (client.closed) return
      await sheetOp({ op: 'colWidth', c0: c, c1: c, px: Math.min(px, 400) })
    }
  }

  /** a column as wide as its longest text (double click on its edge) */
  const fitCol = async (c: number) => {
    let best = 0
    const total = Math.min(s.rows, 20_000)
    for (let r0 = 0; r0 < total; r0 += BLOCK * 8) {
      for (const cells of (await rowsOf(r0, Math.min(total - 1, r0 + BLOCK * 8 - 1))).values()) {
        const cell = cells.find((x) => x[0] === c)
        const css = (cell && styles[cell[2]]) as Record<string, string | number> | undefined
        if (cell?.[1] && css?.whiteSpace !== 'pre-wrap') best = Math.max(best, textWidth(cell[1], css ?? {}, meta.baseFont as Record<string, string>, basePx))
      }
    }
    return best ? Math.min(600, Math.ceil(best) + 12) : colAxis.def
  }

  /**
   * Row heights from what the rows hold: those of wrapped text, wide as its column. `shrink`: rows may get lower
   * (double click on an edge); else they only grow (Wrap text was turned on: a height set by hand is not undone).
   */
  const fitRows = async (range: Range, shrink: boolean) => {
    const live = store.sheet(sheet)
    const sizes: [number, number][] = []
    const last = Math.min(range.r1, range.r0 + 999)
    const held = await rowsOf(range.r0, last)
    for (let r = range.r0; r <= last; r++) {
      const cells = held.get(r) ?? []
      if (live.rowAxis.size(r) === 0) continue
      let h = shrink ? live.rowAxis.def : 0
      for (const [c, text, si] of cells) {
        if (!text) continue
        const css = (styles[si] ?? {}) as Record<string, string | number>
        if (css.textRotate) h = Math.max(h, Math.min(MAX_FIT_ROW, rotatedHeight(text, css, meta.baseFont as Record<string, string>, basePx, Number(css.textRotate))))
        else if (css.whiteSpace === 'pre-wrap') h = Math.max(h, wrapHeight(text, css, meta.baseFont as Record<string, string>, basePx, live.colAxis.size(c) - 6))
        else if (shrink) h = Math.max(h, Math.ceil(fontPx(css.fontSize, basePx) * 1.2) + 3)
      }
      if (h && (shrink ? h !== live.rowAxis.size(r) : h > live.rowAxis.size(r))) sizes.push([r, h])
    }
    for (let i = 0; i < sizes.length; ) {
      let j = i
      while (j + 1 < sizes.length && sizes[j + 1][0] === sizes[j][0] + 1 && sizes[j + 1][1] === sizes[i][1]) j++
      await sheetOp({ op: 'rowHeight', r0: sizes[i][0], r1: sizes[j][0], px: sizes[i][1] })
      i = j + 1
    }
  }

  // in the worker: every cell of the range, also those of rows not shown yet (a whole column, the whole sheet)
  const clearRange = () => void sheetOp({ op: 'clear', range: rect })

  const activeCss = (styles[styleOf(sel.ar, sel.ac)] ?? {}) as Record<string, string | number>
  // whole columns / rows / the sheet: the format goes to them, down to the last row, not only to the cells there are
  const fmtRange = (): Range =>
    sel.whole === 'all' ? { r0: 0, c0: 0, r1: MAX_ROW, c1: MAX_COL } : sel.whole === 'col' ? { ...rect, r0: 0, r1: MAX_ROW } : sel.whole === 'row' ? { ...rect, c0: 0, c1: MAX_COL } : rect
  const fmt = (patch: FmtPatch) => fmtAt(fmtRange(), patch)
  /** `focus`: the grid takes the focus back (the built-in toolbar); a host's control keeps it */
  const fmtAt = (range: Range, patch: FmtPatch, focus = true) => {
    const done = run(client.call('xlsx.format', sheet, range, patch))
    // Wrap text shows all of the text: the rows grow; a number format shows its numbers: the columns widen
    if (patch.wrap) void done.then(() => fitRows(range, false))
    if (patch.numFmt !== undefined) void done.then(() => widen(range))
    if (focus) focusGrid()
    return done
  }
  const toggle = (k: 'b' | 'i' | 'u' | 'strike') => {
    const on = cellFmt[k === 'b' ? 'bold' : k === 'i' ? 'italic' : k === 'u' ? 'underline' : 'strike']
    fmt({ [k]: !on })
  }
  const undo = () => void run(client.call('xlsx.undo'), 'undo')
  const redo = () => void run(client.call('xlsx.redo'), 'redo')
  // the controller's undo / redo and formatting: this sheet's grid runs them while it is shown
  const cellFmt = cellFormat(activeCss, meta.baseFont as { fontFamily?: string; fontSize?: string | number })
  // Commands sent faster than the cell shows their effect (a quick second or third toggle) go by the marks the earlier
  // ones left, not by the format last read from the cell: the worker answers each in turn, and the cell's rows are read
  // again after that. Those marks are kept while any command is unanswered and until the cell shows them; another
  // selection, undo and redo drop them (an undo can bring back the very format they were worked out from).
  const selKey = `${sheet}:${rect.r0}:${rect.c0}:${rect.r1}:${rect.c1}:${sel.whole ?? ''}`
  const marks = (f: TextFormat) => `${f.bold}${f.italic}${f.underline}${f.strike}`
  const pendingFmt = useRef<{ sel: string; fmt: TextFormat; waiting: number } | null>(null)
  {
    const p = pendingFmt.current
    if (p && (p.sel !== selKey || (p.waiting === 0 && marks(p.fmt) === marks(cellFmt)))) pendingFmt.current = null
  }
  const publishFormat = useViewCommands(viewer.ctl, {
    goto: gotoSheet,
    undo: () => {
      pendingFmt.current = null
      if (editable) undo()
    },
    redo: () => {
      pendingFmt.current = null
      if (editable) redo()
    },
    format: (c) => {
      const p = pendingFmt.current && pendingFmt.current.sel === selKey ? pendingFmt.current : null
      const now = p ? p.fmt : cellFmt
      const patch = editable && cellPatch(c, now)
      if (!patch) return
      const q = p ?? { sel: selKey, fmt: now, waiting: 0 }
      q.fmt = marksAfter(c, now)
      q.waiting++
      pendingFmt.current = q
      void fmtAt(fmtRange(), patch, false).finally(() => {
        q.waiting--
      })
    },
  })
  useEffect(() => {
    publishFormat(editable ? cellFmt : null)
    const h = store.history
    viewer.ctl?.update({ canUndo: editable && h.undo > 0, canRedo: editable && h.redo > 0 })
  })

  /** the selection as clipboard text, when all its rows are at hand (else null: the worker has them) */
  const tsv = (): string | null => {
    const r1 = Math.min(rect.r1, s.rows - 1)
    const c1 = Math.min(rect.c1, s.cols - 1)
    if (r1 - rect.r0 > 2000 || !store.loaded(sheet, rect.r0, Math.max(rect.r0, r1))) return null
    const lines: string[] = []
    for (let r = rect.r0; r <= Math.max(rect.r0, r1); r++) {
      const row = store.row(sheet, r)
      const vals: string[] = []
      for (let c = rect.c0; c <= Math.max(rect.c0, c1); c++) vals.push(tsvCell(row?.find((x) => x[0] === c)?.[1] ?? ''))
      lines.push(vals.join('\t'))
    }
    return lines.join('\n')
  }

  /**
   * Copy / cut: the cells are read in the worker as they are (values, formulas, formats: what Paste writes here),
   * all of them, the rows not shown yet included. The clipboard gets their text: at once when the rows are at
   * hand, as the event wants it, else when the worker has read them.
   */
  const copySel = (e: React.ClipboardEvent, cut: boolean) => {
    e.preventDefault()
    const range = { ...rect }
    const local = tsv()
    if (local !== null) e.clipboardData.setData('text/plain', local)
    store.clip = null
    client
      .call('xlsx.copy', sheet, range)
      .then(async (data) => {
        if (local === null) await navigator.clipboard.writeText(data.text)
        store.clip = { text: local ?? data.text, data, cut }
        if (cut) await sheetOp({ op: 'clear', range })
      })
      .catch((err: Error) => setToast(local === null && err.name === 'NotAllowedError' ? str.xlsx.clipboardCopyBlocked : err.message))
  }

  const onKeyDown = (e: React.KeyboardEvent) => {
    if (edit) return
    const mod = e.metaKey || e.ctrlKey
    const k = e.key.toLowerCase()
    if (mod && k === 's') {
      e.preventDefault()
      void viewer.save()
      return
    }
    if (mod && k === 'a' && !e.altKey && !e.shiftKey) {
      e.preventDefault()
      selectSheet()
      return
    }
    if (editable && mod && !e.altKey) {
      if (k === 'z') {
        e.preventDefault()
        if (e.shiftKey) redo()
        else undo()
        return
      }
      if (k === 'y') {
        e.preventDefault()
        redo()
        return
      }
      if (k === 'b' || k === 'i' || k === 'u') {
        e.preventDefault()
        toggle(k)
        return
      }
      if (k === 'd' || k === 'r') {
        e.preventDefault()
        void sheetOp({ op: 'fill', range: rect, dir: k === 'd' ? 'down' : 'right' })
        return
      }
      if (k === '\\') {
        e.preventDefault()
        fmt({ clear: true })
        return
      }
      if (k === 'x' && e.shiftKey) {
        e.preventDefault()
        toggle('strike')
        return
      }
    }
    if (mod && (k === 'f' || k === 'h')) {
      e.preventDefault()
      setFind((f) => ({ ...f, open: true, replace: k === 'h' || f.replace, tick: f.tick + 1 }))
      return
    }
    // Shift+Space selects the row(s), Ctrl+Space the column(s) of the selection
    if (e.code === 'Space' && (e.shiftKey || mod) && !e.altKey) {
      e.preventDefault()
      if (mod) setSel({ ar: 0, ac: rect.c0, fr: Math.max(0, s.rows - 1), fc: rect.c1, whole: 'col' })
      else setSel({ ar: rect.r0, ac: 0, fr: rect.r1, fc: Math.max(0, s.cols - 1), whole: 'row' })
      return
    }
    if (e.key === 'Escape' && painter !== null) {
      setPainter(null)
      return
    }
    const pageRows = Math.max(1, Math.floor(mainH / rowAxis.def) - 1)
    const moves: Record<string, [number, number]> = {
      ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1],
      PageDown: [pageRows, 0], PageUp: [-pageRows, 0],
    }
    if (moves[e.key]) {
      e.preventDefault()
      tabStart.current = null
      const [dr, dc] = moves[e.key]
      if (mod && !e.key.startsWith('Page')) void jump(dr, dc, e.shiftKey)
      else moveBy(dr, dc, e.shiftKey)
      return
    }
    if (e.key === 'Home' || e.key === 'End') {
      e.preventDefault()
      tabStart.current = null
      void homeEnd(e.key === 'End', mod, e.shiftKey)
      return
    }
    if (e.key === 'Enter' || e.key === 'Tab') {
      e.preventDefault()
      if (e.key === 'Enter') moveBy(...enterMove(e.shiftKey, sel.ac), false)
      else moveBy(...tabMove(e.shiftKey, sel.ac), false)
      return
    }
    if (!editable) return
    if (e.altKey && e.code === 'Equal') {
      // AutoSum
      e.preventDefault()
      autoFn('SUM')
      return
    }
    if (mod && !e.altKey && e.code === 'Semicolon') {
      // today's date (with Shift: the time), as a value in a date format
      e.preventDefault()
      stamp(e.shiftKey)
      return
    }
    if (e.key === 'F2') {
      e.preventDefault()
      void startEdit()
    } else if (e.key === 'Delete' || e.key === 'Backspace') {
      e.preventDefault()
      clearRange()
    }
    // (a key that types is not handled here: its text lands in the hidden field of the grid, which starts the edit, so that
    // what an IME or a virtual keyboard composes there is not cut in two)
  }

  /** the date (time) of now in the active cell */
  const stamp = (time: boolean) => {
    const d = new Date()
    const serial = time ? (d.getHours() * 3600 + d.getMinutes() * 60 + d.getSeconds()) / 86400 : Math.floor(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()) / 86400000) + 25569
    const at = { r0: sel.ar, c0: sel.ac, r1: sel.ar, c1: sel.ac }
    lastRun.current = (async () => {
      await run(client.call('xlsx.set', sheet, [{ r: sel.ar, c: sel.ac, input: String(serial) }]))
      await run(client.call('xlsx.format', sheet, at, { numFmt: time ? 'hh:mm' : 'yyyy-mm-dd' }))
      await widen(at)
    })()
  }

  /** the cells the handle would fill if let go over `to`: down or right of the selection, whichever it is further along */
  const fillExtent = (to: { r: number; c: number }): Range | null => {
    const src = fillRef.current?.src ?? rect
    const dr = to.r - src.r1
    const dc = to.c - src.c1
    if (dr > 0 && dr >= dc) return { r0: src.r0, c0: src.c0, r1: to.r, c1: src.c1 }
    if (dc > 0) return { r0: src.r0, c0: src.c0, r1: src.r1, c1: to.c }
    return null
  }

  /**
   * Fill handle: the cells copied down or to the right as far as the handle was dragged. One row (column) is the
   * `fill` operation; several are repeated, as a block of copied cells is.
   */
  const doFill = async (src: Range, to: { r: number; c: number }) => {
    const ext = fillExtent(to)
    if (!ext) return
    const down = ext.r1 > src.r1
    const h = src.r1 - src.r0 + 1
    const w = src.c1 - src.c0 + 1
    if (down ? h === 1 : w === 1) await sheetOp({ op: 'fill', range: ext, dir: down ? 'down' : 'right' })
    else {
      const clip = await client.call('xlsx.copy', sheet, src)
      const rows = down ? ext.r1 - src.r1 : h
      const cols = down ? w : ext.c1 - src.c1
      const cells: ClipCell[][] = []
      for (let i = 0; i < rows; i++) {
        const line: ClipCell[] = []
        for (let j = 0; j < cols; j++) {
          const cell = clip.cells[down ? i % h : i]?.[down ? j : j % w] ?? null
          const tile = down ? Math.floor(i / h) : Math.floor(j / w)
          line.push(shiftedClip(cell, down ? tile * h : 0, down ? 0 : tile * w))
        }
        cells.push(line)
      }
      const range = down ? { r0: src.r1 + 1, c0: src.c0, r1: ext.r1, c1: src.c1 } : { r0: src.r0, c0: src.c1 + 1, r1: src.r1, c1: ext.c1 }
      await sheetOp({ op: 'paste', range, cells, from: { sheet: clip.sheet, r0: src.r0, c0: src.c0 } })
    }
    setSelState({ ar: src.r0, ac: src.c0, fr: ext.r1, fc: ext.c1 })
  }

  /** double click on the edge of a column / row: it is as wide / high as what it holds (all the selected ones, when they are) */
  const fitGrip = async (grip: HTMLElement) => {
    const col = grip.dataset.gripCol
    const i = Number(col ?? grip.dataset.gripRow)
    if (col !== undefined) {
      const full = rect.r0 === 0 && rect.r1 >= s.rows - 1 && i >= rect.c0 && i <= rect.c1
      for (let c = full ? rect.c0 : i; c <= (full ? Math.min(rect.c1, rect.c0 + 199) : i); c++) await sheetOp({ op: 'colWidth', c0: c, c1: c, px: await fitCol(c) })
    } else {
      const full = rect.c0 === 0 && rect.c1 >= s.cols - 1 && i >= rect.r0 && i <= rect.r1
      await fitRows({ r0: full ? rect.r0 : i, c0: 0, r1: full ? rect.r1 : i, c1: 0 }, true)
    }
  }

  const onMouseDown = (e: React.MouseEvent) => {
    if (e.button !== 0) return
    if ((e.target as HTMLElement).closest('[data-fill]') && !edit) {
      e.preventDefault()
      fillRef.current = { src: { ...rect } }
      afterDrag.current = () => {
        const f = fillRef.current
        fillRef.current = null
        setFillTo(null)
        if (f?.to) void doFill(f.src, f.to)
      }
      return
    }
    tabStart.current = null
    const grip = (e.target as HTMLElement).closest('[data-grip-col],[data-grip-row]') as HTMLElement | null
    if (grip) {
      e.preventDefault()
      const col = grip.dataset.gripCol
      const i = Number(col ?? grip.dataset.gripRow)
      setResizing(col !== undefined ? { axis: 'col', i, start: e.clientX, size: colAxis.size(i) } : { axis: 'row', i, start: e.clientY, size: rowAxis.size(i) })
      return
    }
    const el = e.currentTarget as HTMLElement
    const { r, c } = cellAt(e.clientX, e.clientY, el)
    if (edit && pointOk) {
      // formula: clicking / dragging cells inserts references, the input keeps focus
      e.preventDefault()
      if (r < 0 && c < 0) return
      const axis = r < 0 ? 'col' : c < 0 ? 'row' : 'cell'
      let [ar, ac] = [Math.max(r, 0), Math.max(c, 0)]
      const m = axis === 'cell' ? mergeAt(r, c) : undefined
      if (m) [ar, ac] = [m.r0, m.c0]
      const p = e.shiftKey && pointActive && edit.point!.sheet === sheet ? edit.point! : null
      if (p) [ar, ac] = [p.ar, p.ac]
      pointDrag.current = { ar, ac, axis }
      pointRange(ar, ac, r, c, axis)
      return
    }
    if (edit) commit()
    e.preventDefault()
    if (painter !== null && r >= 0 && c >= 0) {
      const xf = painter
      afterDrag.current = () => {
        setPainter(null)
        fmtAt(selRef.current, { xf })
      }
    }
    focusGrid()
    if (r < 0 && c < 0) {
      setSel({ ar: 0, ac: 0, fr: Math.max(0, s.rows - 1), fc: Math.max(0, s.cols - 1), whole: 'all' })
      return
    }
    if (r < 0) setSel({ ar: 0, ac: c, fr: Math.max(0, s.rows - 1), fc: c, whole: 'col' })
    else if (c < 0) setSel({ ar: r, ac: 0, fr: r, fc: Math.max(0, s.cols - 1), whole: 'row' })
    else if (e.shiftKey) setSelState({ ...sel, fr: r, fc: c, whole: undefined })
    else {
      setSel({ ar: r, ac: c, fr: r, fc: c })
      // a cell with a comment shows its thread
      const id = cmAt(r, c)
      if (id) cm?.select(id, false)
      else if (cmState.active) cm?.select(null)
    }
    if (r >= 0 && c >= 0) dragging.current = true
  }

  /** pointed block from the drag anchor to (r, c); header drags select whole columns / rows */
  const pointRange = (ar: number, ac: number, r: number, c: number, axis: 'cell' | 'col' | 'row') => {
    if (axis === 'col') pointAt(0, ac, MAX_ROW, Math.max(0, c))
    else if (axis === 'row') pointAt(ar, 0, Math.max(0, r), MAX_COL)
    else {
      const m0 = mergeAt(r, c)
      // inside the anchor's merge: just the anchor cell
      const m = m0 && !(ar >= m0.r0 && ar <= m0.r1 && ac >= m0.c0 && ac <= m0.c1) ? m0 : undefined
      // dragging over a merge takes all of it
      const fr = m ? (m.r0 < ar ? m.r0 : m.r1) : Math.max(0, r)
      const fc = m ? (m.c0 < ac ? m.c0 : m.c1) : Math.max(0, c)
      pointAt(ar, ac, fr, fc)
    }
  }

  useEffect(() => {
    const up = () => {
      dragging.current = false
      pointDrag.current = null
      const fn = afterDrag.current
      afterDrag.current = null
      if (fn) setTimeout(fn, 0)
    }
    window.addEventListener('mouseup', up)
    return () => window.removeEventListener('mouseup', up)
  }, [])

  useEffect(() => {
    if (!resizing) return
    const axis = resizing.axis === 'col' ? colAxis : rowAxis
    const move = (e: MouseEvent) => {
      const d = (resizing.axis === 'col' ? e.clientX : e.clientY) - resizing.start
      axis.set(resizing.i, Math.max(resizing.axis === 'col' ? 8 : 6, Math.round(resizing.size + d)))
      setDetailTick((t) => t + 1)
    }
    const up = () => {
      const px = axis.size(resizing.i)
      setResizing(null)
      // resize every selected full column / row together
      const full = resizing.axis === 'col' ? rect.r0 === 0 && rect.r1 >= s.rows - 1 && resizing.i >= rect.c0 && resizing.i <= rect.c1 : rect.c0 === 0 && rect.c1 >= s.cols - 1 && resizing.i >= rect.r0 && resizing.i <= rect.r1
      const [a, b] = full ? (resizing.axis === 'col' ? [rect.c0, rect.c1] : [rect.r0, rect.r1]) : [resizing.i, resizing.i]
      void sheetOp(resizing.axis === 'col' ? { op: 'colWidth', c0: a, c1: b, px } : { op: 'rowHeight', r0: a, r1: b, px })
    }
    window.addEventListener('mousemove', move)
    window.addEventListener('mouseup', up, { once: true })
    return () => {
      window.removeEventListener('mousemove', move)
      window.removeEventListener('mouseup', up)
    }
  })

  const sheetOp = (op: SheetOp) => run(client.call('xlsx.sheetOp', sheet, op))
  // the selection holds merged cells (it always covers them whole): Unmerge / toggle unmerges them all
  const merged = s.merges.some((m) => m.r0 >= rect.r0 && m.r1 <= rect.r1 && m.c0 >= rect.c0 && m.c1 <= rect.c1)
  const toggleMerge = () => merge(merged ? 'unmerge' : 'all')
  const merge = async (mode: MergeMode) => {
    if (mode === 'unmerge') return void sheetOp({ op: 'unmerge', range: rect })
    // like Excel, only the upper-left value of each merged block survives: say what goes (Undo brings it back)
    let dropped = 0
    const range = { ...rect }
    if (range.r1 - range.r0 < 5000) {
      for (const [r, cells] of await rowsOf(range.r0, range.r1)) {
        for (const [c, text] of cells) {
          if (c < range.c0 || c > range.c1 || !text) continue
          if (!(mode === 'all' ? r === range.r0 && c === range.c0 : mode === 'across' ? c === range.c0 : r === range.r0)) dropped++
        }
      }
    }
    await sheetOp({ op: 'merge', range, mode })
    if (dropped) setToast(str.xlsx.mergeDropped(dropped, `${MOD}Z`))
  }
  const freeze = (kind: 'none' | 'row' | 'col' | 'sel') =>
    void sheetOp({ op: 'freeze', rows: kind === 'row' ? 1 : kind === 'sel' ? sel.ar : 0, cols: kind === 'col' ? 1 : kind === 'sel' ? sel.ac : 0 })
  /** AutoSum-style: the function over the numbers above (or left of) the active cell */
  const autoFn = (fn = 'SUM') => {
    const { ar, ac } = sel
    if (!oneCell) {
      // a selected block: result goes under it
      const ref = `${colName(rect.c0)}${rect.r0 + 1}:${colName(rect.c1)}${rect.r1 + 1}`
      void run(client.call('xlsx.set', sheet, [{ r: rect.r1 + 1, c: rect.c0, input: `=${fn}(${ref})` }]))
      return
    }
    const num = (r: number, c: number) => store.row(sheet, r)?.find((x) => x[0] === c)?.[3] === 1
    let r = ar - 1
    while (r >= 0 && num(r, ac)) r--
    let ref = ''
    if (r < ar - 1) ref = `${colName(ac)}${r + 2}:${colName(ac)}${ar}`
    else {
      let c = ac - 1
      while (c >= 0 && num(ar, c)) c--
      if (c < ac - 1) ref = `${colName(c + 1)}${ar + 1}:${colName(ac - 1)}${ar + 1}`
    }
    if (!ref) return void startEdit(`=${fn}(`)
    void run(client.call('xlsx.set', sheet, [{ r: ar, c: ac, input: `=${fn}(${ref})` }]))
  }
  const sortBy = async (desc: boolean) => {
    // The rows of the table around the selection, whole (the selected cells alone would be parted from their rows):
    // all of them when the selection is in one row, else the selected ones. Its header stays on top, like frozen rows.
    const picked = { ...rect }
    const col = sel.ac
    let table: { range: Range; header: boolean }
    try {
      table = await client.call('xlsx.region', sheet, picked)
    } catch (e) {
      return setToast((e as Error).message)
    }
    const R = table.range
    let r0 = picked.r0 === picked.r1 ? R.r0 : Math.max(picked.r0, R.r0)
    const r1 = picked.r0 === picked.r1 ? R.r1 : Math.min(picked.r1, R.r1)
    if (r0 === R.r0 && table.header) r0++
    r0 = Math.max(r0, s.frozen.rows)
    if (r1 <= r0) return
    void sheetOp({ op: 'sort', range: { r0, c0: R.c0, r1, c1: R.c1 }, col, desc })
  }
  const rowsOp = (k: string) => {
    const nr = rect.r1 - rect.r0 + 1
    const nc = rect.c1 - rect.c0 + 1
    const ops: Record<string, SheetOp> = {
      rowAbove: { op: 'insertRows', at: rect.r0, n: nr },
      rowBelow: { op: 'insertRows', at: rect.r1 + 1, n: nr },
      colLeft: { op: 'insertCols', at: rect.c0, n: nc },
      colRight: { op: 'insertCols', at: rect.c1 + 1, n: nc },
      delRow: { op: 'deleteRows', at: rect.r0, n: nr },
      delCol: { op: 'deleteCols', at: rect.c0, n: nc },
      hideRows: { op: 'hideRows', a: rect.r0, b: rect.r1, hidden: true },
      hideCols: { op: 'hideCols', a: rect.c0, b: rect.c1, hidden: true },
      // unhide what lies inside (or right around) the selection
      unhideRows: { op: 'hideRows', a: Math.max(0, rect.r0 - (nr === 1 ? 1 : 0)), b: rect.r1 + (nr === 1 ? 1 : 0), hidden: false },
      unhideCols: { op: 'hideCols', a: Math.max(0, rect.c0 - (nc === 1 ? 1 : 0)), b: rect.c1 + (nc === 1 ? 1 : 0), hidden: false },
    }
    if (ops[k]) void sheetOp(ops[k])
  }
  /** merged cells that the block `r0..r1` × `c0..c1` would fill in other cells than their first, with something */
  const mergedHit = (b: Range, filledAt: (r: number, c: number) => boolean) =>
    s.merges.some((m) => {
      if (m.r1 < b.r0 || m.r0 > b.r1 || m.c1 < b.c0 || m.c0 > b.c1) return false
      for (let r = Math.max(m.r0, b.r0); r <= Math.min(m.r1, b.r1); r++) for (let c = Math.max(m.c0, b.c0); c <= Math.min(m.c1, b.c1); c++) if ((r !== m.r0 || c !== m.c0) && filledAt(r, c)) return true
      return false
    })
  const MERGED_PASTE = str.xlsx.mergedPaste
  const pasteText = (raw: string) => {
    const norm = (t: string) => t.replace(/\r\n?/g, '\n').replace(/\n$/, '')
    const at = { r: rect.r0, c: rect.c0 }
    const clip = store.clip
    // the cells copied here: as they were (values, formulas, formats), not as the text of them
    if (clip && norm(raw) === norm(clip.text)) {
      const { cells: copied, range } = clip.data
      const h = copied.length
      const w = Math.max(1, ...copied.map((row) => row.length))
      // one copied block into a selection that holds a whole number of them: it fills the selection (copied cells are not cut ones)
      const selH = rect.r1 - rect.r0 + 1
      const selW = rect.c1 - rect.c0 + 1
      const tiles = !clip.cut && !oneCell && h > 0 && selH % h === 0 && selW % w === 0 && selH * selW <= 100_000 ? { r: selH / h, c: selW / w } : { r: 1, c: 1 }
      const cells: ClipCell[][] = []
      for (let i = 0; i < h * tiles.r; i++) {
        const line: ClipCell[] = []
        for (let j = 0; j < w * tiles.c; j++) {
          const src = copied[i % h]?.[j % w] ?? null
          line.push(shiftedClip(src, Math.floor(i / h) * h, Math.floor(j / w) * w))
        }
        cells.push(line)
      }
      const dest = { r0: at.r, c0: at.c, r1: at.r + cells.length - 1, c1: at.c + w * tiles.c - 1 }
      if (cells.length * w * tiles.c > 1 && mergedHit(dest, (r, c) => !!cells[r - dest.r0]?.[c - dest.c0] && cells[r - dest.r0][c - dest.c0]![1] !== '')) return setToast(MERGED_PASTE)
      void sheetOp({ op: 'paste', range: dest, cells, from: { sheet: clip.data.sheet, r0: range.r0, c0: range.c0 }, ...(clip.cut ? { moved: true } : {}) })
      // moved once: pasted again, they are copied
      if (clip.cut) store.clip = { ...clip, cut: false }
      return
    }
    const rows = parseTsv(raw)
    // one value into a selection of several cells: all of them get it
    if (rows.length === 1 && rows[0].length === 1 && !oneCell) {
      const cells: CellInput[] = []
      if ((rect.r1 - rect.r0 + 1) * (rect.c1 - rect.c0 + 1) > 100_000) return setToast(str.xlsx.fillLimit(100_000, str.locale))
      for (let r = rect.r0; r <= rect.r1; r++) {
        for (let c = rect.c0; c <= rect.c1; c++) {
          const m = mergeAt(r, c)
          if (!m || (m.r0 === r && m.c0 === c)) cells.push({ r, c, input: rows[0][0] })
        }
      }
      void run(client.call('xlsx.set', sheet, cells))
      return
    }
    const cells = rows.flatMap((line, i) => line.map((v, j) => ({ r: at.r + i, c: at.c + j, input: v })))
    const block = { r0: at.r, c0: at.c, r1: at.r + rows.length - 1, c1: at.c + Math.max(1, ...rows.map((l) => l.length)) - 1 }
    if (cells.length > 1 && mergedHit(block, (r, c) => (rows[r - block.r0]?.[c - block.c0] ?? '') !== '')) return setToast(str.xlsx.mergedPaste)
    if (cells.length > 100_000) setToast(str.xlsx.pasteLimit(100_000, cells.length, str.locale))
    if (cells.length) void run(client.call('xlsx.set', sheet, cells.slice(0, 100_000)))
  }
  const ctxItems = (): MenuItem[] => {
    const nr = rect.r1 - rect.r0 + 1
    const nc = rect.c1 - rect.c0 + 1
    const rows = nr === 1 ? str.xlsx.row(String(rect.r0 + 1)) : str.xlsx.rows(String(rect.r0 + 1), String(rect.r1 + 1))
    const cols = nc === 1 ? str.xlsx.col(colName(rect.c0)) : str.xlsx.cols(colName(rect.c0), colName(rect.c1))
    const single = nr === 1 && nc === 1
    const thread = cmAt(sel.ar, sel.ac)
    const comment: MenuItem[] = thread
      ? [{ label: str.edit.showComment, icon: 'comment', run: () => cm?.select(thread) }]
      : cm?.canComment
        ? [{ label: str.edit.comment, icon: 'commentAdd', hint: COMMENT_KEYS, run: () => cm.start() }]
        : []
    return [
      ...comment,
      { label: str.edit.cut, icon: 'cut', hint: `${MOD}X`, sep: comment.length > 0, run: () => document.execCommand('cut') },
      { label: str.edit.copy, icon: 'copy', hint: `${MOD}C`, run: () => document.execCommand('copy') },
      {
        label: str.edit.paste,
        icon: 'paste',
        hint: `${MOD}V`,
        run: () =>
          void navigator.clipboard
            ?.readText()
            .then(pasteText)
            .catch(() => setToast(str.edit.clipboardBlocked(`${MOD}V`))),
      },
      { label: str.xlsx.rowsAbove(nr), icon: 'rowAbove', sep: true, run: () => rowsOp('rowAbove') },
      { label: str.xlsx.rowsBelow(nr), icon: 'rowBelow', run: () => rowsOp('rowBelow') },
      { label: str.xlsx.colsLeft(nc), icon: 'colLeft', run: () => rowsOp('colLeft') },
      { label: str.xlsx.colsRight(nc), icon: 'colRight', run: () => rowsOp('colRight') },
      { label: str.xlsx.deleteWhat(rows), icon: 'delRow', sep: true, run: () => rowsOp('delRow') },
      { label: str.xlsx.deleteWhat(cols), icon: 'delCol', run: () => rowsOp('delCol') },
      { label: str.xlsx.hideWhat(rows), icon: 'hide', sep: true, run: () => rowsOp('hideRows') },
      { label: str.xlsx.hideWhat(cols), icon: 'hide', run: () => rowsOp('hideCols') },
      { label: str.xlsx.unhideRows, icon: 'view', run: () => rowsOp('unhideRows') },
      { label: str.xlsx.unhideCols, icon: 'view', run: () => rowsOp('unhideCols') },
      { label: merged ? str.table.unmerge : str.table.merge, icon: 'merge', sep: true, run: toggleMerge, disabled: single && !merged },
      { label: (single ? str.xlsx.sortSheetAsc : str.xlsx.sortRangeAsc)(colName(sel.ac)), icon: 'sortAZ', run: () => sortBy(false) },
      { label: (single ? str.xlsx.sortSheetDesc : str.xlsx.sortRangeDesc)(colName(sel.ac)), icon: 'sortZA', run: () => sortBy(true) },
      { label: str.xlsx.fillDown, icon: 'down', hint: `${MOD}D`, run: () => void sheetOp({ op: 'fill', range: rect, dir: 'down' }), disabled: nr === 1 },
      { label: str.xlsx.fillRight, hint: `${MOD}R`, run: () => void sheetOp({ op: 'fill', range: rect, dir: 'right' }), disabled: nc === 1 },
      { label: str.xlsx.clearContents, icon: 'trash', hint: str.edit.delKey, sep: true, run: clearRange },
      { label: str.edit.clearFormat, icon: 'clear', hint: `${MOD}\\`, run: () => fmt({ clear: true }) },
      { label: str.xlsx.freezeAt(`${colName(sel.ac)}${sel.ar + 1}`), icon: 'freeze', sep: true, run: () => freeze('sel') },
      { label: str.xlsx.unfreeze, run: () => freeze('none'), disabled: !s.frozen.rows && !s.frozen.cols },
    ]
  }

  // ---------------------------------------------------------------- find & replace
  useEffect(() => {
    if (!find.open || !find.q) {
      setMatches([])
      return
    }
    let live = true
    const id = setTimeout(() => {
      client.call('xlsx.find', sheet, find.q, findOpts).then((m) => live && setMatches(m), () => {})
    }, 150)
    return () => {
      live = false
      clearTimeout(id)
    }
  }, [client, sheet, find.open, find.q, find.tick, findOpts, detailTick])
  const gotoMatch = (i: number) => {
    const m = matches[i]
    if (!m) return
    setFind((f) => ({ ...f, idx: i }))
    setSel({ ar: m.r, ac: m.c, fr: m.r, fc: m.c })
    scrollIntoView(m.r, m.c)
  }
  const findBoxes = find.open ? matches : []

  const onMouseMove = (e: React.MouseEvent) => {
    const pd = pointDrag.current
    if (pd) {
      const { r, c } = cellAt(e.clientX, e.clientY, e.currentTarget as HTMLElement)
      pointRange(pd.ar, pd.ac, r, c, pd.axis)
      return
    }
    const fill = fillRef.current
    if (fill) {
      const { r, c } = cellAt(e.clientX, e.clientY, e.currentTarget as HTMLElement)
      if (r >= 0 && c >= 0 && (fill.to?.r !== r || fill.to?.c !== c)) {
        fill.to = { r, c }
        setFillTo({ r, c })
      }
      return
    }
    if (!dragging.current) return
    const { r, c } = cellAt(e.clientX, e.clientY, e.currentTarget as HTMLElement)
    if (r >= 0 && c >= 0 && (r !== sel.fr || c !== sel.fc)) setSelState((p) => ({ ...p, fr: r, fc: c, whole: undefined }))
  }

  const box = edit && own ? cellBox(edit.r, edit.c) : null
  const fx = edit ? edit.value : shown(detail, sel.ar, sel.ac)
  const editKeys = (e: React.KeyboardEvent<HTMLTextAreaElement>) => {
    if (!edit || e.nativeEvent.isComposing || e.keyCode === 229) return
    const mod = e.metaKey || e.ctrlKey
    if (mod && e.key.toLowerCase() === 's' && !e.altKey) {
      // the cell is written, then the document: the browser's own "Save page as" stays closed
      e.preventDefault()
      commit()
      void lastRun.current.then(() => viewer.save())
      return
    }
    if (sugg.length) {
      if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
        e.preventDefault()
        const n = sugg.length
        setSugIdx((i) => (Math.min(i, n - 1) + (e.key === 'ArrowDown' ? 1 : n - 1)) % n)
        return
      }
      if (e.key === 'Tab' || (e.key === 'Enter' && !e.altKey && !e.shiftKey && !mod)) {
        e.preventDefault()
        accept(sugg[Math.min(sugIdx, sugg.length - 1)])
        return
      }
      if (e.key === 'Escape') {
        e.preventDefault()
        const at = word?.s
        setEdit((p) => (p ? { ...p, noSuggest: at } : p))
        return
      }
    }
    const arrow = ARROWS[e.key]
    if (arrow && !mod && !e.altKey && edit.mode === 'enter') {
      if (pointOk) {
        e.preventDefault()
        movePoint(arrow[0], arrow[1], e.shiftKey)
        return
      }
      if (inCell && !e.shiftKey) {
        e.preventDefault()
        tabStart.current = null
        commit(arrow)
        return
      }
    }
    if ((e.key === 'Home' || e.key === 'End') && !mod && !e.altKey) {
      // macOS scrolls the page on Home/End: move the caret instead
      e.preventDefault()
      const el = e.currentTarget
      const len = el.value.length
      if (e.key === 'Home') el.setSelectionRange(0, e.shiftKey ? (el.selectionEnd ?? 0) : 0, e.shiftKey ? 'backward' : 'none')
      else el.setSelectionRange(e.shiftKey ? (el.selectionStart ?? len) : len, len)
      readCaret(el)
    } else if (e.key === 'F2') {
      e.preventDefault()
      setEdit((p) => (p ? { ...p, mode: p.mode === 'enter' ? 'edit' : 'enter' } : p))
    } else if (e.key === 'F4') {
      e.preventDefault()
      cycleAbsolute()
    } else if (e.key === 'Enter' && mod && !e.altKey) {
      // Ctrl+Enter: the entry into every cell of the selection
      e.preventDefault()
      commitFill()
    } else if (e.key === 'Enter' && (e.altKey || e.shiftKey)) {
      // line break inside the cell: Alt+Enter as in Excel, Shift+Enter as in most text inputs
      e.preventDefault()
      const el = e.currentTarget
      // execCommand keeps the native undo history; setRangeText where it is gone
      if (!el.ownerDocument.execCommand?.('insertText', false, '\n')) {
        el.setRangeText('\n', el.selectionStart, el.selectionEnd, 'end')
        onInput(el)
      }
    } else if (e.key === 'Enter') {
      e.preventDefault()
      commit(enterMove(false, edit.c))
    } else if (e.key === 'Tab') {
      e.preventDefault()
      commit(tabMove(e.shiftKey, edit.c))
    } else if (e.key === 'Escape') {
      e.preventDefault()
      cancel()
    }
  }

  // autocomplete list / argument hint under the active input
  let assist: ReactNode = null
  if (edit && (sugg.length || hint)) {
    const slot = hint && fctx?.call ? argSlot(hint, fctx.call.arg) : -1
    assist = sugg.length ? (
      <div className="de-fassist" role="listbox" onMouseDown={(e) => e.preventDefault()}>
        {sugg.map((f, i) => {
          const on = i === Math.min(sugIdx, sugg.length - 1)
          return (
            <div key={f.name} role="option" aria-selected={on} className={on ? 'de-fsug on' : 'de-fsug'} onMouseEnter={() => setSugIdx(i)} onClick={() => accept(f)}>
              <span className="de-fsug-name">{f.name}</span>
              <span className="de-fsug-args">({[...f.args, ...(f.variadic ? ['...'] : [])].join(', ')})</span>
              {on && <div className="de-fsug-desc">{f.desc}</div>}
            </div>
          )
        })}
      </div>
    ) : (
      <div className="de-fassist de-fhint" onMouseDown={(e) => e.preventDefault()}>
        <div>
          {hint!.name}(
          {[...hint!.args, ...(hint!.variadic ? ['...'] : [])].map((a, i, all) => (
            <span key={i}>
              {i === slot ? <b>{a}</b> : a}
              {i < all.length - 1 ? ', ' : ''}
            </span>
          ))}
          )
        </div>
        <div className="de-fsug-desc">{hint!.desc}</div>
      </div>
    )
  }

  // in-cell editor geometry (shared by the textarea and its highlight layer)
  const edStyle: CSSProperties | undefined =
    edit && box
      ? {
          left: box.x - 1,
          top: box.y - 1,
          minWidth: box.w + 2,
          width: Math.max(box.w + 2, Math.min(480, edit.value.length * 8 + 16)),
          height: Math.max(box.h + 2, edit.value.split('\n').length * 18 + 4),
          ...(meta.baseFont as CSSProperties),
        }
      : undefined
  const hlParts: ReactNode[] = []
  if (edit && refSpans.length) {
    let at = 0
    for (const r of refSpans) {
      if (r.s > at) hlParts.push(edit.value.slice(at, r.s))
      hlParts.push(
        <span key={r.s} style={{ color: r.color }}>
          {edit.value.slice(r.s, r.e)}
        </span>,
      )
      at = r.e
    }
    hlParts.push(edit.value.slice(at) + '\u200b')
  }
  /** the cell is being edited in its own box (not in the formula bar, not on another sheet) */
  const editing = !!edit && inCell && !!edStyle
  const assistPos = (() => {
    if (!edStyle || !assist) return undefined
    const top = (edStyle.top as number) + (edStyle.height as number) + 2
    return top + 200 > size.h && (edStyle.top as number) > 200 ? { left: edStyle.left as number, bottom: size.h - (edStyle.top as number) + 2 } : { left: edStyle.left as number, top }
  })()

  return (
    <>
      <div className="de-fx">
        <div className="de-fx-ref">
          {colName(sel.ac)}
          {sel.ar + 1}
          {!oneCell && `:${colName(rect.c1)}${rect.r1 + 1}`}
        </div>
        {editable ? (
          <textarea
            ref={fxRef}
            className="de-fx-val de-fx-input"
            value={fx}
            rows={1}
            wrap="off"
            spellCheck={false}
            aria-label={str.xlsx.cellContent}
            onFocus={() => !edit && void startEdit(undefined, true)}
            onChange={(e) => onInput(e.target)}
            onSelect={(e) => readCaret(e.currentTarget)}
            onKeyDown={editKeys}
            onScroll={(e) => {
              // one line shows at a time: the caret's, whole (the browser scrolls just enough to reveal the caret)
              const el = e.currentTarget
              const lh = parseFloat(getComputedStyle(el).lineHeight) || el.clientHeight
              const top = Math.round(el.scrollTop / lh) * lh
              if (Math.abs(top - el.scrollTop) >= 1) el.scrollTop = top
            }}
            onBlur={(e) => {
              if (hold.current) return
              // back to the first line for the next cell
              e.currentTarget.scrollTop = 0
              if (edit && (edit.bar || !own)) commit()
            }}
          />
        ) : (
          <div className="de-fx-val" title={fx}>
            {fx}
          </div>
        )}
        {edit && !inCell && assist && <div className="de-fassist-at" style={{ left: 88, top: 30 }}>{assist}</div>}
      </div>
      <div
        ref={scrollRef}
        className="de-grid"
        tabIndex={0}
        onKeyDown={onKeyDown}
        onCopy={(e) => {
          if (!edit) copySel(e, false)
        }}
        onCut={(e) => {
          if (!edit && editable) copySel(e, true)
        }}
        onFocus={(e) => {
          // the grid itself (a click on its scroll bar) hands the keyboard to the hidden field that types into the cells
          if (e.target === e.currentTarget && editable && !edit) editRef.current?.focus({ preventScroll: true })
        }}
        onPaste={(e) => {
          if (edit || !editable) return
          e.preventDefault()
          // other programs may put only HTML on the clipboard (a table): its cells
          const html = e.clipboardData.getData('text/html')
          pasteText(e.clipboardData.getData('text/plain') || (html ? htmlToTsv(html) : ''))
        }}
      >
        <div style={{ position: 'relative', width: Math.max(spacerW, size.w), height: Math.max(spacerH, size.h) }}>
          <div
            className="de-grid-view"
            style={{ width: size.w, height: size.h, ...(meta.baseFont as CSSProperties) }}
            onMouseDown={onMouseDown}
            onMouseMove={onMouseMove}
            onContextMenu={(e) => {
              if (!editable) return
              e.preventDefault()
              const { r, c } = cellAt(e.clientX, e.clientY, e.currentTarget as HTMLElement)
              if (r >= 0 && c >= 0 && (r < rect.r0 || r > rect.r1 || c < rect.c0 || c > rect.c1)) setSel({ ar: r, ac: c, fr: r, fc: c })
              setCtx({ x: e.clientX, y: e.clientY })
            }}
            onDoubleClick={(e) => {
              if (edit) return
              const grip = (e.target as HTMLElement).closest('[data-grip-col],[data-grip-row]') as HTMLElement | null
              if (grip) void fitGrip(grip)
              else void startEdit()
            }}
          >
            {region('main', r0, r1, c0, c1, mainOx, mainOy, HW + fw, HH + fh, mainW, mainH)}
            {fr > 0 && region('top', 0, fr - 1, c0, c1, mainOx, 0, HW + fw, HH, mainW, fh)}
            {fc > 0 && region('left', r0, r1, 0, fc - 1, 0, mainOy, HW, HH + fh, fw, mainH)}
            {fr > 0 && fc > 0 && region('tl', 0, fr - 1, 0, fc - 1, 0, 0, HW, HH, fw, fh)}
            {colHeaders('ch', c0, c1, mainOx, HW + fw, mainW)}
            {fc > 0 && colHeaders('chf', 0, fc - 1, 0, HW, fw)}
            {rowHeaders('rh', r0, r1, mainOy, HH + fh, mainH)}
            {fr > 0 && rowHeaders('rhf', 0, fr - 1, 0, HH, fh)}
            {fr > 0 && <div className="de-freeze-h" style={{ left: 0, top: HH + fh - 1, width: size.w, height: 2 }} />}
            {fc > 0 && <div className="de-freeze-v" style={{ left: HW + fw - 1, top: 0, width: 2, height: size.h }} />}
            <div className="de-corner" style={{ width: HW, height: HH }} />
            {edit && inCell && edStyle && hlParts.length > 0 && (
              <div ref={hlRef} className="de-cell-edit de-cell-hl" style={edStyle} aria-hidden>
                {hlParts}
              </div>
            )}
            {/* The one text field of the grid, always there when the sheet can be edited: it is the cell editor while a cell is
                being edited, and else a hidden field that has the keyboard. Text typed, composed (IME), inserted (emoji picker,
                virtual keyboard) into it starts the edit of the selected cell, in this same field: a composition in progress
                is not cut when the editor opens. */}
            {editable && (
              <textarea
                ref={editRef}
                className={editing ? (hlParts.length ? 'de-cell-edit de-cell-edit-fx' : 'de-cell-edit') : 'de-cell-sink'}
                spellCheck={false}
                autoComplete="off"
                autoCapitalize="off"
                aria-label={str.xlsx.cell}
                // no on-screen keyboard until a cell is being edited (it would open as the sheet does)
                inputMode={edit ? undefined : 'none'}
                value={editing ? edit.value : ''}
                style={editing ? edStyle : SINK_STYLE}
                onMouseDown={(e) => e.stopPropagation()}
                onChange={(e) => (edit ? onInput(e.target) : sinkStart(e.target))}
                onSelect={(e) => readCaret(e.currentTarget)}
                onScroll={(e) => {
                  if (hlRef.current) hlRef.current.scrollTop = e.currentTarget.scrollTop
                }}
                onKeyDown={editKeys}
                onBlur={(e) => {
                  if (!edit || !inCell || hold.current) return
                  // moving into the formula bar keeps editing there
                  if (e.relatedTarget && e.relatedTarget === fxRef.current) return setEdit((p) => (p ? { ...p, bar: true, sync: 0 } : p))
                  commit()
                }}
              />
            )}
            {editing && assist && assistPos && (
              <div className="de-fassist-at" style={assistPos}>
                {assist}
              </div>
            )}
          </div>
        </div>
      </div>
      {editable && (
        <XlsxToolbar
          css={activeCss}
          base={meta.baseFont as { fontFamily?: string; fontSize?: string | number }}
          format={meta.formats[styleOf(sel.ar, sel.ac)] ?? 'General'}
          merged={merged}
          frozen={!!(s.frozen.rows || s.frozen.cols)}
          gridlines={s.gridLines}
          painter={painter !== null}
          onToggle={toggle}
          onFormat={fmt}
          onUndo={undo}
          onRedo={redo}
          onPaint={() => {
            setPainter((p) => (p === null ? styleOf(sel.ar, sel.ac) : null))
            focusGrid()
          }}
          onMerge={merge}
          onFreeze={freeze}
          onGridlines={() => void sheetOp({ op: 'gridlines', show: !s.gridLines })}
          onFunc={(fn) => (fn ? autoFn(fn) : void startEdit('='))}
          onSort={sortBy}
          onRows={rowsOp}
          onFind={() => setFind((f) => ({ ...f, open: true, tick: f.tick + 1 }))}
        />
      )}
      {find.open && (
        <FindBar
          query={find.q}
          index={find.idx}
          count={matches.length}
          editable={editable}
          replace={find.replace}
          placeholder={str.xlsx.findPlaceholder}
          onQuery={(q) => setFind((f) => ({ ...f, q, idx: 0 }))}
          onNav={(d) => matches.length && gotoMatch((find.idx + d + matches.length) % matches.length)}
          onClose={() => {
            setFind((f) => ({ ...f, open: false }))
            focusGrid()
          }}
          onReplace={(text, all) => {
            const m = matches[Math.min(find.idx, matches.length - 1)]
            if (!m) return
            void run(client.call('xlsx.replace', sheet, find.q, text, findOpts, all ? undefined : { r: m.r, c: m.c }))
          }}
          options={
            <>
              {(
                [
                  ['matchCase', str.xlsx.matchCase],
                  ['whole', str.xlsx.entireCell],
                  ['formulas', str.xlsx.inFormulas],
                ] as const
              ).map(([k, label]) => (
                <label key={k} className="de-find-opt">
                  <input type="checkbox" checked={!!findOpts[k]} onChange={(e) => setFindOpts((o) => ({ ...o, [k]: e.target.checked }))} />
                  {label}
                </label>
              ))}
            </>
          }
        />
      )}
      {toast && <Toast text={toast} onDone={() => setToast(null)} />}
      {ctx && <ContextMenu x={ctx.x} y={ctx.y} items={ctxItems()} onClose={() => setCtx(null)} />}
    </>
  )
}
