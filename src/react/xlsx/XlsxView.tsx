import { useEffect, useLayoutEffect, useMemo, useRef, useState, useSyncExternalStore, type CSSProperties, type MutableRefObject, type ReactNode } from 'react'
import type { BookOp, CellDetail, CellInput, CellOut, FindMatch, FindOptions, FmtPatch, Merge, Range, SheetOp, SheetProgress, WorkbookMeta, XlsxEditResult } from '../../core/xlsx/types'
import { colName, MAX_ROWS } from '../../core/xlsx/workbook'
import { detectSeparator, parseCsv } from '../../core/convert'
import { closeFormula, formulaContext, isFormula, rangeText, type FormulaContext } from '../../core/xlsx/formula/assist'
import { argSlot, functionInfo, suggestFunctions, type FnInfo } from '../../core/xlsx/formula/catalog'
import { MAX_COL, MAX_ROW } from '../../core/xlsx/formula/types'
import type { DocumentClient } from '../../core/worker/client'
import { getPeers, publishLocal, subscribePeers } from '../../collab/hosts'
import { claimFocus, useForceUpdate, useScroll, useSize } from '../common/hooks'
import { useFileOps, useViewer, useSelectAll } from '../common/Slots'
import { Status } from '../common/Status'
import { ContextMenu, MOD, type MenuItem } from '../common/ContextMenu'
import { Toast } from '../common/Dialog'
import { FindBar } from '../docx/FindBar'
import { XlsxToolbar, type MergeMode } from './XlsxToolbar'
import { Axis } from './axis'
import { fitNumber, fontPx } from './fit'
import { sheetPrintable } from './print'
import { useCommentsCtl, useCommentsState, type CommentsCtl, type PlacedThread } from '../comments/Comments'
import { COMMENT_KEYS } from '../comments/CommentsPanel'
import type { CommentThread } from '../../core/comments'

const BLOCK = 128
const HH = 22
const MAX_SCROLL_PX = 12_000_000

interface CachedBlock {
  rows: Map<number, CellOut[]>
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
  private listeners = new Set<() => void>()
  private off: () => void

  constructor(private client: DocumentClient) {
    const offs = [
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
      s.frozen = p.frozen ?? s.frozen
      s.gridLines = p.gridLines ?? true
    }
    for (const [r, px] of p.rowSizes) s.rowAxis.set(r, px)
    if (p.recalc) s.blocks.clear()
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
          for (const r of res.rows) rows.set(r.r, r.cells)
          s.blocks.set(b, { rows, complete: res.complete, rowsAt })
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
    s.cols = Math.max(s.cols, res.range.c1 + 1)
    this.invalidate(res.sheet, res.range.r0, res.range.r1)
  }

  row(sheet: number, r: number): CellOut[] | undefined {
    return this.sheets.get(sheet)?.blocks.get(Math.floor(r / BLOCK))?.rows.get(r)
  }

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

  const status = state.error
    ? `Error: ${state.error}`
    : `${state.rows.toLocaleString()} rows × ${state.cols.toLocaleString()} cols${state.done ? '' : ' · loading…'}${state.rows > MAX_ROWS ? ` · over Excel’s ${MAX_ROWS.toLocaleString()}-row limit: save as CSV` : ''}`

  const viewer = useViewer()
  useFileOps({
    printable: () => sheetPrintable(client, cur, store.sheet(cur), meta),
    text: (format) => client.call('xlsx.text', cur, format === 'tsv' ? '\t' : ','),
    stats: () => {
      const s = store.sheet(cur)
      return [
        ['Sheets', `${meta.sheets.length}${meta.sheets.some((x) => x.hidden) ? ` (${meta.sheets.filter((x) => x.hidden).length} hidden)` : ''}`],
        ['Current sheet', meta.sheets[cur]?.name ?? ''],
        ['Used range', `${s.rows.toLocaleString()} rows × ${s.cols.toLocaleString()} columns${s.done ? '' : ' (loading)'}`],
      ]
    },
    importText: async (file, name) => {
      const sep = /\.(tsv|tab)$/i.test((file as File).name ?? '') ? '\t' : undefined
      // a shared session: the rows reach the others as cell edits
      if (client.interceptor) return importRows(file, name, sep)
      // else the worker streams the file into the new sheet, like a sheet's XML (a million rows is fine)
      const res = await client.call('xlsx.import', cur, file, { name, sep })
      store.applied(res, meta)
      viewer.setDirty(true)
      setActive(cur + 1)
    },
  })
  /** the rows of a text file as cell edits of a new sheet (parsed here, sent through the session) */
  const importRows = async (file: Blob, name: string, sep?: string) => {
    const text = await file.text()
    const rows = parseCsv(text, sep ?? detectSeparator(text))
    const res = await client.call('xlsx.bookOp', { op: 'add', name, after: cur })
    store.applied(res, meta)
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
  const [tabMenu, setTabMenu] = useState<{ index: number; x: number; y: number } | null>(null)
  const bookOp = async (op: BookOp, goTo?: number) => {
    try {
      const res = await client.call('xlsx.bookOp', op)
      store.applied(res, meta)
      viewer.setDirty(true)
      if (goTo !== undefined) setActive(Math.max(0, Math.min(goTo, meta.sheets.length - 1)))
      bump()
    } catch (e) {
      alertError(e)
    }
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
      hint: 'Select a cell to comment on.',
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
      />
      <div className="de-tabs" role="tablist">
        {viewer.editable && (
          <button type="button" className="de-tab de-tab-add" title="Add sheet" onClick={() => void bookOp({ op: 'add', after: cur }, cur + 1)}>
            +
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
                if (e.target.value.trim() && e.target.value !== s.name) void bookOp({ op: 'rename', index: s.index, name: e.target.value })
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
            { label: 'Rename', run: () => setRenaming(tabMenu.index) },
            { label: 'Duplicate name → new sheet', run: () => void bookOp({ op: 'add', after: tabMenu.index, name: meta.sheets[tabMenu.index].name + ' (2)' }, tabMenu.index + 1) },
            { label: 'Move left', disabled: tabMenu.index === 0, run: () => void bookOp({ op: 'move', from: tabMenu.index, to: tabMenu.index - 1 }, tabMenu.index - 1) },
            { label: 'Move right', disabled: tabMenu.index >= meta.sheets.length - 1, run: () => void bookOp({ op: 'move', from: tabMenu.index, to: tabMenu.index + 1 }, tabMenu.index + 1) },
            { label: 'Delete sheet', disabled: meta.sheets.length <= 1, run: () => void bookOp({ op: 'delete', index: tabMenu.index }, Math.max(0, tabMenu.index - 1)) },
          ]}
        />
      )}
    </div>
  )
}

function alertError(e: unknown) {
  console.error('[office-engine]', e)
}

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
}

/** text rotation (OOXML degrees, + = counter-clockwise, 255 = stacked) */
function rotated(rot: number): CSSProperties {
  if (rot === 255) return { writingMode: 'vertical-rl', textOrientation: 'upright', whiteSpace: 'nowrap' }
  if (rot === 90) return { writingMode: 'vertical-rl', transform: 'rotate(180deg)', whiteSpace: 'nowrap' }
  if (rot === -90) return { writingMode: 'vertical-rl', whiteSpace: 'nowrap' }
  return { display: 'inline-block', transform: `rotate(${-rot}deg)`, whiteSpace: 'nowrap' }
}

const JUSTIFY: Record<string, string> = { center: 'center', right: 'flex-end', left: 'flex-start', justify: 'flex-start' }

interface Selection {
  ar: number
  ac: number
  fr: number
  fc: number
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

const ARROWS: Record<string, [number, number]> = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] }

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

function SheetGrid({ store, sheet, meta, client, edit, setEdit, hold, onSheet }: GridProps) {
  const viewer = useViewer()
  const editable = viewer.editable
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
  const fxRef = useRef<HTMLInputElement>(null)
  const [sugIdx, setSugIdx] = useState(0)
  const [ctx, setCtx] = useState<{ x: number; y: number } | null>(null)
  /** paint format: style index copied from the active cell, applied to the next selection */
  const [painter, setPainter] = useState<number | null>(null)
  const afterDrag = useRef<(() => void) | null>(null)
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

  const { rowAxis, colAxis, frozen } = s
  const HW = Math.max(40, 12 + 7 * String(Math.max(s.rows, 1000)).length)
  const viewW = Math.max(0, size.w - HW)
  const viewH = Math.max(0, size.h - HH)
  rowAxis.count = Math.max(s.rows + (s.done ? 20 : 1), Math.ceil(viewH / rowAxis.def) + 1, frozen.rows + 1, sel.fr + 2)
  colAxis.count = Math.max(s.cols + 3, 26, Math.ceil(viewW / Math.max(1, colAxis.def)) + 1, frozen.cols + 1, sel.fc + 2)
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
  const mergeAt = (r: number, c: number) => s.merges.find((m) => r >= m.r0 && r <= m.r1 && c >= m.c0 && c <= m.c1)
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
    if (store.focusGrid) {
      store.focusGrid = false
      el?.focus({ preventScroll: true })
    } else claimFocus(el ?? null) // keyboard first: a focused host page control keeps the focus
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
      const [c, raw, si, kind, color] = cell
      const cw = wOverride ?? colAxis.size(c)
      const ch = hOverride ?? rowAxis.size(r)
      if (!cw || !ch) return null
      const css = (styles[si] ?? {}) as Record<string, string | number>
      let text = raw
      if (kind === 1 && raw) text = fitNumber(raw, cw - 6, fontPx(css.fontSize, basePx))
      const { textAlign, textRotate, alignItems, whiteSpace, ...rest } = css
      const rot = Number(textRotate ?? 0)
      const justify = (textAlign && JUSTIFY[textAlign as string]) || (kind === 1 ? 'flex-end' : kind === 2 ? 'center' : 'flex-start')
      let ov = 0
      if (text && !rot && !whiteSpace && justify === 'flex-start' && !merged && filled) {
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
      if (color) style.color = color
      if (whiteSpace) style.whiteSpace = whiteSpace as CSSProperties['whiteSpace']
      // a merged cell hides the gridlines inside it (its own right / bottom edge lines stay visible)
      if (merged && !rest.background) style.background = 'linear-gradient(#fff, #fff) 0 0 / calc(100% - 1px) calc(100% - 1px) no-repeat'
      const overflow = ov > cw
      return (
        <div key={`${r}:${c}`} className={overflow ? 'de-cell de-ov' : 'de-cell'} style={style}>
          {text && <span style={rot ? rotated(rot) : overflow ? { maxWidth: ov - 6, overflow: 'hidden', flex: 'none' } : undefined}>{text}</span>}
        </div>
      )
    }
    for (const m of anchors) {
      const row = store.row(sheet, m.r0)
      const cell = row?.find((x) => x[0] === m.c0) ?? ([m.c0, '', 0, 0] as CellOut)
      out.push(cellEl(m.r0, cell, colAxis.pos(m.c1 + 1) - colAxis.pos(m.c0), rowAxis.pos(m.r1 + 1) - rowAxis.pos(m.r0), undefined, true))
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
      out.push(
        <div key={'peer' + p.site} className="de-peer-range" style={{ left: x0 - ox - 1, top: y0 - oy - 1, width: x1 - x0 + 1, height: y1 - y0 + 1, borderColor: p.color, background: p.color + '12' }}>
          {p.ar >= rr0 && p.ar <= rr1 && p.ac >= cc0 && p.ac <= cc1 && (
            <span className="de-peer-tag" style={{ background: p.color }}>
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

  const focusGrid = () => scrollRef.current?.focus({ preventScroll: true })
  const selectSheet = () => setSel({ ar: 0, ac: 0, fr: Math.max(0, s.rows - 1), fc: Math.max(0, s.cols - 1) })
  // Ctrl/⌘+A from the viewer chrome / the page; a cell being edited keeps its own select-all
  useSelectAll(() => {
    if (edit) return
    focusGrid()
    selectSheet()
  })
  const styleOf = (r: number, c: number) => store.row(sheet, r)?.find((x) => x[0] === c)?.[2] ?? 0

  const run = async (p: Promise<XlsxEditResult | null>) => {
    try {
      const res = await p
      store.applied(res, meta)
      if (res) viewer.setDirty(true)
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
    return d.value
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
    if (value !== e.initial) void run(client.call('xlsx.set', e.sheet, [{ r: e.r, c: e.c, input: value }]))
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

  const inputEl = (): HTMLTextAreaElement | HTMLInputElement | null => (inCell ? editRef.current : fxRef.current)

  // push focus + caret into the active input (start, pointing, autocomplete, sheet switch)
  useLayoutEffect(() => {
    if (!edit) return
    const el = inputEl()
    if (!el) return
    const root = el.getRootNode() as Document | ShadowRoot
    if (root.activeElement !== el) el.focus({ preventScroll: true })
    if (edit.sync) el.setSelectionRange(edit.caret, edit.caretEnd)
  }, [edit?.sync, inCell, edit !== null]) // eslint-disable-line react-hooks/exhaustive-deps

  const readCaret = (el: HTMLTextAreaElement | HTMLInputElement) => {
    const a = el.selectionStart ?? el.value.length
    const b = el.selectionEnd ?? a
    setEdit((p) => (p && (p.caret !== a || p.caretEnd !== b) ? { ...p, caret: a, caretEnd: b } : p))
  }

  const onInput = (el: HTMLTextAreaElement | HTMLInputElement) => {
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
    if (extend) setSelState({ ...sel, fr: r, fc: c })
    else setSel({ ar: r, ac: c, fr: r, fc: c })
    scrollIntoView(r, c)
  }

  const clearRange = () => {
    const cells = []
    const R = { ...rect, r1: Math.min(rect.r1, rect.r0 + 5000), c1: Math.min(rect.c1, rect.c0 + 200) }
    for (let r = R.r0; r <= R.r1; r++) {
      const row = store.row(sheet, r)
      for (let c = R.c0; c <= R.c1; c++) if (row?.some((x) => x[0] === c && x[1] !== '')) cells.push({ r, c, input: '' })
    }
    if (cells.length) void run(client.call('xlsx.set', sheet, cells))
  }

  const activeCss = (styles[styleOf(sel.ar, sel.ac)] ?? {}) as Record<string, string | number>
  const fmt = (patch: FmtPatch) => fmtAt(rect, patch)
  const fmtAt = (range: Range, patch: FmtPatch) => {
    void run(client.call('xlsx.format', sheet, range, patch))
    focusGrid()
  }
  const toggle = (k: 'b' | 'i' | 'u' | 'strike') => {
    const on =
      k === 'b' ? activeCss.fontWeight === 700 : k === 'i' ? activeCss.fontStyle === 'italic' : String(activeCss.textDecoration ?? '').includes(k === 'u' ? 'underline' : 'line-through')
    fmt({ [k]: !on })
  }
  const undo = () => void run(client.call('xlsx.undo'))
  const redo = () => void run(client.call('xlsx.redo'))

  const tsv = () => {
    const lines: string[] = []
    const R = { ...rect, r1: Math.min(rect.r1, rect.r0 + 20000) }
    for (let r = R.r0; r <= R.r1; r++) {
      const row = store.row(sheet, r)
      const vals: string[] = []
      for (let c = R.c0; c <= R.c1; c++) vals.push(row?.find((x) => x[0] === c)?.[1] ?? '')
      lines.push(vals.join('\t'))
    }
    return lines.join('\n')
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
      const [dr, dc] = moves[e.key]
      if (mod) moveBy(dr ? (dr > 0 ? Math.max(0, s.rows - 1 - (e.shiftKey ? sel.fr : sel.ar)) : -1e9) : 0, dc ? (dc > 0 ? Math.max(0, s.cols - 1 - (e.shiftKey ? sel.fc : sel.ac)) : -1e9) : 0, e.shiftKey)
      else moveBy(dr, dc, e.shiftKey)
      return
    }
    if (e.key === 'Enter' || e.key === 'Tab') {
      e.preventDefault()
      if (e.key === 'Enter') moveBy(e.shiftKey ? -1 : 1, 0, false)
      else moveBy(0, e.shiftKey ? -1 : 1, false)
      return
    }
    if (!editable) return
    if (e.key === 'F2') {
      e.preventDefault()
      void startEdit()
    } else if (e.key === 'Delete' || e.key === 'Backspace') {
      e.preventDefault()
      clearRange()
    } else if (e.key.length === 1 && !mod && !e.altKey) {
      e.preventDefault()
      void startEdit(e.key)
    }
  }

  const onMouseDown = (e: React.MouseEvent) => {
    if (e.button !== 0) return
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
      setSel({ ar: 0, ac: 0, fr: Math.max(0, s.rows - 1), fc: Math.max(0, s.cols - 1) })
      return
    }
    if (r < 0) setSel({ ar: 0, ac: c, fr: Math.max(0, s.rows - 1), fc: c })
    else if (c < 0) setSel({ ar: r, ac: 0, fr: r, fc: Math.max(0, s.cols - 1) })
    else if (e.shiftKey) setSelState({ ...sel, fr: r, fc: c })
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
  const toggleMerge = () => void sheetOp({ op: merged ? 'unmerge' : 'merge', range: rect })
  const merge = (mode: MergeMode) => void sheetOp(mode === 'unmerge' ? { op: 'unmerge', range: rect } : { op: 'merge', range: rect, mode })
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
  const sortBy = (desc: boolean) => {
    // one cell: sort the data below the frozen header by the active column
    const single = rect.r0 === rect.r1
    const range = single ? { r0: frozen.rows, c0: 0, r1: Math.max(frozen.rows, s.rows - 1), c1: Math.max(0, s.cols - 1) } : rect
    void sheetOp({ op: 'sort', range, col: sel.ac, desc })
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
  const pasteText = (raw: string) => {
    const text = raw.replace(/\r\n?/g, '\n').replace(/\n$/, '')
    const cells = text.split('\n').flatMap((line, i) => line.split('\t').map((v, j) => ({ r: sel.ar + i, c: sel.ac + j, input: v })))
    if (cells.length) void run(client.call('xlsx.set', sheet, cells.slice(0, 100_000)))
  }
  const ctxItems = (): MenuItem[] => {
    const nr = rect.r1 - rect.r0 + 1
    const nc = rect.c1 - rect.c0 + 1
    const rows = nr === 1 ? `row ${rect.r0 + 1}` : `rows ${rect.r0 + 1}–${rect.r1 + 1}`
    const cols = nc === 1 ? `column ${colName(rect.c0)}` : `columns ${colName(rect.c0)}–${colName(rect.c1)}`
    const single = nr === 1 && nc === 1
    const thread = cmAt(sel.ar, sel.ac)
    const comment: MenuItem[] = thread
      ? [{ label: 'Show comment', icon: 'comment', run: () => cm?.select(thread) }]
      : cm?.canComment
        ? [{ label: 'Comment', icon: 'commentAdd', hint: COMMENT_KEYS, run: () => cm.start() }]
        : []
    return [
      ...comment,
      { label: 'Cut', icon: 'cut', hint: `${MOD}X`, sep: comment.length > 0, run: () => document.execCommand('cut') },
      { label: 'Copy', icon: 'copy', hint: `${MOD}C`, run: () => document.execCommand('copy') },
      {
        label: 'Paste',
        icon: 'paste',
        hint: `${MOD}V`,
        run: () =>
          void navigator.clipboard
            ?.readText()
            .then(pasteText)
            .catch(() => setToast(`The browser blocked clipboard access — use ${MOD}V to paste`)),
      },
      { label: `Insert ${nr} row${nr > 1 ? 's' : ''} above`, icon: 'rowAbove', sep: true, run: () => rowsOp('rowAbove') },
      { label: `Insert ${nr} row${nr > 1 ? 's' : ''} below`, icon: 'rowBelow', run: () => rowsOp('rowBelow') },
      { label: `Insert ${nc} column${nc > 1 ? 's' : ''} left`, icon: 'colLeft', run: () => rowsOp('colLeft') },
      { label: `Insert ${nc} column${nc > 1 ? 's' : ''} right`, icon: 'colRight', run: () => rowsOp('colRight') },
      { label: `Delete ${rows}`, icon: 'delRow', sep: true, run: () => rowsOp('delRow') },
      { label: `Delete ${cols}`, icon: 'delCol', run: () => rowsOp('delCol') },
      { label: `Hide ${rows}`, icon: 'hide', sep: true, run: () => rowsOp('hideRows') },
      { label: `Hide ${cols}`, icon: 'hide', run: () => rowsOp('hideCols') },
      { label: 'Unhide rows', icon: 'view', run: () => rowsOp('unhideRows') },
      { label: 'Unhide columns', icon: 'view', run: () => rowsOp('unhideCols') },
      { label: merged ? 'Unmerge cells' : 'Merge cells', icon: 'merge', sep: true, run: toggleMerge, disabled: single && !merged },
      { label: `Sort ${single ? 'sheet' : 'range'} by ${colName(sel.ac)}, A → Z`, icon: 'sortAZ', run: () => sortBy(false) },
      { label: `Sort ${single ? 'sheet' : 'range'} by ${colName(sel.ac)}, Z → A`, icon: 'sortZA', run: () => sortBy(true) },
      { label: 'Fill down', icon: 'down', hint: `${MOD}D`, run: () => void sheetOp({ op: 'fill', range: rect, dir: 'down' }), disabled: nr === 1 },
      { label: 'Fill right', hint: `${MOD}R`, run: () => void sheetOp({ op: 'fill', range: rect, dir: 'right' }), disabled: nc === 1 },
      { label: 'Clear contents', icon: 'trash', hint: 'Del', sep: true, run: clearRange },
      { label: 'Clear formatting', icon: 'clear', hint: `${MOD}\\`, run: () => fmt({ clear: true }) },
      { label: `Freeze up to ${colName(sel.ac)}${sel.ar + 1}`, icon: 'freeze', sep: true, run: () => freeze('sel') },
      { label: 'Unfreeze', run: () => freeze('none'), disabled: !frozen.rows && !frozen.cols },
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
    if (!dragging.current) return
    const { r, c } = cellAt(e.clientX, e.clientY, e.currentTarget as HTMLElement)
    if (r >= 0 && c >= 0 && (r !== sel.fr || c !== sel.fc)) setSelState((p) => ({ ...p, fr: r, fc: c }))
  }

  const box = edit && own ? cellBox(edit.r, edit.c) : null
  const fx = edit ? edit.value : shown(detail, sel.ar, sel.ac)
  const editKeys = (e: React.KeyboardEvent<HTMLTextAreaElement | HTMLInputElement>) => {
    if (!edit || e.nativeEvent.isComposing) return
    const mod = e.metaKey || e.ctrlKey
    if (sugg.length) {
      if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
        e.preventDefault()
        const n = sugg.length
        setSugIdx((i) => (Math.min(i, n - 1) + (e.key === 'ArrowDown' ? 1 : n - 1)) % n)
        return
      }
      if (e.key === 'Tab' || (e.key === 'Enter' && !e.altKey && !e.shiftKey)) {
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
    } else if (e.key === 'Enter' && !e.altKey) {
      e.preventDefault()
      commit([e.shiftKey ? -1 : 1, 0])
    } else if (e.key === 'Tab') {
      e.preventDefault()
      commit([0, e.shiftKey ? -1 : 1])
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
          <input
            ref={fxRef}
            className="de-fx-val de-fx-input"
            value={fx}
            spellCheck={false}
            aria-label="Cell content"
            onFocus={() => !edit && void startEdit(undefined, true)}
            onChange={(e) => onInput(e.target)}
            onSelect={(e) => readCaret(e.currentTarget)}
            onKeyDown={editKeys}
            onBlur={() => {
              if (hold.current) return
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
          if (edit) return
          e.preventDefault()
          e.clipboardData.setData('text/plain', tsv())
        }}
        onCut={(e) => {
          if (edit || !editable) return
          e.preventDefault()
          e.clipboardData.setData('text/plain', tsv())
          clearRange()
        }}
        onPaste={(e) => {
          if (edit || !editable) return
          e.preventDefault()
          pasteText(e.clipboardData.getData('text/plain'))
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
            onDoubleClick={() => !edit && void startEdit()}
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
            {edit && inCell && edStyle && (
              <>
                {hlParts.length > 0 && (
                  <div ref={hlRef} className="de-cell-edit de-cell-hl" style={edStyle} aria-hidden>
                    {hlParts}
                  </div>
                )}
                <textarea
                  ref={editRef}
                  className={hlParts.length ? 'de-cell-edit de-cell-edit-fx' : 'de-cell-edit'}
                  spellCheck={false}
                  value={edit.value}
                  style={edStyle}
                  onMouseDown={(e) => e.stopPropagation()}
                  onChange={(e) => onInput(e.target)}
                  onSelect={(e) => readCaret(e.currentTarget)}
                  onScroll={(e) => {
                    if (hlRef.current) hlRef.current.scrollTop = e.currentTarget.scrollTop
                  }}
                  onKeyDown={editKeys}
                  onBlur={(e) => {
                    if (hold.current) return
                    // moving into the formula bar keeps editing there
                    if (e.relatedTarget && e.relatedTarget === fxRef.current) return setEdit((p) => (p ? { ...p, bar: true, sync: 0 } : p))
                    commit()
                  }}
                />
                {assist && assistPos && (
                  <div className="de-fassist-at" style={assistPos}>
                    {assist}
                  </div>
                )}
              </>
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
          frozen={!!(frozen.rows || frozen.cols)}
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
          placeholder="Find in sheet"
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
                  ['matchCase', 'Match case'],
                  ['whole', 'Entire cell'],
                  ['formulas', 'In formulas'],
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
