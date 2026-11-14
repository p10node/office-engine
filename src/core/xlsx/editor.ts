/**
 * XLSX editing: cell values/formats live in the worker's SheetData; on save each
 * edited sheet is streamed through a rewriter that only touches edited rows.
 */
import { relOfType, relsPathFor } from '../opc'
import type { DocComment } from '../comments'
import { writeXlsxComments } from './comments'
import {
  clone, el, escText, relIdAttrs, getAttr, isNode, kidOf, kidsOf, local, parseRaw, prefixOf, rootEl, serialize, setAttr, setOrdered,
  type RChild, type RNode,
} from '../raw'
import { parseXml } from '../xml'
import { BUILTIN_FORMATS, isDateFormat } from './numfmt'
import { parseSheetStyles } from './styles'
import type { BookOp, BorderKind, CellInput, ColInfo, FindOptions, FmtPatch, Merge, Range, SheetOp, SheetProgress, XlsxEditResult } from './types'
import { materializeShared, restore, setColWidth, setMerge, setRowHeight, shiftSheet, snapshot, type Snapshot } from './structure'
import { renameSheetInFormula, translateFormula } from './formula'
import { loadDelimited, type DelimitedImport } from './import'
import {
  colName, MAX_COL, MAX_ROWS, parseRef, rowSizesOf, SheetData, T_BOOL, T_DATE, T_EMPTY, T_ERR, T_NUM, T_SST, T_STR, XlsxWorkbook,
} from './workbook'

export const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'

type Entry = [number, unknown, number]
interface CellChange {
  key: number
  prev: Entry | null
  prevF?: string
  next: Entry | null
  nextF?: string
}
interface LayoutState {
  merges?: Merge[]
  colInfo?: ColInfo[]
  rowAttrs?: [number, { ht?: number; hidden?: boolean }][]
  frozen?: { rows: number; cols: number }
  gridLines?: boolean
}

export interface Change {
  sheet: number
  cells: CellChange[]
  range: Range
  /** cheap layout state (merges, sizes, panes) before/after */
  lay?: { before: LayoutState; after: LayoutState }
  /** whole-sheet snapshots (row/column insert/delete) */
  full?: { before: Snapshot; after: Snapshot }
  /** comments the change moved or took away (inserted / deleted rows and columns): their values before / after */
  cm?: { before: [string, DocComment | null][]; after: [string, DocComment | null][] }
}

const FONT_ORDER = ['b', 'i', 'strike', 'condense', 'extend', 'outline', 'shadow', 'u', 'vertAlign', 'sz', 'color', 'name', 'family', 'charset', 'scheme']
const WB_ORDER = [
  'fileVersion', 'fileSharing', 'workbookPr', 'workbookProtection', 'bookViews', 'sheets', 'functionGroups', 'externalReferences',
  'definedNames', 'calcPr', 'oleSize', 'customWorkbookViews', 'pivotCaches', 'smartTagPr', 'smartTagTypes', 'webPublishing',
  'fileRecoveryPr', 'webPublishObjects', 'extLst',
]
const SS_ORDER = ['numFmts', 'fonts', 'fills', 'borders', 'cellStyleXfs', 'cellXfs', 'cellStyles', 'dxfs', 'tableStyles', 'colors', 'extLst']

// ---------------------------------------------------------------- cell storage helpers

function rowOf(d: SheetData, r: number, create: boolean): unknown[] | undefined {
  let row = d.rows[r]
  if (!row && create) d.rows[r] = row = []
  return row
}

export function getEntry(d: SheetData, r: number, c: number): Entry | null {
  const row = d.rows[r]
  if (!row) return null
  for (let k = 0; k < row.length; k += 4) if (row[k] === c) return [row[k + 1] as number, row[k + 2], row[k + 3] as number]
  return null
}

function setEntry(d: SheetData, r: number, c: number, e: Entry | null) {
  // the file does not have this row's edit: its page stays in memory (and must be there: a fresh row would lose the rest)
  if (d.pages && !d.pages.has(r)) throw new Error('Row not in memory')
  d.pages?.pin(r)
  const row = rowOf(d, r, !!e)
  if (!row) return
  let k = 0
  for (; k < row.length; k += 4) if ((row[k] as number) >= c) break
  const exists = k < row.length && row[k] === c
  if (!e) {
    if (exists) row.splice(k, 4)
    if (!row.length) d.rows[r] = undefined
    return
  }
  if (exists) {
    row[k + 1] = e[0]
    row[k + 2] = e[1]
    row[k + 3] = e[2]
  } else row.splice(k, 0, c, e[0], e[1], e[2])
  if (r > d.maxRow) d.maxRow = r
  if (c > d.maxCol) d.maxCol = c
}

function dateSerial(t: string): number | undefined {
  let m = /^(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?$/.exec(t)
  let y: number, mo: number, d: number
  if (m) [y, mo, d] = [Number(m[1]), Number(m[2]), Number(m[3])]
  else {
    m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})(?: (\d{1,2}):(\d{2})(?::(\d{2}))?)?$/.exec(t)
    if (!m) return undefined
    ;[d, mo, y] = [Number(m[1]), Number(m[2]), Number(m[3])]
  }
  const ms = Date.UTC(y, mo - 1, d, Number(m[4] ?? 0), Number(m[5] ?? 0), Number(m[6] ?? 0))
  if (!Number.isFinite(ms)) return undefined
  return ms / 86400000 + 25569
}

function parseInput(input: string, style: number, isDate = false): { e: Entry | null; f?: string } {
  if (input === '') return { e: style ? [T_EMPTY, '', style] : null }
  if (input.startsWith('=') && input.length > 1) return { e: [T_STR, input, style], f: input.slice(1) }
  const t = input.trim()
  if (/^[-+]?(\d+\.?\d*|\.\d+)(e[-+]?\d+)?$/i.test(t)) return { e: [T_NUM, Number(t), style] }
  if (/^[-+]?(\d+\.?\d*|\.\d+)%$/.test(t)) return { e: [T_NUM, Number(t.slice(0, -1)) / 100, style] }
  if (/^(true|false)$/i.test(t)) return { e: [T_BOOL, /^true$/i.test(t) ? 1 : 0, style] }
  if (isDate) {
    const serial = dateSerial(t)
    if (serial !== undefined) return { e: [T_NUM, serial, style] }
  }
  return { e: [T_STR, input, style] }
}

// ---------------------------------------------------------------- styles

class StyleEditor {
  nodes: RChild[]
  root: RNode
  P: string
  private cache = new Map<string, number>()
  /**
   * per section: serialized entry → first index. Derived entries reuse identical ones, so derivation
   * depends only on the table's content — a saved and reopened workbook derives the same indexes
   * as the live one (collaboration snapshots rely on it).
   */
  private found = new Map<string, Map<string, number>>()
  changed = false

  constructor(xml: string) {
    this.nodes = parseRaw(xml)
    this.root = rootEl(this.nodes)!
    this.P = prefixOf(this.root.q)
  }

  private section(name: string): RNode {
    let n = kidOf(this.root, name)
    if (!n) {
      n = el(this.P + name, { count: '0' })
      setOrdered(this.root, n, name, SS_ORDER)
    }
    return n
  }

  private append(sectionName: string, node: RNode): number {
    const sec = this.section(sectionName)
    sec.c.push(node)
    sec.self = false
    const n = kidsOf(sec).length
    setAttr(sec, 'count', String(n))
    this.changed = true
    return n - 1
  }

  private lookup(sectionName: string): Map<string, number> {
    let m = this.found.get(sectionName)
    if (!m) {
      const idx = new Map<string, number>()
      kidsOf(kidOf(this.root, sectionName)).forEach((k, i) => {
        const s = serialize([k])
        if (!idx.has(s)) idx.set(s, i)
      })
      this.found.set(sectionName, (m = idx))
    }
    return m
  }

  /** index of an entry equal to `node`, appended when there is none */
  private intern(sectionName: string, node: RNode): number {
    const s = serialize([node])
    const m = this.lookup(sectionName)
    const hit = m.get(s)
    if (hit !== undefined) return hit
    const i = this.append(sectionName, node)
    m.set(s, i)
    return i
  }

  private numFmtId(code: string): number {
    for (const [id, c] of Object.entries(BUILTIN_FORMATS)) if (c === code) return Number(id)
    const sec = kidOf(this.root, 'numFmts')
    let max = 163
    for (const f of kidsOf(sec, 'numFmt')) {
      const id = Number(getAttr(f, 'numFmtId'))
      if (getAttr(f, 'formatCode') === code) return id
      max = Math.max(max, id)
    }
    this.append('numFmts', el(this.P + 'numFmt', { numFmtId: String(max + 1), formatCode: code }))
    return max + 1
  }

  derive(xfIndex: number, patch: FmtPatch): number {
    const key = xfIndex + JSON.stringify(patch)
    const hit = this.cache.get(key)
    if (hit !== undefined) return hit
    const P = this.P
    const xfs = kidsOf(this.section('cellXfs'), 'xf')
    const xf = clone(xfs[xfIndex] ?? xfs[0] ?? el(P + 'xf', { numFmtId: '0', fontId: '0', fillId: '0', borderId: '0' }))
    const fontKeys = ['b', 'i', 'u', 'strike', 'color', 'sz', 'font'] as const
    if (fontKeys.some((k) => k in patch)) {
      const fonts = kidsOf(this.section('fonts'), 'font')
      const font = clone(fonts[Number(getAttr(xf, 'fontId') ?? 0)] ?? fonts[0] ?? el(P + 'font'))
      const put = (l: string, n: RNode | undefined) => setOrdered(font, n, l, FONT_ORDER)
      if (patch.b !== undefined) put('b', patch.b ? el(P + 'b') : undefined)
      if (patch.i !== undefined) put('i', patch.i ? el(P + 'i') : undefined)
      if (patch.strike !== undefined) put('strike', patch.strike ? el(P + 'strike') : undefined)
      if (patch.u !== undefined) put('u', patch.u ? el(P + 'u') : undefined)
      if (patch.sz !== undefined) put('sz', el(P + 'sz', { val: String(patch.sz) }))
      if ('color' in patch) put('color', patch.color ? el(P + 'color', { rgb: 'FF' + patch.color.replace('#', '').toUpperCase() }) : undefined)
      if (patch.font) {
        put('name', el(P + 'name', { val: patch.font }))
        put('scheme', undefined)
      }
      setAttr(xf, 'fontId', String(this.intern('fonts', font)))
      setAttr(xf, 'applyFont', '1')
    }
    if ('fill' in patch) {
      if (!patch.fill) setAttr(xf, 'fillId', '0')
      else {
        const fill = parseRaw(
          `<${P}fill><${P}patternFill patternType="solid"><${P}fgColor rgb="FF${patch.fill.replace('#', '').toUpperCase()}"/><${P}bgColor indexed="64"/></${P}patternFill></${P}fill>`,
        )[0] as RNode
        setAttr(xf, 'fillId', String(this.intern('fills', fill)))
      }
      setAttr(xf, 'applyFill', '1')
    }
    if (patch.border) {
      const sides = (patch as FmtPatch & { sides?: Record<string, boolean> }).sides ?? {}
      const borders = kidsOf(this.section('borders'), 'border')
      const b = clone(borders[Number(getAttr(xf, 'borderId') ?? 0)] ?? borders[0] ?? el(P + 'border'))
      const color = (patch.borderColor ?? '#000000').replace('#', '').toUpperCase()
      const order = ['left', 'start', 'right', 'end', 'top', 'bottom', 'diagonal', 'vertical', 'horizontal']
      for (const side of ['left', 'right', 'top', 'bottom']) {
        if (!(side in sides)) continue
        const node = sides[side] ? el(P + side, { style: patch.borderStyle ?? 'thin' }, [el(P + 'color', { rgb: 'FF' + color })]) : el(P + side)
        setOrdered(b, node, side, order)
      }
      if (!kidOf(b, 'diagonal')) setOrdered(b, el(P + 'diagonal'), 'diagonal', order)
      setAttr(xf, 'borderId', String(this.intern('borders', b)))
      setAttr(xf, 'applyBorder', '1')
    }
    if ('align' in patch || 'wrap' in patch || 'valign' in patch || 'rotate' in patch) {
      const al = clone(kidOf(xf, 'alignment') ?? el(P + 'alignment'))
      if ('align' in patch) setAttr(al, 'horizontal', patch.align ?? undefined)
      if ('valign' in patch) setAttr(al, 'vertical', patch.valign && patch.valign !== 'bottom' ? patch.valign : undefined)
      if ('wrap' in patch) setAttr(al, 'wrapText', patch.wrap ? '1' : undefined)
      if ('rotate' in patch) {
        const v = patch.rotate
        // OOXML: 1-90 counter-clockwise, 91-180 clockwise (90 + deg), 255 vertical
        setAttr(al, 'textRotation', !v ? undefined : v === 255 ? '255' : v > 0 ? String(Math.min(90, v)) : String(90 + Math.min(90, -v)))
      }
      setOrdered(xf, al.a.trim() ? al : undefined, 'alignment', ['alignment', 'protection', 'extLst'])
      setAttr(xf, 'applyAlignment', '1')
    }
    if (patch.numFmt !== undefined) {
      setAttr(xf, 'numFmtId', String(this.numFmtId(patch.numFmt)))
      setAttr(xf, 'applyNumberFormat', '1')
    }
    const idx = this.intern('cellXfs', xf)
    this.cache.set(key, idx)
    return idx
  }

  xml() {
    return serialize(this.nodes)
  }

  /** collaboration: sizes of the growing sections (to drop what a rolled-back operation derived) */
  mark(): Record<string, number> {
    const out: Record<string, number> = {}
    for (const name of ['numFmts', 'fonts', 'fills', 'borders', 'cellXfs']) out[name] = kidsOf(kidOf(this.root, name)).length
    return out
  }

  truncate(m: Record<string, number>) {
    for (const [name, len] of Object.entries(m)) {
      const sec = kidOf(this.root, name)
      const kids = kidsOf(sec)
      if (!sec || kids.length <= len) continue
      const drop = new Set<RChild>(kids.slice(len))
      sec.c = sec.c.filter((c) => !drop.has(c))
      setAttr(sec, 'count', String(len))
      const idx = this.found.get(name)
      if (idx) for (const [k, v] of idx) if (v >= len) idx.delete(k)
    }
    for (const [k, v] of this.cache) if (v >= m.cellXfs) this.cache.delete(k)
  }
}

// ---------------------------------------------------------------- editor

export class XlsxEditor {
  private undoStack: Change[] = []
  private redoStack: Change[] = []
  private styleEd?: StyleEditor
  private stylesPath?: string
  private formulaTouched = false

  constructor(private wb: XlsxWorkbook) {}

  get dirty() {
    return [...this.wb.sheetsData.values()].some(needsWrite) || !!this.styleEd?.changed || this.bookChanged || this.wb.comments.dirty
  }

  /**
   * Comments of sheet `path` after `n` rows / columns were inserted (n > 0) or deleted (n < 0) at `at`: they move with
   * their cell, and go with it. Returns what changed (for undo).
   */
  private shiftComments(path: string, axis: 'row' | 'col', at: number, n: number): Change['cm'] {
    const store = this.wb.comments
    const before: [string, DocComment | null][] = []
    const after: [string, DocComment | null][] = []
    const set = (id: string, c: DocComment | null) => {
      before.push([id, store.get(id) ?? null])
      after.push([id, c])
      store.set(id, c)
    }
    for (const t of store.threads()) {
      const a = t.root.at
      if (!a || !('sheet' in a) || a.sheet !== path) continue
      const x = axis === 'row' ? a.r : a.c
      const y = n >= 0 ? (x >= at ? x + n : x) : x < at ? x : x < at - n ? null : x + n
      if (y === x) continue
      if (y === null) {
        for (const r of t.replies) set(r.id, null)
        set(t.root.id, null)
      } else set(t.root.id, { ...t.root, at: axis === 'row' ? { ...a, r: y } : { ...a, c: y } })
    }
    return before.length ? { before, after } : undefined
  }

  private sheet(i: number): SheetData {
    const d = this.wb.sheetsData.get(i)
    if (!d || !d.done) throw new Error('Sheet is still loading')
    return d
  }

  private lastRecalc: number[] = []

  private apply(d: SheetData, cells: CellChange[], which: 'prev' | 'next') {
    const sheet = [...this.wb.sheetsData].find(([, x]) => x === d)?.[0] ?? 0
    const formulaKeys: { sheet: number; key: number }[] = []
    this.applyRaw(d, cells, which)
    for (const ch of cells) {
      const f0 = which === 'prev' ? ch.nextF : ch.prevF
      const f1 = which === 'prev' ? ch.prevF : ch.nextF
      if (f0 !== f1) this.wb.calc.formulaChanged(sheet, ch.key)
      if (f1) formulaKeys.push({ sheet, key: ch.key })
    }
    // only value changes matter for dependents (formatting doesn't)
    const changed = cells.filter((c) => {
      if (which === 'prev' ? c.prevF : c.nextF) return false
      const a = c.prev
      const b = c.next
      return !a || !b || a[0] !== b[0] || a[1] !== b[1] || c.prevF !== c.nextF
    })
    this.lastRecalc = changed.length || formulaKeys.length ? [...this.wb.calc.recalc(changed.map((c) => ({ sheet, key: c.key })), formulaKeys)] : []
  }

  private applyRaw(d: SheetData, cells: CellChange[], which: 'prev' | 'next') {
    for (const ch of cells) {
      const r = Math.floor(ch.key / MAX_COL)
      const c = ch.key % MAX_COL
      setEntry(d, r, c, which === 'prev' ? ch.prev : ch.next)
      const f = which === 'prev' ? ch.prevF : ch.nextF
      if (f) d.formulas.set(ch.key, f)
      else d.formulas.delete(ch.key)
      d.edited.add(ch.key)
      if (ch.prevF || ch.nextF) this.formulaTouched = true
    }
  }

  private record(sheet: number, cells: CellChange[], range: Range, lay?: Change['lay']): XlsxEditResult {
    this.undoStack.push({ sheet, cells, range, lay })
    if (this.undoStack.length > 200) this.undoStack.shift()
    this.redoStack = []
    return { sheet, range, recalc: this.lastRecalc, layout: lay ? this.layout(sheet) : undefined }
  }

  /** after rows/columns moved: rebuild the dependency graph and recompute every formula */
  private recalcAll(): number[] {
    this.wb.calc.invalidate()
    const all: { sheet: number; key: number }[] = []
    for (const [sheet, d] of this.wb.sheetsData) for (const key of d.formulas.keys()) all.push({ sheet, key })
    return [...this.wb.calc.recalc([], all)]
  }

  async setCells(sheet: number, inputs: CellInput[]): Promise<XlsxEditResult> {
    const d = this.sheet(sheet)
    await this.prepare(sheet, d, inputs)
    const changes: CellChange[] = []
    const range = { r0: Infinity, c0: Infinity, r1: -1, c1: -1 }
    // text with line breaks: the cell wraps and its row grows to show every line (as Excel does)
    const xfs = this.wb.styles.xfs
    let ed: StyleEditor | undefined
    let rowsBefore: LayoutState['rowAttrs']
    for (const { r, c, input } of inputs) {
      const key = r * MAX_COL + c
      const prev = getEntry(d, r, c)
      const st = prev?.[2] ?? 0
      const { e, f } = parseInput(input, st, isDateFormat(xfs[st]?.numFmt ?? ''))
      if (e && e[0] === T_STR && !f && input.includes('\n')) {
        if (!xfs[st]?.wrap) e[2] = (ed ??= await this.styles()).derive(st, { wrap: true })
        const a = d.rowAttrs.get(r)
        const now = a?.ht !== undefined ? Math.round((a.ht * 4) / 3) : (d.autoRowH.get(r) ?? d.defaultRowH)
        const need = input.split('\n').length * Math.max(d.defaultRowH, Math.round((xfs[st]?.fontPt ?? 11) * 1.6 + 2))
        if (need > now && !a?.hidden && !d.merges.some((m) => m.r0 <= r && r <= m.r1 && m.c0 <= c && c <= m.c1)) {
          rowsBefore ??= layoutState(d).rowAttrs
          setRowHeight(d, r, r, need)
        }
      }
      changes.push({ key, prev, prevF: d.formulas.get(key), next: e, nextF: f })
      range.r0 = Math.min(range.r0, r)
      range.c0 = Math.min(range.c0, c)
      range.r1 = Math.max(range.r1, r)
      range.c1 = Math.max(range.c1, c)
    }
    this.apply(d, changes, 'next')
    const styles = ed ? this.restyled(ed, xfs.length) : undefined
    const res = this.record(sheet, changes, range, rowsBefore && { before: { rowAttrs: rowsBefore }, after: { rowAttrs: layoutState(d).rowAttrs } })
    if (styles?.length) res.styles = styles
    return res
  }

  /** the style table grew: stores it, re-derives its css and returns the entries added after the first `before` */
  private restyled(ed: StyleEditor, before: number): NonNullable<XlsxEditResult['styles']> {
    const xml = ed.xml()
    this.wb.pkg.setPart(this.stylesPath!, xml)
    this.wb.styles = parseSheetStyles(parseXml(xml), this.wb.theme)
    return this.wb.styles.xfs.slice(before).map((x, i) => ({ index: before + i, css: x.css, format: x.numFmt }))
  }

  /**
   * Sheets read from their file by pages: the rows an edit touches, and the rows read by the formulas it adds or
   * whose value it changes, are in memory first (throws when they cannot all be).
   */
  private async prepare(sheet: number, d: SheetData, inputs: CellInput[]): Promise<void> {
    let paged = false
    for (const x of this.wb.sheetsData.values()) if (x.pages) paged = true
    if (!paged) return
    // the rows edited, in runs (an edit far down and one at the top are two)
    const rows = [...new Set(inputs.map((i) => i.r))].sort((a, b) => a - b)
    const gap = d.pages?.size ?? 1
    for (let i = 0; i < rows.length; ) {
      let j = i
      while (j + 1 < rows.length && rows[j + 1] - rows[j] <= gap) j++
      await this.wb.ensure(d, rows[i], rows[j])
      i = j + 1
    }
    const changed = inputs.map((i) => ({ sheet, key: i.r * MAX_COL + i.c }))
    const added = inputs.filter((i) => i.input.length > 1 && i.input.startsWith('=')).map((i) => ({ sheet, key: i.r * MAX_COL + i.c, text: i.input.slice(1) }))
    for (const need of this.wb.calc.reads(changed, added)) {
      const x = this.wb.sheetsData.get(need.sheet)
      if (x?.pages) await this.wb.ensure(x, need.r0, need.r1)
    }
    // later rows may have pushed earlier ones out again
    if (d.pages) for (const r of rows) if (!d.pages.has(r)) throw new Error(`More rows at once than a sheet of ${(d.maxRow + 1).toLocaleString()} rows can hold in memory`)
  }

  private async styles(): Promise<StyleEditor> {
    if (!this.styleEd) {
      const rel = relOfType(this.wb.rels, 'styles')
      this.stylesPath = rel?.target ?? 'xl/styles.xml'
      const xml = this.wb.pkg.has(this.stylesPath)
        ? await this.wb.pkg.text(this.stylesPath)
        : '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs></styleSheet>'
      this.styleEd = new StyleEditor(xml)
    }
    return this.styleEd
  }

  async format(sheet: number, range: Range, patch: FmtPatch): Promise<XlsxEditResult> {
    const d = this.sheet(sheet)
    const ed = await this.styles()
    const r1 = Math.min(range.r1, range.r0 + 200_000)
    await this.wb.ensure(d, range.r0, r1)
    const before = this.wb.styles.xfs.length
    const changes: CellChange[] = []
    let budget = 300_000
    for (let r = range.r0; r <= r1 && budget > 0; r++) {
      for (let c = range.c0; c <= range.c1 && budget > 0; c++, budget--) {
        const key = r * MAX_COL + c
        const prev = getEntry(d, r, c)
        const cellPatch = patch.border ? { ...patch, sides: borderSides(patch.border, range, r, c) } : patch
        if (patch.border && !Object.keys((cellPatch as { sides: object }).sides).length) continue
        const s = patch.clear ? 0 : patch.xf !== undefined ? patch.xf : ed.derive(prev?.[2] ?? 0, cellPatch)
        const next: Entry = prev ? [prev[0], prev[1], s] : [T_EMPTY, '', s]
        const f = d.formulas.get(key)
        changes.push({ key, prev, prevF: f, next, nextF: f })
      }
    }
    this.apply(d, changes, 'next')
    const res = this.record(sheet, changes, range)
    res.styles = this.restyled(ed, before)
    return res
  }

  private step(ch: Change, which: 'prev' | 'next'): XlsxEditResult {
    const d = this.sheet(ch.sheet)
    this.lastRecalc = []
    if (ch.full) restore(d, which === 'prev' ? ch.full.before : ch.full.after)
    if (ch.cm) for (const [id, c] of which === 'prev' ? [...ch.cm.before].reverse() : ch.cm.after) this.wb.comments.set(id, c)
    this.apply(d, ch.cells, which)
    if (ch.lay) applyLayout(d, which === 'prev' ? ch.lay.before : ch.lay.after)
    const recalc = ch.full ? this.recalcAll() : this.lastRecalc
    return { sheet: ch.sheet, range: ch.range, layout: ch.full || ch.lay ? this.layout(ch.sheet) : undefined, recalc }
  }

  undo(): XlsxEditResult | null {
    const ch = this.undoStack.pop()
    if (!ch) return null
    this.redoStack.push(ch)
    return this.step(ch, 'prev')
  }

  redo(): XlsxEditResult | null {
    const ch = this.redoStack.pop()
    if (!ch) return null
    this.undoStack.push(ch)
    return this.step(ch, 'next')
  }

  // ---------------------------------------------------------------- collaboration

  /** takes the change recorded by the last operation out of the (unused) undo history */
  takeRecord(): Change | undefined {
    const ch = this.undoStack.pop()
    this.undoStack = []
    this.redoStack = []
    return ch
  }

  /** exactly undoes a recorded change (rolling back an unconfirmed operation) */
  revertRecord(ch: Change): XlsxEditResult {
    return this.step(ch, 'prev')
  }

  /** sets the cells / layout touched by `ch` to its `which` state, recorded as a new change */
  restoreRecord(ch: Change, which: 'prev' | 'next'): XlsxEditResult {
    const d = this.sheet(ch.sheet)
    const cells: CellChange[] = ch.cells.map((c) => ({
      key: c.key,
      prev: getEntry(d, Math.floor(c.key / MAX_COL), c.key % MAX_COL),
      prevF: d.formulas.get(c.key),
      next: which === 'prev' ? c.prev : c.next,
      nextF: which === 'prev' ? c.prevF : c.nextF,
    }))
    const counter: Change = {
      sheet: ch.sheet,
      cells,
      range: ch.range,
      lay: ch.lay ? { before: layoutState(d), after: which === 'prev' ? ch.lay.before : ch.lay.after } : undefined,
      full: ch.full ? { before: snapshot(d), after: which === 'prev' ? ch.full.before : ch.full.after } : undefined,
      cm: ch.cm
        ? { before: ch.cm.before.map(([id]) => [id, this.wb.comments.get(id) ?? null]), after: which === 'prev' ? [...ch.cm.before].reverse() : ch.cm.after }
        : undefined,
    }
    const res = this.step(counter, 'next')
    this.undoStack.push(counter)
    return res
  }

  /** style table sizes before an operation */
  async styleMark(): Promise<Record<string, number>> {
    return (await this.styles()).mark()
  }

  /** drops styles derived after `m` (the operations that used them were rolled back) */
  async styleTruncate(m: Record<string, number>) {
    const ed = await this.styles()
    const before = ed.xml()
    ed.truncate(m)
    const xml = ed.xml()
    if (xml === before) return
    this.wb.pkg.setPart(this.stylesPath!, xml)
    this.wb.styles = parseSheetStyles(parseXml(xml), this.wb.theme)
  }

  /** style index of a cell (0 when empty) */
  styleAt(sheet: number, r: number, c: number): number {
    return getEntry(this.sheet(sheet), r, c)?.[2] ?? 0
  }

  /** current layout of a sheet (after structural edits) */
  layout(sheet: number): SheetProgress {
    const d = this.sheet(sheet)
    const rowSizes = rowSizesOf(d)
    return {
      sheet, rows: d.maxRow + 1, cols: d.maxCol + 1, rowSizes, done: true, colInfo: d.colInfo, defaultColW: d.defaultColW,
      defaultRowH: d.defaultRowH, frozen: d.frozen, gridLines: d.gridLines, merges: d.merges,
    }
  }

  private async allSheets(): Promise<{ name: string; d: SheetData }[]> {
    await Promise.all(this.wb.sheets.map((s) => this.wb.loadSheet(s.index, () => {})))
    return this.wb.sheets.map((s) => ({ name: s.name, d: this.wb.sheetsData.get(s.index)! }))
  }

  async sheetOp(sheet: number, op: SheetOp): Promise<XlsxEditResult> {
    const d = this.sheet(sheet)
    const name = this.wb.sheets[sheet].name
    const all: Range = { r0: 0, c0: 0, r1: Math.max(0, d.maxRow + 50), c1: Math.max(0, d.maxCol + 10) }
    const lay = (fn: () => void, range = all): XlsxEditResult => {
      const before = layoutState(d)
      fn()
      const ch: Change = { sheet, cells: [], range, lay: { before, after: layoutState(d) } }
      this.undoStack.push(ch)
      this.redoStack = []
      return { sheet, range, layout: this.layout(sheet) }
    }
    switch (op.op) {
      case 'insertRows':
      case 'deleteRows':
      case 'insertCols':
      case 'deleteCols': {
        if (d.pages) throw new Error(`Rows and columns cannot be inserted or deleted in a sheet of ${(d.maxRow + 1).toLocaleString()} rows (its rows are read from the file as they are shown)`)
        const others = await this.allSheets()
        const before = snapshot(d)
        const axis = op.op.endsWith('Rows') ? 'row' : 'col'
        shiftSheet(d, name, axis, op.at, op.op.startsWith('insert') ? op.n : -op.n, others)
        const cm = this.shiftComments(this.wb.sheets[sheet].path, axis, op.at, op.op.startsWith('insert') ? op.n : -op.n)
        this.formulaTouched = true
        const recalc = this.recalcAll()
        const ch: Change = { sheet, cells: [], range: all, full: { before, after: snapshot(d) }, cm }
        this.undoStack.push(ch)
        this.redoStack = []
        return { sheet, range: all, layout: this.layout(sheet), recalc }
      }
      case 'merge':
      case 'unmerge': {
        const R = op.range
        const blocks: Range[] =
          op.op === 'unmerge' || !op.mode || op.mode === 'all'
            ? [R]
            : op.mode === 'across'
              ? Array.from({ length: R.r1 - R.r0 + 1 }, (_, i) => ({ r0: R.r0 + i, r1: R.r0 + i, c0: R.c0, c1: R.c1 }))
              : Array.from({ length: R.c1 - R.c0 + 1 }, (_, i) => ({ r0: R.r0, r1: R.r1, c0: R.c0 + i, c1: R.c0 + i }))
        const res = lay(() => {
          setMerge(d, R, false)
          if (op.op === 'merge') for (const b of blocks) setMerge(d, b, true)
        }, R)
        if (op.op === 'merge') {
          // like Excel: only the top-left value of each merged block survives
          const clear: CellInput[] = []
          for (const b of blocks)
            for (let r = b.r0; r <= b.r1; r++)
              for (let c = b.c0; c <= b.c1; c++) if ((r !== b.r0 || c !== b.c0) && getEntry(d, r, c) && getEntry(d, r, c)?.[1] !== '') clear.push({ r, c, input: '' })
          if (clear.length) {
            await this.setCells(sheet, clear)
            const last = this.undoStack.pop()!
            this.undoStack[this.undoStack.length - 1].cells = last.cells
          }
        }
        return res
      }
      case 'hideRows':
        return lay(() => {
          for (let r = op.a; r <= op.b; r++) {
            const cur = d.rowAttrs.get(r)
            if (op.hidden) d.rowAttrs.set(r, { ...(cur ?? {}), hidden: true })
            else if (cur) {
              const { hidden: _h, ...rest } = cur
              if (rest.ht !== undefined) d.rowAttrs.set(r, rest)
              else d.rowAttrs.delete(r)
            }
            d.dirtyRows.add(r)
          }
        })
      case 'hideCols':
        return lay(() => hideCols(d, op.a, op.b, op.hidden))
      case 'gridlines':
        return lay(() => {
          d.gridLines = op.show
          d.dirtyView = true
        })
      case 'colWidth':
        return lay(() => setColWidth(d, op.c0, op.c1, op.px))
      case 'rowHeight':
        return lay(() => setRowHeight(d, op.r0, op.r1, op.px))
      case 'freeze':
        return lay(() => {
          d.frozen = { rows: op.rows, cols: op.cols }
          d.dirtyView = true
        })
      case 'fill':
        await this.wb.ensure(d, op.range.r0, op.range.r1)
        return this.fill(sheet, d, op.range, op.dir)
      case 'sort':
        await this.wb.ensure(d, op.range.r0, op.range.r1)
        return this.sort(sheet, d, op.range, op.col, !!op.desc)
    }
  }

  /** replaces `query` in matching cells (one undo step); `at` limits it to one cell */
  async replace(sheet: number, query: string, by: string, opts: FindOptions = {}, at?: { r: number; c: number }): Promise<XlsxEditResult | null> {
    const d = this.sheet(sheet)
    const re = new RegExp(opts.whole ? `^${escRe(query)}$` : escRe(query), opts.matchCase ? 'g' : 'gi')
    const inputs: CellInput[] = []
    const cells = at ? [at] : await this.wb.find(sheet, query, opts, 1e9)
    for (const { r, c } of cells) {
      const f = d.formulas.get(r * MAX_COL + c)
      const raw = opts.formulas && f ? '=' + f : (await this.wb.cell(sheet, r, c))?.value
      if (raw === undefined || (f && !opts.formulas)) continue
      const next = raw.replace(re, by)
      if (next !== raw) inputs.push({ r, c, input: next })
    }
    return inputs.length ? this.setCells(sheet, inputs) : null
  }

  private fill(sheet: number, d: SheetData, range: Range, dir: 'down' | 'right'): XlsxEditResult {
    materializeShared(d)
    const changes: CellChange[] = []
    for (let r = range.r0; r <= range.r1; r++) {
      for (let c = range.c0; c <= range.c1; c++) {
        const sr = dir === 'down' ? range.r0 : r
        const sc = dir === 'right' ? range.c0 : c
        if (sr === r && sc === c) continue
        const key = r * MAX_COL + c
        const src = getEntry(d, sr, sc)
        const sf = d.formulas.get(sr * MAX_COL + sc)
        const nf = sf ? translateFormula(sf, r - sr, c - sc) : undefined
        changes.push({ key, prev: getEntry(d, r, c), prevF: d.formulas.get(key), next: src ? [src[0], nf ? '=' + nf : src[1], src[2]] : null, nextF: nf })
      }
    }
    this.apply(d, changes, 'next')
    return this.record(sheet, changes, range)
  }

  private sort(sheet: number, d: SheetData, range: Range, col: number, desc: boolean): XlsxEditResult {
    materializeShared(d)
    const rows: { r: number; cells: (Entry | null)[]; f: (string | undefined)[]; key: unknown }[] = []
    for (let r = range.r0; r <= range.r1; r++) {
      const cells: (Entry | null)[] = []
      const f: (string | undefined)[] = []
      for (let c = range.c0; c <= range.c1; c++) {
        cells.push(getEntry(d, r, c))
        f.push(d.formulas.get(r * MAX_COL + c))
      }
      const e = getEntry(d, r, col)
      const v = e ? (e[0] === T_SST ? this.wb.sst[e[1] as number] : e[1]) : null
      rows.push({ r, cells, f, key: v === '' ? null : v })
    }
    const cmp = (a: unknown, b: unknown) => {
      if (a === null || a === undefined) return 1
      if (b === null || b === undefined) return -1
      if (typeof a === 'number' && typeof b === 'number') return a - b
      if (typeof a === 'number') return -1
      if (typeof b === 'number') return 1
      return String(a).localeCompare(String(b), undefined, { sensitivity: 'base', numeric: true })
    }
    const sorted = [...rows].sort((a, b) => {
      const k = cmp(a.key, b.key)
      return desc && a.key !== null && b.key !== null ? -k : k
    })
    const changes: CellChange[] = []
    sorted.forEach((src, i) => {
      const r = range.r0 + i
      for (let c = range.c0; c <= range.c1; c++) {
        const k = c - range.c0
        const key = r * MAX_COL + c
        const nf = src.f[k] ? translateFormula(src.f[k]!, r - src.r, 0) : undefined
        const e = src.cells[k]
        changes.push({ key, prev: getEntry(d, r, c), prevF: d.formulas.get(key), next: e ? [e[0], nf ? '=' + nf : e[1], e[2]] : null, nextF: nf })
      }
    })
    this.apply(d, changes, 'next')
    return this.record(sheet, changes, range)
  }

  // ---------------------------------------------------------------- workbook (sheet list)

  async bookOp(op: BookOp): Promise<XlsxEditResult> {
    const pkg = this.wb.pkg
    const nodes = parseRaw(await pkg.text(this.wb.mainPath))
    const root = rootEl(nodes)!
    const P = prefixOf(root.q)
    const sheetsEl = kidOf(root, 'sheets')!
    const entries = kidsOf(sheetsEl, 'sheet')
    const sheets = this.wb.sheets
    const remap = (order: number[]) => {
      // order[newIndex] = oldIndex
      const data = new Map(this.wb.sheetsData)
      this.wb.sheetsData.clear()
      order.forEach((old, i) => {
        const d = data.get(old)
        if (d) this.wb.sheetsData.set(i, d)
      })
      const next = order.map((old, i) => ({ ...sheets[old], index: i }))
      sheets.splice(0, sheets.length, ...next)
    }
    switch (op.op) {
      case 'add': {
        const path = pkg.freeName('xl/worksheets', 'sheet', 'xml')
        this.newSheets.set(path, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/><sheetData/></worksheet>')
        pkg.setPart(path, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/><sheetData/></worksheet>')
        await pkg.setContentType(path, 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml')
        const rid = await pkg.addRel(this.wb.mainPath, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet', path)
        const ids = entries.map((e) => Number(getAttr(e, 'sheetId') ?? 0))
        let name = op.name ?? `Sheet${sheets.length + 1}`
        while (sheets.some((s) => s.name.toLowerCase() === name.toLowerCase())) name += '_'
        const node = el(P + 'sheet', { name, sheetId: String(Math.max(0, ...ids) + 1), ...relIdAttrs(root, 'id', rid) })
        const after = op.after ?? sheets.length - 1
        const at = after + 1 >= entries.length ? sheetsEl.c.length : sheetsEl.c.indexOf(entries[after + 1])
        sheetsEl.c.splice(at, 0, node)
        sheetsEl.self = false
        const info = { name, index: after + 1, path, hidden: false }
        const order = sheets.map((_, i) => i)
        sheets.push({ ...info, index: sheets.length })
        order.splice(after + 1, 0, sheets.length - 1)
        remap(order)
        break
      }
      case 'rename': {
        const old = sheets[op.index]
        const name = op.name.trim().slice(0, 31).replace(/[\\/?*[\]:]/g, '_')
        if (!old || !name || sheets.some((s, i) => i !== op.index && s.name.toLowerCase() === name.toLowerCase())) throw new Error('Invalid sheet name')
        setAttr(entries[op.index], 'name', name)
        for (const s of await this.allSheets()) {
          for (const [k, f] of s.d.formulas) {
            const g = renameSheetInFormula(f, old.name, name)
            if (g !== f) {
              s.d.formulas.set(k, g)
              s.d.edited.add(k)
            }
          }
        }
        old.name = name
        break
      }
      case 'delete': {
        if (sheets.length <= 1) throw new Error('A workbook needs at least one sheet')
        const s = sheets[op.index]
        // its comments go with it
        for (const t of this.wb.comments.threads()) {
          const a = t.root.at
          if (!a || !('sheet' in a) || a.sheet !== s.path) continue
          for (const r of t.replies) this.wb.comments.set(r.id, null)
          this.wb.comments.set(t.root.id, null)
        }
        sheetsEl.c.splice(sheetsEl.c.indexOf(entries[op.index]), 1)
        pkg.setPart(s.path, null)
        if (pkg.has(relsPathFor(s.path))) pkg.setPart(relsPathFor(s.path), null)
        await pkg.removeRelsTo(this.wb.mainPath, s.path)
        await pkg.setContentType(s.path, null)
        remap(sheets.map((_, i) => i).filter((i) => i !== op.index))
        // localSheetId indexes in definedNames would now be off; drop the ones of the deleted sheet
        const dn = kidOf(root, 'definedNames')
        if (dn) dn.c = dn.c.filter((c) => !(isNode(c) && getAttr(c, 'localSheetId') === String(op.index)))
        break
      }
      case 'move': {
        const node = entries[op.from]
        sheetsEl.c.splice(sheetsEl.c.indexOf(node), 1)
        const rest = kidsOf(sheetsEl, 'sheet')
        sheetsEl.c.splice(op.to >= rest.length ? sheetsEl.c.length : sheetsEl.c.indexOf(rest[op.to]), 0, node)
        const order = sheets.map((_, i) => i)
        order.splice(op.from, 1)
        order.splice(op.to, 0, op.from)
        remap(order)
        break
      }
    }
    pkg.setPart(this.wb.mainPath, serialize(nodes))
    this.bookChanged = true
    this.undoStack = []
    this.redoStack = []
    return { sheet: 0, range: { r0: 0, c0: 0, r1: 0, c1: 0 }, sheets: sheets.map((s) => ({ ...s })) }
  }

  private bookChanged = false
  /** XML of sheets created in this session (no original to stream from) */
  private newSheets = new Map<string, string>()

  /**
   * Adds a sheet after `after` holding the rows of a delimited text file (CSV / TSV): they stream into the sheet
   * when it is loaded, like a sheet's XML does. `name` defaults to the file's name.
   */
  async importText(after: number, imp: DelimitedImport, name?: string): Promise<XlsxEditResult> {
    const base = (name ?? imp.source.name ?? '').replace(/\.[^./\\]+$/, '').replace(/[\\/?*[\]:]/g, '_').slice(0, 31).trim() || 'Imported'
    const res = await this.bookOp({ op: 'add', name: base, after })
    this.wb.loaders.set(this.wb.sheets[after + 1].path, (d, report) => loadDelimited(d, imp, report))
    return res
  }

  // ---------------------------------------------------------------- save

  async save(): Promise<Blob> {
    // rows still to come from an import are in the model before the sheets are written
    await this.wb.loadPending()
    const pkg = this.wb.pkg
    let any = false
    for (const [index, d] of this.wb.sheetsData) {
      if (d.maxRow >= MAX_ROWS) {
        throw new Error(
          `Sheet “${this.wb.sheets[index].name}” has ${(d.maxRow + 1).toLocaleString()} rows: an Excel workbook holds at most ${MAX_ROWS.toLocaleString()}. Save the sheet as CSV instead.`,
        )
      }
      if (!needsWrite(d)) continue
      any = true
      const info = this.wb.sheets[index]
      pkg.setPart(info.path, () => this.rewrite(info.path, d))
    }
    if (any) {
      // stale cached values / calc chain: let Excel recalculate on open
      const wbNodes = parseRaw(await pkg.text(this.wb.mainPath))
      const root = rootEl(wbNodes)!
      const P = prefixOf(root.q)
      const calc = clone(kidOf(root, 'calcPr') ?? el(P + 'calcPr'))
      setAttr(calc, 'fullCalcOnLoad', '1')
      setOrdered(root, calc, 'calcPr', WB_ORDER)
      pkg.setPart(this.wb.mainPath, serialize(wbNodes))
      const chain = relOfType(this.wb.rels, 'calcChain')
      if (chain && this.formulaTouched && pkg.has(chain.target)) {
        pkg.setPart(chain.target, null)
        await pkg.removeRelsTo(this.wb.mainPath, chain.target)
        await pkg.setContentType(chain.target, null)
      }
    }
    // after the sheets: a sheet given comments here links its drawing of them
    if (this.wb.comments.dirty) await writeXlsxComments(pkg, this.wb.mainPath, this.wb.comments, this.wb.sheets)
    return pkg.save(XLSX_MIME)
  }

  /** Streams the original sheet XML, patching only rows that contain edits. */
  private rewrite(path: string, d: SheetData): ReadableStream<Uint8Array> {
    const pkg = this.wb.pkg
    const wb = this.wb
    // eslint-disable-next-line @typescript-eslint/no-this-alias
    const self = this
    // shared formula groups whose master was overwritten → dependents get explicit formulas
    const fix = new Map<number, string>()
    const broken = new Set<number>()
    for (const [si, g] of d.shared) if (d.edited.has(g.master)) broken.add(si)
    if (broken.size) {
      for (const [key, si] of d.depOf) {
        if (!broken.has(si) || d.edited.has(key)) continue
        const g = d.shared.get(si)!
        const dr = Math.floor(key / MAX_COL) - Math.floor(g.master / MAX_COL)
        const dc = (key % MAX_COL) - (g.master % MAX_COL)
        fix.set(key, translateFormula(g.text, dr, dc))
      }
    }
    const rowCols = new Map<number, Set<number>>()
    for (const r of d.dirtyRows) if (!rowCols.has(r)) rowCols.set(r, new Set())
    for (const key of [...d.edited, ...fix.keys()]) {
      const r = Math.floor(key / MAX_COL)
      let s = rowCols.get(r)
      if (!s) rowCols.set(r, (s = new Set()))
      s.add(key % MAX_COL)
    }
    const rows = [...rowCols.keys()].sort((a, b) => a - b)
    const enc = new TextEncoder()
    return new ReadableStream<Uint8Array>({
      async start(ctrl) {
        const rw = new RowPatcher(rows, rowCols, d, fix, wb)
        const base = self.newSheets.get(path)
        const reader = (base !== undefined ? new Blob([base]).stream() : await pkg.zip.stream(path)).getReader()
        const dec = new TextDecoder()
        for (;;) {
          const { done, value } = await reader.read()
          const text = done ? dec.decode() : dec.decode(value, { stream: true })
          for (const part of await rw.push(text, done)) if (part) ctrl.enqueue(enc.encode(part))
          if (done) break
        }
        ctrl.close()
      },
    })
  }
}

function cellXml(d: SheetData, r: number, c: number, P: string, wb: XlsxWorkbook): string {
  const key = r * MAX_COL + c
  const e = getEntry(d, r, c)
  const f = d.formulas.get(key)
  if (!e && !f) return ''
  const ref = colName(c) + (r + 1)
  const s = e?.[2] ? ` s="${e[2]}"` : ''
  if (f) {
    const [t, v] = e ?? [T_EMPTY, '']
    const fx = `<${P}f>${escText(f)}</${P}f>`
    if (t === T_NUM) return `<${P}c r="${ref}"${s}>${fx}<${P}v>${String(v)}</${P}v></${P}c>`
    if (t === T_BOOL) return `<${P}c r="${ref}"${s} t="b">${fx}<${P}v>${v ? 1 : 0}</${P}v></${P}c>`
    if (t === T_ERR) return `<${P}c r="${ref}"${s} t="e">${fx}<${P}v>${escText(String(v))}</${P}v></${P}c>`
    if (t === T_STR && !String(v).startsWith('=')) return `<${P}c r="${ref}"${s} t="str">${fx}<${P}v>${escText(String(v))}</${P}v></${P}c>`
    return `<${P}c r="${ref}"${s}>${fx}</${P}c>`
  }
  const [t, v] = e!
  switch (t) {
    case T_NUM:
      return `<${P}c r="${ref}"${s}><${P}v>${String(v)}</${P}v></${P}c>`
    case T_SST:
    case T_STR: {
      const text = t === T_SST ? wb.sst[v as number] ?? '' : String(v)
      return `<${P}c r="${ref}"${s} t="inlineStr"><${P}is><${P}t xml:space="preserve">${escText(text)}</${P}t></${P}is></${P}c>`
    }
    case T_BOOL:
      return `<${P}c r="${ref}"${s} t="b"><${P}v>${v ? 1 : 0}</${P}v></${P}c>`
    case T_ERR:
      return `<${P}c r="${ref}"${s} t="e"><${P}v>${escText(String(v))}</${P}v></${P}c>`
    case T_DATE:
      return `<${P}c r="${ref}"${s} t="d"><${P}v>${escText(String(v))}</${P}v></${P}c>`
    default:
      return `<${P}c r="${ref}"${s}/>`
  }
}

function needsWrite(d: SheetData): boolean {
  return d.edited.size > 0 || d.structural || d.dirtyCols || d.dirtyMerges || d.dirtyView || d.dirtyRows.size > 0 || d.formulaChanged
}

function layoutState(d: SheetData): LayoutState {
  return {
    merges: d.merges.map((m) => ({ ...m })),
    colInfo: d.colInfo.map((c) => ({ ...c })),
    rowAttrs: [...d.rowAttrs].map(([k, v]) => [k, { ...v }]),
    frozen: { ...d.frozen },
    gridLines: d.gridLines,
  }
}

function applyLayout(d: SheetData, l: LayoutState) {
  if (l.merges) d.merges = l.merges.map((m) => ({ ...m }))
  if (l.colInfo) d.colInfo = l.colInfo.map((c) => ({ ...c }))
  if (l.rowAttrs) {
    for (const r of d.rowAttrs.keys()) d.dirtyRows.add(r)
    d.rowAttrs = new Map(l.rowAttrs.map(([k, v]) => [k, { ...v }]))
    for (const r of d.rowAttrs.keys()) d.dirtyRows.add(r)
  }
  if (l.frozen) d.frozen = { ...l.frozen }
  if (l.gridLines !== undefined) d.gridLines = l.gridLines
  d.dirtyCols = d.dirtyMerges = d.dirtyView = true
}

/** hides / shows columns a..b (the width before hiding is kept for unhiding) */
function hideCols(d: SheetData, a: number, b: number, hidden: boolean) {
  const out: ColInfo[] = []
  const covered: [number, number][] = []
  for (const ci of d.colInfo) {
    if (ci.max < a || ci.min > b) {
      out.push(ci)
      continue
    }
    if (ci.min < a) out.push({ ...ci, max: a - 1 })
    if (ci.max > b) out.push({ ...ci, min: b + 1 })
    const lo = Math.max(ci.min, a)
    const hi = Math.min(ci.max, b)
    covered.push([lo, hi])
    if (hidden) out.push(ci.w ? { ...ci, min: lo, max: hi, w: 0, prevW: ci.w } : { ...ci, min: lo, max: hi })
    else out.push(ci.w ? { ...ci, min: lo, max: hi } : { min: lo, max: hi, w: ci.prevW || d.defaultColW })
  }
  if (hidden) {
    // columns without an explicit width
    covered.sort((x, y) => x[0] - y[0])
    let at = a
    for (const [lo, hi] of [...covered, [b + 1, b + 1] as [number, number]]) {
      if (lo > at) out.push({ min: at, max: Math.min(lo - 1, b), w: 0, prevW: d.defaultColW })
      at = Math.max(at, hi + 1)
    }
  }
  d.colInfo = out.sort((x, y) => x.min - y.min)
  d.dirtyCols = true
}

function borderSides(kind: BorderKind, R: Range, r: number, c: number): Record<string, boolean> {
  const out: Record<string, boolean> = {}
  const edge = { top: r === R.r0, bottom: r === R.r1, left: c === R.c0, right: c === R.c1 }
  for (const side of ['top', 'bottom', 'left', 'right'] as const) {
    if (kind === 'all') out[side] = true
    else if (kind === 'none') out[side] = false
    else if (kind === 'outer') {
      if (edge[side]) out[side] = true
    } else if (kind === 'inner') {
      if (!edge[side]) out[side] = true
    } else if (kind === side && edge[side]) out[side] = true
  }
  return out
}

/**
 * Streaming sheet rewriter: finds <row> elements with plain string search and copies
 * everything verbatim except rows that contain edits (those are re-serialized).
 * Structural edits regenerate the whole <sheetData>; <cols>, panes and <mergeCells> are patched.
 */
class RowPatcher {
  private buf = ''
  private P = ''
  private next = 0
  private lastRow = -1
  private phase: 'head' | 'data' | 'skip' | 'tail' = 'head'
  private tail = ''

  constructor(
    private rows: number[],
    private rowCols: Map<number, Set<number>>,
    private d: SheetData,
    private fix: Map<number, string>,
    private wb: XlsxWorkbook,
  ) {}

  private newRow(r: number): string {
    const cols = [...(this.rowCols.get(r) ?? [])].sort((a, b) => a - b)
    const cells = cols.map((c) => cellXml(this.d, r, c, this.P, this.wb)).join('')
    const attrs = this.rowAttrXml(r)
    return cells || attrs ? `<${this.P}row r="${r + 1}"${attrs}>${cells}</${this.P}row>` : ''
  }

  private rowAttrXml(r: number): string {
    const a = this.d.rowAttrs.get(r)
    if (!a) return ''
    return (a.ht !== undefined ? ` ht="${a.ht}" customHeight="1"` : '') + (a.hidden ? ' hidden="1"' : '')
  }

  private before(r: number): string {
    let out = ''
    while (this.next < this.rows.length && this.rows[this.next] < r) out += this.newRow(this.rows[this.next++])
    return out
  }

  /** every row of the sheet, regenerated from the model (structural edits); a paged sheet is read page by page */
  private async *allRows(): AsyncGenerator<string> {
    const d = this.d
    let last = Math.max(d.rows.length - 1, d.maxRow)
    for (const r of d.rowAttrs.keys()) if (r > last) last = r
    const size = d.pages?.size ?? Infinity
    let chunk = ''
    for (let r = 0; r <= last; r++) {
      if (d.pages && r % size === 0) await d.pages.ensure(r, r + size - 1)
      const row = d.rows[r]
      const attrs = this.rowAttrXml(r)
      if (!row?.length && !attrs) continue
      let cells = ''
      if (row) {
        const cols: number[] = []
        for (let k = 0; k < row.length; k += 4) cols.push(row[k] as number)
        for (const c of cols.sort((a, b) => a - b)) cells += cellXml(d, r, c, this.P, this.wb)
      }
      chunk += `<${this.P}row r="${r + 1}"${attrs}>${cells}</${this.P}row>`
      if (chunk.length > 1 << 16) {
        yield chunk
        chunk = ''
      }
    }
    if (chunk) yield chunk
  }

  private colsXml(): string {
    const P = this.P
    const cols = [...this.d.colInfo].sort((a, b) => a.min - b.min)
    if (!cols.length) return ''
    return `<${P}cols>${cols.map((c) => `<${P}col min="${c.min + 1}" max="${c.max + 1}" width="${+((c.w || c.prevW || 0) / 7).toFixed(4)}" customWidth="1"${c.w ? '' : ' hidden="1"'}/>`).join('')}</${P}cols>`
  }

  private headFix(head: string): string {
    const d = this.d
    const P = this.P
    let h = head
    if (d.dirtyCols || d.structural) {
      const re = new RegExp(`<${P}cols\\b[\\s\\S]*?<\\/${P}cols>|<${P}cols\\s*\\/>`)
      h = re.test(h) ? h.replace(re, this.colsXml()) : h + this.colsXml()
    }
    if (d.dirtyView || d.structural) {
      const { rows, cols } = d.frozen
      const pane = rows || cols
        ? `<${P}pane${cols ? ` xSplit="${cols}"` : ''}${rows ? ` ySplit="${rows}"` : ''} topLeftCell="${colName(cols)}${rows + 1}" activePane="${rows && cols ? 'bottomRight' : rows ? 'bottomLeft' : 'topRight'}" state="frozen"/>`
        : ''
      const re = new RegExp(`<${P}sheetViews>[\\s\\S]*?<\\/${P}sheetViews>`)
      const m = re.exec(h)
      if (m) {
        const nodes = parseRaw(m[0])
        const view = kidOf(rootEl(nodes), 'sheetView')
        if (view) {
          view.c = view.c.filter((c) => !(isNode(c) && (local(c.q) === 'pane' || local(c.q) === 'selection')))
          if (pane) view.c.unshift(pane)
          setAttr(view, 'showGridLines', d.gridLines ? undefined : '0')
          view.self = view.c.length === 0
          h = h.replace(m[0], serialize(nodes))
        }
      } else if (pane || !d.gridLines) {
        const block = `<${P}sheetViews><${P}sheetView${d.gridLines ? '' : ' showGridLines="0"'} workbookViewId="0">${pane}</${P}sheetView></${P}sheetViews>`
        const at = h.search(new RegExp(`<${P}(sheetFormatPr|cols)\\b`))
        h = at >= 0 ? h.slice(0, at) + block + h.slice(at) : h + block
      }
    }
    if (d.structural) {
      h = h.replace(new RegExp(`<${P}dimension\\b[^>]*\\/>`), `<${P}dimension ref="A1:${colName(Math.max(0, d.maxCol))}${Math.max(1, d.maxRow + 1)}"/>`)
    }
    return h
  }

  private tailFix(tail: string): string {
    const d = this.d
    if (!d.dirtyMerges && !d.structural) return tail
    const P = this.P
    const ms = d.merges
    const xml = ms.length ? `<${P}mergeCells count="${ms.length}">${ms.map((m) => `<${P}mergeCell ref="${colName(m.c0)}${m.r0 + 1}:${colName(m.c1)}${m.r1 + 1}"/>`).join('')}</${P}mergeCells>` : ''
    const re = new RegExp(`<${P}mergeCells\\b[\\s\\S]*?<\\/${P}mergeCells>|<${P}mergeCells\\b[^>]*\\/>`)
    if (re.test(tail)) return tail.replace(re, xml)
    // after the elements that precede mergeCells in the schema
    let at = 0
    const before = /<(\w+:)?(sheetCalcPr|sheetProtection|protectedRanges|scenarios|autoFilter|sortState|dataConsolidate|customSheetViews)\b[\s\S]*?(\/>|<\/(\w+:)?\2>)/g
    for (let m = before.exec(tail); m; m = before.exec(tail)) at = m.index + m[0].length
    return tail.slice(0, at) + xml + tail.slice(at)
  }

  /** feeds decoded text; returns what can be written so far */
  async push(text: string, final: boolean): Promise<string[]> {
    const s = this.buf + text
    this.buf = ''
    const out: string[] = []
    if (this.phase === 'tail') {
      this.tail += s
      if (final) out.push(this.tailFix(this.tail))
      return out
    }
    let i = 0
    if (this.phase === 'head') {
      const m = /<([\w.-]+:)?sheetData(\s[^>]*)?(\/?)>/.exec(s)
      if (!m) {
        if (final) out.push(this.headFix(s))
        else this.buf = s
        return out
      }
      this.P = m[1] ?? ''
      out.push(this.headFix(s.slice(0, m.index)))
      const end = m.index + m[0].length
      if (m[3] === '/') {
        out.push(`<${this.P}sheetData>`)
        if (this.d.structural) for await (const part of this.allRows()) out.push(part)
        else {
          let rows = ''
          while (this.next < this.rows.length) rows += this.newRow(this.rows[this.next++])
          out.push(rows)
        }
        out.push(`</${this.P}sheetData>`)
        this.phase = 'tail'
        return [...out, ...(await this.push(s.slice(end), final))]
      }
      out.push(m[0])
      i = end
      this.phase = this.d.structural ? 'skip' : 'data'
    }
    const dataClose = `</${this.P}sheetData>`
    if (this.phase === 'skip') {
      const dc = s.indexOf(dataClose, i)
      if (dc < 0) {
        this.buf = s.slice(Math.max(i, s.length - dataClose.length))
        return out
      }
      for await (const part of this.allRows()) out.push(part)
      out.push(dataClose)
      this.phase = 'tail'
      return [...out, ...(await this.push(s.slice(dc + dataClose.length), final))]
    }
    const rowOpen = `<${this.P}row`
    let chunk = ''
    let dc = s.indexOf(dataClose, i)
    while (i < s.length) {
      const ro = s.indexOf(rowOpen, i)
      if (dc >= 0 && dc < i) dc = s.indexOf(dataClose, i)
      if (dc >= 0 && (ro < 0 || dc < ro)) {
        chunk += s.slice(i, dc) + this.before(Infinity) + dataClose
        out.push(chunk)
        this.phase = 'tail'
        return [...out, ...(await this.push(s.slice(dc + dataClose.length), final))]
      }
      if (ro < 0) break
      const gt = s.indexOf('>', ro)
      if (gt < 0) {
        chunk += s.slice(i, ro)
        i = ro
        break
      }
      const tag = s.slice(ro, gt + 1)
      const after = tag.charCodeAt(rowOpen.length)
      if (after !== 32 && after !== 62 && after !== 47 && after !== 9 && after !== 10 && after !== 13) {
        chunk += s.slice(i, gt + 1)
        i = gt + 1
        continue
      }
      const m = /\sr="(\d+)"/.exec(tag)
      const r = m ? Number(m[1]) - 1 : this.lastRow + 1
      this.lastRow = r
      const self = tag.endsWith('/>')
      let end: number
      if (self) end = gt + 1
      else {
        const ce = s.indexOf(`</${this.P}row>`, gt)
        if (ce < 0) {
          chunk += s.slice(i, ro)
          i = ro
          this.lastRow = r - 1
          break
        }
        end = ce + this.P.length + 6
      }
      chunk += s.slice(i, ro) + this.before(r)
      if (this.next < this.rows.length && this.rows[this.next] === r) {
        this.next++
        chunk += this.patchRow(s.slice(ro, end), r)
      } else chunk += s.slice(ro, end)
      i = end
    }
    out.push(chunk)
    this.buf = s.slice(i)
    if (final) out.push(this.buf)
    return out
  }

  private patchRow(xml: string, r: number): string {
    const nodes = parseRaw(xml)
    const row = rootEl(nodes)
    if (!row) return xml
    const P = prefixOf(row.q)
    setAttr(row, 'spans', undefined)
    if (this.d.dirtyRows.has(r)) {
      const a = this.d.rowAttrs.get(r)
      setAttr(row, 'ht', a?.ht !== undefined ? String(a.ht) : undefined)
      setAttr(row, 'customHeight', a?.ht !== undefined ? '1' : undefined)
      setAttr(row, 'hidden', a?.hidden ? '1' : undefined)
    }
    const cols = [...(this.rowCols.get(r) ?? [])].sort((a, b) => a - b)
    // index existing cells by column
    const pos = new Map<number, number>()
    let prev = -1
    row.c.forEach((n, i) => {
      if (!isNode(n) || local(n.q) !== 'c') return
      const ref = getAttr(n, 'r')
      const c = ref ? parseRef(ref)[1] : prev + 1
      prev = c
      pos.set(c, i)
    })
    const replacements = new Map<number, RChild>()
    const inserts: { c: number; node: string }[] = []
    for (const c of cols) {
      const key = r * MAX_COL + c
      const at = pos.get(c)
      if (this.d.edited.has(key)) {
        const xml = cellXml(this.d, r, c, P, this.wb)
        if (at !== undefined) replacements.set(at, xml)
        else if (xml) inserts.push({ c, node: xml })
      } else if (at !== undefined && this.fix.has(key)) {
        const cell = row.c[at] as RNode
        const f = kidOf(cell, 'f')
        if (f) {
          f.a = ''
          f.c = [escText(this.fix.get(key)!)]
          f.self = false
        }
      }
    }
    for (const [i, x] of replacements) row.c[i] = x
    if (inserts.length) {
      // place new cells before the first existing cell with a greater column
      const ordered = [...pos.entries()].sort((a, b) => a[0] - b[0])
      for (const ins of inserts.reverse()) {
        const after = ordered.find(([c]) => c > ins.c)
        const idx = after ? after[1] : row.c.length
        row.c.splice(idx, 0, ins.node)
        for (const o of ordered) if (o[1] >= idx) o[1]++
      }
      row.self = false
    }
    return serialize(nodes)
  }
}

const escRe = (s: string) => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
