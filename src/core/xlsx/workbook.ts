import { OpcPackage, relOfType, type Rels } from '../opc'
import { readXlsxComments, XlsxComments } from './comments'
import type { LoadReport, RowPages, SheetLoader } from './import'
import { toCsv } from '../convert'
import { parseTheme, type Theme } from '../theme'
import { Calc } from './calc'
import { desc, kid, kids, parseXml, parseXmlStream, type Attrs, type SaxHandler, type XEl } from '../xml'
import { sheetColor } from '../color'
import { formatValue, formatText } from './numfmt'
import { parseSheetStyles, type SheetStyles } from './styles'
import { expandShared, formulaAt } from './structure'
import { stripFunctionPrefixes, type PivotInfo, type TableInfo } from './formula'
import { readPivots } from './pivot'
import { isoToSerial } from './formula/dates'
import type { Shift } from './follow'
import { MAX_COL, MAX_ROWS, MAX_COL_WIDTH, MAX_ROW_HEIGHT_PT, ooxmlCapped, ooxmlInt, ooxmlNumber } from './limits'
import type { BlockOut, CellDetail, CellOut, ColInfo, FindMatch, FindOptions, Merge, RowAttrs, RowOut, SheetInfo, SheetProgress, TextRun, WorkbookMeta } from './types'

export const T_NUM = 0
export const T_SST = 1
export const T_STR = 2
export const T_BOOL = 3
export const T_ERR = 4
export const T_DATE = 5
export const T_EMPTY = 6

export { MAX_COL, MAX_ROWS } from './limits'

const decodeX = (s: string) => (s.indexOf('_x') >= 0 ? s.replace(/_x([0-9A-Fa-f]{4})_/g, (_, h: string) => String.fromCharCode(parseInt(h, 16))) : s)

export function colIndex(letters: string): number {
  let c = 0
  for (let i = 0; i < letters.length; i++) {
    const code = letters.charCodeAt(i)
    if (code < 65) break
    c = c * 26 + ((code & ~32) - 64)
  }
  return c - 1
}

export function colName(c: number): string {
  let s = ''
  c++
  while (c > 0) {
    const m = (c - 1) % 26
    s = String.fromCharCode(65 + m) + s
    c = Math.floor((c - 1) / 26)
  }
  return s
}

export function parseRef(ref: string): [number, number] {
  let i = 0
  while (i < ref.length && ref.charCodeAt(i) >= 65) i++
  return [Number(ref.slice(i)) - 1, colIndex(ref.slice(0, i))]
}

export function parseRange(ref: string): Merge | undefined {
  const [a, b] = ref.replace(/\$/g, '').split(':')
  if (!a) return undefined
  const [r0, c0] = parseRef(a)
  const [r1, c1] = b ? parseRef(b) : [r0, c0]
  return { r0, c0, r1, c1 }
}

/** Validate file coordinates here, without restricting parseRef/parseRange used by CSV and editing. */
function worksheetRef(ref: string): [number, number] | undefined {
  // (a row number may come with leading zeros: A02)
  if (!/^\$?[A-Z]{1,3}\$?0*[1-9][0-9]{0,6}$/.test(ref)) return undefined
  const [r, c] = parseRef(ref.replace(/\$/g, ''))
  return r < MAX_ROWS && c < MAX_COL ? [r, c] : undefined
}

/** index in a row's entries (column, type, value, style: four slots each, by column) of the first at column `c` or after */
export function colAt(row: unknown[], c: number): number {
  // (the first cells of a row are the ones read most)
  if (!row.length || (row[0] as number) >= c) return 0
  let lo = 0
  let hi = row.length >> 2
  while (lo < hi) {
    const mid = (lo + hi) >> 1
    if ((row[mid << 2] as number) < c) lo = mid + 1
    else hi = mid
  }
  return lo << 2
}

/**
 * Row `r` of a sheet, to change in place: a row a snapshot holds (`SheetData.held`) is copied first, so that what
 * the snapshot was taken of stays as it was.
 */
export function ownRow(d: SheetData, r: number): unknown[] | undefined {
  const row = d.rows[r]
  if (!row || !d.held?.has(row)) return row
  const own = row.slice()
  d.rows[r] = own
  return own
}

export class SheetData {
  /** shared formula groups: si → master cell key + formula text */
  shared = new Map<number, { master: number; text: string }>()
  /** dependent cell key → shared formula si */
  depOf = new Map<number, number>()
  /** array formulas: first cell key → the cells the formula fills. One only while that cell has a formula. */
  arrays = new Map<number, Merge>()
  /**
   * The formulas that are dynamic arrays (by cell key): computed without implicit intersection, their values spill
   * over the cells next to them, as many as they give (`arrays` has the cells of each as it is now).
   */
  dynamic = new Set<number>()
  /** the cells of the sheet's filter (header row included): the rows hidden in it are filtered out */
  filter?: Merge
  /** edited cell keys (r * MAX_COL + c) */
  edited = new Set<number>()
  /** row heights in pt / hidden flags (from the file or the editor) */
  rowAttrs = new Map<number, RowAttrs>()
  /** rows grown to fit a large font (no explicit height in the file): px */
  autoRowH = new Map<number, number>()
  /** attributes of <sheetProtection> when the sheet is protected (`sheet="1"`) */
  protection?: Record<string, string>
  /** whole-sheet rewrite needed (rows/cols inserted or deleted) */
  structural = false
  /** sections to regenerate on save */
  dirtyCols = false
  dirtyMerges = false
  dirtyView = false
  /** rows whose height/visibility changed */
  dirtyRows = new Set<number>()
  /** rows whose cells moved as wholes (a sort): written again from the model, every cell of them */
  rewrite = new Set<number>()
  /** formulas rewritten because another sheet changed */
  formulaChanged = false
  /** formula cells computed anew since the sheet was read: the value the file holds for them is stale */
  revalued = new Set<number>()
  /** the cells of the shared formulas have their own text in `formulas` (`expandShared`) */
  sharedExpanded = false
  /** rows / columns inserted and deleted since the sheet was read, in order: what points at its cells in its XML follows when it is written */
  shifts: Shift[] = []
  /** sheets renamed since ([old, new]): its validations and conditional formats may name them */
  renames: [string, string][] = []
  /** formula cells stored without a cached value (to compute after load) */
  uncached: number[] = []
  /** its formulas were computed after it was read (a workbook that says to compute on load) */
  calculated = false
  /** it has formulas whose value is of the moment they are computed (TODAY, NOW, RAND): computed when the file is read, as Excel does */
  clock = false
  rows: (unknown[] | undefined)[] = []
  /** rows that snapshots (Undo of inserted / deleted rows and columns) hold as they are: see `ownRow` */
  held?: WeakSet<unknown[]>
  /** a sheet too large for memory: its rows are read from the file by pages, those in memory being in `rows` */
  pages?: RowPages
  maxRow = -1
  maxCol = -1
  formulas = new Map<number, string>()
  merges: Merge[] = []
  colInfo: ColInfo[] = []
  defaultColW = 64
  defaultRowH = 20
  frozen = { rows: 0, cols: 0 }
  gridLines = true
  done = false
  error?: string
  loading?: Promise<void>
}

/** attributes of <row> / <col> the model does not edit but a rewrite of the element must keep */
const ROW_KEEP = ['s', 'customFormat', 'outlineLevel', 'collapsed', 'thickTop', 'thickBot', 'ph']
const COL_KEEP = ['style', 'outlineLevel', 'collapsed', 'bestFit', 'phonetic']
const keep = (a: Attrs, names: string[]): Record<string, string> | undefined => {
  let out: Record<string, string> | undefined
  for (const n of names) if (a[n] !== undefined) (out ??= {})[n] = a[n]
  return out
}

class SheetHandler implements SaxHandler {
  private row = -1
  private rowCells: unknown[] | null = null
  private col = -1
  private cType = ''
  private cStyle = 0
  private val = ''
  private formula = ''
  private hasVal = false
  private capture = 0 // 1 v, 2 f, 3 inline t
  private fShared = -1
  private fArray = ''
  private cm = 0
  private inIs = false
  newSizes: [number, number][] = []
  headerReady = false
  private rowHasHt = false
  private rowMaxPt = 0
  /** last column of the row so far, and whether a cell came before it (a row is kept in column order, one cell per column) */
  private rowLast = -1
  private rowUnordered = false

  /** fontPt: font size per cell format, used to auto-grow rows without explicit height */
  constructor(
    private d: SheetData,
    private fontPt: number[] = [],
    private basePt = 11,
    /** whether a cell's metadata (`cm`) says its array formula is a dynamic one */
    private dynamic: (cm: number) => boolean = () => false,
  ) {}

  open(name: string, a: Attrs): boolean | void {
    switch (name) {
      case 'c': {
        if (!this.rowCells) return true
        const r = a.r
        const ref = r === undefined ? [this.row, this.col + 1] : worksheetRef(r)
        // A reference naming another row than the one the cell is in (a file written wrong): the cell is kept, in the
        // row it is in, at the column named. Dropping it would lose its value.
        if (!ref || ref[1] >= MAX_COL) return true
        this.col = ref[1]
        this.cType = a.t ?? ''
        this.cStyle = a.s ? Number(a.s) : 0
        this.cm = a.cm ? Number(a.cm) : 0
        this.val = ''
        this.formula = ''
        this.hasVal = false
        return
      }
      case 'autoFilter': {
        const rg = a.ref ? parseRange(a.ref.replace(/\$/g, '')) : undefined
        if (rg && rg.r0 >= 0 && rg.c0 >= 0 && rg.r1 >= rg.r0 && rg.c1 >= rg.c0) this.d.filter = rg
        return true
      }
      case 'v':
        this.capture = 1
        this.hasVal = true
        return
      case 'f':
        this.capture = 2
        this.fShared = a.t === 'shared' && a.si !== undefined ? Number(a.si) : -1
        this.fArray = a.t === 'array' && a.ref ? a.ref : ''
        return
      case 'is':
        this.inIs = true
        this.hasVal = true
        return
      case 't':
        if (this.inIs) this.capture = 3
        return
      case 'rPh':
        return true
      case 'row': {
        this.headerReady = true
        const row = a.r === undefined ? this.row + 1 : (ooxmlInt(a.r, 1, MAX_ROWS) ?? 0) - 1
        // Discard the whole record before sparse array lengths or progress counts can grow.
        if (row < 0 || row >= MAX_ROWS) return true
        this.row = row
        this.col = -1
        this.rowLast = -1
        this.rowUnordered = false
        this.rowCells = []
        const hidden = a.hidden === '1' || a.hidden === 'true'
        // (a height past Excel's largest: the largest, not the default)
        const ht = ooxmlCapped(a.ht, 0, MAX_ROW_HEIGHT_PT)
        const x = keep(a, ROW_KEEP)
        if (hidden || ht !== undefined || x) this.d.rowAttrs.set(this.row, { ht, hidden: hidden || undefined, x })
        this.rowHasHt = hidden || ht !== undefined
        this.rowMaxPt = 0
        if (hidden) this.newSizes.push([this.row, 0])
        else if (ht !== undefined) {
          const px = Math.round((ht * 4) / 3)
          if (px !== this.d.defaultRowH) this.newSizes.push([this.row, px])
        }
        return
      }
      case 'sheetFormatPr': {
        const h = ooxmlNumber(a.defaultRowHeight, 0, MAX_ROW_HEIGHT_PT)
        if (h) this.d.defaultRowH = Math.max(1, Math.round((h * 4) / 3))
        const w = ooxmlNumber(a.defaultColWidth, 0, MAX_COL_WIDTH)
        const base = ooxmlInt(a.baseColWidth, 0, MAX_COL_WIDTH)
        if (w) this.d.defaultColW = Math.max(1, Math.round(w * 7))
        else if (base !== undefined) this.d.defaultColW = Math.round(base * 7 + 8)
        return
      }
      case 'col': {
        const first = ooxmlInt(a.min, 1, MAX_COL)
        const last = ooxmlInt(a.max, 1, Number.MAX_SAFE_INTEGER)
        if (first === undefined || last === undefined || first > last) return
        const min = first - 1
        // columns past the last one, a width past the largest: kept as far as the sheet goes (the record may also say hidden)
        const max = Math.min(last, MAX_COL) - 1
        const hidden = a.hidden === '1' || a.hidden === 'true'
        const width = ooxmlCapped(a.width, 0, MAX_COL_WIDTH)
        if (a.width !== undefined && width === undefined) return
        const w = hidden ? 0 : width !== undefined ? Math.round(width * 7) : this.d.defaultColW
        const ci: ColInfo = { min, max, w }
        const x = keep(a, COL_KEEP)
        if (x) ci.x = x
        // the width a hidden column comes back with when shown
        if (hidden && width !== undefined) ci.prevW = Math.round(width * 7)
        this.d.colInfo.push(ci)
        return
      }
      case 'sheetProtection':
        if (a.sheet === '1' || a.sheet === 'true') this.d.protection = { ...a }
        return
      case 'pane':
        if (a.state === 'frozen' || a.state === 'frozenSplit') {
          this.d.frozen = { rows: ooxmlInt(a.ySplit, 0, MAX_ROWS - 1) ?? 0, cols: ooxmlInt(a.xSplit, 0, MAX_COL - 1) ?? 0 }
        }
        return
      case 'sheetView':
        if (a.showGridLines === '0' || a.showGridLines === 'false') this.d.gridLines = false
        return
      case 'sheetData':
        this.headerReady = true
        return
      case 'mergeCell': {
        const parts = a.ref?.split(':') ?? []
        const start = worksheetRef(parts[0] ?? '')
        const end = parts.length === 1 ? start : worksheetRef(parts[1] ?? '')
        if (parts.length <= 2 && start && end && start[0] <= end[0] && start[1] <= end[1]) {
          this.d.merges.push({ r0: start[0], c0: start[1], r1: end[0], c1: end[1] })
        }
        return
      }
      case 'extLst':
      case 'conditionalFormatting':
      case 'dataValidations':
      case 'headerFooter':
        return true
    }
  }

  close(name: string): void {
    switch (name) {
      case 'v':
      case 'f':
      case 't':
        this.capture = 0
        return
      case 'is':
        this.inIs = false
        return
      case 'c': {
        const cells = this.rowCells
        if (!cells) return
        const c = this.col
        if (c > this.d.maxCol) this.d.maxCol = c
        const again = c <= this.rowLast
        if (again) this.rowUnordered = true
        else this.rowLast = c
        const key = this.row * MAX_COL + c
        if (this.formula) {
          this.d.formulas.set(key, this.formula)
          if (!this.d.clock && CLOCK.test(this.formula)) this.d.clock = true
        }
        // (a column read twice: the cell read last is the cell, with its formula or without)
        else if (again) this.d.formulas.delete(key)
        if (this.fArray) {
          const rg = this.formula ? parseRange(this.fArray) : undefined
          if (rg && rg.r0 >= 0 && rg.c0 >= 0 && rg.r1 >= rg.r0 && rg.c1 >= rg.c0 && rg.r1 < MAX_ROWS && rg.c1 < MAX_COL) {
            const dynamic = this.cm > 0 && this.dynamic(this.cm)
            if (dynamic) this.d.dynamic.add(key)
            // (a dynamic array formula that fills its own cell only fills no cells: it gives one value, or cannot spill)
            if (!dynamic || rg.r1 > rg.r0 || rg.c1 > rg.c0) this.d.arrays.set(key, rg)
          }
          this.fArray = ''
        }
        if (this.fShared >= 0) {
          if (this.formula) this.d.shared.set(this.fShared, { master: key, text: this.formula })
          else this.d.depOf.set(key, this.fShared)
          this.fShared = -1
        }
        if (!this.hasVal || (this.formula && this.val === '' && this.cType !== 'str')) {
          if (this.formula) {
            this.d.uncached.push(this.row * MAX_COL + c)
            cells.push(c, T_EMPTY, '', this.cStyle)
          } else if (this.cStyle) cells.push(c, T_EMPTY, '', this.cStyle)
          return
        }
        let type = T_NUM
        let v: unknown
        switch (this.cType) {
          case 's':
            type = T_SST
            v = Number(this.val)
            break
          case 'str':
          case 'inlineStr':
            type = T_STR
            v = decodeX(this.val)
            break
          case 'b':
            type = T_BOOL
            v = this.val === '1' || this.val === 'true' ? 1 : 0
            break
          case 'e':
            type = T_ERR
            v = this.val
            break
          case 'd':
            type = T_DATE
            v = this.val
            break
          default: {
            const n = Number(this.val)
            if (this.val === '' || Number.isNaN(n)) {
              type = this.val === '' ? T_EMPTY : T_STR
              v = this.val
            } else v = n
          }
        }
        cells.push(c, type, v, this.cStyle)
        if (!this.rowHasHt) {
          const pt = this.fontPt[this.cStyle] ?? this.basePt
          if (pt > this.rowMaxPt) this.rowMaxPt = pt
        }
        return
      }
      case 'row':
        if (!this.rowHasHt && this.rowMaxPt > this.basePt * 1.15) {
          const px = Math.min(MAX_ROW_HEIGHT_PT * (4 / 3), Math.round(this.rowMaxPt * (4 / 3) * 1.3 + 3))
          if (px > this.d.defaultRowH) {
            this.newSizes.push([this.row, px])
            this.d.autoRowH.set(this.row, px)
          }
        }
        if (this.rowCells && this.rowUnordered) this.rowCells = ordered(this.rowCells)
        if (this.rowCells && this.rowCells.length) this.d.rows[this.row] = this.rowCells
        if (this.row > this.d.maxRow) this.d.maxRow = this.row
        this.rowCells = null
        return
    }
  }

  text(t: string): void {
    if (this.capture === 1) this.val += t
    else if (this.capture === 2) this.formula += t
    else if (this.capture === 3) this.val += t
  }
}

/**
 * The cells of a row in column order, the last one read of a column winning. Files hold them out of order (this
 * engine wrote such rows before 0.15.0), and what edits a row relies on the order.
 */
function ordered(cells: unknown[]): unknown[] {
  const at = new Map<number, number>()
  for (let k = 0; k < cells.length; k += 4) at.set(cells[k] as number, k)
  const out: unknown[] = []
  for (const k of [...at.entries()].sort((a, b) => a[0] - b[0]).map((e) => e[1])) out.push(cells[k], cells[k + 1], cells[k + 2], cells[k + 3])
  return out
}

/** a formula calling a function whose value is of the moment (what the file holds for it is of the day it was saved) */
const CLOCK = /(?:^|[^A-Za-z0-9_.])(?:TODAY|NOW|RAND|RANDBETWEEN|RANDARRAY)\s*\(/i

class SstHandler implements SaxHandler {
  strings: string[] = []
  /** the strings that are rich text (parts with fonts of their own), by index */
  rich = new Map<number, TextRun[]>()
  private cur = ''
  private capture = false
  private runs: TextRun[] | null = null
  private run: TextRun | null = null
  private inPr = false
  constructor(private color: (a: Attrs) => string | undefined = () => undefined) {}
  open(name: string, a: Attrs): boolean | void {
    if (name === 'si') {
      this.cur = ''
      this.runs = null
    } else if (name === 't') {
      this.capture = true
      // (text outside a run, in a string that has runs: a part in the cell's font)
      if (!this.run && this.runs) this.runs.push({ t: '' })
    } else if (name === 'rPh' || name === 'phoneticPr') return true
    else if (name === 'r') {
      this.run = { t: '' }
      ;(this.runs ??= this.cur ? [{ t: this.cur }] : []).push(this.run)
    } else if (name === 'rPr') {
      this.inPr = true
      if (this.run) this.run.own = true
    } else if (this.inPr && this.run) {
      const r = this.run
      const on = a.val !== '0' && a.val !== 'false'
      if (name === 'b') r.b = on
      else if (name === 'i') r.i = on
      else if (name === 'strike') r.strike = on
      else if (name === 'u') r.u = a.val !== 'none'
      else if (name === 'color') r.color = this.color(a)
      else if (name === 'sz') r.sz = Number(a.val) > 0 && Number(a.val) <= 409 ? Number(a.val) : undefined
      else if (name === 'rFont') r.font = a.val
      else if (name === 'vertAlign') r.va = a.val === 'superscript' || a.val === 'subscript' ? a.val : undefined
    }
  }
  close(name: string): void {
    if (name === 't') this.capture = false
    else if (name === 'rPr') this.inPr = false
    else if (name === 'r') this.run = null
    else if (name === 'si') {
      const runs = this.runs?.filter((r) => r.t)
      // (a string of one font is a plain one)
      if (runs?.some((r) => r.own)) this.rich.set(this.strings.length, runs.map((r) => ({ ...r, t: decodeX(r.t) })))
      this.strings.push(decodeX(this.cur))
      this.runs = null
    }
  }
  text(t: string): void {
    if (!this.capture) return
    this.cur += t
    if (this.runs?.length) this.runs[this.runs.length - 1].t += t
  }
}

/** where formatting a workbook that has no styles part puts one */
export const DEFAULT_STYLES_PART = 'xl/styles.xml'

/**
 * The workbook's styles part: the one its relationship names. A workbook that came without styles and was formatted
 * by a version before 0.14.2 holds them in a part no relationship names: it is read as well, or the file opens
 * without any of its formatting.
 */
export function stylesPartOf(pkg: OpcPackage, rels: Rels): string | undefined {
  return relOfType(rels, 'styles')?.target ?? (pkg.has(DEFAULT_STYLES_PART) ? DEFAULT_STYLES_PART : undefined)
}

/**
 * What a cell is exported as (a CSV / TSV file holds text): the value as shown, except a number with no format
 * (General), which keeps its digits (what is shown is rounded to the column's width), and a text another program
 * would take for a formula (starting with = + - @, a tab or a return), which gets an apostrophe in front (CSV
 * injection: OWASP). Numbers, dates and booleans are not texts: `-5` stays one.
 */
function csvField(c: CellOut, raw: number | undefined, numFmt: string | undefined): string {
  const [, text, , kind] = c
  if (kind === 0) return /^[=+\-@\t\r]/.test(text) ? `'${text}` : text
  if (raw !== undefined && Number.isFinite(raw) && (!numFmt || numFmt.toLowerCase() === 'general')) {
    // 15 significant digits are all a double keeps; the plain form, while it is short
    const full = String(Number(raw.toPrecision(15)))
    if (full.length <= 21 && !full.includes('e')) return full
  }
  return text
}

/** a defined name: the formula it stands for (without `=`), and the sheet (by index) it is local to */
export interface DefinedName {
  name: string
  text: string
  local?: number
}

/** a table column's name as written in its part: `_x000a_` is a line break */
const decodeColumn = (s: string) => s.replace(/_x([0-9A-Fa-f]{4})_/g, (_m, h: string) => String.fromCharCode(parseInt(h, 16)))

/**
 * The cell metadata records that mark dynamic array formulas: those of type XLDAPR whose properties say `fDynamic`.
 * A cell names its record by `cm` (counted from 1).
 */
async function readDynamicCm(pkg: OpcPackage, rels: Rels): Promise<Set<number>> {
  const out = new Set<number>()
  const rel = relOfType(rels, 'sheetMetadata')
  if (!rel || !pkg.has(rel.target)) return out
  // (the properties are in an extension list, which the cached reading of a part leaves out)
  const root = parseXml(await pkg.text(rel.target))
  const types = kids(kid(root, 'metadataTypes'), 'metadataType').map((t) => t.a.name)
  const future = kids(root, 'futureMetadata').find((f) => f.a.name === 'XLDAPR')
  const dynamic = kids(future, 'bk').map((bk) => {
    const p = desc(bk, 'dynamicArrayProperties')
    return !!p && (p.a.fDynamic === '1' || p.a.fDynamic === 'true')
  })
  kids(kid(root, 'cellMetadata'), 'bk').forEach((bk, i) => {
    for (const rc of kids(bk, 'rc')) if (types[Number(rc.a.t) - 1] === 'XLDAPR' && dynamic[Number(rc.a.v)]) out.add(i + 1)
  })
  return out
}

export class XlsxWorkbook {
  sheetsData = new Map<number, SheetData>()
  /** the comments (read in `open`) */
  comments = new XlsxComments()
  readonly calc = new Calc(this)
  sst: string[] = []
  /** the shared strings that are rich text, by index: the parts with a font of their own */
  rich = new Map<number, TextRun[]>()
  /** the sheet the file was saved on, when it is not the first */
  active?: number
  /** `<workbookProtection lockStructure>`: sheets cannot be added, renamed, deleted or moved */
  lockStructure = false
  /** the defined names, by name in lower case: what each stands for, and the sheet it is local to (see `loadNames`) */
  names = new Map<string, DefinedName[]>()
  /** the tables (lists with a header row) of all the sheets, for structured references */
  tables: TableInfo[] = []
  /** the pivot tables, as far as GETPIVOTDATA reads them */
  pivots: PivotInfo[] = []
  /** the cell metadata records (`cm`, counted from 1) that mark a dynamic array formula */
  dynamicCm = new Set<number>()
  /** the record a dynamic array formula written by this engine names (set when the file is saved) */
  dynamicRecord = 0
  private sstLoading?: Promise<void>
  /** sheets filled from something else than their XML (a text file being imported), by sheet path: run when the sheet loads */
  readonly loaders = new Map<string, SheetLoader>()

  private constructor(
    readonly pkg: OpcPackage,
    readonly mainPath: string,
    readonly rels: Rels,
    readonly sheets: SheetInfo[],
    public styles: SheetStyles,
    readonly date1904: boolean,
    readonly theme: Theme,
    /** the file says its formula values are not to be trusted (`fullCalcOnLoad`): they are computed when it is read */
    readonly fullCalc = false,
  ) {}

  static async open(pkg: OpcPackage, mainPath: string): Promise<XlsxWorkbook> {
    const root = await pkg.xml(mainPath)
    const rels = await pkg.rels(mainPath)
    const sheets: SheetInfo[] = []
    // (activeTab counts every <sheet>, chart sheets too, which are not read)
    const tab = Number(kid(kid(root, 'bookViews'), 'workbookView')?.a.activeTab ?? 0)
    let active: number | undefined
    kids(kid(root, 'sheets'), 'sheet').forEach((s, at) => {
      const rel = rels.get(s.a['r:id'])
      if (!rel || !rel.type.endsWith('/worksheet')) return
      const hidden = !!s.a.state && s.a.state !== 'visible'
      if (at === tab && !hidden) active = sheets.length
      sheets.push({ name: s.a.name ?? `Sheet${sheets.length + 1}`, index: sheets.length, path: rel.target, hidden })
    })
    const themeRel = relOfType(rels, 'theme')
    const theme = parseTheme(themeRel ? await pkg.xml(themeRel.target) : undefined)
    const stylesPath = stylesPartOf(pkg, rels)
    const styles = parseSheetStyles(stylesPath ? await pkg.xml(stylesPath) : undefined, theme)
    const pr = kid(root, 'workbookPr')
    const date1904 = pr?.a.date1904 === '1' || pr?.a.date1904 === 'true'
    const full = kid(root, 'calcPr')?.a.fullCalcOnLoad
    const wb = new XlsxWorkbook(pkg, mainPath, rels, sheets, styles, date1904, theme, full === '1' || full === 'true')
    const lock = kid(root, 'workbookProtection')?.a.lockStructure
    wb.lockStructure = lock === '1' || lock === 'true'
    wb.active = active
    await wb.loadNames().catch((e: Error) => console.warn('[office-engine] names and tables not read:', e.message))
    wb.dynamicCm = await readDynamicCm(pkg, rels).catch(() => new Set<number>())
    wb.comments = await readXlsxComments(pkg, mainPath, rels, sheets).catch((e: Error) => {
      console.warn('[office-engine] comments not read:', e.message)
      return new XlsxComments()
    })
    return wb
  }

  /**
   * Reads the defined names and the tables from the package as it is now: when the file is opened, and after an
   * edit rewrote them (rows / columns moved, a sheet renamed or deleted).
   */
  async loadNames(): Promise<void> {
    const names = new Map<string, DefinedName[]>()
    const root = await this.pkg.xml(this.mainPath, false)
    // `localSheetId` counts the sheets the workbook lists, chart sheets too: by name, the worksheet it is here (-1: none)
    const listed = kids(kid(root, 'sheets'), 'sheet').map((s) => this.sheets.findIndex((x) => x.name.toLowerCase() === (s.a.name ?? '').toLowerCase()))
    for (const n of kids(kid(root, 'definedNames'), 'definedName')) {
      const name = n.a.name
      // (macros and the names Excel keeps for itself — print areas, filters — are not formulas to evaluate)
      if (!name || n.a.function === '1' || n.a.vbProcedure === '1' || n.a.xlm === '1' || name.startsWith('_xlnm.')) continue
      const local = n.a.localSheetId !== undefined ? (listed[Number(n.a.localSheetId)] ?? -1) : undefined
      // (local to a sheet that is not a worksheet: no formula of a cell can use it)
      if (local !== undefined && local < 0) continue
      const list = names.get(name.toLowerCase())
      const def: DefinedName = { name, text: n.t.trim(), local }
      if (list) list.push(def)
      else names.set(name.toLowerCase(), [def])
    }
    // (the sheets side by side: each has its relationships to read, which a workbook of many sheets would wait for one by one)
    const perSheet = await Promise.all(
      this.sheets.map(async (sheet) => {
        const out: TableInfo[] = []
        const rels = await this.pkg.rels(sheet.path).catch(() => undefined)
        for (const r of rels?.values() ?? []) {
          if (!r.type.endsWith('/table') || r.external || !this.pkg.has(r.target)) continue
          const t = await this.pkg.xml(r.target, false).catch(() => undefined)
          const ref = t?.a.ref ? parseRange(t.a.ref.replace(/\$/g, '')) : undefined
          const name = t?.a.displayName ?? t?.a.name
          if (!t || !ref || !name || !(ref.r0 >= 0 && ref.c0 >= 0 && ref.r1 >= ref.r0 && ref.c1 >= ref.c0)) continue
          out.push({
            name,
            sheet: sheet.name,
            ...ref,
            headers: t.a.headerRowCount === '0' ? 0 : 1,
            totals: Number(t.a.totalsRowCount ?? 0) > 0 ? 1 : 0,
            columns: kids(kid(t, 'tableColumns'), 'tableColumn').map((c) => decodeColumn(c.a.name ?? '')),
          })
        }
        return out
      }),
    )
    this.names = names
    this.tables = perSheet.flat()
    this.pivots = await readPivots(this.pkg, this.sheets, this.date1904).catch(() => [])
    this.calc.namesChanged()
  }

  meta(): WorkbookMeta {
    return {
      sheets: this.sheets,
      styles: this.styles.xfs.map((x) => x.css),
      formats: this.styles.xfs.map((x) => x.numFmt),
      baseFont: this.styles.baseFont,
      date1904: this.date1904,
      ...(this.active ? { active: this.active } : {}),
    }
  }

  private loadSst(onProgress?: (loaded: number, total: number) => void): Promise<void> {
    if (!this.sstLoading) {
      this.sstLoading = (async () => {
        const rel = relOfType(this.rels, 'sharedStrings')
        if (!rel || !this.pkg.has(rel.target)) return
        const h = new SstHandler((a) => sheetColor({ a } as XEl, this.theme))
        this.sst = h.strings
        this.rich = h.rich
        await parseXmlStream(await this.pkg.stream(rel.target, onProgress), h)
      })()
    }
    return this.sstLoading
  }

  /**
   * Streams a sheet into memory, reporting progress so the UI can grow while parsing.
   * Calling again for a loaded sheet replays its state; for a sheet still parsing (for another caller) the state is
   * reported once the parse is done — its header isn't known before, and the view keeps the first header it gets.
   */
  loadSheet(index: number, emit: (p: SheetProgress) => void, onBytes?: (phase: string, loaded: number, total: number) => void): Promise<void> {
    const existing = this.sheetsData.get(index)
    if (existing) {
      if (existing.loading && !existing.done) return existing.loading.then(() => emit(this.snapshot(index, existing, rowSizesOf(existing), true)))
      emit(this.snapshot(index, existing, rowSizesOf(existing), true))
      return existing.loading ?? Promise.resolve()
    }
    const d = new SheetData()
    this.sheetsData.set(index, d)
    d.loading = (async () => {
      const info = this.sheets[index]
      if (!info) throw new Error(`No sheet #${index}`)
      try {
        await this.loadSst((l, t) => onBytes?.('strings', l, t))
        let sentHeader = false
        let last = 0
        let stale = false
        // the view grows while the sheet parses: the header as soon as it is known, then every 150 ms
        const report = async (ready: boolean, sizes: () => [number, number][]) => {
          const now = Date.now()
          if (!sentHeader && ready) {
            sentHeader = true
            emit(this.snapshot(index, d, sizes(), true))
            last = now
            await new Promise((r) => setTimeout(r, 0))
          } else if (now - last > 150) {
            last = now
            emit(this.snapshot(index, d, sizes(), false))
            await new Promise((r) => setTimeout(r, 0))
          }
        }
        const loader = this.loaders.get(info.path)
        let sizes: [number, number][] = []
        if (loader) {
          // the rows come from somewhere else than the sheet's XML (a text file being imported)
          const rep: LoadReport = {
            chunk: () => report(true, () => []),
            reset: () => {
              stale = true
            },
            bytes: (l, t) => onBytes?.('sheet', l, t),
          }
          await loader(d, rep)
        } else {
          const h = new SheetHandler(d, this.styles.xfs.map((x) => x.fontPt), this.styles.xfs[0]?.fontPt ?? 11, (cm) => this.dynamicCm.has(cm))
          await parseXmlStream(await this.pkg.stream(info.path, (l, t) => onBytes?.('sheet', l, t)), h, {
            onChunk: () => report(h.headerReady, () => h.newSizes.splice(0)),
          })
          sizes = h.newSizes.splice(0)
        }
        d.done = true
        emit({ ...this.snapshot(index, d, sizes, !sentHeader), merges: d.merges, ...(stale ? { recalc: true } : {}) })
        // files written by libraries often omit cached formula results, or hold zeros and say to compute: they are computed
        if (d.uncached.length || d.clock || (this.fullCalc && (d.formulas.size || d.depOf.size))) {
          void this.computeStale().then(() => {
            // (where the sheet is now: sheets may have been added or moved meanwhile; nowhere when it was deleted)
            const now = this.sheets.findIndex((s) => s.path === info.path)
            if (now >= 0) emit({ ...this.snapshot(now, d, [], false), recalc: true })
          })
        }
      } catch (e) {
        d.done = true
        d.error = (e as Error).message
        emit({ ...this.snapshot(index, d, [], true), error: d.error })
      }
    })()
    return d.loading
  }

  private stale?: Promise<void>
  /** told the values of a sheet changed after it was read (computed in the background): the views read its cells again */
  onValues?: (p: SheetProgress) => void

  /**
   * Computes the formulas the file gives no value for (all of them when it says to compute on load). In one pass
   * over the workbook, every sheet read first: a formula may read any sheet, and one computed before what it reads
   * would be wrong.
   */
  private computeStale(): Promise<void> {
    const run = async () => {
      await Promise.all(this.sheets.map((s) => this.loadSheet(s.index, () => {})))
      const cells: { sheet: number; key: number }[] = []
      let clock = false
      for (const [sheet, d] of this.sheetsData) {
        if (!d.done || d.error) continue
        if (this.fullCalc && !d.calculated) {
          expandShared(d)
          for (const key of d.formulas.keys()) cells.push({ sheet, key })
        } else for (const key of d.uncached) cells.push({ sheet, key })
        if (d.clock && !d.calculated) clock = true
        d.uncached = []
        d.calculated = true
      }
      if (!cells.length && !clock) return
      this.calc.invalidate()
      const sheets = this.calc.computeMissing(cells, clock)
      // what computing them built (a parsed formula per cell) is not kept for a document that may only be read
      this.calc.invalidate()
      // the sheets shown meanwhile show the values as they are now
      for (const i of sheets) {
        const d = this.sheetsData.get(i)
        if (d?.done) this.onValues?.({ ...this.snapshot(i, d, [], false), recalc: true })
      }
    }
    // one at a time; a sheet read after a pass (added, imported) gets the next
    const next = (this.stale ?? Promise.resolve()).then(run, run)
    this.stale = next
    void next.finally(() => {
      if (this.stale === next) this.stale = undefined
    })
    return next
  }

  /** resolves once the values being computed after reading sheets are in (`computeStale`) */
  async settled(): Promise<void> {
    while (this.stale) await this.stale.catch(() => {})
  }

  /** loads the sheets whose rows come from a loader (an imported text file) and are not in yet: saving writes them from the model */
  async loadPending(): Promise<void> {
    const pending = this.sheets.filter((s) => this.loaders.has(s.path) && !this.sheetsData.get(s.index)?.done)
    await Promise.all(pending.map((s) => this.loadSheet(s.index, () => {})))
  }

  private snapshot(sheet: number, d: SheetData, rowSizes: [number, number][], header: boolean): SheetProgress {
    const p: SheetProgress = { sheet, rows: d.maxRow + 1, cols: d.maxCol + 1, rowSizes, done: d.done }
    if (header) {
      p.colInfo = d.colInfo
      p.defaultColW = d.defaultColW
      p.defaultRowH = d.defaultRowH
      p.frozen = d.frozen
      p.gridLines = d.gridLines
    }
    if (d.done) p.merges = d.merges
    if (d.error) p.error = d.error
    return p
  }

  /** rows [r0, r1] of a sheet in memory (a sheet whose rows are read from its file by pages); throws when they cannot all be */
  async ensure(d: SheetData, r0: number, r1: number): Promise<void> {
    if (d.pages) await d.pages.ensure(Math.max(0, r0), r1)
  }

  /** Formatted cells for rows [r0, r1). */
  async block(sheet: number, r0: number, r1: number): Promise<BlockOut> {
    const d = this.sheetsData.get(sheet)
    const rows: RowOut[] = []
    if (!d) return { sheet, r0, r1, rows, complete: false }
    await this.ensure(d, r0, r1 - 1)
    const xfs = this.styles.xfs
    // (rows past the last cell may have a format of their own)
    const end = d.rowAttrs.size ? r1 : Math.min(r1, Math.max(d.maxRow + 1, d.rows.length))
    for (let r = r0; r < end; r++) {
      const row = d.rows[r]
      const x = d.rowAttrs.size ? d.rowAttrs.get(r)?.x : undefined
      const own = x && (x.customFormat === '1' || x.customFormat === 'true') ? Number(x.s ?? 0) || 0 : undefined
      if (!row) {
        if (own) rows.push({ r, cells: [], s: own })
        continue
      }
      const cells: CellOut[] = []
      for (let k = 0; k < row.length; k += 4) {
        const c = row[k] as number
        const t = row[k + 1] as number
        const v = row[k + 2]
        const s = row[k + 3] as number
        const xf = xfs[s] ?? xfs[0]
        switch (t) {
          case T_NUM: {
            const f = formatValue(v as number, xf.numFmt, this.date1904)
            cells.push(f.color ? [c, f.text, s, 1, f.color] : [c, f.text, s, 1])
            break
          }
          case T_SST: {
            const str = this.sst[v as number] ?? ''
            const text = formatText(str, xf.numFmt)
            const runs = text === str ? this.rich.get(v as number) : undefined
            cells.push(runs ? [c, text, s, 0, undefined, runs] : [c, text, s, 0])
            break
          }
          case T_STR:
            cells.push([c, formatText(v as string, xf.numFmt), s, 0])
            break
          case T_BOOL:
            cells.push([c, v ? 'TRUE' : 'FALSE', s, 2])
            break
          case T_ERR:
            cells.push([c, v as string, s, 2])
            break
          case T_DATE: {
            // a date a library wrote as ISO text: shown with the number format of the cell, else as the text
            const serial = xf.numFmt && xf.numFmt !== 'General' ? isoToSerial(v as string, this.date1904) : null
            if (serial === null) cells.push([c, (v as string).replace('T00:00:00', '').replace('T', ' ').replace(/Z$/, ''), s, 1])
            else {
              const f = formatValue(serial, xf.numFmt, this.date1904)
              cells.push(f.color ? [c, f.text, s, 1, f.color] : [c, f.text, s, 1])
            }
            break
          }
          default:
            cells.push([c, '', s, 0])
        }
      }
      rows.push(own ? { r, cells, s: own } : { r, cells })
    }
    // rows of a page the scan has not passed yet are not known
    const known = d.pages && !d.pages.complete ? Math.min(d.maxRow, d.pages.scanned - 1) : d.maxRow
    return { sheet, r0, r1, rows, complete: d.done || r1 <= known }
  }

  /**
   * The sheet as delimited text (the values as shown), a line per row of the used range less the rows and columns
   * with nothing shown at its end. A Blob: it can be larger than a string should be.
   */
  /** the sheet as delimited text; past `limit` characters (a lower bound of its UTF-8 bytes) it fails with a RangeError */
  async text(sheet: number, sep: string, limit = Infinity): Promise<Blob> {
    const d = this.sheetsData.get(sheet)
    if (!d?.done) throw new Error('Sheet is still loading')
    // the values the file gives no result for are being computed: the text has them
    await this.settled()
    let lastR = d.maxRow
    let lastC = d.maxCol
    if (!d.pages) {
      lastR = lastC = -1
      const last = Math.max(d.maxRow, d.rows.length - 1)
      for (let r = 0; r <= last; r++) {
        const row = d.rows[r]
        if (!row) continue
        for (let k = 0; k < row.length; k += 4) {
          const t = row[k + 1]
          if ((t === T_EMPTY || (t === T_STR && row[k + 2] === '')) && !d.formulas.has(r * MAX_COL + (row[k] as number))) continue
          lastR = r
          if ((row[k] as number) > lastC) lastC = row[k] as number
        }
      }
    }
    // a separator or a line break at least per cell: a sheet that cannot fit is refused before it is written
    const tooLarge = () => new RangeError(`The text of the sheet is larger than ${Math.round(limit / (1024 * 1024))} MB`)
    if ((lastR + 1) * (lastC + 2) > limit) throw tooLarge()
    let total = 0
    // the text goes into Blobs every few MB: a Blob lives outside the heap (and on disk, when large), a string does not
    const blobs: Blob[] = []
    // a byte order mark first: Excel on Windows reads a file without one in the system's code page (every separator)
    let parts: string[] = ['\uFEFF']
    let held = 0
    const step = d.pages?.size ?? 4096
    const blank = () => new Array<string>(lastC + 1).fill('')
    for (let r0 = 0; r0 <= lastR; r0 += step) {
      const r1 = Math.min(r0 + step, lastR + 1)
      const b = await this.block(sheet, r0, r1)
      const lines: string[][] = []
      let next = r0
      for (const row of b.rows) {
        for (; next < row.r; next++) lines.push(blank())
        const line = blank()
        const raw = d.rows[row.r]
        row.cells.forEach((c, i) => {
          if (c[0] > lastC) return
          line[c[0]] = csvField(c, raw?.[i * 4 + 1] === T_NUM && raw[i * 4] === c[0] ? (raw[i * 4 + 2] as number) : undefined, this.styles.xfs[c[2]]?.numFmt)
        })
        lines.push(line)
        next++
      }
      for (; next < r1; next++) lines.push(blank())
      const piece = toCsv(lines, sep)
      total += piece.length
      if (total > limit) throw tooLarge()
      parts.push(piece)
      held += piece.length
      if (held > 1 << 23) {
        blobs.push(new Blob(parts))
        parts = []
        held = 0
      }
    }
    if (parts.length) blobs.push(new Blob(parts))
    return new Blob(blobs, { type: sep === '\t' ? 'text/tab-separated-values' : 'text/csv' })
  }

  /** cells whose shown value (or formula) contains `query`, in row order */
  async find(sheet: number, query: string, opts: FindOptions = {}, limit = 20_000): Promise<FindMatch[]> {
    const d = this.sheetsData.get(sheet)
    const out: FindMatch[] = []
    if (!d || !query) return out
    const norm = (t: string) => (opts.matchCase ? t : t.toLowerCase())
    const q = norm(query)
    const xfs = this.styles.xfs
    if (opts.formulas) expandShared(d)
    const scan = (r: number, row: unknown[]) => {
      for (let k = 0; k < row.length && out.length < limit; k += 4) {
        const c = row[k] as number
        const t = row[k + 1] as number
        const v = row[k + 2]
        const f = opts.formulas ? d.formulas.get(r * MAX_COL + c) : undefined
        let text: string
        if (f) text = '=' + f
        else if (t === T_NUM) text = formatValue(v as number, (xfs[row[k + 3] as number] ?? xfs[0])?.numFmt, this.date1904).text
        else if (t === T_SST) text = this.sst[v as number] ?? ''
        else if (t === T_BOOL) text = v ? 'TRUE' : 'FALSE'
        else if (t === T_EMPTY) text = ''
        else text = String(v)
        if (!text) continue
        const h = norm(text)
        if (opts.whole ? h === q : h.includes(q)) out.push({ r, c, text })
      }
    }
    const last = Math.max(d.maxRow, d.rows.length - 1)
    if (!d.pages) {
      for (let r = 0; r <= last && out.length < limit; r++) {
        const row = d.rows[r]
        if (row) scan(r, row)
      }
    } else {
      // page by page, each read for the look only
      const size = d.pages.size
      for (let p = 0; p * size <= last && out.length < limit; p++) {
        const rows = await d.pages.peek(p)
        for (let k = 0; k < rows.length && out.length < limit; k++) {
          const row = rows[k]
          if (row) scan(p * size + k, row)
        }
      }
    }
    return out
  }

  async cell(sheet: number, r: number, c: number): Promise<CellDetail | undefined> {
    const d = this.sheetsData.get(sheet)
    if (d) await this.ensure(d, r, r)
    const row = d?.rows[r]
    // as it is typed: without the prefixes the file gives the functions newer than its format
    const text = d && formulaAt(d, r * MAX_COL + c)
    const formula = text && stripFunctionPrefixes(text)
    if (!row) return formula ? { value: '', formula: '=' + formula, type: 'formula' } : undefined
    for (let k = 0; k < row.length; k += 4) {
      if (row[k] !== c) continue
      const t = row[k + 1] as number
      const v = row[k + 2]
      const value = t === T_SST ? this.sst[v as number] ?? '' : t === T_BOOL ? (v ? 'TRUE' : 'FALSE') : String(v)
      const type = ['number', 'string', 'string', 'boolean', 'error', 'date', 'empty'][t]
      return { value, formula: formula ? '=' + formula : undefined, type }
    }
    return formula ? { value: '', formula: '=' + formula, type: 'formula' } : undefined
  }

  unload(sheet: number) {
    const d = this.sheetsData.get(sheet)
    if (d?.done) this.sheetsData.delete(sheet)
  }
}

/** pixel heights of every non-default row: explicit heights / hidden rows win over font-grown ones */
export function rowSizesOf(d: SheetData): [number, number][] {
  const out: [number, number][] = []
  for (const [r, px] of d.autoRowH) if (!d.rowAttrs.has(r)) out.push([r, px])
  for (const [r, a] of d.rowAttrs) if (a.ht !== undefined || a.hidden || d.autoRowH.has(r)) out.push([r, a.hidden ? 0 : a.ht !== undefined ? Math.round((a.ht * 4) / 3) : (d.autoRowH.get(r) ?? d.defaultRowH)])
  return out
}
