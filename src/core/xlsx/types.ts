export interface SheetInfo {
  name: string
  index: number
  path: string
  hidden: boolean
}

/** CSS-ish props for a cell format (cellXfs entry). Keys are React style keys. */
export type CellCss = Record<string, string | number>

export interface WorkbookMeta {
  sheets: SheetInfo[]
  styles: CellCss[]
  /** number format code per style */
  formats: string[]
  /** CSS for font 0 (grid default) */
  baseFont: CellCss
  date1904: boolean
}

export interface ColInfo {
  min: number
  max: number
  /** px; 0 when hidden */
  w: number
  /** width before hiding (px) */
  prevW?: number
}

export interface Merge {
  r0: number
  c0: number
  r1: number
  c1: number
}

export interface SheetProgress {
  sheet: number
  /** number of rows known so far (max row index + 1) */
  rows: number
  cols: number
  /** custom row sizes discovered since last progress: [row, px] (0 = hidden) */
  rowSizes: [number, number][]
  done: boolean
  /** only in the first progress message */
  colInfo?: ColInfo[]
  defaultColW?: number
  defaultRowH?: number
  frozen?: { rows: number; cols: number }
  gridLines?: boolean
  /** only when done */
  merges?: Merge[]
  error?: string
  /** formula values were (re)computed: cached rows are stale */
  recalc?: boolean
}

/** [col, text, styleIndex, kind (0 text,1 number,2 center), color?] */
export type CellOut = [number, string, number, number, string?]

export interface RowOut {
  r: number
  cells: CellOut[]
}

export interface BlockOut {
  sheet: number
  r0: number
  r1: number
  rows: RowOut[]
  complete: boolean
}

export interface CellDetail {
  value: string
  formula?: string
  type: string
}

export interface CellInput {
  r: number
  c: number
  /** raw user input ('' clears; '=…' formula) */
  input: string
}

export type BorderKind = 'all' | 'outer' | 'inner' | 'none' | 'top' | 'bottom' | 'left' | 'right'

export interface FmtPatch {
  border?: BorderKind
  borderColor?: string
  b?: boolean
  i?: boolean
  u?: boolean
  strike?: boolean
  color?: string | null
  fill?: string | null
  align?: 'left' | 'center' | 'right' | null
  valign?: 'top' | 'center' | 'bottom' | null
  /** text rotation in degrees (-90…90), 255 = vertical stack, null = none */
  rotate?: number | null
  wrap?: boolean
  numFmt?: string
  sz?: number
  font?: string
  borderStyle?: 'thin' | 'medium' | 'thick' | 'dashed' | 'dotted' | 'double'
  /** reset to the default style */
  clear?: boolean
  /** apply this exact style index (paint format) */
  xf?: number
}

export interface Range {
  r0: number
  c0: number
  r1: number
  c1: number
}

export interface XlsxEditResult {
  sheet: number
  range: Range
  /** style entries appended to the workbook (index → css/format) */
  styles?: { index: number; css: CellCss; format: string }[]
  /** full layout after structural edits (sizes, merges, frozen panes…) */
  layout?: SheetProgress
  /** sheet list changed */
  sheets?: SheetInfo[]
  /** sheets whose formula values changed (cached rows are stale) */
  recalc?: number[]
}

export type SheetOp =
  | { op: 'insertRows' | 'deleteRows' | 'insertCols' | 'deleteCols'; at: number; n: number }
  | { op: 'merge' | 'unmerge'; range: Range; mode?: 'all' | 'across' | 'down' }
  | { op: 'hideRows' | 'hideCols'; a: number; b: number; hidden: boolean }
  | { op: 'gridlines'; show: boolean }
  | { op: 'colWidth'; c0: number; c1: number; px: number }
  | { op: 'rowHeight'; r0: number; r1: number; px: number }
  | { op: 'freeze'; rows: number; cols: number }
  | { op: 'fill'; range: Range; dir: 'down' | 'right' }
  | { op: 'sort'; range: Range; col: number; desc?: boolean }

export type BookOp =
  | { op: 'add'; name?: string; after?: number }
  | { op: 'rename'; index: number; name: string }
  | { op: 'delete'; index: number }
  | { op: 'move'; from: number; to: number }

export interface FindOptions {
  matchCase?: boolean
  /** whole cell contents */
  whole?: boolean
  /** search formula text instead of values */
  formulas?: boolean
}

export interface FindMatch {
  r: number
  c: number
  /** text the match was found in (value or formula) */
  text: string
}
