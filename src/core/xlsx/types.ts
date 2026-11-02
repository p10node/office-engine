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
  /** CSS for font 0 (grid default) */
  baseFont: CellCss
  date1904: boolean
}

export interface ColInfo {
  min: number
  max: number
  /** px; 0 when hidden */
  w: number
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
