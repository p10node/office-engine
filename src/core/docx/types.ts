/** All lengths in points unless stated otherwise. */

export interface RunStyle {
  b?: boolean
  i?: boolean
  /** underline kind ('single', 'double', …) or 'none' */
  u?: string
  strike?: boolean
  dstrike?: boolean
  color?: string
  sz?: number
  font?: string
  hl?: string
  shd?: string
  va?: 'sup' | 'sub'
  caps?: boolean
  smallCaps?: boolean
  vanish?: boolean
  spacing?: number
}

export interface Border {
  w: number
  style: string
  color: string
}

export interface ParaStyle {
  align?: 'left' | 'center' | 'right' | 'justify'
  indL?: number
  indR?: number
  /** positive: first line indent; negative: hanging */
  indFirst?: number
  before?: number
  after?: number
  line?: number
  lineRule?: 'auto' | 'exact' | 'atLeast'
  shd?: string
  bTop?: Border
  bBottom?: Border
  bLeft?: Border
  bRight?: Border
  keepNext?: boolean
  pageBreakBefore?: boolean
  outline?: number
  contextual?: boolean
}

export type Inline =
  | { t: 'text'; text: string; s: RunStyle; href?: string }
  | { t: 'tab'; s: RunStyle }
  | { t: 'br'; s: RunStyle }
  | { t: 'img'; target: string; w: number; h: number; alt?: string; href?: string }
  | { t: 'field'; kind: 'PAGE' | 'NUMPAGES'; s: RunStyle }
  | { t: 'box'; w: number; h: number; blocks: Block[]; fill?: string; border?: string }
  | { t: 'shape'; w: number; h: number; label?: string }

export interface ListMarker {
  text: string
  s: RunStyle
  /** marker followed by tab (default) */
  tab: boolean
}

export interface Paragraph {
  t: 'p'
  ps: ParaStyle
  runs: Inline[]
  marker?: ListMarker
  /** font size of the paragraph mark, for empty paragraph height */
  markSz?: number
  markFont?: string
  /** explicit page break before this paragraph (from w:br type=page or pageBreakBefore) */
  pb?: boolean
}

export interface CellBorders {
  t?: string
  r?: string
  b?: string
  l?: string
}

export interface TableCell {
  blocks: Block[]
  colSpan: number
  rowSpan: number
  w?: number
  shd?: string
  vAlign?: 'top' | 'middle' | 'bottom'
  borders: CellBorders
  pad: [number, number, number, number]
}

export interface TableRow {
  cells: TableCell[]
  h?: number
  exact?: boolean
  header?: boolean
}

export interface Table {
  t: 'tbl'
  cols: number[]
  rows: TableRow[]
  align?: 'left' | 'center' | 'right'
  indL?: number
  w?: number
}

export interface PageSetup {
  w: number
  h: number
  mt: number
  mr: number
  mb: number
  ml: number
  header: number
  footer: number
}

export interface SectionBreak {
  t: 'sect'
  page: PageSetup
  /** 'nextPage' | 'continuous' | 'evenPage' | 'oddPage' */
  kind: string
}

export type Block = Paragraph | Table | SectionBreak

export interface DocxMeta {
  page: PageSetup
  header?: Block[]
  footer?: Block[]
  firstHeader?: Block[]
  firstFooter?: Block[]
  titlePg?: boolean
  /** body default font (CSS) and size */
  font?: string
  sz?: number
  /** final page setup known (end of document reached) */
  final?: boolean
}

export const DEFAULT_PAGE: PageSetup = { w: 612, h: 792, mt: 72, mr: 72, mb: 72, ml: 72, header: 36, footer: 36 }
