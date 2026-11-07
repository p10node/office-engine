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

/** [paragraph id, run ordinal] of the source run in the original XML */
export type SrcRef = [number, number]

/**
 * Direct formatting changed by the user (written back as explicit run properties).
 * null: the run stops setting the property itself, the paragraph style decides.
 */
export interface RunSet {
  b?: boolean | null
  i?: boolean | null
  u?: string | null
  strike?: boolean | null
  color?: string | null
  sz?: number | null
  /** raw font name */
  font?: string | null
  hl?: string | null
  va?: 'sup' | 'sub' | null
}

/**
 * Inline content. Tabs and line breaks are carried as '\t' / '\n' inside text.
 * Everything that isn't 'text' counts as one character when editing.
 * `cm`: the comment threads (root ids) whose range covers the run.
 */
export type Inline =
  | { t: 'text'; text: string; s: RunStyle; href?: string; lk?: string; src?: SrcRef; d?: RunSet; cm?: string[] }
  | {
      t: 'img'
      cm?: string[]
      target: string
      w: number
      h: number
      alt?: string
      href?: string
      src?: SrcRef
      rid?: string
      resized?: boolean
      /** alt text changed by the user */
      altSet?: boolean
      /** picture replaced (rid points to the new image part) */
      repl?: boolean
    }
  | { t: 'field'; kind: 'PAGE' | 'NUMPAGES'; s: RunStyle; src?: SrcRef }
  | { t: 'box'; w: number; h: number; blocks: Block[]; fill?: string; border?: string; src?: SrcRef }
  | { t: 'shape'; w: number; h: number; label?: string; src?: SrcRef }
  | { t: 'ref'; text: string; s: RunStyle; src?: SrcRef }

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
  /** paragraph id (editable paragraphs only); split parts share it */
  pid?: number
  /** read-only (inside fields spanning paragraphs, text boxes, nested tables…) */
  ro?: boolean
  /** paragraph style id */
  sid?: string
  /** list kind + level when numbered */
  lst?: 'bullet' | 'number'
  ilvl?: number
  // ---- editor bookkeeping (main thread) ----
  /** template paragraph (original pid) */
  from?: number
  /** original paragraphs whose bookmarks/comments this one carries */
  orig?: number[]
  /** paragraph-level changes */
  pd?: ParaDirty
  /** changed since load */
  edited?: boolean
  /** keeps the template's section break */
  keepSect?: boolean
  /** render version (bumped to force DOM rebuild after IME composition) */
  ver?: number
  /** formatting for text typed into this (empty) paragraph */
  caret?: Inline
}

/** null (alignment, spacing): the paragraph stops setting the property itself, its style decides. */
export interface ParaDirty {
  style?: string
  align?: 'left' | 'center' | 'right' | 'justify' | null
  list?: 'bullet' | 'number' | 'none'
  ilvl?: number
  indL?: number
  /** line spacing multiple (1, 1.15, 1.5, 2…) */
  line?: number | null
  /** spacing before / after in pt */
  before?: number | null
  after?: number | null
  pageBreak?: boolean
}

export type SpecRun =
  | { text: string; src?: SrcRef; set?: RunSet; lk?: string; cm?: string[] }
  | { obj: 'copy'; src: SrcRef; size?: [number, number]; alt?: string; rid?: string; cm?: string[] }
  | { obj: 'image'; rid: string; w: number; h: number; name?: string; alt?: string; cm?: string[] }
  | { obj: 'field'; kind: 'PAGE' | 'NUMPAGES'; set?: RunSet; cm?: string[] }

export interface ParaSpec {
  pid: number
  from: number
  orig?: number[]
  pd?: ParaDirty
  /** page-break separated parts */
  parts: SpecRun[][]
  pristine?: boolean
  keepSect?: boolean
}

export interface CellSpec {
  src?: [number, number]
  colSpan: number
  rowSpan: number
  paras: ParaSpec[]
  shd?: string
}

export interface TableSpec {
  tid: number
  item?: number
  cols: number[]
  rows: { src?: number; cells: CellSpec[] }[]
}

export interface TableOp {
  tid: number
  item?: number
  /** paragraph the (new) table follows */
  anchor?: number
  /** null deletes the table */
  spec: TableSpec | null
}

export interface EditOp {
  anchor: number
  list: ParaSpec[]
  remove?: number[]
  /** return re-converted blocks (needed after style/list changes) */
  want?: boolean
}

export interface EditResult {
  blocks: Record<number, Block[]>
  markers: Record<number, { text: string; s: RunStyle; tab: boolean; lst?: 'bullet' | 'number' } | null>
}

/** built-in paragraph styles the picker always offers (created on first use when missing) */
export type CanonStyle = 'normal' | 'title' | 'subtitle' | 'h1' | 'h2' | 'h3' | 'h4' | 'h5' | 'h6'

export interface StyleInfo {
  id: string
  name: string
  next?: string
  canon?: CanonStyle
  /** heading level 0-8 (outline) */
  outline?: number
  /** run look of the style, for the style picker */
  preview?: { b?: boolean; i?: boolean; color?: string; sz?: number; font?: string }
}

/**
 * The look given to a paragraph style ("update style to match"). Only the properties present are touched;
 * null: the style stops defining the property and takes it from the style it is based on.
 */
export interface StyleLook {
  b?: boolean | null
  i?: boolean | null
  /** underline kind, 'none' for no underline */
  u?: string | null
  strike?: boolean | null
  /** '#rrggbb', or 'auto' */
  color?: string | null
  sz?: number | null
  /** raw font name */
  font?: string | null
  align?: 'left' | 'center' | 'right' | 'justify' | null
  /** line spacing: a multiple, or pt when `lineRule` says so */
  line?: number | null
  lineRule?: 'exact' | 'atLeast'
  /** spacing before / after in pt */
  before?: number | null
  after?: number | null
}

export type HeaderFooters = Pick<DocxMeta, 'header' | 'footer' | 'firstHeader' | 'firstFooter'>

export interface StyleSetResult {
  styles: StyleInfo[]
  /** paragraph styles that look different now (the style and those based on it); empty when nothing changed */
  changed: string[]
  /** the look that puts the style back as it was */
  undo: StyleLook
  /** the document has no such paragraph style */
  missing?: boolean
  /** headers / footers converted again, when the document has any */
  headers?: HeaderFooters
}

export interface CellBorders {
  t?: string
  r?: string
  b?: string
  l?: string
}

export interface TableCell {
  blocks: Block[]
  /** [row, tc] index in the original table XML (template for cell properties) */
  src?: [number, number]
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
  /** original row index */
  src?: number
  h?: number
  exact?: boolean
  header?: boolean
}

export interface Table {
  t: 'tbl'
  /** body item index of an original top-level table */
  item?: number
  /** set once the table is regenerated by the editor (new or structurally changed) */
  tid?: number
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
  /** quick paragraph styles for the style picker */
  styles?: StyleInfo[]
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
