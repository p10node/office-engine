/** Geometry in px (96 dpi, EMU/9525); font sizes in pt. */

export type Fill =
  | { t: 'solid'; c: string }
  | { t: 'grad'; css: string }
  | { t: 'img'; target: string; tile?: boolean }
  | { t: 'none' }

export interface Stroke {
  c: string
  w: number
  dash?: string
  head?: string
  tail?: string
}

export interface Box {
  x: number
  y: number
  w: number
  h: number
  rot?: number
  flipH?: boolean
  flipV?: boolean
}

export interface TRun {
  text: string
  sz: number
  b?: boolean
  i?: boolean
  u?: boolean
  strike?: boolean
  color?: string
  font?: string
  baseline?: number
  caps?: boolean
  href?: string
  hl?: string
  br?: boolean
  /** slide field (number/date): kept as an object when editing */
  fld?: boolean
  /** [paragraph, run] index in the original shape XML */
  src?: [number, number]
}

export interface TPara {
  algn: 'left' | 'center' | 'right' | 'justify'
  marL: number
  indent: number
  /** line height multiplier or px */
  lnPct?: number
  lnPx?: number
  spcBef: number
  spcAft: number
  bullet?: { text: string; color?: string; font?: string; sz: number }
  /** outline level and bullet kind (editor state) */
  lvl?: number
  buKind?: 'none' | 'bullet' | 'number'
  runs: TRun[]
  endSz: number
  /** template paragraph index in the original shape XML */
  k?: number
}

export interface TextBody {
  anchor: 't' | 'ctr' | 'b'
  ins: [number, number, number, number]
  wrap: boolean
  vert?: string
  paras: TPara[]
  /** placeholder without text (editable, invisible when viewing) */
  empty?: boolean
}

/** `top`: id of the slide-level element (selection / move); `sid`: own id of a text shape (text editing) */
export type SlideEl = (
  | { t: 'shape'; box: Box; geom: string; adj?: Record<string, number>; fill?: Fill; stroke?: Stroke; text?: TextBody; path?: string; sid?: number }
  | { t: 'pic'; box: Box; target?: string; crop?: [number, number, number, number]; stroke?: Stroke; geom?: string; video?: boolean }
  | { t: 'table'; box: Box; cols: number[]; rows: { h: number; cells: TCell[] }[] }
  | { t: 'placeholder'; box: Box; label: string }
) & { top?: number }

/** 'text', 'line', 'image' or any DrawingML preset geometry name (rect, ellipse, star5, chevron…) */
export type ShapeKind = 'text' | 'line' | 'image' | (string & {})

export type ShapeOp =
  | { op: 'xfrm'; top: number; box: Box }
  | { op: 'delete'; top: number }
  | { op: 'fill'; top: number; color: string | null }
  | { op: 'line'; top: number; color?: string | null; w?: number; dash?: string }
  | { op: 'order'; top: number; to: 'front' | 'back' | 'forward' | 'backward' }
  | { op: 'duplicate'; top: number; dx?: number; dy?: number }
  /** vertical text anchor of a text shape (by shape id) */
  | { op: 'anchor'; sid: number; anchor: 't' | 'ctr' | 'b' }
  /** slide background color (null = follow layout) */
  | { op: 'bg'; color: string | null }
  | { op: 'insert'; kind: ShapeKind; box: Box; text?: string; rid?: string }

export type SlideOp =
  | { op: 'add'; after: number; layout?: 'title' | 'content' | 'blank' }
  | { op: 'duplicate'; index: number }
  | { op: 'delete'; index: number }
  | { op: 'move'; from: number; to: number }
  | { op: 'hide'; index: number; hidden: boolean }

export interface TCell {
  text?: TextBody
  fill?: Fill
  bl?: Stroke
  br?: Stroke
  bt?: Stroke
  bb?: Stroke
  colSpan: number
  rowSpan: number
  hidden: boolean
}

export interface SlideModel {
  index: number
  w: number
  h: number
  bg?: Fill
  els: SlideEl[]
  notes?: string
  hidden?: boolean
}

export interface PresentationMeta {
  w: number
  h: number
  count: number
  slides: { index: number; path: string; hidden: boolean }[]
}

export interface PRunSet {
  b?: boolean
  i?: boolean
  u?: boolean
  strike?: boolean
  sz?: number
  color?: string | null
  font?: string
  hl?: string | null
  /** percent of the font size (+30 superscript, -25 subscript), null = normal */
  baseline?: number | null
}

export interface PRunSpec {
  text: string
  src?: [number, number]
  set?: PRunSet
  fld?: boolean
}

export interface PParaSpec {
  k?: number
  algn?: 'left' | 'center' | 'right' | 'justify'
  /** outline level 0-8 */
  lvl?: number
  bullet?: 'none' | 'bullet' | 'number'
  /** line spacing multiplier */
  lnPct?: number
  runs: PRunSpec[]
}
