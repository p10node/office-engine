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
  runs: TRun[]
  endSz: number
}

export interface TextBody {
  anchor: 't' | 'ctr' | 'b'
  ins: [number, number, number, number]
  wrap: boolean
  vert?: string
  paras: TPara[]
}

export type SlideEl =
  | { t: 'shape'; box: Box; geom: string; adj?: Record<string, number>; fill?: Fill; stroke?: Stroke; text?: TextBody; path?: string }
  | { t: 'pic'; box: Box; target?: string; crop?: [number, number, number, number]; stroke?: Stroke; geom?: string; video?: boolean }
  | { t: 'table'; box: Box; cols: number[]; rows: { h: number; cells: TCell[] }[] }
  | { t: 'placeholder'; box: Box; label: string }

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
