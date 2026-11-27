/**
 * Pages of drawing commands: what a layout makes of a document (a slide, a sheet's pages, a document's pages), and
 * what a backend paints (PDF, SVG). Lengths are CSS px (1/96 in), y down, as in the document models; a PDF backend
 * scales them to points.
 */
import type { Font } from '../font/font'

/** [a, b, c, d, e, f]: x' = a·x + c·y + e, y' = b·x + d·y + f (as SVG and PDF) */
export type Matrix = [number, number, number, number, number, number]

/** red, green, blue 0-255, alpha 0-1 */
export interface Rgba {
  r: number
  g: number
  b: number
  a: number
}

/** a path: moves, lines, cubic curves and closes, in absolute coordinates */
export type Seg = { k: 'M' | 'L'; x: number; y: number } | { k: 'C'; x1: number; y1: number; x2: number; y2: number; x: number; y: number } | { k: 'Z' }

export interface Stop {
  /** 0-1 */
  at: number
  color: Rgba
}

export type Paint =
  | { t: 'color'; color: Rgba }
  /** from (x0, y0) to (x1, y1), in the coordinates of the item painted */
  | { t: 'linear'; x0: number; y0: number; x1: number; y1: number; stops: Stop[] }
  /** circles from (fx, fy) r 0 to (cx, cy) r */
  | { t: 'radial'; cx: number; cy: number; r: number; fx: number; fy: number; stops: Stop[] }

export interface StrokeStyle {
  paint: Paint
  width: number
  dash?: number[]
  cap?: 'butt' | 'round' | 'square'
  join?: 'miter' | 'round' | 'bevel'
}

/** the fonts a PDF has without a file: text no declared file serves is set in these (characters of Windows-1252 only) */
export type StdFont = 'Helvetica' | 'Helvetica-Bold' | 'Helvetica-Oblique' | 'Helvetica-BoldOblique'

/** a picture's bytes: JPEG and PNG are drawn; another format is skipped by the backend that cannot draw it */
export interface ImageData {
  bytes: Uint8Array
  /** the MIME type, from the file's first bytes */
  type: string
  /** a key the backend shares the picture by when it is drawn more than once (its part name) */
  key?: string
}

export type DrawItem =
  | { t: 'path'; d: Seg[]; fill?: Paint; rule?: 'nonzero' | 'evenodd'; stroke?: StrokeStyle }
  | {
      t: 'text'
      /** the pen at the start of the baseline */
      x: number
      y: number
      /** px */
      size: number
      /** a font file, or a PDF's own font for text no file serves (then `glyphs` are Windows-1252 codes) */
      font: Font | StdFont
      glyphs: number[]
      /** how far the pen moves after each glyph, px */
      advances: number[]
      /** glyph offsets, px (right and up); absent when all are 0 */
      dx?: number[]
      dy?: number[]
      /**
       * the text the glyphs stand for (all of it, and nothing more), and for each glyph the index of its first character
       * (copy and search in a PDF)
       */
      text: string
      clusters: number[]
      fill: Rgba
      /** a regular file stands in for bold or italic text: the backend thickens or slants it */
      synthetic?: { bold?: boolean; italic?: boolean }
    }
  /** the picture stretched over x, y, w, h */
  | { t: 'image'; image: ImageData; x: number; y: number; w: number; h: number; opacity?: number }
  | { t: 'group'; transform?: Matrix; clip?: Seg[]; clipRule?: 'nonzero' | 'evenodd'; opacity?: number; items: DrawItem[] }
  /** a link over an area (not drawn) */
  | { t: 'link'; x: number; y: number; w: number; h: number; url: string }

export interface Page {
  /** px */
  w: number
  h: number
  items: DrawItem[]
  /** the page's title (a slide's): a PDF's outline */
  title?: string
}
