import { familyKey, fontFamilies, type FontSource } from '../fonts'
import { FontError } from './container'
import { Font } from './font'
import { layoutText, type GlyphRun, type Inexact, type RunOptions, type Shaper } from './text'

/**
 * The font files a host declares (`FontSource`, the list DocViewer's `fonts` takes), read to measure text: which
 * file serves a piece of text set in a CSS font list at a weight and style, chosen as a browser chooses among
 * `@font-face` rules, and how wide the text is in it. Files are fetched and read when text first asks for them.
 */

/** why a measure is not what a browser would set, on top of a run's own reasons */
export type Unsure =
  | Inexact
  /** no declared font file serves the text: nothing is known about its width */
  | 'font'
  /** the file is of another weight or style than asked for: a browser thickens or slants its glyphs itself */
  | 'synthetic'
  /** the file is a variable font, asked for at another weight than the one its tables give widths for */
  | 'variable'

export interface TextSpec {
  text: string
  /** a CSS font-family list, as `cssFont()` builds for a document's font name: `"Calibri", Carlito, sans-serif` */
  font: string
  /** font size: the result's lengths are in its unit */
  size: number
  /** CSS weight (default 400) */
  weight?: number
  italic?: boolean
  /** pair kerning (default on) */
  kerning?: boolean
  /** standard ligatures (default on) */
  ligatures?: boolean
  /** added after every character, in the unit of `size` */
  letterSpacing?: number
}

/** a piece of measured text that one font file serves */
export interface TextPiece {
  /** the piece is `text.slice(from, to)` */
  from: number
  to: number
  /** the family the file is declared for; '' for text no declared file serves */
  family: string
  /** null for text no declared file serves */
  font: Font | null
  /** the glyphs, in font units; null without a font */
  run: GlyphRun | null
  /** in the unit of `size` */
  width: number
}

export interface TextMeasure {
  /** the sum of the pieces that could be measured, in the unit of `size` */
  width: number
  /**
   * How far lines are apart in the first font of the list that has a file ("the primary font" of CSS), in the unit of
   * `size`; all 0 when none has
   */
  ascent: number
  descent: number
  lineGap: number
  /** true when the width is what a browser sets for this text with these fonts */
  exact: boolean
  /** why it is not */
  unsure: Unsure[]
  pieces: TextPiece[]
}

interface Face {
  source: FontSource
  family: string
  /** the weights the file is declared for */
  min: number
  max: number
  style: 'normal' | 'italic' | 'oblique'
  /** unicode-range: null for all characters */
  ranges: [number, number][] | null
  font?: Promise<Font | null>
}

const weightOf = (w: FontSource['weight']): [number, number] => {
  const parts = String(w ?? 400)
    .trim()
    .split(/\s+/)
    .map((p) => (p === 'bold' ? 700 : p === 'normal' ? 400 : Number(p)))
    .filter((n) => Number.isFinite(n) && n > 0)
  if (!parts.length) return [400, 400]
  return [Math.min(...parts), Math.max(...parts)]
}

/** `U+26`, `U+0-7F`, `U+4??`, separated by commas: null when nothing in it is a range */
export function unicodeRanges(text: string | undefined): [number, number][] | null {
  if (!text) return null
  const out: [number, number][] = []
  for (const part of text.split(',')) {
    const m = /^\s*u\+([0-9a-f?]{1,6})(?:-([0-9a-f]{1,6}))?\s*$/i.exec(part)
    if (!m) continue
    if (m[2]) out.push([parseInt(m[1], 16), parseInt(m[2], 16)])
    else out.push([parseInt(m[1].replace(/\?/g, '0'), 16), parseInt(m[1].replace(/\?/g, 'f'), 16)])
  }
  return out.length ? out : null
}

/** how far `face` is from the weight asked for, as CSS orders the candidates (smaller is better) */
function weightRank(face: Face, want: number): number {
  if (want >= face.min && want <= face.max) return 0
  const w = want < face.min ? face.min : face.max
  // 400 and 500 look at each other first; lighter text then goes lighter, bolder text bolder
  if (want >= 400 && want <= 500) return w >= want && w <= 500 ? w - want : w < want ? 1000 + (want - w) : 2000 + (w - 500)
  if (want < 400) return w < want ? want - w : 1000 + (w - want)
  return w > want ? w - want : 1000 + (want - w)
}

const STYLE_ORDER = { italic: ['italic', 'oblique', 'normal'], normal: ['normal', 'oblique', 'italic'] } as const

export interface FontBookOptions {
  /** a shaping engine for the scripts that need one */
  shaper?: Shaper
  /** reads a font file by its address (default: `fetch`) */
  fetch?: (url: string) => Promise<ArrayBuffer | Uint8Array>
  /** told of a font file that could not be fetched or read */
  onError?: (error: Error, source: FontSource) => void
}

export class FontBook {
  private families = new Map<string, Face[]>()
  private stacks = new Map<string, Promise<{ family: string; font: Font; face: Face }[]>>()
  /** the families text asked for (the first of each list) when no declared file could serve it */
  readonly missing = new Set<string>()

  constructor(
    sources: readonly FontSource[],
    private opts: FontBookOptions = {},
  ) {
    for (const source of sources) {
      const family = familyKey(source.family)
      const [min, max] = weightOf(source.weight)
      const style = /italic/i.test(source.style ?? '') ? 'italic' : /oblique/i.test(source.style ?? '') ? 'oblique' : 'normal'
      const face: Face = { source, family, min, max, style, ranges: unicodeRanges(source.unicodeRange) }
      const list = this.families.get(family)
      if (list) list.push(face)
      else this.families.set(family, [face])
    }
  }

  /** fonts by their file: a file declared for several families (Calibri and Carlito) is read once, and is one font */
  private files = new Map<FontSource['src'], Promise<Font | null>>()

  private load(face: Face): Promise<Font | null> {
    const { src } = face.source
    const known = this.files.get(src)
    if (known) return (face.font ??= known)
    const p = (face.font ??= (async () => {
      try {
        let data: ArrayBuffer | ArrayBufferView | Blob
        if (typeof src !== 'string') data = src
        else if (this.opts.fetch) data = await this.opts.fetch(src)
        else {
          const res = await fetch(src)
          if (!res.ok) throw new FontError(`The font file ${src} could not be fetched (HTTP ${res.status})`)
          data = await res.arrayBuffer()
        }
        return await Font.open(data)
      } catch (e) {
        this.opts.onError?.(e instanceof Error ? e : new Error(String(e)), face.source)
        return null
      }
    })())
    this.files.set(src, p)
    return p
  }

  /**
   * The files that serve text in the font list `stack` at this weight and style, in the order of the list: for each
   * family that has declared files, the ones a browser would use (the nearest style, then the nearest weight; all the
   * files of that choice when the family comes in several, by `unicodeRange`).
   */
  private faces(stack: string, weight: number, italic: boolean): Promise<{ family: string; font: Font; face: Face }[]> {
    const key = `${weight}|${italic ? 1 : 0}|${stack}`
    let found = this.stacks.get(key)
    if (!found) {
      found = (async () => {
        const out: { family: string; font: Font; face: Face }[] = []
        for (const family of fontFamilies(stack)) {
          const all = this.families.get(family)
          if (!all) continue
          let pool: Face[] = []
          for (const style of STYLE_ORDER[italic ? 'italic' : 'normal']) {
            pool = all.filter((f) => f.style === style)
            if (pool.length) break
          }
          const best = Math.min(...pool.map((f) => weightRank(f, weight)))
          for (const face of pool) {
            if (weightRank(face, weight) !== best) continue
            const font = await this.load(face)
            if (font) out.push({ family, font, face })
          }
        }
        return out
      })()
      this.stacks.set(key, found)
    }
    return found
  }

  /**
   * Measures `spec.text`. Each character is set in the first file of the font list that is declared for it and has a
   * glyph for it, as a browser falls back from font to font; `pieces` are the runs that came out of one file.
   */
  async measure(spec: TextSpec): Promise<TextMeasure> {
    const weight = spec.weight ?? 400
    const italic = !!spec.italic
    const faces = await this.faces(spec.font, weight, italic)
    const unsure = new Set<Unsure>()
    const pieces: TextPiece[] = []
    const { text } = spec
    // which file each character is set in: -1 for none
    let from = 0
    let cur = -2
    const flush = (to: number) => {
      if (to <= from) return
      const f = cur >= 0 ? faces[cur] : null
      if (!f) {
        unsure.add(faces.length ? 'missing' : 'font')
        pieces.push({ from, to, family: '', font: null, run: null, width: 0 })
      } else {
        const scale = spec.size / f.font.unitsPerEm
        const opts: RunOptions = { kerning: spec.kerning, ligatures: spec.ligatures, letterSpacing: spec.letterSpacing ? spec.letterSpacing / scale : 0, shaper: this.opts.shaper }
        const run = layoutText(f.font, text.slice(from, to), opts)
        for (const r of run.inexact) unsure.add(r)
        if ((weight >= 600 && f.face.max < 600) || (italic && f.face.style === 'normal')) unsure.add('synthetic')
        // a variable font's tables give the widths of one weight, its own: a browser sets the text at the weight asked
        // for, kept within the weights the file is declared for (one weight: that one, whatever is asked for)
        if (f.font.variable && Math.min(Math.max(weight, f.face.min), f.face.max) !== f.font.weight) unsure.add('variable')
        pieces.push({ from, to, family: f.family, font: f.font, run, width: run.width * scale })
      }
      from = to
    }
    for (let i = 0; i < text.length; ) {
      const cp = text.codePointAt(i)!
      let k = faces.findIndex((f) => (!f.face.ranges || f.face.ranges.some(([a, b]) => cp >= a && cp <= b)) && f.font.glyphOf(cp) !== 0)
      // what is drawn with no glyph of its own (a soft hyphen, a joiner) or attaches to the character before it stays in that character's font
      if (k < 0 && cur >= 0 && (cp === 0xad || (cp >= 0x200b && cp <= 0x200f) || cp === 0x2060 || cp === 0xfeff || (cp >= 0x300 && cp <= 0x36f) || (cp >= 0xfe00 && cp <= 0xfe0f))) k = cur
      if (k !== cur) {
        flush(i)
        cur = k
      }
      i += cp > 0xffff ? 2 : 1
    }
    flush(text.length)
    if (!faces.length) {
      const first = fontFamilies(spec.font)[0]
      if (first) this.missing.add(first)
    }
    const primary = faces[0]?.font
    const line = primary?.lineMetrics()
    const scale = primary ? spec.size / primary.unitsPerEm : 0
    return {
      width: pieces.reduce((a, p) => a + p.width, 0),
      ascent: line ? line.ascent * scale : 0,
      descent: line ? line.descent * scale : 0,
      lineGap: line ? line.lineGap * scale : 0,
      exact: !unsure.size,
      unsure: [...unsure],
      pieces,
    }
  }
}
