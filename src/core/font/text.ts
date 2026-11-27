import type { Font } from './font'
import { hasKerning, position, substitute } from './layout'

/**
 * A run of text in one font as glyphs: which ones, how far the pen moves after each, and whether that is what a
 * browser or Word would set. Plain left-to-right text (Latin, Greek, Cyrillic, Armenian, Georgian, Chinese,
 * Japanese, Korean syllables) is laid out here from the font's tables. Text that needs a shaping engine (Arabic,
 * Hebrew, the Indic scripts, Thai, combining marks, emoji sequences…) is handed to the host's `Shaper` when there is
 * one, and otherwise laid out glyph by glyph and said not to be exact.
 */

/** A shaping engine the host supplies (HarfBuzz compiled to WebAssembly, for one): the library ships none. */
export interface Shaper {
  /**
   * Shapes `text`, all of one script and one direction, in the font file `font` (a bare TTF / OTF).
   * Returns the glyphs in the order they are drawn, left to right, with lengths in the font's units.
   * `script`: an ISO 15924 code ("Arab", "Deva"); `features`: OpenType features to turn off or on ("-kern", "-liga").
   */
  shape(font: Uint8Array, text: string, opts: { script: string; direction: 'ltr' | 'rtl'; features?: string[] }): ShapedGlyph[]
}

export interface ShapedGlyph {
  glyph: number
  /** how far the pen moves after the glyph */
  advance: number
  /** how far the glyph is drawn from the pen: right and up */
  dx: number
  dy: number
  /** the index in the text given (UTF-16) of the first character the glyph stands for */
  cluster: number
}

/** why a run is not what a shaping engine would set */
export type Inexact =
  /** a character the font has no glyph for: a browser draws it in another font */
  | 'missing'
  /** text of a script that needs shaping, and no shaper */
  | 'shaping'
  /** the font has rules for these glyphs that are not applied here (contextual forms, mark placement) */
  | 'features'
  /** the font's kerning is in a table of a kind that is not read */
  | 'kerning'

export interface GlyphRun {
  glyphs: number[]
  /** how far the pen moves after each glyph (font units): the glyph's width, kerning and letter spacing */
  advances: number[]
  /** how far each glyph is drawn to the right of the pen, and above the baseline (font units; nearly always 0) */
  dx: number[]
  dy: number[]
  /** for each glyph, the index in the text (UTF-16) of the first character it stands for */
  clusters: number[]
  /** the sum of `advances` */
  width: number
  /** characters the font has no glyph for */
  missing: number
  /** true when this is what a browser's shaping gives for the text in this font */
  exact: boolean
  /** why it is not */
  inexact: Inexact[]
}

export interface RunOptions {
  /** pair kerning (default on, as in browsers; Word has it off unless the text asks for it) */
  kerning?: boolean
  /** standard ligatures such as fi and fl (default on, as in browsers; off in Word unless asked for) */
  ligatures?: boolean
  /** added after every character, in font units. Like CSS, any spacing turns ligatures off */
  letterSpacing?: number
  shaper?: Shaper
}

/** [first, last, OpenType script tag, ISO 15924 code]: the scripts laid out here, in code point order */
const SIMPLE: [number, number, string, string][] = [
  [0x41, 0x5a, 'latn', 'Latn'],
  [0x61, 0x7a, 'latn', 'Latn'],
  [0xaa, 0xaa, 'latn', 'Latn'],
  [0xba, 0xba, 'latn', 'Latn'],
  [0xc0, 0xd6, 'latn', 'Latn'],
  [0xd8, 0xf6, 'latn', 'Latn'],
  [0xf8, 0x2b8, 'latn', 'Latn'],
  [0x370, 0x373, 'grek', 'Grek'],
  [0x375, 0x377, 'grek', 'Grek'],
  [0x37a, 0x37d, 'grek', 'Grek'],
  [0x37f, 0x3ff, 'grek', 'Grek'],
  [0x400, 0x482, 'cyrl', 'Cyrl'],
  [0x48a, 0x52f, 'cyrl', 'Cyrl'],
  [0x531, 0x58f, 'armn', 'Armn'],
  [0x10a0, 0x10ff, 'geor', 'Geor'],
  [0x1c80, 0x1c8f, 'cyrl', 'Cyrl'],
  [0x1c90, 0x1cbf, 'geor', 'Geor'],
  [0x1d00, 0x1d25, 'latn', 'Latn'],
  [0x1e00, 0x1eff, 'latn', 'Latn'],
  [0x1f00, 0x1fff, 'grek', 'Grek'],
  [0x2c60, 0x2c7f, 'latn', 'Latn'],
  [0x2e80, 0x2fdf, 'hani', 'Hani'],
  [0x3005, 0x3007, 'hani', 'Hani'],
  [0x3021, 0x3029, 'hani', 'Hani'],
  [0x3038, 0x303b, 'hani', 'Hani'],
  [0x3041, 0x3096, 'kana', 'Hira'],
  [0x309d, 0x309f, 'kana', 'Hira'],
  [0x30a1, 0x30fa, 'kana', 'Kana'],
  [0x30fd, 0x30ff, 'kana', 'Kana'],
  [0x3131, 0x318e, 'hang', 'Hang'],
  [0x31f0, 0x31ff, 'kana', 'Kana'],
  [0x3400, 0x4dbf, 'hani', 'Hani'],
  [0x4e00, 0x9fff, 'hani', 'Hani'],
  [0xa640, 0xa66e, 'cyrl', 'Cyrl'],
  [0xa720, 0xa7ff, 'latn', 'Latn'],
  [0xac00, 0xd7a3, 'hang', 'Hang'],
  [0xf900, 0xfaff, 'hani', 'Hani'],
  [0xfb00, 0xfb06, 'latn', 'Latn'],
  [0xfb13, 0xfb17, 'armn', 'Armn'],
  [0xff21, 0xff3a, 'latn', 'Latn'],
  [0xff41, 0xff5a, 'latn', 'Latn'],
  [0xff66, 0xff9f, 'kana', 'Kana'],
  [0x20000, 0x323af, 'hani', 'Hani'],
]

/**
 * [first, last, ISO 15924 code, right to left]: what needs a shaping engine. Scripts whose letters join, reorder or
 * stack; and, under `Zinh`, what attaches to the character before it: combining marks, joiners, variation selectors,
 * emoji modifiers.
 */
const SHAPED: [number, number, string, boolean][] = [
  [0x300, 0x36f, 'Zinh', false],
  [0x483, 0x489, 'Zinh', false],
  [0x591, 0x5ff, 'Hebr', true],
  [0x600, 0x6ff, 'Arab', true],
  [0x700, 0x74f, 'Syrc', true],
  [0x750, 0x77f, 'Arab', true],
  [0x780, 0x7bf, 'Thaa', true],
  [0x7c0, 0x7ff, 'Nkoo', true],
  [0x800, 0x83f, 'Samr', true],
  [0x840, 0x85f, 'Mand', true],
  [0x860, 0x86f, 'Syrc', true],
  [0x870, 0x8ff, 'Arab', true],
  [0x900, 0x97f, 'Deva', false],
  [0x980, 0x9ff, 'Beng', false],
  [0xa00, 0xa7f, 'Guru', false],
  [0xa80, 0xaff, 'Gujr', false],
  [0xb00, 0xb7f, 'Orya', false],
  [0xb80, 0xbff, 'Taml', false],
  [0xc00, 0xc7f, 'Telu', false],
  [0xc80, 0xcff, 'Knda', false],
  [0xd00, 0xd7f, 'Mlym', false],
  [0xd80, 0xdff, 'Sinh', false],
  [0xe00, 0xe7f, 'Thai', false],
  [0xe80, 0xeff, 'Laoo', false],
  [0xf00, 0xfff, 'Tibt', false],
  [0x1000, 0x109f, 'Mymr', false],
  [0x1100, 0x11ff, 'Hang', false],
  [0x1780, 0x17ff, 'Khmr', false],
  [0x1800, 0x18af, 'Mong', false],
  [0x1900, 0x1aaf, 'Zinh', false],
  [0x1ab0, 0x1aff, 'Zinh', false],
  [0x1b00, 0x1c4f, 'Zinh', false],
  [0x1cd0, 0x1cff, 'Zinh', false],
  [0x1dc0, 0x1dff, 'Zinh', false],
  [0x200c, 0x200f, 'Zinh', false],
  [0x202a, 0x202e, 'Zinh', false],
  [0x2066, 0x2069, 'Zinh', false],
  [0x20d0, 0x20ff, 'Zinh', false],
  [0x3099, 0x309a, 'Zinh', false],
  [0xa800, 0xabff, 'Zinh', false],
  [0xd7b0, 0xd7ff, 'Hang', false],
  [0xfb1d, 0xfb4f, 'Hebr', true],
  [0xfb50, 0xfdff, 'Arab', true],
  [0xfe00, 0xfe0f, 'Zinh', false],
  [0xfe20, 0xfe2f, 'Zinh', false],
  [0xfe70, 0xfefe, 'Arab', true],
  [0x10800, 0x10fff, 'Zinh', true],
  [0x11000, 0x11fff, 'Zinh', false],
  [0x1e800, 0x1efff, 'Zinh', true],
  [0x1f1e6, 0x1f1ff, 'Zinh', false],
  [0x1f3fb, 0x1f3ff, 'Zinh', false],
  [0xe0000, 0xe01ef, 'Zinh', false],
]

function find<T extends [number, number, ...unknown[]]>(table: T[], cp: number): T | undefined {
  let lo = 0
  let hi = table.length - 1
  while (lo <= hi) {
    const mid = (lo + hi) >> 1
    if (table[mid][1] < cp) lo = mid + 1
    else if (table[mid][0] > cp) hi = mid - 1
    else return table[mid]
  }
  return undefined
}

/** true when `text` has characters that need a shaping engine to be laid out as a browser does */
export function needsShaping(text: string): boolean {
  for (let i = 0; i < text.length; i++) {
    const c = text.charCodeAt(i)
    // (nothing below U+0300 does)
    if (c >= 0x300 && find(SHAPED, text.codePointAt(i)!)) return true
  }
  return false
}

/** `to` made what `from` is, in place (not spread into a call: a long text has more glyphs than a call takes arguments) */
function refill(to: number[], from: readonly number[]): void {
  to.length = from.length
  for (let i = 0; i < from.length; i++) to[i] = from[i]
}

/** characters that take no room and draw nothing: the soft hyphen, zero-width space, word joiner, byte order mark */
const hidden = (cp: number) => cp === 0xad || cp === 0x200b || cp === 0x2060 || cp === 0xfeff

interface Segment {
  from: number
  to: number
  /** OpenType script tag of the font's rules to use */
  tag: string
  /** ISO 15924 code, for a shaper */
  iso: string
  shaped: boolean
  rtl: boolean
}

/**
 * `text` cut where its script changes. Characters of no script of their own (spaces, digits, punctuation) go with
 * the letters before them, or with the first letters when the text starts with them: as browsers cut text before
 * they shape it. What attaches to a character (a combining mark) makes its piece one that needs shaping.
 */
export function segments(text: string): Segment[] {
  const out: Segment[] = []
  let cur: Segment | null = null
  for (let i = 0; i < text.length; ) {
    const cp = text.codePointAt(i)!
    const next = i + (cp > 0xffff ? 2 : 1)
    const shaped = cp >= 0x300 ? find(SHAPED, cp) : undefined
    const simple = shaped ? undefined : find(SIMPLE, cp)
    if (shaped && shaped[2] === 'Zinh') {
      // attached to what is before it: that piece now needs shaping (and is right to left when this is)
      if (!cur) cur = { from: i, to: next, tag: 'DFLT', iso: 'Zyyy', shaped: true, rtl: shaped[3] }
      cur.shaped = true
      cur.to = next
    } else if (shaped || simple) {
      const iso = shaped ? shaped[2] : simple![3]
      const tag = shaped ? 'DFLT' : simple![2]
      if (cur && cur.iso === 'Zyyy') {
        // the text began with characters of no script: they are of this one
        cur.iso = iso
        cur.tag = tag
        cur.shaped ||= !!shaped
        cur.rtl ||= !!shaped?.[3]
        cur.to = next
      } else if (cur && cur.iso === iso) cur.to = next
      else {
        if (cur) out.push(cur)
        cur = { from: i, to: next, tag, iso, shaped: !!shaped, rtl: !!shaped?.[3] }
      }
    } else if (cur) cur.to = next
    else cur = { from: i, to: next, tag: 'DFLT', iso: 'Zyyy', shaped: false, rtl: false }
    i = next
  }
  if (cur) out.push(cur)
  return out
}

/** Lays out `text` in `font`. See `GlyphRun` for what comes back, and `exact` in it for how far it can be trusted. */
export function layoutText(font: Font, text: string, opts: RunOptions = {}): GlyphRun {
  const run: GlyphRun = { glyphs: [], advances: [], dx: [], dy: [], clusters: [], width: 0, missing: 0, exact: true, inexact: [] }
  const why = new Set<Inexact>()
  const kerning = opts.kerning !== false
  const spacing = opts.letterSpacing ?? 0
  const ligatures = opts.ligatures !== false && !spacing
  for (const seg of segments(text)) {
    const piece = text.slice(seg.from, seg.to)
    // characters and their glyphs; the hidden ones are taken out, so that what is around them kerns and joins as if they were not there
    const ids: number[] = []
    const clusters: number[] = []
    const zero: number[] = []
    let missing = 0
    for (let i = 0; i < piece.length; ) {
      const cp = piece.codePointAt(i)!
      if (hidden(cp)) zero.push(seg.from + i)
      else {
        const g = font.glyphOf(cp)
        if (!g) missing++
        ids.push(g)
        clusters.push(seg.from + i)
      }
      i += cp > 0xffff ? 2 : 1
    }
    run.missing += missing
    if (missing) why.add('missing')
    let advances: number[]
    let dx: number[]
    let dy: number[] | null = null
    let glyphs = ids
    let own = true
    const tryShaper = () => {
      if (!opts.shaper || missing) return false
      const features = [...(kerning ? [] : ['-kern']), ...(ligatures ? [] : ['-liga', '-clig'])]
      const shaped = opts.shaper.shape(font.sfnt(), piece, { script: seg.iso === 'Zyyy' ? 'Latn' : seg.iso, direction: seg.rtl ? 'rtl' : 'ltr', features: features.length ? features : undefined })
      glyphs = shaped.map((g) => g.glyph)
      advances = shaped.map((g) => g.advance)
      dx = shaped.map((g) => g.dx)
      dy = shaped.map((g) => g.dy)
      clusters.length = 0
      for (const g of shaped) clusters.push(seg.from + g.cluster)
      zero.length = 0
      own = false
      return true
    }
    advances = []
    dx = []
    if (!(seg.shaped && tryShaper())) {
      if (seg.shaped) why.add('shaping')
      const before = [ids.slice(), clusters.slice()]
      let ok = !font.gsub || substitute(font.gsub, seg.tag, ids, clusters, ligatures)
      advances = ids.map((g) => font.advanceOf(g))
      dx = ids.map(() => 0)
      ok = position(font.gpos, font.kernPairs, seg.tag, ids, advances, dx, kerning) && ok
      // Apple's own tables are not read at all; nor is a table of rules that is damaged or too long
      ok &&= !font.aat && !font.rulesUnread
      if (!ok) {
        // rules of the font that are not applied here: the shaper's result when there is one, else what was laid out without them
        const laid = [ids.slice(), clusters.slice(), advances, dx]
        refill(ids, before[0])
        refill(clusters, before[1])
        if (seg.shaped || !tryShaper()) {
          why.add('features')
          refill(ids, laid[0])
          refill(clusters, laid[1])
          advances = laid[2]
          dx = laid[3]
        }
      }
      if (own && kerning && font.kernUnread && !hasKerning(font.gpos, null, seg.tag)) why.add('kerning')
    }
    for (let i = 0; i < glyphs.length; i++) {
      run.glyphs.push(glyphs[i])
      run.advances.push(advances[i] + spacing)
      run.dx.push(dx[i])
      run.dy.push(dy ? dy[i] : 0)
      run.clusters.push(clusters[i])
    }
    // the hidden characters back in, as glyphs of no width (the space, which draws nothing), where they were
    if (zero.length) {
      const space = font.glyphOf(0x20)
      for (const at of zero) {
        let k = run.clusters.length
        while (k > 0 && run.clusters[k - 1] > at) k--
        run.glyphs.splice(k, 0, space)
        run.advances.splice(k, 0, 0)
        run.dx.splice(k, 0, 0)
        run.dy.splice(k, 0, 0)
        run.clusters.splice(k, 0, at)
      }
    }
  }
  run.width = run.advances.reduce((a, b) => a + b, 0)
  run.inexact = [...why]
  run.exact = !why.size
  return run
}
