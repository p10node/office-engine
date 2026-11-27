import { DAMAGED, FONT_LIMITS, FontError, readSfnt, writeSfnt, type Sfnt } from './container'
import { readGpos, readGsub, readKern, type Positioning, type Substitution } from './layout'

/**
 * A font file read for what laying out text needs: which glyph a character is drawn with, how wide each glyph is,
 * how far lines are apart, and what the font says about pairs of glyphs (kerning) and runs of them (ligatures).
 * Nothing is drawn here and no DOM is used: the same numbers on every browser, in a worker and on a server.
 *
 * All lengths are in the font's own units, `unitsPerEm` to the em: multiply by `size / unitsPerEm`.
 */

const view = (b: Uint8Array) => new DataView(b.buffer, b.byteOffset, b.byteLength)

export interface FontNames {
  /** name 1: the family as the old four-style groups have it ("Arial Narrow") */
  family: string
  /** name 2: "Regular", "Bold Italic"… */
  subfamily: string
  /** name 4 */
  full: string
  /** name 6: the name a PDF calls the font by */
  postScript: string
  /** name 16, when the family has more than four styles ("Avenir Next") */
  typoFamily?: string
}

export interface FontMetrics {
  /** hhea: what most browsers take a line's height from */
  ascent: number
  /** hhea, positive below the baseline */
  descent: number
  lineGap: number
  /** OS/2 typographic values */
  typoAscent: number
  typoDescent: number
  typoLineGap: number
  /** OS/2 Windows values: what Word takes a line's height from */
  winAscent: number
  winDescent: number
  /** fsSelection bit 7: the typographic values are the ones to use */
  useTypo: boolean
  xHeight: number
  capHeight: number
  /** top of the underline's center below the baseline is negative, as in the post table */
  underlinePosition: number
  underlineThickness: number
  strikeoutPosition: number
  strikeoutSize: number
  /** degrees, counter-clockwise from upright: negative for an italic leaning right */
  italicAngle: number
  /** the box of all glyphs: xMin, yMin, xMax, yMax */
  bbox: [number, number, number, number]
}

/** how far lines are apart, above and below the baseline (font units, all positive) */
export interface LineMetrics {
  ascent: number
  descent: number
  /** extra space between two lines, on top of ascent + descent */
  lineGap: number
}

/** the Unicode cmap subtable of a font, searched without unpacking it */
interface Cmap {
  /** format 4: segments */
  segs?: { end: Uint16Array; start: Uint16Array; delta: Int16Array; range: Uint16Array; ids: DataView; idsAt: number }
  /** format 12 (and 13: every character of a group has the group's glyph) */
  groups?: { start: Uint32Array; end: Uint32Array; glyph: Uint32Array; many: boolean }
  /** formats 0 and 6: a plain table from `first` */
  table?: { first: number; glyphs: Uint16Array }
  /** a symbol font: its characters are at U+F000 + the code */
  symbol: boolean
}

function readCmap(b: Uint8Array): Cmap {
  const v = view(b)
  const n = v.getUint16(2)
  let best = -1
  let bestRank = -1
  let symbol = false
  for (let i = 0; i < n; i++) {
    const platform = v.getUint16(4 + i * 8)
    const encoding = v.getUint16(6 + i * 8)
    const offset = v.getUint32(8 + i * 8)
    const format = v.getUint16(offset)
    // all of Unicode before the Basic Multilingual Plane alone, Windows before the Unicode platform, a symbol font last
    const rank =
      platform === 3 && encoding === 10 ? 6 : platform === 0 && (encoding === 4 || encoding === 6) ? 5 : platform === 3 && encoding === 1 ? 4 : platform === 0 ? 3 : platform === 3 && encoding === 0 ? 2 : -1
    if (rank > bestRank && (format === 0 || format === 4 || format === 6 || format === 12 || format === 13)) {
      best = offset
      bestRank = rank
      symbol = platform === 3 && encoding === 0
    }
  }
  if (best < 0) throw new FontError(`${DAMAGED}: no character map that is read (Unicode, formats 0, 4, 6, 12, 13)`)
  const format = v.getUint16(best)
  if (format === 4) {
    const segs = v.getUint16(best + 6) >> 1
    const end = new Uint16Array(segs)
    const start = new Uint16Array(segs)
    const delta = new Int16Array(segs)
    const range = new Uint16Array(segs)
    const at = best + 14
    for (let i = 0; i < segs; i++) {
      end[i] = v.getUint16(at + i * 2)
      start[i] = v.getUint16(at + 2 + segs * 2 + i * 2)
      delta[i] = v.getInt16(at + 2 + segs * 4 + i * 2)
      range[i] = v.getUint16(at + 2 + segs * 6 + i * 2)
    }
    return { segs: { end, start, delta, range, ids: v, idsAt: at + 2 + segs * 6 }, symbol }
  }
  if (format === 12 || format === 13) {
    const count = v.getUint32(best + 12)
    if (count > FONT_LIMITS.records) throw new FontError('The font file is too large: its character map is too long')
    const start = new Uint32Array(count)
    const end = new Uint32Array(count)
    const glyph = new Uint32Array(count)
    for (let i = 0; i < count; i++) {
      start[i] = v.getUint32(best + 16 + i * 12)
      end[i] = v.getUint32(best + 20 + i * 12)
      glyph[i] = v.getUint32(best + 24 + i * 12)
    }
    return { groups: { start, end, glyph, many: format === 13 }, symbol }
  }
  if (format === 6) {
    const first = v.getUint16(best + 6)
    const count = v.getUint16(best + 8)
    const glyphs = new Uint16Array(count)
    for (let i = 0; i < count; i++) glyphs[i] = v.getUint16(best + 10 + i * 2)
    return { table: { first, glyphs }, symbol }
  }
  const glyphs = new Uint16Array(256)
  for (let i = 0; i < 256; i++) glyphs[i] = v.getUint8(best + 6 + i)
  return { table: { first: 0, glyphs }, symbol }
}

function cmapGlyph(c: Cmap, cp: number): number {
  if (c.segs) {
    if (cp > 0xffff) return 0
    const { end, start, delta, range, ids, idsAt } = c.segs
    let lo = 0
    let hi = end.length - 1
    while (lo < hi) {
      const mid = (lo + hi) >> 1
      if (end[mid] < cp) lo = mid + 1
      else hi = mid
    }
    if (lo >= end.length || start[lo] > cp || end[lo] < cp) return 0
    if (!range[lo]) return (cp + delta[lo]) & 0xffff
    const at = idsAt + lo * 2 + range[lo] + (cp - start[lo]) * 2
    // (an offset out of the table is a character without a glyph, as other readers have it)
    if (at + 2 > ids.byteLength) return 0
    const g = ids.getUint16(at)
    return g ? (g + delta[lo]) & 0xffff : 0
  }
  if (c.groups) {
    const { start, end, glyph, many } = c.groups
    let lo = 0
    let hi = start.length - 1
    while (lo < hi) {
      const mid = (lo + hi) >> 1
      if (end[mid] < cp) lo = mid + 1
      else hi = mid
    }
    if (lo >= start.length || start[lo] > cp || end[lo] < cp) return 0
    return many ? glyph[lo] : glyph[lo] + (cp - start[lo])
  }
  const t = c.table!
  return t.glyphs[cp - t.first] ?? 0
}

function readNames(b: Uint8Array | undefined): FontNames {
  const got = new Map<number, { rank: number; text: string }>()
  if (b) {
    const v = view(b)
    const count = v.getUint16(2)
    const strings = v.getUint16(4)
    for (let i = 0; i < count; i++) {
      const rec = 6 + i * 12
      const platform = v.getUint16(rec)
      const encoding = v.getUint16(rec + 2)
      const language = v.getUint16(rec + 4)
      const id = v.getUint16(rec + 6)
      const length = v.getUint16(rec + 8)
      const offset = strings + v.getUint16(rec + 10)
      if (id > 17 || offset + length > b.length) continue
      // Windows English first, then any Windows or Unicode name, then the Macintosh one
      const wide = platform === 3 || platform === 0
      const rank = platform === 3 && language === 0x409 ? 4 : platform === 3 && (encoding === 1 || encoding === 10) ? 3 : platform === 0 ? 2 : platform === 1 && encoding === 0 ? 1 : 0
      if (!rank || rank <= (got.get(id)?.rank ?? 0)) continue
      let text = ''
      if (wide) for (let k = 0; k + 1 < length; k += 2) text += String.fromCharCode(v.getUint16(offset + k))
      else for (let k = 0; k < length; k++) text += String.fromCharCode(b[offset + k])
      got.set(id, { rank, text })
    }
  }
  const name = (id: number) => got.get(id)?.text ?? ''
  const family = name(1) || name(16) || name(4)
  return {
    family,
    subfamily: name(2) || name(17) || 'Regular',
    full: name(4) || family,
    // a PDF name has no spaces and none of ( ) < > [ ] { } / %
    postScript: (name(6) || name(4) || family || 'Font').replace(/[\s()<>[\]{}/%]/g, ''),
    typoFamily: name(16) || undefined,
  }
}

export class Font {
  readonly unitsPerEm: number
  readonly numGlyphs: number
  readonly names: FontNames
  readonly metrics: FontMetrics
  /** OS/2 usWeightClass: 400 regular, 700 bold */
  readonly weight: number
  readonly italic: boolean
  readonly fixedPitch: boolean
  /** `glyf` outlines (TrueType) or `cff` ones (PostScript) */
  readonly outlines: 'glyf' | 'cff'
  /** the container the file came in */
  readonly format: Sfnt['format']
  /** how many fonts the file holds */
  readonly fonts: number
  /**
   * A variable font (an `fvar` table): the numbers read are those of its default instance. Text set at another
   * weight or width of it is drawn with other widths, which are not read.
   */
  readonly variable: boolean
  /** the license lets the font be embedded in a document (OS/2 fsType: not "restricted license") */
  readonly embeddable: boolean
  /** substitutions the font makes by itself: ligatures, and what else it does to plain text */
  readonly gsub: Substitution | null
  /** kerning: the GPOS table's, else the kern table's */
  readonly gpos: Positioning | null
  readonly kernPairs: Map<number, number> | null
  /** true when the font has a kern table of a kind that is not read (so its kerning is not applied) */
  readonly kernUnread: boolean
  /**
   * The font's rules for runs of glyphs are (also) in Apple's own tables (`morx`, `kerx`, `trak`: the fonts macOS
   * comes with), which are not read: text in it is laid out by its glyph widths, and never said to be exact.
   */
  readonly aat: boolean
  /**
   * The font's GSUB, GPOS or kern table could not be read (damaged, or longer than `FONT_LIMITS.records`): the font
   * is used without it, as a browser drops a table it cannot check, and its text is never said to be exact.
   */
  readonly rulesUnread: boolean
  private cmap: Cmap
  private advances: Uint16Array
  private bytes?: Uint8Array<ArrayBuffer>

  private constructor(readonly tables: Map<string, Uint8Array>, sfnt: Sfnt) {
    const need = (tag: string) => {
      const t = tables.get(tag)
      if (!t) throw new FontError(`${DAMAGED}: no ${tag.trim()} table`)
      return t
    }
    const head = view(need('head'))
    const hhea = view(need('hhea'))
    const maxp = view(need('maxp'))
    this.outlines = sfnt.outlines
    this.format = sfnt.format
    this.fonts = sfnt.fonts
    this.unitsPerEm = head.getUint16(18)
    if (this.unitsPerEm < 16 || this.unitsPerEm > 16384) throw new FontError(`${DAMAGED}: ${this.unitsPerEm} units to the em`)
    this.numGlyphs = maxp.getUint16(4)
    const os2 = tables.get('OS/2') && view(tables.get('OS/2')!)
    const post = tables.get('post') && view(tables.get('post')!)
    const has = (v: DataView | undefined, n: number): v is DataView => !!v && v.byteLength >= n
    const ascent = hhea.getInt16(4)
    const descent = -hhea.getInt16(6)
    const fsSelection = has(os2, 64) ? os2.getUint16(62) : 0
    this.metrics = {
      ascent,
      descent,
      lineGap: hhea.getInt16(8),
      typoAscent: has(os2, 74) ? os2.getInt16(68) : ascent,
      typoDescent: has(os2, 74) ? -os2.getInt16(70) : descent,
      typoLineGap: has(os2, 74) ? os2.getInt16(72) : hhea.getInt16(8),
      winAscent: has(os2, 78) ? os2.getUint16(74) : ascent,
      winDescent: has(os2, 78) ? os2.getUint16(76) : descent,
      useTypo: !!(fsSelection & 0x80),
      xHeight: has(os2, 90) && os2.getUint16(0) >= 2 ? os2.getInt16(86) : 0,
      capHeight: has(os2, 90) && os2.getUint16(0) >= 2 ? os2.getInt16(88) : 0,
      underlinePosition: has(post, 12) ? post.getInt16(8) : -Math.round(this.unitsPerEm / 10),
      underlineThickness: has(post, 12) ? post.getInt16(10) : Math.round(this.unitsPerEm / 20),
      strikeoutSize: has(os2, 30) ? os2.getInt16(26) : Math.round(this.unitsPerEm / 20),
      strikeoutPosition: has(os2, 30) ? os2.getInt16(28) : Math.round(this.unitsPerEm * 0.3),
      italicAngle: has(post, 8) ? post.getInt32(4) / 65536 : 0,
      bbox: [head.getInt16(36), head.getInt16(38), head.getInt16(40), head.getInt16(42)],
    }
    this.weight = has(os2, 6) ? os2.getUint16(4) : head.getUint16(44) & 1 ? 700 : 400
    this.italic = !!(fsSelection & 1) || !!(head.getUint16(44) & 2)
    this.fixedPitch = has(post, 16) && post.getUint32(12) !== 0
    this.variable = tables.has('fvar')
    // fsType bit 1: restricted license embedding (bit 9, bitmaps only, does not matter to outlines)
    this.embeddable = !(has(os2, 10) && os2.getUint16(8) & 0x0002)
    this.names = readNames(tables.get('name'))
    this.cmap = readCmap(need('cmap'))

    const hmtx = view(need('hmtx'))
    const numH = Math.min(hhea.getUint16(34), this.numGlyphs)
    if (!numH && this.numGlyphs) throw new FontError(`${DAMAGED}: no glyph widths`)
    this.advances = new Uint16Array(this.numGlyphs)
    for (let g = 0; g < numH; g++) this.advances[g] = hmtx.getUint16(g * 4)
    // the glyphs after the last width given are as wide as that one (a fixed-pitch font gives one width)
    if (numH) this.advances.fill(this.advances[numH - 1], numH)

    const budget = { left: FONT_LIMITS.records }
    let unread = false
    /** a table of rules, or null (and the font marked) when it cannot be read: the widths and the character map still can */
    const rules = <T>(tag: string, read: (b: Uint8Array) => T): T | null => {
      const b = tables.get(tag)
      if (!b) return null
      try {
        return read(b)
      } catch {
        unread = true
        return null
      }
    }
    this.gsub = rules('GSUB', (b) => readGsub(b, budget))
    this.gpos = rules('GPOS', (b) => readGpos(b, budget))
    const kern = rules('kern', (b) => readKern(b, budget))
    this.kernPairs = kern?.pairs ?? null
    this.kernUnread = !!kern?.unread
    this.rulesUnread = unread
    // (a shaping engine takes these before GSUB and GPOS for text set in lines)
    this.aat = tables.has('trak') || tables.has('kerx') || tables.has('morx') || tables.has('mort')
  }

  /**
   * Reads a font file: TTF, OTF, WOFF, WOFF2, or one font (`index`) of a TTC collection.
   * Throws a `FontError` that says what is wrong with a file that cannot be read.
   */
  static async open(data: ArrayBuffer | ArrayBufferView | Blob, index = 0): Promise<Font> {
    const bytes =
      data instanceof Uint8Array ? data : ArrayBuffer.isView(data) ? new Uint8Array(data.buffer, data.byteOffset, data.byteLength) : new Uint8Array(data instanceof Blob ? await data.arrayBuffer() : data)
    const sfnt = await readSfnt(bytes, index)
    try {
      return new Font(sfnt.tables, sfnt)
    } catch (e) {
      // (a read past the end of a table: a DataView's own RangeError)
      throw e instanceof FontError ? e : new FontError(`${DAMAGED}: a table ends early`)
    }
  }

  /** the glyph `codePoint` is drawn with; 0 (the "missing glyph") when the font has none for it */
  glyphOf(codePoint: number): number {
    let g = cmapGlyph(this.cmap, codePoint)
    // a symbol font (Wingdings, Symbol) has its characters in the private use area: text asks for them by their Latin code
    if (!g && this.cmap.symbol && codePoint < 0x100) g = cmapGlyph(this.cmap, 0xf000 + codePoint)
    return g < this.numGlyphs ? g : 0
  }

  /** how far the pen moves after `glyph` */
  advanceOf(glyph: number): number {
    return this.advances[glyph] ?? 0
  }

  /**
   * How far lines of this font are apart.
   * - `'browser'` (default): the hhea values, or the OS/2 typographic ones when the font asks for them (fsSelection
   *   bit 7): what browsers take `line-height: normal` from.
   * - `'word'`: the OS/2 Windows values, with the rest of the hhea line gap between lines: Word's "single" spacing.
   */
  lineMetrics(as: 'browser' | 'word' = 'browser'): LineMetrics {
    const m = this.metrics
    if (as === 'word') {
      const height = m.winAscent + m.winDescent
      return { ascent: m.winAscent, descent: m.winDescent, lineGap: Math.max(0, m.lineGap - (height - (m.ascent + m.descent))) }
    }
    if (m.useTypo) return { ascent: m.typoAscent, descent: m.typoDescent, lineGap: Math.max(0, m.typoLineGap) }
    // (a font whose hhea says nothing is read by its OS/2 values, as the browsers' font libraries do)
    if (!m.ascent && !m.descent) return m.typoAscent || m.typoDescent ? { ascent: m.typoAscent, descent: m.typoDescent, lineGap: Math.max(0, m.typoLineGap) } : { ascent: m.winAscent, descent: m.winDescent, lineGap: 0 }
    return { ascent: m.ascent, descent: m.descent, lineGap: Math.max(0, m.lineGap) }
  }

  /**
   * The font as a bare TTF / OTF file, whatever it was packed in (WOFF, WOFF2, a collection): what a shaping engine
   * or a PDF takes.
   */
  sfnt(): Uint8Array<ArrayBuffer> {
    return (this.bytes ??= writeSfnt(this.tables, this.outlines))
  }
}
