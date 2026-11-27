import { afterEach, describe, expect, test } from 'bun:test'
import { readFileSync } from 'node:fs'
import { cssFont, type FontSource } from '../src/core/fonts'
import { Font, FontBook, FontError, FONT_LIMITS, fontFormat, layoutText, needsShaping, subsetFont, toUnicodeCMap, type Shaper } from '../src/core/font'
import { unicodeRanges } from '../src/core/font/book'
import { segments } from '../src/core/font/text'
import { DocumentClient } from '../src/core/worker/client'
import { Engine } from '../src/core/worker/engine'
import type { Events, MeasuredText } from '../src/core/worker/protocol'
import { buildFont, chainContext, classPairPos, collection, kernTable, layoutTable, ligatureSubst, pairPos, singleSubst, woff, type FontSpec } from './font-tables'

// The test font: glyph 0 the missing glyph, 1 the space, 2–27 A–Z (500, 510… wide), 28–53 a–z (400, 410…), 54 the fi
// ligature, 55 ffi, 56 another A, 57 a composite of A and V, 58 an emoji, 59 a combining acute, 60 a Hebrew alef.
const CHARS: [number, number][] = [[0x20, 1]]
for (let i = 0; i < 26; i++) CHARS.push([0x41 + i, 2 + i], [0x61 + i, 28 + i])
CHARS.push([0xfb01, 54], [0x1f600, 58])
const ADVANCES = [600, 250, ...Array.from({ length: 26 }, (_, i) => 500 + i * 10), ...Array.from({ length: 26 }, (_, i) => 400 + i * 10), 700, 900, 520, 1210, 800]
const g = (c: string) => CHARS.find((x) => x[0] === c.codePointAt(0))![1]
/** exactly the font `test/fixtures/font/boxes.woff2` was made from (see the test of that file) */
const BOXES: FontSpec = { advances: ADVANCES, chars: CHARS, composites: { 57: [g('A'), g('V')] }, cmap12: true }
/** …with a mark and a Hebrew letter, for text that needs shaping */
const MORE: FontSpec = { ...BOXES, advances: [...ADVANCES, 0, 640], chars: [...CHARS, [0x301, 59], [0x5d0, 60]] }

const GSUB = layoutTable(false, [['liga', [0]], ['calt', [1]]], [
  { type: 4, subtables: [ligatureSubst([[[g('f'), g('f'), g('i')], 55], [[g('f'), g('i')], 54]])] },
  // A between x and y becomes the other A (lookup 2): a rule of a kind that is matched, not applied
  { type: 6, subtables: [chainContext([[g('x')]], [[g('A')]], [[g('y')]], [[0, 2]])] },
  { type: 1, subtables: [singleSubst([[g('A'), 56]])] },
])
const GPOS = layoutTable(true, [['kern', [0, 1]]], [
  { type: 2, subtables: [pairPos([[g('A'), g('V'), -80], [g('T'), g('o'), -60]])] },
  { type: 2, extension: true, subtables: [classPairPos([[g('W'), g('Y')]], [[g('a'), g('e')]], [[-40]])] },
])

const open = (spec: FontSpec) => Font.open(buildFont(spec))
const widths = (text: string) => [...text].map((c) => ADVANCES[g(c)])
const sum = (a: number[]) => a.reduce((x, y) => x + y, 0)

describe('a font file: what its tables say', () => {
  test('glyph count, names, weight and the kind of file', async () => {
    const font = await open({ ...BOXES, family: 'Test Read', subfamily: 'Bold Italic', weight: 700, italic: true })
    expect(font.unitsPerEm).toBe(1000)
    expect(font.numGlyphs).toBe(59)
    expect(font.names).toMatchObject({ family: 'Test Read', subfamily: 'Bold Italic', full: 'Test Read Bold Italic', postScript: 'TestRead-BoldItalic' })
    expect(font.weight).toBe(700)
    expect(font.italic).toBe(true)
    expect(font.metrics.italicAngle).toBe(-12)
    expect([font.outlines, font.format, font.fonts]).toEqual(['glyf', 'sfnt', 1])
    expect([font.variable, font.aat, font.rulesUnread, font.embeddable, font.fixedPitch]).toEqual([false, false, false, true, false])
    const plain = await open(BOXES)
    expect([plain.weight, plain.italic, plain.names.subfamily]).toEqual([400, false, 'Regular'])
  })

  test('characters to glyphs: the Basic Multilingual Plane, beyond it, and none', async () => {
    const font = await open(BOXES)
    expect(font.glyphOf(0x41)).toBe(2)
    expect(font.glyphOf(0x7a)).toBe(53)
    expect(font.glyphOf(0x20)).toBe(1)
    expect(font.glyphOf(0xfb01)).toBe(54)
    expect(font.glyphOf(0x1f600)).toBe(58)
    for (const cp of [0, 0x40, 0x5b, 0x7b, 0x20ac, 0xffff, 0x10000, 0x1f601, 0x10ffff]) expect(font.glyphOf(cp)).toBe(0)
    // a map of the Basic Multilingual Plane alone (format 4)
    const bmp = await open({ ...BOXES, cmap12: false })
    expect(bmp.glyphOf(0x61)).toBe(28)
    expect(bmp.glyphOf(0x1f600)).toBe(0)
  })

  test('glyph widths, the last one given for the glyphs after it', async () => {
    const font = await open(BOXES)
    ADVANCES.forEach((w, glyph) => expect(font.advanceOf(glyph)).toBe(w))
    expect(font.advanceOf(59)).toBe(0)
    const short = await open({ ...BOXES, longMetrics: 30 })
    expect(short.advanceOf(29)).toBe(ADVANCES[29])
    for (const glyph of [30, 42, 58]) expect(short.advanceOf(glyph)).toBe(ADVANCES[29])
  })

  test('how far lines are apart: as browsers take it, and as Word does', async () => {
    // Arial's numbers: Word's single spacing is 2355 / 2048 = 1.15
    const arial = await open({ ...BOXES, upem: 2048, hhea: [1854, -434, 67], typo: [1491, -431, 307], win: [1854, 434] })
    expect(arial.lineMetrics()).toEqual({ ascent: 1854, descent: 434, lineGap: 67 })
    expect(arial.lineMetrics('word')).toEqual({ ascent: 1854, descent: 434, lineGap: 67 })
    // Calibri's: the Windows values are taller than the hhea ones, and take the line gap in
    const calibri = await open({ ...BOXES, upem: 2048, hhea: [1536, -512, 452], typo: [1536, -512, 452], win: [1950, 550] })
    expect(calibri.lineMetrics()).toEqual({ ascent: 1536, descent: 512, lineGap: 452 })
    expect(calibri.lineMetrics('word')).toEqual({ ascent: 1950, descent: 550, lineGap: 0 })
    // a font that asks for its typographic values (fsSelection bit 7)
    const typo = await open({ ...BOXES, hhea: [900, -300, 0], typo: [750, -250, 200], useTypo: true })
    expect(typo.lineMetrics()).toEqual({ ascent: 750, descent: 250, lineGap: 200 })
    // a hhea table that says nothing
    const empty = await open({ ...BOXES, hhea: [0, 0, 0], typo: [700, -200, 100] })
    expect(empty.lineMetrics()).toEqual({ ascent: 700, descent: 200, lineGap: 100 })
    expect(arial.metrics).toMatchObject({ underlinePosition: -100, underlineThickness: 50, capHeight: 700, xHeight: 500, winAscent: 1854, typoLineGap: 307 })
  })

  test('what a font says about itself: variable, Apple’s own layout tables, not to be embedded', async () => {
    expect((await open({ ...BOXES, extra: { fvar: [0, 1, 0, 0] } })).variable).toBe(true)
    for (const tag of ['trak', 'morx', 'kerx']) expect((await open({ ...BOXES, extra: { [tag]: [0, 0, 0, 0] } })).aat).toBe(true)
    expect((await open({ ...BOXES, fsType: 2 })).embeddable).toBe(false)
    expect((await open({ ...BOXES, fsType: 8 })).embeddable).toBe(true)
  })
})

describe('a run of text in a font', () => {
  test('glyphs, widths and the characters they stand for', async () => {
    const run = layoutText(await open(BOXES), 'AV To')
    expect(run.glyphs).toEqual([2, 23, 1, 21, 42])
    expect(run.advances).toEqual(widths('AV To'))
    expect(run.clusters).toEqual([0, 1, 2, 3, 4])
    expect(run.width).toBe(sum(widths('AV To')))
    expect(run.dx.concat(run.dy).every((x) => x === 0)).toBe(true)
    expect([run.exact, run.inexact, run.missing]).toEqual([true, [], 0])
    expect(layoutText(await open(BOXES), '').glyphs).toEqual([])
  })

  test('pair kerning from GPOS: pairs of glyphs, pairs of classes, a lookup behind an extension', async () => {
    const font = await open({ ...BOXES, gpos: GPOS })
    const run = layoutText(font, 'AV To WaYe')
    const plain = widths('AV To WaYe')
    expect(run.advances).toEqual([plain[0] - 80, plain[1], plain[2], plain[3] - 60, plain[4], plain[5], plain[6] - 40, plain[7], plain[8] - 40, plain[9]])
    expect(run.exact).toBe(true)
    // pairs the font has nothing for
    expect(layoutText(font, 'VA oT aW').advances).toEqual(widths('VA oT aW'))
    expect(layoutText(font, 'AV To WaYe', { kerning: false }).advances).toEqual(plain)
  })

  test('the kern table is the kerning of a font whose GPOS has none', async () => {
    const kern = kernTable([[g('A'), g('V'), -70], [g('T'), g('o'), -50]])
    const font = await open({ ...BOXES, kern })
    expect(layoutText(font, 'AVA To').advances).toEqual([500 - 70, 710, 500, 250, 690 - 50, 540])
    expect(layoutText(font, 'AV', { kerning: false }).width).toBe(1210)
    // with kerning in GPOS, the kern table is not read for it
    const both = await open({ ...BOXES, gpos: GPOS, kern: kernTable([[g('A'), g('V'), -999], [g('V'), g('A'), -999]]) })
    expect(layoutText(both, 'AVA').advances).toEqual([500 - 80, 710, 500])
    // …but it is when GPOS has other rules only
    const other = await open({ ...BOXES, gpos: layoutTable(true, [['cpsp', [0]]], [{ type: 2, subtables: [pairPos([[g('A'), g('V'), -5]])] }]), kern })
    expect(layoutText(other, 'AV').advances).toEqual([430, 710])
  })

  test('ligatures: the longest first, off when asked or when letters are spaced', async () => {
    const font = await open({ ...BOXES, gsub: GSUB })
    const run = layoutText(font, 'office fit')
    expect(run.glyphs).toEqual([g('o'), 55, g('c'), g('e'), 1, 54, g('t')])
    expect(run.clusters).toEqual([0, 1, 4, 5, 6, 7, 9])
    expect(run.width).toBe(540 + 900 + 420 + 440 + 250 + 700 + 590)
    expect(run.exact).toBe(true)
    expect(layoutText(font, 'office', { ligatures: false }).glyphs).toEqual([...'office'].map(g))
    const spaced = layoutText(font, 'fi', { letterSpacing: 30 })
    expect(spaced.glyphs).toEqual([g('f'), g('i')])
    expect(spaced.advances).toEqual([450 + 30, 480 + 30])
  })

  test('a rule that is not applied here: said, when it would change the run', async () => {
    const font = await open({ ...BOXES, gsub: GSUB })
    // A between x and y: a shaping engine draws the other A
    const hit = layoutText(font, 'xAy')
    expect([hit.exact, hit.inexact]).toEqual([false, ['features']])
    expect(hit.glyphs).toEqual([g('x'), g('A'), g('y')])
    // anywhere else, the rule does nothing
    for (const text of ['xAz', 'Ay', 'xA', 'A', 'yAx']) expect(layoutText(font, text).exact).toBe(true)
    // a rule that only stops the ones after it (no lookup of its own) is not a change
    const stop = layoutTable(false, [['calt', [0]]], [
      { type: 6, subtables: [chainContext([[g('x')]], [[g('f')]], []), chainContext([], [[g('f')]], [], [[0, 1]])] },
      { type: 1, subtables: [singleSubst([[g('f'), 56]])] },
    ])
    const stopped = await open({ ...BOXES, gsub: stop })
    expect(layoutText(stopped, 'xf').exact).toBe(true)
    expect(layoutText(stopped, 'af').exact).toBe(false)
    // a font with Apple's own tables is never exact: they are not read
    expect(layoutText(await open({ ...BOXES, extra: { morx: [0, 0, 0, 0] } }), 'AV').inexact).toEqual(['features'])
  })

  test('a long text whose rules are not applied: laid out without them, however many glyphs it has', async () => {
    const font = await open({ ...BOXES, gsub: GSUB })
    // (more glyphs than a function call takes arguments: some 120,000 in V8, under a million in JavaScriptCore)
    const text = 'xAy' + 'ab'.repeat(600_000)
    const run = layoutText(font, text)
    expect([run.exact, run.inexact, run.glyphs.length]).toEqual([false, ['features'], text.length])
    expect(run.glyphs.slice(0, 5)).toEqual([g('x'), g('A'), g('y'), g('a'), g('b')])
    expect(run.clusters[run.clusters.length - 1]).toBe(text.length - 1)
    expect(run.width).toBe(sum(widths('xAy')) + 600_000 * sum(widths('ab')))
  })

  test('a single substitution of a feature browsers apply by themselves', async () => {
    const font = await open({ ...BOXES, gsub: layoutTable(false, [['locl', [0]]], [{ type: 1, subtables: [singleSubst([[g('A'), 56]])] }]) })
    const run = layoutText(font, 'BAB')
    expect(run.glyphs).toEqual([g('B'), 56, g('B')])
    expect(run.advances).toEqual([510, 520, 510])
    expect(run.exact).toBe(true)
  })

  test('a character without a glyph, and the ones that take no room', async () => {
    const font = await open({ ...BOXES, gpos: GPOS, gsub: GSUB })
    const missing = layoutText(font, 'A€')
    expect([missing.missing, missing.exact, missing.inexact]).toEqual([1, false, ['missing']])
    expect(missing.glyphs).toEqual([2, 0])
    expect(missing.width).toBe(500 + 600)
    // a soft hyphen and a zero-width space: no width, and the letters around them kern and join as if they were not there
    const hidden = layoutText(font, 'A­V f​i')
    expect(hidden.glyphs).toEqual([g('A'), 1, g('V'), 1, 54, 1])
    expect(hidden.advances).toEqual([420, 0, 710, 250, 700, 0])
    expect(hidden.clusters).toEqual([0, 1, 2, 3, 4, 5])
    expect(hidden.exact).toBe(true)
  })

  test('text is cut where its script changes, and what has no script goes with its neighbours', () => {
    const cut = (text: string) => segments(text).map((s) => `${text.slice(s.from, s.to)}|${s.tag}`)
    expect(cut('Hello, world 123!')).toEqual(['Hello, world 123!|latn'])
    expect(cut('"Привет" мир')).toEqual(['"Привет" мир|cyrl'])
    expect(cut('abc абв 1')).toEqual(['abc |latn', 'абв 1|cyrl'])
    expect(cut('漢字かなカナ한국')).toEqual(['漢字|hani', 'かな|kana', 'カナ|kana', '한국|hang'])
    expect(cut('123 + 456')).toEqual(['123 + 456|DFLT'])
    expect(cut('')).toEqual([])
    expect(segments('abc').map((s) => s.iso)).toEqual(['Latn'])
    // kerning does not reach across a change of script
    expect(cut('AБ')).toEqual(['A|latn', 'Б|cyrl'])
  })

  test('which text needs a shaping engine', () => {
    for (const plain of ['Hello', 'Tiếng Việt', 'Größe', 'Привет', 'Ελληνικά', '日本語のテキスト', '한국어', '1 + 2 = 3 €', '😀']) expect(needsShaping(plain)).toBe(false)
    // the same Vietnamese text with its marks as characters of their own
    expect(needsShaping('Tiếng Việt'.normalize('NFD'))).toBe(true)
    for (const shaped of ['مرحبا', 'שלום', 'नमस्ते', 'สวัสดี', 'a‍b', '👍🏽', '🇻🇳', 'é']) expect(needsShaping(shaped)).toBe(true)
  })

  test('text that needs shaping: not exact without a shaper, the shaper’s glyphs with one', async () => {
    const font = await open({ ...MORE, gpos: GPOS })
    const marks = layoutText(font, 'Vé')
    expect([marks.exact, marks.inexact]).toEqual([false, ['shaping']])
    expect(marks.glyphs).toEqual([g('V'), g('e'), 59])
    const calls: { text: string; script: string; direction: string; features?: string[]; sfnt: boolean }[] = []
    const shaper: Shaper = {
      shape(bytes, text, opts) {
        calls.push({ text, ...opts, sfnt: fontFormat(bytes) === 'sfnt' })
        // a made-up result: one glyph for each character, the last one raised
        return [...text].map((_, i) => ({ glyph: 7 + i, advance: 100 * (i + 1), dx: 0, dy: i === text.length - 1 ? 50 : 0, cluster: i }))
      },
    }
    const shaped = layoutText(font, 'AV אא Vé', { shaper, kerning: false })
    // the Latin letters at the start are laid out here, the Hebrew and the marked ones by the shaper
    expect(calls).toEqual([
      { text: 'אא ', script: 'Hebr', direction: 'rtl', features: ['-kern'], sfnt: true },
      { text: 'Vé', script: 'Latn', direction: 'ltr', features: ['-kern'], sfnt: true },
    ])
    expect(shaped.glyphs).toEqual([g('A'), g('V'), 1, 7, 8, 9, 7, 8, 9])
    expect(shaped.advances).toEqual([500, 710, 250, 100, 200, 300, 100, 200, 300])
    expect(shaped.clusters).toEqual([0, 1, 2, 3, 4, 5, 6, 7, 8])
    expect(shaped.dy).toEqual([0, 0, 0, 0, 0, 50, 0, 0, 50])
    expect([shaped.exact, shaped.inexact]).toEqual([true, []])
    // a rule of the font that is not applied here goes to the shaper too
    calls.length = 0
    const ruled = layoutText(await open({ ...BOXES, gsub: GSUB }), 'xAy', { shaper })
    expect(calls.map((c) => c.text)).toEqual(['xAy'])
    expect([ruled.exact, ruled.glyphs]).toEqual([true, [7, 8, 9]])
  })
})

describe('the files a font comes in', () => {
  test('a collection: each font by its index', async () => {
    const ttc = collection([buildFont(BOXES), buildFont({ ...BOXES, family: 'Second', weight: 700 })])
    expect(fontFormat(ttc)).toBe('ttc')
    const first = await Font.open(ttc)
    const second = await Font.open(ttc, 1)
    expect([first.fonts, first.names.family, first.weight]).toEqual([2, 'Test Read', 400])
    expect([second.names.family, second.weight]).toEqual(['Second', 700])
    expect(layoutText(second, 'AV').width).toBe(1210)
    await expect(Font.open(ttc, 2)).rejects.toThrow('The font file holds 2 fonts: there is no font 2')
    await expect(Font.open(buildFont(BOXES), 1)).rejects.toThrow('there is no font 1')
    // on its own again: a file a shaper or a PDF takes
    expect(fontFormat(second.sfnt())).toBe('sfnt')
    expect((await Font.open(second.sfnt())).names.family).toBe('Second')
  })

  test('WOFF: the same font as the plain file', async () => {
    const spec = { ...BOXES, gsub: GSUB, gpos: GPOS }
    const packed = await woff(buildFont(spec))
    expect(fontFormat(packed)).toBe('woff')
    const font = await Font.open(packed)
    const plain = await open(spec)
    expect(font.format).toBe('woff')
    expect(layoutText(font, 'AV office WaYe')).toEqual(layoutText(plain, 'AV office WaYe'))
    for (const [tag, data] of plain.tables) expect(Buffer.compare(Buffer.from(font.tables.get(tag)!), Buffer.from(data))).toBe(0)
  })

  // test/fixtures/font/boxes.woff2 is the font of BOXES packed by fontTools, with the tables WOFF2 rewrites rewritten:
  //   t = TTFont('boxes.ttf'); t.flavor = 'woff2'; t.flavorData = WOFF2FlavorData(transformedTables={'glyf', 'loca', 'hmtx'}); t.save('boxes.woff2')
  test('WOFF2: its outlines and widths are rebuilt as they were', async () => {
    const packed = new Uint8Array(readFileSync(new URL('./fixtures/font/boxes.woff2', import.meta.url)))
    expect(fontFormat(packed)).toBe('woff2')
    const font = await Font.open(packed)
    const plain = await open(BOXES)
    expect([font.format, font.outlines, font.numGlyphs]).toEqual(['woff2', 'glyf', 59])
    expect(font.names).toEqual(plain.names)
    expect(font.metrics).toEqual(plain.metrics)
    expect(Buffer.compare(Buffer.from(font.tables.get('hmtx')!), Buffer.from(plain.tables.get('hmtx')!))).toBe(0)
    expect(layoutText(font, 'AV office \u{1f600}')).toEqual(layoutText(plain, 'AV office \u{1f600}'))
    // every glyph: the same box, the same points, the same parts
    const outline = (f: Font, glyph: number) => {
      const loca = new DataView(f.tables.get('loca')!.buffer, f.tables.get('loca')!.byteOffset)
      const long = new DataView(f.tables.get('head')!.buffer, f.tables.get('head')!.byteOffset).getUint16(50) === 1
      const at = (i: number) => (long ? loca.getUint32(i * 4) : loca.getUint16(i * 2) * 2)
      const data = f.tables.get('glyf')!.subarray(at(glyph), at(glyph + 1))
      if (!data.length) return null
      const v = new DataView(data.buffer, data.byteOffset, data.byteLength)
      const head = [v.getInt16(0), v.getInt16(2), v.getInt16(4), v.getInt16(6), v.getInt16(8)]
      if (head[0] < 0) {
        // a composite: each part's glyph and where it is put (as two words or two bytes: either is the same glyph)
        const parts: number[] = []
        for (let p = 10; ; ) {
          const fl = v.getUint16(p)
          parts.push(v.getUint16(p + 2), fl & 1 ? v.getInt16(p + 4) : v.getInt8(p + 4), fl & 1 ? v.getInt16(p + 6) : v.getInt8(p + 5))
          p += fl & 1 ? 8 : 6
          if (!(fl & 0x20)) return { head, parts }
        }
      }
      // a simple glyph of the test font: one contour of four points on the curve
      const points = v.getUint16(10) + 1
      const flags = [...data.subarray(14, 14 + points)]
      const xy: number[] = []
      let p = 14 + points
      for (const [short, same] of [[0x02, 0x10], [0x04, 0x20]]) {
        let c = 0
        for (const fl of flags) {
          if (fl & short) c += fl & same ? data[p++] : -data[p++]
          else if (!(fl & same)) {
            c += v.getInt16(p)
            p += 2
          }
          xy.push(c)
        }
      }
      return { head, on: flags.map((fl) => fl & 1), xy }
    }
    for (let glyph = 0; glyph < 59; glyph++) expect(outline(font, glyph)).toEqual(outline(plain, glyph))
    expect(outline(font, 2)).toEqual({ head: [1, 50, 0, 450, 700], on: [1, 1, 1, 1], xy: [50, 50, 450, 450, 0, 700, 700, 0] })
    expect(outline(font, 57)).toEqual({ head: [-1, 50, 0, 1160, 700], parts: [2, 0, 0, 23, 500, 0] })
    // and it subsets like the plain file
    expect(subsetFont(font, [57]).glyphs).toEqual(subsetFont(plain, [57]).glyphs)
  })

  test('what is not a font, or not a whole one, is refused in words', async () => {
    await expect(Font.open(new TextEncoder().encode('hello, this is no font at all'))).rejects.toThrow('Not a font file (TTF, OTF, TTC, WOFF or WOFF2)')
    await expect(Font.open(new Uint8Array(3))).rejects.toThrow(FontError)
    const bytes = buildFont({ ...BOXES, gsub: GSUB, gpos: GPOS, kern: kernTable([[2, 23, -70]]) })
    // cut anywhere, with its tables overwritten anywhere: a FontError or a font, never anything else, never forever
    const started = performance.now()
    let refused = 0
    for (let cut = 0; cut < bytes.length; cut += 37) {
      const r = await Font.open(bytes.slice(0, cut)).then(
        () => null,
        (e) => e,
      )
      expect(r).toBeInstanceOf(FontError)
      refused++
    }
    let seed = 12345
    const next = () => (seed = (seed * 1103515245 + 12345) & 0x7fffffff)
    for (let round = 0; round < 300; round++) {
      const broken = bytes.slice()
      for (let k = 0; k < 6; k++) broken[200 + (next() % (broken.length - 200))] = next() & 0xff
      const r = await Font.open(broken).then(
        (f) => {
          layoutText(f, 'AV office WaYe xAy')
          try {
            subsetFont(f, [2, 23, 57])
          } catch (e) {
            return e
          }
          return null
        },
        (e) => e,
      )
      if (r !== null) expect(r).toBeInstanceOf(FontError)
    }
    expect(refused).toBeGreaterThan(50)
    expect(performance.now() - started).toBeLessThan(20_000)
  })

  describe('limits', () => {
    const limits = { ...FONT_LIMITS }
    afterEach(() => void Object.assign(FONT_LIMITS, limits))

    test('a file larger than a font may be is not read', async () => {
      FONT_LIMITS.file = 1000
      await expect(open(BOXES)).rejects.toThrow('The font file is too large')
    })

    test('a WOFF that unpacks to more than it may is stopped', async () => {
      const packed = await woff(buildFont(BOXES))
      FONT_LIMITS.unpacked = 500
      await expect(Font.open(packed)).rejects.toThrow('The font file is too large once unpacked')
    })

    test('tables longer than a font may make the reader hold: the font without its rules, never exact; no font without its character map', async () => {
      FONT_LIMITS.records = 20
      const font = await open({ ...BOXES, cmap12: false, gsub: GSUB, gpos: GPOS })
      expect([font.rulesUnread, font.gsub, font.gpos]).toEqual([true, null, null])
      const run = layoutText(font, 'AV office')
      expect(run.advances).toEqual(widths('AV office'))
      expect([run.exact, run.inexact]).toEqual([false, ['features']])
      expect((await open({ ...BOXES, cmap12: false })).rulesUnread).toBe(false)
      await expect(open(BOXES)).rejects.toThrow('The font file is too large: its character map is too long')
    })
  })
})

describe('a font cut down to the glyphs used', () => {
  const dv = (b: Uint8Array) => new DataView(b.buffer, b.byteOffset, b.byteLength)

  test('TrueType: the glyphs asked for and their parts, renumbered, with their widths', async () => {
    const font = await open({ ...BOXES, gsub: GSUB, gpos: GPOS })
    const sub = subsetFont(font, [g('o'), 57, 55], [[0x6f, g('o')], [0x1f600, 58], [0x41, g('A')]])
    // the missing glyph, A and V (the parts of 57), o, ffi and the composite
    expect([...sub.glyphs]).toEqual([[0, 0], [2, 1], [23, 2], [42, 3], [55, 4], [57, 5]])
    expect(sub.outlines).toBe('glyf')
    expect(fontFormat(sub.bytes)).toBe('sfnt')
    const small = await Font.open(sub.bytes)
    expect(small.numGlyphs).toBe(6)
    expect([0, 1, 2, 3, 4, 5].map((x) => small.advanceOf(x))).toEqual([600, 500, 710, 540, 900, 1210])
    expect(small.names).toEqual(font.names)
    expect(small.metrics).toEqual(font.metrics)
    // the characters asked for whose glyphs are in the subset
    expect([small.glyphOf(0x6f), small.glyphOf(0x41), small.glyphOf(0x1f600), small.glyphOf(0x56)]).toEqual([3, 1, 0, 0])
    // nothing of the font's rules for runs: a PDF draws glyphs where the layout put them
    expect([small.gsub, small.gpos]).toEqual([null, null])
    // the composite names its parts by their new numbers
    const loca = dv(small.tables.get('loca')!)
    const glyf = small.tables.get('glyf')!
    const composite = dv(glyf.subarray(loca.getUint32(20), loca.getUint32(24)))
    expect([composite.getInt16(0), composite.getUint16(12), composite.getUint16(20)]).toEqual([-1, 1, 2])
    expect(sub.bytes.length).toBeLessThan(buildFont(BOXES).length / 2)
    // a whole file: its checksums add up as a font's do
    let total = 0
    for (let i = 0; i < sub.bytes.length; i += 4) total = (total + dv(sub.bytes).getUint32(i)) >>> 0
    expect(total).toBe(0xb1b0afba)
  })

  test('glyphs that are not the font’s are left out; the missing glyph is always there', async () => {
    const font = await open(BOXES)
    expect([...subsetFont(font, []).glyphs]).toEqual([[0, 0]])
    expect([...subsetFont(font, [-1, 59, 1e9, 3, 3]).glyphs]).toEqual([[0, 0], [3, 1]])
    expect((await Font.open(subsetFont(font, []).bytes)).numGlyphs).toBe(1)
  })

  for (const cff of [true, 'cid'] as const) {
    test(`PostScript outlines${cff === 'cid' ? ' of a CID-keyed font' : ''}: their charstrings and subroutines, renumbered`, async () => {
      const font = await open({ ...BOXES, composites: undefined, cff })
      expect(font.outlines).toBe('cff')
      const sub = subsetFont(font, [g('V'), g('a'), 58], [[0x56, g('V')], [0x61, g('a')]])
      expect([...sub.glyphs]).toEqual([[0, 0], [23, 1], [28, 2], [58, 3]])
      expect(sub.outlines).toBe('cff')
      const small = await Font.open(sub.bytes)
      expect([small.outlines, small.numGlyphs]).toEqual(['cff', 4])
      expect([0, 1, 2, 3].map((x) => small.advanceOf(x))).toEqual([600, 710, 400, 800])
      expect([small.glyphOf(0x56), small.glyphOf(0x61)]).toEqual([1, 2])
      // the CFF table alone is what a PDF embeds: it is the subset's, and subsets again to the same glyphs
      expect(Buffer.compare(Buffer.from(sub.cff!), Buffer.from(small.tables.get('CFF ')!))).toBe(0)
      const again = subsetFont(small, [1, 3])
      expect([...again.glyphs]).toEqual([[0, 0], [1, 1], [3, 2]])
      const smaller = await Font.open(again.bytes)
      expect([0, 1, 2].map((x) => smaller.advanceOf(x))).toEqual([600, 710, 800])
      // the charstrings are the font's own, byte for byte: the last glyph's ends the table's index
      const charstring = (f: Font, glyph: number, of: number) => {
        const t = f.tables.get('CFF ')!
        const text = Buffer.from(t).toString('latin1')
        // V: its width, a move, three lines, the call of the subroutine, the end
        const v = Buffer.from([28, 2, 198, 189, 139, 21, 28, 2, 98, 6, 28, 2, 188, 7, 28, 253, 158, 6, 32, 10, 14]).toString('latin1')
        return glyph === of ? text.includes(v) : null
      }
      expect(charstring(font, 23, 23)).toBe(true)
      expect(charstring(small, 1, 1)).toBe(true)
      expect(charstring(smaller, 1, 1)).toBe(true)
    })
  }

  test('the text of each glyph, for a PDF: ToUnicode', () => {
    const cmap = toUnicodeCMap([[3, 'ffi'], [1, 'A'], [2, '\u{1f600}'], [1, 'B'], [4, '']])
    expect(cmap).toContain('1 begincodespacerange\n<0000> <FFFF>\nendcodespacerange\n')
    // sorted, the last text given for a glyph, a ligature as its letters, a character past U+FFFF as its two halves
    expect(cmap).toContain('3 beginbfchar\n<0001> <0042>\n<0002> <D83DDE00>\n<0003> <006600660069>\nendbfchar\n')
    expect(cmap.startsWith('/CIDInit /ProcSet findresource begin')).toBe(true)
    expect(cmap.endsWith('endcmap\nCMapName currentdict /CMap defineresource pop\nend\nend\n')).toBe(true)
    // no more than 100 entries in a block
    const many = toUnicodeCMap(Array.from({ length: 250 }, (_, i) => [i + 1, 'x'] as [number, string]))
    expect(many.match(/\d+ beginbfchar/g)).toEqual(['100 beginbfchar', '100 beginbfchar', '50 beginbfchar'])
  })
})

describe('the font files a host declares', () => {
  const regular = buildFont({ ...BOXES, gpos: GPOS, gsub: GSUB, family: 'Carlito' })
  const bold = buildFont({ ...BOXES, family: 'Carlito', subfamily: 'Bold', weight: 700, advances: ADVANCES.map((w) => w + 20) })
  const italic = buildFont({ ...BOXES, family: 'Carlito', subfamily: 'Italic', italic: true, advances: ADVANCES.map((w) => w - 20) })
  // a second family with one more character, twice as large an em
  const other = buildFont({ ...BOXES, family: 'Liberation Sans', upem: 2000, advances: [...ADVANCES, 1500], chars: [...CHARS, [0x20ac, 59]] })
  const sources: FontSource[] = [
    { family: 'Carlito', src: regular },
    { family: 'Carlito', src: bold, weight: 700 },
    { family: 'Carlito', src: italic, style: 'italic' },
    { family: 'Liberation Sans', src: other },
  ]

  test('text in a document’s font is measured in the declared file its font list names', async () => {
    const book = new FontBook(sources)
    const m = await book.measure({ text: 'AV office', font: cssFont('Calibri')!, size: 20 })
    // Calibri is not declared, Carlito (which its list names next) is
    expect(m.pieces.map((p) => [p.from, p.to, p.family])).toEqual([[0, 9, 'carlito']])
    expect(m.width).toBeCloseTo(((500 - 80 + 710 + 250 + 540 + 900 + 420 + 440) * 20) / 1000, 9)
    expect([m.ascent, m.descent, m.lineGap]).toEqual([16, 4, 0])
    expect([m.exact, m.unsure]).toEqual([true, []])
    expect(m.pieces[0].run!.glyphs).toEqual([g('A'), g('V'), 1, g('o'), 55, g('c'), g('e')])
    const plain = await book.measure({ text: 'AV office', font: 'Carlito', size: 10, kerning: false, ligatures: false, letterSpacing: 1 })
    expect(plain.width).toBeCloseTo((sum(widths('AV office')) * 10) / 1000 + 9, 9)
  })

  test('weight and style choose among a family’s files as CSS does', async () => {
    const book = new FontBook(sources)
    const width = async (weight: number, italicText = false) => (await book.measure({ text: 'A', font: 'Carlito', size: 1000, weight, italic: italicText })).width
    expect(await width(400)).toBe(500)
    expect(await width(700)).toBe(520)
    expect(await width(900)).toBe(520)
    // 500 looks at 400 first, 600 at the bolder file
    expect(await width(500)).toBe(500)
    expect(await width(600)).toBe(520)
    expect(await width(400, true)).toBe(480)
    // bold italic: the italic file is the nearest (its style first), thickened by the browser
    const boldItalic = await book.measure({ text: 'A', font: 'Carlito', size: 1000, weight: 700, italic: true })
    expect([boldItalic.width, boldItalic.exact, boldItalic.unsure]).toEqual([480, false, ['synthetic']])
    // a family with a regular file only: bold and italic text are the browser's own doing
    const one = new FontBook([{ family: 'Liberation Sans', src: other }])
    expect((await one.measure({ text: 'A', font: 'Liberation Sans', size: 2000, weight: 700 })).unsure).toEqual(['synthetic'])
    expect((await one.measure({ text: 'A', font: 'Liberation Sans', size: 2000, italic: true })).unsure).toEqual(['synthetic'])
    expect((await one.measure({ text: 'A', font: 'Liberation Sans', size: 2000 })).exact).toBe(true)
  })

  test('a character the first font lacks is set in the next font of the list that has it', async () => {
    const book = new FontBook(sources)
    const m = await book.measure({ text: 'A€V', font: 'Carlito, "Liberation Sans", sans-serif', size: 100 })
    expect(m.pieces.map((p) => [p.from, p.to, p.family, p.width])).toEqual([
      [0, 1, 'carlito', 50],
      [1, 2, 'liberation sans', 75],
      [2, 3, 'carlito', 71],
    ])
    expect([m.width, m.exact]).toEqual([196, true])
    // no file of the list has it
    const none = await book.measure({ text: 'A中', font: 'Carlito, "Liberation Sans"', size: 100 })
    expect(none.pieces.map((p) => [p.from, p.to, p.family, p.width])).toEqual([
      [0, 1, 'carlito', 50],
      [1, 2, '', 0],
    ])
    expect([none.exact, none.unsure]).toEqual([false, ['missing']])
  })

  test('unicode-range keeps a file to the characters it is declared for', async () => {
    expect(unicodeRanges('U+0-7F, u+131 ,U+4??, nonsense')).toEqual([[0, 0x7f], [0x131, 0x131], [0x400, 0x4ff]])
    expect(unicodeRanges(undefined)).toBeNull()
    expect(unicodeRanges('all')).toBeNull()
    const book = new FontBook([
      { family: 'Split', src: regular, unicodeRange: 'U+41-5A' },
      { family: 'Split', src: bold, unicodeRange: 'U+61-7A' },
    ])
    const m = await book.measure({ text: 'Aa', font: 'Split', size: 1000 })
    expect(m.pieces.map((p) => p.width)).toEqual([500, 420])
  })

  test('text no declared file serves: nothing is known, and it says so', async () => {
    const book = new FontBook(sources)
    const m = await book.measure({ text: 'AV', font: cssFont('Cambria')!, size: 12 })
    expect(m).toEqual({ width: 0, ascent: 0, descent: 0, lineGap: 0, exact: false, unsure: ['font'], pieces: [{ from: 0, to: 2, family: '', font: null, run: null, width: 0 }] })
    expect([...book.missing]).toEqual(['cambria'])
    expect((await new FontBook([]).measure({ text: '', font: 'x', size: 12 })).pieces).toEqual([])
  })

  test('files are fetched when text first asks for them, once; one that fails is told and passed over', async () => {
    const asked: string[] = []
    const errors: string[] = []
    const book = new FontBook(
      [
        { family: 'Carlito', src: 'https://fonts.test/carlito.ttf' },
        { family: 'Carlito', src: 'https://fonts.test/carlito-bold.ttf', weight: 700 },
        { family: 'Caladea', src: 'https://fonts.test/gone.ttf' },
        { family: 'Broken', src: new TextEncoder().encode('not a font') },
        { family: 'Liberation Sans', src: new Blob([other]) },
      ],
      {
        fetch: async (url) => {
          asked.push(url)
          if (url.endsWith('gone.ttf')) throw new Error('HTTP 404')
          return url.includes('bold') ? bold : regular
        },
        onError: (e, source) => errors.push(`${source.family}: ${e.message}`),
      },
    )
    expect(asked).toEqual([])
    expect((await book.measure({ text: 'A', font: 'Carlito', size: 1000 })).width).toBe(500)
    expect((await book.measure({ text: 'AA', font: 'Carlito', size: 1000 })).width).toBe(1000)
    expect(asked).toEqual(['https://fonts.test/carlito.ttf'])
    const fallen = await book.measure({ text: 'A', font: 'Caladea, Broken, "Liberation Sans"', size: 2000 })
    expect([fallen.width, fallen.pieces[0].family]).toEqual([500, 'liberation sans'])
    expect(asked).toEqual(['https://fonts.test/carlito.ttf', 'https://fonts.test/gone.ttf'])
    expect(errors).toEqual(['Caladea: HTTP 404', 'Broken: Not a font file (TTF, OTF, TTC, WOFF or WOFF2)'])
  })

  test('a variable font asked for at another weight than its tables are of', async () => {
    const variable = buildFont({ ...BOXES, family: 'Vary', extra: { fvar: [0, 1, 0, 0] } })
    const book = new FontBook([{ family: 'Vary', src: variable, weight: '100 900' }])
    expect((await book.measure({ text: 'A', font: 'Vary', size: 10, weight: 400 })).exact).toBe(true)
    const bolder = await book.measure({ text: 'A', font: 'Vary', size: 10, weight: 700 })
    expect([bolder.exact, bolder.unsure]).toEqual([false, ['variable']])
    // a file declared for one weight is set at that weight by a browser, whatever weight its tables are of
    const one = new FontBook([{ family: 'Vary', src: variable, weight: 700 }])
    const at700 = await one.measure({ text: 'A', font: 'Vary', size: 10, weight: 700 })
    expect([at700.exact, at700.unsure]).toEqual([false, ['variable']])
    // …and a file declared for none at 400: the tables of this one are of 200
    const light = buildFont({ ...BOXES, family: 'Vary', weight: 200, extra: { fvar: [0, 1, 0, 0] } })
    const plain = await new FontBook([{ family: 'Vary', src: light }]).measure({ text: 'A', font: 'Vary', size: 10 })
    expect([plain.exact, plain.unsure]).toEqual([false, ['variable']])
    // bold text in a file declared for 400 alone: the 400 of its tables, thickened
    const thick = await new FontBook([{ family: 'Vary', src: variable }]).measure({ text: 'A', font: 'Vary', size: 10, weight: 700 })
    expect(thick.unsure).toEqual(['synthetic'])
  })

  test('in the worker: fonts.set and fonts.measure, with no document open', async () => {
    const events: { event: string; data: unknown }[] = []
    const engine = new Engine((event, data) => events.push({ event, data }))
    expect(await engine.call('fonts.measure', [[{ text: 'A', font: 'Carlito', size: 10 }]])).toEqual([
      { width: 0, ascent: 0, descent: 0, lineGap: 0, exact: false, unsure: ['font'], pieces: [{ from: 0, to: 1, family: '', width: 0 }] },
    ])
    await engine.call('fonts.set', [[...sources, { family: 'Caladea', src: new Uint8Array(40) }]])
    const [a, b] = (await engine.call('fonts.measure', [
      [
        { text: 'AV office', font: cssFont('Calibri')!, size: 20 },
        { text: 'A', font: 'Caladea, Carlito', size: 10, weight: 700 },
      ],
    ])) as MeasuredText[]
    // 3680 units at 20 to the 1000: kerned and with its ligature
    expect(a.width).toBeCloseTo(73.6, 9)
    expect(a).toEqual({ width: a.width, ascent: 16, descent: 4, lineGap: 0, exact: true, unsure: [], pieces: [{ from: 0, to: 9, family: 'carlito', width: a.width }] })
    // the bold file of the second family: the first one's file is no font
    expect([b.width, b.exact, b.pieces[0].family]).toEqual([5.2, true, 'carlito'])
    // what crosses to the page is plain data
    expect(structuredClone([a, b])).toEqual([a, b])
    expect(events).toEqual([{ event: 'font:error', data: { family: 'Caladea', message: 'Not a font file (TTF, OTF, TTC, WOFF or WOFF2)' } }])
  })

  test('from a page: addresses are resolved against it, and one that is no address fails alone, when text asks for it', async () => {
    const saved = { document: globalThis.document, fetch: globalThis.fetch }
    const asked: string[] = []
    const errors: Events['font:error'][] = []
    const client = await DocumentClient.open(Bun.file(`${import.meta.dir}/../playground/public/samples/sample.xlsx`) as unknown as Blob, { worker: false })
    client.on('font:error', (e) => errors.push(e))
    try {
      Object.assign(globalThis, {
        document: { baseURI: 'https://host.example/app/' },
        fetch: async (url: string) => {
          asked.push(url)
          if (!url.startsWith('https://host.example/')) throw new TypeError(`${url} cannot be parsed as a URL`)
          return new Response(regular)
        },
      })
      // (the call itself does not throw: the other files are set)
      await client.call('fonts.set', [
        { family: 'Carlito', src: 'fonts/carlito.ttf' },
        { family: 'Broken', src: 'http://[bad' },
      ])
      const [a, b] = await client.call('fonts.measure', [
        { text: 'A', font: 'Carlito', size: 10 },
        { text: 'A', font: 'Broken', size: 10 },
      ])
      expect(asked).toEqual(['https://host.example/app/fonts/carlito.ttf', 'http://[bad'])
      expect([a.width, a.exact, b.unsure]).toEqual([5, true, ['font']])
      expect(errors).toEqual([{ family: 'Broken', message: 'http://[bad cannot be parsed as a URL' }])
    } finally {
      Object.assign(globalThis, saved)
      client.destroy()
    }
  })
})
