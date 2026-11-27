/**
 * Fonts read from their files against a browser: the width `layoutText` gives for a text is the width the browser
 * lays it out at in the same font (kerning and ligatures as the font's tables say), and the font files written here
 * (a subset, a font unpacked from WOFF2) are fonts a browser takes. The fonts are made by test/font-tables.ts: every
 * width in them is a whole number of pixels at 100px, so that nothing depends on how a browser rounds.
 */
import { expect, test } from 'bun:test'
import { readFileSync } from 'node:fs'
import { Font, layoutText, subsetFont, type RunOptions } from '../src/core/font'
import { buildFont, classPairPos, kernTable, layoutTable, ligatureSubst, pairPos, woff, type FontSpec } from '../test/font-tables'
import { useBrowser, viewer } from './harness'

useBrowser()

// glyph 0 the missing glyph, 1 the space, 2–27 A–Z, 28–53 a–z, 54 fi, 55 ffi, 56 unused, 57 a composite of A and V, 58 an emoji
const CHARS: [number, number][] = [[0x20, 1]]
for (let i = 0; i < 26; i++) CHARS.push([0x41 + i, 2 + i], [0x61 + i, 28 + i])
CHARS.push([0xfb01, 54], [0x1f600, 58])
const ADVANCES = [600, 250, ...Array.from({ length: 26 }, (_, i) => 500 + i * 10), ...Array.from({ length: 26 }, (_, i) => 400 + i * 10), 700, 900, 520, 1210, 800]
const g = (c: string) => CHARS.find((x) => x[0] === c.codePointAt(0))![1]
const BOXES: FontSpec = { advances: ADVANCES, chars: CHARS, composites: { 57: [g('A'), g('V')] }, cmap12: true }
const GSUB = layoutTable(false, [['liga', [0]]], [{ type: 4, subtables: [ligatureSubst([[[g('f'), g('f'), g('i')], 55], [[g('f'), g('i')], 54]])] }])
const GPOS = layoutTable(true, [['kern', [0, 1]]], [
  { type: 2, subtables: [pairPos([[g('A'), g('V'), -80], [g('T'), g('o'), -60]])] },
  { type: 2, extension: true, subtables: [classPairPos([[g('W'), g('Y')]], [[g('a'), g('e')]], [[-40]])] },
])

interface Sample {
  font: string
  text: string
  css?: string
  opts?: RunOptions
}

/** the fonts loaded in the page (a file a browser refuses fails here, by its name), then the width of each sample at 100px */
async function inBrowser(fonts: Record<string, Uint8Array>, samples: Sample[]): Promise<number[]> {
  const v = await viewer({}, { wait: false })
  const out = await v.page.evaluate(
    async ([files, list]) => {
      for (const [name, bytes] of files) {
        const face = new FontFace(name, new Uint8Array(bytes))
        await face.load().catch((e) => {
          throw new Error(`${name} is not a font to this browser: ${e}`)
        })
        document.fonts.add(face)
      }
      return list.map((s) => {
        const span = document.createElement('span')
        span.style.cssText = `position:absolute;left:0;top:0;white-space:pre;font:100px "${s.font}";${s.css ?? ''}`
        span.textContent = s.text
        document.body.append(span)
        const w = span.getBoundingClientRect().width
        span.remove()
        return w
      })
    },
    [Object.entries(fonts).map(([name, bytes]) => [name, [...bytes]] as const), samples.map(({ font, text, css }) => ({ font, text, css }))] as const,
  )
  expect(v.errors).toEqual([])
  return out
}

/** the same widths from the font files alone */
async function fromFiles(fonts: Record<string, Uint8Array>, samples: Sample[]): Promise<number[]> {
  const read = new Map<string, Font>()
  for (const [name, bytes] of Object.entries(fonts)) read.set(name, await Font.open(bytes))
  return samples.map((s) => {
    const font = read.get(s.font)!
    const run = layoutText(font, s.text, s.opts)
    expect([s.font, s.text, run.exact, run.missing]).toEqual([s.font, s.text, true, 0])
    return (run.width * 100) / font.unitsPerEm
  })
}

test('text is as wide from the font file as the browser lays it out: kerning, ligatures, spacing', async () => {
  const fonts = {
    'Read Plain': buildFont(BOXES),
    'Read Layout': buildFont({ ...BOXES, gsub: GSUB, gpos: GPOS, kern: kernTable([[g('A'), g('V'), -900]]) }),
    // (a browser gives each of the two glyphs half of a kern table's value: halves that are whole pixels too)
    'Read Kern': buildFont({ ...BOXES, kern: kernTable([[g('A'), g('V'), -80], [g('T'), g('o'), -60]]) }),
    'Read Short': buildFont({ ...BOXES, longMetrics: 30 }),
    'Read Woff': await woff(buildFont({ ...BOXES, gsub: GSUB, gpos: GPOS })),
  }
  const text = 'AV To office WaYe fit VA'
  const samples: Sample[] = [
    { font: 'Read Plain', text },
    { font: 'Read Plain', text: 'A \u{1f600} z' },
    { font: 'Read Layout', text },
    { font: 'Read Layout', text, css: 'font-kerning:none', opts: { kerning: false } },
    { font: 'Read Layout', text, css: 'font-variant-ligatures:none', opts: { ligatures: false } },
    { font: 'Read Layout', text, css: 'font-kerning:none;font-variant-ligatures:none', opts: { kerning: false, ligatures: false } },
    // spaced letters are not joined into ligatures, by CSS and here
    { font: 'Read Layout', text, css: 'letter-spacing:3px', opts: { letterSpacing: 30 } },
    { font: 'Read Kern', text },
    { font: 'Read Kern', text, css: 'font-kerning:none', opts: { kerning: false } },
    { font: 'Read Short', text },
    { font: 'Read Woff', text },
  ]
  const [dom, files] = [await inBrowser(fonts, samples), await fromFiles(fonts, samples)]
  // the fonts do what the test means them to: kerning and ligatures change the width
  expect(files[2]).toBeLessThan(files[3])
  expect(files[2]).toBeLessThan(files[4])
  expect(new Set(files.slice(2, 7)).size).toBe(5)
  expect(files[7]).toBe(files[8] - 14)
  samples.forEach((s, i) => expect([s.font, s.css ?? '', dom[i]]).toEqual([s.font, s.css ?? '', files[i]]))
})

test('a subset, and a font unpacked from WOFF2, are fonts a browser takes, with the widths of the font they are of', async () => {
  const layout = await Font.open(buildFont({ ...BOXES, gsub: GSUB, gpos: GPOS }))
  const chars: [number, number][] = [...'AVo ', '\u{1f600}'].map((c) => [c.codePointAt(0)!, layout.glyphOf(c.codePointAt(0)!)])
  const packed = new Uint8Array(readFileSync(new URL('../test/fixtures/font/boxes.woff2', import.meta.url)))
  const fonts: Record<string, Uint8Array> = {
    'Sub TrueType': subsetFont(layout, [...chars.map((c) => c[1]), 57], chars).bytes,
    'Sub Empty': subsetFont(layout, []).bytes,
    'Packed': packed,
    'Unpacked': (await Font.open(packed)).sfnt(),
  }
  for (const cff of [true, 'cid'] as const) {
    const whole = buildFont({ ...BOXES, composites: undefined, cff })
    fonts[`Whole ${cff}`] = whole
    fonts[`Sub ${cff}`] = subsetFont(await Font.open(whole), chars.map((c) => c[1]), chars).bytes
  }
  const text = 'AVo \u{1f600} oVA'
  const samples: Sample[] = Object.keys(fonts)
    .filter((font) => font !== 'Sub Empty')
    .map((font) => ({ font, text }))
  const [dom, files] = [await inBrowser(fonts, samples), await fromFiles(fonts, samples)]
  samples.forEach((s, i) => expect([s.font, dom[i]]).toEqual([s.font, files[i]]))
  // all the same text in the same boxes
  expect(new Set(files).size).toBe(1)
  expect(files[0]).toBe((500 + 710 + 540 + 250 + 800 + 250 + 540 + 710 + 500) / 10)
})

test('the reader runs in the page too: a WOFF2 file is unpacked where the browser unpacks Brotli, and refused in words where it does not', async () => {
  const packed = new Uint8Array(readFileSync(new URL('../test/fixtures/font/boxes.woff2', import.meta.url)))
  const v = await viewer({}, { wait: false })
  const r = await v.page.evaluate(
    async ([path, bytes]) => {
      const { Font, layoutText } = (await import(path as string)) as typeof import('../src/core/font')
      let brotli = true
      try {
        new DecompressionStream('brotli' as CompressionFormat)
      } catch {
        brotli = false
      }
      try {
        const font = await Font.open(new Uint8Array(bytes as number[]))
        return { brotli, format: font.format, glyphs: font.numGlyphs, width: layoutText(font, 'AV').width, error: '' }
      } catch (e) {
        return { brotli, format: '', glyphs: 0, width: 0, error: String(e) }
      }
    },
    ['/@fs/' + process.cwd() + '/src/core/font/index.ts', [...packed]] as const,
  )
  // (Chromium 156 does not, Firefox 157 does)
  if (r.brotli) expect(r).toEqual({ brotli: true, format: 'woff2', glyphs: 59, width: 1210, error: '' })
  else expect(r.error).toBe('FontError: This font is a WOFF2 file, which this browser cannot unpack: supply it as TTF, OTF or WOFF')
  expect(v.errors).toEqual([])
})
