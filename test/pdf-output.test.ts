import { describe, expect, test } from 'bun:test'
import { readFileSync } from 'node:fs'
import { Font, FontBook, subsetFont } from '../src/core/font'
import { decodeGif, decodePng, jpegInfo } from '../src/core/draw/image'
import { ellipsePath, parseColor, parsePath, rectPath } from '../src/core/draw/path'
import { pageToSvg } from '../src/core/draw/svg'
import { parseTransform, svgItems } from '../src/core/draw/svg-in'
import { layoutParagraphs, type ParaSpec, type RunStyle } from '../src/core/draw/text'
import type { DrawItem, Page } from '../src/core/draw/types'
import { writePdf } from '../src/core/pdf/writer'
import { gradient, slidePage } from '../src/core/pptx/draw'
import type { SlideModel } from '../src/core/pptx/types'
import { Engine } from '../src/core/worker/engine'
import { testFont } from './font-builder'
import { buildFont } from './font-tables'

const BLACK = { r: 0, g: 0, b: 0, a: 1 }

/** a font whose printable characters are all 1000 units (1 em) wide, the space 500 */
const EM = testFont({ family: 'Test Em', advance: 1000, space: 500 })
const book = (family = 'Test Em', bytes: Uint8Array = EM) => new FontBook([{ family, src: bytes }])

async function inflate(b: Uint8Array): Promise<Uint8Array> {
  return new Uint8Array(await new Response(new Blob([b as Uint8Array<ArrayBuffer>]).stream().pipeThrough(new DecompressionStream('deflate'))).arrayBuffer())
}

async function deflate(b: Uint8Array): Promise<Uint8Array> {
  return new Uint8Array(await new Response(new Blob([b as Uint8Array<ArrayBuffer>]).stream().pipeThrough(new CompressionStream('deflate') as unknown as TransformStream<Uint8Array, Uint8Array>)).arrayBuffer())
}

const latin1 = (b: Uint8Array) => {
  let s = ''
  for (let i = 0; i < b.length; i += 0x8000) s += String.fromCharCode(...b.subarray(i, i + 0x8000))
  return s
}

/** a PDF's objects by number, streams unpacked */
async function readPdf(pdf: Uint8Array): Promise<{ text: string; objects: Map<number, { dict: string; stream?: Uint8Array }> }> {
  const text = latin1(pdf)
  const objects = new Map<number, { dict: string; stream?: Uint8Array }>()
  // every offset of the cross-reference table starts its object
  const xref = Number(/startxref\n(\d+)/.exec(text)![1])
  expect(text.slice(xref, xref + 4)).toBe('xref')
  const count = Number(/xref\n0 (\d+)/.exec(text.slice(xref))![1])
  const rows = text.slice(xref).split('\n').slice(2, 2 + count)
  for (let i = 1; i < count; i++) {
    const at = Number(rows[i].slice(0, 10))
    expect(text.slice(at, at + `${i} 0 obj`.length)).toBe(`${i} 0 obj`)
    const end = text.indexOf('\nendobj', at)
    const body = text.slice(at + `${i} 0 obj\n`.length, end)
    const s = body.indexOf('\nstream\n')
    if (s < 0) objects.set(i, { dict: body })
    else {
      const dict = body.slice(0, s)
      const len = Number(/\/Length (\d+)/.exec(dict)![1])
      const start = at + `${i} 0 obj\n`.length + s + '\nstream\n'.length
      let data = pdf.subarray(start, start + len)
      if (dict.includes('/FlateDecode')) data = await inflate(data)
      objects.set(i, { dict, stream: data })
    }
  }
  return { text, objects }
}

/** the text a ToUnicode map gives the glyph codes of a TJ: what a reader copies */
function unicodeOf(cmap: string, hex: string): string {
  const map = new Map<string, string>()
  for (const m of cmap.matchAll(/<([0-9A-F]{4})> <([0-9A-F]+)>/gi)) {
    let s = ''
    for (let i = 0; i < m[2].length; i += 4) s += String.fromCharCode(parseInt(m[2].slice(i, i + 4), 16))
    map.set(m[1].toUpperCase(), s)
  }
  let out = ''
  for (let i = 0; i < hex.length; i += 4) out += map.get(hex.slice(i, i + 4).toUpperCase()) ?? '?'
  return out
}

describe('paths, transforms, colors', () => {
  test('SVG path data: relative commands, H / V, curves and arcs end where they should', () => {
    const d = parsePath('m10 20 h30 v40 H0 l5 5 z')
    expect(d.map((s) => s.k).join('')).toBe('MLLLLZ')
    expect(d[1]).toEqual({ k: 'L', x: 40, y: 20 })
    expect(d[2]).toEqual({ k: 'L', x: 40, y: 60 })
    expect(d[4]).toEqual({ k: 'L', x: 5, y: 65 })
    const c = parsePath('M0 0 Q10 0 10 10 T20 20 S30 30 40 40 A10 10 0 0 1 60 40')
    const last = c[c.length - 1] as { x: number; y: number }
    expect(last.x).toBeCloseTo(60)
    expect(last.y).toBeCloseTo(40)
    // a half circle: two quarter curves
    expect(c.filter((s) => s.k === 'C').length).toBe(5)
    // numbers after a move are lines; what follows an error is dropped
    expect(parsePath('M0 0 10 10 20 0').map((s) => s.k).join('')).toBe('MLL')
    expect(parsePath('M0 0 L10 10 L x').length).toBe(2)
  })

  test('transforms apply right to left, as SVG reads them', () => {
    const m = parseTransform('translate(10 20) scale(2)')
    expect(m).toEqual([2, 0, 0, 2, 10, 20])
    const r = parseTransform('rotate(90 10 10)')
    // (20, 10) turns about (10, 10) to (10, 20)
    expect(r[0] * 20 + r[2] * 10 + r[4]).toBeCloseTo(10)
    expect(r[1] * 20 + r[3] * 10 + r[5]).toBeCloseTo(20)
  })

  test('CSS colors as numbers; anything else is none', () => {
    expect(parseColor('#f80')).toEqual({ r: 255, g: 136, b: 0, a: 1 })
    expect(parseColor('rgba(1,2,3,0.5)')).toEqual({ r: 1, g: 2, b: 3, a: 0.5 })
    expect(parseColor('url(x)')).toBeNull()
    expect(parseColor('transparent')).toBeNull()
  })

  test('the engine\'s gradients span the box as CSS spans them', () => {
    const g = gradient('linear-gradient(90deg, #ff0000 0%, rgba(0,0,255,0.5) 100%)', { x: 0, y: 0, w: 200, h: 100 })
    if (g?.t !== 'linear') throw new Error()
    expect([g.x0, g.y0, g.x1, g.y1].map((v) => Math.round(v))).toEqual([0, 50, 200, 50])
    expect(g.stops.map((s) => s.color.a)).toEqual([1, 0.5])
    const r = gradient('radial-gradient(circle, #fff 0%, #000 100%)', { x: 0, y: 0, w: 60, h: 80 })
    expect(r).toMatchObject({ t: 'radial', cx: 30, cy: 40, r: 50 })
  })
})

describe('pictures unpacked', () => {
  const png = async (w: number, h: number, type: number, depth: number, rows: number[][], extra: [string, number[]][] = [], interlace = 0) => {
    const chunk = (t: string, d: number[]) => {
      const out = [d.length >>> 24, (d.length >> 16) & 255, (d.length >> 8) & 255, d.length & 255, ...[...t].map((c) => c.charCodeAt(0)), ...d, 0, 0, 0, 0]
      return out
    }
    const raw = await deflate(new Uint8Array(rows.flatMap((r) => [0, ...r])))
    const ihdr = [0, 0, 0, w, 0, 0, 0, h, depth, type, 0, 0, interlace]
    return new Uint8Array([0x89, 0x50, 0x4e, 0x47, 13, 10, 26, 10, ...chunk('IHDR', ihdr), ...extra.flatMap(([t, d]) => chunk(t, d)), ...chunk('IDAT', [...raw]), ...chunk('IEND', [])])
  }

  test('PNG: RGBA, a palette with transparency, 1-bit gray', async () => {
    const rgba = await decodePng(await png(2, 1, 6, 8, [[255, 0, 0, 255, 0, 0, 255, 128]]))
    expect(rgba).toMatchObject({ w: 2, h: 1, channels: 3 })
    expect([...rgba.color]).toEqual([255, 0, 0, 0, 0, 255])
    expect([...rgba.alpha!]).toEqual([255, 128])
    const pal = await decodePng(await png(2, 1, 3, 8, [[0, 1]], [['PLTE', [10, 20, 30, 40, 50, 60]], ['tRNS', [0]]]))
    expect([...pal.color]).toEqual([10, 20, 30, 40, 50, 60])
    expect([...pal.alpha!]).toEqual([0, 255])
    const bits = await decodePng(await png(8, 1, 0, 1, [[0b10100000]]))
    expect(bits.channels).toBe(1)
    expect([...bits.color]).toEqual([255, 0, 255, 0, 0, 0, 0, 0])
    expect(bits.alpha).toBeUndefined()
  })

  test('PNG: filters undone (sub, up, average, Paeth)', async () => {
    // one row per filter, each reading back 1 2 3
    const rows = [[1, 1, 1, 1], [2, 0, 0, 0], [3, 1, 1, 1], [4, 0, 0, 0]]
    const raw = await deflate(new Uint8Array(rows.flat()))
    const chunk = (t: string, d: number[]) => [d.length >>> 24, (d.length >> 16) & 255, (d.length >> 8) & 255, d.length & 255, ...[...t].map((c) => c.charCodeAt(0)), ...d, 0, 0, 0, 0]
    const file = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 13, 10, 26, 10, ...chunk('IHDR', [0, 0, 0, 3, 0, 0, 0, 4, 8, 0, 0, 0, 0]), ...chunk('IDAT', [...raw]), ...chunk('IEND', [])])
    const px = await decodePng(file)
    expect([...px.color]).toEqual([1, 2, 3, 1, 2, 3, 1, 2, 3, 1, 2, 3])
  })

  test('GIF: the first frame, its transparent color as alpha', () => {
    const red = Uint8Array.from(atob('R0lGODlhAQABAPAAAP8AAP///yH5BAAAAAAALAAAAAABAAEAAAICRAEAOw=='), (c) => c.charCodeAt(0))
    const px = decodeGif(red)
    expect([...px.color]).toEqual([255, 0, 0])
    expect(px.alpha).toBeUndefined()
    const clear = Uint8Array.from(atob('R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7'), (c) => c.charCodeAt(0))
    expect([...decodeGif(clear).alpha!]).toEqual([0])
  })

  test('JPEG: size and color components from its frame header', () => {
    const sof = [0xff, 0xd8, 0xff, 0xc0, 0, 17, 8, 0, 20, 0, 30, 3, 1, 0x22, 0, 2, 0x11, 1, 3, 0x11, 1, 0xff, 0xd9]
    expect(jpegInfo(new Uint8Array(sof))).toEqual({ w: 30, h: 20, components: 3, adobe: false })
  })
})

describe('text laid out', () => {
  const style = (over: Partial<RunStyle> = {}): RunStyle => ({ font: 'Test Em', size: 10, color: BLACK, ...over })
  const para = (text: string, over: Partial<ParaSpec> = {}): ParaSpec => ({ runs: [{ text, style: style() }], align: 'left', marL: 0, indent: 0, before: 0, after: 0, emptySize: 10, ...over })
  const texts = (items: DrawItem[]) => items.filter((i): i is Extract<DrawItem, { t: 'text' }> => i.t === 'text')

  test('lines break at spaces, a word longer than the line inside it', async () => {
    // each letter 10 px, a space 5 px: "aaa bbb" is 65 px
    const out = await layoutParagraphs(book(), [para('aaa bbb cc'), para('abcdefghij')], 50)
    const lines = texts(out.items).map((t) => [t.text.slice(t.clusters[0], t.clusters[t.clusters.length - 1] + 1), t.x, t.y])
    expect(lines).toEqual([
      ['aaa ', 0, 9],
      ['bbb ', 0, 21],
      ['cc', 0, 33],
      ['abcde', 0, 45],
      ['fghij', 0, 57],
    ])
    // 1.2 × 10 px a line, ascent 8 and descent 2 centred in it: the baseline 9 px down
    expect(out.height).toBeCloseTo(60)
  })

  test('aligned, justified, with a hanging bullet', async () => {
    const center = await layoutParagraphs(book(), [para('ab', { align: 'center' })], 100)
    expect(texts(center.items)[0].x).toBe(40)
    const right = await layoutParagraphs(book(), [para('ab', { align: 'right' })], 100)
    expect(texts(right.items)[0].x).toBe(80)
    // justified: the first line's one space takes what is left; the last line is not stretched
    const just = await layoutParagraphs(book(), [para('aa bb cc', { align: 'justify' })], 60)
    const first = texts(just.items)[0]
    // "aa bb" ends at 60; the space after it is not stretched
    expect(first.advances.reduce((a, b) => a + b, 0)).toBeCloseTo(65)
    expect(first.advances[2]).toBeCloseTo(20)
    const bul = await layoutParagraphs(book(), [para('ab', { marL: 30, indent: -30, bullet: { text: '•', style: style({ font: 'Nothing' }) } })], 100)
    const t = texts(bul.items)
    // the bullet at the hanging indent, the text at the margin
    expect(t[0].x).toBe(0)
    expect(t[0].font).toBe('Helvetica')
    expect(t[1].x).toBe(30)
  })

  test('text no file serves is set in Helvetica and reported', async () => {
    const missing: string[] = []
    const out = await layoutParagraphs(book(), [{ ...para(''), runs: [{ text: 'Größe', style: style({ font: '"Calibri", Carlito, sans-serif', bold: true }) }] }], 500, { onMissing: (f) => missing.push(f) })
    const t = texts(out.items)[0]
    expect(t.font).toBe('Helvetica-Bold')
    // Windows-1252 codes; ß is 0xDF
    expect(t.glyphs).toEqual([0x47, 0x72, 0xf6, 0xdf, 0x65])
    expect(missing).toEqual(['"Calibri", Carlito, sans-serif'])
  })

  test('underline, highlight and link around a run', async () => {
    const out = await layoutParagraphs(book(), [{ ...para(''), runs: [{ text: 'ab', style: style({ underline: true, highlight: { r: 255, g: 255, b: 0, a: 1 }, href: 'https://example.com/' }) }] }], 100)
    expect(out.items.map((i) => i.t)).toEqual(['path', 'text', 'path', 'link'])
    expect(out.items[3]).toMatchObject({ t: 'link', x: 0, w: 20, url: 'https://example.com/' })
  })
})

describe('PDF files', () => {
  test('fonts embedded as subsets whose text can be copied; pictures, links, outline', async () => {
    const font = await Font.open(EM)
    const glyphs = [...'Hi there'].map((c) => font.glyphOf(c.codePointAt(0)!))
    const png = await (async () => {
      const raw = await deflate(new Uint8Array([0, 255, 0, 0, 128]))
      const chunk = (t: string, d: number[]) => [d.length >>> 24, (d.length >> 16) & 255, (d.length >> 8) & 255, d.length & 255, ...[...t].map((c) => c.charCodeAt(0)), ...d, 0, 0, 0, 0]
      return new Uint8Array([0x89, 0x50, 0x4e, 0x47, 13, 10, 26, 10, ...chunk('IHDR', [0, 0, 0, 1, 0, 0, 0, 1, 8, 6, 0, 0, 0]), ...chunk('IDAT', [...raw]), ...chunk('IEND', [])])
    })()
    const page: Page = {
      w: 400,
      h: 300,
      title: 'First',
      items: [
        { t: 'path', d: rectPath(10, 10, 100, 50), fill: { t: 'linear', x0: 10, y0: 0, x1: 110, y1: 0, stops: [{ at: 0, color: BLACK }, { at: 0.5, color: { r: 255, g: 0, b: 0, a: 1 } }, { at: 1, color: { r: 255, g: 255, b: 255, a: 1 } }] } },
        { t: 'path', d: ellipsePath(200, 50, 30, 20), fill: { t: 'color', color: { r: 0, g: 0, b: 255, a: 0.5 } }, stroke: { paint: { t: 'color', color: BLACK }, width: 2, dash: [4, 2] } },
        { t: 'text', x: 10, y: 100, size: 20, font, glyphs, advances: glyphs.map((g) => (font.advanceOf(g) * 20) / 1000), text: 'Hi there', clusters: [0, 1, 2, 3, 4, 5, 6, 7], fill: BLACK },
        { t: 'text', x: 10, y: 150, size: 20, font: 'Helvetica', glyphs: [0x41, 0x80], advances: [13.34, 11.12], text: 'A€', clusters: [0, 1], fill: BLACK },
        { t: 'image', image: { bytes: png, type: 'image/png' }, x: 300, y: 200, w: 50, h: 50 },
        { t: 'group', transform: [1, 0, 0, 1, 100, 0], items: [{ t: 'link', x: 10, y: 80, w: 80, h: 25, url: 'https://example.com/a(b)' }] },
      ],
    }
    const pdf = await writePdf([page, { w: 400, h: 300, items: [] }], { title: 'Täst' })
    expect(latin1(pdf.subarray(0, 8))).toBe('%PDF-1.7')
    const { text, objects } = await readPdf(pdf)
    const all = [...objects.values()]
    expect(text).toContain('/Count 2')
    // the page is 300 × 225 points
    expect(text).toContain('/MediaBox[0 0 300 225]')
    const type0 = all.find((o) => o.dict.includes('/Subtype/Type0'))!
    expect(type0.dict).toMatch(/\/BaseFont\/[A-Z]{6}\+TestEm/)
    expect(all.some((o) => o.dict.includes('/FontFile2'))).toBe(true)
    expect(all.some((o) => o.dict.includes('/BaseFont/Helvetica/Encoding/WinAnsiEncoding'))).toBe(true)
    // the subset is a font of its own, of the glyphs used
    const file = all.find((o) => o.dict.includes('/Length1'))!
    const sub = await Font.open(file.stream!)
    expect(sub.numGlyphs).toBe(new Set([0, ...glyphs]).size)
    // copying the text gives it back
    const cmapRef = Number(/\/ToUnicode (\d+)/.exec(type0.dict)![1])
    const cmap = latin1(objects.get(cmapRef)!.stream!)
    const content = all.map((o) => (o.stream && o.dict.length < 40 ? latin1(o.stream) : '')).find((s) => s.includes('TJ'))!
    const hex = [...content.matchAll(/\/F1 20 Tf\n[^\n]*Tm\n\[(.*)\] TJ/g)][0][1].replace(/[^<>0-9a-f]|<|>/gi, '')
    expect(unicodeOf(cmap, hex)).toBe('Hi there')
    // the gradient as a stitched shading, the half-transparent fill through a graphics state
    expect(text).toContain('/FunctionType 3')
    expect(text).toContain('/ca 0.5')
    expect(content).toContain('[4 2] 0 d')
    // the picture with its alpha as a soft mask
    expect(all.some((o) => o.dict.includes('/SMask'))).toBe(true)
    // the link moved with its group, in points, its address escaped
    expect(text).toContain('/Rect[82.5 146.25 142.5 165]')
    expect(text).toContain('/URI(https://example.com/a\\(b\\))')
    expect(text).toContain('/Title (First)')
    expect(text).toContain('/Title <FEFF005400e400730074>')
  })

  test('PostScript outlines are embedded as OpenType; a font that forbids embedding is replaced by Helvetica', async () => {
    const cff = await Font.open(buildFont({ family: 'Test Cff', cff: true, advances: [500, 600, 700], chars: [[0x41, 1], [0x42, 2]] }))
    const locked = await Font.open(buildFont({ family: 'Test Locked', fsType: 2, advances: [500, 600, 700], chars: [[0x41, 1], [0x42, 2]] }))
    const text = (font: Font): DrawItem => ({ t: 'text', x: 0, y: 20, size: 10, font, glyphs: [1, 2], advances: [6, 7], text: 'AB', clusters: [0, 1], fill: BLACK })
    const warnings: string[] = []
    const pdf = await writePdf([{ w: 100, h: 100, items: [text(cff), text(locked)] }], { onWarning: (w) => warnings.push(w) })
    const { text: s, objects } = await readPdf(pdf)
    expect(s).toContain('/Subtype/CIDFontType0')
    expect(s).toContain('/Subtype/OpenType')
    expect(s).not.toContain('TestLocked')
    expect(warnings).toEqual(['The font Test Locked may not be embedded: its text is set in Helvetica'])
    // its text in Helvetica: the characters, where the glyphs were
    const content = [...objects.values()].map((o) => (o.stream ? latin1(o.stream) : '')).find((c) => c.includes('/F2'))!
    expect(content).toContain('[<41>67<42>] TJ')
  })

  test('a picture in a format a PDF does not take is left out, and said so', async () => {
    const warnings: string[] = []
    const emf = new Uint8Array(64)
    emf.set([1, 0, 0, 0])
    emf.set([0x20, 0x45, 0x4d, 0x46], 40)
    await writePdf([{ w: 10, h: 10, items: [{ t: 'image', image: { bytes: emf, type: 'image/emf' }, x: 0, y: 0, w: 10, h: 10 }] }], { onWarning: (w) => warnings.push(w) })
    expect(warnings).toEqual(['Pictures in EMF are not drawn in a PDF yet'])
  })
})

describe('fonts read once, subsets from Node Buffers', () => {
  test('a file declared for two families is one font', async () => {
    const b = new FontBook([{ family: 'Calibri', src: EM }, { family: 'Carlito', src: EM }])
    const one = await b.measure({ text: 'a', font: 'Calibri', size: 10 })
    const two = await b.measure({ text: 'a', font: 'Carlito', size: 10 })
    expect(one.pieces[0].font).toBe(two.pieces[0].font)
  })

  test('subsetting a font read from a Buffer leaves the font as it was, and keeps each glyph\'s side bearing', async () => {
    const bytes = buildFont({ family: 'Test Lsb', advances: [500, 600, 700, 800], chars: [[0x41, 1], [0x42, 2], [0x43, 3]] })
    const font = await Font.open(Buffer.from(bytes))
    const before = new Uint8Array(font.tables.get('hhea')!)
    const sub = subsetFont(font, [3])
    expect(new Uint8Array(font.tables.get('hhea')!)).toEqual(before)
    const lsb = (f: Font, g: number) => new DataView(f.tables.get('hmtx')!.buffer, f.tables.get('hmtx')!.byteOffset).getInt16(g * 4 + 2)
    const again = await Font.open(sub.bytes)
    expect(lsb(again, sub.glyphs.get(3)!)).toBe(lsb(font, 3))
  })
})

describe('slides as pages', () => {
  const slide = (els: SlideModel['els'], over: Partial<SlideModel> = {}): SlideModel => ({ index: 0, w: 400, h: 300, els, ...over })
  const opts = { fonts: book(), part: async () => null }
  const flat = (items: DrawItem[]): DrawItem[] => items.flatMap((i) => (i.t === 'group' ? flat(i.items) : [i]))

  test('a shape\'s outline flips, its text turns with it but is not flipped', async () => {
    const page = await slidePage(
      slide([
        {
          t: 'shape',
          box: { x: 100, y: 50, w: 200, h: 100, rot: 90, flipH: true },
          geom: 'rect',
          fill: { t: 'solid', c: '#ff0000' },
          stroke: { c: '#000000', w: 2, tail: 'triangle' },
          text: { anchor: 'ctr', ins: [10, 5, 10, 5], wrap: true, paras: [{ algn: 'center', marL: 0, indent: 0, spcBef: 0, spcAft: 0, runs: [{ text: 'ab', sz: 7.5, font: 'Test Em' }], endSz: 7.5 }] },
        },
      ]),
      opts,
    )
    expect(page).toMatchObject({ w: 400, h: 300 })
    const shape = (page.items[0] as Extract<DrawItem, { t: 'group' }>).items[1] as Extract<DrawItem, { t: 'group' }>
    // turned about its centre (200, 100), then placed
    expect(shape.transform![0]).toBeCloseTo(0)
    expect(shape.transform![1]).toBeCloseTo(1)
    expect(shape.transform![4]).toBeCloseTo(250)
    expect(shape.transform![5]).toBeCloseTo(0)
    const outline = shape.items[0] as Extract<DrawItem, { t: 'group' }>
    expect(outline.transform).toEqual([-1, 0, 0, 1, 200, 0])
    // fill, outline, arrowhead
    expect(outline.items.map((i) => i.t)).toEqual(['path', 'path', 'path'])
    const text = flat(shape.items).find((i) => i.t === 'text') as Extract<DrawItem, { t: 'text' }>
    // 7.5 pt is 10 px: "ab" is 20 px wide, centred in 180 px; one 12 px line centred in 90
    const block = shape.items[1] as Extract<DrawItem, { t: 'group' }>
    expect(block.transform).toEqual([1, 0, 0, 1, 10, 5 + 39])
    expect(text.x).toBe(80)
  })

  test('a table\'s rows grow to their text', async () => {
    const cell = (text: string) => ({ text: { anchor: 't' as const, ins: [0, 0, 0, 0] as [number, number, number, number], wrap: true, paras: [{ algn: 'left' as const, marL: 0, indent: 0, spcBef: 0, spcAft: 0, runs: [{ text, sz: 7.5, font: 'Test Em' }], endSz: 7.5 }] }, colSpan: 1, rowSpan: 1, hidden: false, fill: { t: 'solid' as const, c: '#eeeeee' }, bb: { c: '#000000', w: 1 } })
    const page = await slidePage(slide([{ t: 'table', box: { x: 0, y: 0, w: 60, h: 20 }, cols: [30, 30], rows: [{ h: 5, cells: [cell('aaa bbb'), cell('c')] }, { h: 5, cells: [cell('d'), cell('e')] }] }]), opts)
    const items = flat(page.items)
    const borders = items.filter((i): i is Extract<DrawItem, { t: 'path' }> => i.t === 'path' && !!i.stroke)
    // the first row is two lines of 12 px: its bottom border at 24, the second row's at 36
    expect(borders.map((b) => (b.d[0] as { y: number }).y)).toEqual([24, 24, 36, 36])
  })

  test('charts and SVG pictures are drawn as shapes and text', async () => {
    const items = await svgItems(
      '<svg xmlns="http://www.w3.org/2000/svg" width="100" height="50" viewBox="0 0 200 100"><defs><linearGradient id="g"><stop offset="0" stop-color="#f00"/><stop offset="1" stop-color="#00f"/></linearGradient><clipPath id="c"><rect width="50" height="50"/></clipPath></defs><rect width="200" height="100" fill="url(#g)"/><g clip-path="url(#c)" transform="translate(10 0)"><circle cx="20" cy="20" r="10" fill="#0f0" stroke="#000"/></g><text x="100" y="80" font-family="Test Em" font-size="20" text-anchor="middle">ab</text><script>alert(1)</script></svg>',
      book(),
    )
    const top = items[0] as Extract<DrawItem, { t: 'group' }>
    // the view box fitted to 100 × 50
    expect(top.transform).toEqual([0.5, 0, 0, 0.5, 0, 0])
    const all = flat(items)
    expect(all.map((i) => i.t)).toEqual(['path', 'path', 'text'])
    // objectBoundingBox: the gradient across the rectangle
    expect((all[0] as Extract<DrawItem, { t: 'path' }>).fill).toMatchObject({ t: 'linear', x0: 0, x1: 200 })
    const g = top.items[1] as Extract<DrawItem, { t: 'group' }>
    expect(g.transform).toEqual([1, 0, 0, 1, 10, 0])
    expect(g.clip?.length).toBe(5)
    // the text centred on its x: 40 px wide
    const t = top.items[2] as Extract<DrawItem, { t: 'group' }>
    expect(t.transform).toEqual([1, 0, 0, 1, 80, 80])
  })

  test('the sample deck as a PDF: every slide a page, its title in the outline, its chart drawn', async () => {
    const warnings: string[] = []
    const engine = new Engine((e, d) => e === 'warning' && warnings.push((d as { text: string }).text))
    await engine.open(new Uint8Array(readFileSync('playground/public/samples/sample.pptx')), { name: 'sample.pptx' })
    await engine.call('fonts.set', [[{ family: 'Calibri', src: EM }]])
    const blob = (await engine.call('pdf', [{ title: 'Sample' }])) as Blob
    expect(blob.type).toBe('application/pdf')
    const { text, objects } = await readPdf(new Uint8Array(await blob.arrayBuffer()))
    expect(text).toContain('/Count 5')
    expect(text).toContain('/Title (Shapes & images)')
    expect(text).toContain('/Title (Chart \\(placeholder\\))')
    expect(text).toContain('/Title (Sample)')
    // the chart's labels are text
    const contents = [...objects.values()].map((o) => (o.stream && o.dict.length < 40 ? latin1(o.stream) : '')).join('\n')
    expect(contents.split(' TJ').length).toBeGreaterThan(50)
    expect(warnings.every((w) => !/chart/i.test(w))).toBe(true)
    // one slide, and as SVG
    const one = (await engine.call('pdf', [{ slides: [2] }])) as Blob
    expect(latin1(new Uint8Array(await one.arrayBuffer()))).toContain('/Count 1')
    const svg = (await engine.call('pptx.svg', [0])) as string
    expect(svg).toMatch(/^<svg xmlns="http:\/\/www.w3.org\/2000\/svg" width="1279.97" height="720"/)
    expect(svg).toContain('office-engine')
  })

  test('hidden slides are left out unless asked for, also when named', async () => {
    const engine = new Engine(() => {})
    await engine.open(new Uint8Array(readFileSync('playground/public/samples/sample.pptx')), { name: 'sample.pptx' })
    await engine.call('pptx.slides', [{ op: 'hide', index: 1, hidden: true }])
    const count = async (opts: object) => /\/Count (\d+)/.exec(latin1(new Uint8Array(await ((await engine.call('pdf', [opts])) as Blob).arrayBuffer())))![1]
    expect(await count({})).toBe('4')
    expect(await count({ hidden: true })).toBe('5')
    // named, a hidden slide is still left out: a list of slides does not show what was hidden
    expect(await count({ slides: [0, 1] })).toBe('1')
    expect(await count({ slides: [1], hidden: true })).toBe('1')
    expect(engine.call('pdf', [{ slides: [1] }])).rejects.toThrow('Every slide asked for is hidden')
  })

  test('a page as SVG', () => {
    const svg = pageToSvg({ w: 10, h: 10, items: [{ t: 'path', d: rectPath(0, 0, 5, 5), fill: { t: 'color', color: { r: 255, g: 0, b: 0, a: 0.5 } } }] })
    expect(svg).toBe('<svg xmlns="http://www.w3.org/2000/svg" width="10" height="10" viewBox="0 0 10 10"><path d="M0 0L5 0L5 5L0 5Z" fill="rgb(255,0,0)" fill-opacity="0.5"/></svg>')
  })
})

describe('review: hostile input, copied text, licenses', () => {
  const style = (over: Partial<RunStyle> = {}): RunStyle => ({ font: 'Test Em', size: 10, color: BLACK, ...over })
  const para = (runs: { text: string; style: RunStyle }[]): ParaSpec => ({ runs, align: 'left', marL: 0, indent: 0, before: 0, after: 0, emptySize: 10 })
  const texts = (items: DrawItem[]) => items.filter((i): i is Extract<DrawItem, { t: 'text' }> => i.t === 'text')
  const contents = (objects: Map<number, { dict: string; stream?: Uint8Array }>) => [...objects.values()].map((o) => (o.stream && !o.dict.includes('/Length1') && !o.dict.includes('/Subtype/') ? latin1(o.stream) : '')).join('\n')

  test('an SVG whose uses repeat uses is refused, not multiplied; a use of itself is drawn once', async () => {
    let svg = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 10 10"><defs><rect id="a0" width="1" height="1"/>'
    for (let i = 1; i <= 20; i++) svg += `<g id="a${i}"><use href="#a${i - 1}"/><use href="#a${i - 1}"/></g>`
    svg += '</defs><use href="#a20"/></svg>'
    expect(svgItems(svg, book())).rejects.toThrow(/is not drawn/)
    const loop = await svgItems('<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 10 10"><g id="g"><rect width="1" height="1"/><use href="#g"/></g></svg>', book())
    const flat = (items: DrawItem[]): DrawItem[] => items.flatMap((i) => (i.t === 'group' ? flat(i.items) : [i]))
    expect(flat(loop).length).toBe(2)
  })

  test('a PNG is unpacked no further than a little past its pixels; a GIF frame past the limit is refused', async () => {
    const chunk = (t: string, d: number[] | Uint8Array) => [d.length >>> 24, (d.length >> 16) & 255, (d.length >> 8) & 255, d.length & 255, ...[...t].map((c) => c.charCodeAt(0)), ...d, 0, 0, 0, 0]
    // 16 MB of zeros behind a 1 × 1 picture
    const raw = await deflate(new Uint8Array(16 * 1024 * 1024))
    const png = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 13, 10, 26, 10, ...chunk('IHDR', [0, 0, 0, 1, 0, 0, 0, 1, 8, 0, 0, 0, 0]), ...chunk('IDAT', raw), ...chunk('IEND', [])])
    expect(decodePng(png)).rejects.toThrow(/much longer than its pixels/)
    const gif = Uint8Array.from(atob('R0lGODlhAQABAPAAAP8AAP///yH5BAAAAAAALAAAAAABAAEAAAICRAEAOw=='), (c) => c.charCodeAt(0))
    // the frame (its descriptor at byte 27) says 9000 × 9000
    gif.set([0x28, 0x23, 0x28, 0x23], 32)
    expect(() => decodeGif(gif)).toThrow(/not drawn/)
  })

  test('16-bit color keys compare all their bits', async () => {
    const chunk = (t: string, d: number[] | Uint8Array) => [d.length >>> 24, (d.length >> 16) & 255, (d.length >> 8) & 255, d.length & 255, ...[...t].map((c) => c.charCodeAt(0)), ...d, 0, 0, 0, 0]
    const raw = await deflate(new Uint8Array([0, 0xab, 0x01, 0xab, 0x02]))
    const png = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 13, 10, 26, 10, ...chunk('IHDR', [0, 0, 0, 2, 0, 0, 0, 1, 16, 0, 0, 0, 0]), ...chunk('tRNS', [0xab, 0x01]), ...chunk('IDAT', raw), ...chunk('IEND', [])])
    const px = await decodePng(png)
    expect([...px.color]).toEqual([0xab, 0xab])
    expect([...px.alpha!]).toEqual([0, 255])
  })

  test('a line set in two fonts: each glyph run stands for its own characters', async () => {
    const greek = buildFont({ family: 'Test Greek', advances: [500, 600], chars: [[0x3a9, 1]] })
    const b = new FontBook([{ family: 'Test Em', src: EM }, { family: 'Test Greek', src: greek }])
    const out = await layoutParagraphs(b, [para([{ text: 'AΩB', style: style({ font: 'Test Em, Test Greek' }) }])], 500)
    expect(texts(out.items).map((t) => [t.text, t.clusters])).toEqual([
      ['A', [0]],
      ['Ω', [0]],
      ['B', [0]],
    ])
  })

  test('lines are not broken inside a character of two UTF-16 units; tabs draw nothing', async () => {
    const out = await layoutParagraphs(book(), [para([{ text: 'a\u{20000}\u{20000}\u{20000}', style: style({ font: 'Nothing' }) }])], 1)
    for (const t of texts(out.items)) expect(t.text).not.toMatch(/^[\udc00-\udfff]|[\ud800-\udbff]$/)
    const tab = await layoutParagraphs(book(), [para([{ text: 'A\tB', style: style({ font: 'Nothing' }) }])], 500)
    const t = texts(tab.items)[0]
    expect(t.glyphs).toEqual([0x41, 0x42])
    // B at the first tab stop
    expect(t.advances[0]).toBe(96)
  })

  test('a narrow box of long text is laid out in linear time', async () => {
    const t0 = performance.now()
    const out = await layoutParagraphs(book(), [para([{ text: 'a'.repeat(20000), style: style() }])], 1)
    expect(texts(out.items).length).toBe(20000)
    expect(performance.now() - t0).toBeLessThan(3000)
  })

  test('the license: bitmaps only is not embedded, no subsetting embeds the whole font, a damaged font falls back', async () => {
    const font = (fsType: number, family: string) => buildFont({ family, fsType, advances: [500, 600, 700, 800], chars: [[0x41, 1], [0x42, 2], [0x43, 3]] })
    const bitmaps = await Font.open(font(0x0200, 'Test Bitmaps'))
    const whole = await Font.open(font(0x0100, 'Test Whole'))
    const damaged = await Font.open(font(0, 'Test Damaged'))
    // its first glyph made to end past its table
    const loca = damaged.tables.get('loca')!
    new DataView(loca.buffer, loca.byteOffset).setUint32(8, 0x7fffffff)
    const text = (f: Font): DrawItem => ({ t: 'text', x: 0, y: 20, size: 10, font: f, glyphs: [1], advances: [6], text: 'A', clusters: [0], fill: BLACK })
    const warnings: string[] = []
    const pdf = await writePdf([{ w: 100, h: 100, items: [text(bitmaps), text(whole), text(damaged)] }], { onWarning: (w) => warnings.push(w) })
    expect(warnings).toEqual([
      'The font Test Bitmaps may not be embedded but as bitmaps: its text is set in Helvetica',
      expect.stringMatching(/^The font Test Damaged could not be embedded \(.*\): its text is set in Helvetica$/),
    ])
    const { objects } = await readPdf(pdf)
    const file = [...objects.values()].find((o) => o.dict.includes('/Length1'))!
    expect((await Font.open(file.stream!)).numGlyphs).toBe(4)
  })

  test('a glyph that stands for different text in two places is given its text; links are ASCII', async () => {
    const font = await Font.open(EM)
    const g = font.glyphOf(0x66)
    const item = (text: string): DrawItem => ({ t: 'text', x: 0, y: 20, size: 10, font, glyphs: [g], advances: [10], text, clusters: [0], fill: BLACK })
    const pdf = await writePdf([{ w: 100, h: 100, items: [item('fi'), item('ﬁ'), { t: 'link', x: 0, y: 0, w: 10, h: 10, url: 'https://example.org/café(x)' }] }])
    const { text, objects } = await readPdf(pdf)
    const c = contents(objects)
    expect(c.match(/ActualText/g)?.length).toBe(1)
    expect(c).toContain('/Span<</ActualText <FEFFfb01>>>BDC')
    expect(text).toContain('/URI(https://example.org/caf%C3%A9\\(x\\))')
  })

  test('transparent groups are drawn whole; a gradient that fades has a soft mask; numbers have no exponent; hard stops', async () => {
    const red = { r: 255, g: 0, b: 0, a: 1 }
    const pdf = await writePdf([
      {
        w: 100,
        h: 100,
        items: [
          { t: 'group', opacity: 0.5, items: [{ t: 'path', d: rectPath(0, 0, 50, 50), fill: { t: 'color', color: red } }, { t: 'path', d: rectPath(25, 25, 50, 50), fill: { t: 'color', color: { ...red, a: 0.8 } } }] },
          { t: 'path', d: rectPath(0, 0, 1e22, 10), fill: { t: 'linear', x0: 0, y0: 0, x1: 100, y1: 0, stops: [{ at: 0, color: { ...red, a: 0 } }, { at: 0.5, color: red }, { at: 0.5, color: BLACK }, { at: 1, color: BLACK }] } },
        ],
      },
    ])
    const { text, objects } = await readPdf(pdf)
    expect(text).toContain('/Subtype/Form')
    expect(text).toContain('/Group<</S/Transparency/I true/CS/DeviceRGB>>')
    expect(text).toContain('/SMask<</Type/Mask/S/Luminosity/G')
    expect(text).not.toMatch(/\de[+-]?\d/)
    expect(contents(objects)).toContain('1000000000 0 l')
    // the hard stop at 0.5: two intervals, one bound
    expect(text).toContain('/Bounds[0.5]/Encode[0 1 0 1]')
  })

  test('PostScript subsets keyed by CID say Adobe-Identity-0, as the PDF does', async () => {
    const cid = await Font.open(buildFont({ family: 'Test Cid', cff: 'cid', advances: [500, 600, 700], chars: [[0x41, 1], [0x42, 2]] }))
    const sub = subsetFont(cid, [2])
    expect(latin1(sub.cff!)).toContain('Identity')
  })

  test('a slide\'s picture that cannot be read is left out; an empty selection is refused', async () => {
    const warnings: string[] = []
    const page = await slidePage({ index: 0, w: 100, h: 100, els: [{ t: 'pic', box: { x: 0, y: 0, w: 10, h: 10 }, target: 'ppt/media/a.png' }] }, { fonts: book(), part: async () => Promise.reject(new Error('gone')), onWarning: (w) => warnings.push(w) })
    expect(page.w).toBe(100)
    expect(warnings).toEqual(['A picture could not be read: gone'])
    const engine = new Engine(() => {})
    await engine.open(new Uint8Array(readFileSync('playground/public/samples/sample.pptx')), { name: 'sample.pptx' })
    expect(engine.call('pdf', [{ slides: [] }])).rejects.toThrow('No slides were asked for')
    expect(writePdf([])).rejects.toThrow(/at least one page/)
  })

  test('a cropped, flipped picture shows its cropped part mirrored in the frame', async () => {
    const png = Uint8Array.from(atob('R0lGODlhAQABAPAAAP8AAP///yH5BAAAAAAALAAAAAABAAEAAAICRAEAOw=='), (c) => c.charCodeAt(0))
    const page = await slidePage({ index: 0, w: 100, h: 100, els: [{ t: 'pic', box: { x: 0, y: 0, w: 75, h: 10, flipH: true }, target: 'a.gif', crop: [0.25, 0, 0, 0] }] }, { fonts: book(), part: async () => png })
    const frame = (page.items[0] as Extract<DrawItem, { t: 'group' }>).items[1] as Extract<DrawItem, { t: 'group' }>
    const flip = (frame.items[0] as Extract<DrawItem, { t: 'group' }>).items[0] as Extract<DrawItem, { t: 'group' }>
    expect(flip.transform).toEqual([-1, 0, 0, 1, 75, 0])
    // the picture's 100 px with its first 25 cut off
    expect(flip.items[0]).toMatchObject({ t: 'image', x: -25, w: 100 })
  })
})

describe('re-check: budgets, names, boxes', () => {
  const style = (over: Partial<RunStyle> = {}): RunStyle => ({ font: 'Test Em', size: 10, color: BLACK, ...over })
  const para = (runs: { text: string; style: RunStyle }[]): ParaSpec => ({ runs, align: 'left', marL: 0, indent: 0, before: 0, after: 0, emptySize: 10 })
  const texts = (items: DrawItem[]) => items.filter((i): i is Extract<DrawItem, { t: 'text' }> => i.t === 'text')
  const red = { r: 255, g: 0, b: 0, a: 1 }

  test('clip paths count toward the segments an SVG may draw', async () => {
    let d = 'M0 0'
    for (let i = 0; i < 4000; i++) d += ` L${i % 10} ${i % 7}`
    let svg = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 10 10"><defs><clipPath id="c"><path d="${d}"/></clipPath><rect id="r" width="1" height="1" clip-path="url(#c)"/></defs>`
    for (let i = 0; i < 251; i++) svg += '<use href="#r"/>'
    expect(svgItems(svg + '</svg>', book())).rejects.toThrow(/path segments is not drawn/)
  })

  test('a picture part that throws, or rejects with anything, is left out', async () => {
    for (const part of [() => { throw new Error('sync') }, () => Promise.reject(null)]) {
      const warnings: string[] = []
      await slidePage({ index: 0, w: 10, h: 10, els: [{ t: 'pic', box: { x: 0, y: 0, w: 5, h: 5 }, target: 'a.png' }] }, { fonts: book(), part: part as () => Promise<Uint8Array | null>, onWarning: (w) => warnings.push(w) })
      expect(warnings.length).toBe(1)
    }
  })

  test('a link with half a surrogate pair is written, as U+FFFD', async () => {
    const pdf = await writePdf([{ w: 10, h: 10, items: [{ t: 'link', x: 0, y: 0, w: 5, h: 5, url: 'https://example.org/\ud800x' }] }])
    expect(latin1(pdf)).toContain('/URI(https://example.org/%EF%BF%BDx)')
  })

  test('long runs and runs in two fonts stay linear', async () => {
    const font = await Font.open(EM)
    const n = 40000
    const g = font.glyphOf(0x41)
    const t0 = performance.now()
    await writePdf([{ w: 10, h: 10, items: [{ t: 'text', x: 0, y: 10, size: 10, font, glyphs: Array(n).fill(g), advances: Array(n).fill(10), text: 'A'.repeat(n), clusters: Array.from({ length: n }, (_, i) => i), fill: BLACK }] }])
    expect(performance.now() - t0).toBeLessThan(2000)
    const greek = buildFont({ family: 'Test Greek', advances: [500, 600], chars: [[0x3a9, 1]] })
    const b = new FontBook([{ family: 'Test Em', src: EM }, { family: 'Test Greek', src: greek }])
    const t1 = performance.now()
    const out = await layoutParagraphs(b, [para([{ text: 'AΩ'.repeat(6000), style: style({ font: 'Test Em, Test Greek' }) }])], 1)
    expect(texts(out.items).length).toBe(12000)
    expect(performance.now() - t1).toBeLessThan(3000)
  })

  test('masks have names of their own; a form\'s box holds the page wherever it is drawn; whole fonts carry no subset tag', async () => {
    const whole = await Font.open(buildFont({ family: 'Test Whole', fsType: 0x0100, advances: [500, 600], chars: [[0x41, 1]] }))
    const items: DrawItem[] = [
      { t: 'path', d: rectPath(0, 0, 10, 10), fill: { t: 'linear', x0: 0, y0: 0, x1: 10, y1: 0, stops: [{ at: 0, color: { ...red, a: 0 } }, { at: 1, color: red }] } },
      ...Array.from({ length: 5 }, (_, i): DrawItem => ({ t: 'path', d: rectPath(0, 0, 1, 1), fill: { t: 'color', color: { ...red, a: (i + 1) / 10 } } })),
      { t: 'group', transform: [1, 0, 0, 1, -1000010, 0], items: [{ t: 'group', opacity: 0.5, items: [{ t: 'path', d: rectPath(1000010, 0, 50, 50), fill: { t: 'color', color: red } }] }] },
      { t: 'text', x: 0, y: 20, size: 10, font: whole, glyphs: [1], advances: [6], text: 'A', clusters: [0], fill: BLACK },
    ]
    const { text } = await readPdf(await writePdf([{ w: 100, h: 100, items }]))
    const names = [...text.matchAll(/\/ExtGState<<([^>]*)>>/g)].flatMap((m) => [...m[1].matchAll(/\/(\w+) \d+ 0 R/g)].map((x) => x[1]))
    expect(new Set(names).size).toBe(names.length)
    expect(names).toContain('M1')
    // the transparent group's form: in its space the page is at x 1000010 (and the box around it)
    const box = /\/BBox\[([^\]]*)\]\/Group<<\/S\/Transparency\/I true/.exec(text)![1].split(' ').map(Number)
    expect(box[0]).toBeLessThanOrEqual(1000010)
    expect(box[2]).toBeGreaterThanOrEqual(1000110)
    expect(text).toContain('/BaseFont/TestWhole')
  })

  test('a gradient along a line is a shading pattern; stops that round together make one edge', async () => {
    const stops = [0, 0.0002, 0.5001, 0.5002, 0.9998, 1].map((at, i) => ({ at, color: i % 2 ? red : BLACK }))
    const { text, objects } = await readPdf(await writePdf([{ w: 100, h: 100, items: [{ t: 'path', d: [{ k: 'M', x: 0, y: 0 }, { k: 'L', x: 100, y: 0 }], stroke: { paint: { t: 'linear', x0: 0, y0: 0, x1: 100, y1: 0, stops }, width: 4 } }] }]))
    expect(text).toContain('/PatternType 2')
    expect([...objects.values()].map((o) => (o.stream ? latin1(o.stream) : '')).join('')).toContain('/Pattern CS /P1 SCN')
    const bounds = /\/Bounds\[([^\]]*)\]/.exec(text)![1].split(' ').map(Number)
    for (let i = 1; i < bounds.length; i++) expect(bounds[i]).toBeGreaterThan(bounds[i - 1])
  })

  test('a PNG whose data is cut short of its checksum is refused', async () => {
    const chunk = (t: string, d: number[] | Uint8Array) => [d.length >>> 24, (d.length >> 16) & 255, (d.length >> 8) & 255, d.length & 255, ...[...t].map((c) => c.charCodeAt(0)), ...d, 0, 0, 0, 0]
    const raw = await deflate(new Uint8Array([0, 7]))
    const png = (data: Uint8Array) => new Uint8Array([0x89, 0x50, 0x4e, 0x47, 13, 10, 26, 10, ...chunk('IHDR', [0, 0, 0, 1, 0, 0, 0, 1, 8, 0, 0, 0, 0]), ...chunk('IDAT', data), ...chunk('IEND', [])])
    expect([...(await decodePng(png(raw))).color]).toEqual([7])
    expect(decodePng(png(raw.subarray(0, raw.length - 4)))).rejects.toThrow(/damaged/)
  })

  test('a line can end after an ideograph of two UTF-16 units', async () => {
    // x 10 px, the ideograph in Helvetica (a question mark, 5.56 px), y and z 10 px: x and the ideograph fit in 31 px with y
    const out = await layoutParagraphs(book(), [para([{ text: 'x\u{20000}yz', style: style() }])], 21)
    const t = texts(out.items)
    expect(t.map((i) => i.text)).toEqual(['x', '\u{20000}', 'yz'])
    expect(t[1].y).toBe(t[0].y)
    expect(t[2].y).toBeGreaterThan(t[0].y)
  })
})

describe('final check: precision, reuse, edges', () => {
  const red = { r: 255, g: 0, b: 0, a: 1 }
  const blue = { r: 0, g: 0, b: 255, a: 1 }

  test('small scales are written as they are, so boxes and patterns agree with what is drawn', async () => {
    const pdf = await writePdf([
      {
        w: 100,
        h: 100,
        items: [
          { t: 'group', transform: [0.00149, 0, 0, 1, 0, 0], items: [{ t: 'group', opacity: 0.5, items: [{ t: 'path', d: rectPath(60000, 0, 50000, 50), fill: { t: 'color', color: red } }] }] },
          { t: 'group', opacity: 0.5, items: [{ t: 'group', transform: [0.02, 0, 0, 1, 0, 0], items: [{ t: 'group', transform: [0.02, 0, 0, 1, 0, 0], items: [{ t: 'path', d: [{ k: 'M', x: 0, y: 10 }, { k: 'L', x: 200000, y: 10 }], stroke: { paint: { t: 'linear', x0: 0, y0: 0, x1: 200000, y1: 0, stops: [{ at: 0, color: red }, { at: 1, color: blue }] }, width: 2 } }] }] }] },
        ],
      },
    ])
    const { text, objects } = await readPdf(pdf)
    const all = [...objects.values()].map((o) => (o.stream ? latin1(o.stream) : '')).join('\n')
    expect(all).toContain('0.00149 0 0 1 0 0 cm')
    // the form's box, in its space: the page's 75 points are 75 / (0.75 × 0.00149) wide there
    const box = /\/BBox\[([^\]]*)\]/.exec(text)![1].split(' ').map(Number)
    expect(box[2]).toBeCloseTo(100 / 0.00149 + 1, -1)
    // the pattern keeps the product of the two scales, in the space of the form it is drawn in
    expect(text).toMatch(/\/PatternType 2\/Shading \d+ 0 R\/Matrix\[0\.0004 0 0 1 0 0\]/)
  })

  test('a text command changed after one export is exported as it is now', async () => {
    const font = await Font.open(EM)
    const item: Extract<DrawItem, { t: 'text' }> = { t: 'text', x: 0, y: 20, size: 10, font, glyphs: [font.glyphOf(0x41)], advances: [10], text: 'A', clusters: [0], fill: BLACK }
    expect(pageToSvg({ w: 10, h: 10, items: [item] })).toContain('>A</text>')
    item.text = 'B'
    item.glyphs = [font.glyphOf(0x42)]
    expect(pageToSvg({ w: 10, h: 10, items: [item] })).toContain('>B</text>')
  })

  test('a gradient of no length paints its last color; a color fill and outline are painted together', async () => {
    const { objects } = await readPdf(
      await writePdf([
        {
          w: 100,
          h: 100,
          items: [
            { t: 'path', d: [{ k: 'M', x: 0, y: 5 }, { k: 'L', x: 50, y: 5 }], stroke: { paint: { t: 'linear', x0: 20, y0: 5, x1: 20, y1: 5, stops: [{ at: 0, color: red }, { at: 1, color: blue }] }, width: 2 } },
            { t: 'path', d: rectPath(10, 10, 50, 50), fill: { t: 'color', color: { ...red, a: 0.5 } }, stroke: { paint: { t: 'color', color: { ...blue, a: 0.5 } }, width: 6 } },
          ],
        },
      ]),
    )
    const c = [...objects.values()].map((o) => (o.stream ? latin1(o.stream) : '')).join('\n')
    expect(c).toContain('0 0 1 RG')
    expect(c).not.toContain('/Pattern CS')
    expect(c).toMatch(/1 0 0 rg\n0 0 1 RG\n6 w\n0 J\n0 j\n10 10 m .* B\n/)
  })

  test('two fonts of one name keep apart', async () => {
    const whole = (name: string) => buildFont({ family: name, fsType: 0x0100, advances: [500, 600], chars: [[0x41, 1]] })
    const a = await Font.open(whole('Test.A_B'))
    const b = await Font.open(whole('TestAB'))
    const t = (f: Font): DrawItem => ({ t: 'text', x: 0, y: 20, size: 10, font: f, glyphs: [1], advances: [6], text: 'A', clusters: [0], fill: BLACK })
    const text = latin1(await writePdf([{ w: 10, h: 10, items: [t(a), t(b)] }]))
    const names = [...text.matchAll(/\/FontName(\/[^/]+)\//g)].map((m) => m[1])
    expect(new Set(names).size).toBe(2)
  })

  test('a rejection that cannot be made a string is still reported', async () => {
    const warnings: string[] = []
    await slidePage({ index: 0, w: 10, h: 10, els: [{ t: 'pic', box: { x: 0, y: 0, w: 5, h: 5 }, target: 'a.png' }] }, { fonts: book(), part: () => Promise.reject(Object.create(null)), onWarning: (w) => warnings.push(w) })
    expect(warnings).toEqual(['A picture could not be read: an error'])
  })
})
