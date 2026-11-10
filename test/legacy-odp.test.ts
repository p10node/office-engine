import { describe, expect, test } from 'bun:test'
import { readFileSync } from 'node:fs'
import type { DocComment } from '../src/core/comments'
import { convertFlatOdf } from '../src/core/legacy'
import type { SlideEl, SlideModel } from '../src/core/pptx/types'
import { BlobSource } from '../src/core/source'
import { Engine } from '../src/core/worker/engine'
import { ZipArchive } from '../src/core/zip'
import { crc32, writeZip } from '../src/core/zipwrite'

const fixture = (name: string) => new Uint8Array(readFileSync(new URL(`./fixtures/legacy/${name}`, import.meta.url)))

async function open(src: Uint8Array | string, name = 'deck.fodp') {
  const engine = new Engine(() => {})
  const info = await engine.open(new Blob([src as BlobPart]), { name })
  const slides: SlideModel[] = []
  for (let i = 0; i < (info.pptx?.count ?? 0); i++) slides.push((await engine.call('pptx.slide', [i])) as SlideModel)
  /** a part of the converted package */
  const part = async (path: string) => ((await engine.call('part', [path])) as Blob | undefined)?.text()
  return { engine, info, slides, part }
}

type Shape = Extract<SlideEl, { t: 'shape' }>
const shapes = (s: SlideModel) => s.els.filter((e): e is Shape => e.t === 'shape')
const textOf = (e: SlideEl | undefined) => (e && e.t === 'shape' ? (e.text?.paras.map((p) => p.runs.map((r) => (r.br ? '\n' : r.text)).join('')).join('|') ?? '') : '')
const byText = (s: SlideModel, t: string) => shapes(s).find((e) => textOf(e).includes(t))!
/** cm → slide px (96 dpi) */
const px = (cm: number) => (cm / 2.54) * 96
const near = (a: number, b: number, tol = 0.6) => expect(Math.abs(a - b)).toBeLessThan(tol)

// ---------------------------------------------------------------- a flat OpenDocument presentation builder

const NS = [
  'office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"',
  'style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"',
  'text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"',
  'table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"',
  'draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"',
  'fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"',
  'svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"',
  'presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0"',
  'xlink="http://www.w3.org/1999/xlink"',
  'loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0"',
]
  .map((n) => 'xmlns:' + n)
  .join(' ')

const STYLES =
  '<style:default-style style:family="graphic"><style:graphic-properties svg:stroke-color="#3465a4" draw:fill-color="#729fcf"/><style:text-properties style:font-name="Liberation Sans" fo:font-size="18pt"/></style:default-style>' +
  '<style:style style:name="standard" style:family="graphic"><style:graphic-properties draw:stroke="solid" svg:stroke-color="#3465a4" draw:fill="solid" draw:fill-color="#729fcf"/></style:style>' +
  '<style:style style:name="Default-title" style:family="presentation"><style:graphic-properties draw:stroke="none" draw:fill="none" draw:textarea-vertical-align="middle"/><style:paragraph-properties fo:text-align="center"/><style:text-properties fo:font-size="44pt"/></style:style>' +
  '<style:style style:name="Default-outline1" style:family="presentation"><style:graphic-properties draw:stroke="none" draw:fill="none"><text:list-style style:name="Default-outline1">' +
  '<text:list-level-style-bullet text:level="1" text:bullet-char="●"><style:list-level-properties text:space-before="0.3cm" text:min-label-width="0.9cm"/><style:text-properties fo:font-family="OpenSymbol" fo:font-size="45%"/></text:list-level-style-bullet>' +
  '<text:list-level-style-bullet text:level="2" text:bullet-char="–"><style:list-level-properties text:space-before="1.5cm" text:min-label-width="0.9cm"/><style:text-properties fo:font-family="OpenSymbol" fo:font-size="75%"/></text:list-level-style-bullet>' +
  '</text:list-style></style:graphic-properties><style:paragraph-properties fo:margin-top="0.5cm"/><style:text-properties fo:font-size="32pt"/></style:style>' +
  '<style:style style:name="Default-outline2" style:family="presentation" style:parent-style-name="Default-outline1"><style:text-properties fo:font-size="28pt"/></style:style>' +
  '<style:style style:name="Default-notes" style:family="presentation"><style:graphic-properties draw:stroke="none" draw:fill="none"/><style:text-properties fo:font-size="20pt"/></style:style>'

/** 25.4 × 19.05 cm (4:3): 960 × 720 px */
const PAGE = '<style:page-layout style:name="PM1"><style:page-layout-properties fo:page-width="25.4cm" fo:page-height="19.05cm"/></style:page-layout>'

const MASTER =
  '<style:master-page style:name="Default" style:page-layout-name="PM1" draw:style-name="Mdp1">' +
  '<draw:frame presentation:style-name="Default-title" svg:x="1cm" svg:y="1cm" svg:width="23cm" svg:height="3cm" presentation:class="title" presentation:placeholder="true"><draw:text-box/></draw:frame>' +
  '<draw:frame presentation:style-name="Default-outline1" svg:x="1cm" svg:y="5cm" svg:width="23cm" svg:height="12cm" presentation:class="outline" presentation:placeholder="true"><draw:text-box/></draw:frame>' +
  '</style:master-page>'

function fodp(o: { body: string; styles?: string; auto?: string; masters?: string; office?: string }): string {
  return (
    `<?xml version="1.0" encoding="UTF-8"?>\n<office:document ${NS} office:version="1.3" office:mimetype="application/vnd.oasis.opendocument.presentation">` +
    `<office:font-face-decls><style:font-face style:name="Liberation Sans" svg:font-family="'Liberation Sans'"/></office:font-face-decls>` +
    `<office:styles>${STYLES}${o.styles ?? ''}</office:styles>` +
    `<office:automatic-styles>${PAGE}<style:style style:name="Mdp1" style:family="drawing-page"><style:drawing-page-properties draw:fill="none"/></style:style>${o.auto ?? ''}</office:automatic-styles>` +
    `<office:master-styles>${o.masters ?? MASTER}</office:master-styles><office:body><office:presentation>${o.body}</office:presentation></office:body></office:document>`
  )
}
const page = (inner: string, attrs = '') => `<draw:page draw:master-page-name="Default"${attrs}>${inner}</draw:page>`

/** a PNG of w × h pixels (one color), 96 dpi */
function png(w: number, h: number): Uint8Array {
  const chunk = (type: string, data: Uint8Array) => {
    const out = new Uint8Array(12 + data.length)
    const dv = new DataView(out.buffer)
    dv.setUint32(0, data.length)
    for (let i = 0; i < 4; i++) out[4 + i] = type.charCodeAt(i)
    out.set(data, 8)
    dv.setUint32(8 + data.length, crc32(out.subarray(4, 8 + data.length)))
    return out
  }
  const ihdr = new Uint8Array(13)
  const dv = new DataView(ihdr.buffer)
  dv.setUint32(0, w)
  dv.setUint32(4, h)
  ihdr.set([8, 2, 0, 0, 0], 8)
  const raw = new Uint8Array(h * (1 + w * 3))
  // zlib stream of stored (uncompressed) deflate blocks
  const z = [0x78, 0x01, 0x01, raw.length & 0xff, raw.length >> 8, ~raw.length & 0xff, (~raw.length >> 8) & 0xff, ...raw]
  let a = 1
  let b = 0
  for (const v of raw) {
    a = (a + v) % 65521
    b = (b + a) % 65521
  }
  z.push(b >> 8, b & 0xff, a >> 8, a & 0xff)
  const parts = [Uint8Array.of(0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a), chunk('IHDR', ihdr), chunk('IDAT', Uint8Array.from(z)), chunk('IEND', new Uint8Array(0))]
  const out = new Uint8Array(parts.reduce((n, p) => n + p.length, 0))
  let o = 0
  for (const p of parts) out.set(p, o), (o += p.length)
  return out
}
const base64 = (b: Uint8Array) => btoa(String.fromCharCode(...b))

// ---------------------------------------------------------------- the sample deck

describe('OpenDocument presentation fixtures', () => {
  test('sample.odp: masters, placeholders, text formatting, lists, shapes, picture, table, backgrounds, footer, notes', async () => {
    const { info, slides, part } = await open(fixture('sample.odp'), 'sample.odp')
    expect(info.kind).toBe('pptx')
    expect(info.converted).toBe('odp')
    // 28 × 15.75 cm
    expect(info.pptx).toMatchObject({ count: 6 })
    near(info.pptx!.w, px(28))
    near(info.pptx!.h, px(15.75))
    expect(slides.map((s) => !!s.hidden)).toEqual([false, false, false, false, false, true])

    // the master's band (a gradient) under every slide, its footer and slide number where the page shows them
    for (const s of slides) expect(shapes(s)[0]).toMatchObject({ geom: 'rect', fill: { t: 'grad' } })
    near(shapes(slides[0])[0].box.y, px(14.9))
    expect(textOf(byText(slides[0], 'sample deck'))).toBe('office-engine sample deck')
    const number = shapes(slides[2]).find((e) => e.text?.paras[0]?.runs[0]?.fld)!
    expect(textOf(number)).toBe('3')
    expect(shapes(slides[5]).some((e) => textOf(e).includes('sample deck'))).toBe(false)

    // title slide: title and subtitle with their master styles
    const title = byText(slides[0], 'office-engine')
    expect(title.text!.paras[0]).toMatchObject({ algn: 'center', runs: [{ text: 'office-engine', sz: 40, b: true, color: '#111827' }] })
    expect(title.text!.anchor).toBe('ctr')
    near(title.box.x, px(2))
    expect(byText(slides[0], 'Progressive').text!.paras[0].runs[0]).toMatchObject({ sz: 24, color: '#4b5563' })

    // outline: bullets by level with the list's indents, run formatting, link, highlight, superscript
    const body = byText(slides[1], 'Lazy ZIP')
    const paras = body.text!.paras
    expect(paras.map((p) => [p.bullet?.text, p.lvl ?? 0])).toEqual([
      ['●', 0],
      ['–', 1],
      ['●', 0],
      ['●', 0],
    ])
    expect(paras[0].bullet).toMatchObject({ color: '#2563eb' })
    // text at space-before + min-label-width (1.2 cm), the bullet at space-before (0.3 cm)
    near(paras[0].marL, px(1.2))
    near(paras[0].marL + paras[0].indent, px(0.3))
    expect(paras[0].runs.map((r) => [r.text, !!r.b])).toEqual([
      ['Lazy ZIP: only the ', false],
      ['central directory', true],
      [' is read up front', false],
    ])
    expect(paras[0].runs[0]).toMatchObject({ sz: 28, color: '#1f2937' })
    expect(paras[0].runs[0].font).toContain('Arial')
    expect(paras[1].runs[1]).toMatchObject({ text: 'on demand', i: true, sz: 24 })
    expect(paras[2].runs[1]).toMatchObject({ text: 'Web Worker', u: true })
    const last = paras[3].runs
    expect(last.find((r) => r.text === 'documentation')).toMatchObject({ href: 'https://example.com/docs' })
    expect(last.find((r) => r.text === 'open()')).toMatchObject({ hl: '#fef3c7' })
    expect(last.find((r) => r.text === 'open()')!.font).toContain('Courier New')
    expect(last.find((r) => r.text === '2')).toMatchObject({ baseline: 33 })

    // shapes: preset, custom path with equations, rotated star, text box, picture, line with an arrow head
    const s3 = shapes(slides[2])
    expect(s3.map((e) => [e.geom, textOf(e)]).slice(2, 7)).toEqual([
      ['roundRect', 'Rounded'],
      ['ellipse', 'Oval'],
      ['custom', 'Arrow'],
      ['custom', ''],
      ['rect', 'A free text box with colored bold text'],
    ])
    expect(byText(slides[2], 'Rounded')).toMatchObject({ fill: { t: 'solid', c: '#2563eb' }, adj: { adj: 16667 } })
    expect(byText(slides[2], 'Rounded').text!.paras[0].runs[0]).toMatchObject({ color: '#ffffff', b: true, sz: 20 })
    expect(byText(slides[2], 'Oval').stroke).toMatchObject({ c: '#14532d' })
    // the arrow's path from its equations: the shaft at a quarter of the height
    expect(byText(slides[2], 'Arrow').path).toStartWith('M0.00 28.35 L141.73 28.35')
    const star = s3[5]
    expect(star.box.rot).toBe(330)
    expect(byText(slides[2], 'colored bold').text!.paras[0].runs[1]).toMatchObject({ text: 'colored bold text', color: '#dc2626', b: true })
    const pic = slides[2].els.find((e) => e.t === 'pic')!
    expect(pic.t === 'pic' && pic.target).toMatch(/^ppt\/media\/image\d\.png$/)
    near(pic.box.x, px(14))
    const line = s3.find((e) => e.geom === 'line')!
    expect(line.stroke).toMatchObject({ c: '#dc2626', tail: 'triangle' })

    // table: column widths, header and body fills, text colors
    const tbl = slides[3].els.find((e) => e.t === 'table')!
    if (tbl.t !== 'table') throw new Error('no table')
    expect(tbl.cols.map((c) => Math.round(c))).toEqual([227, 151, 151, 151])
    expect(tbl.rows.map((r) => r.cells.map((c) => textOf({ t: 'shape', box: tbl.box, geom: 'rect', text: c.text })))).toEqual([
      ['Metric', '2024', '2025', '2026'],
      ['Users', '380', '750', '424'],
      ['Revenue', '746', '55', '781'],
      ['NPS', '469', '441', '643'],
      ['Latency', '130', '209', '885'],
    ])
    expect(tbl.rows[0].cells[0]).toMatchObject({ fill: { t: 'solid', c: '#1d4ed8' } })
    expect(tbl.rows[0].cells[0].text!.paras[0].runs[0]).toMatchObject({ color: '#ffffff', b: true })
    expect(tbl.rows[1].cells[2]).toMatchObject({ fill: { t: 'solid', c: '#e5e7eb' } })

    // numbered list with a lettered sub-list on a tinted background
    expect(slides[4].bg).toEqual({ t: 'solid', c: '#f1f5f9' })
    expect(byText(slides[4], 'Install').text!.paras.map((p) => p.bullet?.text)).toEqual(['1.', '2.', '3.', 'a)', 'b)'])

    // speaker notes as notes slides
    expect(await part('ppt/notesSlides/notesSlide1.xml')).toContain('Welcome everyone and introduce the project.')
    expect(await part('ppt/notesSlides/notesSlide2.xml')).toContain('Mention the lazy ZIP reader first.')
    expect(await part('ppt/slides/_rels/slide1.xml.rels')).toContain('notesSlide1.xml')
    // master text styles: the outline levels with their bullets
    const master = (await part('ppt/slideMasters/slideMaster1.xml'))!
    expect(master).toContain('<p:bodyStyle><a:lvl1pPr')
    expect(master).toMatch(/<a:lvl2pPr[^>]*>.*<a:buChar char="–"\/>/)
  })

  test('the template (.otp) and the flat XML (.fodp) open the same, the picture inline', async () => {
    for (const name of ['sample.otp', 'sample.fodp']) {
      const { info, slides } = await open(fixture(name), name)
      expect(info.converted).toBe('odp')
      expect(info.pptx!.count).toBe(6)
      expect(textOf(byText(slides[0], 'office-engine'))).toBe('office-engine')
      expect(slides[2].els.some((e) => e.t === 'pic' && e.target)).toBe(true)
    }
  })

  test('saving gives a PowerPoint file that opens again', async () => {
    const { engine } = await open(fixture('sample.odp'), 'sample.odp')
    const pptx = (await engine.call('save', [])) as Blob
    const again = await open(new Uint8Array(await pptx.arrayBuffer()), 'sample.pptx')
    expect(again.info.converted).toBeUndefined()
    expect(again.info.pptx?.count).toBe(6)
    expect(textOf(byText(again.slides[1], 'Lazy ZIP'))).toContain('central directory')
  })
})

// ---------------------------------------------------------------- targeted documents

describe('OpenDocument presentation conversion', () => {
  test('page size from the master page layout; placeholders without a frame take the master one; layouts by slide kind', async () => {
    const xml = fodp({
      auto: '<style:style style:name="pr1" style:family="presentation" style:parent-style-name="Default-title"/>',
      body:
        page('<draw:frame presentation:style-name="pr1" presentation:class="title"><draw:text-box><text:p>Inherited</text:p></draw:text-box></draw:frame>') +
        page('<draw:frame presentation:style-name="pr1" presentation:class="title" svg:x="2cm" svg:y="2cm" svg:width="10cm" svg:height="2cm"><draw:text-box><text:p>Title</text:p></draw:text-box></draw:frame><draw:frame presentation:class="subtitle" svg:x="2cm" svg:y="6cm" svg:width="10cm" svg:height="2cm"><draw:text-box><text:p>Sub</text:p></draw:text-box></draw:frame>'),
    })
    const { info, slides, part } = await open(xml)
    expect(info.pptx).toMatchObject({ w: 960, h: 720, count: 2 })
    const t = byText(slides[0], 'Inherited')
    for (const [k, v] of Object.entries({ x: 1, y: 1, w: 23, h: 3 })) near(t.box[k as 'x'], px(v))
    expect(t.text!.paras[0].runs[0]).toMatchObject({ sz: 44 })
    expect(await part('ppt/slides/slide1.xml')).toContain('<p:ph type="title"/>')
    expect(await part('ppt/slides/slide2.xml')).toContain('<p:ph type="subTitle" idx="1"/>')
    const layout = async (n: number) => /slideLayouts\/(slideLayout\d+\.xml)/.exec((await part(`ppt/slides/_rels/slide${n}.xml.rels`))!)![1]
    expect(await part(`ppt/slideLayouts/${await layout(1)}`)).toContain('type="obj"')
    expect(await part(`ppt/slideLayouts/${await layout(2)}`)).toContain('type="title"')
  })

  test('shapes: transforms, mirrored custom shapes, arcs, polygons, paths, lines, groups, rounded rectangles, sections', async () => {
    const auto =
      '<style:style style:name="gr1" style:family="graphic" style:parent-style-name="standard"><style:graphic-properties draw:fill-color="#ff0000"/></style:style>' +
      '<style:style style:name="gr2" style:family="graphic" style:parent-style-name="standard"><style:graphic-properties draw:stroke="dash" draw:stroke-dash="Fine" svg:stroke-width="0.1cm" draw:marker-start="Circle" draw:marker-end="Arrow"/></style:style>'
    const styles = '<draw:stroke-dash draw:name="Fine" draw:style="rect" draw:dots1="1" draw:dots1-length="0.2cm" draw:distance="0.2cm"/>'
    const xml = fodp({
      styles,
      auto,
      body: page(
        // an ellipse turned a quarter counterclockwise around its corner, then moved
        '<draw:custom-shape draw:style-name="gr1" svg:width="4cm" svg:height="2cm" draw:transform="rotate (1.5707963267949) translate (5cm 10cm)"><draw:enhanced-geometry svg:viewBox="0 0 21600 21600" draw:type="ellipse"/></draw:custom-shape>' +
          // a half disc from an angle-ellipse; mirrored on one axis: the rotation turns the other way
          '<draw:custom-shape draw:style-name="gr1" svg:width="4cm" svg:height="4cm" draw:transform="rotate (0.785398163397448) translate (10cm 2cm)"><draw:enhanced-geometry svg:viewBox="0 0 21600 21600" draw:mirror-horizontal="true" draw:type="mso-spt100" draw:enhanced-path="U 10800 10800 10800 10800 0 180 Z N"/></draw:custom-shape>' +
          '<draw:polygon draw:style-name="gr1" svg:x="1cm" svg:y="1cm" svg:width="2cm" svg:height="2cm" svg:viewBox="0 0 1000 1000" draw:points="0,0 1000,0 500,1000"/>' +
          '<draw:polyline draw:style-name="gr1" svg:x="4cm" svg:y="1cm" svg:width="2cm" svg:height="2cm" svg:viewBox="0 0 1000 1000" draw:points="0,0 1000,1000"/>' +
          '<draw:path draw:style-name="gr1" svg:x="7cm" svg:y="1cm" svg:width="2cm" svg:height="2cm" svg:viewBox="0 0 1000 1000" svg:d="M 0 0 C 0 1000 1000 1000 1000 0 Z"/>' +
          '<draw:line draw:style-name="gr2" svg:x1="1cm" svg:y1="8cm" svg:x2="6cm" svg:y2="5cm"/>' +
          '<draw:g><draw:rect draw:style-name="gr1" svg:x="12cm" svg:y="12cm" svg:width="2cm" svg:height="1cm" draw:corner-radius="0.25cm"/><draw:ellipse draw:style-name="gr1" svg:x="15cm" svg:y="13cm" svg:width="2cm" svg:height="2cm" draw:kind="section" draw:start-angle="0" draw:end-angle="90"/></draw:g>',
      ),
    })
    const { slides, part } = await open(xml)
    const s = shapes(slides[0])
    expect(s.map((e) => e.geom)).toEqual(['ellipse', 'custom', 'custom', 'custom', 'custom', 'line', 'roundRect', 'custom'])
    // the box keeps its size and center; ODF turns counterclockwise, DrawingML clockwise
    const ell = s[0]
    expect(ell.box.rot).toBe(270)
    near(ell.box.w, px(4))
    near(ell.box.x + ell.box.w / 2, px(5 + 1))
    near(ell.box.y + ell.box.h / 2, px(10 - 2))
    expect(s[1].box.rot).toBe(45)
    expect(s[1].path).toContain('C')
    expect(s[2].path).toMatch(/^M0\.00 0\.00 L75\.59 0\.00 L37\.80 75\.59 Z/)
    expect(s[3].fill).toEqual({ t: 'none' })
    expect(s[4].path).toContain('C0.00 75.59 75.59 75.59 75.59 0.00')
    expect(s[5].box).toMatchObject({ x: px(1), y: px(5), flipV: true })
    expect(s[5].stroke).toMatchObject({ head: 'oval', tail: 'triangle' })
    expect(s[5].stroke!.dash).toBeDefined()
    expect(s[6].adj).toEqual({ adj: 25000 })
    near(s[7].box.x, px(15))
    expect(await part('ppt/slides/slide1.xml')).toContain('<p:grpSp>')
  })

  test('text: white space, spaces, tabs and breaks, fields, list headers, numbering start, fonts, decorations, automatic color', async () => {
    const auto =
      '<style:style style:name="dp1" style:family="drawing-page"><style:drawing-page-properties draw:fill="solid" draw:fill-color="#101010"/></style:style>' +
      '<style:style style:name="gr1" style:family="graphic"><style:graphic-properties draw:stroke="none" draw:fill="none"/></style:style>' +
      '<style:style style:name="T1" style:family="text"><style:text-properties style:text-underline-style="solid" style:text-underline-type="double" style:text-line-through-style="solid" fo:text-transform="uppercase"/></style:style>' +
      '<style:style style:name="T2" style:family="text"><style:text-properties fo:color="#00ff00" fo:font-family="\'DejaVu Serif\'" fo:font-size="150%"/></style:style>' +
      '<text:list-style style:name="L1"><text:list-level-style-number text:level="1" style:num-format="i" style:num-prefix="(" style:num-suffix=")" text:start-value="3"><style:list-level-properties text:space-before="0.5cm" text:min-label-width="1cm"/></text:list-level-style-number></text:list-style>'
    const box = (y: number, inner: string) => `<draw:frame draw:style-name="gr1" svg:x="1cm" svg:y="${y}cm" svg:width="20cm" svg:height="2cm"><draw:text-box>${inner}</draw:text-box></draw:frame>`
    const xml = fodp({
      auto,
      body:
        page(
          box(1, '<text:p>\n   Hello   <text:s text:c="2"/>world <text:tab/>tab<text:line-break/>  next  </text:p>') +
            box(4, '<text:p>Page <text:page-number>9</text:page-number> of <text:page-count>9</text:page-count></text:p>') +
            box(7, '<text:list text:style-name="L1"><text:list-header><text:p>Header</text:p></text:list-header><text:list-item><text:p>Third</text:p></text:list-item><text:list-item><text:p>Fourth</text:p></text:list-item></text:list>') +
            box(10, '<text:p><text:span text:style-name="T1">marked</text:span> <text:span text:style-name="T2">green</text:span> <text:a xlink:href="#page2">next slide</text:a></text:p>'),
          ' draw:style-name="dp1"',
        ) + page(box(1, '<text:p>Second</text:p>')),
    })
    const { slides } = await open(xml)
    const [ws, fields, list, deco] = shapes(slides[0])
    expect(ws.text!.paras[0].runs.map((r) => (r.br ? '<br>' : r.text)).join('')).toBe('Hello   world \ttab<br>next')
    expect(textOf(fields)).toBe('Page 1 of 2')
    expect(fields.text!.paras[0].runs[1]).toMatchObject({ fld: true })
    // the header has no label but sits at the text position; numbering in roman numerals from 3
    expect(list.text!.paras.map((p) => p.bullet?.text)).toEqual([undefined, '(iii)', '(iv)'])
    near(list.text!.paras[0].marL, px(1.5))
    expect(list.text!.paras[0].indent).toBe(0)
    near(list.text!.paras[1].marL + list.text!.paras[1].indent, px(0.5))
    const [marked, , green, , link] = deco.text!.paras[0].runs
    expect(marked).toMatchObject({ u: true, strike: true, caps: true, color: '#ffffff' })
    expect(green).toMatchObject({ color: '#00ff00', sz: 27 })
    expect(green.font).toContain('DejaVu Serif')
    expect(marked.font).toContain('Arial')
    expect(link.href).toBe('ppt/slides/slide2.xml')
    // automatic color: black on the default white page
    expect(shapes(slides[1])[0].text!.paras[0].runs[0].color).toBe('#000000')
  })

  test('master pages: background, drawings, footer fields where shown, hidden master drawings, notes size and hidden layers', async () => {
    const masters =
      '<draw:layer-set><draw:layer draw:name="layout"/><draw:layer draw:name="secret" draw:display="none"/></draw:layer-set>' +
      '<style:master-page style:name="Default" style:page-layout-name="PM1" draw:style-name="Mdp2">' +
      '<draw:rect draw:style-name="Mgr1" svg:x="0cm" svg:y="18cm" svg:width="25.4cm" svg:height="1cm"/>' +
      '<draw:frame presentation:style-name="Default-title" svg:x="1cm" svg:y="1cm" svg:width="23cm" svg:height="3cm" presentation:class="title" presentation:placeholder="true"><draw:text-box/></draw:frame>' +
      '<draw:frame svg:x="1cm" svg:y="18cm" svg:width="10cm" svg:height="1cm" presentation:class="footer"><draw:text-box><text:p><presentation:footer/></text:p></draw:text-box></draw:frame>' +
      '<draw:frame svg:x="20cm" svg:y="18cm" svg:width="4cm" svg:height="1cm" presentation:class="page-number"><draw:text-box><text:p>#<text:page-number/></text:p></draw:text-box></draw:frame>' +
      '<draw:frame svg:x="12cm" svg:y="18cm" svg:width="6cm" svg:height="1cm" presentation:class="date-time"><draw:text-box><text:p><presentation:date-time/></text:p></draw:text-box></draw:frame>' +
      '</style:master-page>'
    const auto =
      '<style:style style:name="Mdp2" style:family="drawing-page"><style:drawing-page-properties draw:fill="gradient" draw:fill-gradient-name="Sky"/></style:style>' +
      '<style:style style:name="Mgr1" style:family="graphic"><style:graphic-properties draw:stroke="none" draw:fill="solid" draw:fill-color="#123456"/></style:style>' +
      '<style:style style:name="dp1" style:family="drawing-page"><style:drawing-page-properties presentation:display-footer="true" presentation:display-page-number="true" presentation:display-date-time="true"/></style:style>' +
      '<style:style style:name="dp2" style:family="drawing-page"><style:drawing-page-properties presentation:background-objects-visible="false" presentation:display-footer="false" presentation:visibility="hidden"/></style:style>' +
      '<style:style style:name="gr1" style:family="graphic"><style:graphic-properties draw:stroke="none" draw:fill="solid" draw:fill-color="#ff0000"/></style:style>'
    const styles = '<draw:gradient draw:name="Sky" draw:style="linear" draw:start-color="#000080" draw:end-color="#ffffff" draw:angle="0"/>'
    const xml = fodp({
      styles,
      auto,
      masters,
      body:
        '<presentation:footer-decl presentation:name="ftr1">Quarterly review</presentation:footer-decl><presentation:date-time-decl presentation:name="dtd1" presentation:source="fixed">March 2026</presentation:date-time-decl>' +
        page('<draw:rect draw:style-name="gr1" draw:layer="secret" svg:x="1cm" svg:y="1cm" svg:width="1cm" svg:height="1cm"/>', ' draw:style-name="dp1" presentation:use-footer-name="ftr1" presentation:use-date-time-name="dtd1"') +
        page(
          '<presentation:notes><draw:frame presentation:class="notes" svg:x="1cm" svg:y="10cm" svg:width="18cm" svg:height="10cm"><draw:text-box><text:p>Remember the <text:span>numbers</text:span>.</text:p></draw:text-box></draw:frame></presentation:notes>',
          ' draw:style-name="dp2" presentation:use-footer-name="ftr1"',
        ),
    })
    const { slides, part } = await open(xml)
    expect(slides[0].bg).toMatchObject({ t: 'grad' })
    expect((slides[0].bg as { css: string }).css).toContain('#000080 0%')
    const s1 = shapes(slides[0])
    expect(s1[0]).toMatchObject({ fill: { t: 'solid', c: '#123456' } })
    expect(s1.map(textOf).filter(Boolean)).toEqual(['March 2026', 'Quarterly review', '#1'])
    // the hidden layer's rectangle is not drawn
    expect(s1.some((e) => e.fill?.t === 'solid' && e.fill.c === '#ff0000')).toBe(false)
    // the second page hides the master drawings, the footer, and itself
    expect(slides[1].hidden).toBe(true)
    expect(slides[1].els.length).toBe(0)
    expect(await part('ppt/slides/slide2.xml')).toContain('showMasterSp="0"')
    expect(await part('ppt/notesSlides/notesSlide2.xml')).toContain('Remember the ')
    expect(await part('ppt/presentation.xml')).toContain('<p:notesMasterIdLst>')
  })

  test('a picture cropped by fo:clip, an image fill, a table with merged cells, an embedded object', async () => {
    const auto =
      '<style:style style:name="gr1" style:family="graphic"><style:graphic-properties draw:stroke="none" draw:fill="none" fo:clip="rect(0.75pt, 0.75pt, 0pt, 0.75pt)"/></style:style>' +
      '<style:style style:name="gr2" style:family="graphic"><style:graphic-properties draw:stroke="none" draw:fill="bitmap" draw:fill-image-name="Dots" style:repeat="stretch"/></style:style>' +
      '<style:style style:name="ce1" style:family="table-cell"><style:graphic-properties draw:fill="solid" draw:fill-color="#cccccc"/><style:paragraph-properties fo:border="1pt solid #000000"/></style:style>'
    const img = base64(png(4, 2))
    const styles = `<draw:fill-image draw:name="Dots"><office:binary-data>${img}</office:binary-data></draw:fill-image>`
    const xml = fodp({
      styles,
      auto,
      body: page(
        `<draw:frame draw:style-name="gr1" svg:x="1cm" svg:y="1cm" svg:width="4cm" svg:height="2cm"><draw:image><office:binary-data>${img}</office:binary-data></draw:image></draw:frame>` +
          '<draw:rect draw:style-name="gr2" svg:x="6cm" svg:y="1cm" svg:width="4cm" svg:height="2cm"/>' +
          '<draw:frame svg:x="1cm" svg:y="5cm" svg:width="9cm" svg:height="3cm"><table:table><table:table-column table:number-columns-repeated="3"/>' +
          '<table:table-row><table:table-cell table:style-name="ce1" table:number-columns-spanned="2"><text:p>Wide</text:p></table:table-cell><table:covered-table-cell/><table:table-cell><text:p>C</text:p></table:table-cell></table:table-row>' +
          '<table:table-row><table:table-cell table:number-rows-spanned="1"><text:p>A</text:p></table:table-cell><table:table-cell><text:p>B</text:p></table:table-cell><table:table-cell/></table:table-row>' +
          '</table:table></draw:frame>' +
          '<draw:frame svg:x="12cm" svg:y="5cm" svg:width="6cm" svg:height="4cm"><draw:object xlink:href="./Object 1"/></draw:frame>',
      ),
    })
    const { slides } = await open(xml)
    const pic = slides[0].els.find((e) => e.t === 'pic')!
    // 4 × 2 px at 96 dpi: 3 × 1.5 pt
    expect(pic.t === 'pic' && pic.crop).toEqual([0.25, 0.5, 0.25, 0])
    expect(shapes(slides[0])[0].fill).toMatchObject({ t: 'img' })
    const tbl = slides[0].els.find((e) => e.t === 'table')!
    if (tbl.t !== 'table') throw new Error('no table')
    expect(tbl.cols.length).toBe(3)
    near(tbl.cols[0], px(3))
    expect(tbl.rows[0].cells.map((c) => [c.colSpan, c.hidden])).toEqual([
      [2, false],
      [1, true],
      [1, false],
    ])
    expect(tbl.rows[0].cells[0]).toMatchObject({ fill: { t: 'solid', c: '#cccccc' }, bl: { c: '#000000' } })
    expect(textOf(byText(slides[0], 'Object'))).toBe('Object')
  })

  test('shrink on overflow: a font scale for text that does not fit its frame', async () => {
    const auto = '<style:style style:name="pr1" style:family="presentation" style:parent-style-name="Default-outline1"><style:graphic-properties style:shrink-to-fit="true"/></style:style>'
    const items = Array.from({ length: 12 }, (_, i) => `<text:list-item><text:p>Point number ${i + 1} with some words</text:p></text:list-item>`).join('')
    const xml = fodp({
      auto,
      body: page(
        `<draw:frame presentation:style-name="pr1" presentation:class="outline" svg:x="1cm" svg:y="5cm" svg:width="23cm" svg:height="12cm"><draw:text-box><text:list>${items}</text:list></draw:text-box></draw:frame>` +
          '<draw:frame presentation:style-name="pr1" presentation:class="outline" svg:x="1cm" svg:y="5cm" svg:width="23cm" svg:height="12cm"><draw:text-box><text:list><text:list-item><text:p>Short</text:p></text:list-item></text:list></draw:text-box></draw:frame>',
      ),
    })
    const { slides } = await open(xml)
    const [long, short] = shapes(slides[0])
    const sz = long.text!.paras[0].runs[0].sz
    expect(sz).toBeLessThan(32)
    expect(sz).toBeGreaterThan(8)
    expect(short.text!.paras[0].runs[0].sz).toBe(32)
  })

  test('annotations become comments at their place', async () => {
    const note = (who: string, text: string) =>
      `<officeooo:annotation xmlns:officeooo="http://openoffice.org/2009/office" xmlns:dc="http://purl.org/dc/elements/1.1/" svg:x="2.54cm" svg:y="1.27cm"><dc:creator>${who}</dc:creator><dc:date>2025-10-09T01:09:01</dc:date><text:p>${text}</text:p><text:p>second line</text:p></officeooo:annotation>`
    const { engine, part } = await open(fodp({ body: page(note('Ann Lee', 'Check the &lt;numbers&gt;')) + page(note('Ann Lee', 'Fine')) }))
    const list = (await engine.call('comments.list', [])) as DocComment[]
    expect(list.map((c) => [c.author.name, c.text])).toEqual([
      ['Ann Lee', 'Check the <numbers>\nsecond line'],
      ['Ann Lee', 'Fine\nsecond line'],
    ])
    expect(list[0].time).toStartWith('2025-10-09T01:09:01')
    // 1 inch, half an inch: 96 and 48 px
    expect(list[0].at).toMatchObject({ x: 96, y: 48 })
    expect(await part('ppt/commentAuthors.xml')).toContain('name="Ann Lee" initials="AL" lastIdx="2"')
  })

  test('damaged input: no presentation body, unreadable geometry, deep nesting and huge repeat counts', async () => {
    const empty = `<?xml version="1.0"?><office:document ${NS} office:mimetype="application/vnd.oasis.opendocument.presentation"><office:body/></office:document>`
    await expect(convertFlatOdf(new TextEncoder().encode(empty))).rejects.toThrow('Not an OpenDocument presentation')
    const deep = '<draw:g>'.repeat(60) + '<draw:rect svg:x="1cm" svg:y="1cm" svg:width="1cm" svg:height="1cm"/>' + '</draw:g>'.repeat(60)
    const xml = fodp({
      body: page(
        '<draw:custom-shape svg:x="1cm" svg:y="1cm" svg:width="2cm" svg:height="2cm"><draw:enhanced-geometry svg:viewBox="0 0 21600 21600" draw:type="non-primitive" draw:enhanced-path="M ?f0 ?f1 L $9 Q ( Z N"><draw:equation draw:name="f0" draw:formula="?f1 +1"/><draw:equation draw:name="f1" draw:formula="?f0 *2"/></draw:enhanced-geometry></draw:custom-shape>' +
          deep +
          '<draw:frame svg:x="1cm" svg:y="5cm" svg:width="9cm" svg:height="3cm"><table:table><table:table-column table:number-columns-repeated="1000000"/><table:table-row table:number-rows-repeated="1000000"><table:table-cell table:number-columns-repeated="1000000"/></table:table-row></table:table></draw:frame>',
      ),
    })
    const t0 = performance.now()
    const { slides } = await open(xml)
    expect(performance.now() - t0).toBeLessThan(5000)
    expect(slides.length).toBe(1)
    expect(shapes(slides[0])[0].geom).not.toBe('line')
  })

  test('a package without content is not a presentation', async () => {
    const zip = await writeZip(new Map([['mimetype', 'application/vnd.oasis.opendocument.presentation'], ['content.xml', `<office:document-content ${NS}><office:body><office:text/></office:body></office:document-content>`]]))
    const engine = new Engine(() => {})
    await expect(engine.open(zip, { name: 'broken.odp' })).rejects.toThrow('Not an OpenDocument presentation')
    // the package still reads as a ZIP
    expect((await ZipArchive.open(new BlobSource(zip))).has('content.xml')).toBe(true)
  })
})
