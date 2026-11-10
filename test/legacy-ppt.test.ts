import { describe, expect, test } from 'bun:test'
import { readFileSync } from 'fs'
import { deflateSync } from 'node:zlib'
import type { SlideEl, SlideModel } from '../src/core/pptx/types'
import { Cfb } from '../src/core/legacy/cfb'
import { pptToPptx } from '../src/core/legacy/ppt'
import { BlobSource } from '../src/core/source'
import { Engine } from '../src/core/worker/engine'
import { ZipArchive } from '../src/core/zip'
import { writeCfb } from './cfb-writer'
import { cat, comment10, defaultMaster, filled, fixed, group, lined, oa, oaScheme, P, PptBuilder, progTag, rec, samplePpt, shape, TX, u16, u32 } from './ppt-builder'

const fixture = (name: string) => new Uint8Array(readFileSync(new URL(`./fixtures/legacy/${name}`, import.meta.url)))

async function openPpt(bytes: Uint8Array<ArrayBuffer>) {
  const engine = new Engine(() => {})
  const info = await engine.open(new Blob([bytes]), { name: 'deck.ppt' })
  const slides: SlideModel[] = []
  for (let i = 0; i < (info.pptx?.count ?? 0); i++) slides.push((await engine.call('pptx.slide', [i])) as SlideModel)
  return { engine, info, slides }
}

/** the converted package */
async function pptx(bytes: Uint8Array) {
  return ZipArchive.open(new BlobSource(await pptToPptx(Cfb.open(bytes))))
}

type Shape = Extract<SlideEl, { t: 'shape' }>
const shapes = (s: SlideModel) => s.els.filter((e): e is Shape => e.t === 'shape')
const textOf = (e: SlideEl | undefined) => (e && e.t === 'shape' ? (e.text?.paras.map((p) => p.runs.map((r) => r.text).join('')).join('\n') ?? '') : '')
const withText = (s: SlideModel, t: string) => shapes(s).find((e) => textOf(e) === t || textOf(e).startsWith(t))!
/** master units (576 per inch) → slide px (96 per inch) */
const px = (mu: number) => mu / 6

/** a deck with the default master and one slide of these shapes */
function oneSlide(shapesOf: (b: PptBuilder) => Uint8Array[], opts: { texts?: Parameters<PptBuilder['slide']>[0]['texts'] } = {}) {
  const b = new PptBuilder()
  b.master(defaultMaster())
  b.slide({ shapes: shapesOf(b), texts: opts.texts })
  return b
}

describe('PowerPoint 97-2003', () => {
  test('the sample deck: title master, bullets, numbering, link, shapes, picture, table, hidden slide, footers', async () => {
    const { info, slides } = await openPpt(samplePpt())
    expect(info.kind).toBe('pptx')
    expect(info.converted).toBe('ppt')
    expect(info.pptx).toMatchObject({ w: 960, h: 720, count: 5 })
    expect(slides.map((s) => !!s.hidden)).toEqual([false, false, false, false, true])

    // 1: the title master's gradient background and look; its main master's band is not drawn
    expect(slides[0].bg?.t).toBe('grad')
    expect(withText(slides[0], 'office-engine').text!.paras[0]).toMatchObject({ algn: 'center', runs: [{ text: 'office-engine', b: true, sz: 54, color: '#ffffff' }] })
    expect(textOf(shapes(slides[0])[1])).toBe('Progressive Office viewer for the browser')
    expect(shapes(slides[0]).some((e) => e.fill?.t === 'solid' && e.fill.c === '#2563eb')).toBe(false)

    // 2: bullets by level (master style, a Wingdings bullet), bold / colored / linked runs, numbering
    const s2 = slides[1]
    expect(shapes(s2)[0]).toMatchObject({ fill: { t: 'solid', c: '#2563eb' }, box: { y: px(4176), h: px(144) } })
    expect(textOf(withText(s2, 'Why'))).toBe('Why it is fast')
    const body = withText(s2, 'Lazy ZIP').text!
    expect(body.paras.map((p) => [p.bullet?.text, p.lvl ?? 0])).toEqual([
      ['•', 0],
      ['–', 1],
      ['•', 0],
      ['■', 1],
      ['•', 0],
    ])
    expect(body.paras[0]).toMatchObject({ marL: px(270), indent: -px(270) })
    expect(body.paras[1]).toMatchObject({ marL: px(558), indent: -px(270) })
    expect(body.paras[3].bullet?.color).toBe('#2563eb')
    expect(body.paras[0].runs.map((r) => [r.text, !!r.b, r.sz])).toEqual([
      ['Lazy ZIP: only the ', false, 28],
      ['central directory', true, 28],
      [' is read up front', false, 28],
    ])
    expect(body.paras[1].runs[0]).toMatchObject({ i: true, sz: 24 })
    expect(body.paras[2].runs[1]).toMatchObject({ text: 'Web Worker', color: '#c00000', u: true })
    expect(body.paras[4].runs[1]).toMatchObject({ text: 'example.com', href: 'https://example.com/office-engine', u: true })
    expect(withText(s2, 'Open').text!.paras.map((p) => p.bullet?.text)).toEqual(['1.', '2.', '3.'])
    // footers: the master's placeholders with this slide's values
    expect(textOf(withText(s2, 'office-engine sample'))).toBe('office-engine sample')
    expect(shapes(s2).map(textOf)).toContain('2')

    // 3: shapes
    const s3 = shapes(slides[2])
    expect(withText(slides[2], 'Rounded')).toMatchObject({ geom: 'roundRect', adj: { adj: 16667 }, fill: { t: 'solid', c: '#2563eb' } })
    expect(withText(slides[2], 'Rounded').text!.paras[0].runs[0]).toMatchObject({ b: true, color: '#ffffff' })
    expect(withText(slides[2], 'Oval')).toMatchObject({ geom: 'ellipse', fill: { c: '#10b981' }, stroke: { c: '#065f46', w: 25400 / 9525 } })
    const arrow = withText(slides[2], 'Arrow')
    expect(arrow.geom).toBe('rightArrow')
    expect(arrow.fill).toEqual({ t: 'grad', css: 'linear-gradient(0deg, #fef3c7 0%, #f59e0b 100%)' })
    expect(withText(slides[2], 'A free text box').text!.paras[0].runs[1]).toMatchObject({ b: true, color: '#7c3aed', font: expect.stringContaining('Georgia') })
    const line = s3.find((e) => e.geom === 'line')!
    expect(line.stroke).toMatchObject({ c: '#dc2626', tail: 'triangle' })
    expect(line.stroke?.dash).toBeDefined()
    expect(s3.filter((e) => e.geom === 'custom').map((e) => e.fill?.t)).toEqual(['solid', 'none'])
    expect(s3.find((e) => e.geom === 'custom' && e.fill?.t === 'none')?.path).toContain('A')
    expect(withText(slides[2], 'Rotated').box.rot).toBe(345)
    const pic = slides[2].els.find((e) => e.t === 'pic')!
    expect(pic).toMatchObject({ t: 'pic', box: { x: px(3312), y: px(2160), w: px(1440), h: px(864) } })
    expect(pic.t === 'pic' && pic.crop?.map((v) => Math.round(v * 1000) / 1000)).toEqual([0.1, 0, 0, 0])
    expect(pic.t === 'pic' && pic.target).toMatch(/^ppt\/media\/image1\.png$/)

    // 4: a table stored as a group of cells: the group's coordinates scale them
    const cells = shapes(slides[3]).filter((e) => /^(Metric|20\d\d|Users|Revenue|Latency|\d+)$/.test(textOf(e)) && e.fill?.t === 'solid')
    expect(cells).toHaveLength(16)
    const users = withText(slides[3], 'Users')
    expect(users.box.x).toBeCloseTo(px(720), 5)
    expect(users.box.y).toBeCloseTo(px(1152 + (400 * 1728) / 1600), 5)
    expect(users.box.w).toBeCloseTo(px((1000 * 4320) / 4000), 5)
    expect(withText(slides[3], 'Metric').fill).toEqual({ t: 'solid', c: '#1d4ed8' })
    expect(withText(slides[3], '380').text!.paras[0].algn).toBe('right')

    // 5: hidden, its own background and color scheme (the master's band takes the slide's fill color)
    expect(slides[4].bg).toEqual({ t: 'solid', c: '#fef3c7' })
    expect(shapes(slides[4])[0].fill).toEqual({ t: 'solid', c: '#f59e0b' })
    expect(withText(slides[4], 'Hidden').text!.paras[0].runs[0].color).toBe('#92400e')
  })

  test('the package: masters with layouts and themes, notes, placeholders, content types', async () => {
    const zip = await pptx(samplePpt())
    const names = zip.list().map((e) => e.name)
    expect(names).toEqual(expect.arrayContaining(['ppt/notesMasters/notesMaster1.xml', 'ppt/notesSlides/notesSlide2.xml', 'ppt/media/image1.png', 'ppt/theme/theme1.xml']))
    expect(await zip.text('ppt/notesSlides/notesSlide2.xml')).toContain('Speaker notes: mention the worker and the lazy ZIP reader.')
    expect(await zip.text('ppt/slides/_rels/slide2.xml.rels')).toContain('notesSlide2.xml')
    const types = await zip.text('[Content_Types].xml')
    for (const n of names.filter((n) => /^ppt\/(slides|slideLayouts|slideMasters|notesSlides|notesMasters|theme)\/[^/]+\.xml$/.test(n))) expect(types).toContain(`PartName="/${n}"`)
    expect(types).toContain('Extension="png" ContentType="image/png"')
    // the theme takes the master's color scheme and fonts
    const theme = await zip.text('ppt/theme/theme1.xml')
    expect(theme).toContain('<a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>')
    expect(theme).toContain('<a:accent1><a:srgbClr val="2563EB"/></a:accent1>')
    // placeholders keep their kind; the title master is a "Title Slide" layout hiding the master's shapes
    const s1 = await zip.text('ppt/slides/slide1.xml')
    expect(s1).toContain('<p:ph type="ctrTitle"/>')
    expect(s1).toContain('<p:ph type="subTitle" idx="1"/>')
    const layout = await zip.text((await zip.text('ppt/slides/_rels/slide1.xml.rels')).match(/slideLayouts\/(slideLayout\d+\.xml)/)![0].replace(/^/, 'ppt/'))
    expect(layout).toContain('type="title"')
    expect(layout).toContain('showMasterSp="0"')
    const master = await zip.text('ppt/slideMasters/slideMaster1.xml')
    expect(master).toMatch(/<p:titleStyle><a:lvl1pPr[^>]* algn="ctr"/)
    expect(master).toContain('<p:ph type="sldNum" idx="12"/>')
    expect(await zip.text('ppt/presentation.xml')).toContain('<p:sldSz cx="9144000" cy="6858000"/>')
  })

  test('saving gives a PowerPoint file', async () => {
    const { engine } = await openPpt(samplePpt())
    const saved = (await engine.call('save', [])) as Blob
    const again = new Engine(() => {})
    const info = await again.open(saved, { name: 'deck.pptx' })
    expect(info.converted).toBeUndefined()
    expect(info.pptx?.count).toBe(5)
    expect(textOf(withText((await again.call('pptx.slide', [1])) as SlideModel, 'Why'))).toBe('Why it is fast')
  })

  test('the committed sample opens', async () => {
    const { slides } = await openPpt(fixture('sample.ppt'))
    expect(slides).toHaveLength(5)
    expect(textOf(withText(slides[0], 'office-engine'))).toBe('office-engine')
  })

  test('the latest edit wins; without a Current User stream the last edit is found', async () => {
    const b = oneSlide(() => [shape({ spt: 1, at: [432, 173, 4896, 720], ph: [0x0d, 0], outline: 0 })], { texts: [{ type: TX.title, paras: ['First save'] }] })
    const ids = b.ids
    // a second save rewrites the slide with another title
    const first = b.build()
    const other = oneSlide(() => [shape({ spt: 1, at: [432, 173, 4896, 720], ph: [0x0d, 0], outline: 0 })], { texts: [{ type: TX.title, paras: ['Second save'] }] })
    b.edit(new Map([[ids.slides[0], other.slideRecord(0)]]))
    // the slide's text lives in the document's slide list: the document record is saved again too
    const otherDoc = Cfb.open(other.build()).stream('PowerPoint Document')!
    const docLen = (otherDoc[4] | (otherDoc[5] << 8) | (otherDoc[6] << 16)) + 8
    b.edit(new Map([[1, otherDoc.subarray(0, docLen)]]))
    const second = b.build()
    expect(textOf(shapes((await openPpt(first)).slides[0])[0])).toBe('First save')
    expect(textOf(shapes((await openPpt(second)).slides[0])[0])).toBe('Second save')

    const cfb = Cfb.open(second)
    const noUser = writeCfb({ 'PowerPoint Document': cfb.stream('PowerPoint Document')! })
    expect(textOf(shapes((await openPpt(noUser)).slides[0])[0])).toBe('Second save')
  })

  test('encrypted, PowerPoint 95 and damaged files say what is wrong', async () => {
    const fail = (streams: Record<string, Uint8Array>) => new Engine(() => {}).open(new Blob([writeCfb(streams)]), { name: 'x.ppt' })
    const doc = Cfb.open(samplePpt()).stream('PowerPoint Document')!
    const user = (token: number) => rec(0x0ff6, cat(u32(0x14), u32(token), u32(0), u16(1), u16(0x03f4), [3, 0], u16(0), [65], u32(8)))
    await expect(fail({ 'PowerPoint Document': doc, 'Current User': user(0xf3d1c4df) })).rejects.toThrow('Password-protected')
    await expect(fail({ 'PowerPoint Document': doc, 'Current User': user(0xe391c05f), EncryptedSummary: new Uint8Array(8) })).rejects.toThrow('Password-protected')
    // PowerPoint 95: a 4-byte size before the user name in Current User
    const user95 = cat(u32(4), [65, 66, 67, 68])
    await expect(fail({ 'PowerPoint Document': new Uint8Array(64), 'Current User': user95 })).rejects.toThrow('PowerPoint 95')
    await expect(fail({ 'PowerPoint Document': new Uint8Array(64), 'Current User': user(0xe391c05f) })).rejects.toThrow('damaged')
  })

  test('damaged files fail with an error or open, never hang', async () => {
    const good = samplePpt()
    const cfb = Cfb.open(good)
    const doc = cfb.stream('PowerPoint Document')!
    const pictures = cfb.stream('Pictures')!
    const user = cfb.stream('Current User')!
    let seed = 7
    const rand = (n: number) => ((seed = (seed * 1103515245 + 12345) >>> 0) % n)
    const started = Date.now()
    for (let k = 0; k < 120; k++) {
      const d = doc.slice()
      if (k % 3 === 0) {
        for (let i = 0; i < 20; i++) d[rand(d.length)] = rand(256)
      } else if (k % 3 === 1) {
        // a record length or offset made huge
        const at = rand(d.length - 4)
        d.set([0xff, 0xff, 0xff, rand(2) ? 0x7f : 0xff], at)
      }
      const streams = { 'PowerPoint Document': k % 3 === 2 ? d.subarray(0, rand(d.length)) : d, 'Current User': user, Pictures: pictures }
      try {
        await pptToPptx(Cfb.open(writeCfb(streams)))
      } catch (e) {
        expect(e).toBeInstanceOf(Error)
      }
    }
    expect(Date.now() - started).toBeLessThan(20000)
  })

  test('text: 8-bit characters, line breaks, tabs, superscript, fields, symbol bullets, alignment, spacing', async () => {
    const b = oneSlide(() => [
      shape({
        spt: 202,
        at: [432, 432, 4000, 2000],
        props: [[P.fillBooleans, filled(false)], [P.lineBooleans, lined(false)], [0x0081, 0], [P.anchorText, 2]],
        text: {
          type: TX.other,
          bytes: true,
          paras: [
            { pf: { align: 2, spaceBefore: -144, lineSpacing: 150 }, runs: [{ text: 'Café\u000bE=mc' }, { text: '2', cf: { position: 30 } }, { text: '\tend' }] },
            { pf: { bullet: true, bulletChar: '§', bulletFont: 1, bulletSize: 120 }, runs: [{ text: 'Symbol bullet' }] },
            { runs: [{ text: 'Slide ' }, { text: '*', field: 'slidenum' }] },
          ],
        },
      }),
    ])
    b.font('Wingdings')
    const { slides } = await openPpt(b.build())
    const t = shapes(slides[0]).find((e) => textOf(e).startsWith('Café'))!.text!
    expect(t.anchor).toBe('b')
    expect(t.ins[0]).toBe(0)
    expect(t.paras[0]).toMatchObject({ algn: 'right', lnPct: 1.5, spcBef: 24 })
    expect(t.paras[0].runs.map((r) => [r.text, !!r.br, r.baseline])).toEqual([
      ['Café', false, undefined],
      ['', true, undefined],
      ['E=mc', false, undefined],
      ['2', false, 30],
      ['\tend', false, undefined],
    ])
    expect(t.paras[1].bullet).toMatchObject({ text: '▪', sz: 18 * 1.2 })
    expect(t.paras[2].runs.map((r) => r.text).join('')).toBe('Slide 1')
  })

  test('pictures: JPEG and DIB in the Pictures stream, EMF kept as a file, missing ones left out', async () => {
    const b = new PptBuilder()
    b.master(defaultMaster())
    const jpeg = new Uint8Array([0xff, 0xd8, 0xff, 0xe0, 0, 0x10, 74, 70, 73, 70, 0, 0xff, 0xd9])
    // 2 x 2, 24-bit, rows padded to 4 bytes
    const dib = cat(u32(40), u32(2), u32(2), u16(1), u16(24), u32(0), u32(16), u32(2835), u32(2835), u32(0), u32(0), new Uint8Array(16).fill(0x80))
    const emf = cat([1, 0, 0, 0], u32(108), new Uint8Array(32), [0x20, 0x45, 0x4d, 0x46], new Uint8Array(64))
    const packed = deflateSync(emf)
    const emfBlip = rec(0xf01a, cat(new Uint8Array(16).fill(9), u32(emf.length), u32(0), u32(0), u32(100), u32(100), u32(914400), u32(914400), u32(packed.length), [0, 0xfe], packed), 0x3d4)
    const pics = [b.picture(jpeg, 'jpg'), b.picture(dib, 'dib'), b.rawBlip(emfBlip)]
    b.slide({ shapes: [...pics, 99].map((pib, i) => shape({ spt: 75, at: [432 + i * 1200, 432, 1000, 800], props: [[P.pib, pib]] })) })
    const bytes = b.build()
    const { slides } = await openPpt(bytes)
    expect(slides[0].els.map((e) => e.t === 'pic' && e.target)).toEqual(['ppt/media/image1.jpg', 'ppt/media/image2.bmp', 'ppt/media/image3.emf'])
    const zip = await pptx(bytes)
    expect(await zip.bytes('ppt/media/image1.jpg')).toEqual(jpeg)
    const bmp = await zip.bytes('ppt/media/image2.bmp')
    expect([bmp[0], bmp[1], bmp[10]]).toEqual([0x42, 0x4d, 54])
    expect(bmp.subarray(14)).toEqual(dib)
    expect(await zip.bytes('ppt/media/image3.emf')).toEqual(emf)
    expect(await zip.text('[Content_Types].xml')).toContain('Extension="emf" ContentType="image/x-emf"')
  })

  test('shapes: flips, rotated frames, connectors, shadows, picture fills, master shapes, hidden shapes', async () => {
    const b = new PptBuilder()
    const dm = defaultMaster()
    b.master(dm)
    const pib = b.picture(new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]))
    b.slide({
      shapes: [
        // turned by 90 degrees: the anchor is the turned box, the frame swaps width and height around the center
        shape({ spt: 1, at: [1000, 1000, 600, 200], props: [[P.rotation, fixed(90)], [P.fillColor, oa('FF0000')]], flags: 0x40 }),
        shape({ spt: 32, at: [100, 100, 500, 300], props: [[P.lineColor, oaScheme(1)], [P.lineEndArrowhead, 2]], flags: 0x80 }),
        shape({ spt: 1, at: [2000, 2000, 300, 300], props: [[P.fillType, 3], [P.fillBlip, pib], [0x023f, 0x00020002], [P.lineBooleans, lined(false)]] }),
        shape({ spt: 1, at: [3000, 2000, 300, 300], props: [[P.fillColor, oa('00FF00')], [P.groupBooleans, 0x00020002]] }),
        shape({ spt: 3, at: [3500, 3000, 300, 300], master: dm.ph.title }),
        group([4000, 100, 1000, 1000], [0, 0, 100, 100], [group([0, 0, 50, 50], [0, 0, 10, 10], [shape({ spt: 1, child: true, at: [5, 5, 5, 5], props: [[P.fillColor, oa('0000FF')]] })], true)]),
      ],
    })
    const { slides } = await openPpt(b.build())
    // the hidden shape is left out; the master title has no fill and no line, so a shape following it has neither
    const [turned, connector, picture, nested, ...rest] = shapes(slides[0])
    expect(rest).toHaveLength(0)
    expect(turned.box).toMatchObject({ x: px(1200), y: px(800), w: px(200), h: px(600), rot: 90, flipH: true })
    expect(connector).toMatchObject({ geom: 'straightConnector1', box: { flipV: true }, stroke: { c: '#000000', tail: 'stealth' } })
    expect(picture.fill).toEqual({ t: 'img', target: 'ppt/media/image1.png', tile: false })
    // groups in groups: each scales its children into its frame
    expect(nested.box.x).toBeCloseTo(px(4250), 1)
    expect(nested.box.y).toBeCloseTo(px(350), 1)
    expect(nested.box.w).toBeCloseTo(px(250), 1)
    const zip = await pptx(b.build())
    expect(await zip.text('ppt/slides/slide1.xml')).toContain('<a:outerShdw')
  })

  test('comments become PPTX comments with their author, date, text and place', async () => {
    const b = new PptBuilder()
    b.master(defaultMaster())
    b.slide({ shapes: [], extra: [progTag('___PPT10', [comment10('Ann Lee', 'AL', 'Check the numbers\r\nbefore Friday', [2007, 3, 9, 14, 5, 30], [576, 288])])] })
    const { engine } = await openPpt(b.build())
    const list = (await engine.call('comments.list', [])) as { author: { name: string; initials?: string }; time: string; text: string; at: { x: number; y: number } }[]
    expect(list).toHaveLength(1)
    expect(list[0]).toMatchObject({ author: { name: 'Ann Lee', initials: 'AL' }, time: '2007-03-09T14:05:30Z', text: 'Check the numbers\nbefore Friday', at: { x: 96, y: 48 } })
  })
})
