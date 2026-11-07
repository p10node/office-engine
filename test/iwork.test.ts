import { describe, expect, test } from 'bun:test'
import type { Block, DocxMeta, Paragraph, Table } from '../src/core/docx/types'
import type { SlideEl, SlideModel } from '../src/core/pptx/types'
import type { BlockOut } from '../src/core/xlsx/types'
import { durationText, excelDate, excelFormat, shownText } from '../src/core/iwork/format'
import { readIwa } from '../src/core/iwork/iwa'
import { psFont } from '../src/core/iwork/styles'
import { Msg } from '../src/core/iwork/proto'
import { unsnappy } from '../src/core/iwork/snappy'
import { Engine } from '../src/core/worker/engine'
import { writeZip } from '../src/core/zipwrite'
import { appleDate, charProps, drawable, encode, f32, fixed32, fmt, iwa, IworkBuilder, K, msg, ref, sampleKeynote, sampleNumbers, samplePages, shape, shapeStyle, storage, styleInfo, table, T } from './iwork-builder'

async function open(bytes: Uint8Array<ArrayBuffer>) {
  const events: { event: string; data: any }[] = []
  const engine = new Engine((event, data) => events.push({ event, data }))
  const info = await engine.open(new Blob([bytes]), { name: 'doc.pages' })
  await engine.call('docx.start', [])
  const done = events.find((e) => e.event === 'docx:done')?.data
  if (done?.error) throw new Error(done.error)
  const blocks: Block[] = events.filter((e) => e.event === 'docx:blocks').flatMap((e) => e.data)
  const meta: DocxMeta = events.filter((e) => e.event === 'docx:meta').at(-1)?.data
  return { engine, info, blocks, meta, paras: blocks.filter((b): b is Paragraph => b.t === 'p') }
}
const textOf = (p: Paragraph) => p.runs.map((r) => (r.t === 'text' ? r.text : r.t === 'field' ? `{${r.kind}}` : r.t === 'ref' ? r.text : '')).join('')
const cellText = (t: Table, r: number, c: number) => t.rows[r].cells[c].blocks.map((b) => (b.t === 'p' ? textOf(b) : '')).join('\n')

/** a minimal Pages document around one body storage */
async function pagesWith(build: (b: IworkBuilder, body: number) => { text: string; extra?: Parameters<typeof storage>[1] extends infer S ? Partial<S> : never }) {
  const b = new IworkBuilder()
  const bodyStyle = b.add(T.PARA_STYLE, msg({ 1: styleInfo('Body'), 11: charProps({ size: 12, font: 'Helvetica' }) }))
  const { text, extra } = build(b, bodyStyle)
  const body = storage(b, { text, paras: [[0, bodyStyle]], ...extra })
  b.add(T.DOCUMENT, msg({ 4: ref(body), 30: f32(595), 31: f32(842) }), 1)
  return b.zip()
}

describe('snappy', () => {
  test('literals and back-references (including overlapping ones)', () => {
    // "abcabcabcX": literal "abc", copy (offset 3, length 6), literal "X"
    const block = Uint8Array.of(10, 2 << 2, 97, 98, 99, 1 | ((6 - 4) << 2), 3, 0 << 2, 88)
    expect(new TextDecoder().decode(unsnappy(block))).toBe('abcabcabcX')
    // 2-byte offset copy: "ab" then copy offset 2 length 4 → "ababab"
    const two = Uint8Array.of(6, 1 << 2, 97, 98, 2 | ((4 - 1) << 2), 2, 0)
    expect(new TextDecoder().decode(unsnappy(two))).toBe('ababab')
  })

  test('corrupt input fails instead of reading past the data', () => {
    expect(() => unsnappy(Uint8Array.of(5, 1 | (0 << 2), 9))).toThrow()
    expect(() => unsnappy(Uint8Array.of(4, 0, 97))).toThrow()
  })
})

describe('protobuf reader', () => {
  test('varints, negatives, floats, strings, nested, packed and repeated fields', () => {
    const m = new Msg(
      encode(
        msg({
          1: 300,
          2: -5,
          3: f32(1.5),
          4: 'héllo',
          5: msg({ 1: 7 }),
          6: [1, 2, 3],
          7: new Uint8Array([1, 0x96, 0x01, 3]),
          8: [f32(0.5), f32(2)],
          9: fixed32(0x00020003),
          10: 2 ** 40,
        }),
      ),
    )
    expect(m.u(1)).toBe(300)
    expect(m.u(2)).toBe(-5)
    expect(m.float(3)).toBe(1.5)
    expect(m.str(4)).toBe('héllo')
    expect(m.msg(5)?.u(1)).toBe(7)
    expect(m.ref(5)).toBe(7)
    expect(m.us(6)).toEqual([1, 2, 3])
    // packed varints
    expect(m.us(7)).toEqual([1, 150, 3])
    expect(m.floats(8)).toEqual([0.5, 2])
    expect(m.u(9)).toBe(0x00020003)
    expect(m.u(10)).toBe(2 ** 40)
    expect(m.has(11)).toBe(false)
  })

  test('IWA archives: objects by identifier with their type', () => {
    const objects = readIwa(iwa([{ id: 1, type: 10000, body: encode(msg({ 4: ref(7) })) }, { id: 7, type: 2001, body: encode(msg({ 3: 'x' })) }]))
    expect(objects.get(1)?.type).toBe(10000)
    expect(new Msg(objects.get(7)!.data).str(3)).toBe('x')
  })
})

describe('fonts', () => {
  test('PostScript names give the family and the face', () => {
    expect(psFont('HelveticaNeue-BoldItalic')).toEqual({ family: 'Helvetica Neue', b: true, i: true })
    expect(psFont('TimesNewRomanPSMT')).toEqual({ family: 'Times New Roman' })
    expect(psFont('ArialMT')).toEqual({ family: 'Arial' })
    expect(psFont('LucidaGrande')).toEqual({ family: 'Lucida Grande' })
    expect(psFont('TWKEverett-Medium')).toEqual({ family: 'TWK Everett' })
    expect(psFont('PingFangSC-Regular')).toEqual({ family: 'PingFang SC' })
    expect(psFont('Avenir Next')).toEqual({ family: 'Avenir Next' })
  })
})

describe('pages', () => {
  test('the sample document: styles, formatting, link, lists, image, table, header, footer, footnote', async () => {
    const { info, blocks, paras, meta } = await open(await samplePages())
    expect(info.kind).toBe('docx')
    expect(info.from).toBe('pages')
    expect(meta.page).toMatchObject({ w: 612, h: 792, ml: 72 })

    const title = paras[0]
    expect(textOf(title)).toBe('Quarterly Field Report')
    expect(title.runs[0].t === 'text' && title.runs[0].s).toMatchObject({ b: true, sz: 26 })

    const intro = paras.find((p) => textOf(p).startsWith('This report'))!
    const run = (t: string) => intro.runs.find((r) => r.t === 'text' && r.text === t)
    expect(run('bold')).toMatchObject({ s: { b: true } })
    expect(run('italic, colored')).toMatchObject({ s: { i: true, color: '#b45309' } })
    expect(run('the full report')).toMatchObject({ href: 'https://example.com/report' })

    const items = paras.filter((p) => p.marker)
    expect(items.map((p) => [p.marker!.text, p.ilvl, textOf(p)])).toEqual([
      ['•', 0, 'Coverage grew in every region'],
      ['◦', 1, 'Northern sites reported first'],
      ['•', 0, 'Survey response time fell by a third'],
      ['1.', 0, 'Collect the remaining samples'],
      ['2.', 0, 'Publish the regional summaries'],
    ])

    const img = paras.flatMap((p) => p.runs).find((r) => r.t === 'img')
    expect(img).toMatchObject({ w: 360, h: 120 })

    const tbl = blocks.find((b): b is Table => b.t === 'tbl')!
    expect(tbl.rows.length).toBe(4)
    expect(cellText(tbl, 0, 0)).toBe('Region')
    expect(tbl.rows[0].cells[0].shd).toBe('#1d4ed8')
    expect(cellText(tbl, 1, 2)).toBe('1480.5')
    expect(cellText(tbl, 3, 1)).toBe('21')

    // footnote: a superscript mark in the text, the note after the body
    const noted = paras.find((p) => textOf(p).startsWith('Readings are averaged'))!
    expect(textOf(noted)).toBe('Readings are averaged per site1.')
    expect(paras.at(-1) && textOf(paras.at(-1)!)).toBe('1 Measured on the 2024 survey sample (n = 1,204).')

    // header: left and right parts; footer: the page number field
    const head = meta.header?.find((b): b is Table => b.t === 'tbl')
    expect(head && [cellText(head, 0, 0), cellText(head, 0, 2)]).toEqual(['Northwind Field Lab', 'Q3 report'])
    expect(meta.footer?.filter((b): b is Paragraph => b.t === 'p').map(textOf)).toContain('Page {PAGE}')
  })

  test('saving gives a Word document', async () => {
    const { engine } = await open(await samplePages())
    const docx = (await engine.call('save', [])) as Blob
    const again = await open(new Uint8Array(await docx.arrayBuffer()))
    expect(again.info.from).toBeUndefined()
    expect(textOf(again.paras[0])).toBe('Quarterly Field Report')
  })

  test('paragraph styles: a missing entry keeps the previous style; character styles end where their entry does', async () => {
    let heading = 0
    let strong = 0
    const bytes = await pagesWith((b, body) => {
      heading = b.add(T.PARA_STYLE, msg({ 1: styleInfo('Heading 1', body), 11: charProps({ size: 20, bold: true }), 12: msg({ 27: 0 }) }))
      strong = b.add(T.CHAR_STYLE, msg({ 1: styleInfo('Strong'), 11: charProps({ bold: true }) }))
      // "Head\nOne\nTwo two\n": Heading, (same), Body
      return { text: 'Head\nOne\nTwo two\n', extra: { paras: [[0, heading], [5, undefined], [9, body]], chars: [[0, undefined], [13, strong], [16, undefined]] } }
    })
    const { paras } = await open(bytes)
    expect(paras.map(textOf)).toEqual(['Head', 'One', 'Two two'])
    expect(paras[0].sid).toBe(`P${heading}`)
    expect(paras[1].sid).toBe(`P${heading}`)
    expect(paras[1].ps.outline).toBe(0)
    expect(paras[2].runs.map((r) => (r.t === 'text' ? [r.text, !!r.s.b] : null))).toEqual([
      ['Two ', false],
      ['two', true],
    ])
  })

  test('tabs, line breaks, page breaks and astral characters', async () => {
    const { paras } = await open(await pagesWith(() => ({ text: 'a\tb c\u0005next 😀 page\n' })))
    expect(textOf(paras[0])).toBe('a\tb\nc')
    expect(paras[1].pb).toBe(true)
    expect(textOf(paras[1])).toBe('next 😀 page')
  })

  test('tables: pre-BNC cell storage and merged cells', async () => {
    const bytes = await pagesWith((b) => {
      const att = table(
        b,
        [
          [{ text: 'Wide' }, {}, { text: 'C' }],
          [{ num: 1 }, { num: 2.25 }, { text: 'x' }],
        ],
        { format: 'v3', merges: [[0, 0, 1, 2]] },
      )
      return { text: 'Before\n\ufffc\nAfter\n', extra: { attachments: [[7, att]] } }
    })
    const { blocks } = await open(bytes)
    expect(blocks.map((b) => (b.t === 'p' ? textOf(b) : b.t))).toEqual(['Before', 'tbl', 'After'])
    const t = blocks[1] as Table
    expect(t.rows[0].cells.length).toBe(2)
    expect(t.rows[0].cells[0].colSpan).toBe(2)
    expect(cellText(t, 0, 0)).toBe('Wide')
    expect(cellText(t, 0, 1)).toBe('C')
    expect(cellText(t, 1, 1)).toBe('2.25')
  })

  test('password-protected and iWork ’09 files say what is not supported', async () => {
    const zip = async (files: Record<string, string | Uint8Array>) => new Uint8Array(await (await writeZip(new Map(Object.entries(files)))).arrayBuffer())
    const doc = iwa([{ id: 1, type: 1, body: encode(msg({})) }])
    const fail = (bytes: Uint8Array<ArrayBuffer>) => new Engine(() => {}).open(new Blob([bytes]), {})
    await expect(fail(await zip({ 'Index/Document.iwa': doc, '.iwpv2': 'x' }))).rejects.toThrow('Password-protected')
    await expect(fail(await zip({ 'index.xml.gz': 'x' }))).rejects.toThrow('not supported')
  })
})

async function openKey(bytes: Uint8Array<ArrayBuffer>) {
  const engine = new Engine(() => {})
  const info = await engine.open(new Blob([bytes]), { name: 'deck.key' })
  const slides: SlideModel[] = []
  for (let i = 0; i < info.pptx!.count; i++) slides.push((await engine.call('pptx.slide', [i])) as SlideModel)
  return { engine, info, slides }
}
const textOfShape = (e: SlideEl | undefined) => (e && e.t === 'shape' ? (e.text?.paras.map((p) => p.runs.map((r) => r.text).join('')).join('\n') ?? '') : '')
const shapes = (s: SlideModel) => s.els.filter((e): e is Extract<SlideEl, { t: 'shape' }> => e.t === 'shape')
/** pt → slide px (EMU / 9525) */
const px = (pt: number) => (pt * 4) / 3

describe('keynote', () => {
  test('the sample deck: master objects, placeholders, bullets, cropped image, table, shapes, group, skipped slide', async () => {
    const { info, slides } = await openKey(await sampleKeynote())
    expect(info.kind).toBe('pptx')
    expect(info.from).toBe('keynote')
    expect(info.pptx).toMatchObject({ w: px(1024), h: px(768), count: 5 })
    expect(slides.map((s) => !!s.hidden)).toEqual([false, false, false, false, true])

    // the master's band is under every slide
    for (const s of slides) expect(shapes(s)[0]).toMatchObject({ fill: { t: 'solid', c: '#1d4ed8' }, box: { y: px(744), h: px(24) } })

    expect(shapes(slides[0]).map(textOfShape).slice(1)).toEqual(['Quarterly Review', 'Keynote presentation opened in the browser'])
    const title = shapes(slides[0])[1]
    expect(title.text?.anchor).toBe('ctr')
    expect(title.text?.paras[0]).toMatchObject({ algn: 'center', runs: [{ b: true, sz: 54, color: '#111827' }] })

    // bullets with their levels; the image cropped by its mask; a text-sized box placed by its center
    const body = shapes(slides[1]).find((e) => textOfShape(e).startsWith('Coverage'))!
    expect(body.text!.paras.map((p) => [p.bullet?.text, p.lvl ?? 0])).toEqual([
      ['\u2022', 0],
      ['\u2013', 1],
      ['\u2022', 0],
    ])
    const pic = slides[1].els.find((e) => e.t === 'pic')!
    expect(pic.box).toMatchObject({ x: px(650), y: px(230), w: px(300), h: px(240) })
    expect(pic.t === 'pic' && pic.crop).toEqual([0.125, 0.1, 0.125, 0.1])
    expect(shapes(slides[1]).find((e) => textOfShape(e) === 'Field photo')!.box).toMatchObject({ x: px(700), y: px(490) })

    // table: header fill, banded body rows, no lines (an empty stroke sidecar); a tinted slide background
    expect(slides[2].bg).toEqual({ t: 'solid', c: '#f8fafc' })
    const tbl = slides[2].els.find((e) => e.t === 'table')!
    if (tbl.t !== 'table') throw new Error('no table')
    expect(tbl.rows.map((r) => r.cells.map((c) => c.text?.paras.map((p) => p.runs.map((x) => x.text).join('')).join('')))).toEqual([
      ['Region', 'Sites', 'Readings'],
      ['North', '12', '1480.5'],
      ['South', '9', '1102'],
      ['Total', '21', '2582.5'],
    ])
    expect(tbl.rows.map((r) => r.cells[0].fill)).toEqual([{ t: 'solid', c: '#1d4ed8' }, { t: 'none' }, { t: 'solid', c: '#e5e7eb' }, { t: 'none' }])
    expect(tbl.rows[1].cells[1].bt).toBeUndefined()

    // shapes: preset geometries and a curve; the group's children in its coordinates
    const kinds = shapes(slides[3]).slice(2)
    expect(kinds.map((e) => [e.geom, textOfShape(e)])).toEqual([
      ['roundRect', 'Plan'],
      ['custom', 'Build'],
      ['hexagon', 'Ship'],
      ['rightArrow', ''],
    ])
    expect(kinds.slice(0, 3).map((e) => Math.round((e.box.x * 3) / 4))).toEqual([190, 420, 650])
    expect(kinds[1].path).toContain('C')
  })

  test('saving gives a PowerPoint file', async () => {
    const { engine } = await openKey(await sampleKeynote())
    const pptx = (await engine.call('save', [])) as Blob
    const again = await openKey(new Uint8Array(await pptx.arrayBuffer()))
    expect(again.info.from).toBeUndefined()
    expect(again.info.pptx?.count).toBe(5)
    expect(textOfShape(shapes(again.slides[0])[1])).toBe('Quarterly Review')
  })

  test('rotation, groups placed on the slide, slide order of older files (a tree of slide nodes)', async () => {
    const b = new IworkBuilder()
    const style = shapeStyle(b, { fill: 'FF0000' })
    const a = shape(b, { frame: [100, 100, 50, 50], style, angle: 30 })
    // children already in slide coordinates (inside the group's frame)
    const g = b.add(K.GROUP, msg({ 1: drawable(200, 100, 300, 300), 2: [ref(shape(b, { frame: [300, 300, 80, 80], style })), ref(shape(b, { frame: [420, 320, 80, 80], style }))] }))
    const s1 = b.add(K.SLIDE, msg({ 7: [ref(a)] }))
    const s2 = b.add(K.SLIDE, msg({ 7: [ref(g)] }))
    const n1 = b.add(K.SLIDE_NODE, msg({ 2: ref(s1) }))
    const n2 = b.add(K.SLIDE_NODE, msg({ 2: ref(s2) }))
    const root = b.add(K.SLIDE_NODE, msg({ 1: [ref(n1), ref(n2)] }))
    const show = b.add(K.SHOW, msg({ 3: msg({ 1: ref(root) }), 4: msg({ 1: f32(800), 2: f32(600) }) }))
    b.add(K.DOCUMENT, msg({ 2: ref(show) }), 1)
    const { slides } = await openKey(await b.zip())
    expect(slides.length).toBe(2)
    // iWork turns counterclockwise, DrawingML clockwise
    expect(shapes(slides[0])[0].box.rot).toBe(330)
    expect(shapes(slides[1]).map((e) => Math.round((e.box.x * 3) / 4))).toEqual([300, 420])
  })
})

async function openSheet(bytes: Uint8Array<ArrayBuffer>) {
  const engine = new Engine(() => {})
  const info = await engine.open(new Blob([bytes]), { name: 'book.numbers' })
  const sheet = async (index: number) => {
    await engine.call('xlsx.load', [index])
    await new Promise((r) => setTimeout(r, 50))
    const block = (await engine.call('xlsx.block', [index, 0, 50])) as BlockOut
    return (r: number, c: number) => block.rows.find((x) => x.r === r)?.cells.find((x) => x[0] === c)
  }
  return { engine, info, sheet }
}

describe('numbers', () => {
  test('the sample spreadsheet: one worksheet per table, values with their formats, the table look', async () => {
    const { info, sheet } = await openSheet(await sampleNumbers())
    expect(info.kind).toBe('xlsx')
    expect(info.from).toBe('numbers')
    expect(info.xlsx!.sheets.map((s) => s.name)).toEqual(['Budget', 'Notes - Contacts', 'Notes - Totals'])
    const cell = await sheet(0)
    const row = (r: number) => Array.from({ length: 7 }, (_, c) => cell(r, c)?.[1] ?? '')
    expect(row(0)).toEqual(['Category', 'Planned', 'Actual', 'Share', 'Due', 'Paid', 'Time spent'])
    expect(row(2)).toEqual(['Groceries', '$450.00', '$512.40', '17.6%', 'Jan 31, 2025', '\u2611', '3h 30m'])
    expect(row(3)[5]).toBe('\u2610')
    // numbers and dates stay numbers (right-aligned, formatted by the workbook)
    expect(cell(2, 1)![3]).toBe(1)
    expect(cell(2, 4)![3]).toBe(1)
    // header fill and bold white text, banded body rows
    const style = (r: number, c: number) => info.xlsx!.styles[cell(r, c)![2]]
    expect(style(0, 0)).toMatchObject({ background: '#1d4ed8', fontWeight: 700, color: '#ffffff' })
    expect(style(2, 0).background).toBe('#eff6ff')
    expect(style(1, 0).background).toBeUndefined()

    const totals = await sheet(2)
    expect([totals(2, 0)?.[1], totals(2, 1)?.[1]]).toEqual(['Q2', '$3,120.50'])
  })

  test('saving gives an Excel workbook', async () => {
    const { engine } = await openSheet(await sampleNumbers())
    const xlsx = (await engine.call('save', [])) as Blob
    const again = await openSheet(new Uint8Array(await xlsx.arrayBuffer()))
    expect(again.info.from).toBeUndefined()
    expect(again.info.xlsx!.sheets.length).toBe(3)
    expect((await again.sheet(1))(1, 0)?.[1]).toBe('Ann Lee')
  })

  test('formats: ICU dates, durations, fractions, other bases, currencies', () => {
    expect(excelDate('EEEE, d MMMM yyyy h:mm a', true)).toBe('dddd, d mmmm yyyy h:mm am/pm')
    expect(excelDate("yyyy-MM-dd'T'HH:mm:ss")).toBe('yyyy-mm-dd"T"hh:mm:ss')
    expect(shownText('date', appleDate(2023, 4, 1) + 13 * 3600 + 25 * 60, new Msg(encode(fmt.date('d MMM yyyy h:mm a'))), true)).toBe('1 Apr 2023 1:25 pm')
    expect(durationText(5400, new Msg(encode(fmt.duration(1))))).toBe('1h 30m')
    expect(durationText(3725, new Msg(encode(msg({ 1: 268, 7: 0, 15: 4, 16: 16 }))))).toBe('1:02:05')
    expect(durationText(90061, new Msg(encode(msg({ 1: 268, 7: 2, 40: true }))))).toBe('1 day 1 hour 1 minute 1 second')
    expect(shownText('number', 1.25, new Msg(encode(msg({ 1: 262, 11: 4 }))))).toBe('1 1/4')
    expect(shownText('number', 0.333, new Msg(encode(msg({ 1: 262, 11: 0xffffffff }))))).toBe('1/3')
    expect(shownText('number', 255, new Msg(encode(msg({ 1: 269, 8: 16 }))))).toBe('FF')
    expect(shownText('number', -1234.5, new Msg(encode(fmt.currency('EUR'))))).toBe('\u20ac(1,234.50)')
    expect(shownText('number', 0.125, new Msg(encode(fmt.percent(1))))).toBe('12.5%')
    expect(excelFormat(new Msg(encode(fmt.decimal(2))), 'number')).toBe('#,##0.00')
  })
})
