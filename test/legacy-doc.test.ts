import { describe, expect, test } from 'bun:test'
import { readFileSync } from 'node:fs'
import type { Block, DocxMeta, Inline, Paragraph, Table } from '../src/core/docx/types'
import { Cfb } from '../src/core/legacy/cfb'
import { docToDocx } from '../src/core/legacy/doc'
import { dttm, hyperlink } from '../src/core/legacy/doc/fields'
import { Engine } from '../src/core/worker/engine'
import { ZipArchive } from '../src/core/zip'
import { BlobSource } from '../src/core/source'
import { deflateSync } from 'node:zlib'
import { imageBytes, readBlip } from '../src/core/legacy/doc/art'
import { Bytes } from '../src/core/legacy/doc/bin'
import { C, DocBuilder, P, S, STYLES, cfb, sprm } from './doc-builder'

const le32 = (v: number) => [v & 0xff, (v >> 8) & 0xff, (v >> 16) & 0xff, (v >>> 24) & 0xff]

const FIXTURES = `${import.meta.dir}/fixtures/legacy`

async function open(bytes: Uint8Array, name = 'doc.doc') {
  const events: { event: string; data: any }[] = []
  const engine = new Engine((event, data) => events.push({ event, data }))
  const info = await engine.open(new Blob([bytes as Uint8Array<ArrayBuffer>]), { name })
  await engine.call('docx.start', [])
  const done = events.find((e) => e.event === 'docx:done')?.data
  if (done?.error) throw new Error(done.error)
  const blocks: Block[] = events.filter((e) => e.event === 'docx:blocks').flatMap((e) => e.data)
  const meta: DocxMeta = events.filter((e) => e.event === 'docx:meta').at(-1)?.data
  return { engine, info, blocks, meta, paras: blocks.filter((b): b is Paragraph => b.t === 'p') }
}

/** the parts of the converted DOCX */
async function parts(bytes: Uint8Array) {
  const blob = await docToDocx(Cfb.open(bytes))
  const zip = await ZipArchive.open(new BlobSource(blob))
  const text = async (p: string) => (zip.has(p) ? new TextDecoder().decode(await zip.bytes(p)) : undefined)
  return { zip, text, document: (await text('word/document.xml'))! }
}

const textOf = (p: Paragraph) => p.runs.map((r) => (r.t === 'text' ? r.text : r.t === 'field' ? `{${r.kind}}` : r.t === 'ref' ? `^${r.text}` : r.t === 'img' ? '[img]' : '')).join('')
const cellText = (t: Table, r: number, c: number) => t.rows[r].cells[c].blocks.map((b) => (b.t === 'p' ? textOf(b) : '[table]')).join('\n')
const runOf = (p: Paragraph, text: string) => p.runs.find((r): r is Extract<Inline, { t: 'text' }> => r.t === 'text' && r.text.includes(text))
/** a 2x2 red PNG */
const TINY_PNG = Uint8Array.from(atob('iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAIAAAD91JpzAAAAFklEQVR4nGP8z8DAwMDAxMDAwMDAAAANHQEDasKb6QAAAABJRU5ErkJggg=='), (c) => c.charCodeAt(0))

describe('Word 97 binary → DOCX', () => {
  test('text, character formatting and paragraph properties', async () => {
    const d = new DocBuilder()
    d.p([{ text: 'Plain ' }, { text: 'bold', c: C.bold() }, { text: ' ' }, { text: 'italic', c: C.italic() }, { text: ' ' }, { text: 'under', c: C.underline(3) }, { text: ' ' }, { text: 'red', c: C.color('C00000') }, { text: ' ' }, { text: 'big', c: C.size(20) }, { text: ' ' }, { text: 'mono', c: C.font(3) }, { text: ' ' }, { text: 'mark', c: C.highlight(7) }, { text: ' x' }, { text: '2', c: C.iss(1) }, { text: ' ' }, { text: 'gone', c: C.strike() }, { text: ' ' }, { text: 'caps', c: C.smallCaps() }])
    d.p('Centered', 0, P.jc(1))
    d.p('Indented', 0, [...P.left(720), ...P.first(-360), ...P.right(360)])
    d.p('Spaced', 0, [...P.before(240), ...P.after(480), ...P.line(360)])
    d.p('Exact', 0, P.line(-300, false))
    d.p('Kept', 0, [...P.keep(), ...P.keepNext(), ...P.pageBreakBefore()])
    d.p('Boxed', 0, [...P.border('top', [8, 1, 6, 0]), ...P.shd('FFF2CC')])
    d.p('Tabs\there', 0, P.tabs([[2880, 2, 1]]))
    const { paras } = await open(d.build())
    const p0 = paras[0]
    expect(textOf(p0)).toBe('Plain bold italic under red big mono mark x2 gone caps')
    expect(runOf(p0, 'bold')!.s.b).toBe(true)
    expect(runOf(p0, 'italic')!.s.i).toBe(true)
    expect(runOf(p0, 'under')!.s.u).toBe('double')
    expect(runOf(p0, 'red')!.s.color).toBe('#c00000')
    expect(runOf(p0, 'big')!.s.sz).toBe(20)
    expect(runOf(p0, 'mono')!.s.font).toContain('Courier New')
    expect(runOf(p0, 'mark')!.s.hl).toBeTruthy()
    expect(runOf(p0, '2')!.s.va).toBe('sup')
    expect(runOf(p0, 'gone')!.s.strike).toBe(true)
    expect(runOf(p0, 'caps')!.s.smallCaps).toBe(true)
    // the Normal style: 11pt Arial, 6pt after
    expect(runOf(p0, 'Plain')!.s.sz).toBe(11)
    expect(runOf(p0, 'Plain')!.s.font).toContain('Arial')
    expect(paras[1].ps.align).toBe('center')
    expect(paras[2].ps.indL).toBe(36)
    expect(paras[2].ps.indR).toBe(18)
    expect(paras[2].ps.indFirst).toBe(-18)
    expect(paras[3].ps.before).toBe(12)
    expect(paras[3].ps.after).toBe(24)
    expect(paras[3].ps.line).toBe(1.5)
    expect(paras[4].ps.lineRule).toBe('exact')
    expect(paras[4].ps.line).toBe(15)
    expect(paras[5].pb).toBe(true)
    expect(paras[6].ps.bTop?.color).toBe('#ff0000')
    expect(paras[6].ps.shd).toBe('#fff2cc')
    expect(textOf(paras[7])).toBe('Tabs\there')
  })

  test('styles: built-in names and ids, based-on chains, character styles', async () => {
    const d = new DocBuilder()
    d.p('Report', STYLES.title)
    d.p('Chapter', STYLES.heading1)
    d.p('Section', STYLES.heading2)
    d.p([{ text: 'Some ' }, { text: 'strong', c: C.style(STYLES.strong) }, { text: ' words.' }])
    const bytes = d.build()
    const { paras } = await open(bytes)
    expect(paras.map((p) => p.sid)).toEqual(['Title', 'Heading1', 'Heading2', 'Normal'])
    expect(runOf(paras[1], 'Chapter')!.s.b).toBe(true)
    expect(runOf(paras[1], 'Chapter')!.s.sz).toBe(16)
    // based on Normal: Arial from Normal, color from the heading
    expect(runOf(paras[1], 'Chapter')!.s.font).toContain('Arial')
    expect(runOf(paras[1], 'Chapter')!.s.color).toBe('#1f3864')
    expect(runOf(paras[3], 'strong')!.s.b).toBe(true)
    const { text } = await parts(bytes)
    const styles = (await text('word/styles.xml'))!
    expect(styles).toContain('<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/>')
    expect(styles).toContain('w:styleId="Strong"')
    expect(styles).toContain('<w:outlineLvl w:val="0"/>')
  })

  test('lists: bullets and numbers from the list tables', async () => {
    const d = new DocBuilder()
    const bullets = d.list([{ nfc: 23, text: '\uf0b7', left: 720, hang: 360, font: 1 }])
    const numbers = d.list([
      { nfc: 0, text: '\u0000.', nums: [1], left: 720, hang: 360 },
      { nfc: 4, text: '\u0001)', nums: [1], left: 1440, hang: 360 },
    ])
    d.p('Apples', STYLES.listBullet, P.list(bullets))
    d.p('Pears', STYLES.listBullet, P.list(bullets))
    d.p('First', STYLES.listNumber, P.list(numbers))
    d.p('Nested', STYLES.listNumber, P.list(numbers, 1))
    d.p('Second', STYLES.listNumber, P.list(numbers))
    const { paras } = await open(d.build())
    expect(paras.map((p) => p.marker?.text.trim())).toEqual(['•', '•', '1.', 'a)', '2.'])
    expect(paras[0].lst).toBe('bullet')
    expect(paras[2].lst).toBe('number')
    expect(paras[2].ps.indL).toBe(36)
    expect(paras[3].ps.indL).toBe(72)
    expect(paras[3].ilvl).toBe(1)
  })

  test('tables: widths, merged cells, header row, shading, borders, nested table', async () => {
    const d = new DocBuilder()
    d.p('Before')
    d.table({
      widths: [2000, 3000, 1500],
      header: true,
      rows: [
        [{ paras: ['Region'], shd: 'D9E2F3' }, { paras: ['Sales'], shd: 'D9E2F3' }, { paras: ['Share'], shd: 'D9E2F3' }],
        [{ paras: ['North'], vm: 3 }, { paras: ['120', 'units'] }, { paras: ['60%'] }],
        [{ paras: [''], vm: 1 }, { paras: ['80'] }, { paras: ['40%'] }],
        [{ paras: ['Total'], hm: 1 }, { paras: [''], hm: 2 }, { paras: ['200'], table: { widths: [700, 700], rows: [[{ paras: ['a'] }, { paras: ['b'] }]] } }],
      ],
    })
    d.p('After')
    const { blocks } = await open(d.build())
    const t = blocks.find((b): b is Table => b.t === 'tbl')!
    expect(t.cols).toEqual([100, 150, 75])
    expect(t.rows).toHaveLength(4)
    expect(t.rows[0].header).toBe(true)
    expect(cellText(t, 0, 1)).toBe('Sales')
    expect(t.rows[0].cells[0].shd).toBe('#d9e2f3')
    expect(cellText(t, 1, 1)).toBe('120\nunits')
    expect(t.rows[1].cells[0].rowSpan).toBe(2)
    expect(t.rows[2].cells).toHaveLength(2)
    expect(t.rows[3].cells[0].colSpan).toBe(2)
    expect(cellText(t, 3, 0)).toBe('Total')
    const nested = t.rows[3].cells[1].blocks.find((b): b is Table => b.t === 'tbl')!
    expect(nested.rows[0].cells.map((c) => c.blocks.map((b) => (b.t === 'p' ? textOf(b) : '')).join(''))).toEqual(['a', 'b'])
    expect(t.rows[0].cells[0].borders.t).toContain('solid')
    const texts = blocks.filter((b): b is Paragraph => b.t === 'p').map(textOf)
    expect(texts).toEqual(['Before', 'After'])
  })

  test('fields: hyperlinks, page numbers, other fields keep their result', async () => {
    const d = new DocBuilder()
    d.p([{ text: 'See ' }, ...d.hyperlink('https://example.com/docs', 'the docs'), { text: '.' }])
    d.p([{ text: 'Page ' }, ...d.field('PAGE', '1'), { text: ' of ' }, ...d.field('NUMPAGES', '3')])
    d.p([{ text: 'Today: ' }, ...d.field('DATE \\@ "d MMMM yyyy"', '10 November 2026')])
    const bytes = d.build()
    const { paras } = await open(bytes)
    const link = runOf(paras[0], 'the docs')!
    expect(link.href).toBe('https://example.com/docs')
    expect(link.s.u).toBe('single')
    expect(textOf(paras[0])).toBe('See the docs.')
    expect(textOf(paras[1])).toBe('Page {PAGE} of {NUMPAGES}')
    expect(textOf(paras[2])).toBe('Today: 10 November 2026')
    const { document } = await parts(bytes)
    expect(document).toContain('<w:hyperlink r:id=')
    expect(document).toContain('<w:instrText xml:space="preserve"> DATE \\@ "d MMMM yyyy" </w:instrText>')
  })

  test('sections: page size, orientation, margins, columns, breaks', async () => {
    const d = new DocBuilder()
    d.p('Portrait page')
    d.sectionBreak([...S.page(12240, 15840), ...S.margins(1440, 1080, 1440, 1080)])
    d.p('Landscape page')
    d.sectionBreak([...S.bkc(2), ...S.page(15840, 12240), ...S.landscape(), ...S.margins(720, 720, 720, 720)])
    d.p('Two columns')
    const bytes = d.build([...S.bkc(0), ...S.page(11906, 16838), ...S.cols(2, 360)])
    const { blocks, meta } = await open(bytes)
    const sects = blocks.filter((b) => b.t === 'sect') as Extract<Block, { t: 'sect' }>[]
    expect(sects).toHaveLength(2)
    expect(sects[0].page.w).toBe(612)
    expect(sects[0].page.ml).toBe(54)
    expect(sects[1].page.w).toBe(792)
    expect(sects[1].page.h).toBe(612)
    expect(meta.page.w).toBeCloseTo(595.3, 1)
    expect(meta.page.cols).toBe(2)
    const { document } = await parts(bytes)
    expect(document).toContain('w:orient="landscape"')
    expect(document).toContain('<w:type w:val="continuous"/>')
  })

  test('headers and footers, with a page field', async () => {
    const d = new DocBuilder()
    d.p('Body text')
    d.headers([undefined, 'Company report', undefined, [{ text: 'Page ' }, ...d.field('PAGE', '1')], undefined, undefined])
    const { meta } = await open(d.build())
    const h = meta.header!.filter((b): b is Paragraph => b.t === 'p')
    const f = meta.footer!.filter((b): b is Paragraph => b.t === 'p')
    expect(h.map(textOf)).toEqual(['Company report'])
    expect(h[0].sid).toBe('Header')
    expect(f.map(textOf)).toEqual(['Page {PAGE}'])
  })

  test('footnotes', async () => {
    const d = new DocBuilder()
    d.p([{ text: 'A claim' }, d.footnote('The source.'), { text: ' and another' }, d.footnote('Second source.')])
    const bytes = d.build()
    const { paras } = await open(bytes)
    expect(textOf(paras[0])).toBe('A claim^1 and another^2')
    const { text } = await parts(bytes)
    const notes = (await text('word/footnotes.xml'))!
    expect(notes).toContain('<w:footnote w:id="1"><w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr>')
    expect(notes).toContain('<w:footnoteRef/>')
    expect(notes).toContain('Second source.')
    expect(notes).toContain('w:type="separator" w:id="-1"')
  })

  test('comments, with and without a commented range', async () => {
    const d = new DocBuilder()
    d.p([{ text: 'Revenue grew' }, d.comment('Check this figure.', 'Alice Smith', 'AS', 12), { text: ' this quarter.' }, d.comment('Nice.', 'Bob', 'B')])
    const bytes = d.build()
    const { paras } = await open(bytes)
    expect(textOf(paras[0])).toBe('Revenue grew this quarter.')
    const { text, document } = await parts(bytes)
    const xml = (await text('word/comments.xml'))!
    expect(xml).toContain('<w:comment w:id="0" w:author="Alice Smith" w:initials="AS">')
    expect(xml).toContain('Check this figure.')
    expect(xml).toContain('<w:comment w:id="1" w:author="Bob" w:initials="B">')
    expect(xml).toContain('<w:annotationRef/>')
    expect(document).toMatch(/<w:commentRangeStart w:id="0"\/>.*Revenue grew.*<w:commentRangeEnd w:id="0"\/>.*<w:commentReference w:id="0"\/>/)
    expect(document).not.toContain('<w:commentRangeStart w:id="1"/>')
    expect(document).toContain('<w:commentReference w:id="1"/>')
  })

  test('special characters', async () => {
    const d = new DocBuilder()
    d.p('line\u000bbreak\ttab non\u00a0breaking hard\u001ehyphen soft\u001fhyphen')
    d.p('before page\u000cafter page')
    d.p([{ text: 'symbol ' }, { text: '(', c: [...C.spec(), ...C.symbol(4, 0xf0fc)] }])
    const bytes = d.build()
    const { paras } = await open(bytes)
    expect(textOf(paras[0])).toBe('line\nbreak\ttab non\u00a0breaking hard\u2011hyphen soft\u00adhyphen')
    expect(textOf(paras[1])).toBe('before page')
    expect(paras[2].pb).toBe(true)
    expect(textOf(paras[2])).toBe('after page')
    const { document } = await parts(bytes)
    expect(document).toContain('<w:sym w:font="Wingdings" w:char="F0FC"/>')
  })

  test('pictures: inline from the Data stream, floating from the drawing', async () => {
    const d = new DocBuilder()
    d.p([{ text: 'Inline: ' }, d.picture(TINY_PNG, 1440, 720)])
    d.p([d.floating(TINY_PNG, [0, 0, 2880, 1440]), { text: 'Floating picture' }])
    const bytes = d.build()
    const { paras, engine } = await open(bytes)
    const img = paras[0].runs.find((r) => r.t === 'img') as Extract<Inline, { t: 'img' }>
    expect(img.w).toBe(72)
    expect(img.h).toBe(36)
    const png = await engine.call('part', [img.target])
    expect(new Uint8Array(await (png as Blob).arrayBuffer()).slice(0, 4)).toEqual(Uint8Array.of(0x89, 0x50, 0x4e, 0x47))
    const fl = paras[1].runs.find((r) => r.t === 'img') as Extract<Inline, { t: 'img' }>
    expect(fl.w).toBe(144)
    expect(fl.h).toBe(72)
    const { document } = await parts(bytes)
    expect(document).toContain('<wp:anchor ')
    expect(document).toContain('<wp:positionV relativeFrom="paragraph">')
  })

  test('the piece table: 8-bit and UTF-16 pieces, a piece sprm', async () => {
    const d = new DocBuilder()
    d.splitPieces = true
    d.p('Ünïcödé first half')
    d.p('ascii second half')
    const { paras } = await open(d.build())
    expect(paras.map(textOf)).toEqual(['Ünïcödé first half', 'ascii second half'])
    // the 8-bit piece carries sprmCFBold
    expect(runOf(paras[1], 'second')!.s.b).toBe(true)
    expect(runOf(paras[0], 'first')!.s.b).toBeFalsy()
  })

  test('tracked changes', async () => {
    const d = new DocBuilder()
    d.authors = ['Alice', 'Bob']
    d.p([{ text: 'kept ' }, { text: 'added', c: C.rmIns(0) }, { text: ' ' }, { text: 'removed', c: C.rmDel(1) }])
    const { paras } = await open(d.build())
    expect(runOf(paras[0], 'added')!.tc).toMatchObject({ k: 'ins', author: 'Alice' })
    expect(runOf(paras[0], 'removed')!.tc).toMatchObject({ k: 'del', author: 'Bob' })
  })

  test('the result is a valid package: content types, relationships', async () => {
    const d = new DocBuilder()
    d.p([{ text: 'x' }, d.picture(TINY_PNG, 300, 300), d.footnote('n'), ...d.hyperlink('https://example.com', 'l')])
    const { zip, text } = await parts(d.build())
    const ct = (await text('[Content_Types].xml'))!
    for (const p of ['word/document.xml', 'word/styles.xml', 'word/settings.xml', 'word/footnotes.xml', 'word/fontTable.xml']) {
      expect(zip.has(p)).toBe(true)
      expect(ct).toContain(`PartName="/${p}"`)
    }
    expect(ct).toContain('<Default Extension="png" ContentType="image/png"/>')
    const rels = (await text('word/_rels/document.xml.rels'))!
    expect(rels).toContain('Target="media/image1.png"')
    expect(rels).toContain('TargetMode="External"')
    expect(rels).toContain('relationships/footnotes')
  })

  test('unsupported files fail with a clear error', async () => {
    const enc = new DocBuilder()
    enc.encrypted = true
    enc.p('secret')
    await expect(docToDocx(Cfb.open(enc.build()))).rejects.toThrow('Password-protected Word documents are not supported')
    const old = new DocBuilder()
    old.nFib = 0x68
    old.p('old')
    await expect(docToDocx(Cfb.open(old.build()))).rejects.toThrow('Word 6.0 / Word 95 and older documents are not supported')
    await expect(docToDocx(Cfb.open(cfb([['WordDocument', new Uint8Array(100)]])))).rejects.toThrow()
    await expect(docToDocx(Cfb.open(cfb([['Other', new Uint8Array(10)]])))).rejects.toThrow('no WordDocument stream')
  })

  test('damaged files never hang or crash the worker', async () => {
    const d = new DocBuilder()
    d.p('Heading', STYLES.heading1)
    d.p([{ text: 'Body ' }, ...d.hyperlink('https://example.com', 'link'), d.footnote('note'), d.picture(TINY_PNG, 200, 200)])
    d.table({ widths: [1000, 1000], rows: [[{ paras: ['a'] }, { paras: ['b'] }]] })
    const good = d.build()
    const c0 = Cfb.open(good)
    const wd = c0.stream('WordDocument')!
    const tbl = c0.stream('1Table')!
    let seed = 7
    const rnd = () => ((seed = (seed * 1103515245 + 12345) >>> 0) / 2 ** 32)
    for (let k = 0; k < 150; k++) {
      const w = wd.slice()
      const t = tbl.slice()
      const target = k % 2 ? w : t
      const lo = k % 2 ? 0x20 : 0
      for (let n = 0; n < 12; n++) target[lo + Math.floor(rnd() * (target.length - lo))] = Math.floor(rnd() * 256)
      const bytes = cfb([['WordDocument', w], ['1Table', t], ['Data', c0.stream('Data')!]])
      const t0 = performance.now()
      try {
        await docToDocx(Cfb.open(bytes))
      } catch (e) {
        expect(e).toBeInstanceOf(Error)
      }
      expect(performance.now() - t0).toBeLessThan(2000)
    }
  })

  test('frames: absolutely positioned paragraphs', async () => {
    const d = new DocBuilder()
    d.p('Framed', 0, [...sprm(0x261b, (1 << 4) | (2 << 6)), ...sprm(0x8418, 2880), ...sprm(0x8419, -4), ...sprm(0x841a, 3000), ...sprm(0x2423, 2)])
    const { document } = await parts(d.build())
    expect(document).toContain('<w:framePr w:w="3000" w:wrap="around" w:vAnchor="page" w:hAnchor="page" w:x="2880" w:yAlign="top"/>')
  })

  test('pictures: device-independent bitmaps become BMP files, Windows metafiles get a placeable header', async () => {
    const rec = (inst: number, type: number, body: number[]) => Uint8Array.from([inst << 4 & 0xff, inst >> 4, type & 0xff, type >> 8, ...le32(body.length), ...body])
    // a 1x1 24-bit DIB: BITMAPINFOHEADER (40 bytes) and one padded pixel row
    const dib = [...le32(40), ...le32(1), ...le32(1), 1, 0, 24, 0, ...le32(0), ...le32(4), ...le32(0), ...le32(0), ...le32(0), ...le32(0), 0x00, 0x00, 0xff, 0]
    const bmp = readBlip(new Bytes(rec(0x7a8, 0xf01f, [...new Array(16).fill(0), 0xff, ...dib])), 0)!
    expect(bmp.ext).toBe('bmp')
    expect(String.fromCharCode(bmp.data[0], bmp.data[1])).toBe('BM')
    expect(new DataView(bmp.data.buffer).getUint32(10, true)).toBe(54)
    // a deflated WMF: inflated, with a placeable header from the metafile header's bounds and size
    const wmf = Uint8Array.from([1, 0, 9, 0, 0, 3, 0x20, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0])
    const z = deflateSync(wmf)
    const head = [...le32(wmf.length), ...le32(0), ...le32(0), ...le32(1440), ...le32(720), ...le32(914400), ...le32(457200), ...le32(z.length), 0, 0xfe]
    const blip = readBlip(new Bytes(rec(0x216, 0xf01b, [...new Array(16).fill(0), ...head, ...z])), 0)!
    expect(blip).toMatchObject({ ext: 'wmf', deflated: true })
    const bytes = await imageBytes(blip)
    expect(new DataView(bytes.buffer, bytes.byteOffset).getUint32(0, true)).toBe(0x9ac6cdd7)
    expect(new DataView(bytes.buffer, bytes.byteOffset).getUint16(14, true)).toBe(1440)
    expect(bytes.subarray(22)).toEqual(wmf)
  })

  test('helpers: HYPERLINK instructions and DTTM dates', () => {
    expect(hyperlink(' HYPERLINK "https://a.b/c?d=1" ')).toEqual({ url: 'https://a.b/c?d=1' })
    expect(hyperlink('HYPERLINK \\l "_Toc1" \\o "tip"')).toEqual({ anchor: '_Toc1', tip: 'tip' })
    expect(hyperlink('HYPERLINK "C:\\\\docs\\\\a.doc"')).toEqual({ url: 'C:\\docs\\a.doc' })
    expect(hyperlink('PAGE')).toBeUndefined()
    // 2010-03-08 15:57
    expect(dttm(57 | (15 << 6) | (8 << 11) | (3 << 16) | (110 << 20))).toBe('2010-03-08T15:57:00Z')
    expect(dttm(0)).toBeUndefined()
  })

  test('sprm builder sizes match the reader', () => {
    expect(sprm(0x0835, 1)).toEqual([0x35, 0x08, 1])
    expect(sprm(0x4a43, 24)).toEqual([0x43, 0x4a, 24, 0])
    expect(sprm(0xd608, [1, 2])).toEqual([0x08, 0xd6, 3, 0, 1, 2])
  })
})

describe('Word 97 fixtures', () => {
  test('sample.doc: headings, formatting, lists, table, picture, footnote, header and footer', async () => {
    const { paras, blocks, meta, info } = await open(new Uint8Array(readFileSync(`${FIXTURES}/sample.doc`)), 'sample.doc')
    expect(info.kind).toBe('docx')
    expect((info as { converted?: string }).converted).toBe('doc')
    expect(paras[0].sid).toBe('Title')
    expect(textOf(paras[0])).toBe('Quarterly Report')
    expect(paras.some((p) => p.sid === 'Heading1' && textOf(p) === 'Highlights')).toBe(true)
    expect(paras.filter((p) => p.lst === 'bullet').length).toBeGreaterThanOrEqual(3)
    expect(paras.filter((p) => p.lst === 'number').map((p) => p.marker?.text.trim())).toEqual(['1.', '2.', '3.'])
    const t = blocks.find((b): b is Table => b.t === 'tbl')!
    expect(cellText(t, 0, 0)).toBe('Region')
    expect(t.rows.at(-1)!.cells[0].colSpan).toBe(3)
    expect(paras.some((p) => p.runs.some((r) => r.t === 'img'))).toBe(true)
    expect(paras.some((p) => p.runs.some((r) => r.t === 'text' && r.href === 'https://example.com/'))).toBe(true)
    expect(paras.some((p) => p.runs.some((r) => r.t === 'ref'))).toBe(true)
    expect(meta.header?.some((b) => b.t === 'p' && textOf(b).includes('Quarterly Report'))).toBe(true)
    expect(meta.footer?.some((b) => b.t === 'p' && textOf(b).includes('{PAGE}'))).toBe(true)
  })

  test('textutil.doc (written by macOS textutil): 8-bit text, formatting, table', async () => {
    const { paras, blocks } = await open(new Uint8Array(readFileSync(`${FIXTURES}/textutil.doc`)), 'textutil.doc')
    expect(textOf(paras[0])).toBe('Quarterly Report')
    expect(runOf(paras[0], 'Quarterly')!.s.b).toBe(true)
    const p1 = paras[1]
    expect(runOf(p1, 'bold')!.s.b).toBe(true)
    expect(runOf(p1, 'italic')!.s.i).toBe(true)
    expect(runOf(p1, 'red')!.s.color).toBe('#b00004')
    expect(runOf(p1, 'monospace')!.s.font).toContain('Courier')
    expect(paras.some((p) => p.ps.align === 'center')).toBe(true)
    const t = blocks.find((b): b is Table => b.t === 'tbl')!
    expect(cellText(t, 1, 0)).toBe('North')
    expect(cellText(t, 2, 1)).toBe('95')
  })
})
