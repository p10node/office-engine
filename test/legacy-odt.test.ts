import { describe, expect, test } from 'bun:test'
import { readFileSync } from 'node:fs'
import type { DocComment } from '../src/core/comments'
import type { Block, DocxMeta, Inline, Paragraph, Table } from '../src/core/docx/types'
import { convertFlatOdf, convertOdfZip } from '../src/core/legacy'
import { BlobSource } from '../src/core/source'
import { Engine } from '../src/core/worker/engine'
import { ZipArchive } from '../src/core/zip'
import { writeZip } from '../src/core/zipwrite'

const fixture = (name: string) => new Uint8Array(readFileSync(new URL(`./fixtures/legacy/${name}`, import.meta.url)))

async function open(bytes: Uint8Array, name = 'doc.odt') {
  const events: { event: string; data: any }[] = []
  const engine = new Engine((event, data) => events.push({ event, data }))
  const info = await engine.open(new Blob([bytes as BlobPart]), { name })
  await engine.call('docx.start', [])
  const done = events.find((e) => e.event === 'docx:done')?.data
  if (done?.error) throw new Error(done.error)
  const blocks: Block[] = events.filter((e) => e.event === 'docx:blocks').flatMap((e) => e.data)
  const meta: DocxMeta = events.filter((e) => e.event === 'docx:meta').at(-1)?.data
  return { engine, info, blocks, meta, paras: blocks.filter((b): b is Paragraph => b.t === 'p'), tables: blocks.filter((b): b is Table => b.t === 'tbl') }
}

const runText = (r: Inline) => (r.t === 'text' ? r.text : r.t === 'field' ? `{${r.kind}}` : r.t === 'ref' ? `[${r.text}]` : r.t === 'img' ? '[img]' : r.t === 'box' ? '[box]' : '')
const textOf = (p: Paragraph) => p.runs.map(runText).join('')
const cellText = (t: Table, r: number, c: number) => t.rows[r].cells[c].blocks.map((b) => (b.t === 'p' ? textOf(b) : '[table]')).join('\n')
type TextRun = Extract<Inline, { t: 'text' }>
/** the run whose text is `text` (else the first holding it) */
const run = (p: Paragraph, text: string) =>
  p.runs.find((r): r is TextRun => r.t === 'text' && r.text === text) ?? p.runs.find((r): r is TextRun => r.t === 'text' && r.text.includes(text))

// ---------------------------------------------------------------- a small OpenDocument text builder

const NS = [
  'office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"',
  'style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"',
  'text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"',
  'table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"',
  'draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"',
  'fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"',
  'svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"',
  'xlink="http://www.w3.org/1999/xlink"',
  'loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0"',
]
  .map((n) => 'xmlns:' + n)
  .join(' ')

interface Parts {
  /** office:text content */
  body: string
  /** office:styles of styles.xml */
  styles?: string
  /** office:automatic-styles of content.xml */
  auto?: string
  /** office:automatic-styles of styles.xml (page layouts, header styles) */
  masterAuto?: string
  /** office:master-styles */
  master?: string
  files?: Record<string, Uint8Array>
}

async function odt(p: Parts, mime = 'application/vnd.oasis.opendocument.text'): Promise<Uint8Array> {
  const files = new Map<string, string | Uint8Array>([
    ['mimetype', mime],
    ['content.xml', `<?xml version="1.0" encoding="UTF-8"?><office:document-content ${NS} office:version="1.3"><office:automatic-styles>${p.auto ?? ''}</office:automatic-styles><office:body><office:text>${p.body}</office:text></office:body></office:document-content>`],
    ['styles.xml', `<?xml version="1.0" encoding="UTF-8"?><office:document-styles ${NS} office:version="1.3"><office:styles>${p.styles ?? ''}</office:styles><office:automatic-styles>${p.masterAuto ?? ''}</office:automatic-styles><office:master-styles>${p.master ?? ''}</office:master-styles></office:document-styles>`],
    ...Object.entries(p.files ?? {}),
  ])
  return new Uint8Array(await (await writeZip(files, mime)).arrayBuffer())
}

const flat = (p: Parts) =>
  new TextEncoder().encode(
    `<?xml version="1.0" encoding="UTF-8"?><office:document ${NS} office:version="1.3" office:mimetype="application/vnd.oasis.opendocument.text"><office:styles>${p.styles ?? ''}</office:styles>` +
      `<office:automatic-styles>${p.auto ?? ''}${p.masterAuto ?? ''}</office:automatic-styles><office:master-styles>${p.master ?? ''}</office:master-styles><office:body><office:text>${p.body}</office:text></office:body></office:document>`,
  )

/** the converted DOCX package's parts */
async function docxParts(bytes: Uint8Array): Promise<(path: string) => Promise<string>> {
  const conv = await convertOdfZip(await ZipArchive.open(new BlobSource(new Blob([bytes as BlobPart]))))
  const zip = await ZipArchive.open(new BlobSource(conv!.file))
  return async (path) => (zip.has(path) ? zip.text(path) : '')
}

const style = (name: string, family: string, props: string, extra = '') => `<style:style style:name="${name}" style:family="${family}"${extra}>${props}</style:style>`
const tp = (a: string) => `<style:text-properties ${a}/>`
const pp = (a: string, kids = '') => (kids ? `<style:paragraph-properties ${a}>${kids}</style:paragraph-properties>` : `<style:paragraph-properties ${a}/>`)

const PNG = Uint8Array.from(atob('iVBORw0KGgoAAAANSUhEUgAAAAIAAAABCAYAAAD0In+KAAAAEUlEQVR4nGP4z8DwnwGIGBgAHvQD/X6mH2EAAAAASUVORK5CYII='), (c) => c.charCodeAt(0))

// ---------------------------------------------------------------- fixtures

describe('OpenDocument text fixtures', () => {
  test('sample.odt: styles, outline numbering, lists, table, picture, frame, notes, header and footer', async () => {
    const { info, paras, tables, blocks, meta } = await open(fixture('sample.odt'))
    expect(info.kind).toBe('docx')
    expect(info.converted).toBe('odt')

    const title = paras[0]
    expect(textOf(title)).toBe('Quarterly Report')
    expect(title.sid).toBe('Title')
    expect(title.ps.align).toBe('center')
    expect(paras.find((p) => textOf(p) === 'Prepared for the planning team')?.sid).toBe('Subtitle')
    // the table of contents keeps its text
    expect(paras.find((p) => textOf(p).startsWith('1. Highlights'))?.sid).toBe('TOC1')

    // headings: Word's heading styles, numbered from the outline style
    const h1 = paras.find((p) => textOf(p) === 'Highlights')!
    expect(h1.sid).toBe('Heading1')
    expect(h1.marker?.text).toBe('1. ')
    const results = paras.find((p) => textOf(p) === 'Results')!
    expect(results.marker?.text).toBe('2. ')
    expect(paras.find((p) => textOf(p) === 'Regions')?.marker?.text).toBe('2.1 ')
    // a heading in a paragraph style with a page break
    const appendix = paras.find((p) => textOf(p) === 'Appendix')!
    expect(appendix.sid).toBe('Heading1')
    expect(appendix.pb).toBe(true)

    // character formatting from automatic styles
    const body = paras.find((p) => textOf(p).startsWith('This document shows'))!
    expect(body.sid).toBe('BodyText')
    expect(body.ps.align).toBe('justify')
    expect(run(body, 'bold')!.s.b).toBe(true)
    expect(run(body, 'italic')!.s.i).toBe(true)
    expect(run(body, 'underlined red')!.s).toMatchObject({ u: 'single', color: '#c00000' })
    expect(run(body, 'struck out')!.s.strike).toBe(true)
    expect(run(body, 'highlighted')!.s.shd).toBe('#ffff00')
    expect(run(body, 'large')!.s.sz).toBe(16)
    expect(run(body, 'Small Caps')!.s.smallCaps).toBe(true)
    expect(run(body, 'code()')!.s.font).toContain('Liberation Mono')
    const sup = body.runs.find((r) => r.t === 'text' && r.text === '2' && r.s.va === 'sup')
    expect(sup).toBeDefined()
    expect(body.runs.some((r) => r.t === 'ref')).toBe(true)
    expect(run(body, 'example.com')!.href).toBe('https://example.com/report')
    expect(textOf(paras.find((p) => textOf(p).startsWith('Spacing is kept'))!)).toBe('Spacing is kept:   three spaces, a\ttab and a\nline break.')

    // lists: bullets, numbers, nested levels, a list that goes on after a pause
    const marker = (t: string) => paras.find((p) => textOf(p) === t)?.marker?.text
    expect(marker('Revenue grew 24% year over year')).toBe('•')
    expect(marker('North')).toBe('◦')
    expect(marker('Hire the platform team')).toBe('1.')
    expect(marker('Text documents')).toBe('a)')
    expect(marker('Spreadsheets')).toBe('b)')
    expect(marker('Expand to new markets')).toBe('3.')
    expect(marker('Review the budget')).toBe('4.')

    // the table: a header row, a cell spanning two rows, one spanning three columns, shading
    const t = tables[0]
    expect(t.cols.length).toBe(4)
    expect(t.rows.length).toBe(5)
    expect(t.rows[0].header).toBe(true)
    expect(cellText(t, 0, 0)).toBe('Region')
    expect(t.rows[0].cells[0].shd).toBe('#1f3864')
    expect(t.rows[1].cells[0].rowSpan).toBe(2)
    expect(cellText(t, 1, 0)).toBe('North and South')
    expect(t.rows[2].cells.map((c) => c.blocks.map((b) => (b.t === 'p' ? textOf(b) : '')).join(''))).toEqual(['98', '110', '208'])
    expect(t.rows[4].cells[0].colSpan).toBe(3)
    expect(t.rows[1].cells[0].borders.t).toContain('solid')
    expect(t.rows[1].cells[0].vAlign).toBe('middle')
    expect((t.rows[1].cells[1].blocks[0] as Paragraph).ps.align).toBe('right')
    // 17 cm wide
    expect(t.cols.reduce((a, b) => a + b, 0)).toBeCloseTo(481.9, 0)

    // the picture (12 cm × 6 cm) and the text frame
    const pic = paras.flatMap((p) => p.runs).find((r) => r.t === 'img')
    expect(pic).toMatchObject({ t: 'img' })
    if (pic?.t === 'img') {
      expect(pic.w).toBeCloseTo(340.2, 0)
      expect(pic.h).toBeCloseTo(170.1, 0)
    }
    const box = paras.flatMap((p) => p.runs).find((r) => r.t === 'box')
    expect(box?.t === 'box' && box.blocks.map((b) => (b.t === 'p' ? textOf(b) : '')).join('')).toBe('Note: figures are unaudited.')
    expect(box?.t === 'box' && box.fill?.toLowerCase()).toBe('#e2efda')

    // a two-column section
    const sects = blocks.filter((b) => b.t === 'sect')
    expect(sects.some((s) => s.t === 'sect' && s.page.cols === 2)).toBe(true)

    // page: A4, header and footer from the master page
    expect(meta.page.w).toBeCloseTo(595.3, 0)
    expect(meta.page.h).toBeCloseTo(841.9, 0)
    expect(meta.page.ml).toBeCloseTo(56.7, 0)
    expect(meta.page.header).toBeCloseTo(42.5, 0)
    // the body starts below the header (margin + height + spacing)
    expect(meta.page.mt).toBeCloseTo(42.5 + 17 + 11.3, 0)
    expect(meta.header?.map((b) => (b.t === 'p' ? textOf(b) : '')).join('')).toBe('Quarterly Report — Draft')
    expect(meta.footer?.map((b) => (b.t === 'p' ? textOf(b) : '')).join('')).toBe('Page {PAGE} of {NUMPAGES}')
    // styles offered to the editor
    expect(meta.styles?.some((s) => s.id === 'Heading1' && s.canon === 'h1')).toBe(true)
  })

  test('sample.odt: notes, bookmarks, tab stops and the package parts', async () => {
    const part = await docxParts(fixture('sample.odt'))
    const doc = await part('word/document.xml')
    const notes = await part('word/footnotes.xml')
    expect(notes).toContain('<w:footnote w:id="1">')
    expect(notes).toContain('Figures are rounded to the nearest thousand.')
    expect(notes).toContain('<w:footnoteRef/>')
    expect(doc).toContain('<w:footnoteReference w:id="1"/>')
    expect(doc).toMatch(/<w:bookmarkStart w:id="\d+" w:name="highlights"\/>/)
    expect(doc).toContain('<w:hyperlink w:anchor="highlights"')
    // a right tab stop with a dot leader at 12 cm
    expect(doc).toContain('<w:tab w:val="right" w:leader="dot" w:pos="6803"/>')
    const types = await part('[Content_Types].xml')
    for (const p of ['/word/styles.xml', '/word/numbering.xml', '/word/settings.xml', '/word/footnotes.xml', '/word/header1.xml', '/word/footer2.xml']) expect(types).toContain(`PartName="${p}"`)
    expect(types).toContain('Extension="png"')
    const styles = await part('word/styles.xml')
    expect(styles).toContain('<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Heading"/>')
    expect(styles).toContain('w:styleId="Hyperlink"')
    expect(await part('docProps/core.xml')).toContain('<dc:title>Quarterly Report</dc:title>')
  })

  test('sample.fodt: the flat XML document converts like the package', async () => {
    const zipped = await open(fixture('sample.odt'))
    const flatDoc = await open(fixture('sample.fodt'), 'sample.fodt')
    expect(flatDoc.info.converted).toBe('odt')
    expect(flatDoc.paras.map(textOf)).toEqual(zipped.paras.map(textOf))
    expect(flatDoc.paras.flatMap((p) => p.runs).some((r) => r.t === 'img')).toBe(true)
    expect(flatDoc.meta.header?.length).toBe(1)
  })

  test('template.ott opens as a document', async () => {
    const { info, paras } = await open(fixture('template.ott'), 'notes.ott')
    expect(info.converted).toBe('odt')
    expect(paras.map(textOf)).toEqual(['Meeting notes', 'Date:', 'Attendees:'])
    expect(paras[0].sid).toBe('Heading1')
  })

  test('pandoc and textutil documents', async () => {
    const p = await open(fixture('pandoc.odt'))
    expect(p.paras[0].sid).toBe('Title')
    expect(p.paras.find((x) => textOf(x) === '1. Highlights')?.sid).toBe('Heading1')
    expect(p.paras.find((x) => textOf(x) === 'Revenue grew 24% year over year')?.marker?.text).toBe('•')
    expect(p.paras.find((x) => textOf(x) === 'Expand to APAC')?.marker?.text).toBe('3.')
    expect(cellText(p.tables[0], 0, 3)).toBe('Total')
    expect(p.paras.flatMap((x) => x.runs).some((r) => r.t === 'img')).toBe(true)
    expect(p.meta.footer?.some((b) => b.t === 'p' && b.runs.some((r) => r.t === 'field'))).toBe(true)

    const t = await open(fixture('textutil.odt'))
    expect(t.paras.map(textOf).slice(0, 3)).toEqual(['Sample page', 'Opened from an HTML file: headings, lists, emphasis and links become a document, and the page is saved back as HTML.', 'Checklist'])
    expect(run(t.paras[0], 'Sample')!.s.sz).toBe(24)
    expect(t.paras.find((x) => textOf(x) === 'Open the page')?.marker?.text).toBe('•')
    expect(t.meta.page.w).toBe(612)
  })
})

// ---------------------------------------------------------------- features

describe('OpenDocument text → DOCX', () => {
  test('white space collapses; text:s, text:tab and text:line-break keep theirs', async () => {
    const { paras } = await open(
      await odt({
        body:
          '<text:p>\n  Hello   <text:span>  world</text:span><text:s text:c="2"/>!<text:tab/>x<text:line-break/>  y <text:soft-page-break/>z \u{1F600}&#xD800;</text:p><text:p/>' +
          '<text:p>shown<text:hidden-text text:is-hidden="true" text:string-value="secret"/></text:p><text:p><text:hidden-paragraph text:is-hidden="true"/>gone</text:p>',
      }),
    )
    expect(textOf(paras[0])).toBe('Hello world  !\tx\ny z \u{1F600}')
    expect(paras.length).toBe(4)
    expect(textOf(paras[1])).toBe('')
    expect(textOf(paras[2])).toBe('shown')
    // a paragraph a condition hides
    expect(textOf(paras[3])).toBe('')
  })

  test('named styles become Word styles, automatic styles direct formatting, with inheritance', async () => {
    const styles =
      '<style:default-style style:family="paragraph">' + tp('style:font-name="Liberation Serif" fo:font-size="11pt"') + '</style:default-style>' +
      style('Standard', 'paragraph', '') +
      style('Heading', 'paragraph', tp('fo:font-size="14pt"'), ' style:parent-style-name="Standard"') +
      style('Heading_20_1', 'paragraph', pp('fo:margin-top="12pt" fo:keep-with-next="always"') + tp('fo:font-size="130%" fo:font-weight="bold"'), ' style:display-name="Heading 1" style:parent-style-name="Heading" style:default-outline-level="1"') +
      style('Text_20_body', 'paragraph', pp('fo:margin-bottom="6pt"'), ' style:display-name="Text body" style:parent-style-name="Standard"') +
      style('My_20_Note', 'paragraph', tp('fo:color="#336699"'), ' style:display-name="My Note" style:parent-style-name="Text_20_body"') +
      style('Emphasis', 'text', tp('fo:font-style="italic"'))
    const auto = style('P1', 'paragraph', pp('fo:text-align="center"') + tp('fo:font-weight="bold" fo:font-size="150%"'), ' style:parent-style-name="Text_20_body"') + style('T1', 'text', tp('fo:color="#ff0000"'))
    const bytes = await odt({
      styles,
      auto,
      body:
        '<text:h text:style-name="Heading_20_1" text:outline-level="1">Title</text:h>' +
        '<text:p text:style-name="P1">Centered <text:span text:style-name="Emphasis">em</text:span> <text:span text:style-name="T1">red</text:span></text:p>' +
        '<text:p text:style-name="My_20_Note">note</text:p><text:p>plain</text:p>',
    })
    const { paras, meta } = await open(bytes)
    expect(meta.font).toContain('Liberation Serif')
    expect(meta.sz).toBe(11)
    const [h, p, note, plain] = paras
    expect(h.sid).toBe('Heading1')
    expect(h.ps).toMatchObject({ before: 12, keepNext: true, outline: 0 })
    // 130% of 14pt, to Word's half points
    expect(run(h, 'Title')!.s).toMatchObject({ b: true, sz: 18 })
    expect(p.sid).toBe('BodyText')
    expect(p.ps).toMatchObject({ align: 'center', after: 6 })
    // 150% of 11pt; bold from the automatic style on every run, the character style on top
    expect(run(p, 'Centered')!.s).toMatchObject({ b: true, sz: 16.5 })
    expect(run(p, 'em')!.s).toMatchObject({ b: true, i: true })
    expect(run(p, 'red')!.s.color).toBe('#ff0000')
    expect(note.sid).toBe('MyNote')
    expect(run(note, 'note')!.s.color).toBe('#336699')
    expect(plain.sid).toBe('Normal')

    const part = await docxParts(bytes)
    const st = await part('word/styles.xml')
    expect(st).toContain('<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/>')
    expect(st).toContain('<w:style w:type="paragraph" w:styleId="MyNote"><w:name w:val="My Note"/><w:basedOn w:val="BodyText"/>')
    expect(st).toContain('<w:style w:type="character" w:styleId="Emphasis">')
    expect(await part('word/document.xml')).toContain('<w:rStyle w:val="Emphasis"/>')
  })

  test('character properties', async () => {
    const auto =
      style('T1', 'text', tp('style:text-underline-style="solid" style:text-underline-type="double"')) +
      style('T2', 'text', tp('style:text-line-through-style="solid" style:text-line-through-type="double"')) +
      style('T3', 'text', tp('style:text-position="sub 58%"')) +
      style('T4', 'text', tp('fo:text-transform="uppercase" fo:letter-spacing="1pt"')) +
      style('T5', 'text', tp('fo:background-color="#00ff00" fo:color="#123456" style:text-underline-style="wave"')) +
      style('T6', 'text', tp('style:text-position="33% 100%" fo:font-size="10pt"')) +
      style('T7', 'text', tp('text:display="none"')) +
      style('T8', 'text', tp('fo:font-family="\'DejaVu Sans\'" style:text-scale="80%"'))
    const bytes = await odt({
      auto,
      body: '<text:p><text:span text:style-name="T1">dbl</text:span><text:span text:style-name="T2">x2</text:span>H<text:span text:style-name="T3">2</text:span><text:span text:style-name="T4">caps</text:span><text:span text:style-name="T5">wavy</text:span><text:span text:style-name="T6">up</text:span><text:span text:style-name="T7">hidden</text:span><text:span text:style-name="T8">font</text:span></text:p>',
    })
    const { paras } = await open(bytes)
    const p = paras[0]
    expect(run(p, 'dbl')!.s.u).toBe('double')
    expect(run(p, 'x2')!.s.dstrike).toBe(true)
    expect(p.runs.find((r) => r.t === 'text' && r.text === '2')?.t === 'text' && (p.runs.find((r) => r.t === 'text' && r.text === '2') as any).s.va).toBe('sub')
    expect(run(p, 'caps')!.s).toMatchObject({ caps: true, spacing: 1 })
    expect(run(p, 'wavy')!.s).toMatchObject({ shd: '#00ff00', color: '#123456', u: 'wave' })
    expect(textOf(p)).not.toContain('hidden')
    expect(run(p, 'font')!.s.font).toContain('DejaVu Sans')
    const doc = await (await docxParts(bytes))('word/document.xml')
    // raised by a third of 10pt, in half-points
    expect(doc).toContain('<w:position w:val="7"/>')
    expect(doc).toContain('<w:w w:val="80"/>')
  })

  test('paragraph properties', async () => {
    const auto =
      style('P1', 'paragraph', pp('fo:text-align="end" fo:margin-left="1in" fo:margin-right="0.5in" fo:text-indent="-0.25in" fo:margin-top="3pt" fo:margin-bottom="9pt" fo:line-height="150%"')) +
      style('P2', 'paragraph', pp('fo:line-height="20pt" fo:keep-together="always" fo:break-before="page" fo:background-color="#eeeeee" fo:border="1pt solid #ff0000" fo:padding="2pt"')) +
      style('P3', 'paragraph', pp('style:line-height-at-least="18pt" fo:border-bottom="0.5pt double #0000ff" fo:text-align="justify"', '<style:tab-stops><style:tab-stop style:position="2in" style:type="center"/><style:tab-stop style:position="4in" style:type="char" style:char="."/></style:tab-stops>')) +
      style('P4', 'paragraph', pp('fo:break-after="page"'))
    const bytes = await odt({ auto, body: '<text:p text:style-name="P1">a</text:p><text:p text:style-name="P2">b</text:p><text:p text:style-name="P3">c</text:p><text:p text:style-name="P4">d</text:p><text:p>e</text:p>' })
    const { paras } = await open(bytes)
    expect(paras[0].ps).toMatchObject({ align: 'right', indL: 72, indR: 36, indFirst: -18, before: 3, after: 9, line: 1.5, lineRule: 'auto' })
    expect(paras[1].ps).toMatchObject({ line: 20, lineRule: 'exact', shd: '#eeeeee', pageBreakBefore: true })
    expect(paras[1].ps.bTop).toMatchObject({ w: 1, style: 'solid', color: '#ff0000' })
    expect(paras[2].ps).toMatchObject({ line: 18, lineRule: 'atLeast', align: 'justify' })
    expect(paras[2].ps.bBottom?.style).toBe('double')
    // a break after a paragraph: the next one starts the page
    expect(paras[4].ps.pageBreakBefore).toBe(true)
    const doc = await (await docxParts(bytes))('word/document.xml')
    expect(doc).toContain('<w:keepLines/>')
    expect(doc).toContain('<w:tabs><w:tab w:val="center" w:pos="2880"/><w:tab w:val="decimal" w:pos="5760"/></w:tabs>')
  })

  test('lists: bullets and numbers, nesting, start values, continued and separate lists, headers', async () => {
    const level = (n: number, kind: string, attrs: string) =>
      `<text:list-level-style-${kind} text:level="${n}" ${attrs}><style:list-level-properties text:list-level-position-and-space-mode="label-alignment"><style:list-level-label-alignment text:label-followed-by="listtab" fo:text-indent="-0.25in" fo:margin-left="${0.5 * n}in"/></style:list-level-properties></text:list-level-style-${kind}>`
    const auto =
      `<text:list-style style:name="L1">${level(1, 'bullet', 'text:bullet-char="•"')}${level(2, 'bullet', 'text:bullet-char="–"')}</text:list-style>` +
      `<text:list-style style:name="L2">${level(1, 'number', 'style:num-format="1" style:num-suffix="."')}${level(2, 'number', 'style:num-format="i" style:num-prefix="(" style:num-suffix=")"')}${level(3, 'number', 'style:num-format="1" style:num-suffix="." text:display-levels="3"')}</text:list-style>` +
      // the old way: label width and position
      '<text:list-style style:name="L3"><text:list-level-style-number text:level="1" style:num-format="A" style:num-suffix="." text:start-value="3"><style:list-level-properties text:space-before="0.25in" text:min-label-width="0.25in"/></text:list-level-style-number></text:list-style>'
    const item = (t: string, inner = '') => `<text:list-item><text:p>${t}</text:p>${inner}</text:list-item>`
    const bytes = await odt({
      auto,
      body:
        `<text:list text:style-name="L1">${item('apple', `<text:list>${item('red')}${item('green')}</text:list>`)}${item('pear')}</text:list>` +
        `<text:list text:style-name="L2" xml:id="list1">${item('one')}${item('two', `<text:list>${item('two-a', `<text:list>${item('deep')}</text:list>`)}${item('two-b')}</text:list>`)}<text:list-header><text:p>no label</text:p></text:list-header></text:list>` +
        '<text:p>between</text:p>' +
        `<text:list text:style-name="L2" text:continue-numbering="true">${item('three')}</text:list>` +
        `<text:list text:style-name="L2">${item('again')}<text:list-item text:start-value="7"><text:p>seven</text:p><text:p>second paragraph</text:p></text:list-item>${item('eight')}</text:list>` +
        `<text:list text:style-name="L2" text:continue-list="list1">${item('four')}</text:list>` +
        `<text:list text:style-name="L3">${item('C')}${item('D')}</text:list>`,
    })
    const { paras } = await open(bytes)
    const m = (t: string) => paras.find((p) => textOf(p) === t)?.marker?.text
    expect([m('apple'), m('red'), m('green'), m('pear')]).toEqual(['•', '–', '–', '•'])
    expect([m('one'), m('two'), m('two-a'), m('deep'), m('two-b')]).toEqual(['1.', '2.', '(i)', '2.i.1.', '(ii)'])
    expect(m('no label')).toBeUndefined()
    expect(m('three')).toBe('3.')
    expect([m('again'), m('seven'), m('eight')]).toEqual(['1.', '7.', '8.'])
    expect(m('second paragraph')).toBeUndefined()
    expect(m('four')).toBe('4.')
    expect([m('C'), m('D')]).toEqual(['C.', 'D.'])
    // indents: level 2 at 1 inch with a quarter inch hanging; a paragraph in an item under its text
    const twoA = paras.find((p) => textOf(p) === 'two-a')!
    expect(twoA.ps).toMatchObject({ indL: 72, indFirst: -18 })
    expect(paras.find((p) => textOf(p) === 'second paragraph')!.ps).toMatchObject({ indL: 36, indFirst: 0 })
    expect(paras.find((p) => textOf(p) === 'C')!.ps).toMatchObject({ indL: 36, indFirst: -18 })
  })

  test('tables: columns, spans, covered cells, header rows, cell formatting, nested tables', async () => {
    const auto =
      style('T', 'table', '<style:table-properties style:width="6in" table:align="center"/>') +
      style('T.A', 'table-column', '<style:table-column-properties style:column-width="1in" style:rel-column-width="1000*"/>') +
      style('T.B', 'table-column', '<style:table-column-properties style:column-width="2.5in" style:rel-column-width="2500*"/>') +
      style('T.R', 'table-row', '<style:table-row-properties style:min-row-height="0.5in"/>') +
      style('T.C', 'table-cell', '<style:table-cell-properties fo:background-color="#ffcc00" fo:padding-left="0.1in" fo:border="0.5pt solid #000000" fo:border-right="none" style:vertical-align="bottom"/>')
    const cell = (t: string, attrs = '') => `<table:table-cell${attrs} office:value-type="string"><text:p>${t}</text:p></table:table-cell>`
    const bytes = await odt({
      auto,
      body:
        '<table:table table:name="T" table:style-name="T"><table:table-column table:style-name="T.A"/><table:table-column table:style-name="T.B" table:number-columns-repeated="2"/>' +
        `<table:table-header-rows><table:table-row>${cell('h1')}${cell('h2')}${cell('h3')}</table:table-row></table:table-header-rows>` +
        `<table:table-row table:style-name="T.R">${cell('wide', ' table:number-columns-spanned="2" table:style-name="T.C"')}<table:covered-table-cell/>${cell('tall', ' table:number-rows-spanned="2"')}</table:table-row>` +
        `<table:table-row>${cell('x')}<table:table-cell><table:table><table:table-column table:number-columns-repeated="2"/><table:table-row>${cell('n1')}${cell('n2')}</table:table-row></table:table></table:table-cell><table:covered-table-cell/></table:table-row>` +
        `<table:table-row table:number-rows-repeated="2">${cell('r', ' table:number-columns-repeated="3"')}</table:table-row>` +
        '</table:table><table:table table:name="U"><table:table-column/><table:table-row>' + cell('second') + '</table:table-row></table:table>',
    })
    const { tables } = await open(bytes)
    const t = tables[0]
    expect(t.cols).toEqual([72, 180, 180])
    expect(t.align).toBe('center')
    expect(t.rows.length).toBe(5)
    expect(t.rows[0].header).toBe(true)
    expect(t.rows[1].h).toBe(36)
    expect(t.rows[1].cells[0]).toMatchObject({ colSpan: 2, shd: '#ffcc00', vAlign: 'bottom' })
    expect(t.rows[1].cells[0].pad[3]).toBeCloseTo(7.2)
    expect(t.rows[1].cells[0].borders).toMatchObject({ r: 'none' })
    expect(t.rows[1].cells[0].borders.l).toContain('solid')
    expect(t.rows[1].cells[1].rowSpan).toBe(2)
    expect(cellText(t, 1, 1)).toBe('tall')
    expect(t.rows[2].cells.length).toBe(2)
    const nested = t.rows[2].cells[1].blocks.find((b): b is Table => b.t === 'tbl')!
    expect(cellText(nested, 0, 1)).toBe('n2')
    expect(t.rows[3].cells.map((c) => c.blocks.map((b) => (b.t === 'p' ? textOf(b) : '')).join(''))).toEqual(['r', 'r', 'r'])
    // two tables in a row stay two tables
    expect(tables.length).toBe(2)
    expect(cellText(tables[1], 0, 0)).toBe('second')
  })

  test('frames: inline and anchored pictures, text boxes, linked pictures', async () => {
    const masterAuto = ''
    const auto = style('fr1', 'graphic', '<style:graphic-properties style:wrap="parallel" style:horizontal-pos="from-left" style:horizontal-rel="paragraph" style:vertical-pos="from-top" style:vertical-rel="paragraph"/>')
    const bytes = await odt({
      auto,
      masterAuto,
      files: { 'Pictures/dot.png': PNG },
      body:
        '<text:p>A <draw:frame draw:name="i1" text:anchor-type="as-char" svg:width="1in" svg:height="0.5in"><draw:image xlink:href="Pictures/dot.png"/><svg:desc>Two dots</svg:desc></draw:frame> picture</text:p>' +
        '<text:p><draw:frame draw:style-name="fr1" text:anchor-type="paragraph" svg:x="1in" svg:y="0.25in" svg:width="2in" svg:height="1in"><draw:image xlink:href="Pictures/dot.png"/></draw:frame>anchored</text:p>' +
        '<text:p><draw:frame text:anchor-type="paragraph" svg:width="2in"><draw:text-box fo:min-height="0.5in"><text:p>In a box</text:p></draw:text-box></draw:frame>after</text:p>' +
        '<text:p><draw:a xlink:href="https://example.com/"><draw:frame text:anchor-type="as-char" svg:width="10pt" svg:height="5pt"><draw:image xlink:href="Pictures/dot.png"/></draw:frame></draw:a></text:p>' +
        '<text:p><draw:frame text:anchor-type="as-char" svg:width="1in" svg:height="1in"><draw:image xlink:href="Pictures/missing.png"/></draw:frame>missing</text:p>',
    })
    const { paras } = await open(bytes)
    const img = paras[0].runs.find((r) => r.t === 'img')!
    expect(img).toMatchObject({ w: 72, h: 36, alt: 'Two dots' })
    expect(textOf(paras[0])).toBe('A [img] picture')
    expect(paras[1].runs.find((r) => r.t === 'img')).toMatchObject({ w: 144, h: 72 })
    const box = paras[2].runs.find((r) => r.t === 'box')!
    expect(box.t === 'box' && box.w).toBe(144)
    expect(box.t === 'box' && box.blocks.map((b) => (b.t === 'p' ? textOf(b) : '')).join('')).toBe('In a box')
    expect(paras[3].runs.find((r) => r.t === 'img')).toMatchObject({ href: 'https://example.com/' })
    expect(textOf(paras[4])).toBe('missing')
    const part = await docxParts(bytes)
    const doc = await part('word/document.xml')
    expect(doc).toContain('<wp:positionH relativeFrom="column"><wp:posOffset>914400</wp:posOffset></wp:positionH>')
    expect(doc).toContain('<wp:wrapSquare wrapText="bothSides"/>')
    expect(doc).toContain('<w:txbxContent>')
    expect(await part('word/_rels/document.xml.rels')).toContain('Target="media/image1.png"')
  })

  test('links, bookmarks, notes and fields', async () => {
    const styles = '<text:notes-configuration text:note-class="endnote" style:num-format="i"/>'
    const bytes = await odt({
      styles,
      body:
        '<text:p><text:bookmark text:name="top"/>See <text:a xlink:type="simple" xlink:href="https://example.com/a b">the site</text:a> or <text:a xlink:href="#part%202">part 2</text:a>.</text:p>' +
        '<text:p><text:bookmark-start text:name="part 2"/>Part<text:note text:id="n1" text:note-class="footnote"><text:note-citation>1</text:note-citation><text:note-body><text:p>A footnote.</text:p></text:note-body></text:note> two<text:note text:note-class="endnote"><text:note-citation>i</text:note-citation><text:note-body><text:p>An endnote.</text:p></text:note-body></text:note><text:bookmark-end text:name="part 2"/></text:p>' +
        '<text:p>Page <text:page-number text:select-page="current">3</text:page-number> of <text:page-count>9</text:page-count>, <text:date>2024-05-01</text:date> <text:title>Report</text:title></text:p>',
    })
    const { paras } = await open(bytes)
    expect(run(paras[0], 'the site')!.href).toBe('https://example.com/a%20b')
    expect(run(paras[0], 'part 2')!.href).toBe('#part_2')
    expect(textOf(paras[1])).toBe('Part[1] two[1]')
    expect(textOf(paras[2])).toBe('Page {PAGE} of {NUMPAGES}, 2024-05-01 Report')
    const part = await docxParts(bytes)
    const doc = await part('word/document.xml')
    expect(doc).toMatch(/<w:bookmarkStart w:id="(\d+)" w:name="top"\/><w:bookmarkEnd w:id="\1"\/>/)
    expect(doc).toMatch(/<w:bookmarkStart w:id="(\d+)" w:name="part_2"\/>.*<w:bookmarkEnd w:id="\1"\/>/)
    expect(doc).toContain('<w:endnoteReference w:id="1"/>')
    expect(await part('word/footnotes.xml')).toContain('A footnote.')
    expect(await part('word/endnotes.xml')).toContain('An endnote.')
    const settings = await part('word/settings.xml')
    expect(settings).toContain('<w:endnotePr><w:numFmt w:val="lowerRoman"/><w:endnote w:id="-1"/><w:endnote w:id="0"/></w:endnotePr>')
  })

  test('page layout, headers and footers (left, first), master page changes', async () => {
    const layout = (name: string, w: string, h: string, extra = '') =>
      `<style:page-layout style:name="${name}"><style:page-layout-properties fo:page-width="${w}" fo:page-height="${h}" fo:margin-top="1in" fo:margin-bottom="0.5in" fo:margin-left="1.25in" fo:margin-right="0.75in" ${extra}><style:columns fo:column-count="1"/></style:page-layout-properties>` +
      '<style:header-style><style:header-footer-properties fo:min-height="0.25in" fo:margin-bottom="0.25in"/></style:header-style><style:footer-style/></style:page-layout>'
    const bytes = await odt({
      masterAuto: layout('pm1', '8.5in', '11in') + layout('pm2', '11in', '8.5in', 'style:print-orientation="landscape" style:num-format="i"'),
      master:
        '<style:master-page style:name="Standard" style:page-layout-name="pm1"><style:header><text:p>Odd header</text:p></style:header><style:header-left><text:p>Even header</text:p></style:header-left>' +
        '<loext:header-first><text:p>First header</text:p></loext:header-first><style:footer><text:p>Page <text:page-number>1</text:page-number></text:p></style:footer></style:master-page>' +
        '<style:master-page style:name="Landscape" style:page-layout-name="pm2"><style:header><text:p>Wide header</text:p></style:header></style:master-page>',
      auto: style('P1', 'paragraph', '', ' style:parent-style-name="Standard" style:master-page-name="Landscape"') + style('P2', 'paragraph', pp('style:page-number="5"'), ' style:master-page-name="Standard"'),
      body: '<text:p>portrait</text:p><text:p text:style-name="P1">landscape</text:p><text:p text:style-name="P2">portrait again</text:p>',
    })
    const { blocks, meta } = await open(bytes)
    const sects = blocks.filter((b) => b.t === 'sect')
    expect(sects.length).toBe(2)
    expect(sects[0].t === 'sect' && sects[0].page).toMatchObject({ w: 612, h: 792, ml: 90, mr: 54, header: 72, mt: 108 })
    expect(sects[1].t === 'sect' && sects[1].page).toMatchObject({ w: 792, h: 612 })
    expect(meta.page).toMatchObject({ w: 612, h: 792 })
    expect(meta.titlePg).toBe(true)
    const txt = (bs?: Block[]) => bs?.map((b) => (b.t === 'p' ? textOf(b) : '')).join('')
    expect(txt(meta.header)).toBe('Odd header')
    expect(txt(meta.firstHeader)).toBe('First header')
    expect(txt(meta.footer)).toBe('Page {PAGE}')
    const part = await docxParts(bytes)
    expect(await part('word/settings.xml')).toContain('<w:evenAndOddHeaders/>')
    const doc = await part('word/document.xml')
    expect(doc).toContain('<w:pgNumType w:fmt="lowerRoman"/>')
    expect(doc).toContain('<w:pgNumType w:start="5"/>')
    expect(doc).toMatch(/<w:headerReference w:type="even" r:id="rId\d+"\/>/)
  })

  test('a first page style followed by another one, and a page background', async () => {
    const pl = '<style:page-layout style:name="pm1"><style:page-layout-properties fo:page-width="21cm" fo:page-height="29.7cm" fo:background-color="#fafafa"/></style:page-layout>'
    const bytes = await odt({
      masterAuto: pl,
      master:
        '<style:master-page style:name="First_20_Page" style:display-name="First Page" style:page-layout-name="pm1" style:next-style-name="Standard"/>' +
        '<style:master-page style:name="Standard" style:page-layout-name="pm1"><style:footer><text:p>Default footer</text:p></style:footer></style:master-page>',
      auto: style('P1', 'paragraph', '', ' style:master-page-name="First_20_Page"'),
      body: '<text:p text:style-name="P1">cover</text:p><text:p>body</text:p>',
    })
    const { meta } = await open(bytes)
    expect(meta.titlePg).toBe(true)
    expect(meta.footer?.map((b) => (b.t === 'p' ? textOf(b) : '')).join('')).toBe('Default footer')
    expect(meta.firstFooter?.map((b) => (b.t === 'p' ? textOf(b) : '')).join('')).toBe('')
    const doc = await (await docxParts(bytes))('word/document.xml')
    expect(doc).toContain('<w:background w:color="FAFAFA"/>')
  })

  test('sections, indexes and other containers keep their text', async () => {
    const auto = style('S1', 'section', '<style:section-properties><style:columns fo:column-count="3" fo:column-gap="0.2in"><style:column-sep style:width="0.5pt"/></style:columns></style:section-properties>')
    const bytes = await odt({
      auto,
      body:
        '<text:p>before</text:p><text:section text:name="S" text:style-name="S1"><text:p>in columns</text:p><text:section text:name="inner"><text:p>nested</text:p></text:section></text:section>' +
        '<text:section text:name="hidden" text:display="none"><text:p>hidden</text:p></text:section>' +
        '<text:table-of-content text:name="TOC"><text:table-of-content-source text:outline-level="3"><text:index-title-template>Contents</text:index-title-template></text:table-of-content-source>' +
        '<text:index-body><text:index-title text:name="t"><text:p>Contents</text:p></text:index-title><text:p>Intro<text:tab/>1</text:p></text:index-body></text:table-of-content>' +
        '<text:tracked-changes><text:changed-region text:id="c1"><text:deletion><text:p>deleted</text:p></text:deletion></text:changed-region></text:tracked-changes>' +
        '<text:p>a<text:change text:change-id="c1"/>b<office:annotation><text:p>comment</text:p></office:annotation>c<text:ruby><text:ruby-base>漢</text:ruby-base><text:ruby-text>kan</text:ruby-text></text:ruby></text:p>',
    })
    const { paras, blocks } = await open(bytes)
    expect(paras.map(textOf)).toEqual(['before', 'in columns', 'nested', 'Contents', 'Intro\t1', 'abc漢'])
    const sects = blocks.filter((b) => b.t === 'sect')
    expect(sects.map((s) => s.t === 'sect' && [s.kind, s.page.cols])).toEqual([
      ['nextPage', undefined],
      ['continuous', 3],
    ])
    const doc = await (await docxParts(bytes))('word/document.xml')
    expect(doc).toContain('<w:cols w:num="3" w:space="288" w:sep="1"/>')
  })

  test('annotations become comments', async () => {
    const note = (name: string, who: string, text: string) =>
      `<office:annotation${name ? ` office:name="${name}"` : ''}><dc:creator xmlns:dc="http://purl.org/dc/elements/1.1/">${who}</dc:creator><dc:date xmlns:dc="http://purl.org/dc/elements/1.1/">2024-03-05T10:20:30.123</dc:date><text:p>${text}</text:p></office:annotation>`
    const bytes = await odt({
      body: `<text:p>Before ${note('c1', 'Ann', 'Check this')}covered text<office:annotation-end office:name="c1"/> after.</text:p><text:p>Point${note('', 'Bob', 'A remark')}</text:p>`,
    })
    const { engine, paras } = await open(bytes)
    const list = (await engine.call('comments.list', [])) as DocComment[]
    expect(list.map((c) => [c.author.name, c.text])).toEqual([
      ['Ann', 'Check this'],
      ['Bob', 'A remark'],
    ])
    expect(list[0].time).toStartWith('2024-03-05T10:20:30')
    const ann = list[0].id
    expect(paras[0].runs.filter((r) => r.t === 'text' && r.cm?.includes(ann)).map((r) => (r.t === 'text' ? r.text : '')).join('')).toBe('covered text')
    expect(textOf(paras[0])).toBe('Before covered text after.')
    expect(textOf(paras[1])).toBe('Point')
  })

  test('flat XML documents', async () => {
    const bytes = flat({
      styles: style('Standard', 'paragraph', ''),
      auto: style('T1', 'text', tp('fo:font-weight="bold"')),
      body: `<text:p>Flat <text:span text:style-name="T1">bold</text:span><draw:frame text:anchor-type="as-char" svg:width="2pt" svg:height="1pt"><draw:image><office:binary-data>${btoa(String.fromCharCode(...PNG))}</office:binary-data></draw:image></draw:frame></text:p>`,
    })
    const { info, paras } = await open(bytes, 'doc.fodt')
    expect(info.converted).toBe('odt')
    expect(run(paras[0], 'bold')!.s.b).toBe(true)
    expect(paras[0].runs.find((r) => r.t === 'img')).toMatchObject({ w: 2, h: 1 })
    const conv = await convertFlatOdf(bytes)
    const zip = await ZipArchive.open(new BlobSource(conv.file))
    expect(zip.has('word/media/image1.png')).toBe(true)
  })

  test('outline numbering of headings follows the outline style', async () => {
    const lvl = (n: number) => `<text:outline-level-style text:level="${n}" style:num-format="${n === 1 ? 'I' : '1'}" style:num-suffix="." text:display-levels="${n}"/>`
    // heading styles name the outline style as their list style (LibreOffice 7.6 and later)
    const styles = `<text:outline-style style:name="Outline">${lvl(1)}${lvl(2)}</text:outline-style>` + style('Heading_20_2', 'paragraph', '', ' style:display-name="Heading 2" style:default-outline-level="2" style:list-style-name="Outline"')
    const { paras } = await open(
      await odt({
        styles,
        body:
          '<text:h text:outline-level="1">A</text:h><text:h text:style-name="Heading_20_2" text:outline-level="2"><text:number>I.1.</text:number>A1</text:h><text:h text:outline-level="2">A2</text:h>' +
          '<text:h text:outline-level="1">B</text:h><text:h text:outline-level="2">B1</text:h><text:h text:outline-level="1" text:is-list-header="true">Unnumbered</text:h>',
      }),
    )
    expect(paras.map((p) => p.marker?.text)).toEqual(['I.', 'I.1.', 'I.2.', 'II.', 'II.1.', undefined])
    // the label LibreOffice wrote last (text:number) is not text
    expect(textOf(paras[1])).toBe('A1')
    // headings without a style still get Word's heading styles
    expect(paras.map((p) => p.sid)).toEqual(['Heading1', 'Heading2', 'Heading2', 'Heading1', 'Heading2', 'Heading1'])
  })

  test('damaged and hostile input fails cleanly or stays bounded', async () => {
    // a package without office:text
    const noText = await odt({ body: '' })
    const broken = new TextDecoder().decode(noText)
    expect(broken.length).toBeGreaterThan(0)
    const spreadsheet = await writeZip(
      new Map([
        ['mimetype', 'application/vnd.oasis.opendocument.text'],
        ['content.xml', `<office:document-content ${NS}><office:body><office:spreadsheet/></office:body></office:document-content>`],
      ]),
      'application/vnd.oasis.opendocument.text',
    )
    await expect(open(new Uint8Array(await spreadsheet.arrayBuffer()))).rejects.toThrow(/OpenDocument text/)
    // encrypted parts
    const encrypted = await odt({
      body: '',
      files: {
        'META-INF/manifest.xml': new TextEncoder().encode(
          '<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/>' +
            '<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"><manifest:encryption-data manifest:checksum-type="SHA1/1K" manifest:checksum="x"/></manifest:file-entry></manifest:manifest>',
        ),
      },
    })
    await expect(open(encrypted)).rejects.toThrow(/[Pp]assword/)
    // huge repeat counts and spaces, deep nesting, unknown elements
    let deep = 'x'
    for (let k = 0; k < 200; k++) deep = `<text:span>${deep}</text:span>`
    let lists = '<text:p>deep list</text:p>'
    for (let k = 0; k < 60; k++) lists = `<text:list><text:list-item>${lists}</text:list-item></text:list>`
    const t0 = performance.now()
    const { paras, tables } = await open(
      await odt({
        body:
          `<text:p>a<text:s text:c="999999999"/>b</text:p><text:p>${deep}</text:p>${lists}<foo:bar xmlns:foo="urn:x"><text:p>inside unknown</text:p></foo:bar>` +
          '<table:table><table:table-column table:number-columns-repeated="1000000"/><table:table-row table:number-rows-repeated="1000000"><table:table-cell table:number-columns-repeated="1000000"/></table:table-row></table:table>',
      }),
    )
    expect(performance.now() - t0).toBeLessThan(20000)
    expect(textOf(paras[0]).length).toBeLessThan(2000)
    expect(paras.some((p) => textOf(p) === 'inside unknown')).toBe(true)
    expect(tables[0].cols.length).toBeLessThanOrEqual(63)
  })
})
