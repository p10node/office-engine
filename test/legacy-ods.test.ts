import { describe, expect, test } from 'bun:test'
import { OdfStyles, parseOdfXml } from '../src/core/legacy/odf/common'
import { excelFormat } from '../src/core/legacy/odf/ods-format'
import { excelFormula } from '../src/core/legacy/odf/ods-formula'
import { convertFlatOdf, convertOdfZip } from '../src/core/legacy'
import { Engine } from '../src/core/worker/engine'
import type { BlockOut, CellDetail, SheetProgress } from '../src/core/xlsx/types'
import type { DocComment } from '../src/core/comments'
import { BlobSource } from '../src/core/source'
import { ZipArchive } from '../src/core/zip'
import { writeZip } from '../src/core/zipwrite'

const fixture = async (name: string) => new Uint8Array(await Bun.file(`${import.meta.dir}/fixtures/legacy/${name}`).arrayBuffer())

/** opens a file in the engine; `sheet(i)` loads a sheet and reads its first rows */
async function openBook(bytes: Uint8Array, name = 'book.ods') {
  const events: { event: string; data: any }[] = []
  const engine = new Engine((event, data) => events.push({ event, data }))
  const info = await engine.open(new Blob([bytes as Uint8Array<ArrayBuffer>]), { name })
  const sheet = async (index: number, rows = 100) => {
    await engine.call('xlsx.load', [index])
    const done = () => events.some((e) => e.event === 'xlsx:progress' && e.data.sheet === index && e.data.done)
    for (let i = 0; i < 1000 && !done(); i++) await new Promise((r) => setTimeout(r, 5))
    const all = events.filter((e) => e.event === 'xlsx:progress' && e.data.sheet === index).map((e) => e.data as SheetProgress)
    const progress: SheetProgress = { ...Object.assign({}, ...all), rowSizes: all.flatMap((p) => p.rowSizes) }
    const block = (await engine.call('xlsx.block', [index, 0, rows])) as BlockOut
    const cell = (r: number, c: number) => block.rows.find((x) => x.r === r)?.cells.find((x) => x[0] === c)
    const text = (r: number, c: number) => cell(r, c)?.[1] ?? ''
    const style = (r: number, c: number) => info.xlsx!.styles[cell(r, c)?.[2] ?? 0]
    const detail = async (r: number, c: number) => (await engine.call('xlsx.cell', [index, r, c])) as CellDetail | undefined
    const row = (r: number, n: number) => Array.from({ length: n }, (_, c) => text(r, c))
    return { progress, block, cell, text, style, detail, row }
  }
  return { engine, info, sheet }
}

const NS =
  'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" ' +
  'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" ' +
  'xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" ' +
  'xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/" ' +
  'xmlns:calcext="urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0" xmlns:of="urn:oasis:names:tc:opendocument:xmlns:of:1.2"'

/** an .ods package around spreadsheet content (`auto`: automatic styles of content.xml, `styles`: office:styles) */
async function ods(body: string, opts: { auto?: string; styles?: string; settings?: string; extra?: [string, string][] } = {}): Promise<Uint8Array> {
  const files = new Map<string, string>([
    ['mimetype', 'application/vnd.oasis.opendocument.spreadsheet'],
    [
      'content.xml',
      `<?xml version="1.0" encoding="UTF-8"?><office:document-content ${NS} office:version="1.3"><office:automatic-styles>${opts.auto ?? ''}</office:automatic-styles>` +
        `<office:body><office:spreadsheet>${body}</office:spreadsheet></office:body></office:document-content>`,
    ],
    ['styles.xml', `<?xml version="1.0" encoding="UTF-8"?><office:document-styles ${NS}><office:styles>${opts.styles ?? ''}</office:styles></office:document-styles>`],
    ...(opts.extra ?? []),
  ])
  if (opts.settings) files.set('settings.xml', `<?xml version="1.0" encoding="UTF-8"?><office:document-settings ${NS}><office:settings>${opts.settings}</office:settings></office:document-settings>`)
  return new Uint8Array(await (await writeZip(files)).arrayBuffer())
}

/** the parts of the XLSX an .ods converts to */
async function converted(bytes: Uint8Array) {
  const conv = await convertOdfZip(await ZipArchive.open(new BlobSource(new Blob([bytes as Uint8Array<ArrayBuffer>]))))
  const zip = await ZipArchive.open(new BlobSource(conv!.file))
  return (path: string) => zip.text(path)
}

/** a comment's cell: [row, column] */
const cellOf = (c: DocComment) => (c.at && 'r' in c.at ? [c.at.r, c.at.c] : undefined)

const cellXml = (attrs: string, text?: string) => `<table:table-cell ${attrs}>${text !== undefined ? `<text:p>${text}</text:p>` : ''}</table:table-cell>`

describe('ods → xlsx: the sample', () => {
  test('sheets, values with their formats, formulas, merges, frozen panes, look', async () => {
    const { info, sheet, engine } = await openBook(await fixture('sample.ods'))
    expect(info.kind).toBe('xlsx')
    expect(info.converted).toBe('ods')
    expect(info.xlsx!.sheets.map((s) => [s.name, s.hidden])).toEqual([
      ['Sales', false],
      ['Notes', false],
      ['Data', false],
      ['Lookup', true],
    ])
    const s = await sheet(0)
    expect(s.text(0, 0)).toBe('Sales by region — FY2026')
    expect(s.progress.merges).toContainEqual({ r0: 0, c0: 0, r1: 0, c1: 5 })
    expect(s.progress.frozen).toEqual({ rows: 2, cols: 1 })
    expect(s.row(1, 6)).toEqual(['Region', 'Date', 'Units', 'Price', 'Revenue', 'Margin'])
    // dates, integers, currency, a formula with its cached value, percentages: numbers shown with their formats
    const r2 = s.row(2, 7)
    expect(r2[1]).toBe('2026-01-15')
    expect(r2[3]).toMatch(/^\$\d+\.\d\d$/)
    expect(Number(r2[4].replace(/,/g, ''))).toBeCloseTo(Number(r2[2]) * Number(r2[3].slice(1)), 2)
    expect(r2[5]).toMatch(/^\d+\.\d%$/)
    expect(s.cell(2, 1)![3]).toBe(1)
    expect((await s.detail(2, 4))!.formula).toBe('=C3*D3')
    // a formula on another (hidden) sheet
    expect((await s.detail(2, 6))!.formula).toBe('=VLOOKUP(A3,Lookup!$A$1:$B$5,2,0)')
    expect(s.text(62, 0)).toBe('Total')
    expect((await s.detail(62, 4))!.formula).toBe('=SUM(E3:E62)')
    // negative numbers in red parentheses, booleans, long dates, durations
    expect(s.text(3, 7)).toBe('(42.50)')
    expect(s.cell(3, 7)![4]).toBe('#ff0000')
    expect(s.text(5, 7)).toBe('TRUE')
    expect(s.text(6, 7)).toBe('Sunday, December 31, 2023 6:00 PM')
    expect(s.text(7, 7)).toBe('36:30:00')
    // header look: bold white text on blue, a bottom line; the title larger
    expect(s.style(1, 0)).toMatchObject({ fontWeight: 700, color: '#ffffff', background: '#4472c4', textAlign: 'center' })
    expect(String(s.style(1, 0).borderBottom)).toContain('#8ea9db')
    expect(s.style(0, 0)).toMatchObject({ fontWeight: 700, fontSize: '16pt', color: '#1f3864' })
    expect(s.style(2, 0).background).toBe('#d9e1f2')
    expect(s.style(3, 0).background).toBeUndefined()
    // column widths (3.5cm), the default font
    expect(s.progress.colInfo!.find((c) => c.min === 0)!.w).toBe(132)
    expect(info.xlsx!.baseFont).toMatchObject({ fontSize: '10pt' })
    // the comment on the Revenue header
    const comments = (await engine.call('comments.list', [])) as DocComment[]
    expect(comments.map((c) => [c.author.name, c.text, cellOf(c)])).toEqual([['Sample Author', 'Units × Price, computed by the spreadsheet.', [1, 4]]])

    const notes = await sheet(1)
    expect(notes.text(0, 0)).toContain('\nA second paragraph in the same cell.')
    expect(notes.style(0, 0)).toMatchObject({ whiteSpace: 'pre-wrap', alignItems: 'flex-start' })
    expect(notes.style(0, 1)).toMatchObject({ fontStyle: 'italic', textDecoration: 'underline' })
    expect(notes.text(1, 0)).toBe('See the project website')
    expect(notes.style(2, 0).paddingLeft).toBeDefined()
    expect(notes.style(5, 0)).toMatchObject({ textDecoration: 'line-through', color: '#c00000', fontSize: '12pt' })
    expect(notes.style(5, 1).textRotate).toBe(90)
    expect(notes.progress.gridLines).toBe(false)

    const data = await sheet(2, 500)
    expect(data.progress.rows).toBe(401)
    expect(data.row(400, 2)).toEqual(['400', 'Item 400'])
  })

  test('workbook parts: defined names, hyperlinks, autofilter, tab color, page setup, header and footer', async () => {
    const part = await converted(await fixture('sample.ods'))
    const wb = await part('xl/workbook.xml')
    expect(wb).toContain('<definedName name="Revenue">Sales!$E$3:$E$62</definedName>')
    expect(wb).toContain('<definedName name="_xlnm._FilterDatabase" localSheetId="2" hidden="1">Data!$A$1:$D$401</definedName>')
    const notes = await part('xl/worksheets/sheet2.xml')
    expect(notes).toContain('<tabColor rgb="FFC00000"/>')
    expect(notes).toContain('<hyperlink ref="A2" r:id="rId1"/>')
    expect(await part('xl/worksheets/_rels/sheet2.xml.rels')).toContain('Target="https://example.com/" TargetMode="External"')
    expect(await part('xl/worksheets/sheet3.xml')).toContain('<autoFilter ref="A1:D401"/>')
    const sales = await part('xl/worksheets/sheet1.xml')
    expect(sales).toContain('<pageSetup paperSize="9" orientation="landscape"/>')
    expect(sales).toContain('<oddHeader>&amp;C&amp;A</oddHeader><oddFooter>&amp;CPage &amp;P of &amp;N</oddFooter>')
    expect(sales).toContain('<legacyDrawing r:id=')
    expect(await part('[Content_Types].xml')).toContain('<Override PartName="/xl/comments1.xml"')
  })

  test('saving writes an Excel workbook', async () => {
    const { engine } = await openBook(await fixture('sample.ods'))
    const xlsx = (await engine.call('save', [])) as Blob
    const again = await openBook(new Uint8Array(await xlsx.arrayBuffer()), 'book.xlsx')
    expect(again.info.converted).toBeUndefined()
    expect(again.info.xlsx!.sheets.map((s) => s.name)).toEqual(['Sales', 'Notes', 'Data', 'Lookup'])
    const s = await again.sheet(0)
    expect(s.text(1, 4)).toBe('Revenue')
    expect((await s.detail(62, 4))!.formula).toBe('=SUM(E3:E62)')
  })

  test('flat XML (.fods)', async () => {
    const bytes = await fixture('sample.fods')
    const conv = await convertFlatOdf(bytes)
    expect(conv.format).toBe('ods')
    const { info, sheet } = await openBook(bytes, 'sample.fods')
    expect(info.converted).toBe('ods')
    expect(info.xlsx!.sheets.map((s) => s.name)).toEqual(['Sales', 'Notes', 'Lookup'])
    const s = await sheet(0)
    expect(s.row(1, 3)).toEqual(['Region', 'Date', 'Units'])
    expect(s.text(14, 0)).toBe('Total')
    expect(s.progress.frozen).toEqual({ rows: 2, cols: 1 })
    expect(s.style(1, 0)).toMatchObject({ fontWeight: 700, background: '#4472c4' })
  })
})

describe('ods → xlsx: tables', () => {
  test("LibreOffice's huge trailing repeats are not expanded; repeated content is", async () => {
    const body =
      '<table:table table:name="Big"><table:table-column table:number-columns-repeated="16384"/>' +
      // a row with a value repeated over 4 columns, then empty cells to the last column
      `<table:table-row table:number-rows-repeated="3">${cellXml('office:value-type="float" office:value="7" table:number-columns-repeated="4"', '7')}<table:table-cell table:number-columns-repeated="16380"/></table:table-row>` +
      // a formatted row: its fill runs to the last column
      `<table:table-row><table:table-cell table:style-name="fill" table:number-columns-repeated="16384"/></table:table-row>` +
      '<table:table-row table:number-rows-repeated="1048000"><table:table-cell table:number-columns-repeated="16384"/></table:table-row>' +
      '<table:table-row table:style-name="tall" table:number-rows-repeated="572"><table:table-cell table:style-name="fill" table:number-columns-repeated="16384"/></table:table-row>' +
      '</table:table>'
    const auto =
      '<style:style style:name="fill" style:family="table-cell"><style:table-cell-properties fo:background-color="#ffcc00"/></style:style>' +
      '<style:style style:name="tall" style:family="table-row"><style:table-row-properties style:row-height="1cm"/></style:style>'
    const t0 = performance.now()
    const bytes = await ods(body, { auto })
    const { sheet } = await openBook(bytes)
    const s = await sheet(0)
    expect(performance.now() - t0).toBeLessThan(5000)
    expect(s.progress.rows).toBe(4)
    expect(s.progress.cols).toBe(4)
    expect(s.row(2, 5)).toEqual(['7', '7', '7', '7', ''])
    // the formatted row: cells up to the used width, the rest as the row's format
    expect(s.block.rows.find((r) => r.r === 3)!.cells.length).toBe(4)
    expect(s.style(3, 0).background).toBe('#ffcc00')
    const xml = await (await converted(bytes))('xl/worksheets/sheet1.xml')
    expect(xml).toMatch(/<row r="4" s="\d+" customFormat="1">/)
    expect(xml.length).toBeLessThan(5000)
  })

  test('rows and columns: heights, widths, hidden, groups, default cell styles, merges', async () => {
    const auto =
      '<style:style style:name="co1" style:family="table-column"><style:table-column-properties style:column-width="1in"/></style:style>' +
      '<style:style style:name="co2" style:family="table-column"><style:table-column-properties style:column-width="2in"/></style:style>' +
      '<style:style style:name="ro1" style:family="table-row"><style:table-row-properties style:row-height="0.5in" style:use-optimal-row-height="false"/></style:style>' +
      '<number:number-style style:name="N2"><number:number number:decimal-places="2" number:min-integer-digits="1" number:grouping="true"/></number:number-style>' +
      '<style:style style:name="money" style:family="table-cell" style:data-style-name="N2"/>'
    const body =
      '<table:table table:name="S"><table:table-column table:style-name="co1"/>' +
      '<table:table-column-group><table:table-column table:style-name="co2" table:visibility="collapse" table:default-cell-style-name="money" table:number-columns-repeated="2"/></table:table-column-group>' +
      `<table:table-row table:style-name="ro1">${cellXml('office:value-type="string" table:number-columns-spanned="2" table:number-rows-spanned="2"', 'Merged')}<table:covered-table-cell/>${cellXml('office:value-type="float" office:value="1234.5"', '1,234.50')}</table:table-row>` +
      `<table:table-row-group><table:table-row table:visibility="collapse"><table:covered-table-cell table:number-columns-repeated="2"/>${cellXml('office:value-type="float" office:value="2"', '2.00')}</table:table-row></table:table-row-group>` +
      '</table:table>'
    const { sheet } = await openBook(await ods(body, { auto }))
    const s = await sheet(0)
    expect(s.progress.merges).toEqual([{ r0: 0, c0: 0, r1: 1, c1: 1 }])
    expect(s.text(0, 0)).toBe('Merged')
    // the column's default cell style formats the value
    expect(s.text(0, 2)).toBe('1,234.50')
    expect(s.progress.colInfo!.find((c) => c.min === 0)!.w).toBe(96)
    expect(s.progress.colInfo!.find((c) => c.min === 1)!.w).toBe(0)
    expect(s.progress.rowSizes).toContainEqual([0, 48])
    expect(s.progress.rowSizes).toContainEqual([1, 0])
    const xml = await (await converted(await ods(body, { auto })))('xl/worksheets/sheet1.xml')
    expect(xml).toMatch(/<row r="2" hidden="1" outlineLevel="1">/)
    expect(xml).toMatch(/<col min="2" max="3" width="[\d.]+"[^>]* style="\d+" hidden="1" outlineLevel="1"\/>/)
  })

  test('values: dates with the null date, times, booleans, percentages, errors, paragraphs', async () => {
    const row = (...cells: string[]) => `<table:table-row>${cells.join('')}</table:table-row>`
    const body =
      '<table:calculation-settings><table:null-date table:date-value="1904-01-01"/></table:calculation-settings>' +
      '<table:table table:name="V">' +
      row(cellXml('office:value-type="date" office:date-value="2024-03-01"', '03/01/24'), cellXml('office:value-type="date" office:date-value="2024-03-01T13:30:00"', 'x')) +
      row(cellXml('office:value-type="time" office:time-value="PT36H15M00S"', '36:15:00'), cellXml('office:value-type="boolean" office:boolean-value="true"', 'TRUE')) +
      row(cellXml('office:value-type="percentage" office:value="0.25"', '25%'), cellXml('table:formula="of:=1/0" office:value-type="float" office:value="0" calcext:value-type="error"', '#DIV/0!')) +
      row(`<table:table-cell office:value-type="string"><text:p>one<text:s text:c="2"/>two</text:p><text:p>three</text:p></table:table-cell>`, cellXml('table:formula="of:=&quot;a&quot;&amp;&quot;b&quot;" office:value-type="string" office:string-value="ab"', 'ab')) +
      row(cellXml('office:value-type="currency" office:currency="EUR" office:value="12.5"', '12.50 €')) +
      '</table:table>'
    const { info, sheet } = await openBook(await ods(body))
    expect(info.xlsx!.date1904).toBe(true)
    const s = await sheet(0)
    expect(s.row(0, 2)).toEqual(['2024-03-01', '2024-03-01 13:30:00'])
    expect(s.row(1, 2)).toEqual(['36:15:00', 'TRUE'])
    expect(s.row(2, 2)).toEqual(['25.00%', '#DIV/0!'])
    expect(s.row(3, 2)).toEqual(['one  two\nthree', 'ab'])
    expect((await s.detail(3, 1))!.formula).toBe('="a"&"b"')
    expect(s.text(4, 0)).toBe('€12.50')

    // the usual null date: Excel's serials (and its 29 February 1900)
    const old = await openBook(await ods(`<table:table table:name="D">${row(cellXml('office:value-type="date" office:date-value="1900-01-01"', ''), cellXml('office:value-type="date" office:date-value="2026-01-15"', ''))}</table:table>`))
    expect(old.info.xlsx!.date1904).toBe(false)
    const d = await old.sheet(0)
    expect((await d.detail(0, 0))!.value).toBe('1')
    expect((await d.detail(0, 1))!.value).toBe('46037')
  })

  test('cell styles: fonts, fills, borders, alignment, rotation, inheritance', async () => {
    const styles =
      '<style:default-style style:family="table-cell"><style:text-properties style:font-name="Arial" fo:font-size="11pt"/></style:default-style>' +
      '<style:style style:name="Default" style:family="table-cell"/>' +
      '<style:style style:name="Heading" style:family="table-cell" style:parent-style-name="Default"><style:text-properties fo:font-weight="bold" fo:font-size="14pt"/></style:style>'
    const auto =
      '<style:style style:name="ce1" style:family="table-cell" style:parent-style-name="Heading"><style:table-cell-properties fo:background-color="#e2efda" fo:border="0.06pt solid #000000" fo:border-left="2.5pt dashed #ff0000" style:vertical-align="middle" fo:wrap-option="wrap" style:text-align-source="fix"/><style:paragraph-properties fo:text-align="end"/><style:text-properties fo:color="#2f5597" fo:font-style="italic" style:text-underline-style="solid" style:text-underline-type="double"/></style:style>' +
      '<style:style style:name="ce2" style:family="table-cell"><style:table-cell-properties style:rotation-angle="45" style:diagonal-tl-br="0.75pt solid #00ff00"/><style:text-properties style:text-line-through-style="solid" style:text-position="super 58%"/></style:style>'
    const body = `<table:table table:name="S"><table:table-row>${cellXml('table:style-name="ce1" office:value-type="string"', 'Styled')}${cellXml('table:style-name="ce2" office:value-type="string"', 'Turned')}${cellXml('office:value-type="string"', 'Plain')}</table:table-row></table:table>`
    const bytes = await ods(body, { auto, styles })
    const { info, sheet } = await openBook(bytes)
    const s = await sheet(0)
    expect(info.xlsx!.baseFont).toMatchObject({ fontSize: '11pt' })
    expect(s.style(0, 0)).toMatchObject({ fontWeight: 700, fontSize: '14pt', fontStyle: 'italic', color: '#2f5597', background: '#e2efda', textAlign: 'right', alignItems: 'center', whiteSpace: 'pre-wrap' })
    expect(s.style(0, 0).borderLeft).toBe('2px dashed #ff0000')
    expect(s.style(0, 0).borderTop).toBe('1px solid #000000')
    expect(s.style(0, 1)).toMatchObject({ textRotate: 45, textDecoration: 'line-through' })
    expect(s.style(0, 2)).toEqual({})
    const xml = await (await converted(bytes))('xl/styles.xml')
    expect(xml).toContain('<u val="double"/>')
    expect(xml).toContain('<vertAlign val="superscript"/>')
    expect(xml).toContain('<border diagonalDown="1">')
  })

  test('sheet names Excel does not take are renamed, and the formulas and names follow', async () => {
    const long = 'Q1: Sales/Costs [draft] with a name longer than Excel allows'
    const body =
      `<table:table table:name="${long}"><table:table-row>${cellXml('office:value-type="float" office:value="5"', '5')}</table:table-row></table:table>` +
      `<table:table table:name="Sum"><table:table-row>${cellXml(`table:formula="of:=['${long}'.A1]*2" office:value-type="float" office:value="10"`, '10')}</table:table-row>` +
      `<table:named-expressions><table:named-range table:name="Local" table:base-cell-address="$Sum.$A$1" table:cell-range-address="$Sum.$A$1"/></table:named-expressions></table:table>` +
      `<table:named-expressions><table:named-range table:name="First" table:base-cell-address="$'${long}'.$A$1" table:cell-range-address="$'${long}'.$A$1"/><table:named-expression table:name="Twice" table:expression="of:=[$Sum.$A$1]*2"/></table:named-expressions>`
    const bytes = await ods(body)
    const { info, sheet } = await openBook(bytes)
    const name = info.xlsx!.sheets[0].name
    expect(name).toBe('Q1_ Sales_Costs _draft_ with a')
    const s = await sheet(1)
    expect((await s.detail(0, 0))!.formula).toBe(`='${name}'!A1*2`)
    const wb = await (await converted(bytes))('xl/workbook.xml')
    expect(wb).toContain(`<definedName name="First">'${name}'!$A$1</definedName>`)
    expect(wb).toContain('<definedName name="Twice">Sum!$A$1*2</definedName>')
    expect(wb).toContain('<definedName name="Local" localSheetId="1">Sum!$A$1</definedName>')
  })

  test('settings: frozen rows, active sheet, gridlines, zoom; hidden and right-to-left sheets', async () => {
    const auto =
      '<style:style style:name="ta1" style:family="table"><style:table-properties table:display="true"/></style:style>' +
      '<style:style style:name="taHidden" style:family="table"><style:table-properties table:display="false"/></style:style>' +
      '<style:style style:name="taRtl" style:family="table"><style:table-properties table:display="true" style:writing-mode="rl-tb"/></style:style>'
    const one = (name: string, style: string) => `<table:table table:name="${name}" table:style-name="${style}"><table:table-row>${cellXml('office:value-type="string"', name)}</table:table-row></table:table>`
    const item = (n: string, v: string | number, t = 'int') => `<config:config-item config:name="${n}" config:type="${t}">${v}</config:config-item>`
    const settings =
      '<config:config-item-set config:name="ooo:view-settings"><config:config-item-map-indexed config:name="Views"><config:config-item-map-entry>' +
      '<config:config-item-map-named config:name="Tables">' +
      `<config:config-item-map-entry config:name="Two">${item('VerticalSplitMode', 2, 'short')}${item('VerticalSplitPosition', 3)}${item('PositionBottom', 10)}${item('ZoomValue', 150)}${item('ShowGrid', 'false', 'boolean')}</config:config-item-map-entry>` +
      '</config:config-item-map-named>' +
      `${item('ActiveTable', 'Two', 'string')}</config:config-item-map-entry></config:config-item-map-indexed></config:config-item-set>`
    const bytes = await ods(one('One', 'ta1') + one('Two', 'taRtl') + one('Three', 'taHidden'), { auto, settings })
    const { info, sheet } = await openBook(bytes)
    expect(info.xlsx!.sheets.map((s) => s.hidden)).toEqual([false, false, true])
    const s = await sheet(1)
    expect(s.progress.frozen).toEqual({ rows: 3, cols: 0 })
    expect(s.progress.gridLines).toBe(false)
    const part = await converted(bytes)
    expect(await part('xl/workbook.xml')).toContain('<workbookView activeTab="1"/>')
    const two = await part('xl/worksheets/sheet2.xml')
    expect(two).toContain('<sheetView workbookViewId="0" tabSelected="1" showGridLines="0" rightToLeft="1" zoomScale="150">')
    expect(two).toContain('<pane ySplit="3" topLeftCell="A11" activePane="bottomLeft" state="frozen"/>')
  })

  test('hyperlinks into the workbook and comments shown on the sheet', async () => {
    const link = '<table:table-cell office:value-type="string"><text:p><text:a xlink:href="#Two.B2" xlink:type="simple">Go</text:a></text:p></table:table-cell>'
    const note = '<office:annotation office:display="true"><dc:creator>Ann</dc:creator><text:p>First line</text:p><text:p>Second line</text:p></office:annotation>'
    const body =
      `<table:table table:name="One"><table:table-row>${link}<table:table-cell office:value-type="float" office:value="3">${note}<text:p>3</text:p></table:table-cell></table:table-row></table:table>` +
      `<table:table table:name="Two"><table:table-row>${cellXml('office:value-type="string"', 'x')}</table:table-row></table:table>`
    const bytes = await ods(body)
    const { engine } = await openBook(bytes)
    const comments = (await engine.call('comments.list', [])) as DocComment[]
    expect(comments.map((c) => [c.author.name, c.text, cellOf(c)])).toEqual([['Ann', 'First line\nSecond line', [0, 1]]])
    const part = await converted(bytes)
    expect(await part('xl/worksheets/sheet1.xml')).toContain('<hyperlink ref="A1" location="Two!B2" display="Two.B2"/>')
    expect(await part('xl/drawings/vmlDrawing1.vml')).toContain('visibility:visible')
  })

  test('password-protected packages are refused with a clear error', async () => {
    const manifest =
      '<?xml version="1.0" encoding="UTF-8"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/>' +
      '<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"><manifest:encryption-data manifest:checksum-type="SHA1/1K" manifest:checksum="x"/></manifest:file-entry></manifest:manifest>'
    const files = new Map<string, string>([
      ['mimetype', 'application/vnd.oasis.opendocument.spreadsheet'],
      ['content.xml', '\u0013§garbled<<'],
      ['META-INF/manifest.xml', manifest],
    ])
    const bytes = new Uint8Array(await (await writeZip(files)).arrayBuffer())
    const engine = new Engine(() => {})
    await expect(engine.open(new Blob([bytes]), { name: 'locked.ods' })).rejects.toThrow('Password-protected OpenDocument files are not supported')
  })
})

describe('ods → xlsx: formulas and number formats', () => {
  test('OpenFormula to Excel syntax', () => {
    const cases: [string, string | undefined][] = [
      ['of:=SUM([.A1:.B3])', 'SUM(A1:B3)'],
      ['of:=[$Sheet2.A1]+[.$B$2]', 'Sheet2!A1+$B$2'],
      ["of:=SUM(['My Sheet'.A1:.C3])", "SUM('My Sheet'!A1:C3)"],
      ["of:=[$'O''Neil'.A1]", "'O''Neil'!A1"],
      ['of:=SUM([Sheet1.A1:Sheet3.B2])', 'SUM(Sheet1:Sheet3!A1:B2)'],
      ['of:=IF([.A1]>0;"yes";"no")', 'IF(A1>0,"yes","no")'],
      ['of:={1;2|3;4}', '{1,2;3,4}'],
      ['of:=SUM([.A1:.A3]~[.C1:.C3])', 'SUM(A1:A3,C1:C3)'],
      ['of:=SUM([.A1:.C3]![.B2:.D4])', 'SUM(A1:C3 B2:D4)'],
      ['of:=SUM([.A:.A]) + SUM([.1:.2])', 'SUM(A:A)+SUM(1:2)'],
      ['of:=COM.MICROSOFT.TEXTJOIN(",";1;[.A1:.A3])', '_xlfn.TEXTJOIN(",",1,A1:A3)'],
      ['of:=COM.MICROSOFT.FILTER([.A1:.A9];[.B1:.B9]>0)', '_xlfn._xlws.FILTER(A1:A9,B1:B9>0)'],
      ['of:=IFNA(VLOOKUP([.A1];[.B1:.C9];2;0);"")', '_xlfn.IFNA(VLOOKUP(A1,B1:C9,2,0),"")'],
      ['of:=CEILING([.A1];2)', '_xlfn.CEILING.MATH(A1,2)'],
      ['of:=LEGACY.NORMSDIST(1)', 'NORMSDIST(1)'],
      ['of:=([.A2]=[.C2])AND([.B2]=[.D2])', 'AND((A2=C2),(B2=D2))'],
      ['of:="say ""hi"" ; [.A1]"', '"say ""hi"" ; [.A1]"'],
      ['of:=[#REF!]+1', '#REF!+1'],
      ["of:=['file:///x.ods'#$Sheet1.A1]", '#REF!'],
      ['oooc:=SUM([.A1:.B2])', 'SUM(A1:B2)'],
      ['msoxl:=SUM(A1:B2)', 'SUM(A1:B2)'],
      ['of:=TRUE()', 'TRUE()'],
      ['of:Err:502', undefined],
      ['of:=SUM(1;2', undefined],
    ]
    for (const [odf, xl] of cases) expect([odf, excelFormula(odf)]).toEqual([odf, xl])
    expect(excelFormula("of:=['Old name'.A1]", (s) => (s === 'Old name' ? 'New' : s))).toBe('New!A1')
  })

  test('data styles to Excel number formats', () => {
    const content = parseOdfXml(
      `<office:document-content ${NS}><office:automatic-styles>` +
        '<number:number-style style:name="int"><number:number number:decimal-places="0" number:min-integer-digits="1" number:grouping="true"/></number:number-style>' +
        '<number:number-style style:name="dec"><number:number number:decimal-places="3" number:min-decimal-places="1" number:min-integer-digits="1"/></number:number-style>' +
        '<number:number-style style:name="gen"><number:number number:min-integer-digits="1"/></number:number-style>' +
        '<number:number-style style:name="k"><number:number number:decimal-places="1" number:min-integer-digits="1" number:display-factor="1000"/><number:text>K</number:text></number:number-style>' +
        '<number:number-style style:name="sci"><number:scientific-number number:decimal-places="2" number:min-integer-digits="1" number:min-exponent-digits="2"/></number:number-style>' +
        '<number:number-style style:name="frac"><number:fraction number:min-integer-digits="0" number:min-numerator-digits="1" number:min-denominator-digits="2"/></number:number-style>' +
        '<number:number-style style:name="phone"><number:number number:decimal-places="0" number:min-integer-digits="7"><number:embedded-text number:position="4">-</number:embedded-text></number:number></number:number-style>' +
        '<number:number-style style:name="negP0"><number:number number:decimal-places="2" number:min-integer-digits="1" number:grouping="true"/></number:number-style>' +
        '<number:number-style style:name="neg"><style:text-properties fo:color="#ff0000"/><number:text>-</number:text><number:number number:decimal-places="2" number:min-integer-digits="1" number:grouping="true"/><style:map style:condition="value()&gt;=0" style:apply-style-name="negP0"/></number:number-style>' +
        '<number:percentage-style style:name="pct"><number:number number:decimal-places="2" number:min-integer-digits="1"/><number:text>%</number:text></number:percentage-style>' +
        '<number:currency-style style:name="eurP0"><number:number number:decimal-places="2" number:min-integer-digits="1" number:grouping="true"/><number:text> </number:text><number:currency-symbol>€</number:currency-symbol></number:currency-style>' +
        '<number:currency-style style:name="eur"><number:text>-</number:text><number:number number:decimal-places="2" number:min-integer-digits="1" number:grouping="true"/><number:text> </number:text><number:currency-symbol>€</number:currency-symbol><style:map style:condition="value()&gt;=0" style:apply-style-name="eurP0"/></number:currency-style>' +
        '<number:date-style style:name="date"><number:day-of-week number:style="long"/><number:text>, </number:text><number:day/><number:text>. </number:text><number:month number:textual="true" number:style="long"/><number:text> </number:text><number:year number:style="long"/></number:date-style>' +
        '<number:date-style style:name="dt"><number:day number:style="long"/><number:text>.</number:text><number:month number:style="long"/><number:text>.</number:text><number:year/><number:text> </number:text><number:hours number:style="long"/><number:text>:</number:text><number:minutes number:style="long"/><number:text> </number:text><number:am-pm/></number:date-style>' +
        '<number:time-style style:name="dur" number:truncate-on-overflow="false"><number:hours/><number:text>:</number:text><number:minutes number:style="long"/><number:text>:</number:text><number:seconds number:style="long" number:decimal-places="2"/></number:time-style>' +
        '<number:boolean-style style:name="bool"><number:boolean/></number:boolean-style>' +
        '<number:text-style style:name="txt"><number:text>Code </number:text><number:text-content/></number:text-style>' +
        '<number:number-style style:name="three"><number:number number:decimal-places="0" number:min-integer-digits="1"/><style:map style:condition="value()&gt;0" style:apply-style-name="int"/><style:map style:condition="value()&lt;0" style:apply-style-name="neg"/></number:number-style>' +
        '</office:automatic-styles></office:document-content>',
    )
    const st = new OdfStyles(undefined, content)
    const f = (name: string) => excelFormat(st, name)
    expect(f('int')).toBe('#,##0')
    expect(f('dec')).toBe('0.0##')
    expect(f('gen')).toBeUndefined()
    expect(f('k')).toBe('0.0,"K"')
    expect(f('sci')).toBe('0.00E+00')
    expect(f('frac')).toBe('# ?/??')
    expect(f('phone')).toBe('000-0000')
    expect(f('neg')).toBe('#,##0.00;[Red]-#,##0.00')
    expect(f('pct')).toBe('0.00%')
    expect(f('eur')).toBe('#,##0.00 [$€];-#,##0.00 [$€]')
    expect(f('date')).toBe('dddd, d". "mmmm yyyy')
    expect(f('dt')).toBe('dd"."mm"."yy hh:mm AM/PM')
    expect(f('dur')).toBe('[h]:mm:ss.00')
    expect(f('bool')).toBe('"TRUE";"TRUE";"FALSE"')
    expect(f('txt')).toBe('"Code "@')
    expect(f('three')).toBe('#,##0;[Red]-#,##0.00;0')
  })
})
