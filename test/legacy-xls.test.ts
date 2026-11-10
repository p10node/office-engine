import { describe, expect, test } from 'bun:test'
import { readFileSync } from 'fs'
import { applyTint, INDEXED } from '../src/core/color'
import { convertCfb } from '../src/core/legacy'
import { biffToXlsx } from '../src/core/legacy/xls'
import { latin1 } from '../src/core/legacy/xls/biff'
import { formulaText, type FormulaEnv } from '../src/core/legacy/xls/formula'
import { BlobSource } from '../src/core/source'
import { Engine } from '../src/core/worker/engine'
import type { BlockOut, SheetProgress } from '../src/core/xlsx/types'
import { ZipArchive } from '../src/core/zip'
import { writeZip } from '../src/core/zipwrite'
import { writeCfb } from './cfb-writer'
import {
  blank, boolErr, cat, colInfo, EOF, f64, formula, label, labelSst, merge, mulBlank, mulRk, number, pane, ptg, rec,
  rk, row, str16, stringRec, u16, u32, u8, window2, workbookStream, xlsFile,
} from './xls-builder'

const fixture = (name: string) => new Uint8Array(readFileSync(new URL(`./fixtures/legacy/${name}`, import.meta.url)))

async function openXls(bytes: Uint8Array<ArrayBuffer>) {
  const progress: SheetProgress[] = []
  const engine = new Engine((event, data) => {
    if (event === 'xlsx:progress') progress.push(data as SheetProgress)
  })
  const info = await engine.open(new Blob([bytes]), { name: 'book.xls' })
  const sheet = async (index: number, rows = 60) => {
    await engine.call('xlsx.load', [index])
    // formulas without a result in the file are computed after the load
    await new Promise((r) => setTimeout(r, 50))
    const block = (await engine.call('xlsx.block', [index, 0, rows])) as BlockOut
    const cell = (r: number, c: number) => block.rows.find((x) => x.r === r)?.cells.find((x) => x[0] === c)
    const mine = progress.filter((p) => p.sheet === index)
    return {
      cell,
      text: (r: number, c: number) => cell(r, c)?.[1] ?? '',
      style: (r: number, c: number) => info.xlsx!.styles[cell(r, c)?.[2] ?? 0],
      /** the layout the view gets: the first progress message has the header, the last the merges */
      head: mine[0],
      last: mine[mine.length - 1],
      /** row heights (px, 0 hidden) */
      sizes: new Map(mine.flatMap((p) => p.rowSizes)),
      formula: async (r: number, c: number) => ((await engine.call('xlsx.cell', [index, r, c])) as { formula?: string } | undefined)?.formula,
    }
  }
  return { engine, info, sheet }
}

/** the parts of the XLSX file a workbook converts to */
async function xlsxParts(bytes: Uint8Array) {
  const { file } = await convertCfb(bytes)
  const zip = await ZipArchive.open(new BlobSource(file))
  return { text: (path: string) => zip.text(path), has: (path: string) => zip.has(path) }
}

describe('xls: formulas', () => {
  const env: FormulaEnv = {
    biff: 8,
    sheetRef: (i) => ['Sheet1!', "'My Sheet'!", 'Sheet1:Data!', ''][i] ?? null,
    sheetRef5: () => null,
    name: (i) => ['Rate', '_xlfn.IFERROR'][i - 1],
    nameX: (_, i) => ['EDATE'][i - 1],
    text8: (b) => latin1(b),
  }
  const f = (tokens: number[], extra?: number[], r = 0, c = 0, rel = false) => formulaText(env, Uint8Array.from(tokens), extra && Uint8Array.from(extra), r, c, rel, 0)

  test('operators, operands and precedence as written', () => {
    // A1+B$2*-$C3%
    expect(f(cat(ptg.ref(0, 0), ptg.ref(1, 1, false, true), ptg.ref(2, 2, true, false), ptg.op(0x13), ptg.op(0x14), ptg.op(0x05), ptg.op(0x03)))).toBe('A1+B$2*-$C3%')
    // ("a"&"b")<>"say ""hi"""
    expect(f(cat(ptg.str('a'), ptg.str('b'), ptg.op(0x08), ptg.paren(), ptg.str('say "hi"'), ptg.op(0x0e)))).toBe('("a"&"b")<>"say ""hi"""')
    expect(f(cat(ptg.num(1.5), ptg.int(2), ptg.op(0x07), ptg.bool(true), ptg.op(0x0b)))).toBe('1.5^2=TRUE')
    expect(f(cat(ptg.err(0x07)))).toBe('#DIV/0!')
    // ranges, unions and intersections of references
    expect(f(cat(ptg.area(0, 3, 0, 1, false), ptg.ref(5, 2), ptg.op(0x0f)))).toBe('$A$1:$B$4 C6')
    // whole columns and rows of the 65,536 × 256 grid
    expect(f(cat(ptg.area(0, 65535, 0, 1)))).toBe('A:B')
    expect(f(cat(ptg.area(4, 4, 0, 255, false)))).toBe('$5:$5')
  })

  test('functions: fixed and variable arguments, SUM / IF / CHOOSE attributes, spaces, missing arguments', () => {
    // ROUND(A1,2)
    expect(f(cat(ptg.ref(0, 0), ptg.int(2), ptg.func(27)))).toBe('ROUND(A1,2)')
    // SUM(A1:A3) written with the one-argument SUM attribute
    expect(f(cat(ptg.area(0, 2, 0, 0), ptg.attrSum()))).toBe('SUM(A1:A3)')
    // IF(A1, ,1): the IF jump, a missing argument, a space before the next token
    expect(f(cat(ptg.ref(0, 0), ptg.attrIf(4), ptg.missArg(), ptg.attrSpace(1), ptg.int(1), ptg.funcVar(1, 3)))).toBe('IF(A1,,1)')
    // CHOOSE(2,"a","b") with its jump table
    expect(f(cat(ptg.int(2), ptg.attrChoose([8, 12, 16]), ptg.str('a'), ptg.str('b'), ptg.funcVar(100, 3)))).toBe('CHOOSE(2,"a","b")')
    // NOW() and a function the table has no entry for (a macro command)
    expect(f(ptg.func(74))).toBe('NOW()')
    expect(() => f(cat(ptg.int(1), [0x42, 1, ...u16(0x8001)]))).toThrow()
  })

  test('3D references, names, add-in and Excel 2007 functions', () => {
    expect(f(cat(ptg.area3d(1, 0, 1, 0, 1, false), ptg.funcVar(4, 1)))).toBe("SUM('My Sheet'!$A$1:$B$2)")
    expect(f(ptg.ref3d(2, 2, 2))).toBe('Sheet1:Data!C3')
    expect(f(cat(ptg.name(1), ptg.int(2), ptg.op(0x05)))).toBe('Rate*2')
    // EDATE(A1,1): an add-in function, named by its first argument
    expect(f(cat(ptg.nameX(4, 1), ptg.ref(0, 0), ptg.int(1), ptg.funcVar(255, 3)))).toBe('EDATE(A1,1)')
    // IFERROR is stored as a "_xlfn." name in an .xls file
    expect(f(cat(ptg.name(2), ptg.ref(0, 0), ptg.int(0), ptg.funcVar(255, 3)))).toBe('IFERROR(A1,0)')
    // a deleted sheet
    expect(f(ptg.ref3d(9, 0, 0))).toBe('#REF!')
  })

  test('shared formulas: references relative to the cell', () => {
    // in D6 (row 5, column 3): one row up and one column left, then the same row in absolute column A
    expect(f(cat(ptg.refN(-1, -1), ptg.refN(0, 0, true, false), ptg.op(0x03)), undefined, 5, 3, true)).toBe('C5+$A6')
    // the same tokens in row 32,770 (rows past 32,767 wrap as unsigned)
    expect(f(ptg.refN(1, 0), undefined, 32768, 1, true)).toBe('B32770')
  })

  test('array constants from the extra data', () => {
    const extra = cat(u8(1), u16(1), [1], f64(1), [1], f64(2), [2], str16('a'), [4, 1, 0, 0, 0, 0, 0, 0, 0], [16, 0x2a, 0, 0, 0, 0, 0, 0, 0])
    // 2 columns × 2 rows, then one more value read by nothing
    expect(f(cat(ptg.array(), ptg.funcVar(4, 1)), extra)).toBe('SUM({1,2;"a",TRUE})')
  })

  test('damaged token streams fail instead of reading past the end', () => {
    expect(() => f([0x24, 1])).toThrow()
    expect(() => f(cat(ptg.op(0x03)))).toThrow()
    expect(() => f(cat(ptg.int(1), ptg.int(2)))).toThrow()
    expect(() => f([0xff])).toThrow()
  })
})

describe('xls: the sample workbook (xlwt)', () => {
  test('values, number formats, formulas and their computed results', async () => {
    const { info, sheet } = await openXls(fixture('sample.xls'))
    expect(info.kind).toBe('xlsx')
    expect(info.converted).toBe('xls')
    expect(info.xlsx!.sheets.map((s) => s.name)).toEqual(['Sales', 'Notes', 'Data'])
    const s = await sheet(0)
    expect(s.text(0, 0)).toBe('Sales by region — FY2026')
    expect([0, 1, 2, 3, 4, 5].map((c) => s.text(1, c))).toEqual(['Region', 'Date', 'Units', 'Price', 'Revenue', 'Margin'])
    expect(s.text(2, 1)).toBe('2026-01-16')
    expect(s.text(2, 3)).toMatch(/^\$\d+\.\d\d$/)
    expect(s.text(2, 5)).toMatch(/^-?\d+\.\d%$/)
    // xlwt stores no results: the formulas are computed when the sheet loads
    expect(await s.formula(2, 4)).toBe('=C3*D3')
    const units = Number(s.text(2, 2))
    const price = Number(s.text(2, 3).slice(1))
    expect(s.text(2, 4)).toBe((units * price).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }))
    expect(await s.formula(26, 4)).toBe('=SUM(E3:E26)')
    expect(s.text(3, 7)).toBe('(42.50)')
    expect(s.cell(3, 7)?.[4]).toBe('#ff0000')
    expect(s.text(4, 7)).toBe('TRUE')
    expect(s.text(5, 7)).toBe('Sunday, December 31, 2023 6:00 PM')
    // references to other sheets
    expect(await s.formula(7, 7)).toBe('=Notes!B1&" "&Data!B2')
    expect(s.text(7, 7)).toBe('Italic Item 1')
    expect(await s.formula(8, 7)).toBe('=VLOOKUP("North",A3:F26,3,FALSE)')
  })

  test('the look: fonts, fills, borders, alignment, merged title, frozen panes, sizes', async () => {
    const { sheet } = await openXls(fixture('sample.xls'))
    const s = await sheet(0)
    expect(s.style(0, 0)).toMatchObject({ fontWeight: 700, fontSize: '14pt', color: '#000080', textAlign: 'center' })
    expect(s.style(1, 0)).toMatchObject({ fontWeight: 700, color: '#ffffff', background: '#000080', borderBottom: '2px solid #000000' })
    expect(s.style(2, 0)).toMatchObject({ borderLeft: '1px solid #000', borderTop: '1px solid #000' })
    expect(s.style(26, 0)).toMatchObject({ background: '#ffff99', borderTop: '3px double #000' })
    expect(s.last.merges).toEqual([{ r0: 0, c0: 0, r1: 0, c1: 5 }])
    expect(s.head.frozen).toEqual({ rows: 2, cols: 1 })
    expect(s.head.colInfo!.find((c) => c.min === 0)!.w).toBe(98)
    // the title row is 30 pt high
    expect(s.sizes.get(0)).toBe(40)

    const notes = await sheet(1)
    expect(notes.style(0, 0)).toMatchObject({ whiteSpace: 'pre-wrap', alignItems: 'flex-start' })
    expect(notes.style(0, 1)).toMatchObject({ fontStyle: 'italic' })
    expect(notes.style(1, 0)).toMatchObject({ fontWeight: 700, color: '#ff0000' })
    expect(notes.style(1, 1)).toMatchObject({ textDecoration: 'underline' })
    expect(notes.style(2, 0)).toMatchObject({ textDecoration: 'line-through' })
    expect(notes.style(2, 1)).toMatchObject({ background: '#ccffcc', textAlign: 'center' })
    expect(notes.style(3, 0)).toMatchObject({ textRotate: 45 })
    expect(notes.style(3, 1)).toMatchObject({ paddingLeft: '20px' })
    expect(notes.style(6, 0).fontFamily).toContain('Times New Roman')
    expect(notes.text(4, 0)).toBe('Rich: bold and blue italic')
    // hidden row 6 and column C
    expect(notes.sizes.get(5)).toBe(0)
    expect(notes.head.colInfo!.find((c) => c.min === 2)!.w).toBe(0)
  })

  test('page setup, header and footer, rich text runs', async () => {
    const x = await xlsxParts(fixture('sample.xls'))
    const xml = await x.text('xl/worksheets/sheet1.xml')
    expect(xml).toContain('<pageSetUpPr fitToPage="1"/>')
    expect(xml).toContain('<pageMargins left="0.5" right="0.5" top="0.61" bottom="0.37" header="0.1" footer="0.1"/>')
    expect(xml).toContain('orientation="landscape"')
    expect(xml).toContain('paperSize="9"')
    expect(xml).toContain('<oddHeader>&amp;C&amp;BSales by region</oddHeader><oddFooter>&amp;LConfidential&amp;RPage &amp;P of &amp;N</oddFooter>')
    const sst = await x.text('xl/sharedStrings.xml')
    expect(sst).toMatch(/<si><r><t xml:space="preserve">Rich: <\/t><\/r><r><rPr><rFont val="Arial"\/><b\/>(<color rgb="FF000000"\/>)?<sz val="10"\/><\/rPr><t>bold<\/t><\/r>/)
    // the parts Excel needs
    expect(await x.text('[Content_Types].xml')).toContain('/xl/worksheets/sheet3.xml')
    expect(await x.text('xl/styles.xml')).toContain('<numFmt numFmtId="')
  })

  test('saving writes an Excel workbook that opens again', async () => {
    const { engine } = await openXls(fixture('sample.xls'))
    const xlsx = (await engine.call('save', [])) as Blob
    const again = await openXls(new Uint8Array(await xlsx.arrayBuffer()))
    expect(again.info.converted).toBeUndefined()
    const s = await again.sheet(0)
    expect(s.text(0, 0)).toBe('Sales by region — FY2026')
    expect(await s.formula(2, 4)).toBe('=C3*D3')
  })
})

describe('xls: strings and dates', () => {
  test('a shared string table over CONTINUE records, 8-bit and 16-bit', async () => {
    const { sheet } = await openXls(fixture('strings.xls'))
    const s = await sheet(0, 400)
    const words = ['alpha', 'beta', 'gamma', 'delta', 'epsilon', 'zeta', 'eta', 'theta']
    for (let i = 0; i < 300; i++) {
      let t = `${String(i).padStart(4, '0')} ${[0, 1, 2, 3].map((k) => words[(i + k) % 8]).join(' ')}`
      if (i % 3 === 0) t += ' été 中文 Ω'
      expect(s.text(i, 0)).toBe(t)
    }
    expect(s.text(0, 1)).toBe(Array.from({ length: 1500 }, (_, k) => words[k % 8] + ' ').join(''))
    expect(s.text(1, 1)).toBe('中文'.repeat(2500))
  })

  test('the 1904 date system', async () => {
    const { info, sheet } = await openXls(fixture('dates1904.xls'))
    expect(info.xlsx!.date1904).toBe(true)
    const s = await sheet(0)
    expect(s.text(0, 0)).toBe('2024-02-29')
    expect(s.text(1, 0)).toBe('1 Jan 2001 12:30')
  })

  test('a string that changes from 8-bit to 16-bit characters at a CONTINUE record; rich text and phonetic data skipped', async () => {
    const sst = [
      // "Hello Wörld" split after "Hello " (8-bit), continued as 16-bit
      cat(u16(0x00fc), u16(8 + 3 + 6), u32(3), u32(3), u16(11), u8(0), [...'Hello '].map((ch) => ch.charCodeAt(0))),
      rec(0x003c, [1, ...[...'Wörld'].flatMap((ch) => u16(ch.charCodeAt(0)))]),
      // a rich string (2 runs) and a string with phonetic data, the next string's header at the start of a CONTINUE
      rec(0x003c, u16(4), u8(0x08), u16(2), [...'Bold'].map((ch) => ch.charCodeAt(0)), u16(0), u16(5), u16(2), u16(0)),
      rec(0x003c, u16(3), u8(0x04), u32(6), [...'End'].map((ch) => ch.charCodeAt(0)), [1, 2, 3, 4, 5, 6]),
    ]
    const file = xlsFile({ sst, fonts: [{}, {}, {}, {}, { bold: true }], sheets: [{ name: 'S', records: [labelSst(0, 0, 0), labelSst(1, 0, 1), labelSst(2, 0, 2)] }] })
    const { sheet } = await openXls(file)
    const s = await sheet(0)
    expect([s.text(0, 0), s.text(1, 0), s.text(2, 0)]).toEqual(['Hello Wörld', 'Bold', 'End'])
    const x = await xlsxParts(file)
    expect(await x.text('xl/sharedStrings.xml')).toContain('<si><r><rPr><rFont val="Arial"/><b/><sz val="10"/></rPr><t>Bo</t></r><r><rPr><rFont val="Arial"/><sz val="10"/></rPr><t>ld</t></r></si>')
  })
})

/** a workbook with what xlwt does not write */
function featureBook() {
  const strings = ['Name', 'Value', 'Total']
  const shared = cat(ptg.refN(-2, 0), ptg.int(2), ptg.op(0x05))
  const records = [
    window2({ frozen: true, selected: true }),
    pane(0, 1),
    colInfo(0, 0, 20),
    colInfo(1, 2, 12, false),
    row(0, 20, { custom: true }),
    labelSst(0, 0, 0, 16),
    labelSst(0, 1, 1, 16),
    label(1, 0, 'Apples'),
    rk(1, 1, 3),
    label(2, 0, 'Pears'),
    number(2, 1, 4.25),
    // B4:B6: a shared formula (=B2*2, =B3*2, =B4*2)
    mulRk(1, 3, [1, 2, 3]),
    formula(3, 1, ptg.exp(3, 1), 6),
    rec(0x04bc, u16(3), u16(5), u8(1), u8(1), u8(0), u8(3), u16(shared.length), shared),
    formula(4, 1, ptg.exp(3, 1), 8.5),
    formula(5, 1, ptg.exp(3, 1), 12),
    // D5:E5 an array formula {=B2:C2*10} with a string row of results
    formula(4, 3, ptg.exp(4, 3), 30),
    rec(0x0221, u16(4), u16(4), u8(3), u8(4), u16(0), u32(0), u16(13), cat([0x65, ...u16(1), ...u16(1), ...u16(1 | 0xc000), ...u16(2 | 0xc000)], ptg.int(10), ptg.op(0x05))),
    formula(4, 4, ptg.exp(4, 3), 0),
    // a string result in a STRING record, a boolean and an error result
    formula(6, 0, cat(ptg.ref(1, 0), ptg.str('!'), ptg.op(0x08)), [0, 0]),
    stringRec('Apples!'),
    formula(6, 1, cat(ptg.ref(1, 1), ptg.int(3), ptg.op(0x0b)), [1, 1]),
    formula(6, 2, cat(ptg.int(1), ptg.int(0), ptg.op(0x06)), [2, 0x07]),
    // =SUM(Other!A1:A2)+Rate, a reference to another workbook (value kept), a name
    formula(7, 1, cat(ptg.area3d(1, 0, 1, 0, 0), ptg.attrSum(), ptg.name(1), ptg.op(0x03)), 11),
    formula(7, 2, ptg.ref3d(2, 0, 0), 99),
    labelSst(7, 0, 2, 17),
    boolErr(8, 0, 1, false),
    boolErr(8, 1, 0x2a, true),
    blank(8, 2, 17),
    mulBlank(9, 0, [17, 17, 15]),
    merge([10, 11, 0, 2]),
    // comments: OBJ (a note) + TXO with its text in CONTINUE records, then NOTE
    rec(0x005d, u16(0x15), u16(0x12), u16(0x19), u16(1), u16(0x4011), u32(0), u32(0), u32(0), u32(0), u16(0), u16(0)),
    rec(0x01b6, u16(0x0212), u16(0), u16(0), u32(0), u16(17), u16(16), u16(0), u32(0)),
    rec(0x003c, chars8('Ann Lee:\nCheck it')),
    rec(0x003c, u16(0), u16(0), u32(0), u16(18), u16(0), u32(0)),
    rec(0x001c, u16(1), u16(1), u16(0), u16(1), str16('Ann Lee'), u8(0)),
    // hyperlinks: a URL and a place in the workbook
    hlinkUrl(1, 0, 'https://example.com/apples', 'Apples'),
    rec(0x0800, u16(0x0800), u16(1), u16(1), u16(0), u16(0), [...'Open the page'].flatMap((ch) => u16(ch.charCodeAt(0))), u16(0)),
    hlinkPlace(2, 0, 'Other!A1'),
    label(12, 0, 'Themed', 18),
    // tab color: palette entry 10 (red)
    rec(0x0862, u16(0x0862), u16(0), u32(0), u32(0), u32(0x14), u32(10), u32(0)),
  ]
  const extra = [
    rec(0x01ae, u16(2), u16(0x0401)),
    rec(0x01ae, u16(1), u16(5), u8(0), [...'other'].map((c) => c.charCodeAt(0)), str16('Sheet1')),
    rec(0x0017, u16(3), u16(0), u16(0), u16(0), u16(0), u16(1), u16(1), u16(1), u16(0), u16(0)),
    rec(0x0018, u16(0), u8(0), u8(4), u16(9), u16(0), u16(0), u32(0), chars8('Rate'), ptg.num(0.5)),
    rec(0x0018, u16(0x20), u8(0), u8(1), u16(11), u16(0), u16(1), u32(0), [0, 6], ptg.area3d(0, 0, 9, 0, 2, false)),
    // XFEXT (Excel 2007+): XF 18's fill is theme color accent 1, 80% lighter, its text #123456
    rec(0x087d, u16(0x087d), u16(0), u32(0), u32(0), u16(0), u16(18), u16(0), u16(2), cat(u16(4), u16(20), u16(3), u16(26213), u32(4), u32(0), u32(0)), cat(u16(13), u16(20), u16(2), u16(0), u32(0x563412), u32(0), u32(0))),
    // a custom palette: color 12 (blue by default) becomes orange
    rec(0x0092, u16(56), ...INDEXED.slice(8).map((h, i) => (i === 4 ? [0xff, 0x88, 0x00, 0] : [0, 2, 4].map((k) => parseInt(h.slice(1 + k, 3 + k), 16)).concat(0)))),
  ]
  return xlsFile({
    sst: strings,
    fonts: [{}, {}, {}, {}, { bold: true }],
    xfs: [{ font: 5, pattern: 1, fore: 12, border: 8 }, { fmt: 4, align: 2 }, { pattern: 1, fore: 12 }],
    sheets: [
      { name: 'Main', records },
      { name: 'Other', records: [number(0, 0, 5), number(1, 0, 6)] },
      { name: 'Chart1', type: 2, records: [] },
      { name: 'Secret', state: 2, records: [label(0, 0, 'hidden')] },
    ],
    extra,
  })
}

/** 8-bit characters after a flag byte (no length) */
function chars8(s: string) {
  return [0, ...[...s].map((c) => c.charCodeAt(0))]
}

const URL_CLSID = [0xe0, 0xc9, 0xea, 0x79, 0xf9, 0xba, 0xce, 0x11, 0x8c, 0x82, 0x00, 0xaa, 0x00, 0x4b, 0xa9, 0x0b]
const STD_LINK = [0xd0, 0xc9, 0xea, 0x79, 0xf9, 0xba, 0xce, 0x11, 0x8c, 0x82, 0x00, 0xaa, 0x00, 0x4b, 0xa9, 0x0b]
const hstr = (s: string) => cat(u32(s.length + 1), ...[...s].map((c) => u16(c.charCodeAt(0))), u16(0))

function hlinkUrl(r: number, c: number, url: string, display: string) {
  const u = cat(...[...url].map((ch) => u16(ch.charCodeAt(0))), u16(0))
  return rec(0x01b8, u16(r), u16(r), u16(c), u16(c), STD_LINK, u32(2), u32(0x17), hstr(display), URL_CLSID, u32(u.length), u)
}

function hlinkPlace(r: number, c: number, place: string) {
  return rec(0x01b8, u16(r), u16(r), u16(c), u16(c), STD_LINK, u32(2), u32(0x1c), hstr(place), hstr(place))
}

describe('xls: records xlwt does not write', () => {
  test('shared and array formulas, string / boolean / error results, 3D references and names', async () => {
    const { info, sheet } = await openXls(featureBook())
    expect(info.xlsx!.sheets.map((s) => [s.name, s.hidden])).toEqual([
      ['Main', false],
      ['Other', false],
      ['Chart1', false],
      ['Secret', true],
    ])
    const s = await sheet(0)
    expect([await s.formula(3, 1), await s.formula(4, 1), await s.formula(5, 1)]).toEqual(['=B2*2', '=B3*2', '=B4*2'])
    expect([s.text(3, 1), s.text(4, 1), s.text(5, 1)]).toEqual(['6', '8.5', '12'])
    expect(await s.formula(4, 3)).toBe('=B2:C2*10')
    expect(await s.formula(4, 4)).toBeUndefined()
    expect([s.text(4, 3), s.text(4, 4)]).toEqual(['30', '0'])
    expect([s.text(6, 0), s.text(6, 1), s.text(6, 2)]).toEqual(['Apples!', 'TRUE', '#DIV/0!'])
    expect(await s.formula(6, 0)).toBe('=A2&"!"')
    expect(await s.formula(7, 1)).toBe('=SUM(Other!A1:A2)+Rate')
    // another workbook: the value only
    expect(await s.formula(7, 2)).toBeUndefined()
    expect(s.text(7, 2)).toBe('99')
    expect([s.text(8, 0), s.text(8, 1)]).toEqual(['TRUE', '#N/A'])
    expect(s.text(1, 0)).toBe('Apples')

    const x = await xlsxParts(featureBook())
    const wb = await x.text('xl/workbook.xml')
    expect(wb).toContain('<definedName name="Rate">0.5</definedName>')
    expect(wb).toContain('<definedName name="_xlnm.Print_Area" localSheetId="0">Main!$A$1:$C$10</definedName>')
    expect(wb).toContain('<sheet name="Secret" sheetId="4" state="veryHidden" r:id="rId4"/>')
    const xml = await x.text('xl/worksheets/sheet1.xml')
    expect(xml).toContain('<f t="array" ref="D5:E5">B2:C2*10</f>')
  })

  test('formats, panes, sizes, merges, tab color, palette', async () => {
    const { sheet } = await openXls(featureBook())
    const s = await sheet(0)
    // XF 16: bold font, a solid fill of palette color 12 (changed to orange), thin borders
    expect(s.style(0, 0)).toMatchObject({ fontWeight: 700, background: '#ff8800', borderLeft: '1px solid #000000' })
    // XF 17: number format 4 (#,##0.00), centered
    expect(s.text(7, 0)).toBe('Total')
    expect(s.style(7, 0)).toMatchObject({ textAlign: 'center' })
    expect(s.cell(8, 2)?.[2]).toBe(s.cell(7, 0)?.[2])
    // XFEXT colors win over the palette color the XF keeps (the theme: Office 2007's)
    expect(s.style(12, 0)).toMatchObject({ background: applyTint('#4f81bd', 26213 / 32767), color: '#123456' })
    expect(s.head.frozen).toEqual({ rows: 1, cols: 0 })
    expect(s.head.colInfo).toContainEqual({ min: 0, max: 0, w: 140 })
    expect(s.last.merges).toEqual([{ r0: 10, c0: 0, r1: 11, c1: 2 }])
    const xml = await (await xlsxParts(featureBook())).text('xl/worksheets/sheet1.xml')
    expect(xml).toContain('<tabColor rgb="FFFF0000"/>')
    expect(xml).toContain('<row r="1" ht="20" customHeight="1">')
  })

  test("the workbook's own theme colors (THEME record), split over a CONTINUEFRT12 record", async () => {
    const themeXml =
      '<?xml version="1.0"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Custom"><a:themeElements><a:clrScheme name="Custom">' +
      '<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>' +
      '<a:dk2><a:srgbClr val="222222"/></a:dk2><a:lt2><a:srgbClr val="EEEEEE"/></a:lt2><a:accent1><a:srgbClr val="AA2200"/></a:accent1>' +
      '</a:clrScheme></a:themeElements></a:theme>'
    const zip = new Uint8Array(await (await writeZip(new Map([['theme/theme/theme1.xml', themeXml]]))).arrayBuffer())
    const half = zip.length >> 1
    const frt = (rt: number) => cat(u16(rt), u16(0), u32(0), u32(0))
    const extra = [
      rec(0x0896, frt(0x0896), u32(0), zip.subarray(0, half)),
      rec(0x087f, frt(0x087f), zip.subarray(half)),
      rec(0x087d, frt(0x087d), u16(0), u16(16), u16(0), u16(1), cat(u16(4), u16(20), u16(3), u16(0), u32(4), u32(0), u32(0))),
    ]
    const { sheet } = await openXls(xlsFile({ xfs: [{ pattern: 1, fore: 12 }], extra, sheets: [{ name: 'S', records: [label(0, 0, 'x', 16)] }] }))
    expect((await sheet(0)).style(0, 0).background).toBe('#aa2200')
  })

  test('comments and hyperlinks', async () => {
    const { engine, sheet } = await openXls(featureBook())
    await sheet(0)
    const comments = (await engine.call('comments.list', [])) as { author: { name: string }; text: string; at?: { r: number; c: number } }[]
    expect(comments.map((c) => [c.author.name, c.text, c.at?.r, c.at?.c])).toEqual([['Ann Lee', 'Check it', 1, 1]])
    const x = await xlsxParts(featureBook())
    const xml = await x.text('xl/worksheets/sheet1.xml')
    expect(xml).toContain('<hyperlink ref="A2" r:id="rId1" display="Apples" tooltip="Open the page"/>')
    expect(xml).toContain('<hyperlink ref="A3" location="Other!A1" display="Other!A1"/>')
    expect(xml).toContain('<legacyDrawing r:id="rId3"/>')
    const rels = await x.text('xl/worksheets/_rels/sheet1.xml.rels')
    expect(rels).toContain('Target="https://example.com/apples" TargetMode="External"')
    expect(rels).toContain('Target="../comments1.xml"')
    expect(x.has('xl/drawings/vmlDrawing1.vml')).toBe(true)
  })
})

describe('xls: other files', () => {
  test('a password-protected workbook says so', async () => {
    const file = xlsFile({ head: [rec(0x002f, u16(1), u16(1), u16(1), new Array(48).fill(0))], sheets: [{ name: 'S', records: [] }] })
    await expect(convertCfb(file)).rejects.toThrow('Password-protected Excel workbooks are not supported')
  })

  test('Excel 5.0 / 95 (BIFF5): 8-bit strings in the code page, no shared strings', async () => {
    const biff5 = (s: string) => [...s].map((c) => c.charCodeAt(0))
    const cyr = [0xcf, 0xf0, 0xe8, 0xe2, 0xe5, 0xf2] // "Привет" in windows-1251
    const font5 = rec(0x0031, u16(200), u16(0), u16(0x7fff), u16(400), u16(0), u8(0), u8(0), u8(0), u8(0), u8(5), biff5('Arial'))
    const xf5 = (style: boolean) => rec(0x00e0, u16(0), u16(0), u16(style ? 0xfff5 : 1), u8(0x20), u8(0), u32(0), u32(0))
    const sheetBody = cat(
      rec(0x0809, u16(0x0500), u16(0x0010), u16(0), u16(0)),
      rec(0x0204, u16(0), u16(0), u16(15), u16(cyr.length), cyr),
      rec(0x0203, u16(0), u16(1), u16(15), f64(2.5)),
      // =B1*2
      rec(0x0006, u16(1), u16(1), u16(15), f64(5), u16(0), u32(0), u16(8), [0x44, ...u16(0xc000), 1], ptg.int(2), ptg.op(0x05)),
      EOF(),
    )
    const head = cat(
      rec(0x0809, u16(0x0500), u16(0x0005), u16(0), u16(0)),
      rec(0x0042, u16(1251)),
      font5, font5, font5, font5,
      ...Array.from({ length: 15 }, () => xf5(true)), xf5(false),
    )
    const sheetRec = (pos: number) => rec(0x0085, u32(pos), u8(0), u8(0), u8(4), biff5('List'))
    const globalsLen = head.length + sheetRec(0).length + EOF().length
    const stream = Uint8Array.from(cat(head, sheetRec(globalsLen), EOF(), sheetBody))
    const { sheet, info } = await openXls(writeCfb({ Book: stream }))
    expect(info.xlsx!.sheets[0].name).toBe('List')
    const s = await sheet(0)
    expect(s.text(0, 0)).toBe('Привет')
    expect(s.text(1, 1)).toBe('5')
    expect(await s.formula(1, 1)).toBe('=B1*2')
  })

  test('a worksheet stream without workbook globals or compound file (as SheetJS writes)', async () => {
    const stream = Uint8Array.from(cat(rec(0x0809, u16(0), u16(0x1000), u16(0), u16(0)), number(0, 0, 1.5), rec(0x0204, u16(1), u16(0), u16(16), u16(3), [0x61, 0x62, 0x63]), EOF()))
    const zip = await ZipArchive.open(new BlobSource(await biffToXlsx(stream)))
    expect(await zip.text('xl/worksheets/sheet1.xml')).toContain('<c r="A1"><v>1.5</v></c></row><row r="2"><c r="A2" t="s"><v>0</v></c>')
    expect(await zip.text('xl/sharedStrings.xml')).toContain('<si><t>abc</t></si>')
  })

  test('damaged records and truncated streams: converted as far as they go, never hanging', async () => {
    const stream = workbookStream({ sst: ['a', 'b'], sheets: [{ name: 'S', records: [labelSst(0, 0, 0), number(0, 1, 1), labelSst(1, 0, 9)] }] })
    for (const cut of [40, 200, stream.length - 30]) {
      const part = stream.slice(0, cut)
      await convertCfb(writeCfb({ Workbook: part })).catch((e: Error) => expect(e.message).toBeTruthy())
    }
    // records claiming more data than there is, garbage after the sheet
    const junk = Uint8Array.from(cat(stream, rec(0x0006, [1, 2, 3]), [0xfd, 0x00, 0xff, 0xff, 1, 2]))
    const { file } = await convertCfb(writeCfb({ Workbook: junk }))
    expect(file.size).toBeGreaterThan(0)
    // a string index past the shared strings: an empty cell
    const s = await (await openXls(writeCfb({ Workbook: stream }))).sheet(0)
    expect([s.text(0, 0), s.text(0, 1), s.text(1, 0)]).toEqual(['a', '1', ''])
  })

  test('a 65,536-row sheet converts in a few seconds', async () => {
    const rows = 65536
    const parts: Uint8Array[] = []
    let size = 0
    for (let r = 0; r < rows; r++) {
      const cells = [labelSst(r, 0, r % 100), number(r, 1, r + 0.5), mulRk(r, 2, [r, r + 1]), formula(r, 4, cat(ptg.ref(r, 3), ptg.int(2), ptg.op(0x05)), (r + 1) * 2)]
      for (const c of cells) {
        const b = Uint8Array.from(c)
        parts.push(b)
        size += b.length
      }
    }
    const body = new Uint8Array(size)
    let o = 0
    for (const p of parts) {
      body.set(p, o)
      o += p.length
    }
    const file = xlsFile({ sst: Array.from({ length: 100 }, (_, i) => `Item ${i}`), sheets: [{ name: 'Big', records: [body] }] })
    const t = performance.now()
    const { file: xlsx } = await convertCfb(file)
    expect(performance.now() - t).toBeLessThan(10_000)
    const { sheet } = await openXls(new Uint8Array(await new Blob([file]).arrayBuffer()))
    const s = await sheet(0, 5)
    expect([s.text(1, 0), s.text(1, 1), s.text(1, 2), s.text(1, 4)]).toEqual(['Item 1', '1.5', '1', '4'])
    expect(xlsx.size).toBeGreaterThan(100_000)
  }, 60_000)
})
