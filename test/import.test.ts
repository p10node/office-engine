import { describe, expect, test } from 'bun:test'
import { fieldValue, parseCsv } from '../src/core/convert'
import type { ByteSource } from '../src/core/source'
import { Engine } from '../src/core/worker/engine'
import { loadDelimited, type LoadReport } from '../src/core/xlsx/import'
import type { BlockOut, RowOut, SheetProgress } from '../src/core/xlsx/types'
import { SheetData, T_BOOL, T_NUM, T_STR } from '../src/core/xlsx/workbook'

const enc = (s: string) => new TextEncoder().encode(s)
const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms))

/** bytes handed out `chunk` at a time: a file can be cut anywhere (inside a quote, a CRLF, a multi-byte character) */
class Chunked implements ByteSource {
  readonly size: number
  constructor(
    private bytes: Uint8Array,
    private chunk: number,
    readonly type?: string,
    readonly name?: string,
  ) {
    this.size = bytes.length
  }
  async read(offset: number, length: number) {
    return this.bytes.subarray(offset, offset + length)
  }
  stream(offset: number, length: number) {
    const end = Math.min(offset + length, this.bytes.length)
    const { bytes, chunk } = this
    let pos = offset
    return new ReadableStream<Uint8Array>({
      pull(ctrl) {
        if (pos >= end) return ctrl.close()
        const next = Math.min(pos + chunk, end)
        ctrl.enqueue(bytes.slice(pos, next))
        pos = next
      },
    })
  }
}

const quiet: LoadReport = { chunk: async () => {}, reset: () => {} }

type Cell = [number, number, unknown]
/** the model's rows as [col, type, value] triples (undefined: a row without cells) */
function cellsOf(d: SheetData): (Cell[] | undefined)[] {
  const out: (Cell[] | undefined)[] = []
  for (let r = 0; r < d.rows.length; r++) {
    const row = d.rows[r]
    if (!row) {
      out.push(undefined)
      continue
    }
    const cells: Cell[] = []
    for (let k = 0; k < row.length; k += 4) cells.push([row[k] as number, row[k + 1] as number, row[k + 2]])
    out.push(cells)
  }
  return out
}

/** what the whole-text parser gives, typed the same way */
function expected(text: string, sep?: string): (Cell[] | undefined)[] {
  const rows = parseCsv(text, sep).map((row) => {
    const cells: Cell[] = []
    row.forEach((v, c) => {
      if (v === '') return
      const k = fieldValue(v, sep === ';')
      cells.push(k === null ? [c, T_STR, v] : typeof k === 'boolean' ? [c, T_BOOL, k ? 1 : 0] : [c, T_NUM, k])
    })
    return cells.length ? cells : undefined
  })
  while (rows.length && rows[rows.length - 1] === undefined) rows.pop()
  return rows
}

async function load(bytes: Uint8Array, chunk: number, opts: { sep?: string; type?: string; budget?: number; pageRows?: number } = {}, report = quiet) {
  const d = new SheetData()
  await loadDelimited(d, { source: new Chunked(bytes, chunk, opts.type), sep: opts.sep, budget: opts.budget, pageRows: opts.pageRows }, report)
  return d
}

describe('delimited text streamed into a sheet', () => {
  const SAMPLE = 'id,name,note,ok\r\n1,"Pear, green","say ""hi""",TRUE\n2,Apple,"multi\nline",false\r\n\n3,,x,\r4,"a""b"c,,\n"",7,"",008\n5,last line,no break,true'

  test('cut anywhere, the rows are those of the whole text', async () => {
    const bytes = enc(SAMPLE)
    const want = expected(SAMPLE)
    expect(want.length).toBe(8)
    for (const chunk of [1, 2, 3, 5, 7, 11, 64, bytes.length]) {
      const d = await load(bytes, chunk)
      expect(cellsOf(d)).toEqual(want)
      expect(d.maxRow).toBe(7)
      expect(d.maxCol).toBe(3)
      expect(d.structural).toBe(true)
      expect(d.dirtyCols).toBe(true)
    }
  })

  test('values: numbers, booleans, codes and text; a semicolon file reads 12,5 as a number', async () => {
    const d = await load(enc('n;b;code;t;big;sp\n12,5;false;007;1.234,5;1234567890123456; 42 \n'), 4)
    expect(cellsOf(d)[1]).toEqual([
      [0, T_NUM, 12.5],
      [1, T_BOOL, 0],
      [2, T_STR, '007'],
      [3, T_STR, '1.234,5'],
      [4, T_STR, '1234567890123456'],
      [5, T_NUM, 42],
    ])
    const c = await load(enc('a,b,c,d,e\n1e3,-.5,+7,TRUE,yes\n'), 3)
    expect(cellsOf(c)[1]).toEqual([
      [0, T_NUM, 1000],
      [1, T_NUM, -0.5],
      [2, T_NUM, 7],
      [3, T_BOOL, 1],
      [4, T_STR, 'yes'],
    ])
  })

  test('tabs, and a separator given beats the detected one', async () => {
    expect(cellsOf(await load(enc('a\tb\n1\t2\n'), 2))).toEqual(expected('a\tb\n1\t2\n'))
    const d = await load(enc('a,b\tc\n'), 100, { sep: '\t' })
    expect(cellsOf(d)[0]).toEqual([
      [0, T_STR, 'a,b'],
      [1, T_STR, 'c'],
    ])
  })

  test('multi-byte UTF-8 split between chunks', async () => {
    const text = 'é,日本語,😀 ok\nçà,"x\n😀",fin\n'
    const bytes = enc(text)
    for (const chunk of [1, 2, 3, 4, 5, bytes.length]) expect(cellsOf(await load(bytes, chunk))).toEqual(expected(text))
  })

  test('byte order marks: UTF-8, UTF-16 LE and BE', async () => {
    const text = 'name,qty\nPêche,3\n'
    const want = expected(text)
    expect(cellsOf(await load(new Uint8Array([0xef, 0xbb, 0xbf, ...enc(text)]), 5))).toEqual(want)
    const le = new Uint8Array(2 + text.length * 2)
    le[0] = 0xff
    le[1] = 0xfe
    const be = new Uint8Array(le.length)
    be[0] = 0xfe
    be[1] = 0xff
    for (let i = 0; i < text.length; i++) {
      const code = text.charCodeAt(i)
      le[2 + i * 2] = code & 0xff
      le[3 + i * 2] = code >> 8
      be[2 + i * 2] = code >> 8
      be[3 + i * 2] = code & 0xff
    }
    expect(cellsOf(await load(le, 3))).toEqual(want)
    expect(cellsOf(await load(be, 3))).toEqual(want)
  })

  test('not UTF-8: read again as Windows-1252, the rows so far dropped first', async () => {
    const ascii = 'id,name\n1,Ann\n2,Bob\n'
    const bytes = new Uint8Array([...enc(ascii), 0x33, 0x2c, 0x50, 0xe9, 0x63, 0x68, 0x65, 0x0a]) // 3,Pêche (é as 0xe9)
    let resets = 0
    const d = await load(bytes, 4, {}, {
      chunk: async () => {},
      reset: () => {
        resets++
      },
    })
    expect(resets).toBe(1)
    expect(cellsOf(d)).toEqual(expected(ascii + '3,Péche\n'))
    // a declared charset is trusted as it is
    resets = 0
    const e = await load(bytes, 4, { type: 'text/csv; charset=iso-8859-1' }, { chunk: async () => {}, reset: () => resets++ })
    expect(resets).toBe(0)
    expect(cellsOf(e)[3]).toEqual([
      [0, T_NUM, 3],
      [1, T_STR, 'Péche'],
    ])
  })

  test('column widths come from the first rows; later ones do not change them', async () => {
    const lines = ['short,long column here,n']
    for (let i = 0; i < 6000; i++) lines.push(`a,${i === 5500 ? 'x'.repeat(80) : 'bb'},${i}`)
    const d = await load(enc(lines.join('\n')), 1 << 14)
    // 'long column here' = 16 chars → 18 chars wide; the 80-char text at row 5500 came after the header
    expect(d.colInfo).toEqual([{ min: 1, max: 1, w: 18 * 7 }])
    expect(d.maxRow).toBe(6000)
    // a text wider than 60 characters is capped
    const w = await load(enc(`${'y'.repeat(90)},"two\nlines of text here"\n`), 100)
    expect(w.colInfo).toEqual([
      { min: 0, max: 0, w: 60 * 7 },
      { min: 1, max: 1, w: 20 * 7 },
    ])
  })

  test('the view is told as the rows come: header first, then chunks', async () => {
    const lines = ['a,b']
    for (let i = 0; i < 12_000; i++) lines.push(`${i},x${i}`)
    const seen: number[] = []
    const d = new SheetData()
    await loadDelimited(d, { source: new Chunked(enc(lines.join('\n')), 1 << 12) }, {
      chunk: async () => {
        seen.push(d.maxRow + 1)
      },
      reset: () => {},
    })
    // nothing reported before 5000 rows (the widths sample), then growing counts
    expect(seen.length).toBeGreaterThan(2)
    expect(seen[0]).toBeGreaterThanOrEqual(5000)
    for (let i = 1; i < seen.length; i++) expect(seen[i]).toBeGreaterThanOrEqual(seen[i - 1])
    expect(d.maxRow).toBe(12_000)
  })

  test('edges: empty file, blank lines, an unclosed quote, a quote at the very end', async () => {
    expect(cellsOf(await load(new Uint8Array(0), 4))).toEqual([])
    const e = await load(new Uint8Array(0), 4)
    expect(e.maxRow).toBe(-1)
    expect(cellsOf(await load(enc('\n\n'), 1))).toEqual([])
    expect(cellsOf(await load(enc('a\n\n\nb\n'), 1))).toEqual(expected('a\n\n\nb\n'))
    expect(cellsOf(await load(enc('a,"unclosed\nquote'), 3))).toEqual(expected('a,"unclosed\nquote'))
    expect(cellsOf(await load(enc('a,"end"'), 2))).toEqual(expected('a,"end"'))
    expect(cellsOf(await load(enc('a,"end"\r'), 7))).toEqual(expected('a,"end"\r'))
  })
})

async function open(src: Blob | ArrayBuffer | Uint8Array, name?: string) {
  const events: { event: string; data: any }[] = []
  const engine = new Engine((event, data) => events.push({ event, data }))
  const info = await engine.open(src, name ? { name } : {})
  return { engine, info, events }
}

async function shown(engine: Engine, sheet = 0, rows = 50) {
  await engine.call('xlsx.load', [sheet])
  await sleep(50)
  const block: any = await engine.call('xlsx.block', [sheet, 0, rows])
  return block.rows.map((row: any) => row.cells.map((c: any) => c[1]))
}

describe('CSV through the engine', () => {
  test('opens as a workbook whose sheet streams in; saved, it is a plain workbook with the rows', async () => {
    const lines = ['id,city,population']
    for (let i = 1; i <= 7000; i++) lines.push(`${i},City ${i},${i * 1000}`)
    const file = new File([lines.join('\n')], 'cities.csv', { type: 'text/csv' })
    const { engine, info, events } = await open(file)
    expect(info).toMatchObject({ kind: 'xlsx', name: 'cities.csv', imported: 'csv' })
    expect(info.xlsx!.sheets.map((s) => s.name)).toEqual(['cities'])
    const rows = await shown(engine, 0, 3)
    expect(rows).toEqual([
      ['id', 'city', 'population'],
      ['1', 'City 1', '1000'],
      ['2', 'City 2', '2000'],
    ])
    const progress = events.filter((e) => e.event === 'xlsx:progress').map((e) => e.data)
    // 'population' (10 characters) is the one header wider than 9; the rows' texts are shorter
    expect(progress[0].colInfo).toEqual([{ min: 2, max: 2, w: 12 * 7 }])
    expect(progress[0].defaultRowH).toBeDefined()
    expect(progress.at(-1)).toMatchObject({ rows: 7001, cols: 3, done: true })
    const detail: any = await engine.call('xlsx.cell', [0, 7000, 2])
    expect(detail).toMatchObject({ value: '7000000', type: 'number' })

    const saved = (await engine.call('save', [])) as Blob
    const again = await open(saved)
    expect(again.info.imported).toBeUndefined()
    expect(again.info.xlsx!.sheets.map((s) => s.name)).toEqual(['cities'])
    const back = await shown(again.engine, 0, 7001)
    expect(back.length).toBe(7001)
    expect(back[7000]).toEqual(['7000', 'City 7000', '7000000'])
    const layout = again.events.filter((e) => e.event === 'xlsx:progress').map((e) => e.data)[0]
    expect(layout.colInfo).toEqual([{ min: 2, max: 2, w: 12 * 7 }])
  })

  test('saving before the sheet was ever loaded keeps the rows', async () => {
    const { engine } = await open(new File(['a,b\n1,2\n3,4\n'], 'data.csv'))
    const saved = (await engine.call('save', [])) as Blob
    const again = await open(saved)
    expect(await shown(again.engine)).toEqual([
      ['a', 'b'],
      ['1', '2'],
      ['3', '4'],
    ])
  })

  test('edits on top of the streamed rows are saved with them', async () => {
    const { engine } = await open(new File(['a,b\n1,2\n'], 'data.csv'))
    await shown(engine)
    await engine.call('xlsx.set', [0, [{ r: 1, c: 1, input: '=A2*10' }, { r: 2, c: 0, input: 'new' }]])
    expect(await shown(engine)).toEqual([
      ['a', 'b'],
      ['1', '10'],
      ['new'],
    ])
    const again = await open((await engine.call('save', [])) as Blob)
    expect(await shown(again.engine)).toEqual([
      ['a', 'b'],
      ['1', '10'],
      ['new'],
    ])
  })

  test('xlsx.import: a new sheet after the current one with the file’s rows, named after the file', async () => {
    const { engine } = await open(new File(['a,b\n1,2\n'], 'data.csv'))
    await shown(engine)
    const res: any = await engine.call('xlsx.import', [0, new File(['city\tpop\nHanoi\t8.4\n'], 'cities.tsv'), { sep: '\t' }])
    expect(res.sheets.map((s: any) => s.name)).toEqual(['data', 'cities'])
    expect(await shown(engine, 1)).toEqual([
      ['city', 'pop'],
      ['Hanoi', '8.4'],
    ])
    // a name given, a buffer instead of a file, the separator detected
    const res2: any = await engine.call('xlsx.import', [0, enc('x;y\n1,5;2\n').buffer, { name: 'EU data' }])
    expect(res2.sheets.map((s: any) => s.name)).toEqual(['data', 'EU data', 'cities'])
    expect(await shown(engine, 1)).toEqual([
      ['x', 'y'],
      ['1.5', '2'],
    ])
    expect(await engine.call('dirty', [])).toBe(true)
    const again = await open((await engine.call('save', [])) as Blob)
    expect(again.info.xlsx!.sheets.map((s) => s.name)).toEqual(['data', 'EU data', 'cities'])
    expect(await shown(again.engine, 2)).toEqual([
      ['city', 'pop'],
      ['Hanoi', '8.4'],
    ])
    expect(await shown(again.engine, 0)).toEqual([
      ['a', 'b'],
      ['1', '2'],
    ])
  })

  test('an empty .csv is an empty workbook; a TSV opens by its extension', async () => {
    const { engine, info } = await open(new Uint8Array(0), 'empty.csv')
    expect(info.kind).toBe('xlsx')
    expect(await shown(engine)).toEqual([])
    const again = await open((await engine.call('save', [])) as Blob)
    expect(again.info.kind).toBe('xlsx')
    const tsv = await open(enc('a,b\tc\n'), 'd.tsv')
    expect(tsv.info.imported).toBe('tsv')
    expect(await shown(tsv.engine)).toEqual([['a,b', 'c']])
  })
})

// ---------------------------------------------------------------- rows past the budget: read from the file by pages

/** rows with ids, repeated names, amounts and quoted notes; a quoted line break every 50 rows, a blank line every 97 */
function bigText(n: number): string {
  const lines = ['id,name,amount,note']
  for (let i = 1; i <= n; i++) {
    if (i % 97 === 0) lines.push('')
    lines.push(i % 50 === 0 ? `${i},"two\nlines",${(i * 1.5).toFixed(1)},x` : `${i},Name ${i % 37},${(i * 1.5).toFixed(1)},"n, ${i}"`)
  }
  return lines.join('\n') + '\n'
}

function cellsOfRow(row: unknown[] | undefined): Cell[] | undefined {
  if (!row) return undefined
  const cells: Cell[] = []
  for (let k = 0; k < row.length; k += 4) cells.push([row[k] as number, row[k + 1] as number, row[k + 2]])
  return cells
}

/** every row of a paged sheet, page by page, as `cellsOf` gives them */
async function pagedCells(d: SheetData): Promise<(Cell[] | undefined)[]> {
  const pages = d.pages!
  const out: (Cell[] | undefined)[] = []
  for (let p = 0; p * pages.size <= d.maxRow; p++) {
    const rows = await pages.peek(p)
    for (let k = 0; k < pages.size && p * pages.size + k <= d.maxRow; k++) out.push(cellsOfRow(rows[k]))
  }
  while (out.length && out[out.length - 1] === undefined) out.pop()
  return out
}

describe('rows past the budget stay in the file, read by pages', () => {
  const TEXT = bigText(3000)
  const want = expected(TEXT)

  test('the first pages are in memory, the rest counted; every page reads back right', async () => {
    const d = await load(enc(TEXT), 777, { budget: 1000, pageRows: 50 })
    const pages = d.pages!
    expect(pages).toBeDefined()
    expect(pages.complete).toBe(true)
    expect(d.maxRow).toBe(want.length - 1)
    expect(d.maxCol).toBe(3)
    // the budget was passed while page 5 or 6 streamed in: those before it stayed
    expect(pages.count).toBeGreaterThanOrEqual(5)
    expect(pages.count).toBeLessThanOrEqual(7)
    expect(pages.has(10)).toBe(true)
    expect(pages.has(2000)).toBe(false)
    expect(d.rows[2000]).toBeUndefined()
    expect(await pagedCells(d)).toEqual(want)
    // the same file under the budget: every row in memory, no pages
    const whole = await load(enc(TEXT), 777)
    expect(whole.pages).toBeUndefined()
    expect(cellsOf(whole)).toEqual(want)
  })

  test('pages come as rows are needed and the least recently used leave; pinned (edited) pages stay; too many rows at once is refused', async () => {
    const d = await load(enc(TEXT), 1000, { budget: 1000, pageRows: 50 })
    const pages = d.pages!
    await pages.ensure(2000, 2100)
    expect(pages.has(2000) && pages.has(2050) && pages.has(2100)).toBe(true)
    expect(cellsOfRow(d.rows[2050])).toEqual(want[2050])
    expect(pages.cells).toBeLessThanOrEqual(1000)
    expect(pages.has(10)).toBe(false)
    expect(d.rows[10]).toBeUndefined()
    pages.pin(2050)
    await pages.ensure(0, 120)
    expect(pages.has(10)).toBe(true)
    expect(pages.has(2050)).toBe(true)
    expect(pages.has(2000)).toBe(false)
    expect(pages.cells).toBeLessThanOrEqual(1000)
    await expect(pages.ensure(0, 2000)).rejects.toThrow(/More rows at once/)
    // rows past the file's are never paged
    expect(pages.has(5000)).toBe(true)
  })

  test('UTF-16 (units split between chunks) and Windows-1252 (read again) files page too', async () => {
    const u16 = new Uint8Array(2 + TEXT.length * 2)
    u16[0] = 0xff
    u16[1] = 0xfe
    for (let i = 0; i < TEXT.length; i++) {
      const c = TEXT.charCodeAt(i)
      u16[2 + 2 * i] = c & 0xff
      u16[3 + 2 * i] = c >> 8
    }
    const d = await load(u16, 333, { budget: 1000, pageRows: 50 })
    expect(d.pages).toBeDefined()
    expect(await pagedCells(d)).toEqual(want)

    const latin = TEXT.replace('Name 5,', 'café,').replace('Name 7,', 'naïve,')
    const bytes = Uint8Array.from(latin, (ch) => ch.charCodeAt(0))
    let resets = 0
    const e = await load(bytes, 500, { budget: 1000, pageRows: 50 }, { chunk: async () => {}, reset: () => resets++ })
    expect(resets).toBe(1)
    expect(e.pages).toBeDefined()
    expect(await pagedCells(e)).toEqual(expected(latin))
  })

  test('any text cut into pages reads back as the whole text does', async () => {
    let seed = 7
    const rnd = (n: number) => {
      seed = (seed * 1103515245 + 12345) & 0x7fffffff
      return seed % n
    }
    const alphabet = ['a', 'b', '"', ',', ',', '\n', '\n', '\r', ' ', '1']
    for (let i = 0; i < 300; i++) {
      let text = ''
      const n = 1 + rnd(60)
      for (let k = 0; k < n; k++) text += alphabet[rnd(alphabet.length)]
      const d = await load(enc(text), 1 + rnd(9), { sep: ',', budget: 1, pageRows: 1 + rnd(3) })
      const got = d.pages ? await pagedCells(d) : cellsOf(d)
      const exp = expected(text, ',')
      expect([text, got]).toEqual([text, exp])
      expect([text, d.maxRow]).toEqual([text, exp.length - 1])
    }
  })

  test('through the engine: blocks, find, the cell, an edit far down, formulas, sort, export and save', async () => {
    const eng = new Engine(() => {})
    await eng.open(new Blob([TEXT], { type: 'text/csv' }), { name: 'big.csv', residentCells: 1000, pageRows: 50 })
    let last: SheetProgress | undefined
    await (eng as any).xlsx.loadSheet(0, (p: SheetProgress) => {
      last = p
    })
    expect(last!.rows).toBe(want.length)
    expect(last!.done).toBe(true)
    const texts = (rows: RowOut[]) => rows.map((r) => [r.r, r.cells.map((c) => [c[0], c[1]])])
    const wantTexts = (r0: number, r1: number) => {
      const out: [number, (string | number)[][]][] = []
      for (let r = r0; r < r1; r++) if (want[r]) out.push([r, want[r]!.map((c) => [c[0], String(c[2])])])
      return out
    }
    const far = (await eng.call('xlsx.block', [0, 2900, 2960])) as BlockOut
    expect(far.complete).toBe(true)
    expect(texts(far.rows)).toEqual(wantTexts(2900, 2960))
    expect(await eng.call('xlsx.cell', [0, want.length - 1, 1])).toMatchObject({ value: 'two\nlines' })
    expect(await eng.call('xlsx.find', [0, 'n, 2999'])).toEqual([{ r: want.length - 2, c: 3, text: 'n, 2999' }])
    // an edit far down, and a formula near the top reading a far row: both rows' pages come in and stay
    const r = 2500
    const id = Number(want[r]![0][2])
    await eng.call('xlsx.set', [0, [{ r, c: 1, input: 'edited' }, { r: 10, c: 5, input: `=A${r + 1}*2` }]])
    expect(await eng.call('xlsx.cell', [0, 10, 5])).toMatchObject({ value: String(id * 2), formula: `=A${r + 1}*2` })
    for (const r0 of [0, 1000, 1500, 2000, 2900, 500]) await eng.call('xlsx.block', [0, r0, r0 + 128])
    const edited = (await eng.call('xlsx.block', [0, r, r + 1])) as BlockOut
    expect(edited.rows[0].cells[1]).toEqual([1, 'edited', 0, 0])
    // a formula over every row of the sheet cannot be computed from memory
    await expect(eng.call('xlsx.set', [0, [{ r: 0, c: 6, input: '=SUM(A:A)' }]])).rejects.toThrow(/More rows at once/)
    // rows cannot move, a small range sorts
    await expect(eng.call('xlsx.sheetOp', [0, { op: 'insertRows', at: 5, n: 1 }])).rejects.toThrow(/cannot be inserted/)
    await eng.call('xlsx.sheetOp', [0, { op: 'sort', range: { r0: 2800, c0: 0, r1: 2810, c1: 3 }, col: 0, desc: true }])
    const sorted = (await eng.call('xlsx.block', [0, 2800, 2801])) as BlockOut
    expect(sorted.rows[0].cells[0][1]).toBe(String(Number(want[2810]![0][2])))
    // the whole sheet as CSV, page by page
    const csv = await ((await eng.call('xlsx.text', [0, ','])) as Blob).text()
    const lines = csv.split('\r\n')
    // a paged sheet's whole used range (the formula made it 6 columns wide), a byte order mark first (text() drops it)
    expect([...new Uint8Array(await ((await eng.call('xlsx.text', [0, ','])) as Blob).arrayBuffer()).slice(0, 3)]).toEqual([0xef, 0xbb, 0xbf])
    expect(lines[0]).toBe('id,name,amount,note,,')
    expect(lines[r]).toBe(`${id},edited,${(id * 1.5).toFixed(1)},"n, ${id}",,`)
    expect(lines.length).toBe(want.length + 1)
    // saved as a workbook: every row, with the edits
    const saved = (await eng.call('save', [])) as Blob
    const again = new Engine(() => {})
    await again.open(saved, { name: 'big.xlsx' })
    let l2: SheetProgress | undefined
    await (again as any).xlsx.loadSheet(0, (p: SheetProgress) => {
      l2 = p
    })
    expect(l2!.rows).toBe(want.length)
    const b2 = (await again.call('xlsx.block', [0, r, r + 1])) as BlockOut
    expect(b2.rows[0].cells.map((c) => c[1])).toEqual([String(id), 'edited', (id * 1.5).toFixed(1), `n, ${id}`])
    expect(texts(((await again.call('xlsx.block', [0, 2900, 2960])) as BlockOut).rows)).toEqual(wantTexts(2900, 2960))
  })

  test('Excel’s row limit: a sheet with more rows is not saved as a workbook, its CSV still is', async () => {
    const eng = new Engine(() => {})
    await eng.open(new Blob(['a,b\n1,2\n'], { type: 'text/csv' }), { name: 's.csv' })
    await (eng as any).xlsx.loadSheet(0, () => {})
    await eng.call('xlsx.set', [0, [{ r: 1_048_576, c: 0, input: 'x' }]])
    await expect(eng.call('save', [])).rejects.toThrow(/1,048,577 rows: an Excel workbook holds at most 1,048,576/)
    const csv = await ((await eng.call('xlsx.text', [0, ','])) as Blob).text()
    expect(csv.endsWith('x,\r\n')).toBe(true)
  })
})
