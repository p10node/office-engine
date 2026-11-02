import { describe, expect, test } from 'bun:test'
import { existsSync } from 'node:fs'
import { Engine } from '../src/core/worker/engine'
import type { Block, DocxMeta, Paragraph } from '../src/core/docx/types'
import type { SheetProgress } from '../src/core/xlsx/types'
import { formatValue } from '../src/core/xlsx/numfmt'
import { SaxParser, TreeBuilder } from '../src/core/xml'

const S = (n: string) => `${import.meta.dir}/../playground/public/samples/${n}`
const file = (n: string) => Bun.file(S(n)) as unknown as Blob

function collect() {
  const events: { event: string; data: any }[] = []
  const engine = new Engine((event, data) => events.push({ event, data }))
  return { engine, events }
}

describe('xml', () => {
  test('chunked streaming + skip + alternate content', () => {
    const xml = `<?xml version="1.0"?><w:document xmlns:w="x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p a="1&amp;2"><w:t xml:space="preserve"> a &lt; b </w:t></w:p><mc:AlternateContent><mc:Choice><w:p><w:t>choice</w:t></w:p></mc:Choice><mc:Fallback><w:p><w:t>fallback</w:t></w:p></mc:Fallback></mc:AlternateContent><w:p r:id="rId1"/></w:body></w:document>`
    const out: any[] = []
    const tb = new TreeBuilder({ emitDepth: 2, onElement: (e) => out.push(e) })
    const p = new SaxParser(tb)
    for (let i = 0; i < xml.length; i += 7) p.write(xml.slice(i, i + 7))
    p.end()
    expect(out.length).toBe(3)
    expect(out[0].a.a).toBe('1&2')
    expect(out[0].c[0].t).toBe(' a < b ')
    expect(out[1].c[0].t).toBe('choice')
    expect(out[2].a['r:id']).toBe('rId1')
  })
})

describe('numfmt', () => {
  const f = (v: number, fmt: string) => formatValue(v, fmt).text
  test('numbers', () => {
    expect(f(1234567.891, 'General')).toBe('1234567.891')
    expect(f(1234567.891, '#,##0.00')).toBe('1,234,567.89')
    expect(f(-42.5, '#,##0.00;[Red](#,##0.00)')).toBe('(42.50)')
    expect(formatValue(-42.5, '#,##0.00;[Red](#,##0.00)').color).toBe('#ff0000')
    expect(f(0.256, '0.0%')).toBe('25.6%')
    expect(f(12.5, '"$"#,##0.00')).toBe('$12.50')
    expect(f(1500, '[$€-407] #,##0')).toBe('€ 1,500')
    expect(f(0.000123, '0.00E+00')).toBe('1.23E-04')
    expect(f(5551234567, '(000) 000-0000')).toBe('(555) 123-4567')
    expect(f(1.5, '# ?/?')).toBe('1 1/2')
    expect(f(-3, '0')).toBe('-3')
  })
  test('dates', () => {
    expect(f(45291.75, 'yyyy-mm-dd hh:mm')).toBe('2023-12-31 18:00')
    expect(f(45291.75, 'dddd, mmmm d, yyyy h:mm AM/PM')).toBe('Sunday, December 31, 2023 6:00 PM')
    expect(f(45291, 'm/d/yyyy')).toBe('12/31/2023')
    expect(f(1.5, '[h]:mm:ss')).toBe('36:00:00')
  })
})

describe('docx', () => {
  test('sample.docx streams blocks with styles, lists, tables, images, header/footer', async () => {
    const { engine, events } = collect()
    const info = await engine.open(file('sample.docx'), { name: 'sample.docx' })
    expect(info.kind).toBe('docx')
    await engine.call('docx.start', [])
    const blocks: Block[] = events.filter((e) => e.event === 'docx:blocks').flatMap((e) => e.data)
    const meta: DocxMeta = events.filter((e) => e.event === 'docx:meta').at(-1)!.data
    expect(meta.final).toBe(true)
    expect(meta.page.w).toBeGreaterThan(500)
    expect(meta.footer?.length).toBeGreaterThan(0)
    const footerRuns = (meta.footer![0] as Paragraph).runs
    expect(footerRuns.some((r) => r.t === 'field' && r.kind === 'PAGE')).toBe(true)
    const paras = blocks.filter((b): b is Paragraph => b.t === 'p')
    const title = paras.find((p) => p.runs.some((r) => r.t === 'text' && r.text === 'Quarterly Report'))!
    expect((title.runs[0] as any).s.sz).toBeGreaterThan(20)
    const bullets = paras.filter((p) => p.marker?.text === '•')
    expect(bullets.length).toBe(3)
    const numbered = paras.filter((p) => /^\d\.$/.test(p.marker?.text ?? ''))
    expect(numbered.map((p) => p.marker!.text)).toEqual(['1.', '2.', '3.'])
    const link = paras.flatMap((p) => p.runs).find((r: any) => r.href)
    expect((link as any).href).toBe('https://example.com')
    expect(paras.flatMap((p) => p.runs).some((r) => r.t === 'img')).toBe(true)
    const table = blocks.find((b) => b.t === 'tbl') as any
    expect(table.rows.length).toBe(5)
    expect(table.rows[4].cells[0].colSpan).toBe(3)
    expect(table.rows[0].cells[0].shd).toBe('#d9e2f3')
    expect(table.rows[1].cells[0].borders.t).toContain('solid')
    expect(paras.some((p) => p.pb)).toBe(true)
  })

  test('large docx streams progressively', async () => {
    if (!existsSync(S('large-doc.docx'))) return
    const { engine, events } = collect()
    await engine.open(file('large-doc.docx'), {})
    const t0 = performance.now()
    await engine.call('docx.start', [])
    const batches = events.filter((e) => e.event === 'docx:blocks')
    const total = batches.reduce((n, b) => n + b.data.length, 0)
    console.log(`large docx: ${total} blocks in ${batches.length} batches, ${(performance.now() - t0).toFixed(0)} ms`)
    expect(total).toBeGreaterThan(20000)
  })
})

describe('xlsx', () => {
  test('sample.xlsx formats & merges', async () => {
    const { engine, events } = collect()
    const info = await engine.open(file('sample.xlsx'), {})
    expect(info.kind).toBe('xlsx')
    expect(info.xlsx!.sheets.map((s) => s.name)).toEqual(['Sales', 'Notes', 'Data'])
    await engine.call('xlsx.load', [0])
    await new Promise((r) => setTimeout(r, 300))
    const prog: SheetProgress[] = events.filter((e) => e.event === 'xlsx:progress').map((e) => e.data)
    const last = prog.at(-1)!
    expect(last.done).toBe(true)
    expect(last.merges).toEqual([{ r0: 0, c0: 0, r1: 0, c1: 5 }])
    expect(prog.find((p) => p.frozen)!.frozen).toEqual({ rows: 2, cols: 1 })
    const block: any = await engine.call('xlsx.block', [0, 0, 10])
    const cell = (r: number, c: number) => block.rows.find((x: any) => x.r === r)?.cells.find((x: any) => x[0] === c)
    expect(cell(0, 0)[1]).toBe('Sales by region — FY2026')
    expect(cell(1, 0)[1]).toBe('Region')
    expect(cell(2, 1)[1]).toMatch(/^2026-\d\d-\d\d$/)
    expect(cell(2, 3)[1]).toMatch(/^\$\d+\.\d\d$/)
    expect(cell(2, 5)[1]).toMatch(/%$/)
    expect(cell(2, 7)[1]).toBe('1234567.891')
    expect(cell(3, 7)[1]).toBe('(42.50)')
    const detail: any = await engine.call('xlsx.cell', [0, 2, 4])
    expect(detail.formula).toBe('=C3*D3')
  })

  test('large xlsx loads', async () => {
    if (!existsSync(S('large-sheet.xlsx'))) return
    const { engine, events } = collect()
    await engine.open(file('large-sheet.xlsx'), {})
    const t0 = performance.now()
    await engine.call('xlsx.load', [0])
    while (!events.some((e) => e.event === 'xlsx:progress' && e.data.done)) await new Promise((r) => setTimeout(r, 50))
    const n = events.filter((e) => e.event === 'xlsx:progress').length
    const last = events.filter((e) => e.event === 'xlsx:progress').at(-1)!.data
    console.log(`large xlsx: ${last.rows} rows x ${last.cols} cols, ${n} progress events, ${(performance.now() - t0).toFixed(0)} ms`)
    expect(last.rows).toBe(300001)
    const block: any = await engine.call('xlsx.block', [0, 150000, 150010])
    expect(block.rows.length).toBe(10)
  }, 120000)
})

describe('pptx', () => {
  test('sample.pptx slides', async () => {
    const { engine } = collect()
    const info = await engine.open(file('sample.pptx'), {})
    expect(info.kind).toBe('pptx')
    expect(info.pptx!.count).toBe(5)
    expect(Math.round(info.pptx!.w)).toBe(1280)
    const s0: any = await engine.call('pptx.slide', [0])
    const texts = s0.els.filter((e: any) => e.text).flatMap((e: any) => e.text.paras.flatMap((p: any) => p.runs.map((r: any) => r.text)))
    expect(texts).toContain('office-engine')
    const title = s0.els.find((e: any) => e.text?.paras[0].runs[0].text === 'office-engine')
    expect(title.text.paras[0].runs[0].sz).toBeGreaterThan(30)
    const s1: any = await engine.call('pptx.slide', [1])
    const body = s1.els.find((e: any) => e.text?.paras.length >= 5)
    expect(body.text.paras[0].bullet?.text).toBeTruthy()
    expect(body.text.paras[1].marL).toBeGreaterThan(body.text.paras[0].marL)
    const s2: any = await engine.call('pptx.slide', [2])
    expect(s2.els.filter((e: any) => e.t === 'shape' && e.fill?.t === 'solid').length).toBeGreaterThanOrEqual(3)
    expect(s2.els.some((e: any) => e.t === 'shape' && e.geom === 'ellipse')).toBe(true)
    expect(s2.els.some((e: any) => e.t === 'pic' && e.target)).toBe(true)
    const s3: any = await engine.call('pptx.slide', [3])
    const tbl = s3.els.find((e: any) => e.t === 'table')
    expect(tbl.rows.length).toBe(5)
    expect(tbl.rows[0].cells[0].fill?.t).toBe('solid')
    const s4: any = await engine.call('pptx.slide', [4])
    expect(s4.els.some((e: any) => e.t === 'placeholder' && e.label === 'Chart')).toBe(true)
  })
})
