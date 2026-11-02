import { describe, expect, test } from 'bun:test'
import { blankPptx, detectSeparator, markdownToHtml, parasToDocx, parseCsv, rowsToXlsx, textToParas } from '../src/core/convert'
import type { Block, Paragraph } from '../src/core/docx/types'
import { Engine } from '../src/core/worker/engine'

const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms))

async function openDocx(src: Blob) {
  const events: { event: string; data: any }[] = []
  const engine = new Engine((event, data) => events.push({ event, data }))
  const info = await engine.open(src, {})
  await engine.call('docx.start', [])
  const blocks: Block[] = events.filter((e) => e.event === 'docx:blocks').flatMap((e) => e.data)
  return { engine, info, paras: blocks.filter((b): b is Paragraph => b.t === 'p') }
}
const textOf = (p: Paragraph) => p.runs.map((r) => (r.t === 'text' ? r.text : '')).join('')

describe('csv', () => {
  test('quotes, escaped quotes, newlines in fields, CRLF', () => {
    expect(parseCsv('a,b\r\n"x, y","say ""hi"""\n"multi\nline",3\n')).toEqual([
      ['a', 'b'],
      ['x, y', 'say "hi"'],
      ['multi\nline', '3'],
    ])
  })

  test('separator detection', () => {
    expect(detectSeparator('a;b;c\n1;2;3')).toBe(';')
    expect(detectSeparator('a\tb\n1\t2')).toBe('\t')
    expect(detectSeparator('a,b\n1,2')).toBe(',')
    expect(parseCsv('a\tb\n1\t2')).toEqual([
      ['a', 'b'],
      ['1', '2'],
    ])
  })

  test('workbook keeps numbers, booleans and codes', async () => {
    const blob = await rowsToXlsx(parseCsv('name,qty,code,ok\nApple,12.5,007,TRUE\n'), 'Fruits')
    const events: { event: string; data: any }[] = []
    const engine = new Engine((event, data) => events.push({ event, data }))
    const info = await engine.open(blob, {})
    expect(info.kind).toBe('xlsx')
    expect(info.xlsx!.sheets.map((s) => s.name)).toEqual(['Fruits'])
    await engine.call('xlsx.load', [0])
    await sleep(50)
    const block: any = await engine.call('xlsx.block', [0, 0, 10])
    const cell = (r: number, c: number) => block.rows.find((x: any) => x.r === r)?.cells.find((x: any) => x[0] === c)
    expect(cell(0, 0)[1]).toBe('name')
    expect(cell(1, 1)[1]).toBe('12.5')
    expect(cell(1, 1)[3]).toBe(1) // number
    expect(cell(1, 2)[1]).toBe('007')
    expect(cell(1, 2)[3]).toBe(0) // text
    expect(cell(1, 3)[1]).toBe('TRUE')
    const detail: any = await engine.call('xlsx.cell', [0, 1, 1])
    expect(detail.type).toBe('number')
  })


  test('semicolon files: decimal comma', async () => {
    const blob = await rowsToXlsx(parseCsv('a;b\n12,5;1.234,5\n', ';'), 'S', true)
    const engine = new Engine(() => {})
    await engine.open(blob, {})
    await engine.call('xlsx.load', [0])
    await sleep(50)
    const block: any = await engine.call('xlsx.block', [0, 0, 10])
    const row = block.rows.find((x: any) => x.r === 1).cells
    expect(row[0][1]).toBe('12.5')
    expect(row[0][3]).toBe(1)
    expect(row[1][3]).toBe(0) // thousands separators stay text
  })
})

describe('docx from text', () => {
  test('plain text: one paragraph per line', async () => {
    const { info, paras } = await openDocx(await parasToDocx(textToParas('Hello\n\nWorld\twide')))
    expect(info.kind).toBe('docx')
    expect(paras.map(textOf)).toEqual(['Hello', '', 'World\twide'])
  })

  test('headings, lists, formatting and links', async () => {
    const blob = await parasToDocx([
      { runs: [{ text: 'Title' }], heading: 1 },
      { runs: [{ text: 'bold', set: { b: true } }, { text: ' and ' }, { text: 'link', href: 'https://example.com', set: { u: 'single', color: '#0563c1' } }] },
      { runs: [{ text: 'one' }], list: 'number' },
      { runs: [{ text: 'two' }], list: 'number' },
      { runs: [{ text: 'dot' }], list: 'bullet' },
      { runs: [{ text: 'again' }], list: 'number' },
    ])
    const { paras } = await openDocx(blob)
    expect(paras[0].ps.outline).toBe(0)
    const runs = paras[1].runs.filter((r) => r.t === 'text') as Extract<Paragraph['runs'][number], { t: 'text' }>[]
    expect(runs[0].s.b).toBe(true)
    expect(runs.find((r) => r.text === 'link')?.href).toBe('https://example.com')
    expect(paras.slice(2).map((p) => p.marker?.text)).toEqual(['1.', '2.', '•', '1.'])
  })

  test('save after open keeps content (valid package for the writer)', async () => {
    const { engine } = await openDocx(await parasToDocx([{ runs: [{ text: 'x' }], heading: 2 }]))
    const out = (await engine.call('save', [])) as Blob
    const again = await openDocx(out)
    expect(again.paras.map(textOf)).toEqual(['x'])
  })
})

describe('markdown', () => {
  test('block structure', () => {
    const html = markdownToHtml('# Title\n\nSome **bold** and *it* `code` [a](https://x.y)\nnext line\n\n- a\n- b\n\n1. one\n2. two\n\n> quote\n\n```\nx < y\n```\n\n| h1 | h2 |\n|---|---|\n| 1 | 2 |')
    expect(html).toContain('<h1>Title</h1>')
    expect(html).toContain('<p>Some <b>bold</b> and <i>it</i> <code>code</code> <a href="https://x.y">a</a> next line</p>')
    expect(html).toContain('<ul>\n<li>a</li>\n<li>b</li>\n</ul>')
    expect(html).toContain('<ol>\n<li>one</li>\n<li>two</li>\n</ol>')
    expect(html).toContain('<blockquote><i>quote</i></blockquote>')
    expect(html).toContain('<pre><code>x &lt; y</code></pre>')
    expect(html).toContain('<td>1</td><td>2</td>')
  })

  test('no script links, escaped html', () => {
    const html = markdownToHtml('[x](javascript:alert(1)) <b>raw</b>')
    expect(html).not.toContain('javascript:')
    expect(html).toContain('&lt;b&gt;raw&lt;/b&gt;')
  })
})

describe('blank presentation', () => {
  test('opens, adds slides from its layouts, edits text, saves', async () => {
    const engine = new Engine(() => {})
    const info = await engine.open(await blankPptx(), {})
    expect(info.kind).toBe('pptx')
    expect(info.pptx!.count).toBe(1)
    expect(info.pptx!.w).toBe(1280)
    const s0: any = await engine.call('pptx.slide', [0])
    const title = s0.els.find((e: any) => e.t === 'shape' && e.text)
    expect(title.box.w).toBeGreaterThan(0)
    const meta: any = await engine.call('pptx.slides', [{ op: 'add', after: 0, layout: 'content' }])
    expect(meta.count).toBe(2)
    const s1: any = await engine.call('pptx.slide', [1])
    expect(s1.els.filter((e: any) => e.t === 'shape').length).toBe(2)
    await engine.call('pptx.setText', [0, title.sid, [{ runs: [{ text: 'Hello' }] }]])
    const out = (await engine.call('save', [])) as Blob
    const again = new Engine(() => {})
    const info2 = await again.open(out, {})
    expect(info2.pptx!.count).toBe(2)
    const t: any = await again.call('pptx.slide', [0])
    expect(JSON.stringify(t.els)).toContain('Hello')
  })
})
