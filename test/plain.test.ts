import { describe, expect, test } from 'bun:test'
import type { Block, Paragraph } from '../src/core/docx/types'
import { htmlToParas, parseHtml } from '../src/core/html'
import { decodeText, plainFormat } from '../src/core/plain'
import { Engine } from '../src/core/worker/engine'

const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms))
const enc = (s: string) => new TextEncoder().encode(s)
const textOf = (p: { runs: { t: string; text?: string }[] }) => p.runs.map((r) => (r.t === 'text' ? r.text : '')).join('')

async function open(src: Blob | ArrayBuffer | Uint8Array, name?: string) {
  const events: { event: string; data: any }[] = []
  const engine = new Engine((event, data) => events.push({ event, data }))
  const info = await engine.open(src, name ? { name } : {})
  return { engine, info, events }
}

async function cells(engine: Engine) {
  await engine.call('xlsx.load', [0])
  await sleep(50)
  const block: any = await engine.call('xlsx.block', [0, 0, 50])
  return block.rows.map((row: any) => row.cells.map((c: any) => c[1]))
}

async function paras(engine: Engine, events: { event: string; data: any }[]) {
  await engine.call('docx.start', [])
  const blocks: Block[] = events.filter((e) => e.event === 'docx:blocks').flatMap((e) => e.data)
  return blocks.filter((b): b is Paragraph => b.t === 'p')
}

describe('plain sources open as Office files', () => {
  test('CSV file: a workbook, sheet named after the file, editable and saved as xlsx', async () => {
    const file = new File(['name,qty\nApple,12\n"Pear, green",3\n'], 'fruits.csv', { type: 'text/csv' })
    const { engine, info } = await open(file)
    expect(info).toMatchObject({ kind: 'xlsx', name: 'fruits.csv', imported: 'csv', size: file.size })
    expect(info.from).toBeUndefined()
    expect(info.xlsx!.sheets.map((s) => s.name)).toEqual(['fruits'])
    expect(await cells(engine)).toEqual([
      ['name', 'qty'],
      ['Apple', '12'],
      ['Pear, green', '3'],
    ])
    const saved = (await engine.call('save', [])) as Blob
    const again = await open(saved)
    expect(again.info.kind).toBe('xlsx')
    expect(again.info.imported).toBeUndefined()
  })

  test('TSV, and a semicolon CSV from a comma-decimal locale', async () => {
    const tsv = await open(enc('a\tb\n1,5\tx\n'), 'data.tsv')
    expect(tsv.info.imported).toBe('tsv')
    expect(await cells(tsv.engine)).toEqual([
      ['a', 'b'],
      ['1,5', 'x'],
    ])
    const semi = await open(enc('a;b\n12,5;x\n'), 'eu.csv')
    expect(await cells(semi.engine)).toEqual([
      ['a', 'b'],
      ['12.5', 'x'],
    ])
  })

  test('text, Markdown and HTML: a document', async () => {
    const txt = await open(enc('line one\nline two'), 'notes.txt')
    expect(txt.info).toMatchObject({ kind: 'docx', imported: 'text' })
    expect((await paras(txt.engine, txt.events)).map(textOf)).toEqual(['line one', 'line two'])

    const md = await open(enc('# Title\n\nSome **bold** [link](https://example.com)\n\n- one\n- two\n'), 'readme.md')
    expect(md.info.imported).toBe('markdown')
    const mp = await paras(md.engine, md.events)
    expect(mp.map(textOf)).toEqual(['Title', 'Some bold link', 'one', 'two'])
    expect(mp[0].ps.outline).toBe(0)
    expect(mp[2].marker?.text).toBe('•')
    const bold = mp[1].runs.find((r) => r.t === 'text' && r.text === 'bold')
    expect(bold?.t === 'text' && bold.s.b).toBe(true)

    // in the worker there is no DOMParser: the light parser reads the file
    const html = await open(enc('<!doctype html><title>T</title><h2>Web &amp; more</h2><p>Para <i>it</i><p>next<ol><li>a<li>b</ol>'), 'page.html')
    expect(html.info.imported).toBe('html')
    expect((await paras(html.engine, html.events)).map(textOf)).toEqual(['Web & more', 'Para it', 'next', 'a', 'b'])
  })

  test('no name: the type, then the content, tells the format', async () => {
    expect((await open(new Blob(['x,y\n1,2\n'], { type: 'text/csv' }))).info.imported).toBe('csv')
    expect((await open(new Blob(['# Not a heading'], { type: 'text/plain' }))).info.imported).toBe('text')
    expect((await open(enc('id,city,pop\n1,Hanoi,8.4\n2,Hue,0.65\n'))).info.imported).toBe('csv')
    expect((await open(enc('a\tb\n1\t2\n'))).info.imported).toBe('tsv')
    expect((await open(enc('<html><body><p>Hi</p></body></html>'))).info.imported).toBe('html')
    expect((await open(enc('Dear team,\nthe meeting moved, again.\nThanks'))).info.imported).toBe('text')
  })

  test('an Office name is not read as text; binary and empty files are not text', async () => {
    // e.g. a missing file served as the site's index.html
    await expect(open(enc('<!doctype html><p>Not found</p>'), 'report.docx')).rejects.toThrow(/Not a ZIP/)
    await expect(open(new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0]), 'chart.csv')).rejects.toThrow('“chart.csv” is not a text file')
    await expect(open(new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]))).rejects.toThrow(/Not a ZIP/)
    await expect(open(enc('%PDF-1.7\n%âãÏÓ\n1 0 obj\n<< /Type /Catalog >>\nendobj\n'), 'scan')).rejects.toThrow(/Not a ZIP/)
    await expect(open(new Uint8Array(0))).rejects.toThrow(/Not a ZIP/)
    // an empty .csv is an empty workbook
    expect((await open(new Uint8Array(0), 'empty.csv')).info.kind).toBe('xlsx')
  })

  test('a package named .csv still opens as the package', async () => {
    const { info: xlsx } = await open(new File(['a,b\n'], 'x.csv'))
    const saved = (await (await open(new File(['a,b\n'], 'x.csv'))).engine.call('save', [])) as Blob
    expect(xlsx.kind).toBe('xlsx')
    const { info } = await open(saved, 'mislabelled.csv')
    expect(info.kind).toBe('xlsx')
    expect(info.imported).toBeUndefined()
  })
})

describe('plain format detection', () => {
  const fmt = (name: string | undefined, type: string | undefined, s: string, whole = true) => plainFormat(name, type, enc(s), whole)

  test('extension first, then MIME type, then content', () => {
    expect(fmt('a.CSV', undefined, 'x')).toBe('csv')
    expect(fmt('a.tab', undefined, 'x')).toBe('tsv')
    expect(fmt('a.log', undefined, 'x')).toBe('text')
    expect(fmt('a.markdown', undefined, 'x')).toBe('markdown')
    expect(fmt('a.htm', undefined, 'x')).toBe('html')
    expect(fmt('a.txt', 'text/html', '<p>x</p>')).toBe('text')
    expect(fmt('export', 'text/csv; charset=utf-8', 'x')).toBe('csv')
    expect(fmt('export', 'application/octet-stream', 'a,b\n1,2')).toBe('csv')
    expect(fmt('a.xlsx', 'text/csv', 'a,b')).toBeNull()
    expect(fmt('a.key', undefined, 'x')).toBeNull()
  })

  test('a cut-off last line does not decide the column count', () => {
    expect(fmt(undefined, undefined, 'a,b,c\n1,2,3\n4,5', false)).toBe('csv')
    expect(fmt(undefined, undefined, 'a,b,c\n1,2,3\n4,5', true)).toBe('text')
  })
})

describe('text decoding', () => {
  test('BOMs, declared charset, UTF-8, else Windows-1252', () => {
    expect(decodeText(new Uint8Array([0xef, 0xbb, 0xbf, 0x61]))).toBe('a')
    expect(decodeText(new Uint8Array([0xff, 0xfe, 0x61, 0x00, 0xe9, 0x00]))).toBe('aé')
    expect(decodeText(new Uint8Array([0xfe, 0xff, 0x00, 0x61, 0x00, 0xe9]))).toBe('aé')
    expect(decodeText(enc('Hà Nội'))).toBe('Hà Nội')
    expect(decodeText(new Uint8Array([0x93, 0x63, 0x61, 0x66, 0xe9, 0x94]))).toBe('“café”')
    expect(decodeText(new Uint8Array([0xe9]), 'text/csv; charset=iso-8859-1')).toBe('é')
  })
})

describe('HTML without a DOM', () => {
  test('implied end tags, entities, raw text, head', () => {
    const p = htmlToParas(
      '<html><head><title>T</title><style>p{color:red}</style><script>if (a<b) x()</script></head>' +
        '<body><h1>A&nbsp;&lt;b&gt; &#8211; &#x41;&eacute;&unknown;</h1><p>one<p>two<ul><li>x<li>y</ul>' +
        '<table><tr><td>c1<td>c2<tr><td>d1</table><pre>\n  kept\n  lines</pre></body></html>',
    )
    expect(p.map(textOf)).toEqual(['A\u00a0<b> – Aé&unknown;', 'one', 'two', 'x', 'y', 'c1\tc2', 'd1', 'kept\n  lines'])
    expect(p[0].heading).toBe(1)
    expect(p.slice(3, 5).map((x) => x.list)).toEqual(['bullet', 'bullet'])
  })

  test('inline styles, links, and no script links', () => {
    const p = htmlToParas(
      '<p><b>b</b><span style="font-weight:700; color: #FF0000; font-size:12pt">s</span><a href="https://x.y?a=1&amp;b=2">l</a>' +
        '<a href=" java\tscript:alert(1)">bad</a><a href=\'#top\'>top</a><font color="#00f" face="Georgia, serif">f</font></p>',
    )
    const runs = p[0].runs.filter((r) => r.t === 'text') as any[]
    expect(runs.map((r) => r.text)).toEqual(['b', 's', 'l', 'bad', 'top', 'f'])
    expect(runs[0].d).toEqual({ b: true })
    expect(runs[1].d).toEqual({ b: true, color: '#ff0000', sz: 12 })
    expect(runs[2].href).toBe('https://x.y?a=1&b=2')
    expect(runs[3].href).toBeUndefined()
    expect(runs[4].href).toBe('#top')
    expect(runs[5].d).toMatchObject({ color: '#0000ff', font: 'Georgia' })
  })

  test('text after the head, without <body>, is shown', () => {
    expect(htmlToParas('<head><title>x</title><p>shown').map(textOf)).toEqual(['shown'])
    expect(htmlToParas('<head><meta charset="utf-8">Loose text').map(textOf)).toEqual(['Loose text'])
  })

  test('tags in attribute values and stray "<"', () => {
    const root = parseHtml('<p title="a > b" data-x=\'<i>\'>1 < 2</p>') as any
    expect(root.childNodes[0].getAttribute('title')).toBe('a > b')
    expect(root.childNodes[0].childNodes[0].textContent).toBe('1 < 2')
  })
})
