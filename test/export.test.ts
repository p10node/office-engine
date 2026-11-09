import { describe, expect, test } from 'bun:test'
import type { Block, Paragraph } from '../src/core/docx/types'
import { htmlToParas } from '../src/core/html'
import { Engine } from '../src/core/worker/engine'
import { docxToHtml, docxToMarkdown, pictureData } from '../src/react/file/text'

const enc = (s: string) => new TextEncoder().encode(s)

async function blocksOf(text: string, name: string) {
  const events: { event: string; data: any }[] = []
  const engine = new Engine((event, data) => events.push({ event, data }))
  await engine.open(enc(text), { name })
  await engine.call('docx.start', [])
  const blocks: Block[] = events.filter((e) => e.event === 'docx:blocks').flatMap((e) => e.data)
  return { engine, blocks }
}

const MD = '# Title\n\nSome **bold** and *it* with a [link](https://example.com) and `code`.\n\n- one\n- two\n\n1. first\n2. second\n\n```mermaid\ngraph LR\n  A --> B\n```\n'

/** a paragraph of plain runs, as the engine sends them */
const para = (text: string, extra: Partial<Paragraph> = {}): Paragraph => ({ t: 'p', ps: {}, runs: [{ t: 'text', text, s: {} }], ...extra })
const item = (text: string, depth: number, ordered = false): Paragraph => para(text, { ps: { indL: 36 * (depth + 1) }, marker: { text: ordered ? '1.' : '•', s: {}, tab: true } })

describe('a document saved as HTML', () => {
  test('headings, formatting, links, code, lists and the pictures of a Markdown file', async () => {
    const { engine, blocks } = await blocksOf(MD, 'notes.md')
    const images = await pictureData(blocks, (t) => engine.call('part', [t]) as Promise<Blob>)
    expect([...images.keys()]).toEqual(['word/media/diagram1.svg'])
    expect(images.get('word/media/diagram1.svg')).toStartWith('data:image/svg+xml;base64,')
    const html = docxToHtml(blocks, { title: 'notes', images })
    expect(html).toStartWith('<!doctype html>\n<html><head><meta charset="utf-8">')
    expect(html).toContain('<title>notes</title>')
    // the heading's own color and the link's underline are not formatting of the text
    expect(html).toContain('<h1>Title</h1>')
    expect(html).toContain('<p>Some <strong>bold</strong> and <em>it</em> with a <a href="https://example.com">link</a> and <code>code</code>.</p>')
    expect(html).toContain('<ul>\n<li>one</li>\n<li>two</li></ul>\n<ol>\n<li>first</li>\n<li>second</li></ol>')
    expect(html).toMatch(/<p style="text-align:center"><img src="data:image\/svg\+xml;base64,[A-Za-z0-9+/=]+" alt="Mermaid diagram" width="\d+" height="\d+"><\/p>/)
    // read back, the page holds the same text
    const texts = htmlToParas(html).map((p) => p.runs.map((r) => (r.t === 'text' ? r.text : '')).join(''))
    expect(texts.filter(Boolean)).toEqual(['Title', 'Some bold and it with a link and code.', 'one', 'two', 'first', 'second'])
    // Markdown keeps the picture too, and inline code
    const md = docxToMarkdown(blocks, images)
    expect(md).toContain('with a [link](https://example.com) and `code`.')
    expect(md).toMatch(/\n!\[Mermaid diagram\]\(data:image\/svg\+xml;base64,[A-Za-z0-9+/=]+\)\n/)
    // without pictures given, none is written
    expect(docxToHtml(blocks)).not.toContain('<img')
  })

  test('nested lists, a list of another kind, tables, alignment, escaping', () => {
    const blocks: Block[] = [
      item('a', 0),
      item('b', 0),
      item('b.1', 1),
      item('b.1.x', 2),
      item('b.2', 1),
      item('c', 0),
      item('n1', 0, true),
      para('x < y & "z"', { ps: { align: 'right' } }),
      { t: 'tbl', rows: [{ cells: [{ blocks: [para('A')] }, { blocks: [para('B'), para('C')] }] }] } as unknown as Block,
      para(''),
    ]
    const html = docxToHtml(blocks)
    expect(html).toContain(
      '<ul>\n<li>a</li>\n<li>b\n<ul>\n<li>b.1\n<ul>\n<li>b.1.x</li></ul></li>\n<li>b.2</li></ul></li>\n<li>c</li></ul>\n<ol>\n<li>n1</li></ol>\n<p style="text-align:right">x &lt; y &amp; &quot;z&quot;</p>\n<table>\n<tr><td>A</td><td>B<br>C</td></tr>\n</table>',
    )
  })

  test('run formatting: underline, strike, super / subscript, colors, line breaks; deleted text left out', () => {
    const p: Paragraph = {
      t: 'p',
      ps: {},
      runs: [
        { t: 'text', text: 'u', s: { u: 'single' } },
        { t: 'text', text: 's', s: { strike: true } },
        { t: 'text', text: '2', s: { va: 'sup' } },
        { t: 'text', text: 'red', s: { color: '#ff0000', hl: 'yellow' } },
        { t: 'text', text: 'gone', s: {}, tc: { k: 'del' } },
        { t: 'text', text: 'a\nb', s: {} },
        { t: 'text', text: 'bo', s: { b: true } },
        { t: 'text', text: 'ld', s: { b: true } },
      ],
    }
    expect(docxToHtml([p])).toContain('<p><u>u</u><s>s</s><sup>2</sup><span style="color:#ff0000;background:yellow">red</span>a<br>b<strong>bold</strong></p>')
  })
})
