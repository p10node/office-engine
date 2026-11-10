import { describe, expect, test } from 'bun:test'
import { deflateSync } from 'node:zlib'
import type { Block, DocxMeta, Inline, Paragraph, Table } from '../src/core/docx/types'
import { rtfToDocx } from '../src/core/legacy/rtf'
import { BlobSource } from '../src/core/source'
import { Engine } from '../src/core/worker/engine'
import { ZipArchive } from '../src/core/zip'
import { crc32 } from '../src/core/zipwrite'

/** each character one byte (RTF is 7-bit text with \'hh escapes; 8-bit bytes for code page tests) */
const bytes = (s: string) => Uint8Array.from(s, (c) => c.charCodeAt(0) & 0xff)

const FONTS = String.raw`{\fonttbl{\f0\froman\fcharset0 Times New Roman;}{\f1\fswiss\fcharset0 Arial;}{\f2\fnil\fcharset2 Symbol;}{\f3\fnil\fcharset128 MS Mincho;}}`
const COLORS = String.raw`{\colortbl;\red255\green0\blue0;\red0\green0\blue255;\red255\green255\blue0;\red200\green200\blue200;}`
/** an RTF document around a body (and extra header tables) */
const doc = (body: string, head = '') => String.raw`{\rtf1\ansi\ansicpg1252\deff0` + FONTS + COLORS + head + '\n' + body + '}'

async function open(rtf: string | Uint8Array) {
  const events: { event: string; data: any }[] = []
  const engine = new Engine((event, data) => events.push({ event, data }))
  const info = await engine.open(new Blob([typeof rtf === 'string' ? bytes(rtf) : (rtf as Uint8Array<ArrayBuffer>)]), { name: 'doc.rtf' })
  await engine.call('docx.start', [])
  const done = events.find((e) => e.event === 'docx:done')?.data
  if (done?.error) throw new Error(done.error)
  const blocks: Block[] = events.filter((e) => e.event === 'docx:blocks').flatMap((e) => e.data)
  const meta: DocxMeta = events.filter((e) => e.event === 'docx:meta').at(-1)?.data
  return { info, blocks, meta, paras: blocks.filter((b): b is Paragraph => b.t === 'p'), tables: blocks.filter((b): b is Table => b.t === 'tbl') }
}

/** a part of the converted package */
async function part(rtf: string | Uint8Array, name = 'word/document.xml'): Promise<string> {
  const zip = await ZipArchive.open(new BlobSource(await rtfToDocx(typeof rtf === 'string' ? bytes(rtf) : rtf)))
  return zip.has(name) ? new TextDecoder().decode(await zip.bytes(name)) : ''
}

async function names(rtf: string | Uint8Array): Promise<string[]> {
  const zip = await ZipArchive.open(new BlobSource(await rtfToDocx(typeof rtf === 'string' ? bytes(rtf) : rtf)))
  return zip.list().map((e) => e.name)
}

const textOf = (p: Paragraph) =>
  p.runs.map((r) => (r.t === 'text' ? r.text : r.t === 'field' ? `{${r.kind}}` : r.t === 'ref' ? `[${r.text}]` : r.t === 'img' ? '<img>' : '')).join('')
const texts = (paras: Paragraph[]) => paras.map(textOf)
const runs = (p: Paragraph) => p.runs.filter((r): r is Extract<Inline, { t: 'text' }> => r.t === 'text')
const run = (p: Paragraph, text: string) => runs(p).find((r) => r.text === text)!
const cellText = (t: Table, r: number, c: number) => t.rows[r].cells[c].blocks.map((b) => (b.t === 'p' ? textOf(b) : '<table>')).join('|')

/** a PNG of w × h pixels */
function png(w: number, h: number): Uint8Array {
  const chunk = (type: string, data: Uint8Array) => {
    const body = new Uint8Array(4 + data.length)
    body.set(bytes(type))
    body.set(data, 4)
    const out = new Uint8Array(12 + data.length)
    new DataView(out.buffer).setUint32(0, data.length)
    out.set(body, 4)
    new DataView(out.buffer).setUint32(8 + data.length, crc32(body))
    return out
  }
  const ihdr = new Uint8Array(13)
  new DataView(ihdr.buffer).setUint32(0, w)
  new DataView(ihdr.buffer).setUint32(4, h)
  ihdr.set([8, 2, 0, 0, 0], 8)
  const raw = new Uint8Array(h * (1 + w * 3)).fill(0x80)
  const parts = [Uint8Array.of(0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a), chunk('IHDR', ihdr), chunk('IDAT', new Uint8Array(deflateSync(raw))), chunk('IEND', new Uint8Array())]
  const out = new Uint8Array(parts.reduce((n, p) => n + p.length, 0))
  let k = 0
  for (const p of parts) {
    out.set(p, k)
    k += p.length
  }
  return out
}
const hex = (b: Uint8Array) => Buffer.from(b).toString('hex').replace(/(.{64})/g, '$1\n')
/** a 24-bit DIB (BITMAPINFOHEADER + pixels) of w × h */
function dib(w: number, h: number): Uint8Array {
  const row = Math.ceil((w * 3) / 4) * 4
  const b = new Uint8Array(40 + row * h)
  const v = new DataView(b.buffer)
  v.setUint32(0, 40, true)
  v.setInt32(4, w, true)
  v.setInt32(8, h, true)
  v.setUint16(12, 1, true)
  v.setUint16(14, 24, true)
  v.setUint32(20, row * h, true)
  return b
}

describe('RTF tokenizer', () => {
  test('control words with parameters, control symbols, hex bytes', async () => {
    const { paras } = await open(doc(String.raw`\pard\fi-360\li720 caf\'e9 \{braces\} back\\slash and\~space\par`))
    expect(textOf(paras[0])).toBe('caf\u00e9 {braces} back\\slash and\u00a0space')
    expect(paras[0].ps.indFirst).toBe(-18)
    expect(paras[0].ps.indL).toBe(36)
  })

  test('\\u with \\uc skip counts, also inherited by nested groups', async () => {
    const { paras } = await open(doc(String.raw`\pard\uc2 A\u8364\'80\'80 B{\u8212\'97\'97 C}\uc1 \u8220?x\u-10179 ?\u-8704 ?\par`))
    expect(textOf(paras[0])).toBe('A\u20ac B\u2014 C\u201cx\ud83d\ude00')
  })

  test('\\bin data, ignorable and ignored destinations are skipped', async () => {
    const { paras } = await open(
      doc(String.raw`\pard A{\*\unknowndest hidden {nested} text}B\bin3 x}yC{\info{\title Title}{\author Me}}D{\*\generator Writer;}E{\xe index}F\par`),
    )
    expect(textOf(paras[0])).toBe('ABCDEF')
  })

  test('deep nesting, unbalanced and truncated files are read without failing', async () => {
    const deep = await open(doc('\\pard ' + '{'.repeat(5000) + 'deep' + '}'.repeat(5000) + ' end\\par'))
    expect(textOf(deep.paras[0])).toBe(' end')
    const truncated = await open(String.raw`{\rtf1\ansi abc {\b def`)
    expect(textOf(truncated.paras[0])).toBe('abc def')
    expect(run(truncated.paras[0], 'def').s.b).toBe(true)
    const extra = await open(String.raw`{\rtf1 one\par}}} two\par}`)
    expect(texts(extra.paras)).toEqual(['one'])
  })

  test('files that are not RTF are rejected', async () => {
    await expect(rtfToDocx(bytes('hello world'))).rejects.toThrow('Not an RTF file')
    await expect(rtfToDocx(bytes('{\\x1 nope}'))).rejects.toThrow('Not an RTF file')
  })
})

describe('code pages', () => {
  test('\\ansicpg', async () => {
    const { paras } = await open(String.raw`{\rtf1\ansi\ansicpg1251 \'cf\'f0\'e8\'e2\'e5\'f2\par}`)
    expect(textOf(paras[0])).toBe('Привет')
  })

  test('font charsets: double-byte Shift-JIS, symbol fonts', async () => {
    const { paras } = await open(doc(String.raw`\pard\f3 \'82\'a0\'82\'a2\f0  abc \f2\'b7\par`))
    expect(textOf(paras[0])).toBe('あい abc ')
  })

  test('\\mac and \\pc', async () => {
    expect(textOf((await open(String.raw`{\rtf1\mac \'8e\'87\par}`)).paras[0])).toBe('éá')
    expect(textOf((await open(String.raw`{\rtf1\pc \'82\'f1\par}`)).paras[0])).toBe('é±')
  })
})

const STYLES = String.raw`{\stylesheet{\ql\sa120\f0\fs24 Normal;}{\s1\ql\sb240\keepn\b\f1\fs32\cf2\sbasedon0\snext0\outlinelevel0 heading 1;}{\s2\qc\i\fs28\sbasedon0\snext0 My Title,Alias;}{\*\cs10\additive Default Paragraph Font;}{\*\cs15\additive\ul\cf2\sbasedon10 Hyperlink;}{\*\ts11\tsrowd Normal Table;}}`

describe('styles', () => {
  test('stylesheet entries become Word styles: built-in ids, basedOn / next, complete definitions', async () => {
    const styles = await part(doc(String.raw`\pard\s1\ql\sb240\keepn\outlinelevel0\b\f1\fs32\cf2 Title\par`, STYLES), 'word/styles.xml')
    expect(styles).toContain('<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/>')
    expect(styles).toMatch(/w:styleId="Heading1"><w:name w:val="heading 1"\/><w:basedOn w:val="Normal"\/><w:next w:val="Normal"\/>/)
    // an entry states the style in full: "spacing after" left out of heading 1 is 0, not Normal's 120
    expect(styles).toMatch(/w:styleId="Heading1">.*<w:spacing w:before="240" w:after="0"/)
    expect(styles).toContain('w:styleId="MyTitle"><w:name w:val="My Title"/>')
    expect(styles).toMatch(/w:type="character" w:styleId="Hyperlink">.*<w:color w:val="0000FF"\/><w:u w:val="single"\/>/)
    expect(styles).not.toContain('Normal Table')
  })

  test('paragraphs and runs are written as their difference from their style', async () => {
    const rtf = doc(
      String.raw`\pard\s1\ql\sb240\keepn\outlinelevel0\b\f1\fs32\cf2 Bold heading {\b0 plain part}\par` +
        String.raw`\pard\plain\s1\ql\sb240\keepn\outlinelevel0\f1\fs32\cf2 Heading without bold\par` +
        String.raw`\pard\plain\ql\sa120 Body {\cs15\ul\cf2 styled} text\par`,
      STYLES,
    )
    const xml = await part(rtf)
    expect(xml).toContain('<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t xml:space="preserve">Bold heading </w:t></w:r>')
    expect(xml).toContain('<w:rStyle w:val="Hyperlink"/></w:rPr><w:t xml:space="preserve">styled</w:t>')
    const { paras, meta } = await open(rtf)
    expect(paras.map((p) => p.sid)).toEqual(['Heading1', 'Heading1', 'Normal'])
    expect(run(paras[0], 'Bold heading ').s.b).toBe(true)
    expect(run(paras[0], 'plain part').s.b).toBe(false)
    // \plain left bold out: the style's bold is turned off
    expect(run(paras[1], 'Heading without bold').s.b).toBe(false)
    expect(run(paras[1], 'Heading without bold').s.color).toBe('#0000ff')
    expect(run(paras[2], 'styled').s).toMatchObject({ u: 'single', color: '#0000ff' })
    expect(paras[0].ps).toMatchObject({ before: 12, after: 0, keepNext: true, outline: 0 })
    expect(meta.styles!.find((s) => s.id === 'Heading1')?.canon).toBe('h1')
  })
})

describe('character formatting', () => {
  test('bold, italic, underlines, strike, caps, size, font, colors, highlight, position, hidden, spacing', async () => {
    const { paras } = await open(
      doc(
        String.raw`\pard {\b b}{\i i}{\ul u}{\uldb d}{\ulwave w}{\ul\ulnone n}{\strike s}{\striked1 ds}{\caps c}{\scaps sc}{\fs40 big}{\f1 arial}` +
          String.raw`{\cf1 red}{\highlight3 hl}{\cb4 shade}{\chcbpat1 shd2}{\super sup}{\sub sub}{\up6 up}{\v hidden}{\expndtw20 wide}{\b\i\plain plain}\par`,
      ),
    )
    const p = paras[0]
    expect(textOf(p)).not.toContain('hidden')
    expect(run(p, 'b').s.b).toBe(true)
    expect(run(p, 'i').s.i).toBe(true)
    expect(run(p, 'u').s.u).toBe('single')
    expect(run(p, 'd').s.u).toBe('double')
    expect(run(p, 'w').s.u).toBe('wave')
    expect(run(p, 'n').s.u ?? 'none').toBe('none')
    expect(run(p, 's').s.strike).toBe(true)
    expect(run(p, 'ds').s.dstrike).toBe(true)
    expect(run(p, 'c').s.caps).toBe(true)
    expect(run(p, 'sc').s.smallCaps).toBe(true)
    expect(run(p, 'big').s.sz).toBe(20)
    expect(run(p, 'arial').s.font).toContain('Arial')
    expect(run(p, 'red').s.color).toBe('#ff0000')
    expect(run(p, 'hl').s.hl).toBe('#ffff00')
    expect(run(p, 'shade').s.shd).toBe('#c8c8c8')
    expect(run(p, 'shd2').s.shd).toBe('#ff0000')
    expect(run(p, 'sup').s.va).toBe('sup')
    expect(run(p, 'sub').s.va).toBe('sub')
    expect(run(p, 'wide').s.spacing).toBe(1)
    expect(run(p, 'plain').s).toEqual({ font: expect.stringContaining('Times New Roman'), sz: 12 })
    const xml = await part(doc(String.raw`\pard {\up6 up}{\dn4 down}{\v hidden}{\outl o}\par`))
    expect(xml).toContain('<w:position w:val="6"/>')
    expect(xml).toContain('<w:position w:val="-4"/>')
    expect(xml).toContain('<w:vanish/>')
    expect(xml).toContain('<w:outline/>')
  })
})

describe('paragraph formatting', () => {
  test('alignment, indents, spacing, line spacing, keep, page break before', async () => {
    const { paras } = await open(
      doc(
        String.raw`\pard\qc centered\par\pard\qr right\par\pard\qj\li720\ri360\fi360 justified\par` +
          String.raw`\pard\sb240\sa120\sl360\slmult1 spaced\par\pard\sl-300 exact\par\pard\sl300 least\par\pard\keepn\pagebb broken\par`,
      ),
    )
    expect(paras.map((p) => p.ps.align)).toEqual(['center', 'right', 'justify', undefined, undefined, undefined, undefined])
    expect(paras[2].ps).toMatchObject({ indL: 36, indR: 18, indFirst: 18 })
    expect(paras[3].ps).toMatchObject({ before: 12, after: 6, line: 1.5, lineRule: 'auto' })
    expect(paras[4].ps).toMatchObject({ line: 15, lineRule: 'exact' })
    expect(paras[5].ps).toMatchObject({ line: 15, lineRule: 'atLeast' })
    expect(paras[6].ps.keepNext).toBe(true)
    expect(paras[6].pb).toBe(true)
  })

  test('tabs, borders and shading', async () => {
    const rtf = doc(String.raw`\pard\tqr\tldot\tx9000\tqc\tx4000\tx1000 a\tab b\par\pard\brdrt\brdrs\brdrw15\brdrcf1\brsp40\brdrb\brdrdb\box\brdrdot\cbpat3 boxed\par\pard\cfpat1\shading5000 half\par`)
    const xml = await part(rtf)
    expect(xml).toContain('<w:tabs><w:tab w:val="left" w:pos="1000"/><w:tab w:val="center" w:pos="4000"/><w:tab w:val="right" w:leader="dot" w:pos="9000"/></w:tabs>')
    const { paras } = await open(rtf)
    expect(textOf(paras[0])).toBe('a\tb')
    expect(paras[1].ps.bTop).toMatchObject({ style: 'dotted' })
    expect(paras[1].ps.bLeft).toMatchObject({ style: 'dotted' })
    expect(paras[1].ps.shd).toBe('#ffff00')
    // 50% red pattern over white
    expect(paras[2].ps.shd).toBe('#ff8080')
  })

  test('line, page and column breaks, tabs and special characters', async () => {
    const rtf = doc(String.raw`\pard one\line two\page three\column four\tab five\emdash\endash\bullet\lquote\rquote\ldblquote\rdblquote soft\-hyphen non\_breaking\par`)
    expect(await part(rtf)).toContain('<w:br w:type="column"/>')
    const { paras } = await open(rtf)
    expect(texts(paras)).toEqual(['one\ntwo', 'three\nfour\tfive\u2014\u2013\u2022\u2018\u2019\u201c\u201dsoft\u00adhyphen non\u2011breaking'])
    expect(paras[1].pb).toBe(true)
  })
})

const LISTS = String.raw`{\*\listtable{\list\listtemplateid1\listhybrid{\listlevel\levelnfc0\leveljc0\levelfollow0\levelstartat1{\leveltext\'02\'00.;}{\levelnumbers\'01;}\fi-360\li720}{\listlevel\levelnfc4\leveljc0\levelfollow0\levelstartat1{\leveltext\'03(\'01);}{\levelnumbers\'02;}\fi-360\li1440}\listid10}{\list\listtemplateid2{\listlevel\levelnfc23\leveljc0\levelfollow0\levelstartat1{\leveltext\'01\u-3913 ?;}{\levelnumbers;}\f2\fi-360\li720}\listid20}}{\*\listoverridetable{\listoverride\listid10\listoverridecount0\ls1}{\listoverride\listid20\listoverridecount0\ls2}{\listoverride\listid10\listoverridecount1{\lfolevel\listoverridestartat\levelstartat5}\ls3}}`

describe('lists', () => {
  test('list tables and overrides', async () => {
    const { paras } = await open(
      doc(
        String.raw`\pard\ls1\ilvl0\fi-360\li720{\listtext 1.\tab}one\par\pard\ls1\ilvl1{\listtext (a)\tab}sub\par\pard\ls1\ilvl0 two\par` +
          String.raw`\pard\ls2{\listtext\f2 \'b7\tab}bullet\par\pard\ls3 restarted\par\pard plain\par`,
        LISTS,
      ),
    )
    expect(texts(paras)).toEqual(['one', 'sub', 'two', 'bullet', 'restarted', 'plain'])
    expect(paras.map((p) => p.marker?.text)).toEqual(['1.', '(a)', '2.', '•', '5.', undefined])
    expect(paras[0].ps).toMatchObject({ indL: 36, indFirst: -18 })
    expect(paras[1].ps).toMatchObject({ indL: 72, indFirst: -18 })
    expect(paras[3].lst).toBe('bullet')
    const numbering = await part(doc('\\pard\\ls1 x\\par', LISTS), 'word/numbering.xml')
    expect(numbering).toContain('<w:lvlText w:val="%1."/>')
    expect(numbering).toContain('<w:lvlText w:val="(%2)"/>')
    expect(numbering).toContain('<w:num w:numId="3"><w:abstractNumId w:val="0"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="5"/></w:lvlOverride></w:num>')
  })

  test('legacy \\pn numbering (Word 95, WordPad)', async () => {
    const { paras } = await open(
      doc(
        String.raw`{\pntext\f2\'b7\tab}{\*\pn\pnlvlblt\pnf2\pnindent0{\pntxtb\'b7}}\fi-360\li720 bullet\par` +
          String.raw`\pard{\pntext 1.\tab}{\*\pn\pnlvlbody\pndec\pnstart1{\pntxta .}}\fi-360\li720 first\par` +
          String.raw`{\pntext 2.\tab}{\*\pn\pnlvlbody\pndec\pnstart1{\pntxta .}}second\par` +
          String.raw`{\*\pn\pnlvlcont}continued\par\pard plain\par`,
      ),
    )
    expect(texts(paras)).toEqual(['bullet', 'first', 'second', 'continued', 'plain'])
    expect(paras.map((p) => p.marker?.text)).toEqual(['•', '1.', '2.', undefined, undefined])
  })
})

describe('tables', () => {
  test('cell widths, row position, borders, shading, header rows, heights', async () => {
    const { tables } = await open(
      doc(
        String.raw`\trowd\trgaph108\trleft-108\trhdr\trrh400\clbrdrt\brdrs\brdrw10\clbrdrb\brdrdb\clcbpat3\cellx2000\clvertalc\clbrdrl\brdrs\brdrcf1\cellx5000` +
          String.raw`\pard\intbl Head A\cell Head B\cell\row\trowd\trgaph108\trleft-108\trrh-300\cellx2000\cellx5000\pard\intbl a\cell b\cell\row\pard after\par`,
      ),
    )
    const t = tables[0]
    expect(t.cols).toEqual([105.4, 150])
    expect(t.indL).toBe(-5.4)
    expect(t.rows[0]).toMatchObject({ header: true, h: 20, exact: false })
    expect(t.rows[1]).toMatchObject({ h: 15, exact: true })
    expect(cellText(t, 0, 0)).toBe('Head A')
    expect(cellText(t, 1, 1)).toBe('b')
    const c = t.rows[0].cells[0]
    expect(c.shd).toBe('#ffff00')
    expect(c.borders.t).toBe('0.5pt solid #000000')
    expect(c.borders.b).toBe('2.25pt double #000000')
    expect(t.rows[0].cells[1].borders.l).toBe('0.5pt solid #ff0000')
    expect(t.rows[0].cells[1].vAlign).toBe('middle')
    expect(t.rows[0].cells[0].pad).toEqual([0, 5.4, 0, 5.4])
  })

  test('horizontal and vertical merges', async () => {
    const { tables } = await open(
      doc(
        String.raw`\trowd\clmgf\cellx1000\clmrg\cellx2000\clvmgf\cellx3000\pard\intbl wide\cell\cell tall\cell\row` +
          String.raw`\trowd\cellx1000\cellx2000\clvmrg\cellx3000\pard\intbl a\cell b\cell\cell\row`,
      ),
    )
    const t = tables[0]
    expect(t.rows[0].cells.map((c) => [c.colSpan, c.rowSpan])).toEqual([[2, 1], [1, 2]])
    expect(t.rows[1].cells.length).toBe(2)
    expect(cellText(t, 0, 0)).toBe('wide')
    expect(cellText(t, 0, 1)).toBe('tall')
  })

  test('rows without a definition of their own repeat the last one (Word 97)', async () => {
    const { tables, paras } = await open(doc(String.raw`\trowd\cellx1000\cellx2000\pard\intbl a\cell b\cell\row\intbl c\cell d\cell\row\pard after\par`))
    expect(tables[0].rows.length).toBe(2)
    expect(cellText(tables[0], 1, 1)).toBe('d')
    expect(texts(paras)).toEqual(['after'])
  })

  test('rows with differing cells share a grid', async () => {
    const { tables } = await open(doc(String.raw`\trowd\cellx1000\cellx3000\pard\intbl a\cell b\cell\row\trowd\cellx2000\cellx3000\pard\intbl c\cell d\cell\row\pard\par`))
    const t = tables[0]
    expect(t.cols).toEqual([50, 50, 50])
    expect(t.rows[0].cells.map((c) => c.colSpan)).toEqual([1, 2])
    expect(t.rows[1].cells.map((c) => c.colSpan)).toEqual([2, 1])
  })

  test('nested tables', async () => {
    const { tables } = await open(
      doc(
        String.raw`\trowd\cellx4000\cellx6000\pard\intbl\itap1 outer\par` +
          String.raw`\pard\intbl\itap2 inner a\nestcell inner b\nestcell{\*\nesttableprops\trowd\cellx1000\cellx2000\nestrow}{\nonesttables\par}` +
          String.raw`\pard\intbl\itap1 below\cell right\cell{\trowd\cellx4000\cellx6000\row}\pard\itap0 after\par`,
      ),
    )
    const t = tables[0]
    expect(t.cols).toEqual([200, 100])
    expect(cellText(t, 0, 0)).toBe('outer|<table>|below')
    const inner = t.rows[0].cells[0].blocks[1] as Table
    expect(inner.cols).toEqual([50, 50])
    expect(cellText(inner, 0, 1)).toBe('inner b')
    expect(cellText(t, 0, 1)).toBe('right')
  })

  test('a paragraph closed before \\cell does not leave an empty one (pandoc)', async () => {
    const { tables } = await open(doc(String.raw`{\trowd\trgaph120\cellx2000\cellx4000\trkeep\intbl{{\pard\intbl a\par}\cell}{{\pard\intbl b\par}\cell}\intbl\row}`))
    expect(tables[0].rows[0].cells.map((c) => c.blocks.length)).toEqual([1, 1])
  })
})

describe('sections, headers and footers', () => {
  test('page size, margins, landscape, columns', async () => {
    const { meta } = await open(doc(String.raw`\paperw11906\paperh16838\margl1134\margr850\margt1440\margb720\landscape\sectd\cols2\colsx360 text\par`))
    expect(meta.page).toMatchObject({ w: 841.9, h: 595.3, ml: 56.7, mr: 42.5, mt: 72, mb: 36, cols: 2, colGap: 18 })
  })

  test('sections: breaks, their own page setup', async () => {
    const { blocks } = await open(doc(String.raw`\sectd\pgwsxn12240\pghsxn15840 one\sect\sectd\sbknone\lndscpsxn\pgwsxn15840\pghsxn12240 two\par`))
    const sect = blocks.filter((b) => b.t === 'sect')
    expect(sect.length).toBe(1)
    expect(sect[0]).toMatchObject({ page: { w: 612, h: 792 } })
    const xml = await part(doc(String.raw`\sectd one\sect\sectd\sbknone\lndscpsxn\pgwsxn15840\pghsxn12240 two\par`))
    expect(xml).toContain('<w:type w:val="continuous"/><w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/>')
  })

  test('headers and footers, first page, page number fields', async () => {
    const rtf = doc(
      String.raw`\sectd\titlepg{\header\pard\qc Report header\par}{\footer\pard Page {\field{\*\fldinst PAGE}{\fldrslt 1}} of {\field{\*\fldinst NUMPAGES}{\fldrslt 3}} \chpgn\par}{\headerf\pard First page\par}\pard body\par`,
    )
    const { meta } = await open(rtf)
    expect(meta.titlePg).toBe(true)
    expect(textOf(meta.header![0] as Paragraph)).toBe('Report header')
    expect(textOf(meta.footer![0] as Paragraph)).toBe('Page {PAGE} of {NUMPAGES} {PAGE}')
    expect(textOf(meta.firstHeader![0] as Paragraph)).toBe('First page')
    expect(meta.firstFooter).toBeUndefined()
    const files = await names(rtf)
    expect(files).toContain('word/header1.xml')
    expect(files).toContain('word/footer1.xml')
  })

  test('left / right headers of facing pages', async () => {
    const rtf = doc(String.raw`\facingp\sectd{\headerl\pard Left\par}{\headerr\pard Right\par} body\par`)
    const xml = await part(rtf)
    expect(xml).toMatch(/<w:headerReference w:type="default" r:id="rId\d+"\/><w:headerReference w:type="even" r:id="rId\d+"\/>/)
    expect(await part(rtf, 'word/settings.xml')).toContain('<w:evenAndOddHeaders/>')
    expect(textOf((await open(rtf)).meta.header![0] as Paragraph)).toBe('Right')
  })
})

describe('notes, fields, bookmarks', () => {
  test('footnotes and endnotes', async () => {
    const rtf = doc(String.raw`\pard Text{\super\chftn}{\footnote\pard\plain{\super\chftn} The note.}\par More{\super\chftn}{\footnote\ftnalt\pard{\super\chftn} End note {\field{\*\fldinst HYPERLINK "https://example.com/end"}{\fldrslt link}}.}\par`)
    const { paras } = await open(rtf)
    expect(texts(paras)).toEqual(['Text[1]', 'More[1]'])
    const notes = await part(rtf, 'word/footnotes.xml')
    expect(notes).toContain('<w:footnote w:type="separator" w:id="-1">')
    expect(notes).toMatch(/<w:footnote w:id="1"><w:p><w:r><w:rPr><w:vertAlign w:val="superscript"\/><\/w:rPr><w:footnoteRef\/><\/w:r><w:r><w:t xml:space="preserve"> The note.<\/w:t>/)
    expect(await part(rtf, 'word/endnotes.xml')).toContain('End note ')
    // links and pictures of a note are relationships of its own part
    expect(await part(rtf, 'word/_rels/endnotes.xml.rels')).toContain('Target="https://example.com/end"')
    expect(await part(rtf)).toContain('<w:endnoteReference w:id="1"/>')
    expect(await part(rtf, 'word/settings.xml')).toContain('<w:footnotePr><w:footnote w:id="-1"/><w:footnote w:id="0"/></w:footnotePr>')
  })

  test('fields: links (external and to a bookmark), page numbers, other results kept', async () => {
    const rtf = doc(
      String.raw`\pard{\*\bkmkstart target}Target{\*\bkmkend target} {\field{\*\fldinst {HYPERLINK "https://example.com/a?b=1"}}{\fldrslt {\ul site}}} ` +
        String.raw`{\field{\*\fldinst HYPERLINK \\l "target"}{\fldrslt back}} {\field{\*\fldinst DATE \\@ "yyyy"}{\fldrslt 2026}} {\field{\*\fldinst PAGE \\* MERGEFORMAT}{\fldrslt 4}}\par`,
    )
    const { paras } = await open(rtf)
    const p = paras[0]
    expect(textOf(p)).toBe('Target site back 2026 {PAGE}')
    expect(run(p, 'site').href).toBe('https://example.com/a?b=1')
    expect(run(p, 'back').href).toBe('#target')
    const xml = await part(rtf)
    expect(xml).toMatch(/<w:bookmarkStart w:id="(\d+)" w:name="target"\/>.*<w:bookmarkEnd w:id="\1"\/>/)
    expect(xml).toContain('<w:fldSimple w:instr=" PAGE \\* MERGEFORMAT ">')
  })
})

describe('pictures', () => {
  test('PNG as hex: goal size, scale, cropping', async () => {
    const data = hex(png(4, 2))
    const rtf = doc(String.raw`\pard{\pict\pngblip\picw4\pich2\picwgoal1440\pichgoal720\picscalex50\picscaley200 ` + data + String.raw`}{\pict\pngblip\picw4\pich2 ` + data + `}\\par`)
    const { paras } = await open(rtf)
    const imgs = paras[0].runs.filter((r) => r.t === 'img')
    expect(imgs.map((r) => [r.w, r.h])).toEqual([[36, 72], [3, 1.5]])
    // the same picture is stored once
    expect((await names(rtf)).filter((n) => n.startsWith('word/media/'))).toEqual(['word/media/image1.png'])
    const crop = await part(doc(String.raw`\pard{\pict\pngblip\picwgoal1440\pichgoal720\piccropl360\piccropr360 ` + data + '}\\par'))
    expect(crop).toContain('<a:srcRect l="25000" t="0" r="25000" b="0"/>')
    expect(crop).toContain('<wp:extent cx="457200" cy="457200"/>')
  })

  test('binary JPEG data (\\bin), bitmaps (DIB → BMP), a metafile wrapping one bitmap', async () => {
    const jpeg = Uint8Array.of(0xff, 0xd8, 0xff, 0xc0, 0, 17, 8, 0, 20, 0, 30, 3, 1, 0x22, 0, 2, 0x11, 1, 3, 0x11, 1, 0xff, 0xd9)
    const head = bytes(doc('\\pard{\\pict\\jpegblip\\bin' + jpeg.length + ' ').slice(0, -1))
    const tail = bytes(String.raw`}{\pict\dibitmap0\picwgoal300\pichgoal300 ` + hex(dib(2, 2)) + '}\\par}')
    const rtf = new Uint8Array(head.length + jpeg.length + tail.length)
    rtf.set(head)
    rtf.set(jpeg, head.length)
    rtf.set(tail, head.length + jpeg.length)
    const { paras } = await open(rtf)
    const imgs = paras[0].runs.filter((r) => r.t === 'img')
    expect(imgs.map((r) => [r.w, r.h])).toEqual([[22.5, 15], [15, 15]])
    const files = await names(rtf)
    expect(files).toContain('word/media/image1.jpg')
    expect(files).toContain('word/media/image2.bmp')
    // a WMF that only draws a bitmap (META_STRETCHDIB)
    const bmp = dib(3, 1)
    const recWords = 3 + 11 + bmp.length / 2
    const wmf = new Uint8Array(18 + recWords * 2 + 6)
    const v = new DataView(wmf.buffer)
    v.setUint16(0, 1, true)
    v.setUint16(2, 9, true)
    v.setUint16(4, 0x300, true)
    v.setUint32(18, recWords, true)
    v.setUint16(22, 0x0f43, true)
    wmf.set(bmp, 18 + 6 + 22)
    v.setUint32(18 + recWords * 2, 3, true)
    const meta = await names(doc(String.raw`\pard{\pict\wmetafile8\picwgoal600\pichgoal200 ` + hex(wmf) + '}\\par'))
    expect(meta).toContain('word/media/image1.bmp')
  })

  test('\\shppict is preferred to \\nonshppict, which is used when the first is a metafile', async () => {
    const a = hex(png(1, 1))
    const b = hex(png(2, 2))
    const both = await open(doc(String.raw`\pard{\*\shppict{\pict\pngblip\picwgoal200\pichgoal200 ` + a + String.raw`}}{\nonshppict{\pict\pngblip\picwgoal400\pichgoal400 ` + b + '}}\\par'))
    expect(both.paras[0].runs.filter((r) => r.t === 'img').map((r) => r.w)).toEqual([10])
    const emf = await open(doc(String.raw`\pard{\*\shppict{\pict\emfblip 0100000000}}{\nonshppict{\pict\pngblip\picwgoal400\pichgoal400 ` + b + '}}\\par'))
    expect(emf.paras[0].runs.filter((r) => r.t === 'img').map((r) => r.w)).toEqual([20])
    const wmf = await open(doc(String.raw`\pard{\pict\wmetafile8 0100090000035500000000000000}\par`))
    expect(wmf.paras[0].runs.filter((r) => r.t === 'img')).toEqual([])
  })

  test('floating pictures and text boxes (\\shp)', async () => {
    const rtf = doc(
      String.raw`\pard before{\shp{\*\shpinst\shpleft100\shptop200\shpright1540\shpbottom920\shpbxcolumn\shpbypara\shpwr2{\sp{\sn shapeType}{\sv 75}}{\sp{\sn pib}{\sv {\pict\pngblip ` +
        hex(png(2, 1)) +
        String.raw`}}}}{\shprslt fallback}} after{\shp{\*\shpinst\shpleft0\shptop0\shpright2880\shpbottom1440{\sp{\sn fillColor}{\sv 255}}{\shptxt\pard boxed text\par}}{\shprslt\pard fallback\par}}\par`,
    )
    const { paras } = await open(rtf)
    const p = paras[0]
    expect(textOf(p)).toBe('before<img> after')
    expect(p.runs.find((r) => r.t === 'img')).toMatchObject({ w: 72, h: 36 })
    const box = p.runs.find((r) => r.t === 'box')!
    expect(box).toMatchObject({ w: 144, h: 72, fill: '#FF0000' })
    expect(textOf((box as { blocks: Block[] }).blocks[0] as Paragraph)).toBe('boxed text')
    expect(await part(rtf)).toContain('<wp:positionH relativeFrom="column"><wp:posOffset>63500</wp:posOffset></wp:positionH>')
  })
})

/** 2016-01-29 18:50 as an RTF date (minutes, hours, day, month, years since 1900 in bit fields; Friday) */
const DTTM = (50 | (18 << 6) | (29 << 11) | (1 << 16) | (116 << 20) | (5 << 29)) >>> 0

describe('tracked changes', () => {
  test('\\revised and \\deleted become w:ins / w:del with their author and date', async () => {
    const rtf = doc(String.raw`\pard keep {\revised\revauth1\revdttm${DTTM} added}{\deleted\revauthdel2 removed} end\par`, String.raw`{\*\revtbl {Unknown;}{Ann Author;}{Bob Reviewer;}}`)
    const xml = await part(rtf)
    expect(xml).toMatch(/<w:ins w:id="\d+" w:author="Ann Author" w:date="2016-01-29T18:50:00Z"><w:r><w:t xml:space="preserve">added<\/w:t><\/w:r><\/w:ins>/)
    expect(xml).toMatch(/<w:del w:id="\d+" w:author="Bob Reviewer"><w:r><w:delText xml:space="preserve">removed<\/w:delText><\/w:r><\/w:del>/)
    const { paras } = await open(rtf)
    expect(paras[0].runs.map((r) => (r.t === 'text' ? [r.text, r.tc?.k] : []))).toEqual([['keep ', undefined], ['added', 'ins'], ['removed', 'del'], [' end', undefined]])
  })
})

const fixture = async (name: string) => new Uint8Array(await Bun.file(`${import.meta.dir}/fixtures/legacy/${name}`).arrayBuffer())

describe('fixtures', () => {
  test('sample.rtf (Word layout): styles, lists, table, picture, header / footer, footnote, links', async () => {
    const { info, paras, tables, meta } = await open(await fixture('sample.rtf'))
    expect(info.converted).toBe('rtf')
    expect(info.kind).toBe('docx')
    expect(paras[0]).toMatchObject({ sid: 'Title' })
    expect(textOf(paras[0])).toBe('Quarterly Report')
    const headings = paras.filter((p) => p.sid === 'Heading1' || p.sid === 'Heading2').map(textOf)
    expect(headings).toEqual(['Highlights', 'Next steps', 'Figures', 'Table'])
    expect(paras.filter((p) => p.marker).map((p) => p.marker!.text)).toEqual(['•', '•', '◦', '•', '1.', '2.', 'a.', '3.'])
    const body = paras[1]
    expect(textOf(body)).toContain('café € and “quotes”. Visit example.com.[1]')
    expect(run(body, 'example.com').href).toBe('https://example.com')
    expect(run(body, 'bold').s.b).toBe(true)
    expect(run(body, '2').s.va).toBeDefined()
    expect(paras.some((p) => p.runs.some((r) => r.t === 'img'))).toBe(true)
    const t = tables[0]
    expect(t.rows.length).toBe(5)
    expect(t.rows[0].header).toBe(true)
    expect(t.rows[0].cells[0].shd).toBe('#d9e2f3')
    expect(t.rows[4].cells[0].colSpan).toBe(3)
    expect(cellText(t, 4, 1)).toBe('750')
    expect(textOf(meta.header![0] as Paragraph)).toBe('Office Engine\tQuarterly Report\t2026')
    expect(textOf(meta.footer![0] as Paragraph)).toBe('Page {PAGE} of {NUMPAGES}')
    expect(meta.styles!.find((s) => s.id === 'Normal')?.preview?.font).toContain('Calibri')
    expect(meta.styles!.map((s) => s.canon).filter(Boolean)).toEqual(expect.arrayContaining(['normal', 'title', 'h1', 'h2']))
    const back = paras.find((p) => textOf(p).startsWith('Note: '))!
    expect(run(back, 'Highlights').href).toBe('#highlights')
    expect(back.ps.shd).toBe('#f2f2f2')
    expect(paras.find((p) => textOf(p) === 'Appendix')?.pb).toBe(true)
  })

  test('textutil.rtf (Cocoa): formatting, lists, table, link', async () => {
    const { paras, tables } = await open(await fixture('textutil.rtf'))
    expect(textOf(paras[0])).toBe('Project Notes')
    expect(runs(paras[0])[0].s).toMatchObject({ b: true, sz: 24 })
    expect(run(paras[1], 'a link').href).toBe('https://example.com/docs')
    expect(run(paras[1], 'red').s.color).toBe('#b00004')
    expect(paras.filter((p) => p.marker).map((p) => [textOf(p), p.marker!.text])).toEqual([
      ['Write the parser', '•'],
      ['Test real files', '•'],
      ['Word and pandoc', '◦'],
      ['First step', '1'],
      ['Second step', '2'],
    ])
    expect(paras.find((p) => textOf(p) === 'Centered line')?.ps.align).toBe('center')
    expect(cellText(tables[0], 1, 0)).toBe('Alpha')
    expect(tables[0].rows[2].cells[0].colSpan).toBe(2)
    expect(textOf(paras.at(-1)!)).toBe('Café costs 5 € — done.')
  })

  test('pandoc.rtf: picture, lists, table, link', async () => {
    const { paras, tables } = await open(await fixture('pandoc.rtf'))
    expect(textOf(paras[0])).toBe('Quarterly Report')
    expect(paras.some((p) => p.runs.some((r) => r.t === 'img' && r.w === 324 && r.h === 162))).toBe(true)
    expect(paras.filter((p) => p.marker).length).toBe(6)
    expect(cellText(tables[0], 0, 0)).toBe('Region')
    expect(tables[0].rows[0].cells[0].blocks.length).toBe(1)
    expect(paras.some((p) => runs(p).some((r) => r.href === 'https://example.com'))).toBe(true)
  })
})

describe('large files', () => {
  test('a few megabytes of text and hex pictures convert quickly', async () => {
    const pic = hex(png(64, 64))
    const parts: string[] = []
    for (let k = 0; k < 6000; k++) {
      parts.push(String.raw`{\pard\sa120 Paragraph ${k} with {\b bold} and {\i italic} text caf\'e9 \u8364\'80.\par}` + '\n')
      if (k % 100 === 0) parts.push(String.raw`{\pard\qc{\pict\pngblip\picwgoal720\pichgoal720 ` + pic + '}\\par}\n')
    }
    const rtf = doc(parts.join(''))
    const t = performance.now()
    const blob = await rtfToDocx(bytes(rtf))
    expect(performance.now() - t).toBeLessThan(10000)
    expect(blob.size).toBeGreaterThan(1000)
    const { paras } = await open(rtf)
    expect(paras.length).toBe(6060)
  }, 30000)
})
