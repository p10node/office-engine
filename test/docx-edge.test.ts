import { describe, expect, test } from 'bun:test'
import { Engine } from '../src/core/worker/engine'
import type { Block, DocxMeta, HFSetResult, Paragraph } from '../src/core/docx/types'
import { ZipArchive } from '../src/core/zip'
import { BlobSource } from '../src/core/source'
import { rebuildZip } from '../src/core/zipwrite'

const S = (n: string) => `${import.meta.dir}/../playground/public/samples/${n}`
const DOCX = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'

async function open(src: Blob) {
  const events: { event: string; data: any }[] = []
  const engine = new Engine((event, data) => events.push({ event, data }))
  await engine.open(src, {})
  await engine.call('docx.start', [])
  const blocks: Block[] = events.filter((e) => e.event === 'docx:blocks').flatMap((e) => e.data)
  const meta = events.filter((e) => e.event === 'docx:meta').at(-1)!.data as DocxMeta
  return { engine, blocks, meta }
}

/** the sample with document.xml rewritten by `fn` */
async function variant(name: string, fn: (xml: string) => string): Promise<Blob> {
  const zip = await ZipArchive.open(new BlobSource(Bun.file(S(name)) as unknown as Blob))
  return rebuildZip(zip, new Map([['word/document.xml', fn(await zip.text('word/document.xml'))]]), DOCX)
}
const mainXml = async (b: Blob) => (await ZipArchive.open(new BlobSource(b))).text('word/document.xml')
const lineOf = (b: Block | undefined) => (b?.t === 'p' ? b.runs.map((r) => (r.t === 'text' ? r.text : r.t === 'field' ? `{${r.kind}}` : '')).join('') : '')

describe('documents without a section (no body sectPr)', () => {
  const noSect = () => variant('sample.docx', (x) => x.replace(/<w:sectPr[\s\S]*?<\/w:sectPr>/, ''))

  test('page setup is still written', async () => {
    const { engine } = await open(await noSect())
    await engine.call('docx.page', [{ w: 841.9, h: 595.3, cols: 2 }])
    const re = await open((await engine.call('save', [])) as Blob)
    expect(re.meta.page.w).toBeCloseTo(841.9, 1)
    expect(re.meta.page.cols).toBe(2)
  })

  test('a new header is still referenced', async () => {
    const { engine } = await open(await noSect())
    await engine.call('docx.hf', [{ parts: { header: [{ text: 'Hello' }] }, titlePg: true }])
    const re = await open((await engine.call('save', [])) as Blob)
    expect(lineOf(re.meta.header?.[0])).toBe('Hello')
    expect(re.meta.titlePg).toBe(true)
  })
})

describe('several sections', () => {
  test('page setup reaches every section; a new header is referenced from every section', async () => {
    // a section break inside the body (a paragraph carrying its own sectPr)
    const src = await variant('sample.docx', (x) => x.replace('<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Appendix</w:t>', '<w:p><w:pPr><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:pPr></w:p><w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Appendix</w:t>'))
    const { engine } = await open(src)
    await engine.call('docx.page', [{ ml: 36 }])
    await engine.call('docx.hf', [{ parts: { firstFooter: [{ text: 'x' }] } }])
    const xml = await mainXml((await engine.call('save', [])) as Blob)
    const sects = xml.match(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/g)!
    expect(sects.length).toBe(2)
    for (const s of sects) {
      expect(s).toContain('w:left="720"')
      expect(s).toMatch(/<w:footerReference w:type="first" r:id="rId\d+"\/>/)
    }
  })
})

describe('header / footer text', () => {
  test('markup characters, empty lines, both fields on one line', async () => {
    const { engine } = await open(new Blob([await Bun.file(S('sample.docx')).arrayBuffer()]))
    const res = (await engine.call('docx.hf', [{ parts: { footer: [{ text: 'A <b> & "q" {PAGE}/{NUMPAGES} {PAGES}' }, { text: '' }, { text: 'end' }] } }])) as HFSetResult
    expect(res.headers.footer?.map(lineOf)).toEqual(['A <b> & "q" {PAGE}/{NUMPAGES} {PAGES}', '', 'end'])
    const re = await open((await engine.call('save', [])) as Blob)
    expect(re.meta.footer?.map(lineOf)).toEqual(['A <b> & "q" {PAGE}/{NUMPAGES} {PAGES}', '', 'end'])
  })

  test('emptying a part leaves an empty paragraph (a valid part)', async () => {
    const { engine } = await open(new Blob([await Bun.file(S('sample.docx')).arrayBuffer()]))
    const res = (await engine.call('docx.hf', [{ parts: { header: [] } }])) as HFSetResult
    expect(res.headers.header?.map(lineOf)).toEqual([''])
  })
})

describe('tracked changes: files', () => {
  test('a file with tracked changes saved untouched is byte-identical', async () => {
    const src = new Blob([await Bun.file(S('tracked.docx')).arrayBuffer()])
    const { engine } = await open(src)
    expect(await mainXml((await engine.call('save', [])) as Blob)).toBe(await mainXml(src))
  })

  test('word count and text export see deleted text? the paragraph text keeps it (shown struck through)', async () => {
    const { blocks } = await open(new Blob([await Bun.file(S('tracked.docx')).arrayBuffer()]))
    const p = blocks.find((b): b is Paragraph => b.t === 'p' && lineOf(b).startsWith('Revenue'))!
    expect(lineOf(p)).toBe('Revenue grew 24%31% year over year')
  })

  test('a deletion with a comment range and a bookmark inside survives an edit of its paragraph', async () => {
    const src = await variant('tracked.docx', (x) => x.replace('<w:del w:id="101"', '<w:bookmarkStart w:id="7" w:name="bm"/><w:del w:id="101"').replace('<w:delText>24%</w:delText></w:r></w:del>', '<w:delText>24%</w:delText></w:r></w:del><w:bookmarkEnd w:id="7"/>'))
    const { engine, blocks } = await open(src)
    const { paraSpec } = await import('../src/core/docx/spec')
    const p = blocks.find((b): b is Paragraph => b.t === 'p' && lineOf(b).startsWith('Revenue'))!
    const runs = [...p.runs, { t: 'text' as const, text: '!', s: {} }]
    await engine.call('docx.edit', [{ anchor: p.pid, list: [paraSpec([{ ...p, runs, edited: true }])] }])
    const xml = await mainXml((await engine.call('save', [])) as Blob)
    expect(xml).toContain('<w:bookmarkStart w:id="7" w:name="bm"/>')
    expect(xml).toMatch(/<w:del [^>]*><w:r><w:delText xml:space="preserve">24%<\/w:delText><\/w:r><\/w:del>/)
  })
})

// ---------------------------------------------------------------- tables
import { tableSpec } from '../src/core/docx/spec'
import type { Table } from '../src/core/docx/types'
import { cellsIn, gridOf, mergeCells, newTable, rectOf, setBorders, unmergeCell } from '../src/react/docx/tables'

let pid = 7_000_000
const tp = (text = ''): Paragraph => ({ t: 'p', ps: {}, runs: text ? [{ t: 'text', text, s: {} }] : [], pid: pid++, edited: true })
const grid = (r: number, c: number): Table => {
  const t = newTable(r, c, 100 * c, () => tp())
  t.rows.forEach((row, ri) => row.cells.forEach((cell, ci) => (cell.blocks = [tp(`${ri}${ci}`)])))
  return t
}
const texts = (t: Table) => t.rows.map((r) => r.cells.map((c) => c.blocks.map(lineOf).join('+') + (c.colSpan > 1 ? `>${c.colSpan}` : '') + (c.rowSpan > 1 ? `v${c.rowSpan}` : '')))

describe('table edge cases', () => {
  test('a range crossing a cell merged down grows to its full height', () => {
    const t = mergeCells(grid(3, 3), { r0: 0, g0: 2, r1: 2, g1: 2 })! // last column merged down
    expect(rectOf(t, [1, 1], [1, 1])).toEqual({ r0: 1, g0: 1, r1: 1, g1: 1 })
    expect(rectOf(t, [1, 0], [0, 2])).toEqual({ r0: 0, g0: 0, r1: 2, g1: 2 })
  })

  test('merging the whole table, then unmerging it, gives the grid back (text in the first cell)', () => {
    const m = mergeCells(grid(2, 2), { r0: 0, g0: 0, r1: 1, g1: 1 })!
    expect(texts(m)).toEqual([['00+01+10+11>2v2'], []])
    const u = unmergeCell(m, 0, 0, () => tp())!
    expect(u.rows.map((r) => r.cells.length)).toEqual([2, 2])
  })

  test('a fully merged table survives save + reopen with all its rows', async () => {
    const { engine, blocks } = await open(new Blob([await Bun.file(S('sample.docx')).arrayBuffer()]))
    const orig = blocks.find((b): b is Table => b.t === 'tbl')!
    const n = orig.rows.length
    const g = gridOf(orig)
    const m = mergeCells({ ...orig, tid: 9_200_001 }, { r0: 0, g0: 0, r1: n - 1, g1: g.nCols - 1 })!
    await engine.call('docx.table', [{ tid: m.tid, item: m.item, spec: tableSpec(m) }])
    const re = await open((await engine.call('save', [])) as Blob)
    const t = re.blocks.find((b): b is Table => b.t === 'tbl')!
    expect(t.rows.length).toBe(n)
    expect(t.rows[0].cells[0].colSpan).toBe(g.nCols)
    expect(t.rows[0].cells[0].rowSpan).toBe(n)
  })

  test('a ragged row (fewer cells) does not break merging or borders', () => {
    const t = grid(2, 3)
    t.rows[1].cells.pop()
    expect(gridOf(t).at(1, 2)).toBeUndefined()
    const m = mergeCells(t, rectOf(t, [1, 0], [1, 1])!)!
    expect(texts(m)[1]).toEqual(['10+11>2'])
    const b = setBorders(t, { r0: 0, g0: 0, r1: 1, g1: 2 }, 'outer', '1pt solid #000000')
    expect(b.rows[0].cells[2].borders.r).toBe('1pt solid #000000')
  })

  test('a single cell is not "merged"; unmerging a plain cell does nothing', () => {
    const t = grid(2, 2)
    expect(mergeCells(t, { r0: 0, g0: 0, r1: 0, g1: 0 })).toBeNull()
    expect(unmergeCell(t, 1, 1, () => tp())).toBeNull()
    expect(cellsIn(t, { r0: 5, g0: 5, r1: 6, g1: 6 })).toEqual([])
  })

  test('top border of a cell below a merged-across cell reaches it', () => {
    const t = mergeCells(grid(2, 3), { r0: 0, g0: 0, r1: 0, g1: 2 })!
    const b = setBorders(t, { r0: 1, g0: 1, r1: 1, g1: 1 }, 'top', '2pt solid #ff0000')
    expect(b.rows[1].cells[1].borders.t).toBe('2pt solid #ff0000')
    expect(b.rows[0].cells[0].borders.b).toBe('2pt solid #ff0000')
  })
})

// ---------------------------------------------------------------- tracking
import type { Inline, TrackMark } from '../src/core/docx/types'
import { deleteRange, insertText, type RRun } from '../src/react/edit/rich'
import { resolveRuns, trackEdit } from '../src/react/docx/track'

const ME: TrackMark = { k: 'ins', author: 'Vy', date: '2026-11-08T10:00:00Z' }
const P = (runs: Inline[]): Paragraph => ({ t: 'p', ps: {}, runs, pid: 1 })
const T = (text: string, tc?: TrackMark): Inline => ({ t: 'text', text, s: {}, ...(tc ? { tc } : {}) })
const at = (o: number) => ({ a: { pa: '0', o }, f: { pa: '0', o } })
const show = (p: Paragraph) => p.runs.map((r) => (r.t !== 'text' ? '￼' : r.tc ? `[${r.tc.k === 'ins' ? '+' : '-'}${r.text}]` : r.text)).join('')

describe('tracking edge cases', () => {
  test('Backspace over an emoji keeps the whole pair', () => {
    const b = P([T('a😀b')])
    const n = P(deleteRange(b.runs as RRun[], 1, 3) as Inline[])
    expect(show(trackEdit([b], [n], '0', at(1), at(3), ME).next[0])).toBe('a[-😀]b')
  })

  test('typing at the very start and the very end', () => {
    const b = P([T('mid')])
    expect(show(trackEdit([b], [P(insertText(b.runs as RRun[], 0, '>') as Inline[])], '0', at(1), at(0), ME).next[0])).toBe('[+>]mid')
    expect(show(trackEdit([b], [P(insertText(b.runs as RRun[], 3, '<') as Inline[])], '0', at(4), at(3), ME).next[0])).toBe('mid[+<]')
  })

  test('typing the same letter that follows the caret is still put where it was typed', () => {
    const b = P([T('aab')])
    const n = P(insertText(b.runs as RRun[], 1, 'a') as Inline[])
    expect(show(trackEdit([b], [n], '0', at(2), at(1), ME).next[0])).toBe('a[+a]ab')
  })

  test("deleting across someone else's insertion and deletion", () => {
    const lan: TrackMark = { k: 'ins', author: 'Lan' }
    const b = P([T('x'), T('NEW', lan), T('old', { k: 'del', author: 'Lan' }), T('y')])
    const n = P(deleteRange(b.runs as RRun[], 0, 7) as Inline[])
    // "x" and Lan's insertion become my deletion (the reviewer still sees them), Lan's deletion stays hers
    const r = trackEdit([b], [n], '0', at(0), { a: { pa: '0', o: 0 }, f: { pa: '0', o: 7 } }, ME)
    expect(show(r.next[0])).toBe('[-xNEW][-old]y')
    expect(r.next[0].runs.filter((x) => x.t === 'text' && x.tc).map((x) => (x as { tc: TrackMark }).tc.author)).toEqual(['Vy', 'Lan'])
  })

  test('pictures removed while tracking go for real, the text around them is tracked', () => {
    const img: Inline = { t: 'img', target: 'x.png', w: 1, h: 1 }
    const b = P([T('a'), img, T('b')])
    const n = P(deleteRange(b.runs as RRun[], 0, 3) as Inline[])
    expect(show(trackEdit([b], [n], '0', at(0), { a: { pa: '0', o: 0 }, f: { pa: '0', o: 3 } }, ME).next[0])).toBe('[-ab]')
  })

  test('accepting part of a change leaves the rest of it (the two ends join again)', () => {
    const p = P([T('abcd', { k: 'del', author: 'Lan' })])
    expect(show(P(resolveRuns(p.runs, 1, 3, true)))).toBe('[-ad]')
    expect(show(P(resolveRuns(p.runs, 1, 3, false)))).toBe('[-a]bc[-d]')
  })
})

// ---------------------------------------------------------------- review findings (regressions)
import { docxToMarkdown, docxToText } from '../src/react/file/text'
import { paraSpec } from '../src/core/docx/spec'
import { DocumentClient } from '../src/core/worker/client'
import { DocxEditor, type DocState } from '../src/react/docx/DocxEditor'
import { opInsert } from '../src/react/edit/input'

async function editorOn(src: Blob) {
  const client = await DocumentClient.open(src, { worker: false })
  const st: DocState = { blocks: [], heights: new WeakMap(), dirty: new Set(), frontier: 0 }
  client.on('docx:blocks', (b) => st.blocks.push(...b))
  const done = new Promise<void>((r) => client.on('docx:done', () => r()))
  void client.call('docx.start')
  await done
  return { client, st, ed: new DocxEditor(client, st, () => {}) }
}
const tracked = async () => new Blob([await Bun.file(S('tracked.docx')).arrayBuffer()])

describe('review findings', () => {
  test('typing right after a deleted run (tracking off) is plain text', async () => {
    const { ed, st, client } = await editorOn(await tracked())
    const pi = st.blocks.findIndex((b) => b.t === 'p' && lineOf(b).startsWith('Revenue'))
    // after "24%" (deleted), before "31%" (inserted): offset 16
    opInsert(ed, { a: { pa: String(pi), o: 16 }, f: { pa: String(pi), o: 16 } }, 'X')
    expect(show(st.blocks[pi] as Paragraph)).toBe('Revenue grew [-24%]X[+31%] year over year')
    client.destroy()
  })

  test('a deleted picture stays deleted when its paragraph is edited', async () => {
    const src = await variant('sample.docx', (x) => x.replace(/<w:r>(<w:drawing>[\s\S]*?<\/w:drawing>)<\/w:r>/, '<w:del w:id="301" w:author="Lan"><w:r>$1</w:r></w:del>'))
    const { engine, blocks } = await open(src)
    const p = blocks.find((b): b is Paragraph => b.t === 'p' && b.runs.some((r) => r.t === 'img'))!
    expect(p.runs.find((r) => r.t === 'img')?.tc?.k).toBe('del')
    await engine.call('docx.edit', [{ anchor: p.pid, list: [paraSpec([{ ...p, runs: [...p.runs, { t: 'text', text: ' caption', s: {} }], edited: true }])] }])
    const xml = await mainXml((await engine.call('save', [])) as Blob)
    expect(xml).toMatch(/<w:del [^>]*w:author="Lan"[^>]*><w:r><w:drawing>/)
  })

  test('text export and markdown leave deleted text out', async () => {
    const { blocks } = await open(await tracked())
    expect(docxToText(blocks)).toContain('Revenue grew 31% year over year')
    expect(docxToMarkdown(blocks)).toContain('Expand to Vietnam')
    expect(docxToText(blocks)).not.toContain('24%')
  })

  test("tracking: deleting someone else's insertion leaves a deletion", () => {
    const b = P([T('ab'), T('XYZ', { k: 'ins', author: 'Ann' })])
    const n = P(deleteRange(b.runs as RRun[], 2, 5) as Inline[])
    const r = trackEdit([b], [n], '0', at(2), { a: { pa: '0', o: 2 }, f: { pa: '0', o: 5 } }, { ...ME, author: 'Bob' })
    expect(show(r.next[0])).toBe('ab[-XYZ]')
  })

  test('undoing a page / first-page change leaves every section as it was', async () => {
    const src = await variant('sample.docx', (x) => x.replace('<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Appendix</w:t>', '<w:p><w:pPr><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="2880" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:pPr></w:p><w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Appendix</w:t>'))
    const { engine } = await open(src)
    const before = await mainXml(src)
    const p = (await engine.call('docx.page', [{ ml: 36 }])) as { undo: object }
    await engine.call('docx.page', [p.undo])
    const h = (await engine.call('docx.hf', [{ titlePg: true }])) as HFSetResult
    await engine.call('docx.hf', [h.undo])
    expect(await mainXml((await engine.call('save', [])) as Blob)).toBe(before)
  })

  test('change ids stay unique across save → reopen → edit → save', async () => {
    const { ed, st, client } = await editorOn(await tracked())
    ed.author = () => 'Vy'
    ed.setTracking(true)
    const pa = (s: DocState, t: string) => String(s.blocks.findIndex((b) => b.t === 'p' && lineOf(b).startsWith(t)))
    opInsert(ed, { a: { pa: pa(st, 'Hire'), o: 0 }, f: { pa: pa(st, 'Hire'), o: 0 } }, 'A ')
    await new Promise((r) => setTimeout(r, 30))
    const first = (await client.call('save')) as Blob
    const two = await editorOn(first)
    two.ed.author = () => 'Vy'
    two.ed.setTracking(true)
    opInsert(two.ed, { a: { pa: pa(two.st, 'Ship'), o: 0 }, f: { pa: pa(two.st, 'Ship'), o: 0 } }, 'B ')
    await new Promise((r) => setTimeout(r, 30))
    const xml = await mainXml((await two.client.call('save')) as Blob)
    const ids = [...xml.matchAll(/<w:(?:ins|del) w:id="(\d+)"/g)].map((m) => m[1])
    expect(ids.length).toBeGreaterThanOrEqual(7)
    expect(new Set(ids).size).toBe(ids.length)
    client.destroy()
    two.client.destroy()
  })

  test('unmerged cells and the continuation cells of a vertical merge carry the merged cell’s look', async () => {
    const { engine, blocks } = await open(new Blob([await Bun.file(S('sample.docx')).arrayBuffer()]))
    const orig = blocks.find((b): b is Table => b.t === 'tbl')!
    // body cell (1,0) merged down over (2,0), filled and outlined; header row stays as it is
    let t: Table = mergeCells({ ...orig, tid: 9_300_001 }, { r0: 1, g0: 0, r1: 2, g1: 0 })!
    t = { ...t, rows: t.rows.map((r, ri) => (ri !== 1 ? r : { ...r, cells: r.cells.map((c, ci) => (ci ? c : { ...c, shd: '#ffcc00', set: { shd: true } })) })) }
    t = setBorders(t, { r0: 1, g0: 0, r1: 2, g1: 0 }, 'outer', '3pt solid #ff0000')
    await engine.call('docx.table', [{ tid: t.tid, item: t.item, spec: tableSpec(t) }])
    const xml = await mainXml((await engine.call('save', [])) as Blob)
    const cont = /<w:tc><w:tcPr>(?:(?!<\/w:tcPr>)[\s\S])*<w:vMerge\/>(?:(?!<\/w:tcPr>)[\s\S])*<\/w:tcPr>/.exec(xml)![0]
    expect(cont).toContain('w:fill="FFCC00"')
    expect(cont).toContain('<w:bottom w:val="single" w:sz="24"')
    // unmerge: the new cell takes the merged cell's fill, not the header's
    const re = await open((await engine.call('save', [])) as Blob)
    const rt = { ...re.blocks.find((b): b is Table => b.t === 'tbl')!, tid: 9_300_002 }
    const u = unmergeCell(rt, 1, 0, () => tp())!
    await re.engine.call('docx.table', [{ tid: u.tid, item: u.item, spec: tableSpec(u) }])
    const re2 = await open((await re.engine.call('save', [])) as Blob)
    expect(re2.blocks.find((b): b is Table => b.t === 'tbl')!.rows[2].cells[0].shd).toBe('#ffcc00')
  })

  test('a header part with a self-closing root is rewritten as valid XML', async () => {
    const zip = await ZipArchive.open(new BlobSource(Bun.file(S('sample.docx')) as unknown as Blob))
    const rels = await zip.text('word/_rels/document.xml.rels')
    const hdr = /Id="(rId\d+)"[^>]*Target="(header\d*\.xml)"|Target="(header\d*\.xml)"[^>]*Id="(rId\d+)"/.exec(rels)!
    const part = `word/${hdr[2] ?? hdr[3]}`
    const src = await rebuildZip(zip, new Map([[part, '<?xml version="1.0" encoding="UTF-8"?>\n<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>']]), DOCX)
    const { engine } = await open(src)
    await engine.call('docx.hf', [{ parts: { header: [{ text: 'ok' }] } }])
    const re = await open((await engine.call('save', [])) as Blob)
    expect(lineOf(re.meta.header?.[0])).toBe('ok')
  })

  test('accept all is one undo step', async () => {
    const { ed, client } = await editorOn(await tracked())
    expect(ed.review(true, true)).toBe(5)
    expect(ed.changeCount()).toBe(0)
    ed.undo()
    expect(ed.changeCount()).toBe(5)
    expect(ed.canUndo).toBe(false)
    ed.redo()
    expect(ed.changeCount()).toBe(0)
    client.destroy()
  })
})

describe('saving in between', () => {
  test('page setup and first-page changes undone after a save are gone from the next save', async () => {
    const src = new Blob([await Bun.file(S('sample.docx')).arrayBuffer()])
    const { engine } = await open(src)
    const before = await mainXml(src)
    const p = (await engine.call('docx.page', [{ w: 841.9, h: 595.3, cols: 2 }])) as { undo: object }
    const h = (await engine.call('docx.hf', [{ titlePg: true }])) as HFSetResult
    expect(await mainXml((await engine.call('save', [])) as Blob)).not.toBe(before)
    await engine.call('docx.page', [p.undo])
    await engine.call('docx.hf', [h.undo])
    expect(await mainXml((await engine.call('save', [])) as Blob)).toBe(before)
  })
})
