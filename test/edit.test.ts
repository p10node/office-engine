import { describe, expect, test } from 'bun:test'
import { Engine } from '../src/core/worker/engine'
import { paraSpec } from '../src/core/docx/spec'
import type { Block, EditResult, Paragraph, StyleSetResult } from '../src/core/docx/types'
import { ZipArchive } from '../src/core/zip'
import { BlobSource } from '../src/core/source'
import { rebuildZip } from '../src/core/zipwrite'

const S = (n: string) => `${import.meta.dir}/../playground/public/samples/${n}`
const file = (n: string) => Bun.file(S(n)) as unknown as Blob

async function openDocx(src: Blob) {
  const events: { event: string; data: any }[] = []
  const engine = new Engine((event, data) => events.push({ event, data }))
  await engine.open(src, {})
  await engine.call('docx.start', [])
  const blocks: Block[] = events.filter((e) => e.event === 'docx:blocks').flatMap((e) => e.data)
  return { engine, blocks, events }
}

const textOf = (p: Paragraph) => p.runs.map((r) => (r.t === 'text' ? r.text : '￼')).join('')
const paras = (blocks: Block[]) => blocks.filter((b): b is Paragraph => b.t === 'p')

async function partText(blob: Blob, path: string) {
  const zip = await ZipArchive.open(new BlobSource(blob))
  return zip.text(path)
}

describe('docx write-back', () => {
  test('unedited save reproduces document.xml exactly', async () => {
    const { engine } = await openDocx(file('sample.docx'))
    const orig = await partText(file('sample.docx'), 'word/document.xml')
    // force regeneration of document.xml through the edit path with a pristine spec
    const out = (await engine.call('save', [])) as Blob
    expect(await partText(out, 'word/document.xml')).toBe(orig)
  })

  test('edit text + bold + new paragraph + list survives save/reopen', async () => {
    const { engine, blocks } = await openDocx(file('sample.docx'))
    const ps = paras(blocks)
    const target = ps.find((p) => textOf(p).startsWith('This document shows'))!
    // replace the first text run and bold it, insert a new paragraph after
    const edited: Paragraph = {
      ...target,
      edited: true,
      runs: [{ t: 'text', text: 'Hello, café world! ', s: {}, src: (target.runs[0] as any).src, d: { b: true } }, ...target.runs.slice(1)],
    }
    const extra: Paragraph = {
      t: 'p', ps: {}, pid: 1_000_000, from: target.pid, edited: true, pd: { list: 'bullet' },
      runs: [{ t: 'text', text: 'New line — ✓', s: {}, src: (target.runs[0] as any).src, d: { i: true, color: '#c00000' } }],
    }
    const res = (await engine.call('docx.edit', [{ anchor: target.pid, list: [paraSpec([edited]), paraSpec([extra])], want: true }])) as EditResult
    const nb = res.blocks[1_000_000][0] as Paragraph
    expect(nb.marker?.text).toBe('•')
    expect(nb.lst).toBe('bullet')

    // delete the paragraph after the image caption
    const cap = ps.find((p) => textOf(p).startsWith('Figure 1'))!
    await engine.call('docx.edit', [{ anchor: cap.pid, list: [], remove: [] }])

    const saved = (await engine.call('save', [])) as Blob
    const xml = await partText(saved, 'word/document.xml')
    expect(xml).toContain('Hello, café world!')
    expect(xml).not.toContain('Figure 1 — generated gradient image')

    const re = await openDocx(saved)
    const rps = paras(re.blocks)
    const p1 = rps.find((p) => textOf(p).startsWith('Hello, café'))!
    expect(p1).toBeDefined()
    const r0 = p1.runs[0] as any
    expect(r0.s.b).toBe(true)
    expect(textOf(p1)).toContain('example.com')
    expect((p1.runs.find((r: any) => r.href) as any)?.href).toBe('https://example.com')
    const p2 = rps[rps.indexOf(p1) + 1]
    expect(textOf(p2)).toBe('New line — ✓')
    expect((p2.runs[0] as any).s.i).toBe(true)
    expect((p2.runs[0] as any).s.color).toBe('#c00000')
    expect(p2.marker?.text).toBe('•')
    // image still there, numbering still renumbered correctly
    expect(rps.some((p) => p.runs.some((r) => r.t === 'img'))).toBe(true)
    expect(rps.filter((p) => /^\d\.$/.test(p.marker?.text ?? '')).map((p) => p.marker!.text)).toEqual(['1.', '2.', '3.'])
  })

  test('heading style change + numbered list renumbering', async () => {
    const { engine, blocks } = await openDocx(file('sample.docx'))
    const ps = paras(blocks)
    const nums = ps.filter((p) => /^\d\.$/.test(p.marker?.text ?? ''))
    // remove the first numbered item from the list → others renumber
    const res = (await engine.call('docx.edit', [{ anchor: nums[0].pid, list: [{ ...paraSpec([{ ...nums[0], edited: true }]), pd: { list: 'none' } }], want: true }])) as EditResult
    expect(res.markers[nums[1].pid!]?.text).toBe('1.')
    expect(res.markers[nums[2].pid!]?.text).toBe('2.')
    const p = ps.find((x) => textOf(x) === 'Revenue grew 24% year over year')!
    const r2 = (await engine.call('docx.edit', [{ anchor: p.pid, list: [{ ...paraSpec([{ ...p, edited: true }]), pd: { style: 'Heading1', list: 'none' } }], want: true }])) as EditResult
    const hp = r2.blocks[p.pid!][0] as Paragraph
    expect(hp.sid).toBe('Heading1')
    expect((hp.runs[0] as any).s.sz).toBeGreaterThan(12)
    expect(hp.marker).toBeUndefined()
  })
})

async function openXlsx(src: Blob, sheet = 0) {
  const events: { event: string; data: any }[] = []
  const engine = new Engine((event, data) => events.push({ event, data }))
  const info = await engine.open(src, {})
  await engine.call('xlsx.load', [sheet])
  while (!events.some((e) => e.event === 'xlsx:progress' && e.data.sheet === sheet && e.data.done)) await new Promise((r) => setTimeout(r, 10))
  const cell = async (r: number, c: number) => {
    const b: any = await engine.call('xlsx.block', [sheet, r, r + 1])
    return b.rows[0]?.cells.find((x: any) => x[0] === c)
  }
  return { engine, info, cell, events }
}

describe('xlsx write-back', () => {
  test('values, formulas, styles, undo survive save/reopen', async () => {
    const mem = new Blob([await Bun.file(S('sample.xlsx')).arrayBuffer()])
    const { engine, cell } = await openXlsx(mem)
    await engine.call('xlsx.set', [0, [
      { r: 2, c: 0, input: 'Zürich' },
      { r: 2, c: 2, input: '1234.5' },
      { r: 70, c: 1, input: 'new row' },
      { r: 3, c: 9, input: '=C4*2' },
      { r: 4, c: 0, input: '' },
    ]])
    expect((await cell(2, 0))[1]).toBe('Zürich')
    const fmt: any = await engine.call('xlsx.format', [0, { r0: 2, c0: 0, r1: 2, c1: 1 }, { b: true, fill: '#ffff00', align: 'center' }])
    expect(fmt.styles.length).toBeGreaterThan(0)
    expect(fmt.styles[0].css.fontWeight).toBe(700)
    await engine.call('xlsx.set', [0, [{ r: 5, c: 0, input: 'temp' }]])
    await engine.call('xlsx.undo', [])
    expect((await cell(5, 0))[1]).not.toBe('temp')

    const saved = (await engine.call('save', [])) as Blob
    const re = await openXlsx(saved)
    expect((await re.cell(2, 0))[1]).toBe('Zürich')
    expect((await re.cell(2, 2))[1]).toBe('1234.5')
    expect((await re.cell(70, 1))[1]).toBe('new row')
    expect((await re.cell(4, 0))?.[1] ?? '').toBe('')
    const styled = await re.cell(2, 0)
    const css = re.info.xlsx!.styles[styled[2]]
    expect(css.fontWeight).toBe(700)
    expect(css.textAlign).toBe('center')
    const f: any = await re.engine.call('xlsx.cell', [0, 3, 9])
    expect(f.formula).toBe('=C4*2')
    // untouched data intact
    expect((await re.cell(1, 0))[1]).toBe('Region')
    expect((await re.cell(0, 0))[1]).toBe('Sales by region — FY2026')
  })
})

describe('pptx write-back', () => {
  test('shape text edit survives save/reopen', async () => {
    const mem = new Blob([await Bun.file(S('sample.pptx')).arrayBuffer()])
    const engine = new Engine(() => {})
    await engine.open(mem, {})
    const s2: any = await engine.call('pptx.slide', [2])
    const shape = s2.els.find((e: any) => e.text?.paras[0].runs[0]?.text === 'Rounded')
    expect(shape.sid).toBeGreaterThan(0)
    const r0 = shape.text.paras[0].runs[0]
    const res: any = await engine.call('pptx.setText', [2, shape.sid, [
      { k: 0, algn: 'center', runs: [{ text: 'Rounded corner', src: r0.src, set: { b: true, color: '#ffff00' } }] },
      { k: 0, runs: [{ text: 'line 2\nline 3', src: r0.src, set: { sz: 12 } }] },
    ]])
    const again = res.els.find((e: any) => e.sid === shape.sid)
    expect(again.text.paras.length).toBe(2)
    expect(again.text.paras[0].runs[0].src).toEqual(r0.src)
    // second edit on top of the first keeps referring to the original XML
    await engine.call('pptx.setText', [2, shape.sid, [{ k: 0, algn: 'center', runs: [{ text: 'Rounded corner!', src: again.text.paras[0].runs[0].src, set: { b: true, color: '#ffff00' } }] }]])
    const saved = (await engine.call('save', [])) as Blob
    const re = new Engine(() => {})
    await re.open(saved, {})
    const s: any = await re.call('pptx.slide', [2])
    const sh = s.els.find((e: any) => e.sid === shape.sid)
    expect(sh.text.paras[0].runs[0].text).toBe('Rounded corner!')
    expect(sh.text.paras[0].runs[0].b).toBe(true)
    expect(sh.text.paras[0].runs[0].color).toBe('#ffff00')
    expect(sh.text.paras[0].algn).toBe('center')
    // other slides untouched
    const s0: any = await re.call('pptx.slide', [0])
    expect(s0.els.some((e: any) => e.text?.paras[0]?.runs[0]?.text === 'office-engine')).toBe(true)
  })
})

describe('docx full editing', () => {
  test('table regen + new table + image + link + spacing', async () => {
    const { engine, blocks } = await openDocx(new Blob([await Bun.file(S('sample.docx')).arrayBuffer()]))
    const ps = paras(blocks)
    const tbl = blocks.find((b) => b.t === 'tbl') as any
    expect(tbl.item).toBeGreaterThan(0)
    // add a row to the existing table (copy of row 1) and change a cell
    const cellPara = (c: any) => c.blocks[0] as Paragraph
    const specOf = (c: any, text?: string) => {
      const p = cellPara(c)
      return { src: c.src, colSpan: c.colSpan, rowSpan: c.rowSpan, paras: [paraSpec([text === undefined ? p : { ...p, edited: true, runs: [{ ...(p.runs[0] as any), text }] }])] }
    }
    const rows = tbl.rows.map((r: any, ri: number) => ({ src: r.src, cells: r.cells.map((c: any, ci: number) => specOf(c, ri === 1 && ci === 0 ? 'North' : undefined)) }))
    rows.splice(2, 0, { src: tbl.rows[1].src, cells: tbl.rows[1].cells.map((c: any) => ({ src: c.src, colSpan: c.colSpan, rowSpan: 1, paras: [{ pid: 2_000_000 + Math.random() * 1000 | 0, from: cellPara(c).pid, parts: [[{ text: 'new' }]] }] })) })
    await engine.call('docx.table', [{ tid: 9_000_001, item: tbl.item, spec: { tid: 9_000_001, item: tbl.item, cols: tbl.cols, rows } }])

    // insert a brand new 2x2 table after the first paragraph
    const p0 = ps.find((p) => textOf(p).startsWith('This document shows'))!
    const mk = (t: string, i: number) => ({ colSpan: 1, rowSpan: 1, paras: [{ pid: 3_000_000 + i, from: p0.pid!, pd: { style: '', list: 'none' as const }, keepSect: false, orig: [], parts: [[{ text: t }]] }] })
    await engine.call('docx.table', [{ tid: 9_000_002, anchor: p0.pid, spec: { tid: 9_000_002, cols: [200, 200], rows: [{ cells: [mk('a1', 1), mk('b1', 2)] }, { cells: [mk('a2', 3), mk('b2', 4)] }] } }])

    // image + link + spacing on paragraph 0
    const png = Uint8Array.from(atob('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=='), (c) => c.charCodeAt(0))
    const img: any = await engine.call('docx.image', [png.buffer, 'png'])
    const lk: any = await engine.call('docx.link', ['https://eldora.vn'])
    const p = { ...p0, edited: true, pd: { line: 1.5, before: 12 }, runs: [...p0.runs, { t: 'text', text: ' link', s: {}, lk, href: 'https://eldora.vn' }, { t: 'img', target: img.path, w: 50, h: 50, rid: img.rid }] } as Paragraph
    await engine.call('docx.edit', [{ anchor: p0.pid, list: [paraSpec([p])] }])

    const saved = (await engine.call('save', [])) as Blob
    const re = await openDocx(saved)
    const tables = re.blocks.filter((b) => b.t === 'tbl') as any[]
    expect(tables.length).toBe(2)
    const newT = tables[0]
    expect(newT.rows.length).toBe(2)
    expect(textOf(newT.rows[1].cells[1].blocks[0])).toBe('b2')
    const oldT = tables[1]
    expect(oldT.rows.length).toBe(6)
    expect(textOf(oldT.rows[1].cells[0].blocks[0])).toBe('North')
    expect(textOf(oldT.rows[2].cells[0].blocks[0])).toBe('new')
    expect(oldT.rows[0].cells[0].shd).toBe('#d9e2f3')
    expect(oldT.rows[5].cells[0].colSpan).toBe(3)
    const rp = paras(re.blocks).find((x) => textOf(x).startsWith('This document shows'))!
    expect(rp.ps.line).toBe(1.5)
    expect(rp.ps.before).toBe(12)
    expect(rp.runs.some((r) => r.t === 'img' && r.target.startsWith('word/media/image'))).toBe(true)
    expect((rp.runs.find((r: any) => r.href === 'https://eldora.vn') as any)?.text).toContain('link')
  })
})

describe('docx styles + pictures', () => {
  test('missing built-in styles are created; alt text and replaced pictures are written', async () => {
    const zip = await ZipArchive.open(new BlobSource(new Blob([await Bun.file(S('sample.docx')).arrayBuffer()])))
    const styles = (await zip.text('word/styles.xml')).replace(/<w:style [^>]*w:styleId="(Heading\d|Title|Subtitle)"[^>]*>[\s\S]*?<\/w:style>/g, '')
    const stripped = await rebuildZip(zip, new Map([['word/styles.xml', styles]]))
    const { engine, blocks, events } = await openDocx(stripped)
    const meta = events.filter((e) => e.event === 'docx:meta').pop()!.data
    expect(meta.styles.some((x: any) => x.canon === 'h1')).toBe(false)
    const r: any = await engine.call('docx.style', ['h1'])
    expect(r.id).toBe('Heading1')
    expect(r.styles.find((x: any) => x.canon === 'h1')?.outline).toBe(0)
    // asking again reuses it
    expect(((await engine.call('docx.style', ['h1'])) as any).id).toBe('Heading1')
    const p0 = paras(blocks).find((p) => textOf(p).startsWith('This document shows'))!
    const res = (await engine.call('docx.edit', [{ anchor: p0.pid, list: [{ ...paraSpec([{ ...p0, edited: true }]), pd: { style: 'Heading1' } }], want: true }])) as EditResult
    const h = res.blocks[p0.pid!][0] as Paragraph
    expect(h.sid).toBe('Heading1')
    expect(h.ps.outline).toBe(0)
    expect(h.runs.find((x) => x.t === 'text')!.s.sz).toBe(16)

    // picture: alt text + replacement keep the original drawing
    const picPara = paras(blocks).find((p) => p.runs.some((x) => x.t === 'img'))!
    const png = Uint8Array.from(atob('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=='), (c) => c.charCodeAt(0))
    const img: any = await engine.call('docx.image', [png.buffer, 'png'])
    const runs = picPara.runs.map((x) => (x.t === 'img' ? { ...x, alt: 'Doanh thu "Q1"', altSet: true, rid: img.rid, target: img.path, repl: true, w: 100, h: 50, resized: true } : x))
    await engine.call('docx.edit', [{ anchor: picPara.pid, list: [paraSpec([{ ...picPara, edited: true, runs }])] }])
    const saved = (await engine.call('save', [])) as Blob
    const doc = await partText(saved, 'word/document.xml')
    expect(doc).toContain('descr="Doanh thu &quot;Q1&quot;"')
    expect(doc).toContain(`r:embed="${img.rid}"`)
    expect(doc).toContain('cx="1270000"')
    expect(await partText(saved, 'word/styles.xml')).toContain('w:styleId="Heading1"')
    const re = await openDocx(saved)
    const pic = paras(re.blocks).flatMap((p) => p.runs).find((x) => x.t === 'img') as any
    expect(pic.alt).toBe('Doanh thu "Q1"')
    expect(pic.target).toBe(img.path)
  })
})

describe('docx style looks', () => {
  const src = async () => new Blob([await Bun.file(S('sample.docx')).arrayBuffer()])
  const first = (bl: Block[]) => bl[0] as Paragraph
  const look = (p: Paragraph) => (p.runs.find((x) => x.t === 'text') as Extract<Paragraph['runs'][number], { t: 'text' }>).s

  test('a style takes another look in the document itself: its paragraphs follow, the file keeps it, the way back is exact', async () => {
    const file = await src()
    const { engine, blocks } = await openDocx(file)
    const heads = paras(blocks).filter((p) => p.sid === 'Heading1')
    const body = paras(blocks).find((p) => textOf(p).startsWith('This document shows'))!
    expect(heads.length).toBeGreaterThan(1)
    const r = (await engine.call('docx.styleSet', ['Heading1', { color: '#c00000', sz: 20, b: false, font: 'Georgia', align: 'center', before: 6 }])) as StyleSetResult
    // the style and those based on it, nothing else
    expect(r.changed).toContain('Heading1')
    expect(r.changed).not.toContain('Normal')
    expect(r.changed).not.toContain('Heading2')
    expect(r.undo).toEqual({ color: '#365f91', sz: 14, b: true, font: 'Calibri', align: null, before: 24 })
    expect(r.styles.find((x) => x.canon === 'h1')!.preview).toMatchObject({ color: '#c00000', sz: 20 })
    expect(r.headers?.header?.length).toBeGreaterThan(0)

    const conv = (await engine.call('docx.convert', [[...heads, body].map((p) => paraSpec([p]))])) as Block[][]
    for (const bl of conv.slice(0, heads.length)) {
      expect(look(first(bl))).toMatchObject({ color: '#c00000', sz: 20, font: '"Georgia", serif' })
      expect(look(first(bl)).b).toBeFalsy()
      expect(first(bl).ps).toMatchObject({ align: 'center', before: 6 })
    }
    // a paragraph of another style keeps its look
    expect(first(conv.at(-1)!).ps.align).toBeUndefined()
    expect(look(first(conv.at(-1)!))).toEqual(look(body))

    // the look it has, or values no look can hold: nothing changes
    expect(((await engine.call('docx.styleSet', ['Heading1', { color: '#c00000', sz: 20, align: 'center' }])) as StyleSetResult).changed).toEqual([])
    expect(((await engine.call('docx.styleSet', ['Heading1', { color: 'red"/><w:x', sz: 'big', align: 'up' }])) as StyleSetResult).changed).toEqual([])
    expect(((await engine.call('docx.styleSet', ['NoSuchStyle', { sz: 30 }])) as StyleSetResult).changed).toEqual([])

    // only the styles part of this file is written
    const saved = (await engine.call('save', [])) as Blob
    expect(await partText(saved, 'word/document.xml')).toBe(await partText(file, 'word/document.xml'))
    const def = /<w:style [^>]*w:styleId="Heading1">.*?<\/w:style>/.exec(await partText(saved, 'word/styles.xml'))![0]
    expect(def).toContain('<w:rFonts w:ascii="Georgia" w:hAnsi="Georgia" w:eastAsia="Georgia" w:cs="Georgia"/><w:color w:val="C00000"/><w:sz w:val="40"/><w:szCs w:val="40"/>')
    expect(def).toContain('<w:spacing w:before="120" w:after="0"/><w:jc w:val="center"/>')
    // bold is what the style it is based on gives: not repeated in it
    expect(def).not.toMatch(/<w:b(Cs)?[ \/]/)
    const re = await openDocx(saved)
    for (const p of paras(re.blocks).filter((x) => x.sid === 'Heading1')) expect(look(p)).toMatchObject({ color: '#c00000', sz: 20 })

    // back: theme font and theme colour included, as the file came
    expect(((await engine.call('docx.styleSet', ['Heading1', r.undo])) as StyleSetResult).undo).toEqual({ color: '#c00000', sz: 20, b: null, font: 'Georgia', align: 'center', before: 6 })
    expect(await partText((await engine.call('save', [])) as Blob, 'word/styles.xml')).toBe(await partText(file, 'word/styles.xml'))
  })

  test('the default style: every style based on it follows, list markers and headers included', async () => {
    const { engine, blocks, events } = await openDocx(await src())
    const meta = events.filter((e) => e.event === 'docx:meta').pop()!.data
    const item = paras(blocks).find((p) => p.lst === 'number')!
    expect(item.marker!.s.sz).toBe(11)
    const r = (await engine.call('docx.styleSet', ['Normal', { sz: 13, color: '#1f4e79' }])) as StyleSetResult
    expect(r.changed).toEqual(expect.arrayContaining(['Normal', 'Heading1', 'ListNumber', 'Header']))
    const [bl] = (await engine.call('docx.convert', [[paraSpec([item])]])) as Block[][]
    // the number keeps its text and takes the size of the paragraph
    expect(first(bl).marker).toMatchObject({ text: item.marker!.text, s: { sz: 13, color: '#1f4e79' } })
    expect(look(first(bl)).sz).toBe(13)
    const header = (b: Block[] | undefined) => look(b!.find((x): x is Paragraph => x.t === 'p' && x.runs.length > 0)!)
    expect(header(meta.header).sz).toBe(11)
    expect(header(r.headers!.header)).toMatchObject({ sz: 13, color: '#1f4e79' })
    // a heading sets its own size and colour: only what it inherits moved
    const head = paras(blocks).find((p) => p.sid === 'Heading1')!
    const [hb] = (await engine.call('docx.convert', [[paraSpec([head])]])) as Block[][]
    expect(look(first(hb))).toEqual(look(head))
  })

  test('formatting a run or a paragraph sets itself can be dropped: the style decides again', async () => {
    const { engine, blocks } = await openDocx(await src())
    const p = paras(blocks).find((x) => textOf(x).startsWith('This document shows'))!
    const bold = p.runs.findIndex((x) => x.t === 'text' && x.s.b)
    const big = p.runs.findIndex((x) => x.t === 'text' && (x.s.sz ?? 0) > 11)
    expect(bold).toBeGreaterThan(-1)
    expect(big).toBeGreaterThan(-1)
    const runs = p.runs.map((x, i) => (i === bold ? { ...x, d: { b: null } } : i === big ? { ...x, d: { sz: null, font: null } } : x)) as Paragraph['runs']
    const centred = { ...paraSpec([{ ...p, edited: true }]), pd: { align: 'center' as const, before: 9, line: 2 } }
    await engine.call('docx.edit', [{ anchor: p.pid, list: [centred] }])
    const res = (await engine.call('docx.edit', [{ anchor: p.pid, list: [{ ...paraSpec([{ ...p, edited: true, runs }]), pd: { align: null, before: null, line: null } }], want: true }])) as EditResult
    const q = res.blocks[p.pid!][0] as Paragraph
    expect(textOf(q)).toBe(textOf(p))
    expect(q.runs.every((x) => x.t !== 'text' || (!x.s.b && x.s.sz === 11))).toBe(true)
    expect(q.ps).toEqual(p.ps)
    const xml = await partText((await engine.call('save', [])) as Blob, 'word/document.xml')
    const out = /<w:p[ >](?:(?!<\/w:p>)[\s\S])*?This document shows[\s\S]*?<\/w:p>/.exec(xml)![0]
    expect(out).not.toContain('<w:b/>')
    expect(out).not.toContain('<w:jc')
  })
})

describe('pptx toolbar ops', () => {
  test('arrange, duplicate, rotate/flip, dash, anchor, background, undo; list/level/spacing/highlight text', async () => {
    const engine = new Engine(() => {})
    await engine.open(new Blob([await Bun.file(S('sample.pptx')).arrayBuffer()]), {})
    const s2: any = await engine.call('pptx.slide', [2])
    const tops = () => [...new Set<number>(cur.els.map((e: any) => e.top))]
    let cur = s2
    const rounded = s2.els.find((e: any) => e.text?.paras[0]?.runs[0]?.text === 'Rounded')
    const order0 = tops()
    // one step forward swaps with the next shape
    cur = await engine.call('pptx.shapes', [2, [{ op: 'order', top: rounded.top, to: 'forward' }]])
    const order1 = tops()
    const i0 = order0.indexOf(rounded.top)
    expect(order1.indexOf(rounded.top)).toBe(i0 + 1)
    cur = await engine.call('pptx.shapes', [2, [{ op: 'duplicate', top: rounded.top }]])
    const copies = cur.els.filter((e: any) => e.text?.paras[0]?.runs[0]?.text === 'Rounded')
    expect(copies.length).toBe(2)
    expect(Math.round(copies[1].box.x - copies[0].box.x)).toBe(16)
    cur = await engine.call('pptx.shapes', [2, [
      { op: 'xfrm', top: rounded.top, box: { ...rounded.box, rot: 90, flipH: true } },
      { op: 'line', top: rounded.top, w: 3, dash: 'dash' },
      { op: 'anchor', sid: rounded.sid, anchor: 'b' },
      { op: 'bg', color: '#fff2cc' },
    ]])
    const r = cur.els.find((e: any) => e.top === rounded.top)
    expect(r.box.rot).toBe(90)
    expect(r.box.flipH).toBe(true)
    expect(r.stroke.w).toBeCloseTo(4, 0)
    expect(r.text.anchor).toBe('b')
    expect(cur.bg?.c ?? cur.bg?.css).toContain('fff2cc')
    // undo the last batch, redo it
    const u: any = await engine.call('pptx.undo', [])
    expect(u.index).toBe(2)
    expect(u.model.els.find((e: any) => e.top === rounded.top).box.rot).toBeFalsy()
    const re: any = await engine.call('pptx.undo', [true])
    expect(re.model.els.find((e: any) => e.top === rounded.top).box.rot).toBe(90)
    // text: numbered list, level 1, spacing, highlight + superscript
    const s1: any = await engine.call('pptx.slide', [1])
    const body = s1.els.find((e: any) => e.text?.paras.length > 2)
    const specs = body.text.paras.map((p: any, k: number) => ({
      k,
      bullet: 'number',
      lvl: k === 0 ? 1 : undefined,
      lnPct: 1.5,
      runs: p.runs.map((x: any, j: number) => ({ text: x.text, src: x.src, set: j === 0 ? { hl: '#ffff00', baseline: 30 } : undefined })),
    }))
    const m1: any = await engine.call('pptx.setText', [1, body.sid, specs])
    const t = m1.els.find((e: any) => e.sid === body.sid).text
    expect(t.paras[0].lvl).toBe(1)
    expect(t.paras[0].buKind).toBe('number')
    expect(t.paras[1].bullet.text).toMatch(/^\d+\.$/)
    expect(t.paras[0].lnPct).toBeCloseTo(1.5)
    expect(t.paras[0].runs[0].hl).toBe('#ffff00')
    expect(t.paras[0].runs[0].baseline).toBe(30)
    const saved = (await engine.call('save', [])) as Blob
    const x2 = await partText(saved, 'ppt/slides/slide2.xml')
    expect(x2).toContain('<a:buAutoNum type="arabicPeriod"/>')
    expect(x2).toContain('<a:spcPct val="150000"/>')
    expect(x2).toContain('baseline="30000"')
  })
})

describe('pptx full editing', () => {
  test('shapes + slides survive save/reopen', async () => {
    const engine = new Engine(() => {})
    const info: any = await engine.open(new Blob([await Bun.file(S('sample.pptx')).arrayBuffer()]), {})
    const s2: any = await engine.call('pptx.slide', [2])
    const rounded = s2.els.find((e: any) => e.text?.paras[0]?.runs[0]?.text === 'Rounded')
    const pic = s2.els.find((e: any) => e.t === 'pic')
    expect(rounded.top).toBeGreaterThan(0)
    const m: any = await engine.call('pptx.shapes', [2, [
      { op: 'xfrm', top: rounded.top, box: { x: 10, y: 20, w: 300, h: 100 } },
      { op: 'fill', top: rounded.top, color: '#ff0000' },
      { op: 'delete', top: pic.top },
      { op: 'insert', kind: 'text', box: { x: 100, y: 500, w: 400, h: 50 }, text: 'New text' },
      { op: 'insert', kind: 'ellipse', box: { x: 600, y: 500, w: 100, h: 100 } },
    ]])
    expect(m.els.some((e: any) => e.t === 'pic')).toBe(false)
    const r2 = m.els.find((e: any) => e.top === rounded.top)
    expect(Math.round(r2.box.x)).toBe(10)
    expect(r2.fill.c).toBe('#ff0000')
    expect(m.els.some((e: any) => e.text?.paras[0]?.runs[0]?.text === 'New text')).toBe(true)
    // slides: duplicate slide 0, add a blank-ish content slide at end, move it first, delete original slide 4, hide slide 1
    let meta: any = await engine.call('pptx.slides', [{ op: 'duplicate', index: 0 }])
    expect(meta.count).toBe(6)
    meta = await engine.call('pptx.slides', [{ op: 'add', after: 5 }])
    expect(meta.count).toBe(7)
    meta = await engine.call('pptx.slides', [{ op: 'move', from: 6, to: 0 }])
    meta = await engine.call('pptx.slides', [{ op: 'delete', index: 5 }])
    meta = await engine.call('pptx.slides', [{ op: 'hide', index: 1, hidden: true }])
    expect(meta.count).toBe(6)
    const saved = (await engine.call('save', [])) as Blob
    const re = new Engine(() => {})
    const ri: any = await re.open(saved, {})
    expect(ri.pptx.count).toBe(6)
    const first: any = await re.call('pptx.slide', [0])
    expect(first.els.some((e: any) => e.t === 'shape' && e.text?.empty)).toBe(true)
    const dup: any = await re.call('pptx.slide', [2])
    expect(dup.els.some((e: any) => e.text?.paras[0]?.runs[0]?.text === 'office-engine')).toBe(true)
    expect(dup.hidden).toBeFalsy()
    const hid: any = await re.call('pptx.slide', [1])
    expect(hid.hidden).toBe(true)
    const shapes: any = await re.call('pptx.slide', [4])
    expect(shapes.els.some((e: any) => e.text?.paras[0]?.runs[0]?.text === 'New text')).toBe(true)
    void info
  })
})

describe('xlsx formulas + structure', () => {
  test('computes missing values, recalcs dependents, shifts on row insert, survives save', async () => {
    const mem = new Blob([await Bun.file(S('sample.xlsx')).arrayBuffer()])
    const { engine, cell, events } = await openXlsx(mem)
    // sample.xlsx Revenue column (E) is =C*D without cached values → computed after load
    while (!events.some((e) => e.event === 'xlsx:progress' && e.data.recalc)) await new Promise((r) => setTimeout(r, 10))
    const c3 = Number((await cell(2, 2))[1])
    const d3 = Number((await cell(2, 3))[1].replace(/[$,]/g, ''))
    expect(Number((await cell(2, 4))[1].replace(/,/g, ''))).toBeCloseTo(c3 * d3, 1)
    // edit an input → dependent recalculates
    const r1: any = await engine.call('xlsx.set', [0, [{ r: 2, c: 2, input: '10' }]])
    expect(r1.recalc).toContain(0)
    expect(Number((await cell(2, 4))[1].replace(/,/g, ''))).toBeCloseTo(10 * d3, 1)
    // new formula referencing a range
    await engine.call('xlsx.set', [0, [{ r: 70, c: 2, input: '=SUM(C3:C5)' }]])
    const sum = Number((await cell(70, 2))[1])
    // insert 2 rows above row 4 → formula range grows, value same
    const res: any = await engine.call('xlsx.sheetOp', [0, { op: 'insertRows', at: 3, n: 2 }])
    expect(res.layout.rows).toBeGreaterThan(70)
    const f: any = await engine.call('xlsx.cell', [0, 72, 2])
    expect(f.formula).toBe('=SUM(C3:C7)')
    expect(Number((await cell(72, 2))[1])).toBe(sum)
    // merge, column width, freeze, fill down, then delete a column
    await engine.call('xlsx.sheetOp', [0, { op: 'merge', range: { r0: 80, c0: 0, r1: 81, c1: 1 } }])
    await engine.call('xlsx.sheetOp', [0, { op: 'colWidth', c0: 7, c1: 7, px: 150 }])
    await engine.call('xlsx.sheetOp', [0, { op: 'freeze', rows: 1, cols: 0 }])
    await engine.call('xlsx.set', [0, [{ r: 90, c: 0, input: '=A3' }]])
    await engine.call('xlsx.sheetOp', [0, { op: 'fill', range: { r0: 90, c0: 0, r1: 92, c1: 0 }, dir: 'down' }])
    expect(((await engine.call('xlsx.cell', [0, 92, 0])) as any).formula).toBe('=A5')
    await engine.call('xlsx.format', [0, { r0: 90, c0: 0, r1: 92, c1: 1 }, { border: 'all' }])
    await engine.call('xlsx.bookOp', [{ op: 'rename', index: 0, name: 'Revenue' }])
    await engine.call('xlsx.bookOp', [{ op: 'add', name: 'New' }])
    const saved = (await engine.call('save', [])) as Blob
    const re = await openXlsx(saved)
    expect(re.info.xlsx!.sheets.map((s) => s.name)).toEqual(['Revenue', 'Notes', 'Data', 'New'])
    const p = re.events.filter((e) => e.event === 'xlsx:progress').map((e) => e.data)
    const last = p.filter((x) => x.done).at(-1)
    expect(last.merges).toContainEqual({ r0: 80, c0: 0, r1: 81, c1: 1 })
    expect(p.find((x) => x.frozen)?.frozen).toEqual({ rows: 1, cols: 0 })
    expect(p.find((x) => x.colInfo)?.colInfo.some((ci: any) => ci.min === 7 && ci.w === 150)).toBe(true)
    expect(Number((await re.cell(72, 2))[1])).toBe(sum)
    expect(((await re.engine.call('xlsx.cell', [0, 92, 0])) as any).formula).toBe('=A5')
  }, 60000)
})

describe('xlsx row heights', () => {
  test('rows grown by a large font keep their height after inserting rows', async () => {
    const mem = new Blob([await Bun.file(S('sample.xlsx')).arrayBuffer()])
    const { engine, events } = await openXlsx(mem)
    const loaded = new Map<number, number>()
    for (const e of events) if (e.event === 'xlsx:progress' && e.data.sheet === 0) for (const [r, px] of e.data.rowSizes) loaded.set(r, px)
    const title = loaded.get(0)
    expect(title).toBeGreaterThan(20)
    const res: any = await engine.call('xlsx.sheetOp', [0, { op: 'insertRows', at: 0, n: 1 }])
    const after = new Map<number, number>(res.layout.rowSizes)
    expect(after.get(1)).toBe(title)
    expect(after.has(0)).toBe(false)
    const undo: any = await engine.call('xlsx.undo', [])
    expect(new Map<number, number>(undo.layout.rowSizes).get(0)).toBe(title)
  })

  test('text with line breaks: the cell wraps and its row grows to show every line, undone in one step', async () => {
    const mem = new Blob([await Bun.file(S('sample.xlsx')).arrayBuffer()])
    const { engine, cell } = await openXlsx(mem)
    const rowPx = (res: any, r: number) => new Map<number, number>(res.layout?.rowSizes).get(r)

    const res: any = await engine.call('xlsx.set', [0, [{ r: 70, c: 9, input: 'one\ntwo\nthree' }, { r: 71, c: 9, input: 'plain' }]])
    expect(res.styles.map((x: any) => x.css.whiteSpace)).toEqual(['pre-wrap'])
    expect((await cell(70, 9))[2]).toBe(res.styles[0].index)
    expect(rowPx(res, 70)).toBe(60)
    expect(rowPx(res, 71)).toBeUndefined()
    // the same style again, a row already tall enough: nothing new
    const again: any = await engine.call('xlsx.set', [0, [{ r: 70, c: 10, input: 'a\nb' }]])
    expect(again.styles).toBeUndefined()
    expect(again.layout).toBeUndefined()
    expect((await cell(70, 10))[2]).toBe(res.styles[0].index)
    await engine.call('xlsx.undo', [])

    // a formula spread over lines, a number: no text to wrap
    const other: any = await engine.call('xlsx.set', [0, [{ r: 72, c: 9, input: '=1+\n2' }, { r: 73, c: 9, input: '5\n' }]])
    expect(other.styles).toBeUndefined()
    expect(other.layout).toBeUndefined()
    expect((await cell(72, 9))[1]).toBe('3')
    // a merged cell spans rows of its own: its first row is left alone
    await engine.call('xlsx.sheetOp', [0, { op: 'merge', range: { r0: 75, c0: 9, r1: 76, c1: 9 } }])
    const merged: any = await engine.call('xlsx.set', [0, [{ r: 75, c: 9, input: 'x\ny\nz' }]])
    expect((await cell(75, 9))[2]).toBe(res.styles[0].index)
    expect(merged.layout).toBeUndefined()

    const saved = (await engine.call('save', [])) as Blob
    const re = await openXlsx(saved)
    expect((await re.cell(70, 9))[1]).toBe('one\ntwo\nthree')
    expect(re.info.xlsx!.styles[(await re.cell(70, 9))[2]].whiteSpace).toBe('pre-wrap')
    expect(re.info.xlsx!.styles[(await re.cell(75, 9))[2]].whiteSpace).toBe('pre-wrap')
    const sizes = new Map<number, number>()
    for (const e of re.events) if (e.event === 'xlsx:progress' && e.data.sheet === 0) for (const [r, px] of e.data.rowSizes) sizes.set(r, px)
    expect(sizes.get(70)).toBe(60)
    expect(sizes.has(75)).toBe(false)

    // one undo: the text, the wrap and the row height
    for (let i = 0; i < 3; i++) await engine.call('xlsx.undo', [])
    const undone: any = await engine.call('xlsx.undo', [])
    expect(rowPx(undone, 70)).toBeUndefined()
    expect(await cell(70, 9)).toBeUndefined()
    const redone: any = await engine.call('xlsx.redo', [])
    expect(rowPx(redone, 70)).toBe(60)
    expect((await cell(70, 9))[1]).toBe('one\ntwo\nthree')
  })
})

describe('xlsx toolbar ops', () => {
  test('alignment, rotation, clear/paint format, merge across, hide, gridlines, find & replace', async () => {
    const mem = new Blob([await Bun.file(S('sample.xlsx')).arrayBuffer()])
    const { engine, cell } = await openXlsx(mem)
    const f1: any = await engine.call('xlsx.format', [0, { r0: 2, c0: 0, r1: 3, c1: 0 }, { valign: 'top', rotate: 45, borderStyle: 'dashed', border: 'outer' }])
    const css = f1.styles.at(-1).css
    expect(css.alignItems).toBe('flex-start')
    expect(css.textRotate).toBe(45)
    const xf = (await cell(2, 0))[2]
    // paint format copies the exact style; clear resets it
    await engine.call('xlsx.format', [0, { r0: 5, c0: 5, r1: 5, c1: 5 }, { xf }])
    expect((await cell(5, 5))[2]).toBe(xf)
    await engine.call('xlsx.format', [0, { r0: 2, c0: 0, r1: 2, c1: 0 }, { clear: true }])
    expect((await cell(2, 0))[2]).toBe(0)
    // merge across: one merge per row
    const m: any = await engine.call('xlsx.sheetOp', [0, { op: 'merge', range: { r0: 80, c0: 0, r1: 82, c1: 2 }, mode: 'across' }])
    expect(m.layout.merges.filter((x: any) => x.r0 >= 80).map((x: any) => `${x.r0}:${x.c0}-${x.r1}:${x.c1}`)).toEqual(['80:0-80:2', '81:0-81:2', '82:0-82:2'])
    // hide / unhide columns keep the width
    const w0 = m.layout.colInfo.find((c: any) => c.min <= 1 && c.max >= 1)?.w ?? m.layout.defaultColW
    const h: any = await engine.call('xlsx.sheetOp', [0, { op: 'hideCols', a: 1, b: 2, hidden: true }])
    expect(h.layout.colInfo.find((c: any) => c.min <= 1 && c.max >= 1).w).toBe(0)
    const u: any = await engine.call('xlsx.sheetOp', [0, { op: 'hideCols', a: 1, b: 2, hidden: false }])
    expect(u.layout.colInfo.find((c: any) => c.min <= 1 && c.max >= 1).w).toBe(w0)
    await engine.call('xlsx.sheetOp', [0, { op: 'hideRows', a: 10, b: 11, hidden: true }])
    await engine.call('xlsx.sheetOp', [0, { op: 'hideCols', a: 7, b: 7, hidden: true }])
    const g: any = await engine.call('xlsx.sheetOp', [0, { op: 'gridlines', show: false }])
    expect(g.layout.gridLines).toBe(false)
    // find & replace
    const found: any = await engine.call('xlsx.find', [0, 'central', {}])
    expect(found.length).toBeGreaterThan(2)
    expect((await engine.call('xlsx.find', [0, 'central', { matchCase: true }]) as any).length).toBe(0)
    const rep: any = await engine.call('xlsx.replace', [0, 'Central', 'Trung tâm', {}])
    expect(rep.range).toBeTruthy()
    expect((await engine.call('xlsx.find', [0, 'Central', {}]) as any).length).toBe(0)
    expect((await engine.call('xlsx.find', [0, 'Trung tâm', { whole: true }]) as any).length).toBe(found.length)
    const saved = (await engine.call('save', [])) as Blob
    const xml = await partText(saved, 'xl/worksheets/sheet1.xml')
    expect(xml).toContain('showGridLines="0"')
    expect(xml).toMatch(/<row r="11"[^>]*hidden="1"/)
    expect(xml).toMatch(/<col min="8" max="8" width="[\d.]+" customWidth="1" hidden="1"\/>/)
    expect(xml).toContain('<mergeCell ref="A82:C82"/>')
    const styles = await partText(saved, 'xl/styles.xml')
    expect(styles).toContain('textRotation="45"')
    expect(styles).toContain('vertical="top"')
    expect(styles).toContain('style="dashed"')
  })
})

describe('xlsx new sheet', () => {
  test('add sheet, edit it, save twice', async () => {
    const { engine, events } = await openXlsx(new Blob([await Bun.file(S('sample.xlsx')).arrayBuffer()]))
    await engine.call('xlsx.bookOp', [{ op: 'add', name: 'Summary', after: 0 }])
    await engine.call('xlsx.load', [1])
    while (!events.some((e) => e.event === 'xlsx:progress' && e.data.sheet === 1 && e.data.done)) await new Promise((r) => setTimeout(r, 10))
    expect(events.some((e) => e.event === 'xlsx:progress' && e.data.sheet === 1 && e.data.error)).toBe(false)
    await engine.call('xlsx.set', [1, [{ r: 0, c: 0, input: '=Sales!C3*2' }, { r: 1, c: 1, input: 'hello' }]])
    const b: any = await engine.call('xlsx.block', [1, 0, 2])
    expect(Number(b.rows[0].cells[0][1])).toBeGreaterThan(0)
    await engine.call('save', [])
    await engine.call('xlsx.set', [1, [{ r: 2, c: 0, input: '42' }]])
    const saved = (await engine.call('save', [])) as Blob
    const re = await openXlsx(saved, 1)
    expect(re.info.xlsx!.sheets[1].name).toBe('Summary')
    expect((await re.cell(1, 1))[1]).toBe('hello')
    expect((await re.cell(2, 0))[1]).toBe('42')
    expect(((await re.engine.call('xlsx.cell', [1, 0, 0])) as any).formula).toBe('=Sales!C3*2')
  })
})

describe('xlsx state survives save/reopen exactly (collaboration snapshots)', () => {
  test('circular references evaluate to #REF! however often they are recalculated', async () => {
    const mem = new Blob([await Bun.file(S('sample.xlsx')).arrayBuffer()])
    const { engine, cell } = await openXlsx(mem)
    for (let i = 0; i < 3; i++) {
      // recalculated each time: the previous value must not leak into the result
      await engine.call('xlsx.set', [0, [{ r: 79, c: 0, input: '=A80&"!"' }]])
      await engine.call('xlsx.set', [0, [{ r: 80, c: 0, input: '=B81+1' }, { r: 80, c: 1, input: '=A81+1' }]])
    }
    expect((await cell(79, 0))[1]).toBe('#REF!')
    expect((await cell(80, 0))[1]).toBe('#REF!')
    expect((await cell(80, 1))[1]).toBe('#REF!')
    const re = await openXlsx((await engine.call('save', [])) as Blob)
    expect((await re.cell(79, 0))[1]).toBe('#REF!')
    expect((await re.cell(80, 0))[1]).toBe('#REF!')
  })

  test('formatting after reopening reuses the styles derived before saving', async () => {
    const mem = new Blob([await Bun.file(S('sample.xlsx')).arrayBuffer()])
    const { engine, cell } = await openXlsx(mem)
    await engine.call('xlsx.set', [0, [{ r: 90, c: 0, input: 'a' }, { r: 91, c: 0, input: 'b' }]])
    await engine.call('xlsx.format', [0, { r0: 90, c0: 0, r1: 90, c1: 0 }, { b: true, fill: '#ffee00' }])
    const live = (await cell(90, 0))[2]
    const re = await openXlsx((await engine.call('save', [])) as Blob)
    const before = re.info.xlsx!.styles.length
    await re.engine.call('xlsx.format', [0, { r0: 91, c0: 0, r1: 91, c1: 0 }, { b: true, fill: '#ffee00' }])
    // same base style + same patch → the identical style already in the file, nothing appended
    expect((await re.cell(91, 0))[2]).toBe(live)
    const saved = await ZipArchive.open(new BlobSource((await re.engine.call('save', [])) as Blob))
    const xfs = new TextDecoder().decode(await saved.bytes('xl/styles.xml')).match(/<cellXfs[^>]*count="(\d+)"/)?.[1]
    expect(Number(xfs)).toBe(before)
  })
})
