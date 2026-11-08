import { describe, expect, test } from 'bun:test'
import { DocumentClient } from '../src/core/worker/client'
import type { Block, Paragraph } from '../src/core/docx/types'
import { CollabHub } from '../src/collab/hub'
import { registerHost } from '../src/collab/hosts'
import { startCollab, type Collab } from '../src/collab'
import { MemoryTransport } from '../src/collab/transport'
import { DocxEditor, type DocState } from '../src/react/docx/DocxEditor'
import { opBackspace, opDelete, opEnter, opInsert } from '../src/react/edit/input'
import { plen, runsText, type RRun } from '../src/react/edit/rich'
import { BlobSource } from '../src/core/source'
import { ZipArchive } from '../src/core/zip'
import { rebuildZip } from '../src/core/zipwrite'

const S = (n: string) => `${import.meta.dir}/../playground/public/samples/${n}`
const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms))

/** `lag`: ms the engine takes to answer an edit, as a worker does (an edit made meanwhile is sent before the answer is back) */
async function load(src: Blob, lag = 0) {
  const client = await DocumentClient.open(src, { worker: false })
  if (lag) {
    const direct = client.direct.bind(client) as (m: string, ...a: unknown[]) => Promise<unknown>
    client.direct = ((m: string, ...a: unknown[]) => direct(m, ...a).then(async (r) => (m === 'docx.edit' ? (await sleep(lag), r) : r))) as typeof client.direct
  }
  const st: DocState = { blocks: [], heights: new WeakMap(), dirty: new Set(), frontier: 0 }
  client.on('docx:blocks', (b) => st.blocks.push(...b))
  const done = new Promise<void>((r) => client.on('docx:done', () => r()))
  void client.call('docx.start')
  await done
  return { client, st }
}

async function open(hub: CollabHub, name: string, latency: () => number, src: Blob = Bun.file(S('sample.docx')) as unknown as Blob, lag = 0) {
  const { client, st } = await load(src, lag)
  const editor = new DocxEditor(client, st, () => {})
  registerHost(client, 'docx', { editor, st })
  const transport = new MemoryTransport(hub, latency)
  const collab = startCollab(client, { transport, room: 'd', user: { name } })
  return { client, st, editor, collab, transport }
}

async function settle(cs: Collab[]) {
  for (let i = 0; i < 300; i++) {
    await sleep(10)
    await Promise.all(cs.map((c) => c.session.idle()))
    if (cs.every((c) => c.pending === 0 && c.status === 'live') && new Set(cs.map((c) => c.session.rev)).size === 1) return
  }
  throw new Error('did not settle')
}

const text = (blocks: Block[]): string[] =>
  blocks.flatMap((b) => (b.t === 'p' ? [`${b.pid}|${b.sid ?? ''}|${runsText(b.runs as RRun[])}`] : b.t === 'tbl' ? b.rows.flatMap((r) => r.cells.flatMap((c) => text(c.blocks))) : []))

/** editable top-level paragraphs */
const editable = (st: DocState) => st.blocks.map((b, i) => [b, i] as const).filter(([b]) => b.t === 'p' && b.pid !== undefined && !b.ro) as [Paragraph, number][]

describe('collab docx', () => {
  for (const start of [3, 8, 21, 55, 77, 101, 131, 202]) {
    test(`concurrent typing, splitting, joining and formatting converge (seed ${start})`, async () => {
      let seed = start
      const rand = () => ((seed = (seed * 1664525 + 1013904223) % 4294967296) / 4294967296)
      const hub = new CollabHub()
      const a = await open(hub, 'Ann', () => Math.floor(rand() * 6))
      const b = await open(hub, 'Bob', () => Math.floor(rand() * 6))
      await settle([a.collab, b.collab])
      for (let i = 0; i < 120; i++) {
        const who = rand() < 0.5 ? a : b
        const paras = editable(who.st)
        // both users often work in the same few paragraphs
        const [p, bi] = paras[Math.floor(rand() * Math.min(paras.length, 4))]
        const pa = String(bi)
        const n = plen(p.runs as RRun[])
        const o = Math.floor(rand() * (n + 1))
        const x = rand()
        const at = (o: number) => ({ a: { pa, o }, f: { pa, o } })
        if (x < 0.45) opInsert(who.editor, at(o), who === a ? 'a' + i : 'B' + i)
        else if (x < 0.6 && n > 0) opDelete(who.editor, { a: { pa, o: Math.max(0, o - 2) }, f: { pa, o } })
        else if (x < 0.7) opEnter(who.editor, at(o))
        else if (x < 0.77 && bi > 0) opBackspace(who.editor, at(0))
        else if (x < 0.85 && n > 1) {
          who.editor.input.lastSel = { a: { pa, o: 0 }, f: { pa, o: Math.min(n, 3) } }
          who.editor.toggle('b')
        } else if (x < 0.87 && n > 1) {
          who.editor.input.lastSel = { a: { pa, o: 0 }, f: { pa, o: Math.min(n, 2) } }
          void who.editor.setLink('example.com/' + i)
        } else if (x < 0.89) {
          who.editor.input.lastSel = at(o)
          who.editor.insertTable(2, 2)
        } else if (x < 0.92) {
          // type into a table cell
          const ti = who.st.blocks.findIndex((b) => b.t === 'tbl')
          if (ti >= 0) opInsert(who.editor, { a: { pa: `${ti}.0.0.0`, o: 0 }, f: { pa: `${ti}.0.0.0`, o: 0 } }, 'c' + i)
        } else if (x < 0.96) who.editor.undo()
        else who.editor.redo()
        if (rand() < 0.5) await sleep(Math.floor(rand() * 4))
      }
      await settle([a.collab, b.collab])
      // pid + text (the style id on a model paragraph is display state: compared on the saved files)
      const ta = text(a.st.blocks).map((x) => x.replace(/\|[^|]*\|/, '||'))
      const tb = text(b.st.blocks).map((x) => x.replace(/\|[^|]*\|/, '||'))
      expect(ta.join('\n')).toMatch(/[aB]\d+/)
      expect(ta).toEqual(tb)

      // the saved files hold the same text
      const [sa, sb] = await Promise.all([a.client.call('save'), b.client.call('save')])
      const ra = await load(sa)
      const rb = await load(sb)
      const strip = (xs: string[]) => xs.map((x) => x.split('|').slice(2).join('|'))
      const styled = (xs: string[]) => xs.map((x) => x.split('|').slice(1).join('|'))
      if (process.env.DEBUG_COLLAB) {
        const da = strip(text(ra.st.blocks))
        const ma = strip(ta)
        const db = strip(text(rb.st.blocks))
        const mb = strip(tb)
        for (let k = 0; k < Math.max(da.length, ma.length); k++) if (da[k] !== ma[k]) console.log('A saved≠model at', k, JSON.stringify(da[k]), JSON.stringify(ma[k]), ta[k])
        for (let k = 0; k < Math.max(db.length, mb.length); k++) if (db[k] !== mb[k]) console.log('B saved≠model at', k, JSON.stringify(db[k]), JSON.stringify(mb[k]), tb[k])
      }
      expect(styled(text(ra.st.blocks))).toEqual(styled(text(rb.st.blocks)))
      expect(strip(text(ra.st.blocks))).toEqual(strip(ta))

      // late joiner
      const c = await open(hub, 'Cid', () => 0)
      await settle([a.collab, b.collab, c.collab])
      expect(text(c.st.blocks).map((x) => x.replace(/\|[^|]*\|/, '||'))).toEqual(ta)
      for (const x of [a, b, c, ra, rb]) x.client.destroy()
      for (const x of [a, b, c]) x.collab.close()
    }, 30_000)
  }

  test('a built-in style the document lacks is created for every editor, late joiners included', async () => {
    const zip = await ZipArchive.open(new BlobSource(new Blob([await Bun.file(S('sample.docx')).arrayBuffer()])))
    const styles = (await zip.text('word/styles.xml')).replace(/<w:style [^>]*w:styleId="(Heading\d|Title|Subtitle)"[^>]*>[\s\S]*?<\/w:style>/g, '')
    const src = await rebuildZip(zip, new Map([['word/styles.xml', styles]]))
    const hub = new CollabHub()
    const a = await open(hub, 'Ann', () => 1, src)
    const b = await open(hub, 'Bob', () => 1, src)
    await settle([a.collab, b.collab])
    expect(a.editor.styles.some((s) => s.canon === 'h1')).toBe(false)
    const [p, bi] = editable(a.st).find(([x]) => runsText(x.runs as RRun[]).startsWith('This document shows'))!
    const at = { a: { pa: String(bi), o: 0 }, f: { pa: String(bi), o: 0 } }
    a.editor.input.lastSel = at
    // the style is added by a shared operation (every copy of the file gets it), then the paragraph takes it
    await a.editor.applyCanon('h1')
    opInsert(a.editor, at, 'A ')
    await settle([a.collab, b.collab])
    const c = await open(hub, 'Cid', () => 0, src)
    await settle([a.collab, b.collab, c.collab])
    for (const x of [a, b, c]) {
      const q = x.st.blocks.find((y) => y.t === 'p' && y.pid === p.pid) as Paragraph
      expect(runsText(q.runs as RRun[])).toStartWith('A This document shows')
      const saved = await load(await x.client.call('save'))
      const r = saved.st.blocks.find((y) => y.t === 'p' && runsText(y.runs as RRun[]).startsWith('A This document shows')) as Paragraph
      expect(r.sid).toBe('Heading1')
      expect(r.ps.outline).toBe(0)
      saved.client.destroy()
    }
    for (const x of [a, b, c]) x.client.destroy()
    for (const x of [a, b, c]) x.collab.close()
  })

  test('a paragraph edited again right after it was restyled shows the look of its style, also where the log is replayed', async () => {
    const sample = Bun.file(S('sample.docx')) as unknown as Blob
    const hub = new CollabHub()
    const a = await open(hub, 'Ann', () => 1, sample, 5)
    await settle([a.collab])
    const [p, bi] = editable(a.st).find(([x]) => runsText(x.runs as RRun[]).startsWith('This document shows'))!
    const at = { a: { pa: String(bi), o: 0 }, f: { pa: String(bi), o: 0 } }
    // typing before the engine answered the style change
    a.editor.setPara({ style: 'Heading2' }, {}, at)
    opInsert(a.editor, at, 'XY ')
    await settle([a.collab])
    // a newcomer replays both edits back to back
    const b = await open(hub, 'Bob', () => 0, sample, 5)
    await settle([a.collab, b.collab])
    await sleep(100)
    /** the look of every character of a paragraph (however its runs are cut) */
    const look = (q: Paragraph[]) => q.flatMap((x) => x.runs.flatMap((r) => (r.t === 'text' ? [...r.text].map((ch) => ch + JSON.stringify(Object.entries(r.s ?? {}).sort())) : [r.t])))
    const shown = (st: DocState) => look(st.blocks.filter((x): x is Paragraph => x.t === 'p' && x.pid === p.pid))
    // what the file says
    const saved = await load(await a.client.call('save'))
    const want = look(saved.st.blocks.filter((x): x is Paragraph => x.t === 'p' && runsText(x.runs as RRun[]).startsWith('XY This document shows')))
    expect(want.length).toBeGreaterThan(20)
    expect(want[0]).not.toBe(look([p])[0])
    expect(shown(a.st)).toEqual(want)
    expect(shown(b.st)).toEqual(want)
    for (const x of [a, b, saved]) x.client.destroy()
    for (const x of [a, b]) x.collab.close()
  })

  test('a rebase puts back a paragraph that stood between two tables, which a pending join had taken away', async () => {
    const hub = new CollabHub()
    const a = await open(hub, 'Ann', () => 0)
    const b = await open(hub, 'Bob', () => 0)
    await settle([a.collab, b.collab])
    const [title, intro] = editable(a.st).map(([p]) => p.pid!)
    const at = (st: DocState, pid: number, o: number) => {
      const pa = String(st.blocks.findIndex((x) => x.t === 'p' && x.pid === pid))
      return { a: { pa, o }, f: { pa, o } }
    }
    // a table follows the second paragraph
    b.editor.input.lastSel = at(b.st, intro, 0)
    b.editor.insertTable(1, 1)
    await settle([a.collab, b.collab])
    // offline, Ann joins the second paragraph to the first
    a.transport.drop()
    opBackspace(a.editor, at(a.st, intro, 0))
    // meanwhile Bob puts a table between the two, then splits the second one: two batches
    b.editor.input.lastSel = at(b.st, title, 0)
    b.editor.insertTable(1, 1)
    await settle([b.collab])
    opEnter(b.editor, at(b.st, intro, 5))
    await settle([b.collab])
    // Ann's join is replayed after the first batch (the paragraph, now apart, goes away on its own), then reverted for the second
    a.transport.connect()
    await settle([a.collab, b.collab])
    const strip = (xs: string[]) => xs.map((x) => x.split('|').slice(2).join('|'))
    const ta = text(a.st.blocks).map((x) => x.replace(/\|[^|]*\|/, '||'))
    expect(ta).toEqual(text(b.st.blocks).map((x) => x.replace(/\|[^|]*\|/, '||')))
    // title + joined text, the first table's cell, what Bob split off, the second table's cell
    expect(strip(ta).slice(0, 4)).toEqual(['Quarterly ReportThis document shows bold, italic, underlined red and large text. Visit example.com.', '', 'document shows bold, italic, underlined red and large text. Visit example.com.', ''])
    // the files say the same, in the same order
    const [ra, rb] = await Promise.all([a, b].map(async (x) => load(await x.client.call('save'))))
    expect(strip(text(ra.st.blocks))).toEqual(strip(ta))
    expect(strip(text(rb.st.blocks))).toEqual(strip(ta))
    for (const x of [a, b, ra, rb]) x.client.destroy()
    for (const x of [a, b]) x.collab.close()
  })

  test('undo after typing in two paragraphs that had the same address in turn only takes back the last one', async () => {
    const hub = new CollabHub()
    const a = await open(hub, 'Ann', () => 0)
    const b = await open(hub, 'Bob', () => 0)
    await settle([a.collab, b.collab])
    const [title, intro] = editable(a.st).map(([p]) => p.pid!)
    const at = (st: DocState, pid: number, o: number) => {
      const pa = String(st.blocks.findIndex((x) => x.t === 'p' && x.pid === pid))
      return { a: { pa, o }, f: { pa, o } }
    }
    opInsert(a.editor, at(a.st, intro, 0), 'x')
    await settle([a.collab, b.collab])
    // Bob splits the title: its second half now stands where the paragraph Ann typed in was
    opEnter(b.editor, at(b.st, title, 9))
    await settle([a.collab, b.collab])
    const half = (a.st.blocks[1] as Paragraph).pid!
    expect(half).not.toBe(intro)
    opInsert(a.editor, at(a.st, half, 0), 'y')
    a.editor.undo()
    await settle([a.collab, b.collab])
    const ta = text(a.st.blocks).map((x) => x.replace(/\|[^|]*\|/, '||'))
    expect(ta.slice(0, 3)).toEqual([`${title}||Quarterly`, `${half}|| Report`, `${intro}||xThis document shows bold, italic, underlined red and large text. Visit example.com.`])
    expect(text(b.st.blocks).map((x) => x.replace(/\|[^|]*\|/, '||'))).toEqual(ta)
    for (const x of [a, b]) x.client.destroy()
    for (const x of [a, b]) x.collab.close()
  })

  test('undoing a table insert does not bring back the paragraph before it, joined to another one meanwhile', async () => {
    const hub = new CollabHub()
    const a = await open(hub, 'Ann', () => 0)
    const b = await open(hub, 'Bob', () => 0)
    await settle([a.collab, b.collab])
    const [, intro] = editable(a.st).map(([p]) => p.pid!)
    const at = (st: DocState, pid: number, o: number) => {
      const pa = String(st.blocks.findIndex((x) => x.t === 'p' && x.pid === pid))
      return { a: { pa, o }, f: { pa, o } }
    }
    b.editor.input.lastSel = at(b.st, intro, 0)
    b.editor.insertTable(1, 1)
    await settle([a.collab, b.collab])
    // offline, Bob takes his table back: an op that names the paragraph before the table
    b.transport.drop()
    b.editor.undo()
    // Ann's join is sequenced first
    opBackspace(a.editor, at(a.st, intro, 0))
    await settle([a.collab])
    b.transport.connect()
    await settle([a.collab, b.collab])
    const strip = (xs: string[]) => xs.map((x) => x.split('|').slice(2).join('|'))
    const ta = text(a.st.blocks).map((x) => x.replace(/\|[^|]*\|/, '||'))
    expect(ta).toEqual(text(b.st.blocks).map((x) => x.replace(/\|[^|]*\|/, '||')))
    expect(strip(ta).slice(0, 2)).toEqual(['Quarterly ReportThis document shows bold, italic, underlined red and large text. Visit example.com.', '1. Highlights'])
    // what is on screen is what the files hold
    const [ra, rb] = await Promise.all([a, b].map(async (x) => load(await x.client.call('save'))))
    expect(strip(text(ra.st.blocks))).toEqual(strip(ta))
    expect(strip(text(rb.st.blocks))).toEqual(strip(ta))
    for (const x of [a, b, ra, rb]) x.client.destroy()
    for (const x of [a, b]) x.collab.close()
  })

  test('a style updated to match the selection: its paragraphs follow for everyone, in the files too, and undo takes it back', async () => {
    const hub = new CollabHub()
    const a = await open(hub, 'Ann', () => 5)
    const b = await open(hub, 'Bob', () => 5)
    await settle([a.collab, b.collab])
    const heads = (st: DocState) => st.blocks.filter((x): x is Paragraph => x.t === 'p' && x.sid === 'Heading1')
    const looks = (st: DocState) => heads(st).map((p) => p.runs.flatMap((r) => (r.t === 'text' ? [`${r.s.color} ${r.s.sz}`] : [])).join())
    /** the paragraphs take the look of their style a moment after the style changed */
    const shown = async (want: string) => {
      for (let i = 0; i < 200 && ![a, b].every((x) => looks(x.st).every((l) => l === want)); i++) await sleep(10)
      for (const x of [a, b]) expect(looks(x.st)).toEqual(heads(x.st).map(() => want))
    }
    const [first, second] = heads(a.st)
    const bi = a.st.blocks.indexOf(first)
    const all = { a: { pa: String(bi), o: 0 }, f: { pa: String(bi), o: plen(first.runs as RRun[]) } }
    a.editor.input.lastSel = all
    a.editor.setColor('#c00000')
    a.editor.setSize(20)
    await settle([a.collab, b.collab])
    expect(looks(b.st)).toEqual(['#c00000 20', '#365f91 14', '#365f91 14'])

    // offline, Bob types into another heading: text that comes with the look the style had
    b.transport.drop()
    const bj = b.st.blocks.findIndex((x) => x.t === 'p' && x.pid === second.pid)
    opInsert(b.editor, { a: { pa: String(bj), o: 3 }, f: { pa: String(bj), o: 3 } }, 'late ')
    expect(await a.editor.updateStyle('Heading1')).toBe(true)
    await settle([a.collab])
    b.transport.connect()
    await settle([a.collab, b.collab])
    await shown('#c00000 20')
    expect(runsText((a.st.blocks.find((x) => x.t === 'p' && x.pid === second.pid) as Paragraph).runs as RRun[])).toContain('late ')
    // the paragraph picked leaves colour and size to the style now
    const picked = (st: DocState) => (st.blocks.find((x) => x.t === 'p' && x.pid === first.pid) as Paragraph).runs.find((r) => r.t === 'text') as { d?: object }
    for (const x of [a, b]) expect(picked(x.st).d).toMatchObject({ color: null, sz: null })
    // asking again changes nothing
    a.editor.input.lastSel = all
    expect(await a.editor.updateStyle('Heading1')).toBe(false)
    // who joins later replays the room: the same look
    const c = await open(hub, 'Cid', () => 0)
    await settle([a.collab, b.collab, c.collab])
    for (let i = 0; i < 200 && looks(c.st).join() !== looks(a.st).join(); i++) await sleep(10)
    expect(looks(c.st)).toEqual(['#c00000 20', '#c00000 20', '#c00000 20'])
    expect(text(c.st.blocks)).toEqual(text(a.st.blocks))
    c.client.destroy()
    c.collab.close()

    // the files hold the style, and the paragraph without its own colour
    const saved = await Promise.all([a, b].map(async (x) => load(await x.client.call('save'))))
    for (const x of saved) expect(looks(x.st)).toEqual(['#c00000 20', '#c00000 20', '#c00000 20'])
    const xml = await (await ZipArchive.open(new BlobSource(await b.client.call('save')))).text('word/document.xml')
    expect(/<w:p[ >](?:(?!<\/w:p>)[\s\S])*?1\. Highlights[\s\S]*?<\/w:p>/.exec(xml)![0]).not.toContain('C00000')

    // Ann takes it back: the style as it was, her paragraph with the colour and size she gave it
    a.editor.undo()
    await settle([a.collab, b.collab])
    for (let i = 0; i < 200 && looks(b.st).join() !== '#c00000 20,#365f91 14,#365f91 14'; i++) await sleep(10)
    for (const x of [a, b]) expect(looks(x.st)).toEqual(['#c00000 20', '#365f91 14', '#365f91 14'])
    for (const x of [a, b]) expect(picked(x.st).d).toEqual({ color: '#c00000', sz: 20 })
    a.editor.redo()
    await settle([a.collab, b.collab])
    await shown('#c00000 20')
    for (const x of [a, b, ...saved]) x.client.destroy()
    for (const x of [a, b]) x.collab.close()
  })

  test('a style updated from a paragraph of another style, in a table: the paragraph gets the style, what the table style adds stays out of it', async () => {
    const { client, st } = await load(Bun.file(S('sample.docx')) as unknown as Blob)
    const editor = new DocxEditor(client, st, () => {})
    const heads = () => st.blocks.filter((x): x is Paragraph => x.t === 'p' && x.sid === 'Heading2')
    const before = heads().length
    expect(before).toBeGreaterThan(0)
    const ti = st.blocks.findIndex((x) => x.t === 'tbl')
    const cell = () => (st.blocks[ti] as Extract<Block, { t: 'tbl' }>).rows[0].cells[0].blocks[0] as Paragraph
    const italic = (p: Paragraph) => p.runs.every((r) => r.t !== 'text' || r.s.i === true)
    // the table style packs its paragraphs: no space after, single line spacing — the document's paragraphs have both
    const body = st.blocks.find((x): x is Paragraph => x.t === 'p' && x.sid === 'Normal')!
    expect([cell().sid, cell().ps.after, cell().ps.line]).toEqual(['Normal', 0, 1])
    expect([body.ps.after, body.ps.line]).toEqual([10, 1.15])
    const pa = `${ti}.0.0.0`
    editor.input.lastSel = { a: { pa, o: 0 }, f: { pa, o: plen(cell().runs as RRun[]) } }
    editor.toggle('i')
    expect(await editor.updateStyle('@h2')).toBe(true)
    for (let i = 0; i < 200 && !(heads().every(italic) && cell().sid === 'Heading2'); i++) await sleep(10)
    // the headings take the look the cell's text has outside a table: italic, the document's own spacing
    expect(heads().length).toBe(before)
    for (const p of heads()) expect([italic(p), p.ps.after, p.ps.line, p.ps.outline]).toEqual([true, 10, 1.15, 1])
    // the cell is a heading now, and looks as it did
    expect([cell().sid, italic(cell()), cell().ps.after, cell().ps.line]).toEqual(['Heading2', true, 0, 1])
    const saved = await load(await client.call('save'))
    const again = (saved.st.blocks[ti] as Extract<Block, { t: 'tbl' }>).rows[0].cells[0].blocks[0] as Paragraph
    expect([again.sid, italic(again), again.ps.after, again.ps.line]).toEqual(['Heading2', true, 0, 1])
    for (const x of [client, saved.client]) x.destroy()
  })

  test('text typed while the operation adding a missing style is on its way is kept for everyone', async () => {
    const zip = await ZipArchive.open(new BlobSource(new Blob([await Bun.file(S('sample.docx')).arrayBuffer()])))
    const styles = (await zip.text('word/styles.xml')).replace(/<w:style [^>]*w:styleId="(Heading\d|Title|Subtitle)"[^>]*>[\s\S]*?<\/w:style>/g, '')
    const src = await rebuildZip(zip, new Map([['word/styles.xml', styles]]))
    const hub = new CollabHub()
    const a = await open(hub, 'Ann', () => 5, src)
    const b = await open(hub, 'Bob', () => 5, src)
    await settle([a.collab, b.collab])
    const [p, bi] = editable(a.st).find(([x]) => runsText(x.runs as RRun[]).startsWith('This document shows'))!
    const at = { a: { pa: String(bi), o: 0 }, f: { pa: String(bi), o: 0 } }
    a.editor.input.lastSel = at
    // the style is added by an operation the hub orders first; the paragraph takes the style once it is there
    const styled = a.editor.applyCanon('h1')
    await a.collab.session.idle()
    opInsert(a.editor, at, 'typed meanwhile ')
    await styled
    await settle([a.collab, b.collab])
    const c = await open(hub, 'Cid', () => 0, src)
    await settle([a.collab, b.collab, c.collab])
    for (const x of [a, b, c]) {
      const q = x.st.blocks.find((y) => y.t === 'p' && y.pid === p.pid) as Paragraph
      expect(runsText(q.runs as RRun[])).toStartWith('typed meanwhile This document shows')
      expect(q.sid).toBe('Heading1')
    }
    const saved = await load(await b.client.call('save'))
    const r = saved.st.blocks.find((y) => y.t === 'p' && runsText(y.runs as RRun[]).startsWith('typed meanwhile This document shows')) as Paragraph
    expect(r.sid).toBe('Heading1')
    for (const x of [a, b, c, saved]) x.client.destroy()
    for (const x of [a, b, c]) x.collab.close()
  })
})

describe('collab docx table formatting', () => {
  test('merged cells, fill, borders and column widths reach the other editor and its saved file', async () => {
    const hub = new CollabHub()
    const a = await open(hub, 'Ann', () => 3)
    const b = await open(hub, 'Bob', () => 5)
    await settle([a.collab, b.collab])
    const ti = a.st.blocks.findIndex((x) => x.t === 'tbl')
    a.editor.input.lastSel = { a: { pa: `${ti}.1.0.0`, o: 0 }, f: { pa: `${ti}.1.2.0`, o: 0 } }
    expect(a.editor.mergeCells()).toBe(true)
    await settle([a.collab, b.collab])
    a.editor.input.lastSel = { a: { pa: `${ti}.2.0.0`, o: 0 }, f: { pa: `${ti}.2.1.0`, o: 0 } }
    a.editor.setCellShade('#ffcc00')
    a.editor.setCellBorders('outer', '1.5pt solid #ff0000')
    await settle([a.collab, b.collab])
    const ta = a.st.blocks[ti] as import('../src/core/docx/types').Table
    b.editor.setColumnWidths(ti, ta.cols.map((w, i) => (i === 0 ? w + 30 : i === 1 ? w - 30 : w)))
    await settle([a.collab, b.collab])

    for (const who of [a, b]) {
      const t = who.st.blocks[ti] as import('../src/core/docx/types').Table
      expect(t.rows[1].cells[0].colSpan).toBe(3)
      expect(t.rows[2].cells.slice(0, 2).map((c) => c.shd)).toEqual(['#ffcc00', '#ffcc00'])
      expect(t.rows[2].cells[0].borders.l).toBe('1.5pt solid #ff0000')
      expect(t.cols[0]).toBeCloseTo(ta.cols[0] + 30, 1)
    }
    const saved = await load((await b.client.call('save')) as Blob)
    const t = saved.st.blocks.find((x) => x.t === 'tbl') as import('../src/core/docx/types').Table
    expect(t.rows[1].cells[0].colSpan).toBe(3)
    expect(t.rows[2].cells[1].shd).toBe('#ffcc00')
    expect(t.rows[2].cells[1].borders.r).toBe('1.5pt solid #ff0000')
    for (const x of [a, b, saved]) x.client.destroy()
  })
})

describe('collab docx page setup', () => {
  test('a page setup change reaches the other editor, its file and a late joiner; undo goes back for everyone', async () => {
    const hub = new CollabHub()
    const a = await open(hub, 'Ann', () => 3)
    const b = await open(hub, 'Bob', () => 5)
    await settle([a.collab, b.collab])
    const seen: number[] = []
    b.editor.onPage = (p) => seen.push(p.w)
    await a.editor.setPage({ w: 841.9, h: 595.3, cols: 2 })
    await settle([a.collab, b.collab])
    expect(seen.at(-1)).toBeCloseTo(841.9, 1)
    const saved = (await b.client.call('save')) as Blob
    const re = await DocumentClient.open(saved, { worker: false })
    const meta = await new Promise<any>((r) => {
      re.on('docx:meta', (m) => m.final && r(m))
      void re.call('docx.start')
    })
    expect(meta.page.cols).toBe(2)
    expect(meta.page.w).toBeCloseTo(841.9, 1)

    const c = await open(hub, 'Cid', () => 0)
    await settle([a.collab, b.collab, c.collab])
    expect(c.editor.pageNow?.cols).toBe(2)

    a.editor.undo()
    await settle([a.collab, b.collab, c.collab])
    expect(seen.at(-1)).toBe(612)
    for (const x of [a, b, c]) x.client.destroy()
    re.destroy()
  })
})
