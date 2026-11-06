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
  const collab = startCollab(client, { transport: new MemoryTransport(hub, latency), room: 'd', user: { name } })
  return { client, st, editor, collab }
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
})
