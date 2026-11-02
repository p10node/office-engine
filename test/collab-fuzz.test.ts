/**
 * Deterministic fuzzing: the hub's messages wait in queues and the test decides when each one is
 * delivered, so any interleaving of local edits and remote batches is reproducible from its seed.
 * DIVERGE_SEEDS=1-400 widens the search.
 */
import { describe, expect, test } from 'bun:test'
import { DocumentClient } from '../src/core/worker/client'
import type { Paragraph } from '../src/core/docx/types'
import { CollabHub, type HubConnection } from '../src/collab/hub'
import { registerHost } from '../src/collab/hosts'
import { startCollab } from '../src/collab'
import type { Transport } from '../src/collab/transport'
import type { ClientMsg, HubMsg } from '../src/collab/protocol'
import { decode, encode } from '../src/collab/codec'
import { DocxEditor, type DocState } from '../src/react/docx/DocxEditor'
import { opBackspace, opDelete, opEnter, opInsert } from '../src/react/edit/input'
import { plen, runsText, type RRun } from '../src/react/edit/rich'
import { ZipArchive } from '../src/core/zip'
import { openSource } from '../src/core/source'

const S = (n: string) => `${import.meta.dir}/../playground/public/samples/${n}`

/** messages wait in queues until the test delivers them */
class Wire {
  queues: { deliver: () => void; key: string }[][] = []
  pending() {
    return this.queues.filter((q) => q.length)
  }
}

class ManualTransport implements Transport {
  onMessage?: (m: HubMsg) => void
  onStatus?: (s: 'connecting' | 'open' | 'closed') => void
  private conn?: HubConnection
  private up: { deliver: () => void; key: string }[] = []
  private down: { deliver: () => void; key: string }[] = []
  constructor(private hub: CollabHub, wire: Wire) {
    wire.queues.push(this.up, this.down)
  }
  connect() {
    this.conn = this.hub.connect((m) => {
      const w = encode(m)
      this.down.push({ key: 'down ' + m.t, deliver: () => this.onMessage?.(decode(w)) })
    })
    this.onStatus?.('open')
  }
  send(m: ClientMsg) {
    const w = encode(m)
    this.up.push({ key: 'up ' + m.t, deliver: () => this.conn!.receive(decode(w)) })
  }
  close() {}
}

async function load(src: Blob) {
  const client = await DocumentClient.open(src, { worker: false })
  const st: DocState = { blocks: [], heights: new WeakMap(), dirty: new Set(), frontier: 0 }
  client.on('docx:blocks', (b) => st.blocks.push(...b))
  const done = new Promise<void>((r) => client.on('docx:done', () => r()))
  void client.call('docx.start')
  await done
  return { client, st }
}

const text = (st: DocState) => st.blocks.flatMap((b) => (b.t === 'p' ? [`${b.pid}:${runsText(b.runs as RRun[])}`] : b.t === 'tbl' ? ['T' + b.tid] : []))

async function runSeed(start: number, steps: number, verbose = false): Promise<string | null> {
  let seed = start
  const rand = () => ((seed = (seed * 1664525 + 1013904223) % 4294967296) / 4294967296)
  const hub = new CollabHub()
  const wire = new Wire()
  const mk = async (name: string) => {
    const { client, st } = await load(Bun.file(S('sample.docx')) as unknown as Blob)
    const editor = new DocxEditor(client, st, () => {})
    registerHost(client, 'docx', { editor, st })
    const collab = startCollab(client, { transport: new ManualTransport(hub, wire), room: 'd', user: { name } })
    return { name, client, st, editor, collab }
  }
  const a = await mk('A')
  const b = await mk('B')
  const idle = async () => {
    for (let k = 0; k < 3; k++) {
      await Promise.all([a.collab.session.idle(), b.collab.session.idle()])
      await new Promise((r) => setTimeout(r, 0))
    }
  }
  const flushAll = async () => {
    for (let guard = 0; guard < 10000; guard++) {
      await idle()
      const q = wire.pending()
      if (!q.length) return
      q[0].shift()!.deliver()
    }
  }
  await flushAll()
  const log: string[] = []
  for (let i = 0; i < steps; i++) {
    if (rand() < 0.45 && wire.pending().length) {
      const q = wire.pending()
      const m = q[Math.floor(rand() * q.length)].shift()!
      log.push('deliver ' + m.key)
      m.deliver()
      await idle()
      continue
    }
    const who = rand() < 0.5 ? a : b
    const paras = who.st.blocks.map((b, i) => [b, i] as const).filter(([b]) => b.t === 'p' && b.pid !== undefined && !b.ro) as [Paragraph, number][]
    const [p, bi] = paras[Math.floor(rand() * Math.min(paras.length, 4))]
    const pa = String(bi)
    const n = plen(p.runs as RRun[])
    const o = Math.floor(rand() * (n + 1))
    const x = rand()
    const at = (o: number) => ({ a: { pa, o }, f: { pa, o } })
    const w = who.name
    if (x < 0.45) { log.push(`${w} ins ${p.pid}@${o}`); opInsert(who.editor, at(o), w + i) }
    else if (x < 0.6 && n > 0) { log.push(`${w} del ${p.pid}@${o}`); opDelete(who.editor, { a: { pa, o: Math.max(0, o - 2) }, f: { pa, o } }) }
    else if (x < 0.7) { log.push(`${w} enter ${p.pid}@${o}`); opEnter(who.editor, at(o)) }
    else if (x < 0.77 && bi > 0) { log.push(`${w} join ${p.pid}`); opBackspace(who.editor, at(0)) }
    else if (x < 0.85 && n > 1) { log.push(`${w} bold ${p.pid}`); who.editor.input.lastSel = { a: { pa, o: 0 }, f: { pa, o: Math.min(n, 3) } }; who.editor.toggle('b') }
    else if (x < 0.89) { log.push(`${w} table ${p.pid}`); who.editor.input.lastSel = at(o); who.editor.insertTable(2, 2) }
    else if (x < 0.92) { const ti = who.st.blocks.findIndex((b) => b.t === 'tbl'); log.push(`${w} cell ${ti}`); if (ti >= 0) opInsert(who.editor, { a: { pa: `${ti}.0.0.0`, o: 0 }, f: { pa: `${ti}.0.0.0`, o: 0 } }, 'c' + i) }
    else if (x < 0.96) { log.push(`${w} undo`); who.editor.undo() }
    else { log.push(`${w} redo`); who.editor.redo() }
    await idle()
  }
  await flushAll()
  const ta = text(a.st)
  const tb = text(b.st)
  a.collab.close()
  b.collab.close()
  a.client.destroy()
  b.client.destroy()
  if (ta.join('\n') === tb.join('\n')) return null
  if (verbose) {
    console.log(log.slice(-12).join('\n'))
    for (let k = 0; k < Math.max(ta.length, tb.length); k++) if (ta[k] !== tb[k]) console.log(k, 'A=', ta[k], '  B=', tb[k])
    const fmt = (o: any) => o.t === 'text' ? `text ${o.pid}@${o.at}-${o.del}+${o.ins.map((r: any) => r.text ?? '#').join('')}` : o.t === 'paras' ? `paras [${o.before}] -> [${o.next.map((p: any) => p.pid + ':' + runsText(p.runs))}]` : o.t === 'blocks' ? `blocks [${o.before.map((k: any) => JSON.stringify(k))}] -> [${o.next.map((b: any) => (b.t === 'p' ? 'p' + b.pid : 'T' + b.tid))}]` : o.t
    for (const bt of hub.rooms.get('d')!.log.slice(-8)) console.log('batch', bt.seq, 'site', bt.site, 'base', bt.base, bt.ops.map(fmt).join(' ; '))
  }
  return 'diverged'
}


async function dump(client: DocumentClient) {
  const meta = client.info.xlsx!
  const out: unknown[] = []
  for (const s of meta.sheets) {
    await client.call('xlsx.load', s.index)
    await new Promise((r) => setTimeout(r, 20))
    const b = await client.call('xlsx.block', s.index, 0, 40)
    out.push(s.name, b.rows.map((r) => [r.r, r.cells.map((c) => [c[0], c[1], c[2]])]))
  }
  const zip = await ZipArchive.open(await openSource(await client.call('save')))
  out.push(new TextDecoder().decode(await zip.bytes('xl/styles.xml')))
  return JSON.stringify(out)
}
async function runSeedX(start: number, steps: number, verbose = false) {
  let seed = start
  const rand = () => ((seed = (seed * 1664525 + 1013904223) % 4294967296) / 4294967296)
  const hub = new CollabHub()
  const wire = new Wire()
  const mk = async (name: string) => {
    const client = await DocumentClient.open(Bun.file(S('sample.xlsx')) as unknown as Blob, { worker: false })
    await client.call('xlsx.load', 0)
    await new Promise((r) => setTimeout(r, 30))
    const collab = startCollab(client, { transport: new ManualTransport(hub, wire), room: 'x', user: { name } })
    return { name, client, collab }
  }
  const a = await mk('A')
  const b = await mk('B')
  const inflight: Promise<unknown>[] = []
  const idle = async () => {
    for (let k = 0; k < 4; k++) {
      await Promise.all([a.collab.session.idle(), b.collab.session.idle()])
      await new Promise((r) => setTimeout(r, 0))
    }
  }
  const flushAll = async () => {
    for (let g = 0; g < 10000; g++) {
      await idle()
      const q = wire.pending()
      if (!q.length) return
      q[0].shift()!.deliver()
    }
  }
  await flushAll()
  const log: string[] = []
  for (let i = 0; i < steps; i++) {
    if (rand() < 0.45 && wire.pending().length) {
      const q = wire.pending()
      const m = q[Math.floor(rand() * q.length)].shift()!
      log.push('deliver ' + m.key)
      m.deliver()
      await idle()
      continue
    }
    const who = rand() < 0.5 ? a : b
    const x = rand()
    const r = Math.floor(rand() * 12)
    const c = Math.floor(rand() * 5)
    const run = (label: string, p: Promise<unknown>) => { log.push(`${who.name} ${label}`); inflight.push(p.catch(() => {})) }
    if (x < 0.45) run(`set ${r},${c}`, who.client.call('xlsx.set', 0, [{ r, c, input: rand() < 0.3 ? `=A${r + 2}*2` : `v${i}` }]))
    else if (x < 0.65) { const patch = rand() < 0.5 ? { b: rand() < 0.5 } : { fill: rand() < 0.5 ? '#ffee00' : '#00ccff' }; run(`fmt ${r},${c} ${JSON.stringify(patch)}`, who.client.call('xlsx.format', 0, { r0: r, c0: c, r1: r + 1, c1: c + 1 }, patch)) }
    else if (x < 0.72) { const op = rand() < 0.6 ? 'insertRows' : 'deleteRows'; run(`${op} ${r}`, who.client.call('xlsx.sheetOp', 0, { op, at: r, n: 1 } as never)) }
    else if (x < 0.8) run('undo', who.client.call('xlsx.undo'))
    else if (x < 0.85) run('redo', who.client.call('xlsx.redo'))
    else if (x < 0.92) run(`colWidth ${c}`, who.client.call('xlsx.sheetOp', 0, { op: 'colWidth', c0: c, c1: c, px: 40 + i }))
    else run(`fill ${r},${c}`, who.client.call('xlsx.sheetOp', 0, { op: 'fill', range: { r0: r, c0: c, r1: r + 2, c1: c }, dir: 'down' }))
    await idle()
  }
  await flushAll()
  await Promise.all(inflight)
  await flushAll()
  const da = await dump(a.client)
  const db = await dump(b.client)
  a.collab.close(); b.collab.close(); a.client.destroy(); b.client.destroy()
  if (da === db) return null
  if (verbose) {
    console.log(log.slice(-14).join('\n'))
    const A = JSON.parse(da), B = JSON.parse(db)
    for (let k = 0; k < A.length; k++) if (JSON.stringify(A[k]) !== JSON.stringify(B[k])) {
      if (typeof A[k] === 'string') { let p = 0; while (A[k][p] === B[k][p]) p++; console.log('diff part', k, 'A:', A[k].slice(p - 80, p + 200), '\nB:', B[k].slice(p - 80, p + 200)) }
      else { const ra = A[k], rb = B[k]; for (let j = 0; j < Math.max(ra.length, rb.length); j++) if (JSON.stringify(ra[j]) !== JSON.stringify(rb[j])) console.log('row', JSON.stringify(ra[j]), '\n   ', JSON.stringify(rb[j])) }
    }
    const fmt = (o: any) => o.t === 'set' ? `set ${o.cells.map((c: any) => c.r + ',' + c.c + '=' + c.input)}` : o.t === 'fmt' ? `fmt ${JSON.stringify(o.range)} ${JSON.stringify(o.patch)}` : o.t === 'sheet' ? `sheet ${JSON.stringify(o.op)}` : o.t === 'revert' ? `revert ${o.target} ${o.which}` : o.t
    for (const bt of hub.rooms.get('x')!.log) console.log('batch', bt.seq, 'site', bt.site, 'base', bt.base, bt.ops.map((o: any) => o.id + ' ' + fmt(o)).join(' ; '))

  }
  return 'diverged'
}

const [lo, hi] = (process.env.DIVERGE_SEEDS ?? '1-12').split('-').map(Number)

describe('collab fuzz (deterministic interleavings)', () => {
  test('docx', async () => {
    for (let s = lo; s <= hi; s++) expect(`seed ${s}: ${await runSeed(s, 60, true)}`).toBe(`seed ${s}: null`)
  }, 600_000)
  test('xlsx', async () => {
    for (let s = lo; s <= hi; s++) expect(`seed ${s}: ${await runSeedX(s, 50, true)}`).toBe(`seed ${s}: null`)
  }, 600_000)
})
