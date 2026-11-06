import { describe, expect, test } from 'bun:test'
import { CollabHub } from '../src/collab/hub'
import { CollabSession, type Adapter } from '../src/collab/session'
import { MemoryTransport } from '../src/collab/transport'
import { decode, encode } from '../src/collab/codec'

type TOp = { t: 'ins'; at: number; s: string } | { t: 'del'; at: number; n: number } | { t: 'set'; k: string; v: number } | { t: 'clear' }

/** toy document: a string + registers; `clear` is pessimistic */
class TextAdapter implements Adapter<TOp> {
  text = ''
  regs: Record<string, number> = {}
  transform(op: TOp, p: TOp): TOp | null {
    if (p.t === 'clear') return op.t === 'set' ? op : null
    if (op.t === 'set' || op.t === 'clear') return op
    if (p.t === 'ins') {
      if (op.t === 'ins') return op.at >= p.at ? { ...op, at: op.at + p.s.length } : op
      if (op.at >= p.at) return { ...op, at: op.at + p.s.length }
      if (op.at + op.n <= p.at) return op
      return { ...op, n: op.n + p.s.length }
    }
    if (p.t === 'del') {
      const pe = p.at + p.n
      if (op.t === 'ins') return op.at <= p.at ? op : op.at >= pe ? { ...op, at: op.at - p.n } : { ...op, at: p.at }
      const s = op.at
      const e = op.at + op.n
      const ns = s <= p.at ? s : s >= pe ? s - p.n : p.at
      const ne = e <= p.at ? e : e >= pe ? e - p.n : p.at
      return ne > ns ? { ...op, at: ns, n: ne - ns } : null
    }
    return op
  }
  apply(op: TOp) {
    const undo = { text: this.text, regs: { ...this.regs } }
    if (op.t === 'ins') this.text = this.text.slice(0, op.at) + op.s + this.text.slice(op.at)
    else if (op.t === 'del') this.text = this.text.slice(0, op.at) + this.text.slice(op.at + op.n)
    else if (op.t === 'set') this.regs[op.k] = op.v
    else this.text = ''
    return { undo, result: this.text.length }
  }
  revert(u: unknown) {
    const s = u as { text: string; regs: Record<string, number> }
    this.text = s.text
    this.regs = s.regs
  }
  pessimistic(op: TOp) {
    return op.t === 'clear'
  }
}

function rng(seed: number) {
  return () => {
    seed = (seed * 1664525 + 1013904223) % 4294967296
    return seed / 4294967296
  }
}

const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms))

async function settle(sessions: CollabSession<TOp>[]) {
  for (let i = 0; i < 200; i++) {
    await sleep(15)
    await Promise.all(sessions.map((s) => s.idle()))
    if (sessions.every((s) => s.pending === 0 && s.status === 'live') && new Set(sessions.map((s) => s.rev)).size === 1) return
  }
  throw new Error('did not settle')
}

describe('collab', () => {
  test('codec keeps binary payloads', () => {
    const v = decode<{ a: Uint8Array; b: string }>(encode({ a: new Uint8Array([0, 255, 7]), b: 'x' }))
    expect([...v.a]).toEqual([0, 255, 7])
    expect(v.b).toBe('x')
  })

  for (const seed of [1, 2, 3, 4, 5]) {
    test(`random concurrent edits converge (seed ${seed})`, async () => {
      const rand = rng(seed)
      const hub = new CollabHub()
      const clients = [0, 1, 2].map((i) => {
        const adapter = new TextAdapter()
        const transport = new MemoryTransport(hub, () => Math.floor(rand() * 8))
        const session = new CollabSession<TOp>(transport, adapter, { room: 'r', user: { name: `u${i}` } })
        return { adapter, transport, session }
      })
      await settle(clients.map((c) => c.session))
      for (let step = 0; step < 300; step++) {
        const c = clients[Math.floor(rand() * clients.length)]
        const len = c.adapter.text.length
        const x = rand()
        let op: TOp
        if (x < 0.5 || len === 0) op = { t: 'ins', at: Math.floor(rand() * (len + 1)), s: String.fromCharCode(97 + Math.floor(rand() * 26)) }
        else if (x < 0.8) {
          const at = Math.floor(rand() * len)
          op = { t: 'del', at, n: 1 + Math.floor(rand() * Math.min(3, len - at)) }
        } else if (x < 0.98) op = { t: 'set', k: 'k' + Math.floor(rand() * 3), v: step }
        else op = { t: 'clear' }
        void c.session.submit(op)
        if (rand() < 0.02) {
          // connection drop + automatic reconnect
          c.transport.drop()
          setTimeout(() => c.transport.connect(), 5)
        }
        if (rand() < 0.3) await sleep(Math.floor(rand() * 4))
      }
      await settle(clients.map((c) => c.session))
      const texts = clients.map((c) => c.adapter.text)
      expect(new Set(texts).size).toBe(1)
      expect(clients[0].session.rev).toBeGreaterThan(10)
      expect(new Set(clients.map((c) => JSON.stringify(c.adapter.regs))).size).toBe(1)

      // a late joiner replays the log and reaches the same state
      const late = new TextAdapter()
      const ls = new CollabSession<TOp>(new MemoryTransport(hub), late, { room: 'r', user: { name: 'late' } })
      await settle([ls, ...clients.map((c) => c.session)])
      expect(late.text).toBe(texts[0])
      for (const c of clients) c.session.close()
      ls.close()
    })
  }

  test('an operation that cannot be prepared is skipped: the room goes on, late joiners included', async () => {
    /** `set bad` (once shared) and `clear` (pessimistic) can be neither prepared nor applied */
    class Flaky extends TextAdapter {
      private bad = (op: TOp) => (op.t === 'set' && op.k === 'bad') || op.t === 'clear'
      prepare(ops: TOp[]) {
        if (ops.some(this.bad)) throw new Error('cannot prepare')
      }
      override apply(op: TOp, kind = 'local') {
        if (this.bad(op) && (kind !== 'local' || op.t === 'clear')) throw new Error('cannot apply')
        return super.apply(op)
      }
    }
    const error = console.error
    console.error = () => {}
    try {
      const hub = new CollabHub()
      const open = (name: string) => {
        const adapter = new Flaky()
        return { adapter, session: new CollabSession<TOp>(new MemoryTransport(hub, () => 1), adapter, { room: 'f', user: { name } }) }
      }
      const a = open('Ann')
      const b = open('Bob')
      await settle([a.session, b.session])
      // fails where it arrives as a remote operation…
      await a.session.submit({ t: 'set', k: 'bad', v: 1 })
      await a.session.submit({ t: 'ins', at: 0, s: 'one' })
      await settle([a.session, b.session])
      // …and for its author, once the hub has ordered it
      expect(await a.session.submit({ t: 'clear' })).toBeNull()
      await settle([a.session, b.session])
      await b.session.submit({ t: 'ins', at: 3, s: ' two' })
      await settle([a.session, b.session])
      const c = open('Cid')
      await settle([a.session, b.session, c.session])
      for (const x of [a, b, c]) expect(x.adapter.text).toBe('one two')
      expect(a.session.rev).toBe(4)
      for (const x of [a, b, c]) x.session.close()
    } finally {
      console.error = error
    }
  })

  test('presence and peers', async () => {
    const hub = new CollabHub()
    const a = new CollabSession<TOp>(new MemoryTransport(hub), new TextAdapter(), { room: 'p', user: { name: 'Ann' } })
    const b = new CollabSession<TOp>(new MemoryTransport(hub), new TextAdapter(), { room: 'p', user: { name: 'Bob' } })
    await settle([a, b])
    a.setPresence({ at: 3 })
    await sleep(120)
    const ann = [...b.peers.values()].find((p) => p.name === 'Ann')
    expect(ann?.presence).toEqual({ at: 3 })
    expect(ann?.color).toBeTruthy()
    a.close()
    await sleep(20)
    expect([...b.peers.values()].some((p) => p.name === 'Ann')).toBe(false)
    b.close()
  })
})
