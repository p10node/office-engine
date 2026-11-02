import { describe, expect, test } from 'bun:test'
import { DocumentClient } from '../src/core/worker/client'
import { CollabHub } from '../src/collab/hub'
import { compact, startCollab, type Collab } from '../src/collab'
import { MemoryTransport } from '../src/collab/transport'
import { MemoryBlobStore, RoomStore } from '../src/collab/store'
import type { Batch } from '../src/collab/protocol'

const SAMPLE = `${import.meta.dir}/../playground/public/samples/sample.xlsx`
const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms))

interface Peer {
  client: DocumentClient
  collab: Collab
  transport: MemoryTransport
}

async function open(hub: CollabHub, bytes: Uint8Array, name: string, base = 0, latency: () => number = () => 0): Promise<Peer> {
  const client = await DocumentClient.open(new Blob([bytes as BlobPart]), { worker: false })
  await client.call('xlsx.load', 0)
  await sleep(30)
  const transport = new MemoryTransport(hub, latency)
  const collab = startCollab(client, { transport, room: 'x', user: { name }, base })
  return { client, collab, transport }
}

async function settle(ps: Peer[]) {
  for (let i = 0; i < 400; i++) {
    await sleep(10)
    await Promise.all(ps.map((p) => p.collab.session.idle()))
    if (ps.every((p) => p.collab.pending === 0 && p.collab.status === 'live') && new Set(ps.map((p) => p.collab.session.rev)).size === 1) return
  }
  throw new Error(`did not settle: ${ps.map((p) => `${p.collab.status}/${p.collab.session.rev}/${p.collab.pending}`).join(' ')}`)
}

/** shown text + styles of the first rows of every sheet */
async function dump(client: DocumentClient) {
  const out: unknown[] = []
  for (const s of client.info.xlsx!.sheets) {
    await client.call('xlsx.load', s.index)
    await sleep(20)
    const b = await client.call('xlsx.block', s.index, 0, 40)
    out.push(s.name, b.rows.map((r) => [r.r, r.cells.map((c) => [c[0], c[1], c[2]])]))
  }
  return JSON.stringify(out)
}

/** server side: materializes the room's head + batches into a new stored snapshot */
async function snapshot(hub: CollabHub, store: RoomStore) {
  const head = (await store.head('x'))!
  const from = (await store.snapshot('x', head.seq))!
  const batches = hub.rooms.get('x')!.log.filter((b) => b.seq > head.seq)
  const client = await DocumentClient.open(new Blob([(await store.file('x', head.seq))! as BlobPart]), { worker: false })
  const snap = await compact(client, { from, batches })
  const file = new Uint8Array(await (await client.call('save')).arrayBuffer())
  client.destroy()
  const next = await store.save('x', head, { ...snap, sites: hub.rooms.get('x')!.nextSite }, file)
  expect(next?.seq).toBe(snap.seq)
  expect(hub.snapshot('x', snap)).toBe(true)
  return snap.seq
}

/** random cell edits / formatting / row structure by `who` */
function edits(who: Peer[], n: number, rand: () => number, tag: string) {
  const ops: Promise<unknown>[] = []
  for (let i = 0; i < n; i++) {
    const p = who[Math.floor(rand() * who.length)]
    const r = Math.floor(rand() * 12)
    const c = Math.floor(rand() * 5)
    const x = rand()
    let call: Promise<unknown>
    if (x < 0.6) call = p.client.call('xlsx.set', 0, [{ r, c, input: rand() < 0.3 ? `=A${r + 2}&"!"` : `${tag}${i}` }])
    else if (x < 0.8) call = p.client.call('xlsx.format', 0, { r0: r, c0: c, r1: r, c1: c + 1 }, { b: rand() < 0.5 })
    else if (x < 0.9) call = p.client.call('xlsx.sheetOp', 0, { op: rand() < 0.6 ? 'insertRows' : 'deleteRows', at: r, n: 1 })
    else call = p.client.call('xlsx.undo')
    ops.push(call.catch(() => {}))
  }
  return Promise.all(ops)
}

describe('collab snapshots', () => {
  for (const start of [3, 17, 51]) test(`late batches, joining from snapshots, restart and history replay converge (seed ${start})`, async () => {
    let seed = start
    const rand = () => ((seed = (seed * 1664525 + 1013904223) % 4294967296) / 4294967296)
    const bytes = new Uint8Array(await Bun.file(SAMPLE).arrayBuffer())
    const store = new RoomStore(new MemoryBlobStore(), { flushMs: 0, retain: 5 })
    const hooks = { onBatch: (r: string, b: Batch) => store.append(r, b), loadLog: (r: string) => store.load(r) }
    const hub = new CollabHub(hooks)
    await store.create('x', bytes, { name: 'sample.xlsx', type: 'application/octet-stream' })

    const lat = () => Math.floor(rand() * 4)
    const a = await open(hub, bytes, 'Ann', 0, lat)
    const b = await open(hub, bytes, 'Bob', 0, lat)
    await settle([a, b])
    await edits([a, b], 30, rand, 'p')
    await settle([a, b])

    // Ann goes offline and keeps editing: her batch will be based on a state older than the snapshot
    a.transport.drop()
    // structural ops wait for the hub's order: they resolve after reconnecting
    const offline = edits([a], 6, rand, 'off')
    await sleep(20)
    await edits([b], 10, rand, 'b')
    await settle([b])
    const s1 = await snapshot(hub, store)
    expect(s1).toBe(b.collab.session.rev)
    a.transport.connect()
    await edits([b], 4, rand, 'c')
    await offline
    await settle([a, b])
    expect(hub.rooms.get('x')!.log.some((x) => x.seq > s1 && x.base < s1)).toBe(true)

    // a newcomer opens the snapshot file and only replays the batches after it (ctx merges Ann's late batch)
    const c = await open(hub, (await store.file('x', s1))!, 'Cid', s1, lat)
    await settle([a, b, c])
    await edits([a, b, c], 30, rand, 'q')
    await settle([a, b, c])
    const want = await dump(a.client)
    expect(want.match(/"[pqbc]\d+"/g)?.length ?? 0).toBeGreaterThan(5)
    expect(await dump(b.client)).toBe(want)
    expect(await dump(c.client)).toBe(want)

    const s2 = await snapshot(hub, store)

    // hub restart: the room is rebuilt from storage (latest snapshot + `retain` batches before it)
    await store.flush()
    const hub2 = new CollabHub({ ...hooks, retain: 2, keep: 1 })
    const d = await open(hub2, (await store.file('x', s2))!, 'Dan', s2, lat)
    await settle([d])
    expect(hub2.rooms.get('x')!.start).toBe(Math.max(0, s2 - 5))
    expect(d.collab.session.rev).toBe(a.collab.session.rev)
    expect(await dump(d.client)).toBe(want)

    // a newer snapshot drops older history from memory: joining from a dropped point is refused
    await edits([d], 12, rand, 'd')
    await settle([d])
    const s3 = await snapshot(hub2, store)
    expect(hub2.rooms.get('x')!.start).toBe(s3 - 2)
    const stale = await open(hub2, bytes, 'Old', 0)
    for (let i = 0; i < 50 && stale.collab.status !== 'error'; i++) await sleep(10)
    expect(stale.collab.status).toBe('error')
    const final = await dump(d.client)

    // version history: any seq = the original upload + the stored batches up to it
    const versions = await store.versions('x')
    expect(versions.map((v) => v.seq)).toEqual([0, s1, s2, s3])
    const all = await store.batches('x', 0)
    expect(all.map((x) => x.seq)).toEqual(Array.from({ length: s3 }, (_, i) => i + 1))
    expect(all.every((x) => typeof x.time === 'number' && typeof x.user === 'string')).toBe(true)
    const replay = await DocumentClient.open(new Blob([bytes as BlobPart]), { worker: false })
    expect((await compact(replay, { batches: all })).seq).toBe(s3)
    expect(await dump(replay)).toBe(final)
    const mid = await DocumentClient.open(new Blob([bytes as BlobPart]), { worker: false })
    await compact(mid, { batches: all.filter((x) => x.seq <= s2) })
    expect(await dump(mid)).toBe(want)
    mid.destroy()

    for (const p of [a, b, c, d, stale]) {
      p.collab.close()
      p.client.destroy()
    }
    replay.destroy()
  }, 60_000)
})

describe('CollabHub rooms', () => {
  const user = { name: 'Ann' }
  const tick = () => sleep(5)
  function peer(hub: CollabHub) {
    const got: any[] = []
    const conn = hub.connect((m) => got.push(m))
    return { got, conn }
  }
  const log = (n: number, from = 1): Batch[] => Array.from({ length: n }, (_, i) => ({ seq: from + i, site: 1, base: from + i - 1, id: `x${from + i}`, ops: [] }))

  test('messages sent while the room loads are handled after it, in order', async () => {
    let release!: () => void
    const hub = new CollabHub({ loadLog: () => new Promise((r) => (release = () => r(log(2)))) })
    const a = peer(hub)
    a.conn.receive({ t: 'hello', room: 'r', user })
    a.conn.receive({ t: 'ops', id: 'o1', base: 2, ops: ['first'] })
    a.conn.receive({ t: 'ops', id: 'o2', base: 2, ops: ['second'] })
    await tick()
    expect(a.got).toEqual([])
    release()
    await tick()
    expect(a.got.map((m) => m.t)).toEqual(['welcome', 'batch', 'batch'])
    expect(a.got[0]).toMatchObject({ seq: 2, from: 0, reset: false })
    expect(a.got.slice(1).map((m) => [m.b.seq, m.b.ops[0]])).toEqual([[3, 'first'], [4, 'second']])
  })

  test('a reset while the room loads makes waiting clients get the new room', async () => {
    let calls = 0
    const results = [log(5), []]
    const hub = new CollabHub({ loadLog: () => sleep(10).then(() => results[calls++]) })
    const a = peer(hub)
    a.conn.receive({ t: 'hello', room: 'r', user })
    hub.reset('r')
    await sleep(40)
    expect(calls).toBe(2)
    expect(a.got[0]).toMatchObject({ t: 'welcome', seq: 0, log: [] })
    expect(hub.rooms.get('r')?.peers.size).toBe(1)
  })

  test('idle rooms unload and load again on the next join', async () => {
    let calls = 0
    const hub = new CollabHub({ loadLog: () => (calls++, log(3)) })
    const a = peer(hub)
    a.conn.receive({ t: 'hello', room: 'r', user })
    await tick()
    expect(hub.unload('r')).toBe(false)
    a.conn.close()
    expect(hub.unload('r')).toBe(true)
    const b = peer(hub)
    b.conn.receive({ t: 'hello', room: 'r', user })
    await tick()
    expect(calls).toBe(2)
    expect(b.got[0]).toMatchObject({ t: 'welcome', seq: 3 })
  })

  test('snapshots: unknown or future ones are refused; joins must start from a known point', async () => {
    const hub = new CollabHub({ loadLog: () => log(10), retain: 3, keep: 2 })
    await hub.room('r')
    expect(hub.snapshot('nope', { seq: 1, ctx: [] })).toBe(false)
    expect(hub.snapshot('r', { seq: 11, ctx: [] })).toBe(false)
    const ctx = [{ seq: 8, site: 1, ops: [] }]
    expect(hub.snapshot('r', { seq: 8, ctx })).toBe(true)
    expect(hub.rooms.get('r')!.start).toBe(5)

    const join = async (msg: object) => {
      const p = peer(hub)
      p.conn.receive({ t: 'hello', room: 'r', user, ...msg })
      await tick()
      return p.got[0]
    }
    expect(await join({ base: 8 })).toMatchObject({ reset: false, from: 8, ctx, log: log(2, 9) })
    expect(await join({ base: 7 })).toMatchObject({ reset: true, log: [] })
    expect(await join({})).toMatchObject({ reset: true })
    expect(await join({ have: 6, site: 1 })).toMatchObject({ reset: false, from: 6 })
    expect(await join({ have: 4, site: 1 })).toMatchObject({ reset: true })
    expect(await join({ have: 11, site: 1 })).toMatchObject({ reset: true })
  })
})
