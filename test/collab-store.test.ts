import { describe, expect, test } from 'bun:test'
import { MemoryBlobStore, RoomStore } from '../src/collab/store'
import type { Batch } from '../src/collab/protocol'

const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms))
const INFO = { name: 'a.xlsx', type: 'application/octet-stream' }
const bytes = (s: string) => new TextEncoder().encode(s)
const batch = (seq: number, ops: unknown[] = [{ n: seq }]): Batch => ({ seq, site: 1, base: seq - 1, id: `b${seq}`, ops, time: 1000 + seq, user: 'ann' })
const keys = (blobs: MemoryBlobStore) => [...blobs.data.keys()].sort()

describe('RoomStore', () => {
  test('a room starts as snapshot 0 of the uploaded file; unknown rooms are empty', async () => {
    const store = new RoomStore(new MemoryBlobStore())
    expect(await store.head('x')).toBeNull()
    expect(await store.load('x')).toBeUndefined()
    expect(await store.file('x', 0)).toBeNull()
    const head = await store.create('x', bytes('v0'), INFO)
    expect(head).toMatchObject({ seq: 0, size: 2, ...INFO })
    expect(new TextDecoder().decode((await store.file('x', 0))!)).toBe('v0')
    expect(await store.load('x')).toMatchObject({ log: [], snapshot: { seq: 0, ctx: [] } })
    expect(await store.versions('x')).toEqual([{ seq: 0, size: 2, time: head.time }])
  })

  test('batches are buffered into segments and read back in order (binary ops too)', async () => {
    const blobs = new MemoryBlobStore()
    const store = new RoomStore(blobs, { flushMs: 5 })
    await store.create('x', bytes('v0'), INFO)
    for (let s = 1; s <= 3; s++) store.append('x', batch(s))
    await sleep(30)
    store.append('x', batch(4, [{ img: new Uint8Array([1, 2, 255]) }]))
    await store.flush()
    expect(keys(blobs).filter((k) => k.includes('/ops/')).map((k) => k.split('/ops/')[1])).toEqual(['000000000001-000000000003.jsonl', '000000000004-000000000004.jsonl'])
    const all = await store.batches('x', 0)
    expect(all.map((b) => b.seq)).toEqual([1, 2, 3, 4])
    expect(all[3].ops[0]).toEqual({ img: new Uint8Array([1, 2, 255]) })
    expect(all[0]).toMatchObject({ time: 1001, user: 'ann' })
    expect((await store.batches('x', 1, 3)).map((b) => b.seq)).toEqual([2, 3])
  })

  test('failed writes are retried without losing or reordering batches', async () => {
    const blobs = new MemoryBlobStore()
    let fail = 2
    const put = blobs.put.bind(blobs)
    blobs.put = async (k, d) => {
      if (k.includes('/ops/') && fail-- > 0) throw new Error('503 Slow Down')
      return put(k, d)
    }
    const errors: string[] = []
    const store = new RoomStore(blobs, { flushMs: 0, onError: (e) => errors.push(e.message) })
    await store.create('x', bytes('v0'), INFO)
    store.append('x', batch(1))
    await sleep(5)
    store.append('x', batch(2))
    for (let i = 0; i < 100 && (await store.batches('x', 0)).length < 2; i++) await sleep(50)
    expect(errors).toEqual(['503 Slow Down', '503 Slow Down'])
    expect((await store.batches('x', 0)).map((b) => b.seq)).toEqual([1, 2])
  })

  test('snapshots move the head; stale or older ones are refused', async () => {
    const store = new RoomStore(new MemoryBlobStore(), { flushMs: 0 })
    const h0 = await store.create('x', bytes('v0'), INFO)
    for (let s = 1; s <= 6; s++) store.append('x', batch(s))
    await store.flush()
    const ctx = [{ seq: 4, site: 1, ops: [{ n: 4 }] }]
    const h4 = await store.save('x', h0, { seq: 4, ctx, sites: 3 }, bytes('v4'))
    expect(h4).toMatchObject({ seq: 4, size: 2, epoch: h0.epoch })
    expect(await store.save('x', h0, { seq: 3, ctx: [] }, bytes('v3'))).toBeNull()
    expect(await store.snapshot('x', 4)).toMatchObject({ seq: 4, ctx, sites: 3 })
    // the room was replaced meanwhile: a snapshot of the old upload must not land
    const h1 = await store.create('x', bytes('new'), INFO)
    expect(await store.save('x', h4!, { seq: 6, ctx: [] }, bytes('v6'))).toBeNull()
    expect(await store.head('x')).toEqual(h1)
    expect(await store.batches('x', 0)).toEqual([])
    expect(await store.versions('x')).toEqual([{ seq: 0, size: 3, time: h1.time }])
  })

  test('load: latest snapshot + `retain` batches before it; a new upload deletes the old epoch', async () => {
    const blobs = new MemoryBlobStore()
    const store = new RoomStore(blobs, { flushMs: 0, retain: 2 })
    const h0 = await store.create('x', bytes('v0'), INFO)
    for (let s = 1; s <= 8; s++) {
      store.append('x', batch(s))
      await store.flush()
    }
    await store.save('x', h0, { seq: 5, ctx: [] }, bytes('v5'))
    const loaded = (await store.load('x'))!
    expect(loaded.snapshot?.seq).toBe(5)
    expect(loaded.log.map((b) => b.seq)).toEqual([4, 5, 6, 7, 8])
    const before = keys(blobs).length
    await store.create('x', bytes('again'), INFO)
    expect(before).toBeGreaterThan(5)
    expect(keys(blobs).filter((k) => !k.endsWith('head.json')).every((k) => k.includes((store as unknown as { heads: Map<string, { epoch: string }> }).heads.get('x')!.epoch))).toBe(true)
  })

  test('keep: older snapshots and the segments only they need are deleted', async () => {
    const blobs = new MemoryBlobStore()
    const store = new RoomStore(blobs, { flushMs: 0, keep: 2, retain: 1 })
    let head = await store.create('x', bytes('v0'), INFO)
    for (let s = 1; s <= 9; s++) {
      store.append('x', batch(s))
      await store.flush()
      if (s % 3 === 0) head = (await store.save('x', head, { seq: s, ctx: [] }, bytes(`v${s}`)))!
    }
    expect((await store.versions('x')).map((v) => v.seq)).toEqual([6, 9])
    expect(await store.file('x', 3)).toBeNull()
    // replaying from snapshot 6 still works
    expect((await store.batches('x', 6)).map((b) => b.seq)).toEqual([7, 8, 9])
    // batches up to the oldest kept snapshot are inside it
    expect((await store.batches('x', 0)).map((b) => b.seq)).toEqual([7, 8, 9])
  })

  test('room ids never escape the store prefix', async () => {
    const blobs = new MemoryBlobStore()
    const store = new RoomStore(blobs, { prefix: 'collab/' })
    await store.create('../../etc', bytes('x'), INFO)
    await store.create('a/b', bytes('y'), INFO)
    for (const k of keys(blobs)) {
      expect(k.startsWith('collab/')).toBe(true)
      expect(k.split('/')).not.toContain('..')
    }
    expect(new TextDecoder().decode((await store.file('a/b', 0))!)).toBe('y')
  })
})
