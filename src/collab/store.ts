/**
 * Durable rooms on any key → bytes storage (S3, GCS, a directory, memory):
 *
 *   {room}/head.json                          latest snapshot: { epoch, seq, size, name, type, time }
 *   {room}/{epoch}/snap/{seq}.bin             the document at `seq` (seq 0: the uploaded file)
 *   {room}/{epoch}/snap/{seq}.json / .ctx     its metadata / merge context
 *   {room}/{epoch}/ops/{first}-{last}.jsonl   sequenced batches, one segment per flush
 *
 * Saving an edit appends a small segment; the document itself is rewritten only when a snapshot
 * is taken (`compact`). Snapshots and segments together are the version history: any `seq` is
 * a snapshot plus the batches after it. Uploading a new file starts a new epoch.
 */
import { decode, encode } from './codec'
import type { RoomLog } from './hub'
import type { Batch, Entry, Snapshot } from './protocol'

/** Minimal object storage. Keys use `/` separators. */
export interface BlobStore {
  /** null when missing */
  get(key: string): Promise<Uint8Array | null>
  put(key: string, data: Uint8Array): Promise<void>
  /** every key starting with `prefix` (any order) */
  list(prefix: string): Promise<string[]>
  delete(key: string): Promise<void>
}

export class MemoryBlobStore implements BlobStore {
  readonly data = new Map<string, Uint8Array>()
  async get(key: string) {
    return this.data.get(key) ?? null
  }
  async put(key: string, data: Uint8Array) {
    this.data.set(key, data.slice())
  }
  async list(prefix: string) {
    return [...this.data.keys()].filter((k) => k.startsWith(prefix))
  }
  async delete(key: string) {
    this.data.delete(key)
  }
}

export interface RoomHead {
  /** changes with every upload: stale writes of an older upload are ignored */
  epoch: string
  /** latest snapshot */
  seq: number
  size: number
  name: string
  type: string
  time: number
}

/** a stored snapshot */
export interface Version {
  seq: number
  size: number
  time: number
}

export interface RoomStoreOptions {
  /** key prefix, e.g. `collab/` */
  prefix?: string
  /** batches are buffered and written as one segment at most this often (ms, default 250) */
  flushMs?: number
  /** snapshots kept per room (default: all — the version history); older ones and their batches are deleted */
  keep?: number
  /** batches below the latest snapshot loaded for clients reconnecting from older states (default 1000) */
  retain?: number
  /** a write failed (it is retried) */
  onError?: (error: Error, room: string) => void
}

type SnapMeta = Version & { sites?: number }

const pad = (n: number) => String(n).padStart(12, '0')
const enc = new TextEncoder()
const dec = new TextDecoder()
const json = (v: unknown) => enc.encode(JSON.stringify(v))
const parse = <T>(b: Uint8Array | null): T | null => (b ? (JSON.parse(dec.decode(b)) as T) : null)

export class RoomStore {
  private heads = new Map<string, RoomHead | null>()
  private buffers = new Map<string, { batches: Batch[]; timer?: ReturnType<typeof setTimeout>; retry: number }>()
  private queues = new Map<string, Promise<unknown>>()

  constructor(
    readonly blobs: BlobStore,
    private opts: RoomStoreOptions = {},
  ) {}

  private dir(room: string) {
    // dots escaped too: no `..` segments in keys used as file paths
    return `${this.opts.prefix ?? ''}${encodeURIComponent(room).replace(/\./g, '%2E')}/`
  }

  private epochDir(room: string, epoch: string) {
    return `${this.dir(room)}${epoch}/`
  }

  /** runs `fn` after every earlier operation on `room` */
  private run<T>(room: string, fn: () => Promise<T>): Promise<T> {
    const p = (this.queues.get(room) ?? Promise.resolve()).then(fn, fn)
    const tail = p.catch(() => {})
    this.queues.set(room, tail)
    void tail.then(() => this.queues.get(room) === tail && this.queues.delete(room))
    return p
  }

  private async readHead(room: string): Promise<RoomHead | null> {
    if (this.heads.has(room)) return this.heads.get(room)!
    const h = parse<RoomHead>(await this.blobs.get(this.dir(room) + 'head.json'))
    this.heads.set(room, h)
    return h
  }

  /** sets a room's document (snapshot 0 of a new epoch) and forgets its history */
  create(room: string, file: Uint8Array, info: { name: string; type: string }): Promise<RoomHead> {
    const buf = this.buffers.get(room)
    if (buf) {
      clearTimeout(buf.timer)
      this.buffers.delete(room)
    }
    return this.run(room, async () => {
      const old = await this.readHead(room)
      const head: RoomHead = { epoch: Date.now().toString(36) + Math.random().toString(36).slice(2, 6), seq: 0, size: file.length, time: Date.now(), ...info }
      const d = this.epochDir(room, head.epoch)
      await this.blobs.put(`${d}snap/${pad(0)}.bin`, file)
      await this.blobs.put(`${d}snap/${pad(0)}.json`, json({ seq: 0, size: file.length, time: head.time } satisfies SnapMeta))
      await this.blobs.put(this.dir(room) + 'head.json', json(head))
      this.heads.set(room, head)
      if (old) for (const k of await this.blobs.list(this.epochDir(room, old.epoch))) await this.blobs.delete(k)
      return head
    })
  }

  /** latest snapshot of the room (null: no such room) */
  head(room: string): Promise<RoomHead | null> {
    return this.run(room, () => this.readHead(room))
  }

  /** the document of snapshot `seq` */
  async file(room: string, seq: number, epoch?: string): Promise<Uint8Array | null> {
    epoch ??= (await this.head(room))?.epoch
    return epoch ? this.blobs.get(`${this.epochDir(room, epoch)}snap/${pad(seq)}.bin`) : null
  }

  /** snapshot `seq` with its merge context */
  async snapshot(room: string, seq: number, epoch?: string): Promise<(Snapshot & SnapMeta) | null> {
    epoch ??= (await this.head(room))?.epoch
    if (!epoch) return null
    const d = `${this.epochDir(room, epoch)}snap/${pad(seq)}`
    const meta = parse<SnapMeta>(await this.blobs.get(d + '.json'))
    if (!meta) return null
    const ctx = await this.blobs.get(d + '.ctx')
    return { ...meta, ctx: ctx ? decode<Entry[]>(dec.decode(ctx)) : [] }
  }

  /** stored snapshots, oldest first */
  async versions(room: string, epoch?: string): Promise<Version[]> {
    epoch ??= (await this.head(room))?.epoch
    if (!epoch) return []
    const d = `${this.epochDir(room, epoch)}snap/`
    const keys = (await this.blobs.list(d)).filter((k) => k.endsWith('.json')).sort()
    const out: Version[] = []
    for (const k of keys) {
      const m = parse<SnapMeta>(await this.blobs.get(k))
      if (m) out.push({ seq: m.seq, size: m.size, time: m.time })
    }
    return out
  }

  /** stored batches with `after` < seq ≤ `upTo` (flushed ones only), in order */
  async batches(room: string, after: number, upTo = Infinity, epoch?: string): Promise<Batch[]> {
    epoch ??= (await this.head(room))?.epoch
    if (!epoch) return []
    const segs = (await this.blobs.list(`${this.epochDir(room, epoch)}ops/`))
      .map((k) => {
        const m = /(\d+)-(\d+)\.jsonl$/.exec(k)
        return m ? { k, first: Number(m[1]), last: Number(m[2]) } : null
      })
      .filter((s) => s !== null && s.last > after && s.first <= upTo)
      .sort((a, b) => a!.first - b!.first)
    const seen = new Set<number>()
    const out: Batch[] = []
    for (const s of segs) {
      const text = dec.decode((await this.blobs.get(s!.k)) ?? new Uint8Array())
      for (const line of text.split('\n')) {
        if (!line) continue
        const b = decode<Batch>(line)
        if (b.seq > after && b.seq <= upTo && !seen.has(b.seq)) {
          seen.add(b.seq)
          out.push(b)
        }
      }
    }
    return out.sort((a, b) => a.seq - b.seq)
  }

  // ---------------------------------------------------------------- hub hooks

  /** hub `onBatch`: buffers the batch, written with the next segment */
  append(room: string, batch: Batch) {
    let buf = this.buffers.get(room)
    if (!buf) this.buffers.set(room, (buf = { batches: [], retry: 0 }))
    buf.batches.push(batch)
    if (!buf.timer) buf.timer = setTimeout(() => void this.flushRoom(room), this.opts.flushMs ?? 250)
  }

  /** writes buffered batches now (all rooms when `room` is omitted) */
  async flush(room?: string): Promise<void> {
    const rooms = room !== undefined ? [room] : [...this.buffers.keys()]
    await Promise.all(rooms.map((r) => this.flushRoom(r)))
  }

  private flushRoom(room: string): Promise<void> {
    const buf = this.buffers.get(room)
    if (buf) {
      clearTimeout(buf.timer)
      buf.timer = undefined
    }
    return this.run(room, async () => {
      const b = this.buffers.get(room)
      if (!b?.batches.length) return
      const batches = b.batches
      b.batches = []
      const head = await this.readHead(room)
      if (!head) return
      const key = `${this.epochDir(room, head.epoch)}ops/${pad(batches[0].seq)}-${pad(batches[batches.length - 1].seq)}.jsonl`
      try {
        await this.blobs.put(key, enc.encode(batches.map((x) => encode(x)).join('\n') + '\n'))
        b.retry = 0
      } catch (e) {
        this.opts.onError?.(e as Error, room)
        if (this.buffers.get(room) !== b) return
        b.batches.unshift(...batches)
        if (!b.timer) b.timer = setTimeout(() => void this.flushRoom(room), Math.min(30_000, 500 * 2 ** b.retry++))
      }
    })
  }

  /** hub `loadLog`: the latest snapshot and the batches after it (plus `retain` before it) */
  load(room: string): Promise<RoomLog | undefined> {
    return this.flushRoom(room).then(() =>
      this.run(room, async () => {
        this.heads.delete(room)
        const head = await this.readHead(room)
        if (!head) return undefined
        const snapshot = await this.snapshot(room, head.seq, head.epoch)
        const log = await this.batches(room, Math.max(0, head.seq - (this.opts.retain ?? 1000)), Infinity, head.epoch)
        return { log, snapshot: snapshot ?? { seq: head.seq, ctx: [] } }
      }),
    )
  }

  /**
   * Stores a snapshot taken from `base` (the head it was compacted from) and makes it the latest.
   * Returns the new head, or null when the room changed meanwhile (new upload, newer snapshot).
   */
  save(room: string, base: RoomHead, snap: Snapshot & { sites?: number }, file: Uint8Array): Promise<RoomHead | null> {
    return this.run(room, async () => {
      const head = await this.readHead(room)
      if (!head || head.epoch !== base.epoch || snap.seq <= head.seq) return null
      const d = this.epochDir(room, head.epoch)
      const time = Date.now()
      // head moves only once the snapshot is complete
      await Promise.all([
        this.blobs.put(`${d}snap/${pad(snap.seq)}.bin`, file),
        this.blobs.put(`${d}snap/${pad(snap.seq)}.ctx`, enc.encode(encode(snap.ctx))),
        this.blobs.put(`${d}snap/${pad(snap.seq)}.json`, json({ seq: snap.seq, size: file.length, time, sites: snap.sites } satisfies SnapMeta)),
      ])
      const next: RoomHead = { ...head, seq: snap.seq, size: file.length, time }
      await this.blobs.put(this.dir(room) + 'head.json', json(next))
      this.heads.set(room, next)
      if (this.opts.keep) await this.prune(room, next)
      return next
    })
  }

  /** drops snapshots beyond `keep` and the segments no remaining snapshot needs */
  private async prune(room: string, head: RoomHead) {
    const d = `${this.epochDir(room, head.epoch)}snap/`
    const seqs = [...new Set((await this.blobs.list(d)).map((k) => Number(/(\d+)\.\w+$/.exec(k)?.[1])).filter((n) => !Number.isNaN(n)))].sort((a, b) => a - b)
    const drop = seqs.slice(0, Math.max(0, seqs.length - this.opts.keep!))
    for (const s of drop) for (const ext of ['bin', 'json', 'ctx']) await this.blobs.delete(`${d}${pad(s)}.${ext}`)
    const oldest = seqs[drop.length]
    const floor = Math.min(oldest, head.seq - (this.opts.retain ?? 1000))
    for (const k of await this.blobs.list(`${this.epochDir(room, head.epoch)}ops/`)) {
      const last = Number(/-(\d+)\.jsonl$/.exec(k)?.[1])
      if (last <= floor) await this.blobs.delete(k)
    }
  }
}
