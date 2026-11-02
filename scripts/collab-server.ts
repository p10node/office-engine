/**
 * Reference collaboration server (Bun, zero dependencies).
 *
 *   bun scripts/collab-server.ts                        # port 8787, in memory
 *   DATA=./collab-data bun scripts/collab-server.ts     # persisted to a directory
 *   S3_BUCKET=my-bucket bun scripts/collab-server.ts    # persisted to S3 (S3_REGION, S3_ENDPOINT, S3_ACCESS_KEY_ID… see Bun.S3Client)
 *
 *   PREFIX=collab/        key prefix in the store
 *   SNAPSHOT_EVERY=500    batches between snapshots (also taken when a room goes idle)
 *   KEEP=0                snapshots kept per room (0: all — the version history)
 *
 * Edits are saved as small append-only op segments; the document itself is rewritten (a snapshot,
 * on a worker thread) only every SNAPSHOT_EVERY batches and when the last editor leaves.
 *
 * WebSocket   ws://host/collab                      hub protocol (src/collab/protocol.ts)
 * PUT         /rooms/:room/file  (body = file)      sets the room's document (starts a new history)
 * GET         /rooms/:room                          { name, size, seq, snapshot, peers }: open `file?seq=<snapshot>`, join with base = snapshot
 * GET / HEAD  /rooms/:room/file?seq=N              the document at N, header X-Collab-Seq, Range support. N between snapshots
 *                                                   is materialized on demand (version history). Without `seq`: the uploaded
 *                                                   file (seq 0), for clients that don't pass `base`
 * GET         /rooms/:room/history[?after=N]        { seq, snapshots: [{ seq, size, time }], edits: [{ seq, time, user }] }
 */
import { existsSync, mkdirSync, readdirSync, renameSync, rmSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { S3Client } from 'bun'
import { decode, encode } from '../src/collab/codec'
import { compact } from '../src/collab/compact'
import { CollabHub, type HubConnection } from '../src/collab/hub'
import type { Batch, ClientMsg } from '../src/collab/protocol'
import { MemoryBlobStore, RoomStore, type BlobStore, type RoomHead } from '../src/collab/store'
import { DocumentClient } from '../src/core/worker/client'

const PORT = Number(process.env.PORT ?? 8787)
const DATA = process.env.DATA
const BUCKET = process.env.S3_BUCKET ?? process.env.AWS_BUCKET
const SNAPSHOT_EVERY = Number(process.env.SNAPSHOT_EVERY ?? 500)
const KEEP = Number(process.env.KEEP ?? 0)

// ---------------------------------------------------------------- storage

function dirStore(dir: string): BlobStore {
  return {
    async get(key) {
      const f = Bun.file(join(dir, key))
      return (await f.exists()) ? new Uint8Array(await f.arrayBuffer()) : null
    },
    async put(key, data) {
      const path = join(dir, key)
      mkdirSync(dirname(path), { recursive: true })
      await Bun.write(path + '.tmp', data)
      renameSync(path + '.tmp', path)
    },
    async list(prefix) {
      const base = prefix.slice(0, prefix.lastIndexOf('/') + 1)
      if (!existsSync(join(dir, base))) return []
      return readdirSync(join(dir, base), { recursive: true, encoding: 'utf8' })
        .map((p) => base + p.replaceAll('\\', '/'))
        .filter((k) => k.startsWith(prefix) && /\.(json|jsonl|bin|ctx)$/.test(k))
    },
    async delete(key) {
      rmSync(join(dir, key), { force: true })
    },
  }
}

function s3Store(bucket: string): BlobStore {
  const s3 = new S3Client({ bucket })
  return {
    async get(key) {
      try {
        return new Uint8Array(await s3.file(key).arrayBuffer())
      } catch (e) {
        const err = e as { code?: string; message?: string }
        if (err.code === 'NoSuchKey' || /NoSuchKey|not found/i.test(err.message ?? '')) return null
        throw e
      }
    },
    async put(key, data) {
      await s3.write(key, data)
    },
    async list(prefix) {
      const keys: string[] = []
      let token: string | undefined
      do {
        const r = await s3.list({ prefix, continuationToken: token })
        for (const c of r.contents ?? []) keys.push(c.key)
        token = r.isTruncated ? r.nextContinuationToken : undefined
      } while (token)
      return keys
    },
    async delete(key) {
      await s3.delete(key)
    },
  }
}

const store = new RoomStore(BUCKET ? s3Store(BUCKET) : DATA ? dirStore(DATA) : new MemoryBlobStore(), {
  prefix: process.env.PREFIX,
  flushMs: BUCKET ? 250 : 0,
  keep: KEEP || undefined,
  onError: (e, room) => console.error(`[store] ${room}: ${e.message} (retrying)`),
})

// ---------------------------------------------------------------- snapshots

/** documents are rebuilt on a worker thread: big sheets don't stall the sockets */
const worker = () => new Worker(join(import.meta.dir, '../src/core/worker/worker.ts'), { type: 'module' })

/** the document at `upTo`: the stored snapshot `from` + the batches after it */
async function materialize(room: string, epoch: string, from: number, batches: Batch[]) {
  const [meta, bytes] = await Promise.all([store.snapshot(room, from, epoch), store.file(room, from, epoch)])
  if (!meta || !bytes) throw new Error(`snapshot ${from} of ${room} is missing`)
  const client = await DocumentClient.open(new Blob([bytes as BlobPart]), { worker })
  try {
    if (client.kind === 'docx') throw new NoSnapshots()
    const snap = await compact(client, { from: meta, batches })
    const file = new Uint8Array(await (await client.call('save')).arrayBuffer())
    return { snap, file }
  } finally {
    client.destroy()
  }
}

/** DOCX edits live in the view's editor: those rooms keep replaying from the upload */
class NoSnapshots extends Error {}

const running = new Map<string, Promise<void>>()
const again = new Set<string>()
const unsupported = new Set<string>()

/** takes a snapshot of a loaded room (one at a time per room) */
function snapshot(room: string): Promise<void> {
  const cur = running.get(room)
  if (cur) {
    again.add(room)
    return cur
  }
  const p = (async () => {
    const r = hub.rooms.get(room)
    const head = await store.head(room)
    if (!r || !head || r.seq <= head.seq || r.start > head.seq || unsupported.has(head.epoch)) return
    const t0 = performance.now()
    const { snap, file } = await materialize(room, head.epoch, head.seq, r.log.filter((b) => b.seq > head.seq))
    const next = await store.save(room, head, { ...snap, sites: r.nextSite }, file)
    if (next && hub.rooms.get(room) === r) hub.snapshot(room, snap)
    if (next) console.log(`[snapshot] ${room} @${snap.seq} (${(file.length / 1024).toFixed(0)} KB, ${Math.round(performance.now() - t0)} ms)`)
  })()
    .catch(async (e) => {
      if (e instanceof NoSnapshots) unsupported.add((await store.head(room))?.epoch ?? '')
      else console.error(`[snapshot] ${room}:`, e)
    })
    .finally(() => {
      running.delete(room)
      if (again.delete(room)) void snapshot(room)
    })
  running.set(room, p)
  return p
}

const latest = (room: string) => Math.max(0, ...(hub.rooms.get(room)?.snaps.keys() ?? []))

const hub = new CollabHub({
  onBatch: (room, b) => {
    store.append(room, b)
    if (SNAPSHOT_EVERY > 0 && b.seq - latest(room) >= SNAPSHOT_EVERY) void snapshot(room)
  },
  loadLog: (room) => store.load(room),
  onIdle: (room) => {
    void snapshot(room)
      .then(() => store.flush(room))
      .then(() => hub.unload(room))
  },
})

// ---------------------------------------------------------------- files

/** recently served documents (immutable per epoch + seq) */
const files = new Map<string, Uint8Array>()

async function documentAt(room: string, head: RoomHead, seq: number): Promise<Uint8Array | null> {
  const key = `${room}|${head.epoch}|${seq}`
  let bytes = files.get(key)
  if (!bytes) {
    bytes = (await store.file(room, seq, head.epoch)) ?? undefined
    if (!bytes && !unsupported.has(head.epoch)) {
      // between snapshots: the nearest older snapshot + stored batches
      await store.flush(room)
      const base = (await store.versions(room, head.epoch)).filter((v) => v.seq <= seq).pop()
      if (!base) return null
      const batches = await store.batches(room, base.seq, seq, head.epoch)
      if (batches.length !== seq - base.seq) return null
      try {
        bytes = (await materialize(room, head.epoch, base.seq, batches)).file
      } catch (e) {
        if (!(e instanceof NoSnapshots)) throw e
        unsupported.add(head.epoch)
        return null
      }
    }
    if (!bytes) return null
    files.set(key, bytes)
    while (files.size > 8) files.delete(files.keys().next().value!)
  }
  return bytes
}

// ---------------------------------------------------------------- http

const CORS = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Methods': 'GET, HEAD, PUT, OPTIONS',
  'Access-Control-Allow-Headers': 'Range, Content-Type, X-File-Name',
  'Access-Control-Expose-Headers': 'Content-Range, Content-Length, Accept-Ranges, X-File-Name, X-Collab-Seq, ETag',
}

type WsData = { conn?: HubConnection }

Bun.serve<WsData>({
  port: PORT,
  async fetch(req, server) {
    const url = new URL(req.url)
    if (url.pathname === '/collab') {
      if (server.upgrade(req, { data: {} })) return
      return new Response('WebSocket expected', { status: 426 })
    }
    if (req.method === 'OPTIONS') return new Response(null, { headers: CORS })
    const m = /^\/rooms\/([^/]+)(?:\/(file|history))?$/.exec(url.pathname)
    if (!m) return new Response('office-engine collab server', { headers: CORS })
    const room = decodeURIComponent(m[1])
    const what = m[2]

    if (what === 'file' && req.method === 'PUT') {
      const bytes = new Uint8Array(await req.arrayBuffer())
      const name = decodeURIComponent(req.headers.get('x-file-name') ?? 'document')
      const type = req.headers.get('content-type') ?? 'application/octet-stream'
      hub.reset(room)
      const head = await store.create(room, bytes, { name, type })
      return Response.json({ room, size: head.size, seq: 0 }, { headers: CORS })
    }

    const head = await store.head(room)
    if (!head) return new Response('No file for this room', { status: 404, headers: CORS })
    const r = hub.rooms.get(room)
    const seq = r?.seq ?? (await store.batches(room, head.seq)).at(-1)?.seq ?? head.seq

    if (!what) return Response.json({ name: head.name, size: head.size, seq, snapshot: head.seq, peers: r?.peers.size ?? 0 }, { headers: CORS })

    if (what === 'history') {
      await store.flush(room)
      const after = Number(url.searchParams.get('after') ?? 0)
      const edits = (await store.batches(room, after)).map((b) => ({ seq: b.seq, time: b.time, user: b.user }))
      return Response.json({ seq, snapshots: await store.versions(room), edits }, { headers: CORS })
    }

    // no `seq`: the original upload, which is what a client joining without `base` expects
    const at = url.searchParams.has('seq') ? Number(url.searchParams.get('seq')) : 0
    if (!Number.isInteger(at) || at < 0 || at > seq) return new Response('No such version', { status: 404, headers: CORS })
    const bytes = await documentAt(room, head, at)
    if (!bytes) return new Response('Version no longer available', { status: 410, headers: CORS })
    const headers: Record<string, string> = {
      ...CORS,
      'Accept-Ranges': 'bytes',
      'Content-Type': head.type,
      'X-File-Name': encodeURIComponent(head.name),
      'X-Collab-Seq': String(at),
      ETag: `"${head.epoch}-${at}"`,
      'Cache-Control': 'no-store',
    }
    const range = /^bytes=(\d*)-(\d*)$/.exec(req.headers.get('range') ?? '')
    if (range) {
      const size = bytes.length
      let start = range[1] ? Number(range[1]) : size - Number(range[2])
      let end = range[1] && range[2] ? Number(range[2]) : size - 1
      start = Math.max(0, start)
      end = Math.min(size - 1, end)
      if (start > end) return new Response(null, { status: 416, headers: { ...headers, 'Content-Range': `bytes */${size}` } })
      return new Response(req.method === 'HEAD' ? null : bytes.slice(start, end + 1), {
        status: 206,
        headers: { ...headers, 'Content-Range': `bytes ${start}-${end}/${size}`, 'Content-Length': String(end - start + 1) },
      })
    }
    return new Response(req.method === 'HEAD' ? null : (bytes as Uint8Array<ArrayBuffer>), { headers: { ...headers, 'Content-Length': String(bytes.length) } })
  },
  websocket: {
    open(ws) {
      ws.data.conn = hub.connect((msg) => ws.send(encode(msg)))
    },
    message(ws, data) {
      if (typeof data === 'string') ws.data.conn?.receive(decode<ClientMsg>(data))
    },
    close(ws) {
      ws.data.conn?.close()
    },
  },
})

for (const sig of ['SIGINT', 'SIGTERM'] as const) {
  process.on(sig, () => {
    void store.flush().finally(() => process.exit(0))
  })
}

const where = BUCKET ? `s3://${BUCKET}/${process.env.PREFIX ?? ''}` : DATA ? DATA : 'memory'
console.log(`office-engine collab server on http://localhost:${PORT} (ws: /collab), storage: ${where}, snapshot every ${SNAPSHOT_EVERY} batches`)
