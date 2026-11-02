/**
 * Reference collaboration server (Bun, zero dependencies).
 *
 *   bun scripts/collab-server.ts            # port 8787, in memory
 *   PORT=9000 DATA=./collab-data bun scripts/collab-server.ts   # persisted to disk
 *
 * WebSocket   ws://host/collab                      hub protocol (src/collab/protocol.ts)
 * PUT         /rooms/:room/file  (body = file)      sets the room's base document (resets its history)
 * GET / HEAD  /rooms/:room/file                     base document, with HTTP Range support
 * GET         /rooms/:room                          { name, size, seq, peers }
 */
import { appendFileSync, existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { decode, encode } from '../src/collab/codec'
import { CollabHub, type HubConnection } from '../src/collab/hub'
import type { Batch, ClientMsg } from '../src/collab/protocol'

const PORT = Number(process.env.PORT ?? 8787)
const DATA = process.env.DATA
if (DATA) mkdirSync(DATA, { recursive: true })

const safe = (room: string) => room.replace(/[^\w.-]/g, '_').slice(0, 80)
const files = new Map<string, { bytes: Uint8Array<ArrayBuffer>; name: string; type: string }>()

function loadFile(room: string) {
  let f = files.get(room)
  if (!f && DATA && existsSync(join(DATA, safe(room) + '.bin'))) {
    const meta = JSON.parse(readFileSync(join(DATA, safe(room) + '.json'), 'utf8')) as { name: string; type: string }
    f = { bytes: new Uint8Array(readFileSync(join(DATA, safe(room) + '.bin'))) as Uint8Array<ArrayBuffer>, ...meta }
    files.set(room, f)
  }
  return f
}

const hub = new CollabHub({
  onBatch: (room, b) => DATA && appendFileSync(join(DATA, safe(room) + '.log'), encode(b) + '\n'),
  loadLog: (room) => {
    if (!DATA || !existsSync(join(DATA, safe(room) + '.log'))) return undefined
    return readFileSync(join(DATA, safe(room) + '.log'), 'utf8').split('\n').filter(Boolean).map((l) => decode<Batch>(l))
  },
})

const CORS = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Methods': 'GET, HEAD, PUT, OPTIONS',
  'Access-Control-Allow-Headers': 'Range, Content-Type, X-File-Name',
  'Access-Control-Expose-Headers': 'Content-Range, Content-Length, Accept-Ranges, X-File-Name',
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
    const m = /^\/rooms\/([^/]+)(\/file)?$/.exec(url.pathname)
    if (!m) return new Response('office-engine collab server', { headers: CORS })
    const room = decodeURIComponent(m[1])
    if (!m[2]) {
      const f = loadFile(room)
      const r = hub.rooms.get(room)
      return Response.json({ name: f?.name, size: f?.bytes.length ?? 0, seq: r?.seq ?? 0, peers: r?.peers.size ?? 0 }, { headers: CORS })
    }
    if (req.method === 'PUT') {
      const bytes = new Uint8Array(await req.arrayBuffer())
      const name = decodeURIComponent(req.headers.get('x-file-name') ?? 'document')
      const type = req.headers.get('content-type') ?? 'application/octet-stream'
      files.set(room, { bytes, name, type })
      hub.reset(room)
      if (DATA) {
        writeFileSync(join(DATA, safe(room) + '.bin'), bytes)
        writeFileSync(join(DATA, safe(room) + '.json'), JSON.stringify({ name, type }))
        writeFileSync(join(DATA, safe(room) + '.log'), '')
      }
      return Response.json({ room, size: bytes.length }, { headers: CORS })
    }
    const f = loadFile(room)
    if (!f) return new Response('No file for this room', { status: 404, headers: CORS })
    const headers: Record<string, string> = { ...CORS, 'Accept-Ranges': 'bytes', 'Content-Type': f.type, 'X-File-Name': encodeURIComponent(f.name), 'Cache-Control': 'no-store' }
    const range = /^bytes=(\d*)-(\d*)$/.exec(req.headers.get('range') ?? '')
    if (range) {
      const size = f.bytes.length
      let start = range[1] ? Number(range[1]) : size - Number(range[2])
      let end = range[1] && range[2] ? Number(range[2]) : size - 1
      start = Math.max(0, start)
      end = Math.min(size - 1, end)
      if (start > end) return new Response(null, { status: 416, headers: { ...headers, 'Content-Range': `bytes */${size}` } })
      return new Response(req.method === 'HEAD' ? null : f.bytes.slice(start, end + 1), {
        status: 206,
        headers: { ...headers, 'Content-Range': `bytes ${start}-${end}/${size}`, 'Content-Length': String(end - start + 1) },
      })
    }
    return new Response(req.method === 'HEAD' ? null : f.bytes, { headers: { ...headers, 'Content-Length': String(f.bytes.length) } })
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

console.log(`office-engine collab server on http://localhost:${PORT} (ws: /collab)${DATA ? `, data in ${DATA}` : ''}`)
