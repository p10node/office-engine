/**
 * Runtime-agnostic collaboration hub: one sequencer + op log per room.
 * Wire it to any socket server (see scripts/collab-server.ts for Bun):
 *
 *   const conn = hub.connect((msg) => socket.send(encode(msg)))
 *   socket.onmessage = (e) => conn.receive(decode(e.data))
 *   socket.onclose = () => conn.close()
 */
import { PEER_COLORS, type Batch, type ClientMsg, type HubMsg, type PeerInfo } from './protocol'

export interface RoomState {
  id: string
  seq: number
  log: Batch[]
  ids: Set<string>
  peers: Map<number, { info: PeerInfo; send: (m: HubMsg) => void }>
  nextSite: number
}

export interface HubOptions {
  /** called for every sequenced batch (persist it to rebuild rooms after a restart) */
  onBatch?: (room: string, batch: Batch) => void
  /** rebuilds a room log (e.g. from storage) the first time the room is used */
  loadLog?: (room: string) => Batch[] | undefined
  /** called when the last peer left a room */
  onIdle?: (room: string) => void
}

export interface HubConnection {
  receive(msg: ClientMsg): void
  close(): void
}

export class CollabHub {
  readonly rooms = new Map<string, RoomState>()

  constructor(private opts: HubOptions = {}) {}

  room(id: string): RoomState {
    let r = this.rooms.get(id)
    if (!r) {
      const log = this.opts.loadLog?.(id) ?? []
      r = { id, seq: log.length ? log[log.length - 1].seq : 0, log, ids: new Set(log.map((b) => b.id)), peers: new Map(), nextSite: 1 }
      for (const b of log) r.nextSite = Math.max(r.nextSite, b.site + 1)
      this.rooms.set(id, r)
    }
    return r
  }

  /** forgets a room (its log and peers) — e.g. when a new base file is uploaded */
  reset(id: string) {
    const r = this.rooms.get(id)
    if (r) for (const p of r.peers.values()) p.send({ t: 'error', message: 'Room was reset' })
    this.rooms.delete(id)
  }

  connect(send: (m: HubMsg) => void): HubConnection {
    let room: RoomState | undefined
    let site = 0
    const broadcast = (m: HubMsg, except?: number) => {
      if (!room) return
      for (const [s, p] of room.peers) if (s !== except) p.send(m)
    }
    return {
      receive: (msg) => {
        if (msg.t === 'hello') {
          if (room) return
          room = this.room(msg.room)
          site = msg.site && msg.site < room.nextSite && !room.peers.has(msg.site) ? msg.site : room.nextSite++
          const used = new Set([...room.peers.values()].map((p) => p.info.color))
          const color = msg.user.color || PEER_COLORS.find((c) => !used.has(c)) || PEER_COLORS[site % PEER_COLORS.length]
          const info: PeerInfo = { ...msg.user, site, color }
          const reset = msg.have !== undefined && msg.have > room.seq
          const from = msg.have !== undefined && !reset ? msg.have : 0
          send({
            t: 'welcome', site, seq: room.seq, from, reset,
            log: room.log.filter((b) => b.seq > from),
            peers: [...room.peers.values()].map((p) => p.info),
          })
          room.peers.set(site, { info, send })
          broadcast({ t: 'join', peer: info }, site)
          return
        }
        if (!room) return send({ t: 'error', message: 'Not joined' })
        if (msg.t === 'ops') {
          if (room.ids.has(msg.id)) return
          const b: Batch = { seq: ++room.seq, site, base: Math.min(msg.base, room.seq - 1), id: msg.id, ops: msg.ops }
          room.log.push(b)
          room.ids.add(b.id)
          this.opts.onBatch?.(room.id, b)
          broadcast({ t: 'batch', b })
        } else if (msg.t === 'presence') {
          const p = room.peers.get(site)
          if (p) p.info.presence = msg.data
          broadcast({ t: 'presence', site, data: msg.data }, site)
        }
      },
      close: () => {
        if (!room || !room.peers.has(site)) return
        room.peers.delete(site)
        broadcast({ t: 'leave', site })
        if (!room.peers.size) this.opts.onIdle?.(room.id)
      },
    }
  }
}
