/**
 * Wire protocol between collaboration clients and a hub (sequencer).
 *
 * The hub is format-agnostic: it assigns a global sequence number to every
 * batch of operations, stores the log and relays presence. All transformation
 * and conflict handling happens in the clients, deterministically, so every
 * client that applies the log ends up with the same document.
 */

export interface UserInfo {
  /** stable user id (optional, for the host app) */
  id?: string
  name: string
  /** CSS color; the hub assigns one when missing */
  color?: string
}

export interface PeerInfo extends UserInfo {
  /** per-room connection id, assigned by the hub */
  site: number
  color: string
  presence?: unknown
}

/** A batch of operations as stored in the hub log. */
export interface Batch<Op = unknown> {
  seq: number
  site: number
  /** last seq the author had applied when creating the batch */
  base: number
  /** client generated id (dedupes resends after reconnects) */
  id: string
  ops: Op[]
}

export type ClientMsg =
  /** join a room; `have`/`site` on reconnect */
  | { t: 'hello'; room: string; user: UserInfo; have?: number; site?: number }
  | { t: 'ops'; id: string; base: number; ops: unknown[] }
  | { t: 'presence'; data: unknown }

export type HubMsg =
  /** `log` holds every batch with seq > `from`; `reset` when the hub lost history newer than `have` */
  | { t: 'welcome'; site: number; seq: number; from: number; log: Batch[]; peers: PeerInfo[]; reset?: boolean }
  | { t: 'batch'; b: Batch }
  | { t: 'presence'; site: number; data: unknown }
  | { t: 'join'; peer: PeerInfo }
  | { t: 'leave'; site: number }
  | { t: 'error'; message: string }

export const PEER_COLORS = ['#1a73e8', '#e8710a', '#188038', '#d93025', '#9334e6', '#c5221f', '#007b83', '#b06000', '#3f51b5', '#e52592']
