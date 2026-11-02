/**
 * Real-time collaboration for office-engine documents.
 *
 *   const collab = startCollab(client, { transport: 'wss://host/collab', room: 'doc-42', user: { name: 'Ann' } })
 *   collab.on('status', console.log)
 *   collab.close()
 *
 * Every participant must open the same base file; the hub keeps the operation log on top of it.
 * DocViewer does all of this through its `collab` prop.
 */
import type { DocumentClient } from '../core/worker/client'
import type { Methods } from '../core/worker/protocol'
import { DocxCollab, type DocxPresence } from './docx'
import { setPeers, type PeerView } from './hosts'
import type { PeerInfo, UserInfo } from './protocol'
import { CollabSession, type Adapter, type CollabStatus, type SessionEvents } from './session'
import { WebSocketTransport, type Transport } from './transport'
import { XlsxCollab } from './xlsx'
import { PptxCollab } from './pptx'

export interface CollabOptions {
  /** hub connection: a WebSocket URL or any Transport */
  transport: Transport | string
  /** shared document id */
  room: string
  user: UserInfo
  /**
   * Sequence number of the snapshot `src` was opened from (see `compact`); default 0, the original
   * upload. The hub then only sends the batches after it.
   */
  base?: number
}

export interface Collab {
  readonly session: CollabSession<unknown>
  readonly status: CollabStatus
  /** collaborators currently connected (without the local user) */
  readonly peers: PeerInfo[]
  /** local operations not yet confirmed by the hub */
  readonly pending: number
  on<K extends keyof SessionEvents>(event: K, fn: (d: SessionEvents[K]) => void): () => void
  /** screen rectangles of a collaborator's DOCX selection (for custom overlays) */
  rects(root: HTMLElement, peer: PeerView): { caret: DOMRect | null; sel: DOMRect[] } | null
  close(): void
}

interface AnyAdapter extends Adapter<unknown> {
  /** whether `method` changes the document (and becomes a shared operation) */
  captures(method: keyof Methods): boolean
  intercept(method: keyof Methods, args: unknown[]): Promise<unknown> | undefined
  submit?: (op: never) => Promise<unknown>
  push?: (op: never, undo: never) => void
  nextId?: () => string
  onPresence?: (p: never) => void
  setSite?(site: number): void
  dispose?(): void
}

export function startCollab(client: DocumentClient, opts: CollabOptions): Collab {
  const transport = typeof opts.transport === 'string' ? new WebSocketTransport(opts.transport) : opts.transport
  const docx = client.kind === 'docx' ? new DocxCollab(client) : undefined
  const adapter = (docx ?? (client.kind === 'xlsx' ? new XlsxCollab(client) : new PptxCollab(client))) as unknown as AnyAdapter
  const session = new CollabSession<unknown>(transport, adapter, { room: opts.room, user: opts.user, base: opts.base })
  adapter.submit = (op) => session.submit(op)
  adapter.push = (op, undo) => session.push(op, undo)
  adapter.nextId = () => session.nextId()
  adapter.onPresence = (p) => session.setPresence(p)
  client.interceptor = (method, args) => {
    if (!adapter.captures(method)) return undefined
    // editing before the shared history was replayed would fork the document
    if (session.status === 'connecting' || session.status === 'syncing' || session.status === 'error') {
      return Promise.reject(new Error(session.status === 'error' ? session.error ?? 'Collaboration stopped' : 'Document is still syncing'))
    }
    return adapter.intercept(method, args)
  }
  const offs = [
    session.on('peers', (peers) => setPeers(client, peers)),
    session.on('status', (s) => {
      if (s === 'live') adapter.setSite?.(session.site)
    }),
  ]
  return {
    session,
    get status() {
      return session.status
    },
    get peers() {
      return [...session.peers.values()]
    },
    get pending() {
      return session.pending
    },
    on: (event, fn) => session.on(event, fn),
    rects: (root, peer) => {
      const p = peer.presence as DocxPresence | undefined
      return docx && p?.k === 'docx' ? docx.rects(root, p) : null
    },
    close() {
      offs.forEach((f) => f())
      session.close()
      if (client.interceptor) client.interceptor = undefined
      adapter.dispose?.()
      setPeers(client, [])
    },
  }
}

export { CollabHub, type HubOptions, type HubConnection, type RoomLog, type RoomState } from './hub'
export { compact, type CompactOptions } from './compact'
export { RoomStore, MemoryBlobStore, type BlobStore, type RoomHead, type RoomStoreOptions, type Version } from './store'
export { CollabSession, type Adapter, type CollabStatus } from './session'
export { MemoryTransport, WebSocketTransport, type Transport, type TransportStatus } from './transport'
export type { Batch, ClientMsg, Entry, HubMsg, PeerInfo, Snapshot, UserInfo } from './protocol'
export { encode, decode } from './codec'
