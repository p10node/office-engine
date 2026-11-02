/**
 * Server-side snapshots: replays shared history onto a document headlessly, exactly the way a
 * collaborator applies it, so the result can be saved as the room's new base file.
 *
 *   const client = await DocumentClient.open(snapshotFile, { worker: false })
 *   const snap = await compact(client, { from: snapshotMeta, batches })
 *   const file = await client.call('save')        // store it, then hub.snapshot(room, snap)
 */
import type { DocumentClient } from '../core/worker/client'
import { getHost } from './hosts'
import { startCollab } from './index'
import type { Batch, ClientMsg, HubMsg, Snapshot } from './protocol'
import type { Transport, TransportStatus } from './transport'

export interface CompactOptions {
  /** the snapshot `client` was opened from (default: the original upload, seq 0) */
  from?: Snapshot
  /** batches to apply: every batch after `from.seq`, in order (older ones are ignored) */
  batches: Batch[]
  /** canonical batches kept as the new snapshot's context (default 1000) */
  window?: number
}

/** feeds a stored log to a session as if a hub had sent it */
class ReplayTransport implements Transport {
  onMessage?: (msg: HubMsg) => void
  onStatus?: (s: TransportStatus) => void

  constructor(private welcome: HubMsg) {}

  connect() {
    queueMicrotask(() => this.onStatus?.('open'))
  }

  send(msg: ClientMsg) {
    if (msg.t === 'hello') queueMicrotask(() => this.onMessage?.(this.welcome))
  }

  close() {}
}

/**
 * Applies `batches` to `client` (opened on the file of snapshot `from`) and returns the snapshot it
 * now holds. Save the document afterwards with `client.call('save')`.
 */
export async function compact(client: DocumentClient, opts: CompactOptions): Promise<Snapshot> {
  if (client.kind === 'docx' && !getHost(client, 'docx')) throw new Error('compact: DOCX needs its editor host registered (registerHost)')
  const from = opts.from ?? { seq: 0, ctx: [] }
  const batches = opts.batches.filter((b) => b.seq > from.seq).sort((a, b) => a.seq - b.seq)
  batches.forEach((b, i) => {
    if (b.seq !== from.seq + i + 1) throw new Error(`compact: batch ${from.seq + i + 1} is missing`)
  })
  const seq = from.seq + batches.length
  // site 0 is never assigned by a hub: every batch is someone else's
  const transport = new ReplayTransport({ t: 'welcome', site: 0, seq, from: from.seq, log: batches, peers: [], ctx: from.ctx })
  const collab = startCollab(client, { transport, room: '', user: { name: 'compact' }, base: from.seq })
  try {
    await new Promise<void>((resolve, reject) => {
      const check = () => {
        if (collab.status === 'live') resolve()
        else if (collab.status === 'error') reject(new Error(collab.session.error ?? 'compact failed'))
      }
      collab.on('status', check)
      check()
    })
    await collab.session.idle()
    if (collab.session.rev !== seq) throw new Error(`compact: stopped at ${collab.session.rev} of ${seq}`)
    return { seq, ctx: collab.session.context(opts.window ?? 1000) }
  } finally {
    collab.close()
  }
}
