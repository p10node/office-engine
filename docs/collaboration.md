# Real-time collaboration

Everyone who opens the same base file in the same room edits it together, like Google Docs: changes appear live, collaborators show up as avatars and their position is drawn in their color (DOCX caret and selection, XLSX selected range, PPTX selected shape). Edits made while offline sync on reconnect.

```bash
bun run collab   # reference hub on :8787 (DATA=./collab-data or S3_BUCKET=… persists rooms)
bun run dev      # playground: open a file → "Share live" → send the link
```

```tsx
// open the room's latest snapshot and join from it: only later edits are replayed
const { snapshot } = await (await fetch(`https://hub.example.com/rooms/${id}`)).json()

<DocViewer
  src={`https://hub.example.com/rooms/${id}/file?seq=${snapshot}`}
  editable
  collab={{ transport: 'wss://hub.example.com/collab', room: id, base: snapshot, user: { name: 'Ann' } }}
/>
```

`base` must be the sequence number of the file `src` points to. Without it the client assumes the original upload (`/rooms/:room/file` without `seq`) and replays the whole history.

Without React: `mount(el, { src, collab })`, `<office-viewer src="…" collab="wss://…/collab" room="doc-42" user="Ann" base="42">`, or `startCollab(client, options)` on a `DocumentClient` (`@p10node/office-engine/collab`, no React).

## How it works

- The hub (`CollabHub`, runtime-agnostic) only orders batches of operations, keeps the log and relays presence. Plug it into any WebSocket server (`scripts/collab-server.ts` does it for Bun) or implement `Transport` for your own backend.
- Clients apply their edits immediately. When a collaborator's batch was ordered first, a client rolls its pending edits back, applies that batch, then replays its own (transformed) — so every client ends up with the same document. Newcomers open the latest snapshot and replay the batches after it.
- Structural changes (inserting / deleting rows or columns, the sheet list, the slide list) wait for the hub's order (one round trip).

| Format | Shared as                                                         | Same thing edited at once                                                                             |
| ------ | ----------------------------------------------------------------- | ----------------------------------------------------------------------------------------------------- |
| DOCX   | text splices per paragraph, paragraph and table states            | characters merge inside a paragraph; split / join / paragraph format: last writer wins                |
| XLSX   | cell values, formats, layout, fill / sort, row & column structure | last writer wins per cell and property; inserted / deleted rows and columns shift the other edits     |
| PPTX   | shape text, shape edits, slide list                               | last writer wins per shape                                                                            |
| all    | comments (one op per comment set / deleted)                       | last writer wins per comment; DOCX ranges travel with the text, XLSX threads move with rows / columns |

## Storage, snapshots and version history

Saving an edit never rewrites the document:

- Every batch the hub orders is appended to the room's op log. `RoomStore` buffers batches and writes them as one small segment per flush (every 250 ms on S3).
- Every `SNAPSHOT_EVERY` batches, and when the last editor leaves, the server replays the new batches onto the latest snapshot (`compact`, on a worker thread) and stores the result as the next snapshot. That is the only time the document file is written again.
- Snapshots and the op log together are the version history: the document at any `seq` is the nearest older snapshot plus the batches up to `seq`. Each batch records its `time` and `user`.
- Downloading needs no server work: the browser exports the document it already holds.

`RoomStore` works on any key → bytes store (`BlobStore`: `get` / `put` / `list` / `delete`):

| Key                                       | Content                                                     |
| ----------------------------------------- | ----------------------------------------------------------- |
| `{room}/head.json`                        | latest snapshot: `{ epoch, seq, size, name, type, time }`   |
| `{room}/{epoch}/snap/{seq}.bin`           | the document at `seq` (`0`: the uploaded file)              |
| `{room}/{epoch}/snap/{seq}.json` / `.ctx` | snapshot metadata / merge context                           |
| `{room}/{epoch}/ops/{first}-{last}.jsonl` | ordered batches, one segment per flush                      |

Uploading a new file starts a new epoch and deletes the previous one.

### Reference server

| Variable         | Default | Meaning                                                                                                              |
| ---------------- | ------- | -------------------------------------------------------------------------------------------------------------------- |
| `PORT`           | `8787`  | HTTP + WebSocket port                                                                                                |
| `DATA`           | –       | store rooms in this directory                                                                                        |
| `S3_BUCKET`      | –       | store rooms in S3; credentials from `S3_REGION`, `S3_ENDPOINT`, `S3_ACCESS_KEY_ID`… (or `AWS_*`), see `Bun.S3Client` |
| `PREFIX`         | –       | key prefix in the store                                                                                              |
| `SNAPSHOT_EVERY` | `500`   | batches between snapshots (`0`: only when a room goes idle)                                                          |
| `KEEP`           | `0`     | snapshots kept per room; `0` keeps all (full history), otherwise older snapshots and their batches are deleted       |

| Request                                | Result                                                                                                      |
| -------------------------------------- | ----------------------------------------------------------------------------------------------------------- |
| `PUT /rooms/:room/file`                | sets the room's document (body = file, header `X-File-Name`) and starts a new history                       |
| `GET /rooms/:room`                     | `{ name, size, seq, snapshot, peers }`                                                                      |
| `GET /rooms/:room/file?seq=N`          | the document at `N` (Range support, header `X-Collab-Seq`); between snapshots it is built on demand         |
| `GET /rooms/:room/file`                | the uploaded file (`seq` 0)                                                                                 |
| `GET /rooms/:room/history?after=N`     | `{ seq, snapshots: [{ seq, size, time }], edits: [{ seq, time, user }] }`                                   |
| `ws://…/collab`                        | hub protocol (`src/collab/protocol.ts`)                                                                     |

### Your own backend

```ts
import { CollabHub, RoomStore, compact, type BlobStore } from '@p10node/office-engine/collab'
import { DocumentClient } from '@p10node/office-engine/core'

const store = new RoomStore(myBlobStore)              // e.g. S3: get / put / list / delete
const hub = new CollabHub({
  onBatch: (room, b) => store.append(room, b),
  loadLog: (room) => store.load(room),                 // may return a promise
  onIdle: (room) => void takeSnapshot(room),
})

async function takeSnapshot(room: string) {
  const r = hub.rooms.get(room)
  const head = await store.head(room)
  if (!r || !head || r.seq <= head.seq) return
  const client = await DocumentClient.open(new Blob([(await store.file(room, head.seq))!]), { worker: false })
  const snap = await compact(client, { from: (await store.snapshot(room, head.seq))!, batches: r.log.filter((b) => b.seq > head.seq) })
  const file = new Uint8Array(await (await client.call('save')).arrayBuffer())
  client.destroy()
  if (await store.save(room, head, { ...snap, sites: r.nextSite }, file)) hub.snapshot(room, snap)
}
```

- A snapshot carries a merge context (`ctx`, the last `window` = 1000 batches as applied), so batches still based on an older state (an editor coming back online) merge correctly for clients that joined from it.
- `hub.snapshot()` keeps `retain` (1000) batches below the latest snapshot in memory for clients reconnecting from older states; older ones must reopen the document.
- `compact` works headless for XLSX and PPTX. Run it on a worker (`DocumentClient.open(file, { worker: () => new Worker(…) })`) so large sheets don't block the server.

## Limitations

- Undo / redo only touches your own changes
- DOCX rooms are not snapshotted yet (their edits run in the view's editor): newcomers replay the whole history
- Deleting rows / columns can't be undone while collaborating
- One hub process per room; batches not yet flushed (up to 250 ms on S3) are lost if it crashes, and clients ahead of the stored log must reopen
- The reference server has no authentication — put it behind your own auth
