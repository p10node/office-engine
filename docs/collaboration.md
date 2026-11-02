# Real-time collaboration

Everyone who opens the same base file in the same room edits it together, like Google Docs: changes appear live, collaborators show up as avatars and their position is drawn in their color (DOCX caret and selection, XLSX selected range, PPTX selected shape). Edits made while offline sync on reconnect.

```bash
bun run collab   # reference hub on :8787 (DATA=./collab-data persists rooms)
bun run dev      # playground: open a file → "Share live" → send the link
```

```tsx
<DocViewer
  src={`https://hub.example.com/rooms/${id}/file`}   // the shared base document
  editable
  collab={{ transport: 'wss://hub.example.com/collab', room: id, user: { name: 'Ann' } }}
/>
```

Without React: `mount(el, { src, collab })`, `<office-viewer src="…" collab="wss://…/collab" room="doc-42" user="Ann">`, or `startCollab(client, options)` on a `DocumentClient` (`@p10node/office-engine/collab`, no React).

## How it works

- The hub (`CollabHub`, runtime-agnostic) only orders batches of operations, keeps the log and relays presence. Plug it into any WebSocket server (`scripts/collab-server.ts` does it for Bun) or implement `Transport` for your own backend.
- Clients apply their edits immediately. When a collaborator's batch was ordered first, a client rolls its pending edits back, applies that batch, then replays its own (transformed) — so every client ends up with the same document. Late joiners open the base file and replay the log.
- Structural changes (inserting / deleting rows or columns, the sheet list, the slide list) wait for the hub's order (one round trip).

| Format | Shared as                                                         | Same thing edited at once                                                                         |
|--------|-------------------------------------------------------------------|---------------------------------------------------------------------------------------------------|
| DOCX   | text splices per paragraph, paragraph and table states            | characters merge inside a paragraph; split / join / paragraph format: last writer wins            |
| XLSX   | cell values, formats, layout, fill / sort, row & column structure | last writer wins per cell and property; inserted / deleted rows and columns shift the other edits |
| PPTX   | shape text, shape edits, slide list                               | last writer wins per shape                                                                        |

## Limitations

- Undo / redo only touches your own changes
- No snapshots yet: join time grows with the history; start over by uploading a new base file
- Deleting rows / columns can't be undone while collaborating
- The reference server has no authentication — put it behind your own auth
