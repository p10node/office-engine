# API

## React

```tsx
import { DocViewer } from '@p10node/office-engine'

<DocViewer src={file} style={{ height: 600 }} />                 // File | Blob | ArrayBuffer | Uint8Array
<DocViewer src="https://cdn.example.com/report.xlsx" />          // URL → Range requests
<DocViewer src={url} headers={{ Authorization: `Bearer ${t}` }} credentials="include" />
<DocViewer src={file} editable onSave={(blob, name) => upload(blob, name)} />  // editing
<DocViewer src={file} allowSave={false} allowDownload={false} />  // no Save/Download, no image download
<DocViewer src={file} fileMenu={false} />                         // no File menu
<DocViewer src={file} theme="dark" />                             // dark chrome; "auto" follows the OS setting
```

### File menu

| Item                      | DOCX                       | XLSX                           | PPTX                       |
|---------------------------|----------------------------|--------------------------------|----------------------------|
| New                       | blank document             | blank spreadsheet              | blank presentation         |
| Open… (`Ctrl+O`)          | any supported file         | any supported file             | any supported file         |
| Import CSV into new sheet | –                          | `.csv` / `.tsv`                | –                          |
| Save (`Ctrl+S`)           | ✓                          | ✓                              | ✓                          |
| Rename…                   | ✓                          | ✓                              | ✓                          |
| Download as (native)      | `.docx`                    | `.xlsx`                        | `.pptx`                    |
| PDF                       | pages as shown             | current sheet, fit to width    | one slide per page         |
| Web page                  | `.html`, images inlined    | `.html`                        | `.html`, images inlined    |
| Text                      | `.txt`, `.md`              | `.csv`, `.tsv` (current sheet) | `.txt` (incl. notes)       |
| Document details          | pages, words, characters   | sheets, used range             | slides, slide size         |
| Print (`Ctrl+P`)          | ✓                          | ✓                              | ✓                          |
| Host items, Close         | `fileMenuItems`, `onClose` | `fileMenuItems`, `onClose`     | `fileMenuItems`, `onClose` |

Open converts other formats: `.csv` / `.tsv` → workbook, `.txt` / `.md` / `.html` → document (headings, lists, bold / italic, links). Converted files start unsaved; Save writes the Office file.
New / Open replace the shown document until the host passes another `src` (unsaved changes are confirmed first; both are hidden during collaboration).
Rename changes the shown name and the name used by Save / Download (`onRename` receives it); it does not reload the document.
Document details also lists the file's core properties (title, author, created, modified…).
PDF goes through the browser print dialog ("Save as PDF"), so text stays selectable. `allowDownload={false}` also hides Download as / Print and blocks `Ctrl+P`.

```tsx
<DocViewer
  src={file}
  fileMenuItems={[{ label: 'Share…', icon: 'open', onClick: share }, { label: 'Version history', onClick: history }]}
  onRename={(name) => api.rename(id, name)}
  onClose={() => router.back()}
/>
```

CSS is auto-injected, no import needed. Container must have a height.

| Prop               | Type                                           | Default   | Notes                                                                                                |
|--------------------|------------------------------------------------|-----------|------------------------------------------------------------------------------------------------------|
| `src`              | `File \| Blob \| ArrayBuffer \| string \| URL` | –         | Document source                                                                                      |
| `name`             | `string`                                       | auto      | Display name                                                                                         |
| `toolbar`          | `boolean`                                      | `true`    | Top bar (name, status, progress)                                                                     |
| `toolbarExtra`     | `ReactNode`                                    | –         | Extra buttons on the right of the toolbar                                                            |
| `zoom`             | `number \| 'fit'`                              | `'fit'`   | Initial zoom (docx/pptx)                                                                             |
| `thumbnails`       | `boolean`                                      | `true`    | Thumbnail column for pptx                                                                            |
| `headers`          | `Record<string,string>`                        | –         | Headers for URL requests                                                                             |
| `credentials`      | `RequestCredentials`                           | –         | Cross-origin cookies                                                                                 |
| `worker`           | `() => Worker \| false`                        | inline    | Custom worker (CSP), `false` = main thread                                                           |
| `styleTarget`      | `Document \| ShadowRoot`                       | document  | Where CSS is injected                                                                                |
| `theme`            | `'light' \| 'dark' \| 'auto'`                  | `'light'` | Color scheme of the chrome; `'auto'` follows `prefers-color-scheme`. Documents keep their own colors |
| `editable`         | `boolean`                                      | `false`   | Open in edit mode (Pages, Keynote and Numbers files stay view only)                                  |
| `allowEdit`        | `boolean`                                      | `true`    | Show Edit/View button                                                                                |
| `allowSave`        | `boolean`                                      | `true`    | Show Save/Download button, enable Ctrl+S                                                             |
| `allowDownload`    | `boolean`                                      | `true`    | Allow downloads to device (save without `onSave`, images)                                            |
| `fileMenu`         | `boolean`                                      | `true`    | File menu: open, download as PDF / HTML / text, print                                                |
| `allowOpen`        | `boolean`                                      | `true`    | File → New / Open (off during collaboration)                                                         |
| `captureSelectAll` | `boolean`                                      | `true`    | Ctrl/⌘+A from anywhere on the page selects the document; `false`: from the viewer only               |
| `fileMenuItems`    | `FileMenuItem[]`                               | –         | Extra File menu entries (`label`, `onClick`, `icon`, `hint`)                                         |
| `onRename`         | `(name) => void`                               | –         | File → Rename result (extension included)                                                            |
| `onClose`          | `() => void`                                   | –         | Adds File → Close                                                                                    |
| `onSave`           | `(blob, name) => void`                         | download  | Receives saved file (downloads by default)                                                           |
| `onDirtyChange`    | `(dirty) => void`                              | –         | Whether there are unsaved changes                                                                    |
| `collab`           | `CollabOptions`                                | –         | Real-time collaboration, see [collaboration](collaboration.md)                                       |
| `comments`         | `CommentOptions \| false`                      | `{}`      | Comments (see below); `false` hides them                                                             |
| `onLoad`           | `(client) => void`                             | –         |                                                                                                      |
| `onError`          | `(error) => void`                              | –         |                                                                                                      |

### Comments

```tsx
<DocViewer
  src={file}
  editable
  comments={{
    user: { id: me.id, name: me.name },                              // author of the comments written here
    mentions: (q) => api.members(nodeId, q),                         // or a list: [{ id, name }, …]
    onEvent: (e) => e.mentioned.length && api.notify(e.mentioned, e), // notify people mentioned
    open: true,                                                      // panel open on load
  }}
/>
```

| `CommentOptions` | Type                                                      | Default                      |
|------------------|-----------------------------------------------------------|------------------------------|
| `user`           | `{ name, id?, initials? }`                                | the `collab` user, else "Me" |
| `mentions`       | `CommentUser[] \| (query) => CommentUser[] \| Promise<…>` | none (no `@` suggestions)    |
| `canComment`     | `boolean`                                                 | while editing                |
| `onEvent`        | `(e: CommentEvent) => void`                               | –                            |
| `open`           | `boolean`                                                 | `false`                      |
| `focus`          | `string` (a comment id)                                   | –                            |

`focus` shows a thread when the document opens (a link from a notification): the panel opens on it and the document scrolls to its place.

`onEvent` is called for the changes made in this viewer (not for collaborators' — each of them reports their own):

| `CommentEvent` | Meaning                                                                                     |
|----------------|---------------------------------------------------------------------------------------------|
| `type`         | `'add'`, `'reply'`, `'edit'`, `'delete'`, `'resolve'`, `'reopen'`                           |
| `comment`      | the comment (`DocComment`: `id`, `parent?`, `author`, `time`, `text`, `mentions?`, `done?`) |
| `thread`       | the thread's first comment                                                                  |
| `mentioned`    | people mentioned by this change: all of a new comment's, those an edit added                |
| `participants` | the people who wrote the thread's other comments (not the one who made the change)          |
| `quote`        | what the thread points at: the commented text, `Sheet1!B3`, `Slide 2`                       |

A comment's author and mentions carry the `id` the host gave (stored in the file's people / persons / authors list), so a host can tell its users apart after the file was saved and reopened. Ownership (edit / delete) is by `user.id` when there is one, else by name.

Use individual views / build your own UI:

```tsx
import { useDocument, DocxView, XlsxView, PptxView } from '@p10node/office-engine'

const { client, error, progress } = useDocument(src)
if (client?.kind === 'xlsx') return <div className="de-root"><XlsxView client={client} /></div>
```

The views style themselves from CSS variables on `.de-root` (`--de-bg`, `--de-chrome`, `--de-accent`…): override them to match your app. Add `de-dark` to the root element for the dark theme.

## Without a React build

```html
<script src="office-engine.embed.js"></script>   <!-- dist/, bundles React + worker -->

<office-viewer src="/files/report.docx" style="height:600px"></office-viewer>

<div id="v" style="height:600px"></div>
<script>
  const viewer = OfficeEngine.mount(document.getElementById('v'), { src: '/files/deck.pptx' })
  input.onchange = (e) => viewer.update({ src: e.target.files[0] })
  // viewer.destroy()
</script>
```

`<office-viewer>`: attributes `src`, `name`, `editable`, `toolbar="false"`, `thumbnails="false"`, `save="false"`, `download="false"`, `file-menu="false"`, `open="false"`, `capture-select-all="false"`, `theme="dark"` (or `"auto"`), `collab` / `room` / `user` / `base` (real-time collaboration; `user` also writes the comments), `comments="false"`; property `.file = File`; events `load`, `error`, `comment` (`detail`: a `CommentEvent`).

## Core API (no React)

```ts
import { DocumentClient } from '@p10node/office-engine/core'

const doc = await DocumentClient.open(fileOrUrl, { onProgress: console.log })
doc.kind                                   // 'docx' | 'xlsx' | 'pptx'
doc.info.from                              // 'pages' / 'keynote' / 'numbers': an Apple file, converted to docx / pptx / xlsx (view only)
doc.on('docx:blocks', (blocks) => {})      // docx: streamed blocks
await doc.call('docx.start')
doc.call('xlsx.load', 0)                   // xlsx: stream sheet
doc.on('xlsx:progress', (p) => {})
await doc.call('xlsx.block', 0, 1000, 1100) // formatted cells, rows [1000, 1100)
await doc.call('pptx.slide', 12)           // slide model with layout/master/theme resolved
await doc.call('xlsx.set', 0, [{ r: 4, c: 1, input: '=SUM(A1:A4)' }])
const comments = await doc.call('comments.list')  // DocComment[]: threads are roots + replies (threadsOf)
doc.on('comment', ({ id, c }) => {})       // a comment set / deleted (c null), here or by a collaborator
await doc.call('comments.set', c.id, { ...c, done: true }) // XLSX / PPTX threads carry their place (`at`)
const blob = await doc.call('save')        // edited file
doc.destroy()
```

## CSP

The default worker is an inline blob → requires `worker-src blob:`. If that's not allowed, host `dist/office-engine.worker.js` (exported as `@p10node/office-engine/worker`) on the same origin:

```tsx
<DocViewer src={src} worker={() => new Worker('/static/office-engine.worker.js', { type: 'module' })} />
```
