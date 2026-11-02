# API

## React

```tsx
import { DocViewer } from '@p10node/office-engine'

<DocViewer src={file} style={{ height: 600 }} />                 // File | Blob | ArrayBuffer | Uint8Array
<DocViewer src="https://cdn.example.com/report.xlsx" />          // URL → Range requests
<DocViewer src={url} headers={{ Authorization: `Bearer ${t}` }} credentials="include" />
<DocViewer src={file} editable onSave={(blob, name) => upload(blob, name)} />  // editing
```

CSS is auto-injected, no import needed. Container must have a height.

| Prop            | Type                                           | Default  | Notes                                                          |
|-----------------|------------------------------------------------|----------|----------------------------------------------------------------|
| `src`           | `File \| Blob \| ArrayBuffer \| string \| URL` | –        | Document source                                                |
| `name`          | `string`                                       | auto     | Display name                                                   |
| `toolbar`       | `boolean`                                      | `true`   | Top bar (name, status, progress)                               |
| `toolbarExtra`  | `ReactNode`                                    | –        | Extra buttons on the right of the toolbar                      |
| `zoom`          | `number \| 'fit'`                              | `'fit'`  | Initial zoom (docx/pptx)                                       |
| `thumbnails`    | `boolean`                                      | `true`   | Thumbnail column for pptx                                      |
| `headers`       | `Record<string,string>`                        | –        | Headers for URL requests                                       |
| `credentials`   | `RequestCredentials`                           | –        | Cross-origin cookies                                           |
| `worker`        | `() => Worker \| false`                        | inline   | Custom worker (CSP), `false` = main thread                     |
| `styleTarget`   | `Document \| ShadowRoot`                       | document | Where CSS is injected                                          |
| `editable`      | `boolean`                                      | `false`  | Open in edit mode                                              |
| `allowEdit`     | `boolean`                                      | `true`   | Show Edit/View button                                          |
| `onSave`        | `(blob, name) => void`                         | download | Receives saved file (downloads by default)                     |
| `onDirtyChange` | `(dirty) => void`                              | –        | Whether there are unsaved changes                              |
| `collab`        | `CollabOptions`                                | –        | Real-time collaboration, see [collaboration](collaboration.md) |
| `onLoad`        | `(client) => void`                             | –        |                                                                |
| `onError`       | `(error) => void`                              | –        |                                                                |

Use individual views / build your own UI:

```tsx
import { useDocument, DocxView, XlsxView, PptxView } from '@p10node/office-engine'

const { client, error, progress } = useDocument(src)
if (client?.kind === 'xlsx') return <div className="de-root"><XlsxView client={client} /></div>
```

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

`<office-viewer>`: attributes `src`, `name`, `editable`, `toolbar="false"`, `thumbnails="false"`, `collab` / `room` / `user` (real-time collaboration); property `.file = File`; events `load`, `error`.

## Core API (no React)

```ts
import { DocumentClient } from '@p10node/office-engine/core'

const doc = await DocumentClient.open(fileOrUrl, { onProgress: console.log })
doc.kind                                   // 'docx' | 'xlsx' | 'pptx'
doc.on('docx:blocks', (blocks) => {})      // docx: streamed blocks
await doc.call('docx.start')
doc.call('xlsx.load', 0)                   // xlsx: stream sheet
doc.on('xlsx:progress', (p) => {})
await doc.call('xlsx.block', 0, 1000, 1100) // formatted cells, rows [1000, 1100)
await doc.call('pptx.slide', 12)           // slide model with layout/master/theme resolved
await doc.call('xlsx.set', 0, [{ r: 4, c: 1, input: '=SUM(A1:A4)' }])
const blob = await doc.call('save')        // edited file
doc.destroy()
```

## CSP

The default worker is an inline blob → requires `worker-src blob:`. If that's not allowed, host `dist/office-engine.worker.js` (exported as `@p10node/office-engine/worker`) on the same origin:

```tsx
<DocViewer src={src} worker={() => new Worker('/static/office-engine.worker.js', { type: 'module' })} />
```
