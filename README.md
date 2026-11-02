# @eldora/docs-engine

Engine xem **DOCX / XLSX / PPTX** trực tiếp trên browser. Load từng phần như Google Docs, chịu được file nặng, React component nhúng 1 dòng.

## Tính năng

- **Lazy ZIP**: chỉ đọc central directory, giải nén từng part khi cần (`DecompressionStream` native, 0 dependency)
- **HTTP Range**: mở URL → chỉ tải byte cần thiết (block cache 256 KB); server không hỗ trợ Range → tự fallback tải full
- **Streaming XML parser** chạy trong **Web Worker** → UI không bao giờ bị block
- **Progressive render**:
  - DOCX: block (paragraph/table) stream dần → đo chiều cao offscreen theo time-slice → phân trang → chỉ render các trang đang nhìn thấy
  - XLSX: row stream dần, grid ảo hoá 2 chiều (chỉ render ô trong viewport), freeze panes, merge, number format, sheet tab
  - PPTX: slide parse on-demand khi cuộn tới, thumbnail ảo hoá
- Nhúng kiểu: React component, hook, `mount()` vanilla, `<docs-viewer>` web component (shadow DOM), 1 file `<script>`

## Hiệu năng đo được (Chromium headless, M-series)

| File                     | Kích thước | First paint | Load xong            |
| ------------------------ | ---------- | ----------- | -------------------- |
| DOCX 20k đoạn văn        | 1.2 MB     | ~180 ms     | 2073 trang trong ~2s |
| XLSX 300k rows × 12 cols | 24.8 MB    | ~330 ms     | ~2s                  |
| PPTX 305 slides          | 0.3 MB     | ~110 ms     | on-demand            |

## Cài đặt

```bash
bun add @eldora/docs-engine   # peer: react >= 18, react-dom >= 18
```

## Dùng với React

```tsx
import { DocViewer } from '@eldora/docs-engine'

<DocViewer src={file} style={{ height: 600 }} />                 // File | Blob | ArrayBuffer | Uint8Array
<DocViewer src="https://cdn.example.com/report.xlsx" />          // URL → Range requests
<DocViewer src={url} headers={{ Authorization: `Bearer ${t}` }} credentials="include" />
```

CSS tự inject, không cần import. Container cần có chiều cao.

| Prop           | Kiểu                                           | Mặc định | Ghi chú                                       |
| -------------- | ---------------------------------------------- | -------- | --------------------------------------------- |
| `src`          | `File \| Blob \| ArrayBuffer \| string \| URL` | –        | Nguồn tài liệu                                |
| `name`         | `string`                                       | auto     | Tên hiển thị                                  |
| `toolbar`      | `boolean`                                      | `true`   | Thanh trên (tên, trạng thái, progress)        |
| `toolbarExtra` | `ReactNode`                                    | –        | Nút thêm bên phải toolbar                     |
| `zoom`         | `number \| 'fit'`                              | `'fit'`  | Zoom ban đầu (docx/pptx)                      |
| `thumbnails`   | `boolean`                                      | `true`   | Cột thumbnail cho pptx                        |
| `headers`      | `Record<string,string>`                        | –        | Header cho request URL                        |
| `credentials`  | `RequestCredentials`                           | –        | Cookie cross-origin                           |
| `worker`       | `() => Worker \| false`                        | inline   | Worker tuỳ chỉnh (CSP), `false` = main thread |
| `styleTarget`  | `Document \| ShadowRoot`                       | document | Nơi inject CSS                                |
| `onLoad`       | `(client) => void`                             | –        |                                               |
| `onError`      | `(error) => void`                              | –        |                                               |

Dùng từng view riêng / tự dựng UI:

```tsx
import { useDocument, DocxView, XlsxView, PptxView } from '@eldora/docs-engine'

const { client, error, progress } = useDocument(src)
if (client?.kind === 'xlsx') return <div className="de-root"><XlsxView client={client} /></div>
```

## Nhúng không cần React build

```html
<script src="docs-engine.embed.js"></script>   <!-- dist/, đã gồm React + worker -->

<docs-viewer src="/files/report.docx" style="height:600px"></docs-viewer>

<div id="v" style="height:600px"></div>
<script>
  const viewer = DocsEngine.mount(document.getElementById('v'), { src: '/files/deck.pptx' })
  input.onchange = (e) => viewer.update({ src: e.target.files[0] })
  // viewer.destroy()
</script>
```

`<docs-viewer>`: attribute `src`, `name`, `toolbar="false"`, `thumbnails="false"`; property `.file = File`; event `load`, `error`.

## Core API (không React)

```ts
import { DocumentClient } from '@eldora/docs-engine/core'

const doc = await DocumentClient.open(fileOrUrl, { onProgress: console.log })
doc.kind                                   // 'docx' | 'xlsx' | 'pptx'
doc.on('docx:blocks', (blocks) => {})      // docx: stream block
await doc.call('docx.start')
doc.call('xlsx.load', 0)                   // xlsx: stream sheet
doc.on('xlsx:progress', (p) => {})
await doc.call('xlsx.block', 0, 1000, 1100) // ô đã format, rows [1000, 1100)
await doc.call('pptx.slide', 12)           // model slide đã resolve layout/master/theme
doc.destroy()
```

## Kiến trúc

```
main thread                                   worker
───────────                                   ──────
DocViewer ─ useDocument ─ DocumentClient ⇄ RPC ⇄ Engine
  ├ DocxView  (measure → paginate → virtual pages)     ├ ByteSource (Blob slice | HTTP Range + LRU)
  ├ XlsxView  (Axis prefix-sum, 2D virtual grid)       ├ ZipArchive (central dir, inflate on demand)
  └ PptxView  (virtual slides + thumbs)                ├ SaxParser / TreeBuilder (streaming)
                                                       └ DocxDocument | XlsxWorkbook | PptxPresentation
```

```
src/core     parser, không phụ thuộc DOM (chạy trong worker / Bun)
src/react    view components
src/embed.ts mount() + <docs-viewer>
```

## Hỗ trợ

| Định dạng | Có                                                                                                                                                             | Chưa                                                            |
| --------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------- | --------------------------------------------------------------- |
| DOCX      | style kế thừa, run format, list/numbering, bảng (merge, border, shading, table style), ảnh, hyperlink, header/footer, field PAGE, text box, page/section break | footnote body, chart, float wrap chính xác, tách đoạn qua trang |
| XLSX      | shared strings, number/date format, font/fill/border/align, merge, freeze, cột/hàng ẩn, wrap, formula (hiển thị), nhiều sheet                                  | tính lại formula, chart, conditional format, ảnh                |
| PPTX      | master/layout/theme, placeholder kế thừa, ~50 preset shape, custGeom, ảnh (crop), bảng + table style, group, bullet/autonum, autofit                           | chart (placeholder), SmartArt, animation, video playback        |

Không hỗ trợ `.doc/.xls/.ppt` (binary cũ) và file mã hoá.

## Phát triển

```bash
bun install
bun run dev          # playground http://localhost:5173 (drag & drop, URL, sample)
bun test             # test core trên fixture
bun run build        # dist/: ESM lib + docs-engine.embed.js + .d.ts
python3 -m venv .venv && .venv/bin/pip install python-docx openpyxl python-pptx Pillow
.venv/bin/python scripts/fixtures.py --large   # sinh sample + file nặng
```

## CSP

Worker mặc định là blob inline → cần `worker-src blob:`. Nếu không cho phép, host file `dist/docs-engine.worker.js` (export `@eldora/docs-engine/worker`) cùng origin:

```tsx
<DocViewer src={src} worker={() => new Worker('/static/docs-engine.worker.js', { type: 'module' })} />
```
