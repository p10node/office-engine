# @p10node/convert

Office files converted on a server, without LibreOffice: Word, Excel and PowerPoint 97-2003, RTF, OpenDocument, Apple Pages / Keynote / Numbers, Visio, CSV, Markdown, HTML, diagrams and source files to **DOCX / XLSX / PPTX**, presentations to **PDF**, and any document to **text, Markdown, HTML or CSV**. A function, a command (`office-convert`) and an HTTP service that answers on Gotenberg's routes. Node ≥ 22 or Bun, no dependency, no browser.

Part of [office-engine](https://github.com/p10node/office-engine): the same engine that opens these files in the browser, on worker threads.

**PDF is written for presentations** (.pptx, .ppt, .odp, Keynote, Visio), by the engine itself ([#44](https://github.com/p10node/office-engine/issues/44)): Gotenberg's `/forms/libreoffice/convert` answers with their PDF. Spreadsheets ([#47](https://github.com/p10node/office-engine/issues/47)) and documents ([#48](https://github.com/p10node/office-engine/issues/48)) follow; until then it answers `501` for them.

## Install

```sh
npm i -g @p10node/convert      # or: bunx @p10node/convert …
```

## Command

```sh
office-convert report.doc                       # report.docx next to it
office-convert report.doc out/report.docx       # the format from the output's extension (an output that exists: -o and --force)
office-convert *.xls *.ods -o converted/        # several at once, into a directory
office-convert sales.xlsx -t csv --sheet Q3     # a sheet as CSV
office-convert deck.pptx deck.pdf --slides 1-3  # slides as PDF
cat notes.md | office-convert - --name notes.md -t html -o - > notes.html
office-convert serve -p 3000                    # the HTTP service
```

| Option            | Default                                                    |
|-------------------|------------------------------------------------------------|
| `-t, --to`        | the Office format the file opens as (.doc → .docx, .csv → .xlsx) |
| `-o, --out`       | next to each input; a file (one input, `-` for stdout) or a directory |
| `--sheet`         | CSV / TSV: the sheet the file was saved on; an index or a name |
| `--encoding`      | text files: a guess (UTF-8, UTF-16, Windows-1252…)         |
| `-f, --force`     | existing files are not overwritten; an input never is, under any name |
| `-j, --jobs`      | files at once: the number of CPUs                          |
| `--timeout`       | 120 s per file (0: no limit)                               |
| `--max-size`      | 100 MB per file                                            |
| `--slides`        | PDF: every slide; `1-3,5` for those                        |
| `--hidden`        | PDF: hidden slides are left out, also when `--slides` names them |
| `--fonts`         | PDF: a font file or folder to set text in, before the system's (again for more) |
| `--no-system-fonts` | PDF: the system's fonts are used                         |

The exit status is 1 when a file failed (the others are still written), 2 for a wrong command line.

## What converts to what

| Input                                                                  | Opens as | Can be written as            |
|------------------------------------------------------------------------|----------|------------------------------|
| .docx .docm .dotx, .doc, .rtf, .odt .ott .fodt, .pages, .eml .msg*, .md, .txt, .html, .mmd .puml .dot, source files | `docx` | `docx`, `txt`, `md`, `html` |
| .xlsx .xlsm .xltx, .xls, .ods .ots .fods, .numbers, .csv .tsv         | `xlsx`   | `xlsx`, `csv`, `tsv`         |
| .pptx .pptm .potx .ppsx, .ppt, .odp .otp .fodp, .key, .vsdx           | `pptx`   | `pptx`, `pdf`, `txt`         |

\* once e-mail messages are in the engine ([#55](https://github.com/p10node/office-engine/pull/55)).

PDF of a document or a workbook is refused with `not-implemented` until #48 and #47.

The format is read from the file's bytes; the name only helps for text formats. An Office file keeps its variant (a `.docm` stays `.docm`). CSV output has the values as shown (formats applied), formulas computed when the file has no value for them, and a byte order mark so Excel reads it as UTF-8.

## HTTP service

```sh
docker build -f packages/convert/Dockerfile -t office-convert .   # from the repository root
docker run --rm -p 3000:3000 office-convert
```

```sh
curl -F files=@report.doc http://localhost:3000/convert -o report.docx
curl -F files=@a.odt -F files=@b.xls http://localhost:3000/convert -o converted.zip
curl -F files=@sales.xlsx -F to=csv -F sheet=Q3 http://localhost:3000/convert -o q3.csv
curl --data-binary @notes.md 'http://localhost:3000/convert?to=html&name=notes.md' -o notes.html
curl -F files=@deck.pptx -F nativePageRanges=1-3 http://localhost:3000/forms/libreoffice/convert -o deck.pdf   # as Gotenberg
```

| Route                              |                                                                                  |
|------------------------------------|----------------------------------------------------------------------------------|
| `POST /convert`                    | multipart: the files (any field name), `to`, `sheet`, `encoding`, `language`, `slides` (PDF: `1-3,5`), `hidden`; or one file as the body, `to` / `name` in the query |
| `POST /forms/libreoffice/convert`  | Gotenberg's route: the files to PDF, `nativePageRanges` the slides, `exportHiddenSlides`. Presentations; `501` for documents and spreadsheets. What it is asked for and does not do yet is answered `400`, not ignored: `merge` of several files, `pdfa`, `pdfua`, `userPassword`, `ownerPassword`, `splitMode`, `password`. Booleans as Gotenberg reads them (`1`, `t`, `true`…); several files give a ZIP of `deck.pptx.pdf`… |
| `GET /health`                      | `{ "status": "up", "details": … }`, as Gotenberg answers                         |
| `GET /version`                     | the version, as text                                                             |

One file gives the converted file, several a ZIP. Gotenberg's headers are read: `Gotenberg-Output-Filename` names the result, `Gotenberg-Trace` comes back (a new id when there is none). What a file lost on the way (more columns than a sheet holds…) is in `Office-Convert-Warning` headers, percent-encoded.

| Status | When                                                                 |
|--------|----------------------------------------------------------------------|
| 400    | no file, too many files, a format the file cannot be written as, slides that are not there, what Gotenberg's route is asked for and does not do yet |
| 405    | not a POST                                                           |
| 413    | the request is larger than `MAX_SIZE`, or the converted file would be larger than 1 GB |
| 422    | the file cannot be read: damaged, encrypted, not a format the engine knows (`400` on Gotenberg's route, as Gotenberg answers) |
| 501    | PDF of a document or a spreadsheet (#48, #47); answered before the file is converted when its bytes say what it is (an Office or OpenDocument file, RTF), whatever its name |
| 503    | the file took longer than `TIMEOUT`, or every thread is busy and the queue is full (`Retry-After`) |

Settings, as flags of `office-convert serve` or environment variables: `PORT` (3000), `HOST` (0.0.0.0), `WORKERS` (the number of CPUs), `TIMEOUT` (120 s), `MAX_SIZE` (100 MB per request), `MAX_FILES` (20 per request), `FONTS` (font files or folders for PDF, before the system's), `SYSTEM_FONTS=0` (not the system's). Each conversion runs on a worker thread: one past its time, or whose client hung up, ends with its thread, and when one file of a request fails the others are given up. At most 100 files, and 1 GB of them, wait for a thread, and 1 GB of request bodies are held at once (`maxPending` of `createHandler`); more is answered `503`. A gzip-compressed file unpacks to 1 GB at most. The engine's own limits for crafted files (ZIP entries, unpacked sizes, cells, chart points) apply as in the browser.

## Fonts

A PDF's text is set in font files and embedded as subsets, so it can be selected, copied and searched. The fonts are the system's (`/usr/share/fonts`, `~/.fonts`, macOS's and Windows' font folders), read once for their names, styles and weights, and those of `--fonts` / `FONTS` before them; a file is read from the disk when text first needs it. A document's font is used when it is installed, else the one its font list falls back to: Carlito for Calibri, Caladea for Cambria, Liberation for Arial, Times New Roman and Courier New (the same widths, so lines break where PowerPoint breaks them), then the generic family (`sans-serif` is Liberation Sans, Arial or DejaVu Sans, then fonts of other scripts such as Noto Sans CJK). Text no font serves is set in Helvetica, with a warning that names the font.

The Docker image has Carlito, Caladea, Liberation and DejaVu. Fonts of other scripts are a build argument away (`--build-arg EXTRA_FONTS=fonts-noto-cjk`), the fonts a company uses a volume away (`-v ./fonts:/fonts -e FONTS=/fonts`).

## API

```ts
import { convert, Converter, createHandler, findFonts, listen } from '@p10node/convert'

// on this thread
const { data, name, type } = await convert(bytes, { name: 'report.doc', to: 'docx' })
const pdf = await convert(deck, { name: 'deck.pptx', to: 'pdf', slides: [0, 2], fonts: await findFonts(['/srv/fonts']) })

// on a pool of worker threads, with a time limit per file
const converter = new Converter({ workers: 4, timeout: 60_000 })
const csv = await converter.convert(bytes, { name: 'sales.xlsx', to: 'csv', sheet: 'Q3', signal })

// the HTTP routes as a fetch handler: Bun.serve({ fetch }), Hono, Deno, a Worker…
const handler = createHandler(converter, { maxBytes: 50 * 1024 * 1024 })
await listen(handler, { port: 3000 })   // Bun.serve under Bun, node:http under Node
```

`convert` resolves with `{ data: Uint8Array, name, type, ext, kind, from?, warnings }` and rejects with a `ConvertError` (`code`: `unsupported`, `not-implemented`, `too-large`) for a request it cannot do, an `Error` for a file it cannot read. Its limits: `maxOutput` (the file written, 1 GB) and `maxUnpacked` (a gzip-compressed input, 1 GB). For PDF: `fonts` (default `systemFonts()`), `slides` (indexes from 0), `hidden` (hidden slides are written only with it, also when `slides` names them). `Converter` takes `workers`, `timeout`, `queue` (100), `bytes` (1 GB of input running and waiting), `memory` (MB of heap per thread, Node only) and `fonts` (found once in the calling thread and handed to the threads; a function is called when a PDF is first asked for), copies the bytes it is given, and adds `TimeoutError`, `BusyError` and `ThreadError` (the thread ran out of memory or crashed).
