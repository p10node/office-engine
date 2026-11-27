# Development

```bash
bun install
bun run dev          # playground http://localhost:5173 (drag & drop, URL, samples)
bun run collab       # collaboration hub for the playground's "Share live" (port 8787)
bun test             # core tests on fixtures
bun run test:e2e     # browser tests (headless Chromium: bunx playwright install chromium; needs large-deck.pptx, see fixtures.py)
bun run build        # dist/: ESM lib + office-engine.embed.js + .d.ts
bun run build:packages  # packages/*/dist: the standalone SDK packages (@p10node/diagram, mermaid, plantuml, graphviz, markdown, convert)
python3 -m venv .venv && .venv/bin/pip install python-docx openpyxl python-pptx Pillow
.venv/bin/python scripts/fixtures.py --large   # generate samples + heavy files (--large-deck: only the deck the browser tests need)
bun scripts/iwork-samples.ts                   # sample.pages / .key / .numbers (built by test/iwork-builder.ts)
bun scripts/visio-samples.ts                   # sample.vsdx (built by test/visio-builder.ts)
bun scripts/legacy-samples/doc.ts              # .doc / .ppt / .rtf test fixtures and samples (also ppt.ts, rtf.ts; .xls / .odt / .ods / .odp from the *.py scripts: see scripts/legacy-samples/*.md)
bun run screenshots  # README screenshots -> docs/screenshots/ (headless Chromium)
bun run gifs         # README GIFs (300k-row sheet opening, collaboration) -> docs/screenshots/*.gif (headless Chromium + ffmpeg)
```

## Architecture

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
src/core     parsers, DOM-free (runs in worker / Bun)
  legacy     .doc / .xls / .ppt (compound file reader + converters), .rtf, OpenDocument → DOCX / XLSX / PPTX
  iwork      Pages / Keynote / Numbers → DOCX / PPTX / XLSX
  visio      Visio drawings (.vsdx) → PPTX
  math       TeX math: `$…$` / `$$…$$` in Markdown, TeX → MathML (views, web pages) and → OMML (Word file)
  font       font files read for layout: glyph widths, kerning, ligatures, line metrics, subsets (no DOM)
src/react    view components
src/embed.ts mount() + <office-viewer>
src/collab   real-time collaboration (hub, session, per-format adapters)
```

### Font files

`src/core/font` reads font files without a DOM (`container.ts`: TTF, OTF, TTC, WOFF, WOFF2; `font.ts`: character map, widths, metrics; `layout.ts`: GSUB, GPOS, kern; `text.ts`: a run of text as glyphs; `book.ts`: which declared file serves a text; `subset.ts`: subsets and ToUnicode). What it does not apply (contextual rules, marks, Apple's tables, scripts that need shaping) it detects, and says so in `exact`: nothing passes for exact that a shaping engine would set differently. Unit tests (`test/font-read.test.ts`) use fonts written by `test/font-tables.ts`; `e2e/font-read.spec.ts` checks the widths against the browser's own layout. On real fonts, which the repository does not hold, it is checked by hand against HarfBuzz:

```sh
bun scripts/font-check.ts /usr/share/fonts/truetype/*/*.ttf "/System/Library/Fonts/Helvetica.ttc#1"   # needs hb-shape and hb-view
```

### DOCX pages

`DocxView` measures the blocks off screen in time slices (`src/react/docx/measure.ts`: a block's height, where a paragraph's lines and a table's rows are), `paginate()` cuts them into pages (`paginate.ts`: pure, lines and rows given, keep rules; unit tests in `test/docx-paginate.test.ts`) and only the pages in view render. The browser breaks the lines; the offset each line starts at is read (`lineStarts`) only for the paragraphs a page in view or an edit cuts. A paragraph split across pages renders as one element per part, each with the paragraph's `data-pa` address and the offsets it holds (`data-from`, `data-to`): `paraEl()` / `posFromDom()` / `domFromPos()` in `src/react/edit/input.ts` map DOM places to paragraph offsets over the parts, so the caret, selections, find and comments need nothing else. Header rows repeated on the following pages of a table are copies without an address (`.de-head-copy`).

A document longer than a browser lets a box scroll (12,000,000 px) is scrolled through `ScrollMap` (`src/react/docx/longscroll.ts`, unit tests in `test/docx-longscroll.test.ts`): the place shown is the scroll position plus an offset the pages are drawn by. A step of the reader or of the browser leaves the offset alone, so the browser moves the pages by itself; a jump of the scrollbar's thumb sets it; at rest the box is put where its thumb belongs, pages and scroll position in one frame.

## Browser tests

`e2e/*.spec.ts` drive a fixture app (`e2e/app`, served by Vite with `vite.e2e.config.ts`) in headless Chromium through `playwright-core`, one browser and server per file (`e2e/harness.ts`).
The app reads `DocViewer` props from the query string (`?src=samples/sample.docx&editable=1&allowDownload=0&onSave=1…`) and records callbacks in `window.__events`; saved files are re-opened with the engine to check what was written.

| Spec                       | Covers                                                                                                                                   |
|----------------------------|------------------------------------------------------------------------------------------------------------------------------------------|
| `viewer.spec.ts`           | rendering of the 3 formats, outline, zoom, sheet tabs, thumbnails, slideshow, props, errors                                              |
| `legacy.spec.ts`           | .doc / .xls / .ppt / .rtf / OpenDocument samples: converted, shown, edited, saved as OOXML                                               |
| `docx.spec.ts`             | typing, formatting, styles, lists, colors, fonts, links, tables, images, find & replace, undo                                            |
| `docx-page.spec.ts`        | page setup: size, orientation, margins, columns; refused values; numbers typed in a locale                                               |
| `docx-split.spec.ts`       | paragraphs and tables split across pages: layout, caret, editing, header rows                                                            |
| `docx-longscroll.spec.ts`  | a document past the scroll cap: wheel, Page Down, the caret, the last page by steps, the box at rest, zoom                               |
| `docx-hf.spec.ts`          | headers and footers: page numbers, text, different first page, saved and undone                                                          |
| `docx-table.spec.ts`       | table editing: cell selection, merges, colors, borders, alignment, column widths                                                         |
| `docx-tabs.spec.ts`        | tab stops: left / center / right / decimal, leaders, stops of a style, absolute position tabs, typing, HTML export                       |
| `docx-track.spec.ts`       | tracked changes: shown, accepted, rejected, typed while tracking                                                                         |
| `docx-qa2.spec.ts`         | Word editor regressions (QA 2): typing across pages, lists, links, clipboard, tables, find, print                                        |
| `docx-qa3.spec.ts`         | Word editor regressions (QA 3): text boxes anchored in a paragraph, Tab in a list that defines one level                                 |
| `xlsx.spec.ts`             | values, formulas, merges (all / across / down / unmerge), formats, rows / columns, sheets, freeze                                        |
| `xlsx-qa3.spec.ts`         | spreadsheet regressions (QA 3): whole columns / rows next to merged cells, header menus, Undo state, deleting a sheet                    |
| `pptx.spec.ts`             | text, shapes (insert, move, fill, arrange, rotate, duplicate, delete), images, slides, present                                           |
| `pptx-save.spec.ts`        | saving while slide text is being typed or committed: nothing pending is left out                                                         |
| `pptx-qa2.spec.ts`         | presentation regressions (QA 2): Tab levels, undo of slide list edits, table cell text, notes export                                     |
| `pptx-qa3.spec.ts`         | presentation regressions (QA 3): undo / redo of text typed in a shape, the state of the buttons, typing and the slide list               |
| `focus-qa3.spec.ts`        | the keyboard after a dialog, a thumbnail or a sheet tab: focus trap, focus return, Ctrl+S                                                |
| `file-menu.spec.ts`        | every File entry: New, Open (conversions), Import, Save, Rename, Download as, Details, Print                                             |
| `permissions.spec.ts`      | `allowSave` / `allowDownload` combinations, image download, Ctrl+S / Ctrl+P                                                              |
| `color.spec.ts`            | custom colors (system picker, recent list, persistence) in the 3 formats                                                                 |
| `code.spec.ts`             | code blocks in Markdown and source files: syntax colors, editing, bytes and line breaks kept                                             |
| `markdown-math.spec.ts`    | TeX math in Markdown: drawn as MathML, edited around, saved as written, printed and exported                                             |
| `comments.spec.ts`         | comments in the 3 formats: on a selection / cell / slide, replies, mentions, resolve, undo, saved                                        |
| `comments-collab.spec.ts`  | comments in a shared room                                                                                                                |
| `collab.spec.ts`           | two editors on `scripts/collab-server.ts`: live edits, presence, File menu in a room                                                     |
| `collab-qa3.spec.ts`       | two editors in a room: a slide added by one does not end the text the other types; rows inserted by one move the cell the other types in |
| `security.spec.ts`         | crafted files: nothing is requested from another host, no script runs, SVG is not a same-origin document                                 |
| `embed.spec.ts`            | `<office-viewer>` (shadow DOM, attributes, `file`, `fonts`) and `mount()`                                                                |
| `sdk.spec.ts`              | the standalone packages (diagram, mermaid, plantuml, graphviz, markdown) in a page                                                       |
| `fonts.spec.ts`            | host fonts (a generated test font): loaded before measuring, timeout, failure, late web fonts                                            |
| `font-read.spec.ts`        | fonts read from their files: widths against the browser's layout, subsets and unpacked WOFF2 as fonts a browser takes                    |
| `icons.spec.ts`            | host icons in the 3 formats; the built-in ones without `icons`                                                                           |
| `components.spec.ts`       | host UI components: Button, More, Dialog (prompts) in place of the engine's                                                              |
| `ui-layout.spec.ts`        | toolbar overflow (More), host dialogs: focus, submit once                                                                                |
| `controller.spec.ts`       | the controller and headless mode: state, commands, file actions, replaced documents                                                      |
| `chrome.spec.ts`           | host-placed controls (`ZoomBar`, `FormatBar`, `OutlineList`, `SheetTabs`…) around a headless viewer                                      |
| `format-controls.spec.ts`  | host formatting controls: focus kept while typing a size, the logical target of commands                                                 |
| `locale.spec.ts`           | UI texts: German, partial overrides, switching, unknown locales                                                                          |
| `locale-docx.spec.ts`      | documents in German: status, toolbar, page setup                                                                                         |
| `locale-xlsx.spec.ts`      | spreadsheets in a locale: sheet menu texts, numbers typed with a decimal comma, the size box                                             |
| `controller-built.spec.ts` | the built package: the controller across locale / UI updates, a replaced document, save                                                  |
| `format-built.spec.ts`     | the built package: host formatting through the real worker, saved and reopened                                                           |
| `locale-built.spec.ts`     | the built package: browser default formatting, opt-in German, plural forms, clipboard and Visio texts                                    |
| `ui-built.spec.ts`         | the built package: nested providers; every entry imported without a DOM                                                                  |

The `*-built.spec.ts` files run against `dist/`: `bun run build` first (they refuse a stale build).

`E2E_BROWSER=firefox bun run test:e2e` (or `webkit`) runs the suite in Playwright's Firefox / WebKit (`bunx playwright install firefox`): a look at another engine. The suite is kept green in Chromium; what differs elsewhere is listed in [Features](features.md#browsers).

`bun test` only runs `test/` (`bunfig.toml`); `bun run test:e2e` runs `e2e/`.

## Site (GitHub Pages)

`.github/workflows/pages.yml` deploys on every push to `main`:

- [p10node.github.io/office-engine](https://p10node.github.io/office-engine/): docs rendered from `README.md` + `docs/*.md`
- [p10node.github.io/office-engine/demo/](https://p10node.github.io/office-engine/demo/): the playground (heavy samples generated in CI)

One-time setup: **Settings → Pages → Source: GitHub Actions**. Build locally: `bun run build:playground && bun run build:site` → `dist-site/`.

"Share live" needs a collaboration hub; the static demo has none, so pass one: `?collab=https://your-hub.example`.

## Release

### Release process

CI (`.github/workflows/ci.yml`) runs typecheck, tests, the library and SDK package builds and `npm pack --dry-run` of each package, and the browser tests, on every push to `main` and every PR.

To release, push a branch named after the version:

```sh
git switch -c v0.2.0 main
git push -u origin v0.2.0
```

`.github/workflows/release.yml` then runs CI, sets `version` in `package.json` and every `packages/*/package.json` to match the branch (committing `chore(release): v0.2.0` onto the branch if any differs), pushes the `v0.2.0` tag, creates the GitHub release with the npm tarballs attached, and publishes to npm with provenance.

Released together, at one version (the SDK packages compile in the engine's code):

| Package                  | From                |
|--------------------------|---------------------|
| `@p10node/office-engine` | repository root     |
| `@p10node/diagram`       | `packages/diagram`  |
| `@p10node/mermaid`       | `packages/mermaid`  |
| `@p10node/plantuml`      | `packages/plantuml` |
| `@p10node/graphviz`      | `packages/graphviz` |
| `@p10node/markdown`      | `packages/markdown` |
| `@p10node/convert`       | `packages/convert`  |

A package whose version is already on npm is skipped; one that fails to publish does not stop the others (the job fails at the end). Re-running the failed job publishes what is missing; so does publishing the tag again by hand once its branch is gone or its run is old: **Actions → Release → Run workflow**, with the tag (`gh workflow run release.yml -f tag=v0.2.0`). Typical case: a package's trusted publisher was fixed after the release ran.

- Release notes: optional `docs/releases/<tag>.md`, placed above GitHub's generated commit list.
- Prereleases (`v0.2.0-rc.1`) are marked prerelease on GitHub and go to the `next` dist-tag; others go to `latest`.
- Re-running the workflow is safe: an existing tag, release or npm version is reused.
- Versions are immutable: pushing new commits to a branch whose tag already exists fails. Cut a new version branch.

npm auth is [Trusted Publishing](https://docs.npmjs.com/trusted-publishers) (OIDC): the workflow uses no token secret. Each package trusts this repo and workflow, set up once with 2FA:

```sh
for p in office-engine diagram mermaid plantuml graphviz markdown; do
  npm trust github "@p10node/$p" --repo p10node/office-engine --file release.yml --allow-publish
done
```

`npm trust` needs npm >= 11.15.0 and only works on a package that already exists on npm, so a **new package** is bootstrapped once, either way:

- **Placeholder** (as `@p10node/office-engine` was): publish a `0.0.0` by hand with 2FA, run `npm trust` for it, then release as usual

  ```sh
  bun run build:packages
  cd packages/mermaid && npm version 0.0.0 --no-git-tag-version && npm publish --access public && git checkout package.json
  ```

- **Token for one release**: add a granular npm token (read and write on `@p10node`, *Bypass 2FA*) as the repository secret `NPM_TOKEN`, release (packages not on npm yet are published with it, the others through trusted publishing), run `npm trust` for the new packages, then delete the secret and the token
