# Development

```bash
bun install
bun run dev          # playground http://localhost:5173 (drag & drop, URL, samples)
bun run collab       # collaboration hub for the playground's "Share live" (port 8787)
bun test             # core tests on fixtures
bun run test:e2e     # browser tests (headless Chromium: bunx playwright install chromium; needs large-deck.pptx, see fixtures.py)
bun run build        # dist/: ESM lib + office-engine.embed.js + .d.ts
bun run build:packages  # packages/*/dist: the standalone SDK packages (@p10node/diagram, mermaid, plantuml, graphviz, markdown)
python3 -m venv .venv && .venv/bin/pip install python-docx openpyxl python-pptx Pillow
.venv/bin/python scripts/fixtures.py --large   # generate samples + heavy files (--large-deck: only the deck the browser tests need)
bun scripts/iwork-samples.ts                   # sample.pages / .key / .numbers (built by test/iwork-builder.ts)
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
src/react    view components
src/embed.ts mount() + <office-viewer>
src/collab   real-time collaboration (hub, session, per-format adapters)
```

## Browser tests

`e2e/*.spec.ts` drive a fixture app (`e2e/app`, served by Vite with `vite.e2e.config.ts`) in headless Chromium through `playwright-core`, one browser and server per file (`e2e/harness.ts`).
The app reads `DocViewer` props from the query string (`?src=samples/sample.docx&editable=1&allowDownload=0&onSave=1…`) and records callbacks in `window.__events`; saved files are re-opened with the engine to check what was written.

| Spec                  | Covers                                                                                            |
|-----------------------|---------------------------------------------------------------------------------------------------|
| `viewer.spec.ts`      | rendering of the 3 formats, outline, zoom, sheet tabs, thumbnails, slideshow, props, errors       |
| `legacy.spec.ts`      | .doc / .xls / .ppt / .rtf / OpenDocument samples: converted, shown, edited, saved as OOXML        |
| `docx.spec.ts`        | typing, formatting, styles, lists, colors, fonts, links, tables, images, find & replace, undo     |
| `xlsx.spec.ts`        | values, formulas, merges (all / across / down / unmerge), formats, rows / columns, sheets, freeze |
| `pptx.spec.ts`        | text, shapes (insert, move, fill, arrange, rotate, duplicate, delete), images, slides, present    |
| `file-menu.spec.ts`   | every File entry: New, Open (conversions), Import, Save, Rename, Download as, Details, Print      |
| `permissions.spec.ts` | `allowSave` / `allowDownload` combinations, image download, Ctrl+S / Ctrl+P                       |
| `color.spec.ts`       | custom colors (system picker, recent list, persistence) in the 3 formats                          |
| `embed.spec.ts`       | `<office-viewer>` (shadow DOM, attributes, `file`) and `mount()`                                  |
| `collab.spec.ts`      | two editors on `scripts/collab-server.ts`: live edits, presence, File menu in a room              |

`bun test` only runs `test/` (`bunfig.toml`); `bun run test:e2e` runs `e2e/`.

## Site (GitHub Pages)

`.github/workflows/pages.yml` deploys on every push to `main`:

- [p10node.github.io/office-engine](https://p10node.github.io/office-engine/): docs rendered from `README.md` + `docs/*.md`
- [p10node.github.io/office-engine/demo/](https://p10node.github.io/office-engine/demo/): the playground (heavy samples generated in CI)

One-time setup: **Settings → Pages → Source: GitHub Actions**. Build locally: `bun run build:playground && bun run build:site` → `dist-site/`.

"Share live" needs a collaboration hub; the static demo has none, so pass one: `?collab=https://your-hub.example`.

## Release

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
