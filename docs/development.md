# Development

```bash
bun install
bun run dev          # playground http://localhost:5173 (drag & drop, URL, samples)
bun run collab       # collaboration hub for the playground's "Share live" (port 8787)
bun test             # core tests on fixtures
bun run build        # dist/: ESM lib + office-engine.embed.js + .d.ts
python3 -m venv .venv && .venv/bin/pip install python-docx openpyxl python-pptx Pillow
.venv/bin/python scripts/fixtures.py --large   # generate samples + heavy files
bun run screenshots  # README screenshots -> docs/screenshots/ (headless Chromium)
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
src/react    view components
src/embed.ts mount() + <office-viewer>
src/collab   real-time collaboration (hub, session, per-format adapters)
```

## Site (GitHub Pages)

`.github/workflows/pages.yml` deploys on every push to `main`:

- [p10node.github.io/office-engine](https://p10node.github.io/office-engine/): docs rendered from `README.md` + `docs/*.md`
- [p10node.github.io/office-engine/demo/](https://p10node.github.io/office-engine/demo/): the playground (heavy samples generated in CI)

One-time setup: **Settings → Pages → Source: GitHub Actions**. Build locally: `bun run build:playground && bun run build:site` → `dist-site/`.

"Share live" needs a collaboration hub; the static demo has none, so pass one: `?collab=https://your-hub.example`.

## Release

CI (`.github/workflows/ci.yml`) runs typecheck, tests, build and `npm pack --dry-run` on every push to `main` and every PR.

Publishing (`.github/workflows/release.yml`) is triggered by either:

- **Actions → Release → Run workflow**, picking `patch` / `minor` / `major` / `pre*`: bumps `package.json`, publishes to npm with provenance, pushes the `chore(release)` commit + `vX.Y.Z` tag, creates a GitHub release.
- **Pushing a tag** from local: `npm version patch && git push --follow-tags`.

Prerelease versions (`1.2.0-rc.0`) go to the `next` dist-tag, others to `latest`. npm auth: configure [Trusted Publishing](https://docs.npmjs.com/trusted-publishers) for this repo + `release.yml`, or add an `NPM_TOKEN` repository secret.
