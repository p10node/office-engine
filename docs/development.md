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

To release, push a branch named after the version:

```sh
git switch -c v0.2.0 main
git push -u origin v0.2.0
```

`.github/workflows/release.yml` then runs CI, sets `package.json` `version` to match the branch (committing `chore(release): v0.2.0` onto the branch if it differs), pushes the `v0.2.0` tag, creates the GitHub release with the npm tarball attached, and publishes to npm with provenance.

- Release notes: optional `docs/releases/<tag>.md`, placed above GitHub's generated commit list.
- Prereleases (`v0.2.0-rc.1`) are marked prerelease on GitHub and go to the `next` dist-tag; others go to `latest`.
- Re-running the workflow is safe: an existing tag, release or npm version is reused.
- Versions are immutable: pushing new commits to a branch whose tag already exists fails. Cut a new version branch.

npm auth is [Trusted Publishing](https://docs.npmjs.com/trusted-publishers) (OIDC): the workflow uses no token secret. The package trusts this repo and workflow, set up once with 2FA:

```sh
npm trust github @p10node/office-engine --repo p10node/office-engine --file release.yml --allow-publish
```

`npm trust` needs npm >= 11.15.0 and only works on a package that already exists on npm, so the package was bootstrapped with a placeholder `0.0.0` published by hand.
