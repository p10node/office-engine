// GitHub Pages site -> dist-site/
//   index.html   README.md
//   docs/*.html  docs/*.md (+ docs/screenshots)
//   demo/        playground (run `bun run build:playground` first)
// Usage: bun run build:site
import { cp, exists, mkdir, rm } from 'node:fs/promises'
import pkg from '../package.json'

const OUT = 'dist-site'
const HOME = pkg.homepage // absolute links to the site become relative
const REPO = pkg.repository.url.replace(/^git\+/, '').replace(/\.git$/, '')

const PAGES = [
  { src: 'README.md', out: 'index.html', label: 'Overview' },
  { src: 'docs/features.md', out: 'docs/features.html', label: 'Features' },
  { src: 'docs/api.md', out: 'docs/api.html', label: 'API' },
  { src: 'docs/collaboration.md', out: 'docs/collaboration.html', label: 'Collaboration' },
  { src: 'docs/development.md', out: 'docs/development.html', label: 'Development' },
]

const text = (html: string) => html.replace(/<[^>]+>/g, '')

/** Site mirrors the repo layout, so only .md targets and links to the site itself need rewriting. */
function rewrite(href: string, root: string): string {
  if (href.startsWith(HOME)) return root + href.slice(HOME.length) || './'
  if (/^[a-z]+:|^#|^\//i.test(href)) return href
  return href.replace(/(^|\/)README\.md(?=#|$)/, '$1index.html').replace(/\.md(?=#|$)/, '.html')
}

function page(p: (typeof PAGES)[number], body: string): string {
  const root = p.out.includes('/') ? '../' : ''
  const sections = [...body.matchAll(/<h2 id="([^"]+)">(.*?)<\/h2>/g)].map(([, id, h]) => `<li><a href="#${id}">${text(h)}</a></li>`).join('')
  const nav = PAGES.map((q) => {
    const cur = q === p
    return `<a href="${root}${q.out}"${cur ? ' aria-current="page"' : ''}>${q.label}</a>${cur && sections ? `<ul>${sections}</ul>` : ''}`
  }).join('')
  const title = p.out === 'index.html' ? 'office-engine: view & edit DOCX, XLSX, PPTX in the browser' : `${p.label} · office-engine`
  return `<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>${title}</title>
<meta name="description" content="${pkg.description}">
<style>${CSS}</style>
</head>
<body>
<header class="top">
  <a class="brand" href="${root || './'}">office-engine</a>
  <nav><a class="demo" href="${root}demo/">Live demo</a><a href="${REPO}">GitHub</a></nav>
</header>
<div class="layout">
  <aside><nav>${nav}</nav></aside>
  <main>${body}</main>
</div>
</body>
</html>
`
}

const CSS = `
:root{--bg:#fff;--fg:#1d2127;--muted:#5f6670;--line:#e3e6ea;--soft:#f5f6f8;--accent:#1a73e8;color-scheme:light}
@media (prefers-color-scheme:dark){:root{--bg:#15171a;--fg:#e6e8eb;--muted:#9aa1ab;--line:#2c3036;--soft:#1e2125;--accent:#7aacff;color-scheme:dark}}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--fg);font:15px/1.65 system-ui,-apple-system,'Segoe UI',Roboto,sans-serif}
a{color:var(--accent);text-decoration:none}a:hover{text-decoration:underline}
.top{position:sticky;top:0;z-index:1;display:flex;align-items:center;gap:16px;padding:12px 24px;background:var(--bg);border-bottom:1px solid var(--line)}
.brand{font-weight:700;color:var(--fg);font-size:16px}
.top nav{margin-left:auto;display:flex;gap:18px;align-items:center}
.top .demo{border:1px solid var(--accent);border-radius:6px;padding:4px 12px}
.layout{display:grid;grid-template-columns:220px minmax(0,1fr);max-width:1180px;margin:0 auto;padding:0 24px}
aside nav{position:sticky;top:57px;padding:28px 16px 28px 0;display:flex;flex-direction:column;gap:2px;font-size:14px}
aside a{color:var(--muted);padding:4px 10px;border-radius:6px}
aside a[aria-current]{color:var(--fg);background:var(--soft);font-weight:600}
aside ul{list-style:none;margin:2px 0 8px;padding:0 0 0 10px;border-left:1px solid var(--line);margin-left:10px}
aside ul a{display:block;font-size:13px;padding:2px 10px}
main{min-width:0;padding:28px 0 80px 32px;max-width:900px}
h1{font-size:32px;line-height:1.2;margin:0 0 16px}
h2{font-size:22px;margin:40px 0 12px;padding-top:8px;border-top:1px solid var(--line)}
h3{font-size:17px;margin:28px 0 8px}
h2,h3{scroll-margin-top:70px}
code{font:13px/1.5 ui-monospace,SFMono-Regular,Menlo,Consolas,monospace;background:var(--soft);padding:1px 5px;border-radius:4px}
pre{background:var(--soft);border:1px solid var(--line);border-radius:8px;padding:14px 16px;overflow-x:auto}
pre code{background:none;padding:0}
.table{overflow-x:auto;margin:16px 0}
table{border-collapse:collapse;font-size:14px;width:100%}
th,td{border:1px solid var(--line);padding:6px 10px;text-align:left;vertical-align:top}
th{background:var(--soft)}
img{max-width:100%;border:1px solid var(--line);border-radius:6px;display:block}
@media (max-width:760px){.layout{grid-template-columns:1fr;padding:0 16px}aside nav{position:static;flex-direction:row;flex-wrap:wrap;padding:12px 0 0}aside ul{display:none}main{padding:16px 0 60px}.top{padding:10px 16px}}
`

await rm(OUT, { recursive: true, force: true })
await mkdir(`${OUT}/docs`, { recursive: true })
for (const p of PAGES) {
  const root = p.out.includes('/') ? '../' : ''
  const body = Bun.markdown
    .html(await Bun.file(p.src).text(), { headings: { ids: true } })
    .replace(/\b(href|src)="([^"]*)"/g, (_, attr, url) => `${attr}="${rewrite(url, root)}"`)
    .replace(/<img src="([^"]+)"[^>]*>/g, '<a href="$1">$&</a>') // full-size on click, like GitHub
    .replace(/<table>/g, '<div class="table"><table>')
    .replace(/<\/table>/g, '</table></div>')
  await Bun.write(`${OUT}/${p.out}`, page(p, body))
}
await cp('docs/screenshots', `${OUT}/docs/screenshots`, { recursive: true })
if (await exists('dist-playground')) await cp('dist-playground', `${OUT}/demo`, { recursive: true })
else console.warn('dist-playground missing -> no demo/ (run `bun run build:playground`)')
console.log(`${OUT}/: ${PAGES.map((p) => p.out).join(', ')}`)
