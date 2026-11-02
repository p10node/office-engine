import type { ReactNode } from 'react'
import { createRoot } from 'react-dom/client'
import { stylesText } from '../injectStyles'

/** Content laid out for paper: PDF / print and standalone HTML. */
export interface Printable {
  /** the pages (`.de-print-page` elements) or flowing content */
  node: ReactNode
  /** CSS `@page` size: '612pt 792pt', 'landscape', 'auto' */
  size: string
  /** CSS `@page` margin */
  margin: string
  /** shown to the user (e.g. output was truncated) */
  notice?: string
}

const PRINT_CSS = `
html, body { margin: 0; padding: 0; background: #fff; }
.de-print {
  color: #000;
  font: 13px/1.4 system-ui, -apple-system, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif;
  -webkit-print-color-adjust: exact;
  print-color-adjust: exact;
}
.de-print *, .de-print *::before, .de-print *::after { box-sizing: border-box; }
.de-print img { max-width: none; }
.de-print-page { position: relative; overflow: hidden; background: #fff; break-after: page; }
.de-print-page:last-child { break-after: auto; }
.de-print-page > .de-page { left: 0; top: 0; box-shadow: none; }
.de-print-table { border-collapse: collapse; table-layout: fixed; }
.de-print-table td { padding: 0 3px 1px; overflow: hidden; }
.de-print-table.de-print-grid td { border: 1px solid #e2e4e7; }
@media screen {
  body { background: #e9ebee; }
  .de-print { padding: 24px 16px; }
  .de-print-page { margin: 0 auto 24px; box-shadow: 0 1px 3px rgba(15, 23, 42, 0.18); }
  .de-print-sheet { width: max-content; max-width: none; margin: 0 auto; padding: 16px; background: #fff; box-shadow: 0 1px 3px rgba(15, 23, 42, 0.18); }
}
@media print {
  .de-print-fit { zoom: var(--de-fit, 1); }
}
`

const cssFor = (p: Printable) => `${stylesText}\n${PRINT_CSS}\n@page { size: ${p.size}; margin: ${p.margin}; }\n`

const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms))

/** waits until React and async media (slide pictures, fills) stop changing the DOM, then for images and fonts */
async function settle(doc: Document, host: HTMLElement) {
  await sleep(0)
  await new Promise<void>((resolve) => {
    let quiet: ReturnType<typeof setTimeout>
    const end = setTimeout(done, 8000)
    const mo = new MutationObserver(() => {
      clearTimeout(quiet)
      quiet = setTimeout(done, 250)
    })
    function done() {
      mo.disconnect()
      clearTimeout(quiet)
      clearTimeout(end)
      resolve()
    }
    mo.observe(host, { subtree: true, childList: true, attributes: true })
    quiet = setTimeout(done, 250)
  })
  const urls = new Set<string>()
  for (const el of host.querySelectorAll<HTMLElement>('[style*="url("]')) {
    for (const m of el.style.backgroundImage.matchAll(/url\(["']?([^"')]+)["']?\)/g)) urls.add(m[1])
  }
  const loads: Promise<unknown>[] = [...host.querySelectorAll('img')].map((img) =>
    img.complete ? null : new Promise((r) => ((img.onload = r), (img.onerror = r))),
  ) as Promise<unknown>[]
  for (const u of urls) {
    loads.push(
      new Promise((r) => {
        const img = new Image()
        img.onload = img.onerror = r
        img.src = u
      }),
    )
  }
  await Promise.race([Promise.all(loads), sleep(10_000)])
  await Promise.race([doc.fonts?.ready, sleep(3000)])
}

interface Mounted {
  frame: HTMLIFrameElement
  doc: Document
  host: HTMLElement
  dispose(): void
}

/** renders the printable into an off-screen same-origin iframe (keeps host page CSS out) */
async function mount(p: Printable, title: string): Promise<Mounted> {
  const frame = document.createElement('iframe')
  frame.setAttribute('aria-hidden', 'true')
  frame.tabIndex = -1
  frame.style.cssText = 'position:fixed;left:-20000px;top:0;width:1100px;height:800px;border:0;visibility:hidden'
  document.body.appendChild(frame)
  const doc = frame.contentDocument!
  doc.open()
  doc.write('<!doctype html><html><head><meta charset="utf-8"></head><body></body></html>')
  doc.close()
  doc.title = title
  const style = doc.createElement('style')
  style.textContent = cssFor(p)
  doc.head.appendChild(style)
  const host = doc.createElement('div')
  host.className = 'de-print'
  doc.body.appendChild(host)
  const root = createRoot(host)
  root.render(p.node)
  let disposed = false
  const dispose = () => {
    if (disposed) return
    disposed = true
    root.unmount()
    frame.remove()
  }
  try {
    await settle(doc, host)
  } catch (e) {
    dispose()
    throw e
  }
  return { frame, doc, host, dispose }
}

let printing: Mounted | null = null

/** Opens the browser print dialog for the content (choose "Save as PDF" there). */
export async function printDocument(p: Printable, title: string): Promise<void> {
  printing?.dispose()
  const m = await mount(p, title)
  printing = m
  const w = m.frame.contentWindow!
  const cleanup = () => {
    if (printing === m) printing = null
    // give the print pipeline a moment before the frame goes away
    setTimeout(() => m.dispose(), 1000)
  }
  w.addEventListener('afterprint', cleanup, { once: true })
  setTimeout(cleanup, 10 * 60_000)
  // the PDF file name proposed by "Save as PDF" comes from the top document's title in some browsers
  const prev = document.title
  document.title = title
  try {
    w.focus()
    // Safari prints the whole page with frame.print(); execCommand prints the frame there
    let ok = false
    try {
      ok = w.document.execCommand('print', false)
    } catch {
      ok = false
    }
    if (!ok) w.print()
  } finally {
    document.title = prev
  }
}

const esc = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')

function toDataUrl(blob: Blob): Promise<string> {
  return new Promise((resolve, reject) => {
    const r = new FileReader()
    r.onload = () => resolve(r.result as string)
    r.onerror = () => reject(r.error)
    r.readAsDataURL(blob)
  })
}

/** Standalone web page of the content (pictures embedded). */
export async function exportHtml(p: Printable, title: string): Promise<Blob> {
  const m = await mount(p, title)
  try {
    let body = m.host.outerHTML.replace(/ contenteditable="(?:true|false)"/g, '')
    const blobs = new Set(body.match(/blob:[^"'()\s&]+/g) ?? [])
    for (const url of blobs) {
      try {
        const data = await toDataUrl(await (await fetch(url)).blob())
        body = body.split(url).join(data)
      } catch {
        // revoked: the picture is left out
      }
    }
    const html =
      `<!doctype html>\n<html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">` +
      `<title>${esc(title)}</title><style>${cssFor(p)}</style></head><body>${body}</body></html>\n`
    return new Blob([html], { type: 'text/html;charset=utf-8' })
  } finally {
    m.dispose()
  }
}
