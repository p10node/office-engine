import { cssColor } from '../../core/color'
import type { Block, Inline, Paragraph, RunStyle, Table } from '../../core/docx/types'
import type { SlideModel, TextBody } from '../../core/pptx/types'
import { safeHref } from '../../core/url'

// ---------------------------------------------------------------- DOCX

/** text as it reads with its tracked changes accepted: deleted text is left out */
const plain = (runs: Inline[]) => runs.map((r) => (r.t === 'text' || r.t === 'ref' ? (r.s.vanish || r.tc?.k === 'del' ? '' : r.text) : '')).join('')

const cellText = (blocks: Block[]) =>
  blocks
    .filter((b): b is Paragraph => b.t === 'p')
    .map((p) => plain(p.runs))
    .join(' ')
    .replace(/\s+/g, ' ')
    .trim()

/** Plain text: one line per paragraph, list markers kept, table cells separated by tabs. */
export function docxToText(blocks: Block[]): string {
  const out: string[] = []
  for (const b of blocks) {
    if (b.t === 'p') out.push((b.marker ? b.marker.text + (b.marker.tab ? '\t' : ' ') : '') + plain(b.runs))
    else if (b.t === 'tbl') for (const row of b.rows) out.push(row.cells.map((c) => cellText(c.blocks)).join('\t'))
  }
  return out.join('\n') + '\n'
}

const mdEsc = (s: string) => s.replace(/([\\`*_[\]<>])/g, '\\$1')

/** a code font: the run is code */
const MONO = /courier|mono|consolas|menlo|fira code|source code/i

/** pictures of a document as data URLs, by part path (see pictureData) */
export type Pictures = Map<string, string>

/** `heading`: bold is the heading's own look, not emphasis; `images`: pictures written as ![alt](data:…) */
function mdInline(runs: Inline[], heading = false, images?: Pictures): string {
  let out = ''
  let i = 0
  while (i < runs.length) {
    const r = runs[i]
    if (r.t === 'img' && r.tc?.k !== 'del' && images?.has(r.target)) {
      i++
      out += `![${mdEsc(r.alt ?? '')}](${images.get(r.target)})`
      continue
    }
    if ((r.t !== 'text' && r.t !== 'ref') || r.tc?.k === 'del') {
      i++
      continue
    }
    // a script link is exported as its text
    if (r.t === 'text' && safeHref(r.href)) {
      // a link spans the consecutive runs with the same target
      let text = ''
      const href = r.href!
      while (i < runs.length && runs[i].t === 'text' && (runs[i] as { href?: string }).href === href) {
        const x = runs[i++] as Inline & { t: 'text' }
        if (x.tc?.k !== 'del') text += x.text
      }
      out += `[${mdEsc(text)}](${href.replace(/[()\s]/g, encodeURIComponent)})`
      continue
    }
    i++
    if (r.s.vanish) continue
    const text = r.text.replace(/\t/g, ' ').replace(/\n/g, '  \n')
    const core = text.trim()
    if (!core) {
      out += text
      continue
    }
    const mark = (r.s.b && !heading ? '**' : '') + (r.s.i && !heading ? '*' : '')
    const strike = r.s.strike ? '~~' : ''
    const lead = text.slice(0, text.indexOf(core[0]))
    const trail = text.slice(text.lastIndexOf(core[core.length - 1]) + 1)
    const body = r.s.font && MONO.test(r.s.font) ? `\`${core.replace(/`/g, '``')}\`` : mdEsc(core)
    out += `${lead}${strike}${mark}${body}${[...mark].reverse().join('')}${strike}${trail}`
  }
  return out.replace(/\*\*\*\*/g, '')
}

function mdTable(t: Table): string {
  const rows = t.rows.map((r) => r.cells.map((c) => cellText(c.blocks).replace(/\|/g, '\\|') || ' '))
  if (!rows.length) return ''
  const n = Math.max(...rows.map((r) => r.length))
  const line = (r: string[]) => `| ${Array.from({ length: n }, (_, k) => r[k] ?? ' ').join(' | ')} |`
  return [line(rows[0]), `|${' --- |'.repeat(n)}`, ...rows.slice(1).map(line)].join('\n')
}

/** Markdown: headings (outline levels), lists, bold / italic / strike, links, tables, pictures (as data URLs when given). */
export function docxToMarkdown(blocks: Block[], images?: Pictures): string {
  const out: string[] = []
  let prevList = false
  for (const b of blocks) {
    if (b.t === 'tbl') {
      out.push(mdTable(b), '')
      prevList = false
      continue
    }
    if (b.t !== 'p') continue
    const lvl = b.ps.outline
    const heading = lvl !== undefined && lvl < 6
    const text = mdInline(b.runs, heading, images).trim()
    if (b.marker) {
      const ordered = /^[\da-zA-Z]{1,4}[.)]$/.test(b.marker.text.trim())
      const depth = Math.max(0, Math.round(((b.ps.indL ?? 36) - 36) / 36))
      if (!prevList && out.length && out[out.length - 1] !== '') out.push('')
      out.push(`${'   '.repeat(depth)}${ordered ? '1.' : '-'} ${text}`)
      prevList = true
      continue
    }
    if (prevList) out.push('')
    prevList = false
    if (!text) continue
    out.push(heading ? `${'#'.repeat(lvl + 1)} ${text}` : text, '')
  }
  return out.join('\n').replace(/\n{3,}/g, '\n\n').trim() + '\n'
}

// ---------------------------------------------------------------- DOCX → HTML

const hEsc = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')

/**
 * one run's text with its formatting as elements; a heading's bold and color and a link's color and underline are
 * the look of the heading / link itself, not formatting of the text
 */
function styled(text: string, s: RunStyle, heading: boolean, link = false): string {
  let h = hEsc(text).replace(/\n/g, '<br>')
  if (!h.trim()) return h
  if (s.font && MONO.test(s.font)) h = `<code>${h}</code>`
  if (s.b && !heading) h = `<strong>${h}</strong>`
  if (s.i) h = `<em>${h}</em>`
  if (s.u && s.u !== 'none' && !link) h = `<u>${h}</u>`
  if (s.strike || s.dstrike) h = `<s>${h}</s>`
  if (s.va === 'sup') h = `<sup>${h}</sup>`
  else if (s.va === 'sub') h = `<sub>${h}</sub>`
  // into a style attribute of the page written out: colors only (the text of anything else would be markup there)
  const color = cssColor(s.color)
  const hl = cssColor(s.hl)
  const css = [color && color !== 'auto' && !heading && !link ? `color:${color}` : '', hl && hl !== 'none' ? `background:${hl}` : ''].filter(Boolean).join(';')
  return css ? `<span style="${css}">${h}</span>` : h
}

/** consecutive text runs alike (same look, link and tracking) as one, so a word does not come out as three <strong>s */
function merged(runs: Inline[]): Inline[] {
  const out: Inline[] = []
  for (const r of runs) {
    const last = out[out.length - 1]
    if (r.t === 'text' && last?.t === 'text' && last.href === r.href && last.tc?.k === r.tc?.k && JSON.stringify(last.s) === JSON.stringify(r.s)) out[out.length - 1] = { ...last, text: last.text + r.text }
    else out.push(r)
  }
  return out
}

function htmlInline(runs: Inline[], heading = false, images?: Pictures): string {
  let out = ''
  let i = 0
  runs = merged(runs)
  while (i < runs.length) {
    const r = runs[i++]
    if (r.tc?.k === 'del') continue
    if (r.t === 'img') {
      const url = images?.get(r.target)
      if (!url) continue
      const img = `<img src="${url}" alt="${hEsc(r.alt ?? '')}" width="${Math.round((r.w * 96) / 72)}" height="${Math.round((r.h * 96) / 72)}">`
      out += safeHref(r.href) ? `<a href="${hEsc(r.href!)}">${img}</a>` : img
      continue
    }
    if (r.t !== 'text' && r.t !== 'ref') continue
    if (r.s.vanish) continue
    if (r.t === 'text' && safeHref(r.href)) {
      // a link spans the consecutive runs with the same target
      const href = r.href!
      let inner = styled(r.text, r.s, heading, true)
      while (i < runs.length && runs[i].t === 'text' && (runs[i] as { href?: string }).href === href) {
        const x = runs[i++] as Inline & { t: 'text' }
        if (x.tc?.k !== 'del' && !x.s.vanish) inner += styled(x.text, x.s, heading, true)
      }
      out += `<a href="${hEsc(href)}">${inner}</a>`
      continue
    }
    out += styled(r.text, r.s, heading)
  }
  return out
}

const cellHtml = (blocks: Block[], images?: Pictures) =>
  blocks
    .filter((b): b is Paragraph => b.t === 'p')
    .map((p) => htmlInline(p.runs, false, images).trim())
    .filter(Boolean)
    .join('<br>')

const HTML_CSS =
  'body{max-width:48em;margin:2em auto;padding:0 1em;font:16px/1.5 system-ui,-apple-system,"Segoe UI",Roboto,sans-serif;color:#222}' +
  'table{border-collapse:collapse}td{border:1px solid #ccc;padding:4px 8px;vertical-align:top}img{max-width:100%;height:auto}' +
  'code{font-family:ui-monospace,Menlo,Consolas,monospace;background:#f3f3f3;padding:0 .2em}'

/**
 * A web page of the document: headings (outline levels), paragraphs, nested lists, bold / italic / underline / strike,
 * links, code, colors, tables, pictures (as data URLs when given). The text reads with its tracked changes accepted.
 */
export function docxToHtml(blocks: Block[], opts: { title?: string; images?: Pictures } = {}): string {
  const { images } = opts
  const out: string[] = []
  // open lists, an <li> open in each; closing tags go on the line of the last item
  const lists: ('ul' | 'ol')[] = []
  const closeTo = (n: number) => {
    while (lists.length > n) out[out.length - 1] += `</li></${lists.pop()}>`
  }
  for (const b of blocks) {
    if (b.t === 'tbl') {
      closeTo(0)
      out.push(`<table>\n${b.rows.map((r) => `<tr>${r.cells.map((c) => `<td>${cellHtml(c.blocks, images)}</td>`).join('')}</tr>`).join('\n')}\n</table>`)
      continue
    }
    if (b.t !== 'p') continue
    const lvl = b.ps.outline
    const heading = lvl !== undefined && lvl < 6
    const inner = htmlInline(b.runs, heading, images).trim()
    if (b.marker) {
      const tag = /^[\da-zA-Z]{1,4}[.)]$/.test(b.marker.text.trim()) ? 'ol' : 'ul'
      const depth = Math.max(0, Math.round(((b.ps.indL ?? 36) - 36) / 36)) + 1
      // a shallower item, or another kind of list at this depth: the deeper / other list ends
      closeTo(lists[depth - 1] === tag ? depth : depth - 1)
      if (depth > lists.length) {
        // deeper: a list inside the open item
        while (lists.length < depth) {
          lists.push(tag)
          out.push(`<${tag}>`)
        }
      } else out[out.length - 1] += '</li>'
      out.push(`<li>${inner}`)
      continue
    }
    closeTo(0)
    if (!inner) continue
    const align = b.ps.align === 'center' || b.ps.align === 'right' ? ` style="text-align:${b.ps.align}"` : ''
    out.push(heading ? `<h${lvl + 1}${align}>${inner}</h${lvl + 1}>` : `<p${align}>${inner}</p>`)
  }
  closeTo(0)
  return (
    `<!doctype html>\n<html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">` +
    `<title>${hEsc(opts.title ?? '')}</title>\n<style>${HTML_CSS}</style></head>\n<body>\n${out.join('\n')}\n</body></html>\n`
  )
}

/** The pictures of a document (paragraphs and table cells) as data URLs, by part path; `part` reads a package part. */
export async function pictureData(blocks: Block[], part: (target: string) => Promise<Blob>): Promise<Pictures> {
  const targets = new Set<string>()
  const walk = (bs: Block[]) => {
    for (const b of bs) {
      if (b.t === 'p') for (const r of b.runs) if (r.t === 'img') targets.add(r.target)
      if (b.t === 'tbl') for (const row of b.rows) for (const c of row.cells) walk(c.blocks)
    }
  }
  walk(blocks)
  const out: Pictures = new Map()
  for (const t of targets) {
    try {
      const blob = await part(t)
      const bytes = new Uint8Array(await blob.arrayBuffer())
      let s = ''
      for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode(...bytes.subarray(i, i + 0x8000))
      out.set(t, `data:${blob.type || 'application/octet-stream'};base64,${btoa(s)}`)
    } catch {
      // a missing part: the picture is left out
    }
  }
  return out
}

// ---------------------------------------------------------------- PPTX

const bodyText = (t: TextBody | undefined) =>
  t
    ? t.paras
        .map((p) => (p.bullet ? p.bullet.text + ' ' : '') + p.runs.map((r) => (r.br ? '\n' : r.text)).join(''))
        .filter((l) => l.trim())
        .join('\n')
    : ''

/** Text of each slide (shapes, tables, speaker notes). */
export function slidesToText(slides: SlideModel[]): string {
  const out: string[] = []
  for (const s of slides) {
    out.push(`Slide ${s.index + 1}`, '')
    for (const el of s.els) {
      if (el.t === 'shape') {
        const t = bodyText(el.text)
        if (t) out.push(t, '')
      } else if (el.t === 'table') {
        for (const row of el.rows) out.push(row.cells.filter((c) => !c.hidden).map((c) => bodyText(c.text).replace(/\s+/g, ' ')).join('\t'))
        out.push('')
      }
    }
    if (s.notes?.trim()) out.push('Notes:', s.notes.trim(), '')
    out.push('')
  }
  return out.join('\n').replace(/\n{3,}/g, '\n\n').trim() + '\n'
}

// ---------------------------------------------------------------- CSV

/** RFC 4180 output (CRLF; fields with separators, quotes or line breaks are quoted). */
export { toCsv } from '../../core/convert'
