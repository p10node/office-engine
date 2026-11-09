/**
 * Markdown → HTML for showing in a page: CommonMark blocks (headings, paragraphs, nested lists, quotes, code,
 * rules) and GitHub's tables, task lists, strikethrough and bare links. The HTML is safe to set as innerHTML:
 * text is escaped, raw HTML in the source is shown as text, and links / images keep only harmless URLs.
 * ```mermaid / ```plantuml / ```dot code blocks become SVG drawings.
 */
import { DIAGRAM_NAME, FENCE_LANG, renderDiagrams } from '../../../src/core/diagram'
import { safeHref } from '../../../src/core/url'

export interface MarkdownOptions {
  /** Draw ```mermaid, ```plantuml / ```puml and ```dot / ```graphviz code blocks as SVG. Default true. */
  diagrams?: boolean
  /** `id` attributes on headings (from their text) so `#links` reach them. Default true. */
  headingIds?: boolean
}

interface Ctx {
  refs: Map<string, { href: string; title?: string }>
  diagrams: boolean
  ids: boolean
  slugs: Map<string, number>
}

const esc = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')

const FENCE = /^( {0,3})(`{3,}|~{3,})[ \t]*([^\s`]*)[^`\n]*$/
const ATX = /^ {0,3}(#{1,6})(?:[ \t]+(.*?))?(?:[ \t]+#+)?[ \t]*$/
const SETEXT = /^ {0,3}(=+|-+)[ \t]*$/
const HR = /^ {0,3}([-*_])(?:[ \t]*\1){2,}[ \t]*$/
const QUOTE = /^ {0,3}> ?(.*)$/
const LIST = /^( {0,3})([-*+]|\d{1,9}[.)])(?:([ \t]{1,4})(.*)|[ \t]*)$/
const DELIM = /^ {0,3}\|?[ \t]*:?-+:?[ \t]*(?:\|[ \t]*:?-+:?[ \t]*)*\|?[ \t]*$/
const REF = /^ {0,3}\[([^\]]+)\]:[ \t]*<?([^\s>]+)>?(?:[ \t]+(?:"([^"]*)"|'([^']*)'|\(([^)]*)\)))?[ \t]*$/

const label = (s: string) => s.trim().replace(/\s+/g, ' ').toLowerCase()

/** Markdown text → HTML. */
export function markdownToHtml(md: string, opts: MarkdownOptions = {}): string {
  const lines = md.replace(/^﻿/, '').replace(/\r\n?/g, '\n').replace(/\t/g, '    ').split('\n')
  const ctx: Ctx = { refs: new Map(), diagrams: opts.diagrams !== false, ids: opts.headingIds !== false, slugs: new Map() }
  // link reference definitions (outside code blocks) are read first and leave no text
  const kept: string[] = []
  let fence: string | null = null
  for (const l of lines) {
    const f = FENCE.exec(l)
    if (fence) {
      if (f && f[2][0] === fence[0] && f[2].length >= fence.length && !f[3]) fence = null
      kept.push(l)
      continue
    }
    if (f) fence = f[2]
    const r = !fence && REF.exec(l)
    if (r) {
      const key = label(r[1])
      if (!ctx.refs.has(key)) ctx.refs.set(key, { href: r[2], title: r[3] ?? r[4] ?? r[5] })
      continue
    }
    kept.push(l)
  }
  return blocks(kept, ctx, false)
}

/** a line that starts a block other than a paragraph (ends a lazy continuation) */
const startsBlock = (l: string) => FENCE.test(l) || ATX.test(l) || HR.test(l) || QUOTE.test(l) || LIST.test(l)

function blocks(lines: string[], ctx: Ctx, tight: boolean): string {
  const out: string[] = []
  let para: string[] = []
  const flush = () => {
    if (!para.length) return
    const html = inline(para.map((l) => l.replace(/^ +/, '')).join('\n').replace(/\s+$/, ''), ctx)
    out.push(tight ? html : `<p>${html}</p>`)
    para = []
  }
  let i = 0
  while (i < lines.length) {
    const line = lines[i]
    if (!line.trim()) {
      flush()
      i++
      continue
    }
    let m: RegExpExecArray | null
    if ((m = FENCE.exec(line))) {
      flush()
      const indent = m[1].length
      const mark = m[2]
      const close = new RegExp(`^ {0,3}${mark[0] === '`' ? '`' : '~'}{${mark.length},}[ \\t]*$`)
      const body: string[] = []
      i++
      while (i < lines.length && !close.test(lines[i])) body.push(lines[i++].replace(new RegExp(`^ {0,${indent}}`), ''))
      i++
      out.push(code(body.join('\n'), m[3], ctx))
      continue
    }
    if (!para.length && /^ {4}/.test(line)) {
      const body: string[] = []
      while (i < lines.length && (/^ {4}/.test(lines[i]) || !lines[i].trim())) body.push(lines[i++].slice(4))
      while (body.length && !body[body.length - 1].trim()) body.pop()
      out.push(`<pre><code>${esc(body.join('\n'))}\n</code></pre>`)
      continue
    }
    if ((m = ATX.exec(line))) {
      flush()
      out.push(heading(m[1].length, m[2] ?? '', ctx))
      i++
      continue
    }
    if (para.length && (m = SETEXT.exec(line))) {
      const text = para.join('\n')
      para = []
      out.push(heading(m[1][0] === '=' ? 1 : 2, text, ctx))
      i++
      continue
    }
    if (HR.test(line)) {
      flush()
      out.push('<hr>')
      i++
      continue
    }
    if (QUOTE.test(line)) {
      flush()
      const inner: string[] = []
      while (i < lines.length) {
        const q = QUOTE.exec(lines[i])
        if (q) inner.push(q[1])
        // lazy continuation of a paragraph in the quote
        else if (lines[i].trim() && inner.length && inner[inner.length - 1].trim() && !startsBlock(lines[i])) inner.push(lines[i])
        else break
        i++
      }
      out.push(`<blockquote>\n${blocks(inner, ctx, false)}\n</blockquote>`)
      continue
    }
    if ((m = LIST.exec(line))) {
      // a list interrupts a paragraph only with content, and (numbered) from 1
      const empty = m[4] === undefined || !m[4].trim()
      if (!para.length || (!empty && (!/\d/.test(m[2]) || parseInt(m[2], 10) === 1))) {
        flush()
        i = list(lines, i, out, ctx)
        continue
      }
    }
    if (line.includes('|') && i + 1 < lines.length && DELIM.test(lines[i + 1]) && (lines[i + 1].includes('|') || !line.trim().startsWith('|'))) {
      flush()
      i = table(lines, i, out, ctx)
      continue
    }
    para.push(line)
    i++
  }
  flush()
  return out.join('\n')
}

function list(lines: string[], start: number, out: string[], ctx: Ctx): number {
  const first = LIST.exec(lines[start])!
  const ordered = /\d/.test(first[2])
  const delim = first[2].slice(-1)
  const items: string[][] = []
  let loose = false
  let i = start
  while (i < lines.length) {
    const m = LIST.exec(lines[i])
    if (!m || /\d/.test(m[2]) !== ordered || m[2].slice(-1) !== delim) break
    // content starts after the marker and 1–4 spaces (more: one space, the rest is indented code)
    const gap = m[3] === undefined ? 1 : m[3].length > 4 ? 1 : m[3].length
    const width = m[1].length + m[2].length + gap
    const item = [m[3] === undefined ? '' : m[3].length > 4 ? m[3].slice(1) + m[4] : m[4]]
    i++
    let blank = false
    while (i < lines.length) {
      const l = lines[i]
      if (!l.trim()) {
        item.push('')
        blank = true
        i++
        continue
      }
      const lead = l.length - l.trimStart().length
      if (lead >= width) {
        item.push(l.slice(width))
        blank = false
        i++
        continue
      }
      if (blank) break
      // lazy continuation of the item's paragraph
      if (!startsBlock(l) && item[item.length - 1].trim()) {
        item.push(l)
        i++
        continue
      }
      break
    }
    let trailing = false
    while (item.length > 1 && !item[item.length - 1].trim()) {
      item.pop()
      trailing = true
    }
    // a blank line between two blocks of the item: a loose list (paragraphs)
    for (let k = 1; k + 1 < item.length; k++) if (!item[k].trim() && item[k + 1].trim() && !/^\s/.test(item[k + 1]) && !LIST.test(item[k + 1])) loose = true
    items.push(item)
    if (trailing) {
      const next = i < lines.length ? LIST.exec(lines[i]) : null
      if (next && /\d/.test(next[2]) === ordered && next[2].slice(-1) === delim) loose = true
    }
  }
  const startNo = ordered ? parseInt(first[2], 10) : 1
  out.push(ordered ? `<ol${startNo !== 1 ? ` start="${startNo}"` : ''}>` : '<ul>')
  for (const item of items) {
    const task = /^\[([ xX])\][ \t]+/.exec(item[0])
    if (task) item[0] = item[0].slice(task[0].length)
    const body = blocks(item, ctx, !loose)
    const box = task ? `<input type="checkbox" disabled${task[1] === ' ' ? '' : ' checked'}> ` : ''
    out.push(task ? `<li class="task-list-item">${box}${body}</li>` : `<li>${body}</li>`)
  }
  out.push(ordered ? '</ol>' : '</ul>')
  return i
}

function table(lines: string[], i: number, out: string[], ctx: Ctx): number {
  const cells = (l: string) => {
    let s = l.trim()
    if (s.startsWith('|')) s = s.slice(1)
    if (s.endsWith('|') && !s.endsWith('\\|')) s = s.slice(0, -1)
    return s.split(/(?<!\\)\|/).map((c) => c.trim().replace(/\\\|/g, '|'))
  }
  const head = cells(lines[i])
  const align = cells(lines[i + 1]).map((c) => (c.startsWith(':') && c.endsWith(':') ? 'center' : c.endsWith(':') ? 'right' : c.startsWith(':') ? 'left' : ''))
  i += 2
  const rows: string[][] = []
  while (i < lines.length && lines[i].trim() && !startsBlock(lines[i])) rows.push(cells(lines[i++]))
  const cell = (tag: string, c: string | undefined, k: number) => `<${tag}${align[k] ? ` style="text-align:${align[k]}"` : ''}>${inline(c ?? '', ctx)}</${tag}>`
  let html = `<table>\n<thead>\n<tr>${head.map((c, k) => cell('th', c, k)).join('')}</tr>\n</thead>`
  if (rows.length) html += `\n<tbody>\n${rows.map((r) => `<tr>${head.map((_, k) => cell('td', r[k], k)).join('')}</tr>`).join('\n')}\n</tbody>`
  out.push(html + '\n</table>')
  return i
}

function heading(level: number, text: string, ctx: Ctx): string {
  const html = inline(text.trim(), ctx)
  if (!ctx.ids) return `<h${level}>${html}</h${level}>`
  const base =
    html
      .replace(/<[^>]*>/g, '')
      .replace(/&\w+;|&#\d+;/g, '')
      .toLowerCase()
      .replace(/[^\p{L}\p{N}\s_-]/gu, '')
      .trim()
      .replace(/\s+/g, '-') || 'section'
  const n = ctx.slugs.get(base) ?? 0
  ctx.slugs.set(base, n + 1)
  return `<h${level} id="${esc(n ? `${base}-${n}` : base)}">${html}</h${level}>`
}

function code(body: string, info: string, ctx: Ctx): string {
  const lang = info.toLowerCase().replace(/[^\w+#.-]/g, '')
  const diagram = FENCE_LANG[lang]
  if (diagram && ctx.diagrams) {
    try {
      return renderDiagrams(diagram, body)
        .map((d) => `<figure class="md-diagram md-${diagram}">${d.svg}</figure>`)
        .join('\n')
    } catch (e) {
      return (
        `<div class="md-diagram-error"><p>${esc(DIAGRAM_NAME[diagram])} diagram not drawn: ${esc((e as Error).message)}</p>` +
        `<pre><code class="language-${lang}">${esc(body)}\n</code></pre></div>`
      )
    }
  }
  return `<pre><code${lang ? ` class="language-${lang}"` : ''}>${esc(body)}${body ? '\n' : ''}</code></pre>`
}

/** URLs a picture may load: web, relative, and data: images */
function safeSrc(src: string): string | undefined {
  const s = src.trim()
  if (/^data:image\/(png|jpe?g|gif|webp|svg\+xml|avif|bmp);/i.test(s)) return s
  const scheme = /^([a-z][a-z0-9+.-]*):/i.exec(s.replace(/[\u0000- ]/g, ''))?.[1].toLowerCase()
  return !scheme || scheme === 'http' || scheme === 'https' ? s : undefined
}

const PUNCT = /\\([!"#$%&'()*+,\-./:;<=>?@[\\\]^_`{|}~])/g
const LINK_TAIL = String.raw`\(\s*(<[^>\n]*>|[^\s()]*(?:\([^\s()]*\)[^\s()]*)*)(?:\s+(?:"([^"]*)"|'([^']*)'|\(([^)]*)\)))?\s*\)`
const IMAGE = new RegExp(String.raw`!\[((?:[^\[\]\\]|\\.|\[[^\]]*\])*)\]` + LINK_TAIL, 'g')
const LINK = new RegExp(String.raw`\[((?:[^\[\]\\]|\\.|\[[^\]]*\])*)\]` + LINK_TAIL, 'g')
const REF_LINK = /(!?)\[((?:[^[\]\\]|\\.)*)\](?:\[((?:[^[\]\\]|\\.)*)\])?/g

function inline(src: string, ctx: Ctx): string {
  const slots: string[] = []
  const hold = (html: string) => `\u0000${slots.push(html) - 1}\u0001`
  const unwrap = (u: string) => u.replace(/^<|>$/g, '')
  const link = (href: string, html: string, title?: string) => {
    const h = safeHref(href.replace(PUNCT, '$1'))
    // `html` is markup already: held whole so it is not escaped again
    return hold(h ? `<a href="${esc(h)}"${title ? ` title="${esc(title)}"` : ''}>${html}</a>` : html)
  }
  const image = (src: string, alt: string, title?: string) => {
    const s = safeSrc(src.replace(PUNCT, '$1'))
    const a = esc(alt.replace(PUNCT, '$1').replace(/[*_~`]/g, ''))
    return hold(s ? `<img src="${esc(s)}" alt="${a}"${title ? ` title="${esc(title)}"` : ''}>` : a)
  }
  let s = src.replace(/[\u0000\u0001]/g, '�')
  // code spans first: nothing inside them is markup
  s = s.replace(/(`+)(?!`)([\s\S]*?[^`])\1(?!`)/g, (_, _t: string, c: string) => {
    const t = c.replace(/\n/g, ' ')
    return hold(`<code>${esc(/^ .* $/.test(t) && t.trim() ? t.slice(1, -1) : t)}</code>`)
  })
  s = s.replace(/\\\n/g, () => hold('<br>\n'))
  s = s.replace(PUNCT, (_, c: string) => hold(esc(c)))
  s = s.replace(/<((?:https?|ftp):\/\/[^\s<>]+|mailto:[^\s<>]+)>/gi, (_, u: string) => link(u, esc(u)))
  s = s.replace(/<([\w.+-]+@[\w-]+(?:\.[\w-]+)+)>/g, (_, e: string) => link('mailto:' + e, esc(e)))
  s = s.replace(IMAGE, (_, alt: string, u: string, t1?: string, t2?: string, t3?: string) => image(unwrap(u), alt, t1 ?? t2 ?? t3))
  s = s.replace(LINK, (_, text: string, u: string, t1?: string, t2?: string, t3?: string) => link(unwrap(u), emphasis(text), t1 ?? t2 ?? t3))
  s = s.replace(REF_LINK, (m, bang: string, text: string, ref?: string) => {
    const def = ctx.refs.get(label(ref || text))
    if (!def) return m
    return bang ? image(def.href, text, def.title) : link(def.href, emphasis(text), def.title)
  })
  // bare web addresses
  s = s.replace(/(^|[\s(*_~])((?:https?:\/\/|www\.)[^\s<\u0000]*[^\s<\u0000.,:;"')\]!?*_~])/g, (_, pre: string, u: string) => pre + link(u.startsWith('www.') ? 'http://' + u : u, esc(u)))
  s = emphasis(s)
  // put the held parts back (they may hold others)
  for (let k = 0; k < 4 && s.includes('\u0000'); k++) s = s.replace(/\u0000(\d+)\u0001/g, (_, n: string) => slots[Number(n)])
  return s
}

/** escapes text (held parts aside) and marks strong / emphasis / strikethrough and hard breaks */
function emphasis(t: string): string {
  return esc(t)
    .replace(/\*\*\*(?=\S)([\s\S]*?\S)\*\*\*/g, '<strong><em>$1</em></strong>')
    .replace(/\*\*(?=\S)([\s\S]*?\S)\*\*/g, '<strong>$1</strong>')
    .replace(/(?<![\p{L}\p{N}_])__(?=\S)([\s\S]*?\S)__(?![\p{L}\p{N}_])/gu, '<strong>$1</strong>')
    .replace(/\*(?=[^\s*])([\s\S]*?[^\s*])\*/g, '<em>$1</em>')
    .replace(/(?<![\p{L}\p{N}_])_(?=[^\s_])([\s\S]*?[^\s_])_(?![\p{L}\p{N}_])/gu, '<em>$1</em>')
    .replace(/~~(?=\S)([\s\S]*?\S)~~/g, '<del>$1</del>')
    .replace(/ {2,}\n/g, '<br>\n')
}
