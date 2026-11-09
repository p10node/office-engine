import type { Block, Inline, Paragraph, Table } from '../../core/docx/types'
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

/** `heading`: bold is the heading's own look, not emphasis */
function mdInline(runs: Inline[], heading = false): string {
  let out = ''
  let i = 0
  while (i < runs.length) {
    const r = runs[i]
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
    out += `${lead}${strike}${mark}${mdEsc(core)}${[...mark].reverse().join('')}${strike}${trail}`
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

/** Markdown: headings (outline levels), lists, bold / italic / strike, links, tables. */
export function docxToMarkdown(blocks: Block[]): string {
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
    const text = mdInline(b.runs, heading).trim()
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
