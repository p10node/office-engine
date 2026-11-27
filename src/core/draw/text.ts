/**
 * Paragraphs laid out in a box without a browser: text measured in the host's font files (`FontBook`), lines broken
 * where CSS breaks them (`white-space: pre-wrap`, `overflow-wrap: break-word`), aligned, spaced as CSS spaces lines
 * of a unitless `line-height`, and drawn as glyph runs. Text no file serves is set in Helvetica. Slides, sheets and
 * documents all lay out their text with it.
 */
import type { FontBook } from '../font/book'
import type { Font } from '../font/font'
import { stdCode, stdFont, stdWidth } from './stdfont'
import { rectPath } from './path'
import type { DrawItem, Rgba, StdFont } from './types'

export interface RunStyle {
  /** a CSS font-family list (`cssFont()`) */
  font: string
  /** px */
  size: number
  bold?: boolean
  italic?: boolean
  color: Rgba
  underline?: boolean
  strike?: boolean
  /** raised (positive) or lowered, in percent of the size: drawn smaller, as superscript or subscript */
  baseline?: number
  highlight?: Rgba
  caps?: boolean
  href?: string
}

export interface ParaRun {
  text: string
  style: RunStyle
  /** a line break */
  br?: boolean
}

export interface ParaSpec {
  runs: ParaRun[]
  align: 'left' | 'center' | 'right' | 'justify'
  /** px: the left margin of every line, and how far the first line starts from it (negative: a hanging indent) */
  marL: number
  indent: number
  /** `line-height` as a multiple of each run's size (default 1.2), or in px */
  lineMult?: number
  linePx?: number
  /** px above and below */
  before: number
  after: number
  bullet?: { text: string; style: RunStyle }
  /** the size of an empty paragraph's line (px) */
  emptySize: number
}

export interface LayoutOptions {
  /** break lines at the width (default true) */
  wrap?: boolean
  /** told of text no font file serves (set in Helvetica) */
  onMissing?: (font: string) => void
}

/** a measured piece of a run: glyphs from one font, positions in px, clusters in the run's text */
interface Piece {
  font: Font | StdFont
  glyphs: number[]
  advances: number[]
  dx?: number[]
  dy?: number[]
  clusters: number[]
  /** the run's text the piece stands for: [from, to) */
  from: number
  to: number
  /** clusters in order ('asc'), in reverse (right to left, 'desc'), or neither (a shaper reordered them) */
  order: 'asc' | 'desc' | null
  synthetic?: { bold?: boolean; italic?: boolean }
  ascent: number
  descent: number
}

interface Measured {
  text: string
  style: RunStyle
  /** the drawn size: smaller for superscript and subscript */
  size: number
  rise: number
  pieces: Piece[]
  /** for each UTF-16 unit of the text, how far it moves the pen (a ligature's width goes to its first character) */
  adv: Float64Array
  /** 1 where a cluster starts: a line is never broken inside one (a character of two UTF-16 units, a letter and its marks) */
  starts: Uint8Array
  ascent: number
  descent: number
  br?: boolean
}

/** Helvetica's line metrics as browsers take them from Arial, which stands in for it */
const STD_ASCENT = 0.905
const STD_DESCENT = 0.212

const DEFAULT_TAB = 96

async function measureRun(book: FontBook, run: ParaRun, opts: LayoutOptions): Promise<Measured> {
  const st = run.style
  const text = st.caps ? run.text.toUpperCase() : run.text
  const size = st.baseline ? st.size * 0.65 : st.size
  const rise = st.baseline ? (st.baseline / 100) * st.size : 0
  const weight = st.bold ? 700 : 400
  const m = await book.measure({ text, font: st.font, size, weight, italic: !!st.italic })
  const adv = new Float64Array(text.length)
  const starts = new Uint8Array(text.length)
  const pieces: Piece[] = []
  let ascent = m.ascent
  let descent = m.descent
  for (const p of m.pieces) {
    if (p.font && p.run) {
      const k = size / p.font.unitsPerEm
      const r = p.run
      const clusters = r.clusters.map((c) => c + p.from)
      const advances = r.advances.map((a) => a * k)
      clusters.forEach((c, i) => ((adv[c] += advances[i]), (starts[c] = 1)))
      const line = p.font.lineMetrics()
      const synthetic = { bold: weight >= 600 && p.font.weight < 600, italic: !!st.italic && !p.font.italic }
      pieces.push({
        font: p.font,
        glyphs: r.glyphs,
        advances,
        ...(r.dx.some((v) => v) || r.dy.some((v) => v) ? { dx: r.dx.map((v) => v * k), dy: r.dy.map((v) => v * k) } : {}),
        clusters,
        from: p.from,
        to: p.to,
        order: orderOf(clusters),
        ...(synthetic.bold || synthetic.italic ? { synthetic } : {}),
        ascent: line.ascent * k,
        descent: line.descent * k,
      })
    } else {
      // no file: Helvetica, one code per character
      const font = stdFont(st.bold, st.italic)
      const glyphs: number[] = []
      const advances: number[] = []
      const clusters: number[] = []
      for (let i = p.from; i < p.to; ) {
        const cp = text.codePointAt(i)!
        const code = stdCode(cp)
        const w = (stdWidth(code, !!st.bold) * size) / 1000
        glyphs.push(code)
        advances.push(w)
        clusters.push(i)
        adv[i] += w
        starts[i] = 1
        i += cp > 0xffff ? 2 : 1
      }
      if (!/^\s*$/.test(text.slice(p.from, p.to))) opts.onMissing?.(st.font)
      pieces.push({ font, glyphs, advances, clusters, from: p.from, to: p.to, order: 'asc', ascent: STD_ASCENT * size, descent: STD_DESCENT * size })
    }
  }
  if (!ascent && !descent) {
    ascent = STD_ASCENT * size
    descent = STD_DESCENT * size
  }
  return { text, style: st, size, rise, pieces, adv, starts, ascent, descent, br: run.br }
}

function orderOf(clusters: number[]): 'asc' | 'desc' | null {
  let asc = true
  let desc = true
  for (let i = 1; i < clusters.length; i++) {
    if (clusters[i] < clusters[i - 1]) asc = false
    if (clusters[i] > clusters[i - 1]) desc = false
  }
  return asc ? 'asc' : desc ? 'desc' : null
}

/** the glyphs of a piece that stand for text in [a, b), in drawing order: found by halving where clusters are in order */
function glyphIndices(p: Piece, a: number, b: number): number[] {
  const cl = p.clusters
  const first = (pred: (c: number) => boolean) => {
    let lo = 0
    let hi = cl.length
    while (lo < hi) {
      const mid = (lo + hi) >> 1
      if (pred(cl[mid])) hi = mid
      else lo = mid + 1
    }
    return lo
  }
  const span = (lo: number, hi: number) => Array.from({ length: Math.max(0, hi - lo) }, (_, k) => lo + k)
  if (p.order === 'asc') return span(first((c) => c >= a), first((c) => c >= b))
  if (p.order === 'desc') return span(first((c) => c < b), first((c) => c < a))
  // (reordered by a shaper: every glyph looked at)
  const out: number[] = []
  cl.forEach((c, i) => c >= a && c < b && out.push(i))
  return out
}

/** a place in a paragraph: run and UTF-16 index in its text */
interface Pos {
  run: number
  at: number
}

interface Line {
  from: Pos
  /** exclusive */
  to: Pos
  /** the width of what is drawn (without trailing spaces) */
  width: number
  /** the line ends the paragraph or at a break: not stretched when justified */
  last: boolean
}

const isSpace = (c: string) => c === ' ' || c === '\t' || c === '　'
/** characters a line breaks after or around: CJK ideographs, kana, Hangul (each a word of its own) */
const isCjk = (cp: number) => (cp >= 0x2e80 && cp <= 0x9fff) || (cp >= 0xac00 && cp <= 0xd7af) || (cp >= 0xf900 && cp <= 0xfaff) || (cp >= 0xff00 && cp <= 0xffef) || (cp >= 0x20000 && cp <= 0x323af)

export interface LaidOut {
  /** drawing commands with (0, 0) at the box's top left */
  items: DrawItem[]
  /** the height the paragraphs take */
  height: number
  /** the widest line */
  width: number
}

/**
 * Lays out paragraphs in a box `width` px wide, the first line's top at 0. Text that does not fit the height is laid
 * out all the same: the caller anchors and clips.
 */
export async function layoutParagraphs(book: FontBook, paras: ParaSpec[], width: number, opts: LayoutOptions = {}): Promise<LaidOut> {
  const items: DrawItem[] = []
  let y = 0
  let widest = 0
  for (const p of paras) {
    y += p.before
    const runs: Measured[] = []
    for (const r of p.runs) runs.push(await measureRun(book, r, opts))
    const bullet = p.bullet && p.runs.some((r) => r.text || r.br) ? await measureRun(book, { text: p.bullet.text, style: p.bullet.style }, opts) : undefined
    // the bullet takes the hanging indent, or its own width and 0.4 em
    const hang = p.indent < 0 ? -p.indent : 0
    const bulletW = bullet ? Math.max(hang, bullet.adv.reduce((a, b) => a + b, 0) + (hang ? 0 : 0.4 * bullet.size)) : 0
    const firstX = p.marL + p.indent + bulletW
    const avail = (first: boolean) => Math.max(0, width - (first ? firstX : p.marL))
    const lines = breakLines(runs, avail, opts.wrap !== false)
    const empty = !lines.length || runs.every((r) => !r.text.length)
    if (empty) {
      // an empty paragraph is one line of its end size
      const h = p.linePx ?? p.emptySize * (p.lineMult ?? 1.2)
      y += h + p.after
      continue
    }
    lines.forEach((line, li) => {
      const x0 = li === 0 ? firstX : p.marL
      // the line's height: each run's box (its size × the multiple, or the px height) around its baseline, as CSS
      // centres an inline box's glyphs in its line-height
      let above = 0
      let below = 0
      forEach(runs, line, (r) => {
        const lh = p.linePx ?? r.size * (p.lineMult ?? 1.2)
        const half = (lh - (r.ascent + r.descent)) / 2
        above = Math.max(above, r.ascent + half + r.rise)
        below = Math.max(below, r.descent + half - r.rise)
      })
      if (li === 0 && bullet) {
        const lh = p.linePx ?? bullet.size * (p.lineMult ?? 1.2)
        const half = (lh - (bullet.ascent + bullet.descent)) / 2
        above = Math.max(above, bullet.ascent + half)
        below = Math.max(below, bullet.descent + half)
      }
      // a line of nothing but a break: as high as an empty paragraph
      if (!above && !below) above = p.linePx ?? p.emptySize * (p.lineMult ?? 1.2)
      const baseline = y + above
      const room = avail(li === 0)
      const extra = Math.max(0, room - line.width)
      let x = x0 + (p.align === 'center' ? extra / 2 : p.align === 'right' ? extra : 0)
      // justified: the room left goes to the spaces between words (not to those that end the line)
      const just = { stretch: 0, spaces: 0 }
      if (p.align === 'justify' && !line.last) {
        let pending = 0
        forEach(runs, line, (r, a, b) => {
          for (let i = a; i < b; i++) {
            if (r.text[i] === ' ') pending++
            else if (!isSpace(r.text[i])) (just.spaces += pending), (pending = 0)
          }
        })
        just.stretch = just.spaces ? extra / just.spaces : 0
      }
      if (li === 0 && bullet) emit(items, bullet, 0, bullet.text.length, p.marL + p.indent, baseline, { stretch: 0, spaces: 0 })
      forEach(runs, line, (r, a, b) => {
        x = emit(items, r, a, b, x, baseline, just)
      })
      widest = Math.max(widest, x0 + line.width)
      y = baseline + below
    })
    y += p.after
  }
  return { items, height: y, width: widest }
}

/** each run's part of a line: `fn(run, from, to)` */
function forEach(runs: Measured[], line: Line, fn: (r: Measured, from: number, to: number) => void) {
  for (let k = line.from.run; k <= line.to.run && k < runs.length; k++) {
    const r = runs[k]
    const a = k === line.from.run ? line.from.at : 0
    const b = k === line.to.run ? line.to.at : r.text.length
    if (b > a || (r.br && k < line.to.run)) fn(r, a, b)
  }
}

/** greedy line breaking: a line ends at the last break opportunity that fits, or inside a word longer than the line */
function breakLines(runs: Measured[], avail: (first: boolean) => number, wrap: boolean): Line[] {
  const lines: Line[] = []
  let start: Pos = { run: 0, at: 0 }
  let width = 0
  /** after the last break opportunity: where the line would end, and its width without trailing spaces */
  let brk: { pos: Pos; width: number } | null = null
  /** the width up to the end of the last non-space */
  let inked = 0
  /** the first character of the cluster the loop is in */
  let base = 0
  const push = (to: Pos, w: number, last: boolean) => {
    lines.push({ from: start, to, width: w, last })
    start = to
  }
  for (let k = 0; k < runs.length; k++) {
    const r = runs[k]
    if (r.br) {
      push({ run: k, at: 0 }, inked, true)
      start = { run: k + 1, at: 0 }
      width = inked = 0
      brk = null
      continue
    }
    for (let i = 0; i < r.text.length; i++) {
      const c = r.text[i]
      const cp = r.text.codePointAt(i)!
      let w = r.adv[i]
      // inside a cluster: no break here, nor a line end; after its last unit, a line can end after an ideograph
      if (i > 0 && !r.starts[i] && c !== '\n' && c !== '\r' && c !== '\v' && c !== '\t') {
        width += w
        inked = width
        if (isCjk(base) && (i + 1 >= r.text.length || r.starts[i + 1])) brk = { pos: { run: k, at: i + 1 }, width: inked }
        continue
      }
      base = cp
      if (c === '\t') {
        // to the next default tab stop
        const at = width
        w = (Math.floor(at / DEFAULT_TAB) + 1) * DEFAULT_TAB - at
        r.adv[i] = w
      }
      if (c === '\n' || c === '\r' || c === '\v') {
        push({ run: k, at: i }, inked, true)
        start = { run: k, at: i + 1 }
        width = inked = 0
        brk = null
        continue
      }
      const space = isSpace(c)
      // a break before an ideograph
      if (!space && isCjk(cp) && width > 0) brk = { pos: { run: k, at: i }, width: inked }
      if (!space && wrap && width + w > avail(lines.length === 0) + 0.01 && width > 0) {
        // past the end: break at the last opportunity, else here (a word longer than the line)
        const at = brk ?? { pos: { run: k, at: i }, width: inked }
        push(at.pos, at.width, false)
        // measure again from the new start to here
        width = 0
        inked = 0
        for (let kk = at.pos.run; kk <= k; kk++) {
          const rr = runs[kk]
          for (let ii = kk === at.pos.run ? at.pos.at : 0; ii < (kk === k ? i : rr.text.length); ii++) {
            width += rr.adv[ii]
            if (!isSpace(rr.text[ii])) inked = width
          }
        }
        brk = null
      }
      width += w
      if (!space) inked = width
      // after spaces, hyphens and ideographs, a line can end
      if (space || c === '-' || c === '‐' || c === '/' || isCjk(cp)) {
        const next = i + 1 < r.text.length ? r.text[i + 1] : k + 1 < runs.length ? runs[k + 1].text[0] : ''
        const whole = i + 1 >= r.text.length || r.starts[i + 1]
        if (whole && !(space && next && isSpace(next))) brk = { pos: { run: k, at: i + 1 }, width: inked }
      }
    }
  }
  const end = { run: runs.length - 1, at: runs[runs.length - 1]?.text.length ?? 0 }
  if (runs.length && (start.run < end.run || (start.run === end.run && start.at < end.at) || !lines.length)) push(end, inked, true)
  return lines
}

/** draws `r.text[a, b)` from the pen at `x` on `baseline`; returns where the pen ends */
function emit(items: DrawItem[], r: Measured, a: number, b: number, x: number, baseline: number, just: { stretch: number; spaces: number }): number {
  const st = r.style
  const mark = items.length
  const y = baseline - r.rise
  const startX = x
  // the pieces in the line: they follow one another through the text, so the first is found by halving
  let lo = 0
  let hi = r.pieces.length
  while (lo < hi) {
    const mid = (lo + hi) >> 1
    if (r.pieces[mid].to > a) hi = mid
    else lo = mid + 1
  }
  for (let k = lo; k < r.pieces.length && r.pieces[k].from < b; k++) {
    const p = r.pieces[k]
    // the piece's part of the line: its own text, so a copy gives each character once
    const from = Math.max(a, p.from)
    const to = Math.min(b, p.to)
    if (to <= from) continue
    const indices = glyphIndices(p, from, to)
    const glyphs: number[] = []
    const advances: number[] = []
    const clusters: number[] = []
    const dx: number[] = []
    const dy: number[] = []
    let px = x
    for (const i of indices) {
      const c = p.clusters[i]
      const ch = r.text[c]
      let adv = ch === '\t' ? r.adv[c] : p.advances[i]
      if (just.spaces > 0 && ch === ' ') (adv += just.stretch), just.spaces--
      x += adv
      if (ch === '\t') {
        // a tab moves the pen and draws nothing (its glyph would be a box, or a question mark)
        if (advances.length) advances[advances.length - 1] += adv
        else px = x
        continue
      }
      glyphs.push(p.glyphs[i])
      advances.push(adv)
      clusters.push(c - from)
      if (p.dx) dx.push(p.dx[i]), dy.push(p.dy![i])
    }
    if (!glyphs.length) continue
    const text = r.text.slice(from, to)
    // spaces draw nothing: a run of them is only a move
    if (st.color.a > 0 && /\S/.test(text)) {
      items.push({
        t: 'text',
        x: px,
        y,
        size: r.size,
        font: p.font,
        glyphs,
        advances,
        ...(p.dx ? { dx, dy } : {}),
        text,
        clusters,
        fill: st.color,
        ...(p.synthetic ? { synthetic: p.synthetic } : {}),
      })
    }
  }
  const w = x - startX
  if (w > 0) {
    const top = y - r.ascent
    const h = r.ascent + r.descent
    // (a highlight is behind the run's text)
    if (st.highlight) items.splice(mark, 0, { t: 'path', d: rectPath(startX, top, w, h), fill: { t: 'color', color: st.highlight } })
    // the lines a browser draws: under the baseline by a tenth of the size, through it at a third
    const thick = Math.max(0.5, r.size / 15)
    if (st.underline) items.push({ t: 'path', d: rectPath(startX, y + r.size * 0.1, w, thick), fill: { t: 'color', color: st.color } })
    if (st.strike) items.push({ t: 'path', d: rectPath(startX, y - r.size * 0.3, w, thick), fill: { t: 'color', color: st.color } })
    if (st.href) items.push({ t: 'link', x: startX, y: top, w, h, url: st.href })
  }
  return x
}

/** One line of text, not broken: its drawing commands from the pen at (0, 0) on the baseline, and its width. */
export async function textLine(book: FontBook, text: string, style: RunStyle, opts: LayoutOptions = {}): Promise<{ items: DrawItem[]; width: number }> {
  const r = await measureRun(book, { text, style }, opts)
  const items: DrawItem[] = []
  const width = emit(items, r, 0, r.text.length, 0, 0, { stretch: 0, spaces: 0 })
  return { items, width }
}
