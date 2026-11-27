/**
 * SVG pictures read as drawing commands: the charts and diagrams the engine draws (`Svg` of diagram/svg.ts writes a
 * small, fixed set of elements), and SVG pictures in documents as far as they keep to the same set. Read: `svg`,
 * `g`, `rect`, `circle`, `ellipse`, `line`, `polyline`, `polygon`, `path`, `text` (one line, its font, size, weight,
 * style, anchor, underline), `image` (data: URLs of PNG / JPEG / GIF), `use` of those, presentation attributes and
 * simple `style`, transforms, linear and radial gradients, clip paths; a pattern paints with its first color. Scripts,
 * CSS style sheets, filters, masks and external references are not followed.
 */
import type { FontBook } from '../font/book'
import { parseXml, type XEl } from '../xml'
import { bounds, ellipsePath, multiply, parseColor, parsePath, rectPath, roundRectPath, transformPath, IDENTITY } from './path'
import { sniffImage } from './image'
import { textLine } from './text'
import type { DrawItem, Matrix, Paint, Rgba, Seg, Stop, StrokeStyle } from './types'

/**
 * Elements read from one picture, and path segments drawn, at most: counted as drawn, so `use` repeating `use`s
 * cannot multiply a small file into more (a crafted file cannot make one take the memory)
 */
export const MAX_SVG_ELEMENTS = 100_000
export const MAX_SVG_SEGMENTS = 1_000_000

export interface SvgReadOptions {
  /** the size the picture is drawn at (default its own) */
  width?: number
  height?: number
  onMissing?: (font: string) => void
}

interface Style {
  fill: string
  stroke: string
  strokeWidth: number
  dash?: number[]
  cap?: 'butt' | 'round' | 'square'
  join?: 'miter' | 'round' | 'bevel'
  fillOpacity: number
  strokeOpacity: number
  fillRule: 'nonzero' | 'evenodd'
  font: string
  fontSize: number
  bold: boolean
  italic: boolean
  anchor: 'start' | 'middle' | 'end'
  underline: boolean
}

const ROOT: Style = {
  fill: '#000000',
  stroke: 'none',
  strokeWidth: 1,
  fillOpacity: 1,
  strokeOpacity: 1,
  fillRule: 'nonzero',
  font: 'Helvetica, Arial, sans-serif',
  fontSize: 16,
  bold: false,
  italic: false,
  anchor: 'start',
  underline: false,
}

const local = (n: string) => n.replace(/^.*:/, '')

const num = (v: string | undefined, dflt = 0): number => {
  if (v === undefined) return dflt
  const x = parseFloat(v)
  return Number.isFinite(x) ? x : dflt
}

/** a length that may be a percentage of `of` */
const len = (v: string | undefined, of: number, dflt = 0): number => (v !== undefined && v.trim().endsWith('%') ? (num(v) / 100) * of : num(v, dflt))

/** `transform="…"` as a matrix */
export function parseTransform(v: string | undefined): Matrix {
  let m: Matrix = IDENTITY
  if (!v) return m
  const re = /(matrix|translate|scale|rotate|skewX|skewY)\s*\(([^)]*)\)/g
  for (let x = re.exec(v); x; x = re.exec(v)) {
    const a = x[2].split(/[\s,]+/).filter(Boolean).map(Number)
    let t: Matrix = IDENTITY
    switch (x[1]) {
      case 'matrix':
        if (a.length >= 6) t = [a[0], a[1], a[2], a[3], a[4], a[5]]
        break
      case 'translate':
        t = [1, 0, 0, 1, a[0] || 0, a[1] || 0]
        break
      case 'scale':
        t = [a[0] ?? 1, 0, 0, a[1] ?? a[0] ?? 1, 0, 0]
        break
      case 'rotate': {
        const r = ((a[0] || 0) * Math.PI) / 180
        const c = Math.cos(r)
        const s = Math.sin(r)
        const cx = a[1] || 0
        const cy = a[2] || 0
        t = [c, s, -s, c, cx - c * cx + s * cy, cy - s * cx - c * cy]
        break
      }
      case 'skewX':
        t = [1, 0, Math.tan(((a[0] || 0) * Math.PI) / 180), 1, 0, 0]
        break
      case 'skewY':
        t = [1, Math.tan(((a[0] || 0) * Math.PI) / 180), 0, 1, 0, 0]
        break
    }
    if (t.some((k) => !Number.isFinite(k))) continue
    // the functions apply right to left: the last one first
    m = multiply(t, m)
  }
  return m
}

/** Reads an SVG picture as drawing commands with (0, 0) at its top left. */
export async function svgItems(svg: string, book: FontBook, opts: SvgReadOptions = {}): Promise<DrawItem[]> {
  const root = parseXml(svg)
  if (local(root.n) !== 'svg') throw new Error('Not an SVG picture')
  const ids = new Map<string, XEl>()
  let count = 0
  const index = (el: XEl) => {
    if (++count > MAX_SVG_ELEMENTS) throw new Error(`An SVG picture of more than ${MAX_SVG_ELEMENTS} elements is not drawn`)
    if (el.a.id) ids.set(el.a.id, el)
    for (const c of el.c) index(c)
  }
  index(root)
  const vb = (root.a.viewBox ?? '').split(/[\s,]+/).map(Number)
  const hasVb = vb.length === 4 && vb.every(Number.isFinite) && vb[2] > 0 && vb[3] > 0
  const ownW = len(root.a.width, 0, hasVb ? vb[2] : 300) || (hasVb ? vb[2] : 300)
  const ownH = len(root.a.height, 0, hasVb ? vb[3] : 150) || (hasVb ? vb[3] : 150)
  const W = opts.width ?? ownW
  const H = opts.height ?? ownH
  // the viewBox stretched to the size (preserveAspectRatio none: the engine draws at the size it is asked for)
  const m: Matrix = hasVb ? [W / vb[2], 0, 0, H / vb[3], vb[0] ? (-vb[0] * W) / vb[2] : 0, vb[1] ? (-vb[1] * H) / vb[3] : 0] : [W / ownW, 0, 0, H / ownH, 0, 0]
  const reader = new SvgReader(ids, book, opts, hasVb ? vb[2] : ownW, hasVb ? vb[3] : ownH)
  const items = await reader.children(root, ROOT, 0)
  return [{ t: 'group', transform: m, clip: rectPath(hasVb ? vb[0] : 0, hasVb ? vb[1] : 0, hasVb ? vb[2] : ownW, hasVb ? vb[3] : ownH), items }]
}

const tooLarge = () => new Error(`An SVG picture of more than ${MAX_SVG_ELEMENTS} elements or ${MAX_SVG_SEGMENTS} path segments is not drawn`)

class SvgReader {
  /** elements and path segments drawn so far */
  private drawn = 0
  private segments = 0
  /** the ids a `use` is drawing: one that refers back to itself is drawn once */
  private using = new Set<string>()

  constructor(
    private ids: Map<string, XEl>,
    private book: FontBook,
    private opts: SvgReadOptions,
    private vw: number,
    private vh: number,
  ) {}

  async children(el: XEl, st: Style, depth: number): Promise<DrawItem[]> {
    const out: DrawItem[] = []
    if (depth > 64) return out
    for (const c of el.c) out.push(...(await this.element(c, st, depth + 1)))
    return out
  }

  private style(el: XEl, parent: Style): Style {
    const a: Record<string, string> = { ...el.a }
    // style="a: b; c: d" over the attributes
    for (const decl of (el.a.style ?? '').split(';')) {
      const k = decl.indexOf(':')
      if (k > 0) a[decl.slice(0, k).trim()] = decl.slice(k + 1).trim()
    }
    const s = { ...parent }
    if (a.fill) s.fill = a.fill
    if (a.stroke) s.stroke = a.stroke
    if (a['stroke-width']) s.strokeWidth = len(a['stroke-width'], Math.hypot(this.vw, this.vh) / Math.SQRT2, 1)
    if (a['stroke-dasharray']) s.dash = a['stroke-dasharray'] === 'none' ? undefined : a['stroke-dasharray'].split(/[\s,]+/).map(Number).filter((v) => Number.isFinite(v) && v >= 0)
    if (a['stroke-linecap'] === 'round' || a['stroke-linecap'] === 'square' || a['stroke-linecap'] === 'butt') s.cap = a['stroke-linecap']
    if (a['stroke-linejoin'] === 'round' || a['stroke-linejoin'] === 'bevel' || a['stroke-linejoin'] === 'miter') s.join = a['stroke-linejoin']
    if (a['fill-opacity']) s.fillOpacity = Math.max(0, Math.min(1, num(a['fill-opacity'], 1)))
    if (a['stroke-opacity']) s.strokeOpacity = Math.max(0, Math.min(1, num(a['stroke-opacity'], 1)))
    if (a['fill-rule'] === 'evenodd' || a['fill-rule'] === 'nonzero') s.fillRule = a['fill-rule']
    if (a['font-family']) s.font = a['font-family']
    if (a['font-size']) s.fontSize = num(a['font-size'], s.fontSize)
    if (a['font-weight']) s.bold = a['font-weight'] === 'bold' || a['font-weight'] === 'bolder' || num(a['font-weight'], 400) >= 600
    if (a['font-style']) s.italic = a['font-style'] === 'italic' || a['font-style'] === 'oblique'
    if (a['text-anchor'] === 'middle' || a['text-anchor'] === 'end' || a['text-anchor'] === 'start') s.anchor = a['text-anchor']
    if (a['text-decoration']) s.underline = a['text-decoration'].includes('underline')
    return s
  }

  /** a paint: a color, a gradient, a pattern's first color */
  private paint(v: string, opacity: number, box: [number, number, number, number] | null, el: XEl): Paint | null {
    if (!v || v === 'none' || v === 'transparent') return null
    const ref = /^url\(\s*['"]?#([^'")\s]+)['"]?\s*\)/.exec(v)
    if (ref) {
      const def = this.ids.get(ref[1])
      if (!def) return null
      const kind = local(def.n)
      if (kind === 'linearGradient' || kind === 'radialGradient') return this.gradient(def, opacity, box)
      if (kind === 'pattern') {
        // a pattern paints with its first fill: its background
        const first = def.c.find((c) => c.a.fill && c.a.fill !== 'none')
        const c = parseColor(first?.a.fill)
        return c ? { t: 'color', color: { ...c, a: c.a * opacity } } : null
      }
      return null
    }
    const c = v === 'currentColor' ? parseColor(el.a.color) ?? { r: 0, g: 0, b: 0, a: 1 } : parseColor(v)
    return c ? { t: 'color', color: { ...c, a: c.a * opacity } } : null
  }

  private gradient(def: XEl, opacity: number, box: [number, number, number, number] | null): Paint | null {
    // stops may come from the gradient it refers to
    let stopsEl = def
    for (let k = 0; k < 8 && !stopsEl.c.some((c) => local(c.n) === 'stop'); k++) {
      const href = (stopsEl.a.href ?? stopsEl.a['xlink:href'] ?? '').replace(/^#/, '')
      const next = this.ids.get(href)
      if (!next) break
      stopsEl = next
    }
    const stops: Stop[] = []
    for (const s of stopsEl.c) {
      if (local(s.n) !== 'stop') continue
      const st: Record<string, string> = { ...s.a }
      for (const decl of (s.a.style ?? '').split(';')) {
        const k = decl.indexOf(':')
        if (k > 0) st[decl.slice(0, k).trim()] = decl.slice(k + 1).trim()
      }
      const c = parseColor(st['stop-color'] ?? '#000')
      if (!c) continue
      const at = Math.max(0, Math.min(1, st.offset?.trim().endsWith('%') ? num(st.offset) / 100 : num(st.offset)))
      stops.push({ at: Math.max(at, stops[stops.length - 1]?.at ?? 0), color: { ...c, a: c.a * num(st['stop-opacity'], 1) * opacity } })
    }
    if (!stops.length) return null
    if (stops.length === 1) return { t: 'color', color: stops[0].color }
    const user = def.a.gradientUnits === 'userSpaceOnUse'
    // objectBoundingBox: fractions of the shape's box
    const [x0, y0, x1, y1] = box ?? [0, 0, 1, 1]
    const bw = x1 - x0
    const bh = y1 - y0
    const X = (v: string | undefined, d: string) => (user ? len(v ?? d, this.vw) : x0 + (v?.trim().endsWith('%') ? num(v) / 100 : num(v ?? d.replace('%', ''), 0) / (d.endsWith('%') && v === undefined ? 100 : 1)) * bw)
    const Y = (v: string | undefined, d: string) => (user ? len(v ?? d, this.vh) : y0 + (v?.trim().endsWith('%') ? num(v) / 100 : num(v ?? d.replace('%', ''), 0) / (d.endsWith('%') && v === undefined ? 100 : 1)) * bh)
    if (local(def.n) === 'linearGradient') return { t: 'linear', x0: X(def.a.x1, '0%'), y0: Y(def.a.y1, '0%'), x1: X(def.a.x2, '100%'), y1: Y(def.a.y2, '0%'), stops }
    const cx = X(def.a.cx, '50%')
    const cy = Y(def.a.cy, '50%')
    const r = user ? len(def.a.r ?? '50%', Math.hypot(this.vw, this.vh) / Math.SQRT2) : (def.a.r?.trim().endsWith('%') ? num(def.a.r) / 100 : def.a.r ? num(def.a.r) : 0.5) * Math.hypot(bw, bh) / Math.SQRT2
    return { t: 'radial', cx, cy, r, fx: def.a.fx ? X(def.a.fx, '50%') : cx, fy: def.a.fy ? Y(def.a.fy, '50%') : cy, stops }
  }

  private shapePath(el: XEl): Seg[] | null {
    const a = el.a
    switch (local(el.n)) {
      case 'rect': {
        const w = len(a.width, this.vw)
        const h = len(a.height, this.vh)
        if (w <= 0 || h <= 0) return null
        const x = len(a.x, this.vw)
        const y = len(a.y, this.vh)
        const rx = a.rx ?? a.ry
        const ry = a.ry ?? a.rx
        return rx ? roundRectPath(x, y, w, h, len(rx, this.vw), len(ry, this.vh)) : rectPath(x, y, w, h)
      }
      case 'circle': {
        const r = len(a.r, Math.hypot(this.vw, this.vh) / Math.SQRT2)
        return r > 0 ? ellipsePath(len(a.cx, this.vw), len(a.cy, this.vh), r, r) : null
      }
      case 'ellipse': {
        const rx = len(a.rx, this.vw)
        const ry = len(a.ry, this.vh)
        return rx > 0 && ry > 0 ? ellipsePath(len(a.cx, this.vw), len(a.cy, this.vh), rx, ry) : null
      }
      case 'line':
        return [{ k: 'M', x: len(a.x1, this.vw), y: len(a.y1, this.vh) }, { k: 'L', x: len(a.x2, this.vw), y: len(a.y2, this.vh) }]
      case 'polyline':
      case 'polygon': {
        const p = (a.points ?? '').split(/[\s,]+/).filter(Boolean).map(Number)
        const d: Seg[] = []
        for (let i = 0; i + 1 < p.length; i += 2) d.push({ k: i ? 'L' : 'M', x: p[i], y: p[i + 1] })
        if (local(el.n) === 'polygon' && d.length) d.push({ k: 'Z' })
        return d.length ? d : null
      }
      case 'path':
        return a.d ? parsePath(a.d) : null
    }
    return null
  }

  private clip(el: XEl): { clip?: Seg[]; clipRule?: 'evenodd' } {
    const ref = /^url\(\s*['"]?#([^'")\s]+)['"]?\s*\)/.exec(el.a['clip-path'] ?? '')
    const def = ref && this.ids.get(ref[1])
    if (!def || local(def.n) !== 'clipPath') return {}
    const d: Seg[] = []
    let evenodd = false
    for (const c of def.c) {
      const p = this.shapePath(c)
      if (!p) continue
      // (a clip's segments count as drawn: each element it clips draws them)
      if ((this.segments += p.length) > MAX_SVG_SEGMENTS) throw tooLarge()
      const t = parseTransform(c.a.transform)
      d.push(...(t === IDENTITY ? p : transformPath(p, t)))
      if (c.a['clip-rule'] === 'evenodd') evenodd = true
    }
    return d.length ? { clip: d, ...(evenodd ? { clipRule: 'evenodd' as const } : {}) } : { clip: [] }
  }

  private async element(el: XEl, parent: Style, depth: number): Promise<DrawItem[]> {
    if (++this.drawn > MAX_SVG_ELEMENTS) throw tooLarge()
    const name = local(el.n)
    if (name === 'defs' || name === 'clipPath' || name === 'linearGradient' || name === 'radialGradient' || name === 'pattern' || name === 'style' || name === 'script' || name === 'title' || name === 'desc' || name === 'metadata' || name === 'mask' || name === 'filter' || name === 'symbol' || name === 'marker') return []
    if (el.a.display === 'none' || el.a.visibility === 'hidden') return []
    const st = this.style(el, parent)
    const transform = parseTransform(el.a.transform)
    const opacity = Math.max(0, Math.min(1, num(el.a.opacity, 1)))
    let items: DrawItem[] = []
    if (name === 'g' || name === 'svg' || name === 'a' || name === 'switch') items = await this.children(el, st, depth)
    else if (name === 'use') {
      const id = (el.a.href ?? el.a['xlink:href'] ?? '').replace(/^#/, '')
      const ref = this.ids.get(id)
      if (ref && depth < 32 && !this.using.has(id)) {
        this.using.add(id)
        try {
          items = await this.element(ref, st, depth + 1)
        } finally {
          this.using.delete(id)
        }
        const dx = num(el.a.x)
        const dy = num(el.a.y)
        if (dx || dy) items = [{ t: 'group', transform: [1, 0, 0, 1, dx, dy], items }]
      }
    } else if (name === 'text') items = await this.text(el, st)
    else if (name === 'image') items = this.image(el)
    else {
      const d = this.shapePath(el)
      if (d && (this.segments += d.length) > MAX_SVG_SEGMENTS) throw tooLarge()
      if (d && d.length) {
        const box = bounds(d)
        const fill = name === 'line' ? null : this.paint(st.fill, st.fillOpacity, box, el)
        const sp = this.paint(st.stroke, st.strokeOpacity, box, el)
        const stroke: StrokeStyle | undefined = sp && st.strokeWidth > 0 ? { paint: sp, width: st.strokeWidth, ...(st.dash?.length && st.dash.some((v) => v > 0) ? { dash: st.dash } : {}), ...(st.cap ? { cap: st.cap } : {}), ...(st.join ? { join: st.join } : {}) } : undefined
        if (fill || stroke) items = [{ t: 'path', d, ...(fill ? { fill } : {}), ...(st.fillRule === 'evenodd' ? { rule: 'evenodd' as const } : {}), ...(stroke ? { stroke } : {}) }]
      }
    }
    if (!items.length) return []
    const clip = this.clip(el)
    if (clip.clip && !clip.clip.length) return []
    if (transform === IDENTITY && !clip.clip && opacity >= 1) return items
    return [{ t: 'group', ...(transform !== IDENTITY ? { transform } : {}), ...clip, ...(opacity < 1 ? { opacity } : {}), items }]
  }

  private async text(el: XEl, st: Style): Promise<DrawItem[]> {
    // one line: the element's text and its tspans', at the element's place
    const text = textOf(el).replace(/\s+/g, ' ').trim()
    if (!text) return []
    const fill = this.paint(st.fill, st.fillOpacity, null, el)
    const color: Rgba = fill?.t === 'color' ? fill.color : fill ? (fill.stops[0]?.color ?? { r: 0, g: 0, b: 0, a: 1 }) : { r: 0, g: 0, b: 0, a: 0 }
    if (color.a <= 0) return []
    const x = num(el.a.x?.split(/[\s,]+/)[0])
    const y = num(el.a.y?.split(/[\s,]+/)[0])
    const line = await textLine(this.book, text, { font: st.font, size: st.fontSize, bold: st.bold, italic: st.italic, color, underline: st.underline }, { onMissing: this.opts.onMissing })
    const dx = st.anchor === 'middle' ? -line.width / 2 : st.anchor === 'end' ? -line.width : 0
    return [{ t: 'group', transform: [1, 0, 0, 1, x + dx, y], items: line.items }]
  }

  private image(el: XEl): DrawItem[] {
    const href = el.a.href ?? el.a['xlink:href'] ?? ''
    const m = /^data:([^;,]*)(;base64)?,(.*)$/s.exec(href)
    if (!m || !m[2]) return []
    let bytes: Uint8Array
    try {
      const bin = atob(m[3].replace(/\s+/g, ''))
      bytes = new Uint8Array(bin.length)
      for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i)
    } catch {
      return []
    }
    const kind = sniffImage(bytes)
    if (kind !== 'png' && kind !== 'jpeg' && kind !== 'gif') return []
    const w = len(el.a.width, this.vw)
    const h = len(el.a.height, this.vh)
    if (w <= 0 || h <= 0) return []
    return [{ t: 'image', image: { bytes, type: `image/${kind}` }, x: len(el.a.x, this.vw), y: len(el.a.y, this.vh), w, h }]
  }
}

function textOf(el: XEl): string {
  let s = el.t ?? ''
  for (const c of el.c) if (local(c.n) === 'tspan' || local(c.n) === 'a') s += textOf(c)
  return s
}
