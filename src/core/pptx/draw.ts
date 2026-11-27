/**
 * A slide as a page of drawing commands (#46): shapes, pictures, tables and text placed as the viewer places them
 * (`SlideView`), text laid out in the host's font files. A PDF or SVG backend paints the page.
 */
import type { FontBook } from '../font/book'
import { ellipsePath, multiply, parseColor, parsePath, rectPath, rotate, roundRectPath, translate, IDENTITY } from '../draw/path'
import { sniffImage } from '../draw/image'
import { svgItems } from '../draw/svg-in'
import { layoutParagraphs, type ParaSpec, type RunStyle } from '../draw/text'
import { renderChart } from '../xlsx/chart-svg'
import type { DrawItem, ImageData, Matrix, Page, Paint, Rgba, Seg, StrokeStyle } from '../draw/types'
import { safeHref } from '../url'
import { geomPath, LINE_GEOMS } from './geometry'
import type { Box, Fill, SlideEl, SlideModel, Stroke, TCell, TextBody } from './types'

export interface SlideDrawOptions {
  /** the host's font files */
  fonts: FontBook
  /** a part's bytes (pictures), by its name in the package */
  part: (name: string) => Promise<Uint8Array | null>
  /** the font of text that names none (default Calibri, as PowerPoint's default theme) */
  defaultFont?: string
  onWarning?: (text: string) => void
}

const PT = 96 / 72
const BLACK: Rgba = { r: 0, g: 0, b: 0, a: 1 }
const WHITE: Rgba = { r: 255, g: 255, b: 255, a: 1 }

/** Lays out one slide as a page. */
export async function slidePage(slide: SlideModel, opts: SlideDrawOptions): Promise<Page> {
  const d = new SlideDrawer(opts)
  const items: DrawItem[] = []
  const all = rectPath(0, 0, slide.w, slide.h)
  // the page is white under a background that lets it through
  items.push({ t: 'path', d: all, fill: { t: 'color', color: WHITE } })
  if (slide.bg) items.push(...(await d.fill(slide.bg, all, { x: 0, y: 0, w: slide.w, h: slide.h })))
  for (const el of slide.els) items.push(...(await d.element(el)))
  return { w: slide.w, h: slide.h, items: [{ t: 'group', clip: all, items }], ...(slide.title ? { title: slide.title } : {}) }
}

class SlideDrawer {
  private images = new Map<string, Promise<ImageData | null>>()

  constructor(private opts: SlideDrawOptions) {}

  private warn(text: string) {
    this.opts.onWarning?.(text)
  }

  async element(el: SlideEl): Promise<DrawItem[]> {
    switch (el.t) {
      case 'shape':
        return this.shape(el)
      case 'pic':
        return this.picture(el)
      case 'table':
        return this.table(el)
      case 'chart': {
        // the viewer's chart picture, read back as drawing commands
        const w = Math.max(1, Math.round(el.box.w))
        const h = Math.max(1, Math.round(el.box.h))
        try {
          const items = await svgItems(renderChart(el.model, w, h), this.opts.fonts, { width: el.box.w, height: el.box.h, onMissing: (f) => this.missing(f) })
          return [{ t: 'group', transform: boxMatrix(el.box), items }]
        } catch (e) {
          this.warn(`A chart could not be drawn: ${(e as Error).message}`)
          return []
        }
      }
      case 'placeholder':
        this.warn(`${el.label} is not drawn in a PDF`)
        return []
    }
  }

  private image(target: string): Promise<ImageData | null> {
    let p = this.images.get(target)
    if (!p) {
      p = Promise.resolve()
        .then(() => this.opts.part(target))
        .then(
        (bytes) => {
          if (!bytes) return null
          const kind = sniffImage(bytes)
          return { bytes, type: kind ? `image/${kind === 'svg' ? 'svg+xml' : kind}` : 'application/octet-stream', key: target }
        },
        // a part that cannot be read leaves its picture out, not the slide
        (e: unknown) => {
          this.warn(`A picture could not be read: ${reason(e)}`)
          return null
        },
      )
      this.images.set(target, p)
    }
    return p
  }

  /** a picture stretched over a box: SVG pictures as drawing commands, others as they are */
  private async picItem(image: ImageData, x: number, y: number, w: number, h: number): Promise<DrawItem> {
    if (image.type === 'image/svg+xml') {
      try {
        const items = await svgItems(new TextDecoder().decode(image.bytes), this.opts.fonts, { width: w, height: h, onMissing: (f) => this.missing(f) })
        return { t: 'group', transform: translate(x, y), items }
      } catch (e) {
        this.warn(`An SVG picture could not be drawn: ${(e as Error).message}`)
        return { t: 'group', items: [] }
      }
    }
    return { t: 'image', image, x, y, w, h }
  }

  /** a fill painted inside `clip`, `box` the area a picture or gradient spans */
  async fill(fill: Fill, clip: Seg[], box: { x: number; y: number; w: number; h: number }, rule?: 'evenodd'): Promise<DrawItem[]> {
    if (fill.t === 'solid') {
      const c = parseColor(fill.c)
      return c ? [{ t: 'path', d: clip, fill: { t: 'color', color: c }, ...(rule ? { rule } : {}) }] : []
    }
    if (fill.t === 'grad') {
      const paint = gradient(fill.css, box)
      return paint ? [{ t: 'path', d: clip, fill: paint, ...(rule ? { rule } : {}) }] : []
    }
    if (fill.t === 'img') {
      const image = await this.image(fill.target)
      if (!image) return []
      // (a tiled picture is stretched: tiles are not drawn yet)
      return [{ t: 'group', clip, ...(rule ? { clipRule: rule } : {}), items: [await this.picItem(image, box.x, box.y, box.w, box.h)] }]
    }
    return []
  }

  private async shape(el: Extract<SlideEl, { t: 'shape' }>): Promise<DrawItem[]> {
    const { box } = el
    const items: DrawItem[] = []
    const isLine = LINE_GEOMS.has(el.geom)
    const d = parsePath(el.path ?? geomPath(el.geom, box.w, box.h, el.adj))
    // the outline flips with the shape, its text does not
    const flip: Matrix = [box.flipH ? -1 : 1, 0, 0, box.flipV ? -1 : 1, box.flipH ? box.w : 0, box.flipV ? box.h : 0]
    const shapeItems: DrawItem[] = []
    if (!isLine && el.fill) shapeItems.push(...(await this.fill(el.fill, d, { x: 0, y: 0, w: box.w, h: box.h }, 'evenodd')))
    if (el.stroke) shapeItems.push(...strokeItems(d, el.stroke))
    if (shapeItems.length) items.push(box.flipH || box.flipV ? { t: 'group', transform: flip, items: shapeItems } : { t: 'group', items: shapeItems })
    if (el.text) items.push(...(await this.text(el.text, box.w, box.h)))
    return [{ t: 'group', transform: boxMatrix(box), items }]
  }

  private async picture(el: Extract<SlideEl, { t: 'pic' }>): Promise<DrawItem[]> {
    const { box } = el
    const image = el.target ? await this.image(el.target) : null
    const items: DrawItem[] = []
    const clip = el.geom === 'ellipse' ? ellipsePath(box.w / 2, box.h / 2, box.w / 2, box.h / 2) : el.geom === 'roundRect' ? roundRectPath(0, 0, box.w, box.h, Math.min(box.w, box.h) * 0.1667) : rectPath(0, 0, box.w, box.h)
    if (image) {
      if (!/^image\/(png|jpeg|gif|svg\+xml)$/.test(image.type)) this.warn(`Pictures in ${image.type.replace(/^image\//, '').toUpperCase()} are not drawn in a PDF yet`)
      let x = 0
      let y = 0
      let w = box.w
      let h = box.h
      if (el.crop) {
        const [l, t, r, b] = el.crop
        w = box.w / Math.max(0.01, 1 - l - r)
        h = box.h / Math.max(0.01, 1 - t - b)
        x = -l * w
        y = -t * h
      }
      let pic: DrawItem = await this.picItem(image, x, y, w, h)
      // the cropped picture flipped in its frame, as PowerPoint flips it (about the frame's centre, not the picture's)
      if (box.flipH || box.flipV) pic = { t: 'group', transform: [box.flipH ? -1 : 1, 0, 0, box.flipV ? -1 : 1, box.flipH ? box.w : 0, box.flipV ? box.h : 0], items: [pic] }
      items.push({ t: 'group', clip, items: [pic] })
    }
    if (el.stroke) {
      // a CSS outline: around the box, outside it
      const s = el.stroke
      const c = parseColor(s.c)
      if (c) items.push({ t: 'path', d: rectPath(-s.w / 2, -s.w / 2, box.w + s.w, box.h + s.w), stroke: { paint: { t: 'color', color: c }, width: s.w } })
    }
    return [{ t: 'group', transform: boxMatrix(box), items }]
  }

  private async table(el: Extract<SlideEl, { t: 'table' }>): Promise<DrawItem[]> {
    const { cols, rows } = el
    const xs = [0]
    for (const w of cols) xs.push(xs[xs.length - 1] + w)
    const right = (ci: number, n: number) => xs[Math.min(cols.length, ci + Math.max(1, n))] ?? xs[xs.length - 1]
    // each cell's text laid out first: rows grow to their text, as a table's rows do in the viewer
    const heights = rows.map((r) => r.h)
    const laid = new Map<TCell, DrawItem[]>()
    for (let ri = 0; ri < rows.length; ri++) {
      for (let ci = 0; ci < rows[ri].cells.length; ci++) {
        const cell = rows[ri].cells[ci]
        const t = cell.text
        if (cell.hidden || !t || t.empty) continue
        const w = right(ci, cell.colSpan) - (xs[ci] ?? 0)
        const out = await layoutParagraphs(this.opts.fonts, t.paras.map((p) => this.para(p)), Math.max(0, w - t.ins[0] - t.ins[2]), { wrap: t.wrap, onMissing: (f) => this.missing(f) })
        laid.set(cell, out.items)
        if (cell.rowSpan <= 1) heights[ri] = Math.max(heights[ri], out.height + t.ins[1] + t.ins[3])
      }
    }
    const ys = [0]
    for (const h of heights) ys.push(ys[ys.length - 1] + h)
    const fills: DrawItem[] = []
    const texts: DrawItem[] = []
    const borders: DrawItem[] = []
    const edge = (s: Stroke | undefined, x0: number, y0: number, x1: number, y1: number) => {
      const c = s && parseColor(s.c)
      if (s && c) borders.push({ t: 'path', d: [{ k: 'M', x: x0, y: y0 }, { k: 'L', x: x1, y: y1 }], stroke: { paint: { t: 'color', color: c }, width: Math.max(1, s.w) } })
    }
    for (let ri = 0; ri < rows.length; ri++) {
      for (let ci = 0; ci < rows[ri].cells.length; ci++) {
        const cell = rows[ri].cells[ci]
        if (cell.hidden) continue
        const x = xs[ci] ?? 0
        const y = ys[ri]
        const w = right(ci, cell.colSpan) - x
        const h = ys[Math.min(rows.length, ri + Math.max(1, cell.rowSpan))] - y
        if (cell.fill) fills.push(...(await this.fill(cell.fill, rectPath(x, y, w, h), { x, y, w, h })))
        const items = laid.get(cell)
        // (cells are anchored at the top in the viewer)
        if (items && cell.text) texts.push({ t: 'group', transform: translate(x + cell.text.ins[0], y + cell.text.ins[1]), items })
        edge(cell.bt, x, y, x + w, y)
        edge(cell.bb, x, y + h, x + w, y + h)
        edge(cell.bl, x, y, x, y + h)
        edge(cell.br, x + w, y, x + w, y + h)
      }
    }
    return [{ t: 'group', transform: boxMatrix(el.box), items: [...fills, ...texts, ...borders] }]
  }

  private missed = new Set<string>()

  private missing(font: string) {
    const name = font.replace(/^"([^"]*)".*$/, '$1').split(',')[0].trim()
    if (this.missed.has(name)) return
    this.missed.add(name)
    this.warn(`No font file for ${name}: its text is set in Helvetica`)
  }

  private para(p: TextBody['paras'][number]): ParaSpec {
    const font = this.opts.defaultFont ?? '"Calibri", Carlito, sans-serif'
    const style = (r: { sz: number; b?: boolean; i?: boolean; u?: boolean; strike?: boolean; color?: string; font?: string; baseline?: number; caps?: boolean; href?: string; hl?: string }): RunStyle => ({
      font: r.font ?? font,
      size: r.sz * PT,
      ...(r.b ? { bold: true } : {}),
      ...(r.i ? { italic: true } : {}),
      color: parseColor(r.color) ?? BLACK,
      ...(r.u ? { underline: true } : {}),
      ...(r.strike ? { strike: true } : {}),
      ...(r.baseline ? { baseline: r.baseline } : {}),
      ...(r.caps ? { caps: true } : {}),
      ...(safeHref(r.href) ? { href: safeHref(r.href)! } : {}),
      ...(parseColor(r.hl) ? { highlight: parseColor(r.hl)! } : {}),
    })
    return {
      runs: p.runs.map((r) => ({ text: r.br ? '' : r.text, style: style(r), ...(r.br ? { br: true } : {}) })),
      align: p.algn,
      marL: p.marL,
      indent: p.indent,
      ...(p.lnPx ? { linePx: p.lnPx } : { lineMult: (p.lnPct ?? 1) * 1.2 }),
      before: p.spcBef,
      after: p.spcAft,
      ...(p.bullet ? { bullet: { text: p.bullet.text, style: { font: p.bullet.font ?? p.runs[0]?.font ?? font, size: p.bullet.sz * PT, color: parseColor(p.bullet.color) ?? parseColor(p.runs[0]?.color) ?? BLACK } } } : {}),
      emptySize: p.endSz * PT,
    }
  }

  /** a shape's text in its box (box-local px): insets, vertical anchor, vertical text turned */
  private async text(t: TextBody, w: number, h: number): Promise<DrawItem[]> {
    if (t.empty) return []
    const [l, top, r, b] = t.ins
    const vertical = t.vert === 'vert' || t.vert === 'eaVert' || t.vert === 'vert270' || t.vert === 'mongolianVert' || t.vert === 'wordArtVertRtl'
    // vertical text is laid out across the box's height, then turned
    const bw = vertical ? h : w
    const bh = vertical ? w : h
    const out = await layoutParagraphs(this.opts.fonts, t.paras.map((p) => this.para(p)), Math.max(0, bw - l - r), { wrap: t.wrap, onMissing: (f) => this.missing(f) })
    const room = bh - top - b
    const dy = t.anchor === 'ctr' ? (room - out.height) / 2 : t.anchor === 'b' ? room - out.height : 0
    let m: Matrix = translate(l, top + dy)
    if (vertical) m = multiply(m, t.vert === 'vert270' ? [0, -1, 1, 0, 0, h] : [0, 1, -1, 0, w, 0])
    return [{ t: 'group', transform: m, items: out.items }]
  }
}

/** what went wrong, in words, whatever was thrown (an object with no way to be a string included) */
function reason(e: unknown): string {
  if (e instanceof Error) return e.message
  try {
    return String(e)
  } catch {
    return 'an error'
  }
}

/** box-local px to slide px: the box's place, turned about its centre */
function boxMatrix(b: Box): Matrix {
  let m: Matrix = IDENTITY
  if (b.rot) m = rotate(b.rot, b.w / 2, b.h / 2)
  return multiply(m, translate(b.x, b.y))
}

/** `linear-gradient(Adeg, c p%, …)` / `radial-gradient(circle, …)` as the engine writes them, over `box` as CSS spans them */
export function gradient(css: string, box: { x: number; y: number; w: number; h: number }): Paint | null {
  const stops: { at: number; color: Rgba }[] = []
  const re = /(#[0-9a-f]{3,8}|rgba?\([^)]*\))\s*(-?[\d.]+)?%?/gi
  const body = css.replace(/^[a-z-]+\(\s*(?:-?[\d.]+deg|circle)\s*,/i, '')
  for (let m = re.exec(body); m; m = re.exec(body)) {
    const color = parseColor(m[1])
    if (color) stops.push({ at: m[2] !== undefined ? Number(m[2]) / 100 : NaN, color })
  }
  if (!stops.length) return null
  // stops without a place are spread between their neighbours, as CSS spreads them
  stops.forEach((s, i) => {
    if (Number.isNaN(s.at)) s.at = stops.length === 1 ? 0 : i / (stops.length - 1)
  })
  const cx = box.x + box.w / 2
  const cy = box.y + box.h / 2
  if (/^radial/i.test(css)) {
    const r = Math.hypot(box.w, box.h) / 2
    return { t: 'radial', cx, cy, r, fx: cx, fy: cy, stops }
  }
  const deg = Number(/(-?[\d.]+)deg/.exec(css)?.[1] ?? 180)
  const a = (deg * Math.PI) / 180
  // CSS: 0deg points up, 90deg right; the line is as long as the box is across in that direction
  const dx = Math.sin(a)
  const dy = -Math.cos(a)
  const half = (Math.abs(box.w * dx) + Math.abs(box.h * dy)) / 2
  return { t: 'linear', x0: cx - dx * half, y0: cy - dy * half, x1: cx + dx * half, y1: cy + dy * half, stops }
}

/** a line's dash pattern as the model writes it ("4 3", px) */
const dashOf = (s: string | undefined) => (s ? s.split(/[\s,]+/).map(Number).filter((v) => Number.isFinite(v) && v >= 0) : undefined)

/** a shape's outline, with its arrowheads */
function strokeItems(d: Seg[], s: Stroke): DrawItem[] {
  const c = parseColor(s.c)
  if (!c) return []
  const stroke: StrokeStyle = { paint: { t: 'color', color: c }, width: s.w, join: 'round' }
  const dash = dashOf(s.dash)
  if (dash?.length && dash.some((v) => v > 0)) stroke.dash = dash
  const items: DrawItem[] = [{ t: 'path', d, stroke }]
  // the viewer's markers: a triangle four line widths long, centred on the end, along the line
  const head = (tip: { x: number; y: number }, from: { x: number; y: number }) => {
    const len = Math.hypot(tip.x - from.x, tip.y - from.y)
    if (!len) return
    const ux = (tip.x - from.x) / len
    const uy = (tip.y - from.y) / len
    const size = 4 * s.w
    // refX 5 of a 10-wide marker: the triangle's middle is on the end
    const bx = tip.x - ux * size * 0.5
    const by = tip.y - uy * size * 0.5
    const pt = { x: bx + ux * size, y: by + uy * size }
    const l = { x: bx - uy * size * 0.5, y: by + ux * size * 0.5 }
    const r = { x: bx + uy * size * 0.5, y: by - ux * size * 0.5 }
    items.push({ t: 'path', d: [{ k: 'M', ...l }, { k: 'L', ...pt }, { k: 'L', ...r }, { k: 'Z' }], fill: { t: 'color', color: c } })
  }
  const pts = d.filter((x): x is Exclude<Seg, { k: 'Z' }> => x.k !== 'Z')
  if (pts.length >= 2) {
    const last = pts[pts.length - 1]
    const before = last.k === 'C' ? { x: last.x2, y: last.y2 } : pts[pts.length - 2]
    const first = pts[0]
    const second = pts[1].k === 'C' ? { x: pts[1].x1, y: pts[1].y1 } : pts[1]
    if (s.tail) head(last, before)
    if (s.head) head(first, second)
  }
  return items
}
