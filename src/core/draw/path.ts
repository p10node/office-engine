/** SVG path data, transforms and CSS colors as drawing commands take them. */
import type { Matrix, Rgba, Seg } from './types'

export const IDENTITY: Matrix = [1, 0, 0, 1, 0, 0]

/** `m` then `n`: a point is transformed by `m` first */
export function multiply(m: Matrix, n: Matrix): Matrix {
  return [m[0] * n[0] + m[1] * n[2], m[0] * n[1] + m[1] * n[3], m[2] * n[0] + m[3] * n[2], m[2] * n[1] + m[3] * n[3], m[4] * n[0] + m[5] * n[2] + n[4], m[4] * n[1] + m[5] * n[3] + n[5]]
}

export const translate = (x: number, y: number): Matrix => [1, 0, 0, 1, x, y]
export const scale = (sx: number, sy = sx): Matrix => [sx, 0, 0, sy, 0, 0]
/** degrees clockwise (y down), about (cx, cy) */
export function rotate(deg: number, cx = 0, cy = 0): Matrix {
  const r = (deg * Math.PI) / 180
  const c = Math.cos(r)
  const s = Math.sin(r)
  return [c, s, -s, c, cx - c * cx + s * cy, cy - s * cx - c * cy]
}

/** the matrix that undoes `m`; null for one that flattens (a scale of 0) */
export function invert(m: Matrix): Matrix | null {
  const det = m[0] * m[3] - m[1] * m[2]
  if (!det || !Number.isFinite(det)) return null
  return [m[3] / det, -m[1] / det, -m[2] / det, m[0] / det, (m[2] * m[5] - m[3] * m[4]) / det, (m[1] * m[4] - m[0] * m[5]) / det]
}

export const apply = (m: Matrix, x: number, y: number): [number, number] => [m[0] * x + m[2] * y + m[4], m[1] * x + m[3] * y + m[5]]

export function transformPath(d: Seg[], m: Matrix): Seg[] {
  return d.map((s) => {
    if (s.k === 'Z') return s
    if (s.k === 'C') {
      const [x1, y1] = apply(m, s.x1, s.y1)
      const [x2, y2] = apply(m, s.x2, s.y2)
      const [x, y] = apply(m, s.x, s.y)
      return { k: 'C', x1, y1, x2, y2, x, y }
    }
    const [x, y] = apply(m, s.x, s.y)
    return { k: s.k, x, y }
  })
}

export const rectPath = (x: number, y: number, w: number, h: number): Seg[] => [
  { k: 'M', x, y },
  { k: 'L', x: x + w, y },
  { k: 'L', x: x + w, y: y + h },
  { k: 'L', x, y: y + h },
  { k: 'Z' },
]

/** κ: control points of a quarter circle as a cubic curve */
const K = 0.5522847498

export function ellipsePath(cx: number, cy: number, rx: number, ry: number): Seg[] {
  const ox = rx * K
  const oy = ry * K
  return [
    { k: 'M', x: cx + rx, y: cy },
    { k: 'C', x1: cx + rx, y1: cy + oy, x2: cx + ox, y2: cy + ry, x: cx, y: cy + ry },
    { k: 'C', x1: cx - ox, y1: cy + ry, x2: cx - rx, y2: cy + oy, x: cx - rx, y: cy },
    { k: 'C', x1: cx - rx, y1: cy - oy, x2: cx - ox, y2: cy - ry, x: cx, y: cy - ry },
    { k: 'C', x1: cx + ox, y1: cy - ry, x2: cx + rx, y2: cy - oy, x: cx + rx, y: cy },
    { k: 'Z' },
  ]
}

export function roundRectPath(x: number, y: number, w: number, h: number, rx: number, ry = rx): Seg[] {
  rx = Math.min(Math.abs(rx), w / 2)
  ry = Math.min(Math.abs(ry), h / 2)
  if (!rx || !ry) return rectPath(x, y, w, h)
  const ox = rx * (1 - K)
  const oy = ry * (1 - K)
  return [
    { k: 'M', x: x + rx, y },
    { k: 'L', x: x + w - rx, y },
    { k: 'C', x1: x + w - ox, y1: y, x2: x + w, y2: y + oy, x: x + w, y: y + ry },
    { k: 'L', x: x + w, y: y + h - ry },
    { k: 'C', x1: x + w, y1: y + h - oy, x2: x + w - ox, y2: y + h, x: x + w - rx, y: y + h },
    { k: 'L', x: x + rx, y: y + h },
    { k: 'C', x1: x + ox, y1: y + h, x2: x, y2: y + h - oy, x, y: y + h - ry },
    { k: 'L', x, y: y + ry },
    { k: 'C', x1: x, y1: y + oy, x2: x + ox, y2: y, x: x + rx, y },
    { k: 'Z' },
  ]
}

/** the box a path's points (control points included) lie in: [x0, y0, x1, y1], null for none */
export function bounds(d: Seg[]): [number, number, number, number] | null {
  let x0 = Infinity
  let y0 = Infinity
  let x1 = -Infinity
  let y1 = -Infinity
  const at = (x: number, y: number) => {
    if (x < x0) x0 = x
    if (x > x1) x1 = x
    if (y < y0) y0 = y
    if (y > y1) y1 = y
  }
  for (const s of d) {
    if (s.k === 'Z') continue
    if (s.k === 'C') at(s.x1, s.y1), at(s.x2, s.y2)
    at(s.x, s.y)
  }
  return x0 <= x1 ? [x0, y0, x1, y1] : null
}

/** the largest number of path segments read from one path: a crafted file cannot make one take the memory */
export const MAX_SEGMENTS = 200_000

const NUM = /[-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?/y
const SPACE = /[\s,]*/y

/**
 * SVG path data (`M L H V C S Q T A Z`, relative and absolute) as moves, lines and cubic curves. What follows an error
 * is dropped, as SVG renders a path up to its first error.
 */
export function parsePath(d: string): Seg[] {
  const out: Seg[] = []
  let i = 0
  const skip = () => {
    SPACE.lastIndex = i
    SPACE.exec(d)
    i = SPACE.lastIndex
  }
  const num = (): number | null => {
    skip()
    NUM.lastIndex = i
    const m = NUM.exec(d)
    if (!m) return null
    i = NUM.lastIndex
    return Number(m[0])
  }
  const flag = (): number | null => {
    skip()
    const c = d[i]
    if (c !== '0' && c !== '1') return null
    i++
    return c === '1' ? 1 : 0
  }
  let x = 0
  let y = 0
  let sx = 0
  let sy = 0
  // the last control point, for S and T
  let cx = 0
  let cy = 0
  let prev = ''
  let cmd = ''
  while (out.length < MAX_SEGMENTS) {
    skip()
    if (i >= d.length) break
    const c = d[i]
    if (/[MLHVCSQTAZmlhvcsqtaz]/.test(c)) {
      cmd = c
      i++
    } else if (!cmd || cmd === 'Z' || cmd === 'z') break
    const rel = cmd === cmd.toLowerCase()
    const ox = rel ? x : 0
    const oy = rel ? y : 0
    const C = cmd.toUpperCase()
    if (C === 'Z') {
      out.push({ k: 'Z' })
      x = sx
      y = sy
      prev = 'Z'
      continue
    }
    if (C === 'M' || C === 'L' || C === 'T') {
      const a = num()
      const b = a === null ? null : num()
      if (b === null) break
      if (C === 'T') {
        const [qx, qy] = prev === 'Q' || prev === 'T' ? [2 * x - cx, 2 * y - cy] : [x, y]
        quad(out, x, y, qx, qy, ox + a!, oy + b)
        cx = qx
        cy = qy
      } else out.push({ k: C === 'M' ? 'M' : 'L', x: ox + a!, y: oy + b })
      x = ox + a!
      y = oy + b
      if (C === 'M') {
        sx = x
        sy = y
        // numbers after a move are lines
        cmd = rel ? 'l' : 'L'
      }
      prev = C
      continue
    }
    if (C === 'H' || C === 'V') {
      const a = num()
      if (a === null) break
      if (C === 'H') x = (rel ? x : 0) + a
      else y = (rel ? y : 0) + a
      out.push({ k: 'L', x, y })
      prev = C
      continue
    }
    if (C === 'C' || C === 'S') {
      const n: number[] = []
      for (let k = C === 'C' ? 6 : 4; k > 0; k--) {
        const v = num()
        if (v === null) break
        n.push(v)
      }
      if (n.length < (C === 'C' ? 6 : 4)) break
      const [x1, y1] = C === 'C' ? [ox + n[0], oy + n[1]] : prev === 'C' || prev === 'S' ? [2 * x - cx, 2 * y - cy] : [x, y]
      const r = C === 'C' ? 2 : 0
      const x2 = ox + n[r]
      const y2 = oy + n[r + 1]
      x = ox + n[r + 2]
      y = oy + n[r + 3]
      out.push({ k: 'C', x1, y1, x2, y2, x, y })
      cx = x2
      cy = y2
      prev = C
      continue
    }
    if (C === 'Q') {
      const n: number[] = []
      for (let k = 4; k > 0; k--) {
        const v = num()
        if (v === null) break
        n.push(v)
      }
      if (n.length < 4) break
      quad(out, x, y, ox + n[0], oy + n[1], ox + n[2], oy + n[3])
      cx = ox + n[0]
      cy = oy + n[1]
      x = ox + n[2]
      y = oy + n[3]
      prev = 'Q'
      continue
    }
    // A rx ry rotation large-arc sweep x y
    const rx = num()
    const ry = rx === null ? null : num()
    const rot = ry === null ? null : num()
    const large = rot === null ? null : flag()
    const sweep = large === null ? null : flag()
    const ex = sweep === null ? null : num()
    const ey = ex === null ? null : num()
    if (ey === null) break
    arc(out, x, y, rx!, ry!, rot!, !!large, !!sweep, ox + ex!, oy + ey)
    x = ox + ex!
    y = oy + ey
    prev = 'A'
  }
  return out
}

function quad(out: Seg[], x0: number, y0: number, qx: number, qy: number, x: number, y: number) {
  out.push({ k: 'C', x1: x0 + (2 / 3) * (qx - x0), y1: y0 + (2 / 3) * (qy - y0), x2: x + (2 / 3) * (qx - x), y2: y + (2 / 3) * (qy - y), x, y })
}

/** an SVG elliptical arc as cubic curves (SVG 1.1, F.6.5) */
export function arc(out: Seg[], x0: number, y0: number, rx: number, ry: number, deg: number, large: boolean, sweep: boolean, x: number, y: number) {
  rx = Math.abs(rx)
  ry = Math.abs(ry)
  if ((x0 === x && y0 === y) || !rx || !ry) {
    if (x0 !== x || y0 !== y) out.push({ k: 'L', x, y })
    return
  }
  const phi = (deg * Math.PI) / 180
  const cos = Math.cos(phi)
  const sin = Math.sin(phi)
  const dx = (x0 - x) / 2
  const dy = (y0 - y) / 2
  const x1 = cos * dx + sin * dy
  const y1 = -sin * dx + cos * dy
  const lambda = (x1 * x1) / (rx * rx) + (y1 * y1) / (ry * ry)
  if (lambda > 1) {
    rx *= Math.sqrt(lambda)
    ry *= Math.sqrt(lambda)
  }
  const num = rx * rx * ry * ry - rx * rx * y1 * y1 - ry * ry * x1 * x1
  const den = rx * rx * y1 * y1 + ry * ry * x1 * x1
  const co = (large === sweep ? -1 : 1) * Math.sqrt(Math.max(0, num / den))
  const cxp = (co * rx * y1) / ry
  const cyp = (-co * ry * x1) / rx
  const cx = cos * cxp - sin * cyp + (x0 + x) / 2
  const cy = sin * cxp + cos * cyp + (y0 + y) / 2
  const angle = (ux: number, uy: number, vx: number, vy: number) => {
    const a = Math.atan2(ux * vy - uy * vx, ux * vx + uy * vy)
    return a
  }
  const t1 = angle(1, 0, (x1 - cxp) / rx, (y1 - cyp) / ry)
  let dt = angle((x1 - cxp) / rx, (y1 - cyp) / ry, (-x1 - cxp) / rx, (-y1 - cyp) / ry)
  if (!sweep && dt > 0) dt -= 2 * Math.PI
  else if (sweep && dt < 0) dt += 2 * Math.PI
  const n = Math.max(1, Math.ceil(Math.abs(dt) / (Math.PI / 2) - 1e-9))
  const step = dt / n
  const k = (4 / 3) * Math.tan(step / 4)
  let t = t1
  const pt = (a: number): [number, number] => [cx + rx * Math.cos(a) * cos - ry * Math.sin(a) * sin, cy + rx * Math.cos(a) * sin + ry * Math.sin(a) * cos]
  const tangent = (a: number): [number, number] => [-rx * Math.sin(a) * cos - ry * Math.cos(a) * sin, -rx * Math.sin(a) * sin + ry * Math.cos(a) * cos]
  for (let i = 0; i < n; i++) {
    const [ax, ay] = pt(t)
    const [tx0, ty0] = tangent(t)
    const t2 = t + step
    const [bx, by] = i === n - 1 ? [x, y] : pt(t2)
    const [tx1, ty1] = tangent(t2)
    out.push({ k: 'C', x1: ax + k * tx0, y1: ay + k * ty0, x2: bx - k * tx1, y2: by - k * ty1, x: bx, y: by })
    t = t2
  }
}

const NAMED: Record<string, string> = {
  black: '000000', white: 'ffffff', red: 'ff0000', green: '008000', blue: '0000ff', yellow: 'ffff00', gray: '808080', grey: '808080',
  silver: 'c0c0c0', maroon: '800000', purple: '800080', fuchsia: 'ff00ff', lime: '00ff00', olive: '808000', navy: '000080', teal: '008080',
  aqua: '00ffff', orange: 'ffa500', lightgray: 'd3d3d3', lightgrey: 'd3d3d3', darkgray: 'a9a9a9', darkgrey: 'a9a9a9',
}

/** a CSS color (`#rgb`, `#rrggbb`, `#rrggbbaa`, `rgb()` / `rgba()`, a few names) as numbers; null for none, `transparent` or anything else */
export function parseColor(v: string | undefined | null): Rgba | null {
  if (!v) return null
  let s = v.trim().toLowerCase()
  if (NAMED[s]) s = '#' + NAMED[s]
  let m = /^#([0-9a-f]{3,4}|[0-9a-f]{6}|[0-9a-f]{8})$/.exec(s)
  if (m) {
    let h = m[1]
    if (h.length <= 4) h = h.replace(/./g, '$&$&')
    return { r: parseInt(h.slice(0, 2), 16), g: parseInt(h.slice(2, 4), 16), b: parseInt(h.slice(4, 6), 16), a: h.length === 8 ? parseInt(h.slice(6, 8), 16) / 255 : 1 }
  }
  m = /^rgba?\(\s*([-.\d]+)%?[\s,]+([-.\d]+)%?[\s,]+([-.\d]+)%?(?:[\s,/]+([-.\d]+)(%?))?\s*\)$/.exec(s)
  if (m) {
    const c = (n: string) => Math.max(0, Math.min(255, Math.round(Number(n))))
    const a = m[4] === undefined ? 1 : Number(m[4]) / (m[5] ? 100 : 1)
    return { r: c(m[1]), g: c(m[2]), b: c(m[3]), a: Number.isFinite(a) ? Math.max(0, Math.min(1, a)) : 1 }
  }
  return null
}
