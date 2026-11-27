/**
 * A page of drawing commands as an SVG picture: for tests (compared with the viewer's own rendering) and thumbnails.
 * Text is written as text at the place of each of its characters, in the font's family, so a picture shows the same
 * where the fonts are installed; pictures are embedded as data: URLs.
 */
import { esc } from '../diagram/svg'
import { glyphTexts } from './glyphs'
import type { DrawItem, Page, Paint, Rgba, Seg } from './types'

const n = (v: number) => String(Math.round(v * 100) / 100)

const color = (c: Rgba) => `rgb(${c.r},${c.g},${c.b})`

function pathData(d: Seg[]): string {
  return d
    .map((s) => (s.k === 'Z' ? 'Z' : s.k === 'C' ? `C${n(s.x1)} ${n(s.y1)} ${n(s.x2)} ${n(s.y2)} ${n(s.x)} ${n(s.y)}` : `${s.k}${n(s.x)} ${n(s.y)}`))
    .join('')
}

function base64(b: Uint8Array): string {
  let s = ''
  for (let i = 0; i < b.length; i += 0x8000) s += String.fromCharCode(...b.subarray(i, i + 0x8000))
  return btoa(s)
}

export interface SvgOptions {
  /** embed pictures (default true); false leaves their place empty */
  images?: boolean
}

/** `page` as a self-contained SVG document */
export function pageToSvg(page: Page, opts: SvgOptions = {}): string {
  const defs: string[] = []
  let id = 0
  /** a paint as the value of `attr` (fill, stroke), with its opacity */
  const paint = (p: Paint, attr: 'fill' | 'stroke'): string => {
    if (p.t === 'color') return ` ${attr}="${color(p.color)}"${p.color.a < 1 ? ` ${attr}-opacity="${n(p.color.a)}"` : ''}`
    const key = `g${++id}`
    const stops = p.stops.map((s) => `<stop offset="${n(s.at)}" stop-color="${color(s.color)}"${s.color.a < 1 ? ` stop-opacity="${n(s.color.a)}"` : ''}/>`).join('')
    if (p.t === 'linear') defs.push(`<linearGradient id="${key}" gradientUnits="userSpaceOnUse" x1="${n(p.x0)}" y1="${n(p.y0)}" x2="${n(p.x1)}" y2="${n(p.y1)}">${stops}</linearGradient>`)
    else defs.push(`<radialGradient id="${key}" gradientUnits="userSpaceOnUse" cx="${n(p.cx)}" cy="${n(p.cy)}" r="${n(p.r)}" fx="${n(p.fx)}" fy="${n(p.fy)}">${stops}</radialGradient>`)
    return ` ${attr}="url(#${key})"`
  }
  const item = (it: DrawItem): string => {
    switch (it.t) {
      case 'path': {
        const fill = it.fill ? paint(it.fill, 'fill') : ' fill="none"'
        const s = it.stroke
        let stroke = ''
        if (s && s.width > 0) {
          stroke = `${paint(s.paint, 'stroke')} stroke-width="${n(s.width)}"${s.dash?.length ? ` stroke-dasharray="${s.dash.map(n).join(' ')}"` : ''}${s.cap && s.cap !== 'butt' ? ` stroke-linecap="${s.cap}"` : ''}${s.join && s.join !== 'miter' ? ` stroke-linejoin="${s.join}"` : ''}`
        }
        return `<path d="${pathData(it.d)}"${fill}${it.rule === 'evenodd' ? ' fill-rule="evenodd"' : ''}${stroke}/>`
      }
      case 'text': {
        // each glyph's first character at the glyph's place
        const xs: string[] = []
        let chars = ''
        let x = it.x
        const texts = glyphTexts(it)
        it.glyphs.forEach((_, i) => {
          for (const ch of texts[i]) {
            xs.push(n(x + (it.dx?.[i] ?? 0)))
            chars += ch
          }
          x += it.advances[i]
        })
        const family = typeof it.font === 'string' ? 'Helvetica, Arial, sans-serif' : `'${it.font.names.family.replace(/['\\]/g, '')}'`
        const bold = typeof it.font === 'string' ? it.font.includes('Bold') : it.font.weight >= 600 || it.synthetic?.bold
        const italic = typeof it.font === 'string' ? it.font.includes('Oblique') : it.font.italic || it.synthetic?.italic
        return `<text x="${xs.join(' ')}" y="${n(it.y)}" font-family="${esc(family)}" font-size="${n(it.size)}"${bold ? ' font-weight="bold"' : ''}${italic ? ' font-style="italic"' : ''} fill="${color(it.fill)}"${it.fill.a < 1 ? ` fill-opacity="${n(it.fill.a)}"` : ''} xml:space="preserve">${esc(chars)}</text>`
      }
      case 'image':
        if (opts.images === false) return ''
        return `<image x="${n(it.x)}" y="${n(it.y)}" width="${n(it.w)}" height="${n(it.h)}" preserveAspectRatio="none"${it.opacity !== undefined && it.opacity < 1 ? ` opacity="${n(it.opacity)}"` : ''} href="data:${esc(it.image.type)};base64,${base64(it.image.bytes)}"/>`
      case 'group': {
        let clip = ''
        if (it.clip) {
          const key = `c${++id}`
          defs.push(`<clipPath id="${key}"><path d="${pathData(it.clip)}"${it.clipRule === 'evenodd' ? ' clip-rule="evenodd"' : ''}/></clipPath>`)
          clip = ` clip-path="url(#${key})"`
        }
        // the clip is in the group's coordinates: a transform goes on a group of its own around it
        const inner = `<g${clip}${it.opacity !== undefined && it.opacity < 1 ? ` opacity="${n(it.opacity)}"` : ''}>${it.items.map(item).join('')}</g>`
        return it.transform ? `<g transform="matrix(${it.transform.map(n).join(' ')})">${inner}</g>` : inner
      }
      case 'link':
        return ''
    }
  }
  const body = page.items.map(item).join('')
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${n(page.w)}" height="${n(page.h)}" viewBox="0 0 ${n(page.w)} ${n(page.h)}">${defs.length ? `<defs>${defs.join('')}</defs>` : ''}${body}</svg>`
}
