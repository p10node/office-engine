/**
 * Pages of drawing commands written as a PDF file: no browser, no dependency. Fonts are embedded as subsets with the
 * text of their glyphs (text can be selected, copied and searched), JPEG pictures are copied as they are, PNG and GIF
 * ones unpacked and packed again with their transparency, links and an outline from the pages' titles are kept.
 * Streams are packed with the runtime's `CompressionStream`.
 */
import type { Font } from '../font/font'
import { isCidKeyed, subsetFont, toUnicodeCMap, type FontSubset } from '../font/subset'
import { decodeGif, decodePng, jpegInfo, sniffImage, type Pixels } from '../draw/image'
import { apply, invert, multiply, IDENTITY } from '../draw/path'
import { glyphTexts } from '../draw/glyphs'
import { isBold, stdChar, stdWidth } from '../draw/stdfont'
import type { DrawItem, ImageData, Matrix, Page, Paint, Rgba, Seg, StdFont, Stop } from '../draw/types'

export interface PdfOptions {
  /** the document's title (the file's metadata) */
  title?: string
  author?: string
  subject?: string
  /** the program that made the document (default "office-engine") */
  producer?: string
  /** a bookmark for each page that has a title (default true) */
  outline?: boolean
  /** told of what could not be drawn as it is (a picture of a format that is not read, a font that may not be embedded) */
  onWarning?: (text: string) => void
}

/** px (1/96 in) to points */
const PT = 0.75

/** a PDF number: decimal, never an exponent (past what a page can use, a coordinate is clamped) */
const n = (v: number): string => {
  if (!Number.isFinite(v)) return '0'
  const r = Math.round(Math.max(-1e9, Math.min(1e9, v)) * 1000) / 1000
  return r === 0 ? '0' : r.toFixed(3).replace(/\.?0+$/, '')
}

/**
 * A number to six significant digits, never an exponent: the entries of transforms and the places of stops, which
 * three decimals would flatten (a scale of 0.00149 written as 0.001)
 */
const nm = (v: number): string => {
  if (!Number.isFinite(v) || Math.abs(v) < 1e-12) return '0'
  const c = Math.max(-1e9, Math.min(1e9, v))
  const s = c.toFixed(Math.max(0, Math.min(15, 5 - Math.floor(Math.log10(Math.abs(c))))))
  return s.includes('.') ? s.replace(/\.?0+$/, '') : s
}

/** a matrix as it is written: what is drawn, and the boxes and patterns worked out from it, use the same numbers */
const written = (m: Matrix): Matrix => m.map((v) => Number(nm(v))) as Matrix

/** a URI action's address: ASCII only, as the format asks (other characters percent-encoded as UTF-8), escaped */
// eslint-disable-next-line no-control-regex
const uriString = (url: string) =>
  '(' +
  url
    // (half a surrogate pair is no character: it cannot be encoded, and stands for U+FFFD as in a browser)
    .replace(/[\ud800-\udbff](?![\udc00-\udfff])|(?<![\ud800-\udbff])[\udc00-\udfff]/g, '\ufffd')
    .replace(/[^\x21-\x7e]/gu, (c) => encodeURIComponent(c))
    .replace(/[\\()]/g, '\\$&') +
  ')'

/** a PDF text string: ASCII as it is (escaped), anything else as UTF-16 with its byte order mark */
function pdfString(s: string): string {
  // eslint-disable-next-line no-control-regex
  if (/^[\x20-\x7e]*$/.test(s)) return '(' + s.replace(/[\\()]/g, '\\$&') + ')'
  let hex = 'FEFF'
  for (let i = 0; i < s.length; i++) hex += s.charCodeAt(i).toString(16).padStart(4, '0')
  return `<${hex}>`
}

const pdfName = (s: string) => '/' + s.replace(/[^A-Za-z0-9+_.-]/g, (c) => '#' + c.charCodeAt(0).toString(16).padStart(2, '0'))

async function deflate(data: Uint8Array): Promise<Uint8Array> {
  const stream = new Blob([data as Uint8Array<ArrayBuffer>]).stream().pipeThrough(new CompressionStream('deflate') as unknown as TransformStream<Uint8Array, Uint8Array>)
  return new Uint8Array(await new Response(stream).arrayBuffer())
}

const latin1 = (s: string): Uint8Array => {
  const b = new Uint8Array(s.length)
  for (let i = 0; i < s.length; i++) b[i] = s.charCodeAt(i)
  return b
}

interface FontUse {
  name: string
  /** embedded whole (the license allows no subset): named without a subset's tag */
  whole?: boolean
  /** glyph → the text it stands for (the first glyph of a cluster carries it) */
  glyphs: Map<number, string>
  /** the subset, once every page's glyphs are known */
  sub?: FontSubset
}

interface StdUse {
  name: string
}

/** Writes `pages` as a PDF file. */
export async function writePdf(pages: Page[], opts: PdfOptions = {}): Promise<Uint8Array<ArrayBuffer>> {
  // (a file of no page is one readers refuse)
  if (!pages.length) throw new Error('A PDF file needs at least one page')
  const w = new PdfWriter(opts)
  return w.write(pages)
}

class PdfWriter {
  private objects: (Uint8Array | undefined)[] = [undefined]
  private fonts = new Map<Font, FontUse>()
  private std = new Map<StdFont, StdUse>()
  /** fonts that may not be embedded: their text is set in Helvetica */
  private refused = new Set<Font>()
  private images = new Map<string | Uint8Array, { name: string; ref: number } | null>()
  private alphas = new Map<string, string>()
  private shadings: { name: string; ref: number }[] = []
  private warned = new Set<string>()

  constructor(private opts: PdfOptions) {}

  private warn(text: string) {
    if (this.warned.has(text)) return
    this.warned.add(text)
    this.opts.onWarning?.(text)
  }

  /** a new object number */
  private alloc(): number {
    this.objects.push(undefined)
    return this.objects.length - 1
  }

  private set(ref: number, body: string | Uint8Array) {
    this.objects[ref] = typeof body === 'string' ? latin1(body) : body
  }

  private async stream(ref: number, dict: string, data: Uint8Array, compress = true) {
    const body = compress ? await deflate(data) : data
    const head = latin1(`<<${dict}${compress ? '/Filter/FlateDecode' : ''}/Length ${body.length}>>\nstream\n`)
    const tail = latin1('\nendstream')
    const out = new Uint8Array(head.length + body.length + tail.length)
    out.set(head)
    out.set(body, head.length)
    out.set(tail, head.length + body.length)
    this.objects[ref] = out
  }

  async write(pages: Page[]): Promise<Uint8Array<ArrayBuffer>> {
    const catalog = this.alloc()
    const tree = this.alloc()
    // the glyphs every font draws, before any page is written: pages name the subsets' glyphs
    for (const p of pages) this.collect(p.items)
    for (const [font, use] of this.fonts) {
      if (this.refused.has(font)) continue
      try {
        if (font.fsType & 0x0100) {
          // the license allows the whole font only: its glyphs as they are
          if (isCidKeyed(font)) throw new Error('it may not be subset, and its glyphs are not its characters')
          use.sub = { bytes: font.sfnt(), glyphs: new Map([0, ...use.glyphs.keys()].map((g) => [g, g])), outlines: font.outlines }
          use.whole = true
        } else use.sub = subsetFont(font, use.glyphs.keys())
      } catch (e) {
        // a damaged font (or one that cannot be embedded whole) takes its text to Helvetica, not the file down
        this.refused.add(font)
        this.warn(`The font ${font.names.family} could not be embedded (${(e as Error).message}): its text is set in Helvetica`)
      }
    }
    const kids: number[] = []
    const outline: { title: string; page: number; h: number }[] = []
    for (const page of pages) {
      const ref = this.alloc()
      kids.push(ref)
      const content = this.alloc()
      const W = page.w * PT
      const H = page.h * PT
      const ctx: Ctx = { ops: [], m: [PT, 0, 0, -PT, 0, H], base: IDENTITY, page: [W, H], links: [], res: new Set() }
      ctx.ops.push(`${n(PT)} 0 0 ${n(-PT)} 0 ${n(H)} cm`)
      await this.items(page.items, ctx)
      await this.stream(content, '', latin1(ctx.ops.join('\n')))
      const annots: number[] = []
      for (const l of ctx.links) {
        const a = this.alloc()
        this.set(a, `<</Type/Annot/Subtype/Link/Rect[${l.rect.map(n).join(' ')}]/Border[0 0 0]/A<</S/URI/URI${uriString(l.url)}>>>>`)
        annots.push(a)
      }
      this.set(ref, `<</Type/Page/Parent ${tree} 0 R/MediaBox[0 0 ${n(W)} ${n(H)}]/Resources ${this.resources(ctx.res)}/Contents ${content} 0 R${annots.length ? `/Annots[${annots.map((a) => `${a} 0 R`).join(' ')}]` : ''}>>`)
      if (page.title) outline.push({ title: page.title, page: ref, h: H })
    }
    this.set(tree, `<</Type/Pages/Kids[${kids.map((k) => `${k} 0 R`).join(' ')}]/Count ${kids.length}>>`)
    await this.writeFonts()
    let outlines = ''
    if (this.opts.outline !== false && outline.length) {
      const root = this.alloc()
      const refs = outline.map(() => this.alloc())
      outline.forEach((o, i) => {
        const links = `${i ? `/Prev ${refs[i - 1]} 0 R` : ''}${i < refs.length - 1 ? `/Next ${refs[i + 1]} 0 R` : ''}`
        this.set(refs[i], `<</Title ${pdfString(o.title)}/Parent ${root} 0 R${links}/Dest[${o.page} 0 R/XYZ 0 ${n(o.h)} 0]>>`)
      })
      this.set(root, `<</Type/Outlines/First ${refs[0]} 0 R/Last ${refs[refs.length - 1]} 0 R/Count ${refs.length}>>`)
      outlines = `/Outlines ${root} 0 R/PageMode/UseOutlines`
    }
    this.set(catalog, `<</Type/Catalog/Pages ${tree} 0 R${outlines}>>`)
    const info = this.alloc()
    const d = new Date()
    const date = `D:${d.getUTCFullYear()}${[d.getUTCMonth() + 1, d.getUTCDate(), d.getUTCHours(), d.getUTCMinutes(), d.getUTCSeconds()].map((v) => String(v).padStart(2, '0')).join('')}Z`
    const meta: [string, string | undefined][] = [['Title', this.opts.title], ['Author', this.opts.author], ['Subject', this.opts.subject], ['Producer', this.opts.producer ?? 'office-engine']]
    this.set(info, `<<${meta.filter(([, v]) => v).map(([k, v]) => `/${k} ${pdfString(v!)}`).join('')}/CreationDate(${date})>>`)
    return this.file(catalog, info)
  }

  private file(catalog: number, info: number): Uint8Array<ArrayBuffer> {
    const head = latin1('%PDF-1.7\n%\xe2\xe3\xcf\xd3\n')
    const chunks: Uint8Array[] = [head]
    const offsets: number[] = [0]
    let at = head.length
    for (let i = 1; i < this.objects.length; i++) {
      const body = this.objects[i] ?? latin1('null')
      const open = latin1(`${i} 0 obj\n`)
      const close = latin1('\nendobj\n')
      offsets.push(at)
      chunks.push(open, body, close)
      at += open.length + body.length + close.length
    }
    let xref = `xref\n0 ${this.objects.length}\n0000000000 65535 f \n`
    for (let i = 1; i < offsets.length; i++) xref += `${String(offsets[i]).padStart(10, '0')} 00000 n \n`
    xref += `trailer\n<</Size ${this.objects.length}/Root ${catalog} 0 R/Info ${info} 0 R>>\nstartxref\n${at}\n%%EOF\n`
    chunks.push(latin1(xref))
    let size = 0
    for (const c of chunks) size += c.length
    const out = new Uint8Array(size)
    size = 0
    for (const c of chunks) out.set(c, size), (size += c.length)
    return out
  }

  // ---- what the pages use -----------------------------------------------------

  private collect(items: DrawItem[]) {
    for (const it of items) {
      if (it.t === 'group') this.collect(it.items)
      else if (it.t === 'text' && typeof it.font !== 'string') {
        const font = it.font
        if (!this.refused.has(font) && (!font.embeddable || font.fsType & 0x0200)) {
          this.refused.add(font)
          this.warn(`The font ${font.names.family} may not be embedded${font.embeddable ? ' but as bitmaps' : ''}: its text is set in Helvetica`)
        }
        if (this.refused.has(font)) continue
        let use = this.fonts.get(font)
        if (!use) this.fonts.set(font, (use = { name: `F${this.fonts.size + this.std.size + 1}`, glyphs: new Map() }))
        // the first glyph of a cluster stands for its characters (several for a ligature), the others for none
        const texts = glyphTexts(it)
        it.glyphs.forEach((g, i) => {
          if (!use.glyphs.get(g)) use.glyphs.set(g, texts[i])
        })
      }
    }
  }

  private resources(res: Set<string>): string {
    const fonts: string[] = []
    const xobjects: string[] = []
    const states: string[] = []
    const shadings: string[] = []
    const patterns: string[] = []
    for (const r of res) {
      const [kind, name, ref] = r.split(' ')
      if (kind === 'F') fonts.push(`/${name} ${ref} 0 R`)
      else if (kind === 'X') xobjects.push(`/${name} ${ref} 0 R`)
      else if (kind === 'G') states.push(`/${name} ${ref} 0 R`)
      else if (kind === 'S') shadings.push(`/${name} ${ref} 0 R`)
      else if (kind === 'P') patterns.push(`/${name} ${ref} 0 R`)
    }
    return `<<${fonts.length ? `/Font<<${fonts.join('')}>>` : ''}${xobjects.length ? `/XObject<<${xobjects.join('')}>>` : ''}${states.length ? `/ExtGState<<${states.join('')}>>` : ''}${shadings.length ? `/Shading<<${shadings.join('')}>>` : ''}${patterns.length ? `/Pattern<<${patterns.join('')}>>` : ''}>>`
  }

  private fontRefs = new Map<string, number>()
  private tags = new Set<string>()
  private baseNames = new Set<string>()
  private forms = 0

  private fontRef(name: string): number {
    let ref = this.fontRefs.get(name)
    if (ref === undefined) this.fontRefs.set(name, (ref = this.alloc()))
    return ref
  }

  private async writeFonts() {
    for (const [font, use] of this.fonts) {
      const sub = use.sub
      if (this.refused.has(font) || !sub) continue
      const ref = this.fontRef(use.name)
      const upm = font.unitsPerEm
      const k = 1000 / upm
      // a tag of six capitals from the font and the glyphs kept, as the subsets of one font in two documents differ;
      // two subsets of one document never share one
      let hash = 0
      for (const c of font.names.postScript + font.names.family) hash = (hash * 31 + c.charCodeAt(0)) >>> 0
      for (const g of [...use.glyphs.keys()].sort((a, b) => a - b)) hash = (hash * 31 + g) >>> 0
      let tag = ''
      for (let tries = 0; !tag || this.tags.has(tag); tries++) {
        tag = ''
        let h = (hash + tries * 7919) >>> 0
        for (let i = 0; i < 6; i++) (tag += String.fromCharCode(65 + (h % 26))), (h = Math.floor(h / 26))
      }
      this.tags.add(tag)
      // (the name as it is, escaped: two fonts whose names differ only in punctuation keep apart)
      const ps = (font.names.postScript || font.names.family || 'Font').replace(/[^\x21-\x7e]/g, '')
      let base = use.whole ? ps : `${tag}+${ps}`
      // (two fonts of one name, two files of a family that say the same: kept apart, as a reader may share fonts by name)
      for (let k = 2; this.baseNames.has(base); k++) base = `${use.whole ? ps : `${tag}+${ps}`}-${k}`
      this.baseNames.add(base)
      const m = font.metrics
      const flags = 4 | (font.fixedPitch ? 1 : 0) | (font.italic ? 64 : 0)
      const file = this.alloc()
      const cff = font.outlines === 'cff'
      await this.stream(file, cff ? '/Subtype/OpenType' : `/Length1 ${sub.bytes.length}`, sub.bytes)
      const desc = this.alloc()
      this.set(
        desc,
        `<</Type/FontDescriptor/FontName${pdfName(base)}/Flags ${flags}/FontBBox[${m.bbox.map((v) => n(v * k)).join(' ')}]/ItalicAngle ${n(m.italicAngle)}/Ascent ${n(m.ascent * k)}/Descent ${n(-m.descent * k)}/CapHeight ${n((m.capHeight || m.ascent) * k)}/StemV ${font.weight >= 600 ? 120 : 80}/${cff ? 'FontFile3' : 'FontFile2'} ${file} 0 R>>`,
      )
      // widths of the subset's glyphs, in order
      const widths: [number, number][] = [...sub.glyphs].map(([old, now]) => [now, font.advanceOf(old) * k] as [number, number]).sort((a, b) => a[0] - b[0])
      const W: string[] = []
      for (let i = 0; i < widths.length; ) {
        let j = i
        while (j + 1 < widths.length && widths[j + 1][0] === widths[j][0] + 1) j++
        W.push(`${widths[i][0]}[${widths.slice(i, j + 1).map(([, v]) => n(v)).join(' ')}]`)
        i = j + 1
      }
      const cid = this.alloc()
      this.set(
        cid,
        `<</Type/Font/Subtype/${cff ? 'CIDFontType0' : 'CIDFontType2'}/BaseFont${pdfName(base)}/CIDSystemInfo<</Registry(Adobe)/Ordering(Identity)/Supplement 0>>/FontDescriptor ${desc} 0 R/DW ${n(font.advanceOf(0) * k)}/W[${W.join('')}]${cff ? '' : '/CIDToGIDMap/Identity'}>>`,
      )
      const toUni = this.alloc()
      await this.stream(toUni, '', latin1(unescape(encodeURIComponent(toUnicodeCMap([...use.glyphs].map(([g, text]) => [sub.glyphs.get(g)!, text] as [number, string]))))))
      this.set(ref, `<</Type/Font/Subtype/Type0/BaseFont${pdfName(base)}/Encoding/Identity-H/DescendantFonts[${cid} 0 R]/ToUnicode ${toUni} 0 R>>`)
    }
    for (const [font, use] of this.std) {
      const ref = this.fontRef(use.name)
      const bold = isBold(font)
      const widths: string[] = []
      for (let c = 32; c <= 255; c++) widths.push(String(stdWidth(c, bold)))
      this.set(ref, `<</Type/Font/Subtype/Type1/BaseFont/${font}/Encoding/WinAnsiEncoding/FirstChar 32/LastChar 255/Widths[${widths.join(' ')}]>>`)
    }
  }

  // ---- page content -------------------------------------------------------------

  private async items(items: DrawItem[], ctx: Ctx) {
    for (const it of items) await this.item(it, ctx)
  }

  private async item(it: DrawItem, ctx: Ctx) {
    switch (it.t) {
      case 'path':
        return this.path(it, ctx)
      case 'text':
        return this.text(it, ctx)
      case 'image':
        return this.image(it, ctx)
      case 'link': {
        const pts = [apply(ctx.m, it.x, it.y), apply(ctx.m, it.x + it.w, it.y), apply(ctx.m, it.x, it.y + it.h), apply(ctx.m, it.x + it.w, it.y + it.h)]
        const xs = pts.map((p) => p[0])
        const ys = pts.map((p) => p[1])
        ctx.links.push({ rect: [Math.min(...xs), Math.min(...ys), Math.max(...xs), Math.max(...ys)], url: it.url })
        return
      }
      case 'group': {
        if (!it.items.length) return
        if (it.opacity === undefined || it.opacity >= 1) {
          ctx.ops.push('q')
          await this.groupBody(it, ctx)
          ctx.ops.push('Q')
          return
        }
        // a transparent group is drawn whole, then made transparent (an isolated transparency group): its parts do not
        // show through one another, as in the viewer
        const inner: Ctx = { ops: [], m: ctx.m, base: ctx.m, page: ctx.page, links: ctx.links, res: new Set() }
        await this.groupBody(it, inner)
        const form = await this.form(inner, '/Group<</S/Transparency/I true/CS/DeviceRGB>>')
        if (!form) return
        const { name, ref } = form
        ctx.res.add(`X ${name} ${ref}`)
        ctx.ops.push('q')
        this.alpha(ctx, it.opacity, it.opacity)
        ctx.ops.push(`/${name} Do`, 'Q')
        return
      }
    }
  }

  /** a group's transform, clip and items, into `ctx` */
  private async groupBody(it: Extract<DrawItem, { t: 'group' }>, ctx: Ctx) {
    const saved = ctx.m
    if (it.transform) {
      const t = written(it.transform)
      ctx.ops.push(`${t.map(nm).join(' ')} cm`)
      ctx.m = multiply(t, ctx.m)
    }
    if (it.clip) ctx.ops.push((pathOps(it.clip) || '0 0 m') + (it.clipRule === 'evenodd' ? ' W* n' : ' W n'))
    await this.items(it.items, ctx)
    ctx.m = saved
  }

  /**
   * What was drawn into `inner` as a form XObject, its box the page in the form's space (what is off the page is not
   * seen anyway); null when that space is flattened to nothing
   */
  private async form(inner: Ctx, extra: string): Promise<{ name: string; ref: number } | null> {
    const back = invert(inner.base)
    if (!back) return null
    const [W, H] = inner.page
    const pts = [apply(back, 0, 0), apply(back, W, 0), apply(back, 0, H), apply(back, W, H)]
    const xs = pts.map((p) => p[0])
    const ys = pts.map((p) => p[1])
    const box = [Math.min(...xs) - 1, Math.min(...ys) - 1, Math.max(...xs) + 1, Math.max(...ys) + 1]
    const ref = this.alloc()
    const name = `Fm${++this.forms}`
    await this.stream(ref, `/Type/XObject/Subtype/Form/BBox[${box.map(n).join(' ')}]${extra}/Resources ${this.resources(inner.res)}`, latin1(inner.ops.join('\n')))
    return { name, ref }
  }

  /** an ExtGState for these alphas, made current */
  private alpha(ctx: Ctx, fill: number, stroke: number) {
    const key = `${n(fill)} ${n(stroke)}`
    let name = this.alphas.get(key)
    let ref: number
    if (!name) {
      name = `G${this.alphas.size + 1}`
      ref = this.alloc()
      this.set(ref, `<</Type/ExtGState/ca ${n(fill)}/CA ${n(stroke)}>>`)
      this.alphas.set(key, name)
      this.alphaRefs.set(name, ref)
    } else ref = this.alphaRefs.get(name)!
    ctx.res.add(`G ${name} ${ref}`)
    ctx.ops.push(`/${name} gs`)
  }

  private alphaRefs = new Map<string, number>()

  /** a shading object of the gradient's colors, or (`alpha`) of its stops' alphas as gray */
  private shadingObject(p: Exclude<Paint, { t: 'color' }>, alpha: boolean): number {
    const ref = this.alloc()
    const coords = p.t === 'linear' ? [p.x0, p.y0, p.x1, p.y1] : [p.fx, p.fy, 0, p.cx, p.cy, p.r]
    const fn = stopsFunction(p.stops, alpha ? (c) => n(c.a) : rgb)
    this.set(ref, `<</ShadingType ${p.t === 'linear' ? 2 : 3}/ColorSpace/${alpha ? 'DeviceGray' : 'DeviceRGB'}/Coords[${coords.map(n).join(' ')}]/Function ${fn}/Extend[true true]>>`)
    return ref
  }

  /** the gradient as a shading of the content stream (`sh`) */
  private shading(ctx: Ctx, p: Exclude<Paint, { t: 'color' }>, alpha = false): string {
    const ref = this.shadingObject(p, alpha)
    const name = `Sh${this.shadings.length + 1}`
    this.shadings.push({ name, ref })
    ctx.res.add(`S ${name} ${ref}`)
    return name
  }

  /** the gradient as a shading pattern, to stroke with: placed in the current space, as the shading of `sh` is */
  private pattern(ctx: Ctx, p: Exclude<Paint, { t: 'color' }>): string | null {
    const back = invert(ctx.base)
    if (!back) return null
    const ref = this.alloc()
    this.set(ref, `<</PatternType 2/Shading ${this.shadingObject(p, false)} 0 R/Matrix[${multiply(ctx.m, back).map(nm).join(' ')}]>>`)
    const name = `P${++this.patterns}`
    ctx.res.add(`P ${name} ${ref}`)
    return name
  }

  private patterns = 0

  /** the gradient's alphas as a soft mask, made current: a mask of the same shading in gray */
  private async gradientMask(ctx: Ctx, p: Exclude<Paint, { t: 'color' }>, clip: string) {
    const inner: Ctx = { ops: [], m: ctx.m, base: ctx.m, page: ctx.page, links: [], res: new Set() }
    const sh = this.shading(inner, p, true)
    if (clip) inner.ops.push(clip)
    inner.ops.push(`/${sh} sh`)
    const form = await this.form(inner, '/Group<</S/Transparency/CS/DeviceGray>>')
    if (!form) return
    const gs = this.alloc()
    this.set(gs, `<</Type/ExtGState/SMask<</Type/Mask/S/Luminosity/G ${form.ref} 0 R>>>>`)
    // (a name of their own: the alphas' are G1, G2…)
    const name = `M${++this.masks}`
    ctx.res.add(`G ${name} ${gs}`)
    ctx.ops.push(`/${name} gs`)
  }

  private masks = 0

  private async path(it: Extract<DrawItem, { t: 'path' }>, ctx: Ctx) {
    if (!it.d.length) return
    const ops = pathOps(it.d)
    const evenodd = it.rule === 'evenodd'
    const fill = it.fill && settled(it.fill)
    const stroke = it.stroke && it.stroke.width > 0 ? { ...it.stroke, paint: settled(it.stroke.paint) } : undefined
    if (fill && fill.t !== 'color') {
      // a gradient: the path clips, the shading paints
      const name = this.shading(ctx, fill)
      const clip = ops + (evenodd ? ' W* n' : ' W n')
      ctx.ops.push('q', clip)
      const alphas = fill.stops.map((s) => s.color.a)
      if (alphas.some((a) => a !== alphas[0])) await this.gradientMask(ctx, fill, clip)
      else if (alphas[0] < 1) this.alpha(ctx, alphas[0], 1)
      ctx.ops.push(`/${name} sh`, 'Q')
    }
    const solid = fill?.t === 'color' && fill.color.a > 0 ? fill.color : undefined
    const sp = stroke && !(stroke.paint.t === 'color' && stroke.paint.color.a <= 0) ? stroke.paint : undefined
    // a color fill and a color outline are painted together (B): where they overlap, the outline is not over the fill
    const together = solid && sp?.t === 'color'
    if (solid && !together) {
      ctx.ops.push('q')
      if (solid.a < 1) this.alpha(ctx, solid.a, 1)
      ctx.ops.push(`${rgb(solid)} rg`, ops + (evenodd ? ' f*' : ' f'), 'Q')
    }
    if (!stroke || !sp) return
    ctx.ops.push('q')
    if (sp.t === 'color') {
      if (sp.color.a < 1 || (together && solid.a < 1)) this.alpha(ctx, together ? solid.a : 1, sp.color.a)
      if (together) ctx.ops.push(`${rgb(solid)} rg`)
      ctx.ops.push(`${rgb(sp.color)} RG`)
    } else {
      // a gradient along the line: a shading pattern, its alphas a soft mask where they change
      const name = this.pattern(ctx, sp)
      if (!name) return void ctx.ops.push('Q')
      const alphas = sp.stops.map((s) => s.color.a)
      if (alphas.some((a) => a !== alphas[0])) await this.gradientMask(ctx, sp, '')
      else if (alphas[0] < 1) this.alpha(ctx, 1, alphas[0])
      ctx.ops.push(`/Pattern CS /${name} SCN`)
    }
    ctx.ops.push(`${n(stroke.width)} w`)
    if (stroke.dash?.length) ctx.ops.push(`[${stroke.dash.map(n).join(' ')}] 0 d`)
    ctx.ops.push(`${{ butt: 0, round: 1, square: 2 }[stroke.cap ?? 'butt']} J`, `${{ miter: 0, round: 1, bevel: 2 }[stroke.join ?? 'miter']} j`, ops + (together ? (evenodd ? ' B*' : ' B') : ' S'), 'Q')
  }

  private text(it: Extract<DrawItem, { t: 'text' }>, ctx: Ctx) {
    if (!it.glyphs.length || it.fill.a <= 0) return
    if (typeof it.font !== 'string' && this.refused.has(it.font)) return this.stdText(asStd(it), ctx)
    if (typeof it.font === 'string') return this.stdText(it as StdText, ctx)
    const use = this.fonts.get(it.font)!
    const ref = this.fontRef(use.name)
    ctx.res.add(`F ${use.name} ${ref}`)
    const font = it.font
    const k = it.size / font.unitsPerEm
    // a glyph stands for one text in the whole file: where this text has it stand for another (a ligature of the font
    // and the one character U+FB01), the text is given as it is
    const texts = glyphTexts(it)
    const actual = it.glyphs.some((g, i) => use.glyphs.get(g) !== texts[i])
    if (actual) ctx.ops.push(`/Span<</ActualText ${pdfString(it.text)}>>BDC`)
    this.glyphText(it, ctx, use.name, (g) => (use.sub!.glyphs.get(g) ?? 0).toString(16).padStart(4, '0'), (g) => font.advanceOf(g) * k)
    if (actual) ctx.ops.push('EMC')
  }

  private stdText(it: StdText, ctx: Ctx) {
    let use = this.std.get(it.font)
    if (!use) this.std.set(it.font, (use = { name: `F${this.fonts.size + this.std.size + 1}` }))
    ctx.res.add(`F ${use.name} ${this.fontRef(use.name)}`)
    const bold = isBold(it.font)
    this.glyphText(it, ctx, use.name, (g) => g.toString(16).padStart(2, '0'), (g) => (stdWidth(g, bold) * it.size) / 1000)
  }

  /** the glyphs as one TJ, each moved from where the font's width puts it to where the layout did */
  private glyphText(it: Extract<DrawItem, { t: 'text' }>, ctx: Ctx, font: string, code: (g: number) => string, width: (g: number) => number) {
    const ops = ctx.ops
    ops.push('q')
    if (it.fill.a < 1) this.alpha(ctx, it.fill.a, it.fill.a)
    ops.push(`${rgb(it.fill)} rg`)
    const syn = it.synthetic
    if (syn?.bold) ops.push(`${rgb(it.fill)} RG`, `${n(it.size * 0.03)} w`, '2 Tr')
    const skew = syn?.italic ? 0.21 : 0
    ops.push('BT', `/${font} ${n(it.size)} Tf`)
    const placed = it.dx?.some((v) => v) || it.dy?.some((v) => v)
    if (placed) {
      // glyphs moved off the pen: each at its own place
      let x = it.x
      it.glyphs.forEach((g, i) => {
        ops.push(`1 0 ${n(skew)} -1 ${n(x + (it.dx?.[i] ?? 0))} ${n(it.y - (it.dy?.[i] ?? 0))} Tm <${code(g)}> Tj`)
        x += it.advances[i]
      })
    } else {
      ops.push(`1 0 ${n(skew)} -1 ${n(it.x)} ${n(it.y)} Tm`)
      let tj = ''
      let run = ''
      it.glyphs.forEach((g, i) => {
        run += code(g)
        const off = it.advances[i] - width(g)
        if (Math.abs(off) > 1e-3 && i < it.glyphs.length - 1) {
          tj += `<${run}>${n((-off * 1000) / it.size)}`
          run = ''
        }
      })
      if (run) tj += `<${run}>`
      ops.push(`[${tj}] TJ`)
    }
    ops.push('ET', 'Q')
  }

  private async image(it: Extract<DrawItem, { t: 'image' }>, ctx: Ctx) {
    if (it.w <= 0 || it.h <= 0) return
    const xo = await this.xobject(it.image)
    if (!xo) return
    ctx.res.add(`X ${xo.name} ${xo.ref}`)
    ctx.ops.push('q')
    if (it.opacity !== undefined && it.opacity < 1) this.alpha(ctx, it.opacity, it.opacity)
    ctx.ops.push(`${nm(it.w)} 0 0 ${nm(-it.h)} ${nm(it.x)} ${nm(it.y + it.h)} cm`, `/${xo.name} Do`, 'Q')
  }

  private async xobject(img: ImageData): Promise<{ name: string; ref: number } | null> {
    const key = img.key ?? img.bytes
    const known = this.images.get(key)
    if (known !== undefined) return known
    let out: { name: string; ref: number } | null = null
    const kind = sniffImage(img.bytes)
    try {
      if (kind === 'jpeg') {
        const info = jpegInfo(img.bytes)
        if (info) {
          const ref = this.alloc()
          const cs = info.components === 1 ? 'DeviceGray' : info.components === 4 ? 'DeviceCMYK' : 'DeviceRGB'
          const decode = info.components === 4 && info.adobe ? '/Decode[1 0 1 0 1 0 1 0]' : ''
          await this.stream(ref, `/Type/XObject/Subtype/Image/Width ${info.w}/Height ${info.h}/ColorSpace/${cs}/BitsPerComponent 8${decode}/Filter/DCTDecode`, img.bytes, false)
          // (DCTDecode named in the dictionary: the bytes are the JPEG as it is)
          out = { name: `Im${this.images.size + 1}`, ref }
        }
      } else if (kind === 'png' || kind === 'gif') {
        const px = kind === 'png' ? await decodePng(img.bytes) : decodeGif(img.bytes)
        out = { name: `Im${this.images.size + 1}`, ref: await this.pixels(px) }
      }
    } catch (e) {
      this.warn(`A picture could not be read: ${(e as Error).message}`)
    }
    if (!out && kind !== 'png' && kind !== 'jpeg' && kind !== 'gif') this.warn(`Pictures in ${(kind ?? 'an unknown format').toUpperCase()} are not drawn in a PDF yet`)
    this.images.set(key, out)
    return out
  }

  private async pixels(px: Pixels): Promise<number> {
    const ref = this.alloc()
    let smask = ''
    if (px.alpha) {
      const s = this.alloc()
      await this.stream(s, `/Type/XObject/Subtype/Image/Width ${px.w}/Height ${px.h}/ColorSpace/DeviceGray/BitsPerComponent 8`, px.alpha)
      smask = `/SMask ${s} 0 R`
    }
    await this.stream(ref, `/Type/XObject/Subtype/Image/Width ${px.w}/Height ${px.h}/ColorSpace/${px.channels === 1 ? 'DeviceGray' : 'DeviceRGB'}/BitsPerComponent 8${smask}`, px.color)
    return ref
  }
}

type StdText = Extract<DrawItem, { t: 'text' }> & { font: StdFont }

interface Ctx {
  ops: string[]
  /** user space to PDF space, for links */
  m: Matrix
  /** the content stream's own space to PDF space: the page's is the identity, a form's the space it is drawn in */
  base: Matrix
  /** the page's size in points: what a form's box must hold */
  page: [number, number]
  links: { rect: number[]; url: string }[]
  /** resources the page uses: "F name ref", "X name ref"… */
  res: Set<string>
}

/** text of a font that may not be embedded, set in Helvetica: each cluster's characters where its glyphs were */
function asStd(it: Extract<DrawItem, { t: 'text' }>): StdText {
  const font = it.font as Font
  const std = font.weight >= 600 ? (font.italic ? 'Helvetica-BoldOblique' : 'Helvetica-Bold') : font.italic ? 'Helvetica-Oblique' : 'Helvetica'
  const bold = isBold(std)
  const glyphs: number[] = []
  const advances: number[] = []
  const clusters: number[] = []
  const texts = glyphTexts(it)
  // each cluster as wide as its glyphs together
  const widths = new Map<number, number>()
  it.clusters.forEach((c, k) => widths.set(c, (widths.get(c) ?? 0) + it.advances[k]))
  it.glyphs.forEach((_, i) => {
    const from = it.clusters[i] ?? 0
    if (!texts[i]) return
    // its last character takes up what its others do not; a tab is pen movement: a space, whose width is the layout's
    let width = widths.get(from) ?? 0
    const chars = [...texts[i].replace(/\t/g, ' ')]
    chars.forEach((ch, k) => {
      const code = stdCodeOf(ch)
      const w = k < chars.length - 1 ? (stdWidth(code, bold) * it.size) / 1000 : width
      glyphs.push(code)
      advances.push(w)
      clusters.push(from)
      width -= w
    })
  })
  return { ...it, font: std, glyphs, advances, clusters, dx: undefined, dy: undefined, synthetic: undefined }
}

function stdCodeOf(ch: string): number {
  const cp = ch.codePointAt(0)!
  for (let c = 32; c < 256; c++) if (stdChar(c).codePointAt(0) === cp) return c
  return 0x3f
}

function pathOps(d: Seg[]): string {
  const out: string[] = []
  for (const s of d) {
    if (s.k === 'M') out.push(`${n(s.x)} ${n(s.y)} m`)
    else if (s.k === 'L') out.push(`${n(s.x)} ${n(s.y)} l`)
    else if (s.k === 'C') out.push(`${n(s.x1)} ${n(s.y1)} ${n(s.x2)} ${n(s.y2)} ${n(s.x)} ${n(s.y)} c`)
    else out.push('h')
  }
  return out.join(' ')
}

const rgb = (c: Rgba) => `${n(c.r / 255)} ${n(c.g / 255)} ${n(c.b / 255)}`

/** a gradient of no length (its two points at one place, a circle of no radius) paints its last stop's color, as SVG says */
function settled(p: Paint): Paint {
  if (p.t === 'color') return p
  const flat = p.t === 'linear' ? p.x0 === p.x1 && p.y0 === p.y1 : !(p.r > 0)
  const last = [...p.stops].sort((a, b) => a.at - b.at).pop()
  return flat || !last ? { t: 'color', color: last?.color ?? { r: 0, g: 0, b: 0, a: 0 } } : p
}

/**
 * A function from 0-1 to the stops' colors (`color` writes one): an interpolation between two, a stitching of them for
 * more. Two stops at one place make a hard edge: the interval between them, of no width, is left out.
 */
function stopsFunction(stops: Stop[], color: (c: Rgba) => string): string {
  // (places as they are written: stops that round to one place make one hard edge, not an interval of no width)
  const s = (stops.length ? [...stops] : [{ at: 0, color: { r: 0, g: 0, b: 0, a: 1 } }]).map((x) => ({ ...x, at: Number(nm(Math.min(1, Math.max(0, x.at)))) })).sort((a, b) => a.at - b.at)
  if (s[0].at > 0) s.unshift({ at: 0, color: s[0].color })
  if (s[s.length - 1].at < 1) s.push({ at: 1, color: s[s.length - 1].color })
  const lerp = (a: Rgba, b: Rgba) => `<</FunctionType 2/Domain[0 1]/C0[${color(a)}]/C1[${color(b)}]/N 1>>`
  const parts: { from: number; fn: string }[] = []
  for (let i = 0; i < s.length - 1; i++) if (s[i + 1].at > s[i].at) parts.push({ from: s[i].at, fn: lerp(s[i].color, s[i + 1].color) })
  if (!parts.length) return lerp(s[s.length - 1].color, s[s.length - 1].color)
  if (parts.length === 1) return parts[0].fn
  return `<</FunctionType 3/Domain[0 1]/Functions[${parts.map((p) => p.fn).join('')}]/Bounds[${parts.slice(1).map((p) => nm(p.from)).join(' ')}]/Encode[${parts.map(() => '0 1').join(' ')}]>>`
}
