/**
 * The tables of a font file, whatever it is packed in: a bare TrueType / OpenType file, one font of a collection
 * (TTC), WOFF (its tables deflated) or WOFF2 (Brotli, with the `glyf`, `loca` and `hmtx` tables rewritten).
 * No dependency: the unpacking is the runtime's own `DecompressionStream`.
 */

/** a font file that cannot be read: damaged, of a kind not read here, or too large */
export class FontError extends Error {
  constructor(message: string) {
    super(message)
    this.name = 'FontError'
  }
}

/** what a font file can make the reader read and hold */
export const FONT_LIMITS = {
  /** bytes of a font file */
  file: 64 * 1024 * 1024,
  /** bytes of its tables once unpacked (WOFF, WOFF2) */
  unpacked: 128 * 1024 * 1024,
  /** tables in a file */
  tables: 512,
  /** kerning pairs, ligatures, lookup subtables read from one font */
  records: 2_000_000,
}

export const DAMAGED = 'The font file is damaged'

export interface Sfnt {
  /** `glyf` outlines (TrueType) or `CFF ` ones (PostScript) */
  outlines: 'glyf' | 'cff'
  format: 'sfnt' | 'woff' | 'woff2'
  /** how many fonts the file holds (more than one: a collection) */
  fonts: number
  tables: Map<string, Uint8Array>
}

const tagAt = (b: Uint8Array, at: number) => String.fromCharCode(b[at], b[at + 1], b[at + 2], b[at + 3])
const view = (b: Uint8Array) => new DataView(b.buffer, b.byteOffset, b.byteLength)

const TRUETYPE = 0x00010000
const TRUE = 0x74727565 // 'true'
const OTTO = 0x4f54544f
const TTCF = 0x74746366
const WOFF = 0x774f4646
const WOF2 = 0x774f4632

/** the kind of file by its first bytes: null for anything that is not a font read here */
export function fontFormat(b: Uint8Array): 'sfnt' | 'ttc' | 'woff' | 'woff2' | null {
  if (b.length < 12) return null
  const sig = view(b).getUint32(0)
  return sig === TRUETYPE || sig === TRUE || sig === OTTO ? 'sfnt' : sig === TTCF ? 'ttc' : sig === WOFF ? 'woff' : sig === WOF2 ? 'woff2' : null
}

/** Unpacks `data` (zlib or Brotli) to at most `limit` bytes. */
async function unpack(data: Uint8Array, format: 'deflate' | 'brotli', limit: number): Promise<Uint8Array> {
  let stream: DecompressionStream
  try {
    stream = new DecompressionStream(format as CompressionFormat)
  } catch {
    // (Brotli is the newer of the two: a runtime without it reads WOFF2 fonts only once they are unpacked elsewhere)
    throw new FontError('This font is a WOFF2 file, which this browser cannot unpack: supply it as TTF, OTF or WOFF')
  }
  const reader = new Blob([data as BlobPart]).stream().pipeThrough(stream as unknown as TransformStream<Uint8Array, Uint8Array>).getReader()
  const parts: Uint8Array[] = []
  let size = 0
  try {
    for (;;) {
      const { done, value } = await reader.read()
      if (done) break
      size += value.length
      if (size > limit) {
        void reader.cancel().catch(() => {})
        throw new FontError('The font file is too large once unpacked')
      }
      parts.push(value)
    }
  } catch (e) {
    throw e instanceof FontError ? e : new FontError(`${DAMAGED}: it does not unpack`)
  }
  const out = new Uint8Array(size)
  let at = 0
  for (const p of parts) {
    out.set(p, at)
    at += p.length
  }
  return out
}

function outlinesOf(flavor: number, tables: Map<string, Uint8Array>): 'glyf' | 'cff' {
  if (tables.has('CFF ')) return 'cff'
  if (tables.has('glyf')) return 'glyf'
  if (tables.has('CFF2')) throw new FontError('This font has CFF2 outlines (a variable PostScript font), which are not read')
  // a font of bitmaps or of color pictures only
  if (flavor === OTTO) throw new FontError(`${DAMAGED}: no outlines`)
  return 'glyf'
}

function sfntTables(b: Uint8Array, at: number): Map<string, Uint8Array> {
  const v = view(b)
  const n = v.getUint16(at + 4)
  if (n > FONT_LIMITS.tables) throw new FontError(`${DAMAGED}: too many tables`)
  const tables = new Map<string, Uint8Array>()
  for (let i = 0; i < n; i++) {
    const rec = at + 12 + i * 16
    const offset = v.getUint32(rec + 8)
    const length = v.getUint32(rec + 12)
    if (offset + length > b.length) throw new FontError(`${DAMAGED}: a table ends after the file`)
    tables.set(tagAt(b, rec), b.subarray(offset, offset + length))
  }
  return tables
}

async function woffTables(b: Uint8Array): Promise<Map<string, Uint8Array>> {
  const v = view(b)
  const n = v.getUint16(12)
  if (n > FONT_LIMITS.tables) throw new FontError(`${DAMAGED}: too many tables`)
  const tables = new Map<string, Uint8Array>()
  let total = 0
  for (let i = 0; i < n; i++) {
    const rec = 44 + i * 20
    const offset = v.getUint32(rec + 4)
    const comp = v.getUint32(rec + 8)
    const orig = v.getUint32(rec + 12)
    if (offset + comp > b.length) throw new FontError(`${DAMAGED}: a table ends after the file`)
    total += orig
    if (total > FONT_LIMITS.unpacked) throw new FontError('The font file is too large once unpacked')
    const raw = b.subarray(offset, offset + comp)
    // (a table that does not get smaller is stored as it is)
    const data = comp < orig ? await unpack(raw, 'deflate', orig) : raw
    if (data.length !== orig) throw new FontError(`${DAMAGED}: a table is not of its announced size`)
    tables.set(tagAt(b, rec), data)
  }
  return tables
}

/** the tags a WOFF2 directory names by a number */
const WOFF2_TAGS =
  'cmap head hhea hmtx maxp name OS/2 post cvt  fpgm glyf loca prep CFF  VORG EBDT EBLC gasp hdmx kern LTSH PCLT VDMX vhea vmtx BASE GDEF GPOS GSUB EBSC JSTF MATH CBDT CBLC COLR CPAL SVG  sbix acnt avar bdat bloc bsln cvar fdsc feat fmtx fvar gvar hsty just lcar mort morx opbd prop trak Zapf Silf Glat Gloc Feat Sill'
    .match(/.{4} ?/g)!
    .map((t) => t.slice(0, 4))

/** a cursor over bytes, for the variable-length numbers of WOFF2 */
class Cursor {
  private v: DataView
  constructor(
    readonly b: Uint8Array,
    public at = 0,
  ) {
    this.v = view(b)
  }
  u8() {
    return this.v.getUint8(this.at++)
  }
  u16() {
    const x = this.v.getUint16(this.at)
    this.at += 2
    return x
  }
  i16() {
    const x = this.v.getInt16(this.at)
    this.at += 2
    return x
  }
  u32() {
    const x = this.v.getUint32(this.at)
    this.at += 4
    return x
  }
  /** UIntBase128 */
  base128() {
    let x = 0
    for (let i = 0; i < 5; i++) {
      const c = this.u8()
      // (no leading zeros, no more than 32 bits)
      if ((i === 0 && c === 0x80) || x > 0x1ffffff) break
      x = x * 128 + (c & 0x7f)
      if (!(c & 0x80)) return x
    }
    throw new FontError(`${DAMAGED}: a number of its directory`)
  }
  /** 255UInt16 */
  u255() {
    const c = this.u8()
    return c === 253 ? this.u16() : c === 255 ? this.u8() + 253 : c === 254 ? this.u8() + 506 : c
  }
  take(n: number) {
    if (n < 0 || this.at + n > this.b.length) throw new RangeError('past the end')
    const x = this.b.subarray(this.at, this.at + n)
    this.at += n
    return x
  }
}

async function woff2Tables(b: Uint8Array, index: number): Promise<{ tables: Map<string, Uint8Array>; fonts: number; flavor: number }> {
  const c = new Cursor(b, 4)
  let flavor = c.u32()
  c.at = 12
  const n = c.u16()
  if (n > FONT_LIMITS.tables) throw new FontError(`${DAMAGED}: too many tables`)
  c.at = 20
  const packed = c.u32()
  c.at = 48
  const dir: { tag: string; length: number; transform: number; stored: number; from: number }[] = []
  let total = 0
  for (let i = 0; i < n; i++) {
    const flags = c.u8()
    const tag = (flags & 63) === 63 ? tagAt(c.take(4), 0) : WOFF2_TAGS[flags & 63]
    const version = flags >> 6
    const length = c.base128()
    // glyf and loca are rewritten unless their version says 3, any other table only when it says so
    const transform = tag === 'glyf' || tag === 'loca' ? (version === 3 ? -1 : version) : version === 0 ? -1 : version
    const stored = transform >= 0 ? c.base128() : length
    dir.push({ tag, length, transform, stored, from: total })
    total += stored
    if (total > FONT_LIMITS.unpacked) throw new FontError('The font file is too large once unpacked')
  }
  let fonts = 1
  let mine: number[] | null = null
  if (flavor === TTCF) {
    c.u32()
    fonts = c.u255()
    for (let f = 0; f < fonts; f++) {
      const count = c.u255()
      const fl = c.u32()
      const idx: number[] = []
      for (let i = 0; i < count; i++) idx.push(c.u255())
      if (f === index) {
        mine = idx
        flavor = fl
      }
    }
    if (!mine) throw new FontError(`The font file holds ${fonts} fonts: there is no font ${index}`)
  }
  const data = await unpack(c.take(packed), 'brotli', total)
  if (data.length !== total) throw new FontError(`${DAMAGED}: its tables are not of their announced size`)
  const raw = new Map<string, { bytes: Uint8Array; transform: number; length: number }>()
  for (const i of mine ?? dir.keys()) {
    const d = dir[i]
    if (!d) throw new FontError(`${DAMAGED}: a table of its collection`)
    raw.set(d.tag, { bytes: data.subarray(d.from, d.from + d.stored), transform: d.transform, length: d.length })
  }
  const tables = new Map<string, Uint8Array>()
  for (const [tag, t] of raw) if (t.transform < 0) tables.set(tag, t.bytes)
  const glyf = raw.get('glyf')
  if (glyf && glyf.transform >= 0) {
    if (glyf.transform !== 0) throw new FontError(`${DAMAGED}: its outlines are packed in a way that is not known`)
    const out = woff2Glyf(glyf.bytes)
    tables.set('glyf', out.glyf)
    tables.set('loca', out.loca)
    // the head table says how long a loca entry is: as the rebuilt one has them
    const head = tables.get('head')
    if (head && head.length >= 54) {
      const h = head.slice()
      view(h).setUint16(50, out.long ? 1 : 0)
      tables.set('head', h)
    }
  }
  const hmtx = raw.get('hmtx')
  if (hmtx && hmtx.transform >= 0) {
    if (hmtx.transform !== 1) throw new FontError(`${DAMAGED}: its widths are packed in a way that is not known`)
    tables.set('hmtx', woff2Hmtx(hmtx.bytes, tables))
  }
  return { tables, fonts, flavor }
}

/**
 * The `hmtx` table from WOFF2's short form: the left side bearings it leaves out are the glyphs' xMin, in `glyf`.
 */
function woff2Hmtx(b: Uint8Array, tables: Map<string, Uint8Array>): Uint8Array {
  const hhea = tables.get('hhea')
  const maxp = tables.get('maxp')
  const glyf = tables.get('glyf')
  const loca = tables.get('loca')
  const head = tables.get('head')
  if (!hhea || !maxp || !glyf || !loca || !head) throw new FontError(`${DAMAGED}: its widths need tables it does not have`)
  const numH = view(hhea).getUint16(34)
  const numGlyphs = view(maxp).getUint16(4)
  const long = view(head).getUint16(50) === 1
  const lv = view(loca)
  const gv = view(glyf)
  const xMin = (g: number) => {
    const a = long ? lv.getUint32(g * 4) : lv.getUint16(g * 2) * 2
    const e = long ? lv.getUint32(g * 4 + 4) : lv.getUint16(g * 2 + 2) * 2
    return e > a ? gv.getInt16(a + 2) : 0
  }
  const c = new Cursor(b)
  const flags = c.u8()
  const out = new Uint8Array(numH * 4 + (numGlyphs - numH) * 2)
  const ov = view(out)
  for (let g = 0; g < numH; g++) ov.setUint16(g * 4, c.u16())
  for (let g = 0; g < numH; g++) ov.setInt16(g * 4 + 2, flags & 1 ? xMin(g) : c.i16())
  for (let g = numH; g < numGlyphs; g++) ov.setInt16(numH * 4 + (g - numH) * 2, flags & 2 ? xMin(g) : c.i16())
  return out
}

/** The `glyf` and `loca` tables from WOFF2's form: the parts of every glyph in streams of their own. */
function woff2Glyf(b: Uint8Array): { glyf: Uint8Array; loca: Uint8Array; long: boolean } {
  const h = new Cursor(b)
  h.u16()
  const options = h.u16()
  const numGlyphs = h.u16()
  h.u16()
  const sizes = [h.u32(), h.u32(), h.u32(), h.u32(), h.u32(), h.u32(), h.u32()]
  const streams = sizes.map((n) => new Cursor(h.take(n)))
  const [contours, points, flagsS, glyphs, composites, bboxes, instructions] = streams
  const overlap = options & 1 ? h.take((numGlyphs + 7) >> 3) : null
  const bitmap = bboxes.take(((numGlyphs + 31) >> 5) << 2)
  const out: Uint8Array[] = []
  let size = 0
  const i16 = (o: number[], x: number) => o.push((x >> 8) & 255, x & 255)
  for (let g = 0; g < numGlyphs; g++) {
    const n = contours.i16()
    const explicit = !!(bitmap[g >> 3] & (0x80 >> (g & 7)))
    const o: number[] = []
    if (n === 0) {
      if (explicit) throw new FontError(`${DAMAGED}: an empty glyph with a box`)
    } else if (n > 0) {
      const ends: number[] = []
      let total = 0
      for (let k = 0; k < n; k++) ends.push((total += points.u255()) - 1)
      const fl: number[] = []
      const xs: number[] = []
      const ys: number[] = []
      let x = 0
      let y = 0
      for (let p = 0; p < total; p++) {
        const f = flagsS.u8()
        const on = !(f & 0x80)
        const t = f & 0x7f
        let dx: number
        let dy: number
        // the coordinate triplets of the WOFF2 specification (section 5.2)
        if (t < 10) {
          dx = 0
          dy = ((t & 14) << 7) + glyphs.u8()
          if (!(t & 1)) dy = -dy
        } else if (t < 20) {
          dx = (((t - 10) & 14) << 7) + glyphs.u8()
          if (!((t - 10) & 1)) dx = -dx
          dy = 0
        } else if (t < 84) {
          const b0 = t - 20
          const b1 = glyphs.u8()
          dx = 1 + (b0 & 0x30) + (b1 >> 4)
          dy = 1 + ((b0 & 0x0c) << 2) + (b1 & 0x0f)
          if (!(b0 & 1)) dx = -dx
          if (!(b0 & 2)) dy = -dy
        } else if (t < 120) {
          const b0 = t - 84
          dx = 1 + (((b0 / 12) | 0) << 8) + glyphs.u8()
          dy = 1 + ((((b0 % 12) >> 2) | 0) << 8) + glyphs.u8()
          if (!(b0 & 1)) dx = -dx
          if (!(b0 & 2)) dy = -dy
        } else if (t < 124) {
          const b1 = glyphs.u8()
          const b2 = glyphs.u8()
          const b3 = glyphs.u8()
          dx = (b1 << 4) + (b2 >> 4)
          dy = ((b2 & 0x0f) << 8) + b3
          if (!((t - 120) & 1)) dx = -dx
          if (!((t - 120) & 2)) dy = -dy
        } else {
          dx = glyphs.u16()
          dy = glyphs.u16()
          if (!((t - 124) & 1)) dx = -dx
          if (!((t - 124) & 2)) dy = -dy
        }
        x += dx
        y += dy
        xs.push(x)
        ys.push(y)
        fl.push(on ? 1 : 0)
      }
      const ins = instructions.take(glyphs.u255())
      let box: number[]
      if (explicit) box = [bboxes.i16(), bboxes.i16(), bboxes.i16(), bboxes.i16()]
      else if (!total) box = [0, 0, 0, 0]
      else {
        // (reduce, not Math.min(...xs): a glyph can have more points than a call takes arguments)
        box = [xs.reduce((a, c) => Math.min(a, c)), ys.reduce((a, c) => Math.min(a, c)), xs.reduce((a, c) => Math.max(a, c)), ys.reduce((a, c) => Math.max(a, c))]
      }
      i16(o, n)
      for (const v of box) i16(o, v)
      for (const e of ends) i16(o, e)
      i16(o, ins.length)
      for (const v of ins) o.push(v)
      // flags, then x and y as differences, each in the shortest form of the glyf table
      const fb: number[] = []
      const xb: number[] = []
      const yb: number[] = []
      let px = 0
      let py = 0
      for (let p = 0; p < total; p++) {
        let f = fl[p]
        if (p === 0 && overlap && overlap[g >> 3] & (0x80 >> (g & 7))) f |= 0x40
        const dx = xs[p] - px
        const dy = ys[p] - py
        px = xs[p]
        py = ys[p]
        if (dx === 0) f |= 0x10
        else if (dx > -256 && dx < 256) {
          f |= dx > 0 ? 0x12 : 0x02
          xb.push(Math.abs(dx))
        } else i16(xb, dx)
        if (dy === 0) f |= 0x20
        else if (dy > -256 && dy < 256) {
          f |= dy > 0 ? 0x24 : 0x04
          yb.push(Math.abs(dy))
        } else i16(yb, dy)
        fb.push(f)
      }
      for (const v of fb) o.push(v)
      for (const v of xb) o.push(v)
      for (const v of yb) o.push(v)
    } else {
      // a composite: its components as they are, and its box, which is always given
      if (!explicit) throw new FontError(`${DAMAGED}: a composite glyph without a box`)
      i16(o, -1)
      for (let k = 0; k < 4; k++) i16(o, bboxes.i16())
      let more = true
      let hasIns = false
      while (more) {
        const f = composites.u16()
        more = !!(f & 0x20)
        hasIns ||= !!(f & 0x100)
        const args = (f & 1 ? 4 : 2) + (f & 8 ? 2 : f & 0x40 ? 4 : f & 0x80 ? 8 : 0)
        i16(o, f)
        for (const v of composites.take(2 + args)) o.push(v)
      }
      if (hasIns) {
        const ins = instructions.take(glyphs.u255())
        i16(o, ins.length)
        for (const v of ins) o.push(v)
      }
    }
    // glyphs start at even offsets (a short loca counts in pairs of bytes)
    if (o.length & 1) o.push(0)
    out.push(Uint8Array.from(o))
    size += o.length
    if (size > FONT_LIMITS.unpacked) throw new FontError('The font file is too large once unpacked')
  }
  const long = size > 0x1fffe
  const glyf = new Uint8Array(size)
  const loca = new Uint8Array((numGlyphs + 1) * (long ? 4 : 2))
  const lv = view(loca)
  let at = 0
  out.forEach((o, g) => {
    if (long) lv.setUint32(g * 4, at)
    else lv.setUint16(g * 2, at >> 1)
    glyf.set(o, at)
    at += o.length
  })
  if (long) lv.setUint32(numGlyphs * 4, at)
  else lv.setUint16(numGlyphs * 2, at >> 1)
  return { glyf, loca, long }
}

/**
 * The tables of the font in `bytes`; `index`: which font of a collection (the first by default).
 * Throws a `FontError` that says what is wrong with a file that cannot be read.
 */
export async function readSfnt(bytes: Uint8Array, index = 0): Promise<Sfnt> {
  if (bytes.length > FONT_LIMITS.file) throw new FontError('The font file is too large')
  const format = fontFormat(bytes)
  if (!format) throw new FontError('Not a font file (TTF, OTF, TTC, WOFF or WOFF2)')
  try {
    const v = view(bytes)
    if (format === 'woff') {
      if (index) throw new FontError(`The font file holds 1 font: there is no font ${index}`)
      const tables = await woffTables(bytes)
      return { outlines: outlinesOf(v.getUint32(4), tables), format, fonts: 1, tables }
    }
    if (format === 'woff2') {
      const { tables, fonts, flavor } = await woff2Tables(bytes, index)
      if (index >= fonts) throw new FontError(`The font file holds ${fonts} font${fonts > 1 ? 's' : ''}: there is no font ${index}`)
      return { outlines: outlinesOf(flavor, tables), format, fonts, tables }
    }
    let at = 0
    let fonts = 1
    if (format === 'ttc') {
      fonts = v.getUint32(8)
      if (index >= fonts) throw new FontError(`The font file holds ${fonts} fonts: there is no font ${index}`)
      at = v.getUint32(12 + index * 4)
    } else if (index) throw new FontError(`The font file holds 1 font: there is no font ${index}`)
    const tables = sfntTables(bytes, at)
    return { outlines: outlinesOf(v.getUint32(at), tables), format: 'sfnt', fonts, tables }
  } catch (e) {
    // (a read past the end of the bytes: a DataView's own RangeError)
    throw e instanceof FontError ? e : new FontError(`${DAMAGED}: it ends early`)
  }
}

function checksum(b: Uint8Array): number {
  const v = view(b)
  let sum = 0
  const whole = b.length & ~3
  for (let i = 0; i < whole; i += 4) sum = (sum + v.getUint32(i)) >>> 0
  let last = 0
  for (let i = whole; i < b.length; i++) last |= b[i] << (24 - (i - whole) * 8)
  return (sum + last) >>> 0
}

/**
 * A bare font file of these tables, as a browser or a PDF reader takes one: the directory sorted by tag, every table
 * at a multiple of four bytes, the checksums and the head table's adjustment set.
 */
export function writeSfnt(tables: Map<string, Uint8Array>, outlines: 'glyf' | 'cff'): Uint8Array<ArrayBuffer> {
  const tags = [...tables.keys()].sort()
  const n = tags.length
  let size = 12 + n * 16
  const offsets = tags.map((t) => {
    const at = size
    size += (tables.get(t)!.length + 3) & ~3
    return at
  })
  const out = new Uint8Array(size)
  const v = view(out)
  v.setUint32(0, outlines === 'cff' ? OTTO : TRUETYPE)
  v.setUint16(4, n)
  const pow = n ? 31 - Math.clz32(n) : 0
  v.setUint16(6, (1 << pow) * 16)
  v.setUint16(8, pow)
  v.setUint16(10, n * 16 - (1 << pow) * 16)
  let headAt = -1
  tags.forEach((t, i) => {
    let data = tables.get(t)!
    if (t === 'head' && data.length >= 12) {
      // (the adjustment is counted with this field at zero)
      data = data.slice()
      view(data).setUint32(8, 0)
      headAt = offsets[i]
    }
    out.set(data, offsets[i])
    const rec = 12 + i * 16
    for (let k = 0; k < 4; k++) out[rec + k] = t.charCodeAt(k)
    v.setUint32(rec + 4, checksum(data))
    v.setUint32(rec + 8, offsets[i])
    v.setUint32(rec + 12, data.length)
  })
  if (headAt >= 0) v.setUint32(headAt + 8, (0xb1b0afba - checksum(out)) >>> 0)
  return out
}
