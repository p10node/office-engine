import { DAMAGED, FontError, writeSfnt } from './container'
import type { Font } from './font'

/**
 * A font cut down to the glyphs a document uses, for embedding (a PDF holds only what it draws): the glyphs
 * renumbered from 0, everything else about them as in the font. TrueType outlines are copied glyph by glyph, with
 * the parts composite glyphs are made of; PostScript ones (CFF) keep their charstrings and all subroutines.
 * The result is a font file of its own: a browser loads it, and a PDF takes it as it is.
 */

const view = (b: Uint8Array) => new DataView(b.buffer, b.byteOffset, b.byteLength)

export interface FontSubset {
  /** a bare font file (TTF, or OTF for PostScript outlines) of the glyphs kept */
  bytes: Uint8Array<ArrayBuffer>
  /** glyph of the font → glyph of the subset. Glyph 0 (the "missing glyph") is always kept, as 0 */
  glyphs: Map<number, number>
  outlines: 'glyf' | 'cff'
  /** PostScript outlines: the CFF table alone, which is what a PDF embeds for them (`FontFile3`, `CIDFontType0C`) */
  cff?: Uint8Array
}

class Out {
  b: number[] = []
  u8(v: number) {
    this.b.push(v & 0xff)
    return this
  }
  u16(v: number) {
    this.b.push((v >> 8) & 0xff, v & 0xff)
    return this
  }
  u32(v: number) {
    return this.u16((v >>> 16) & 0xffff).u16(v & 0xffff)
  }
  bytes(v: ArrayLike<number>) {
    for (let i = 0; i < v.length; i++) this.b.push(v[i])
    return this
  }
  get length() {
    return this.b.length
  }
  done() {
    return Uint8Array.from(this.b)
  }
}

/**
 * A cmap table for `chars` (code point → glyph): a format 4 subtable for the Basic Multilingual Plane, and a
 * format 12 one for everything when a character is beyond it.
 */
function writeCmap(chars: [number, number][]): Uint8Array {
  const sorted = [...new Map(chars.filter(([cp]) => cp >= 0 && cp <= 0x10ffff)).entries()].sort((a, b) => a[0] - b[0])
  /** runs of characters whose glyphs follow each other as they do */
  const runs = (max: number) => {
    const out: { start: number; end: number; glyph: number }[] = []
    for (const [cp, g] of sorted) {
      if (cp > max) break
      const last = out[out.length - 1]
      if (last && cp === last.end + 1 && g === last.glyph + (cp - last.start)) last.end = cp
      else out.push({ start: cp, end: cp, glyph: g })
    }
    return out
  }
  const bmp = runs(0xfffe)
  // the format ends with a segment for U+FFFF alone, mapped to no glyph
  const segs = [...bmp, { start: 0xffff, end: 0xffff, glyph: 0 }]
  const n = segs.length
  const f4 = new Out()
  const pow = 31 - Math.clz32(n)
  f4.u16(4).u16(16 + n * 8).u16(0).u16(n * 2).u16((1 << pow) * 2).u16(pow).u16(n * 2 - (1 << pow) * 2)
  for (const s of segs) f4.u16(s.end)
  f4.u16(0)
  for (const s of segs) f4.u16(s.start)
  segs.forEach((s, i) => f4.u16(i === n - 1 ? 1 : (s.glyph - s.start) & 0xffff))
  for (let i = 0; i < n; i++) f4.u16(0)
  const astral = sorted.length > 0 && sorted[sorted.length - 1][0] > 0xffff
  // format 4 says its length in 16 bits: a map of more runs than fit (thousands of scattered characters) is written in format 12 alone
  const wide = f4.length > 0xffff
  const out = new Out()
  if (wide) out.u16(0).u16(1).u16(3).u16(10).u32(12)
  else {
    out.u16(0).u16(astral ? 2 : 1)
    out.u16(3).u16(1).u32(astral ? 20 : 12)
    if (astral) out.u16(3).u16(10).u32(20 + f4.length)
    out.bytes(f4.b)
  }
  if (astral || wide) {
    const groups = runs(0x10ffff)
    out.u16(12).u16(0).u32(16 + groups.length * 12).u32(0).u32(groups.length)
    for (const g of groups) out.u32(g.start).u32(g.end).u32(g.glyph)
  }
  return out.done()
}

/** the tables every subset has: the glyph count and widths of the glyphs kept, a character map, no glyph names */
function commonTables(font: Font, kept: number[], chars: [number, number][], glyphs: Map<number, number>): Map<string, Uint8Array> {
  const t = font.tables
  const tables = new Map<string, Uint8Array>()
  for (const tag of ['OS/2', 'name', 'head', 'hhea']) if (t.has(tag)) tables.set(tag, t.get(tag)!.slice())
  const n = kept.length
  // every glyph with its own width: the short form (one width for the last glyphs) saves nothing in a subset
  view(tables.get('hhea')!).setUint16(34, n)
  const old = view(t.get('hmtx')!)
  const numH = view(t.get('hhea')!).getUint16(34)
  const hmtx = new Uint8Array(n * 4)
  const hv = view(hmtx)
  kept.forEach((g, i) => {
    const at = g < numH ? g * 4 + 2 : numH * 4 + (g - numH) * 2
    hv.setUint16(i * 4, font.advanceOf(g))
    hv.setInt16(i * 4 + 2, at + 2 <= old.byteLength ? old.getInt16(at) : 0)
  })
  tables.set('hmtx', hmtx)
  // post, version 3: the measures of the underline, and no names
  const post = new Uint8Array(32)
  post.set((t.get('post') ?? new Uint8Array(0)).subarray(0, 32))
  view(post).setUint32(0, 0x00030000)
  tables.set('post', post)
  const mapped: [number, number][] = []
  for (const [cp, g] of chars) if (glyphs.has(g)) mapped.push([cp, glyphs.get(g)!])
  tables.set('cmap', writeCmap(mapped))
  return tables
}

function subsetGlyf(font: Font, used: Iterable<number>, chars: [number, number][]): FontSubset {
  const t = font.tables
  const glyf = t.get('glyf')
  const loca = t.get('loca')
  if (!glyf || !loca) throw new FontError(`${DAMAGED}: no outlines`)
  const long = view(t.get('head')!).getUint16(50) === 1
  const lv = view(loca)
  const gv = view(glyf)
  const range = (g: number): [number, number] => {
    const a = long ? lv.getUint32(g * 4) : lv.getUint16(g * 2) * 2
    const b = long ? lv.getUint32(g * 4 + 4) : lv.getUint16(g * 2 + 2) * 2
    if (b < a || b > glyf.length) throw new FontError(`${DAMAGED}: a glyph ends after its table`)
    return [a, b]
  }
  /** the components of a composite glyph: where each glyph number is, and the number */
  const components = (g: number): [number, number][] => {
    const [a, b] = range(g)
    if (b - a < 10 || gv.getInt16(a) >= 0) return []
    const out: [number, number][] = []
    let p = a + 10
    while (p + 4 <= b) {
      const flags = gv.getUint16(p)
      out.push([p + 2 - a, gv.getUint16(p + 2)])
      p += 4 + (flags & 1 ? 4 : 2) + (flags & 8 ? 2 : flags & 0x40 ? 4 : flags & 0x80 ? 8 : 0)
      if (!(flags & 0x20)) break
    }
    return out
  }
  // the glyphs asked for, and the ones they are made of (which can be made of others)
  const keep = new Set<number>()
  const todo = [0, ...used]
  while (todo.length) {
    const g = todo.pop()!
    if (!(g >= 0 && g < font.numGlyphs) || keep.has(g)) continue
    keep.add(g)
    for (const [, c] of components(g)) todo.push(c)
  }
  const kept = [...keep].sort((a, b) => a - b)
  const glyphs = new Map(kept.map((g, i) => [g, i]))
  const parts: Uint8Array[] = []
  const offsets: number[] = []
  let size = 0
  for (const g of kept) {
    offsets.push(size)
    const [a, b] = range(g)
    const data = glyf.slice(a, b)
    const dv = view(data)
    // a composite names its parts by their numbers in the subset
    for (const [at, c] of components(g)) dv.setUint16(at, glyphs.get(c) ?? 0)
    parts.push(data)
    size += (data.length + 3) & ~3
  }
  offsets.push(size)
  const newGlyf = new Uint8Array(size)
  parts.forEach((p, i) => newGlyf.set(p, offsets[i]))
  const newLoca = new Uint8Array(offsets.length * 4)
  offsets.forEach((o, i) => view(newLoca).setUint32(i * 4, o))
  const tables = commonTables(font, kept, chars, glyphs)
  tables.set('glyf', newGlyf)
  tables.set('loca', newLoca)
  view(tables.get('head')!).setUint16(50, 1)
  const maxp = t.get('maxp')!.slice()
  view(maxp).setUint16(4, kept.length)
  tables.set('maxp', maxp)
  // what the glyphs' instructions use
  for (const tag of ['cvt ', 'fpgm', 'prep', 'gasp']) if (t.has(tag)) tables.set(tag, t.get(tag)!)
  return { bytes: writeSfnt(tables, 'glyf'), glyphs, outlines: 'glyf' }
}

// ---------------------------------------------------------------- CFF

/** the items of a CFF INDEX at `at`, and where it ends */
function readIndex(b: Uint8Array, at: number): { items: Uint8Array[]; end: number } {
  const v = view(b)
  const count = v.getUint16(at)
  if (!count) return { items: [], end: at + 2 }
  const size = v.getUint8(at + 2)
  if (size < 1 || size > 4) throw new FontError(`${DAMAGED}: an index of its outlines`)
  const off = (i: number) => {
    let x = 0
    for (let k = 0; k < size; k++) x = x * 256 + v.getUint8(at + 3 + i * size + k)
    return x
  }
  const base = at + 3 + (count + 1) * size - 1
  const items: Uint8Array[] = []
  let prev = off(0)
  for (let i = 1; i <= count; i++) {
    const next = off(i)
    if (next < prev || base + next > b.length) throw new FontError(`${DAMAGED}: an index of its outlines`)
    items.push(b.subarray(base + prev, base + next))
    prev = next
  }
  return { items, end: base + prev }
}

function writeIndex(out: Out, items: ArrayLike<number>[]) {
  out.u16(items.length)
  if (!items.length) return
  let total = 1
  for (const it of items) total += it.length
  const size = total < 0x100 ? 1 : total < 0x10000 ? 2 : total < 0x1000000 ? 3 : 4
  out.u8(size)
  let at = 1
  const off = (x: number) => {
    for (let k = size - 1; k >= 0; k--) out.u8(x / 2 ** (k * 8))
  }
  off(at)
  for (const it of items) off((at += it.length))
  for (const it of items) out.bytes(it)
}

/** an entry of a CFF DICT: the operator (two-byte ones as 1200 + the second byte), its whole-number operands, and its bytes */
interface DictEntry {
  op: number
  nums: number[]
  raw: Uint8Array
}

function readDict(b: Uint8Array): DictEntry[] {
  const out: DictEntry[] = []
  let nums: number[] = []
  let start = 0
  for (let i = 0; i < b.length; ) {
    const c = b[i]
    if (c <= 21) {
      const op = c === 12 ? 1200 + b[i + 1] : c
      i += c === 12 ? 2 : 1
      out.push({ op, nums, raw: b.subarray(start, i) })
      nums = []
      start = i
    } else if (c === 28) {
      nums.push(((b[i + 1] << 8) | b[i + 2]) << 16 >> 16)
      i += 3
    } else if (c === 29) {
      nums.push((b[i + 1] << 24) | (b[i + 2] << 16) | (b[i + 3] << 8) | b[i + 4])
      i += 5
    } else if (c === 30) {
      // a real number, in nibbles up to an f: kept as it is written, never needed as a number here
      i++
      while (i < b.length && (b[i] & 0x0f) !== 0x0f && (b[i] & 0xf0) !== 0xf0) i++
      i++
      nums.push(NaN)
    } else if (c >= 32 && c <= 246) {
      nums.push(c - 139)
      i++
    } else if (c >= 247 && c <= 250) {
      nums.push((c - 247) * 256 + b[i + 1] + 108)
      i += 2
    } else if (c >= 251 && c <= 254) {
      nums.push(-(c - 251) * 256 - b[i + 1] - 108)
      i += 2
    } else throw new FontError(`${DAMAGED}: a dictionary of its outlines`)
  }
  return out
}

/** a whole number as a DICT operand of five bytes, whatever its size: so that a dictionary's length does not depend on the offsets in it */
const int5 = (out: Out, x: number) => out.u8(29).u32(x >>> 0)
const opBytes = (out: Out, op: number) => (op >= 1200 ? out.u8(12).u8(op - 1200) : out.u8(op))

/** `entries` as bytes, with the operands of the operators in `set` replaced (by five-byte numbers) and those in `drop` left out */
function writeDict(entries: DictEntry[], set: Map<number, number[]>, drop: Set<number> = new Set()): number[] {
  const out = new Out()
  const done = new Set<number>()
  for (const e of entries) {
    if (drop.has(e.op)) continue
    const to = set.get(e.op)
    if (to) {
      for (const x of to) int5(out, x)
      opBytes(out, e.op)
      done.add(e.op)
    } else out.bytes(e.raw)
  }
  for (const [op, to] of set) {
    if (done.has(op)) continue
    for (const x of to) int5(out, x)
    opBytes(out, op)
  }
  return out.b
}

const CHARSET = 15
const ENCODING = 16
const CHARSTRINGS = 17
const PRIVATE = 18
const SUBRS = 19
const ROS = 1230
const FDARRAY = 1236
const FDSELECT = 1237
const UNIQUE_ID = 13
const XUID = 14

function subsetCff(font: Font, used: Iterable<number>, chars: [number, number][]): FontSubset {
  const cff = font.tables.get('CFF ')!
  const v = view(cff)
  if (v.getUint8(0) !== 1) throw new FontError('This font has PostScript outlines of a version that is not read')
  const names = readIndex(cff, v.getUint8(2))
  const tops = readIndex(cff, names.end)
  const strings = readIndex(cff, tops.end)
  const gsubrs = readIndex(cff, strings.end)
  if (tops.items.length !== 1) throw new FontError(`${DAMAGED}: its outlines hold ${tops.items.length} fonts`)
  const top = readDict(tops.items[0])
  const get = (op: number) => top.find((e) => e.op === op)?.nums
  const charStringsAt = get(CHARSTRINGS)?.[0]
  if (!charStringsAt) throw new FontError(`${DAMAGED}: its outlines have no glyphs`)
  const charStrings = readIndex(cff, charStringsAt).items
  const numGlyphs = charStrings.length
  const cid = !!get(ROS)

  const keep = new Set<number>([0])
  for (const g of used) if (g >= 0 && g < numGlyphs) keep.add(g)
  const kept = [...keep].sort((a, b) => a - b)
  const glyphs = new Map(kept.map((g, i) => [g, i]))

  // the name (or the CID) of every glyph of the font: glyph 0 has none
  const ids = new Uint16Array(numGlyphs)
  const charsetAt = get(CHARSET)?.[0] ?? 0
  if (charsetAt <= 2) {
    // a predefined set: the standard one names glyph n by string n (the two "expert" ones are of fonts of small capitals and fractions only)
    if (charsetAt !== 0) throw new FontError('This font names its glyphs by a predefined expert set, which is not read')
    for (let g = 1; g < numGlyphs; g++) ids[g] = g
  } else {
    const format = v.getUint8(charsetAt)
    let p = charsetAt + 1
    if (format === 0) for (let g = 1; g < numGlyphs; g++, p += 2) ids[g] = v.getUint16(p)
    else if (format === 1 || format === 2) {
      for (let g = 1; g < numGlyphs; ) {
        const first = v.getUint16(p)
        const left = format === 1 ? v.getUint8(p + 2) : v.getUint16(p + 2)
        p += format === 1 ? 3 : 4
        for (let k = 0; k <= left && g < numGlyphs; k++, g++) ids[g] = first + k
      }
    } else throw new FontError(`${DAMAGED}: the names of its glyphs`)
  }

  /** a Private DICT and the subroutines after it, as bytes that can go anywhere: [the DICT with its subroutines right after it, the DICT's length] */
  const privateBytes = (size: number, at: number): { bytes: number[]; size: number } => {
    if (at + size > cff.length) throw new FontError(`${DAMAGED}: a dictionary of its outlines`)
    const dict = readDict(cff.subarray(at, at + size))
    const subrs = dict.find((e) => e.op === SUBRS)?.nums[0]
    if (!subrs) return { bytes: [...cff.subarray(at, at + size)], size }
    // the subroutines are found from the start of the DICT: put right after it
    const length = writeDict(dict, new Map([[SUBRS, [0]]])).length
    const bytes = writeDict(dict, new Map([[SUBRS, [length]]]))
    const index = readIndex(cff, at + subrs)
    for (let i = at + subrs; i < index.end; i++) bytes.push(cff[i])
    return { bytes, size: length }
  }

  // the parts after the Top DICT, each placed once the length of the Top DICT is known
  const head = new Out()
  head.bytes(cff.subarray(0, v.getUint8(2)))
  head.bytes(cff.subarray(v.getUint8(2), names.end))
  const tail = new Out()
  tail.bytes(cff.subarray(tops.end, gsubrs.end))

  const charset = new Out()
  // format 0: the name of each glyph kept; a CID font is renumbered, its CIDs are the glyphs of the subset
  charset.u8(0)
  for (let i = 1; i < kept.length; i++) charset.u16(cid ? i : ids[kept[i]])

  const strs = new Out()
  writeIndex(strs, kept.map((g) => charStrings[g]))

  const set = new Map<number, number[]>()
  const drop = new Set([ENCODING, UNIQUE_ID, XUID])
  let build: (base: number) => number[]
  if (cid) {
    const fdArrayAt = get(FDARRAY)?.[0]
    const fdSelectAt = get(FDSELECT)?.[0]
    if (!fdArrayAt || !fdSelectAt) throw new FontError(`${DAMAGED}: its outlines have no font dictionaries`)
    const fds = readIndex(cff, fdArrayAt).items.map((d) => readDict(d))
    const fdOf = new Uint8Array(numGlyphs)
    const format = v.getUint8(fdSelectAt)
    if (format === 0) for (let g = 0; g < numGlyphs; g++) fdOf[g] = v.getUint8(fdSelectAt + 1 + g)
    else if (format === 3) {
      const n = v.getUint16(fdSelectAt + 1)
      for (let i = 0; i < n; i++) {
        const first = v.getUint16(fdSelectAt + 3 + i * 3)
        const fd = v.getUint8(fdSelectAt + 5 + i * 3)
        const next = v.getUint16(fdSelectAt + 6 + i * 3)
        for (let g = first; g < next && g < numGlyphs; g++) fdOf[g] = fd
      }
    } else throw new FontError(`${DAMAGED}: the font dictionary of its glyphs`)
    const select = new Out()
    select.u8(0)
    for (const g of kept) select.u8(fdOf[g])
    const privates = fds.map((d) => {
      const p = d.find((e) => e.op === PRIVATE)?.nums
      return p ? privateBytes(p[0], p[1]) : { bytes: [], size: 0 }
    })
    build = (base) => {
      const charsetAt2 = base + tail.length
      const selectAt = charsetAt2 + charset.length
      const stringsAt = selectAt + select.length
      const fdAt = stringsAt + strs.length
      // the font dictionaries are as long whatever offsets they hold: written once to know where the private ones start
      const dicts = (at: number[]) => fds.map((d, i) => writeDict(d, new Map([[PRIVATE, [privates[i].size, at[i]]]])))
      const probe = new Out()
      writeIndex(probe, dicts(fds.map(() => 0)))
      let p = fdAt + probe.length
      const at = privates.map((x) => {
        const here = p
        p += x.bytes.length
        return here
      })
      const fdIndex = new Out()
      writeIndex(fdIndex, dicts(at))
      set.set(CHARSET, [charsetAt2]).set(FDSELECT, [selectAt]).set(CHARSTRINGS, [stringsAt]).set(FDARRAY, [fdAt])
      const out = [...tail.b, ...charset.b, ...select.b, ...strs.b, ...fdIndex.b]
      for (const x of privates) for (const byte of x.bytes) out.push(byte)
      return out
    }
  } else {
    const p = get(PRIVATE)
    const priv = p ? privateBytes(p[0], p[1]) : { bytes: [], size: 0 }
    build = (base) => {
      const charsetAt2 = base + tail.length
      const stringsAt = charsetAt2 + charset.length
      const privateAt = stringsAt + strs.length
      set.set(CHARSET, [charsetAt2]).set(CHARSTRINGS, [stringsAt])
      if (p) set.set(PRIVATE, [priv.size, privateAt])
      return [...tail.b, ...charset.b, ...strs.b, ...priv.bytes]
    }
  }
  // (the offsets do not change the Top DICT's length: they are five bytes each, also in the first pass that measures it)
  build(0)
  const probe = new Out()
  writeIndex(probe, [writeDict(top, set, drop)])
  const rest = build(head.length + probe.length)
  const topIndex = new Out()
  writeIndex(topIndex, [writeDict(top, set, drop)])
  const table = Uint8Array.from([...head.b, ...topIndex.b, ...rest])

  const tables = commonTables(font, kept, chars, glyphs)
  tables.set('CFF ', table)
  const maxp = new Uint8Array(6)
  view(maxp).setUint32(0, 0x00005000)
  view(maxp).setUint16(4, kept.length)
  tables.set('maxp', maxp)
  return { bytes: writeSfnt(tables, 'cff'), glyphs, outlines: 'cff', cff: table }
}

/**
 * The font with only the glyphs `used` (and what they are made of), renumbered. `chars`: the characters the subset's
 * own character map should give (code point, glyph of the font): what a browser needs to draw text in it; a PDF
 * draws by glyph number and needs none.
 */
export function subsetFont(font: Font, used: Iterable<number>, chars: Iterable<[number, number]> = []): FontSubset {
  try {
    return font.outlines === 'cff' ? subsetCff(font, used, [...chars]) : subsetGlyf(font, used, [...chars])
  } catch (e) {
    throw e instanceof FontError ? e : new FontError(`${DAMAGED}: its outlines end early`)
  }
}

const hex4 = (n: number) => n.toString(16).toUpperCase().padStart(4, '0')

/**
 * The ToUnicode CMap of a PDF font drawn by glyph numbers: for each glyph the text it stands for (several
 * characters for a ligature), so that text can be selected, searched and copied. `entries`: glyph (two bytes, as
 * drawn) and its text.
 */
export function toUnicodeCMap(entries: Iterable<[number, string]>): string {
  const lines: string[] = []
  for (const [glyph, text] of [...new Map(entries).entries()].sort((a, b) => a[0] - b[0])) {
    if (!text || glyph < 0 || glyph > 0xffff) continue
    let utf16 = ''
    for (let i = 0; i < text.length; i++) utf16 += hex4(text.charCodeAt(i))
    lines.push(`<${hex4(glyph)}> <${utf16}>`)
  }
  let body = ''
  // a block holds at most 100 entries
  for (let i = 0; i < lines.length; i += 100) {
    const block = lines.slice(i, i + 100)
    body += `${block.length} beginbfchar\n${block.join('\n')}\nendbfchar\n`
  }
  return (
    '/CIDInit /ProcSet findresource begin\n12 dict begin\nbegincmap\n' +
    '/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def\n/CMapName /Adobe-Identity-UCS def\n/CMapType 2 def\n' +
    '1 begincodespacerange\n<0000> <FFFF>\nendcodespacerange\n' +
    body +
    'endcmap\nCMapName currentdict /CMap defineresource pop\nend\nend\n'
  )
}
