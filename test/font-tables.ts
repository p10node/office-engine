/**
 * Writes fonts for the tests of `src/core/font`: a font of box glyphs with the tables a test asks for (character
 * maps of both kinds, a kern table, GSUB and GPOS lookups, composite glyphs, PostScript outlines), and the
 * containers a font comes in (a collection, WOFF). Made when a test runs: no font file in the repository.
 * `test/font-builder.ts` is the plain font of the viewer's tests; this one is for reading fonts.
 */

export class Bytes {
  b: number[] = []
  u8(v: number) {
    this.b.push(v & 0xff)
    return this
  }
  u16(v: number) {
    return this.u8(v >> 8).u8(v)
  }
  i16(v: number) {
    return this.u16(v < 0 ? v + 0x10000 : v)
  }
  u32(v: number) {
    return this.u16((v >>> 16) & 0xffff).u16(v & 0xffff)
  }
  tag(s: string) {
    for (const c of s) this.u8(c.charCodeAt(0))
    return this
  }
  zeros(n: number) {
    for (let i = 0; i < n; i++) this.u8(0)
    return this
  }
  add(v: ArrayLike<number>) {
    for (let i = 0; i < v.length; i++) this.b.push(v[i])
    return this
  }
}

const table = (fill: (o: Bytes) => void) => {
  const o = new Bytes()
  fill(o)
  return o.b
}

// ---------------------------------------------------------------- GSUB / GPOS

/** a coverage table (format 1) of `glyphs` */
export const coverage = (glyphs: number[]) =>
  table((o) => {
    const sorted = [...glyphs].sort((a, b) => a - b)
    o.u16(1).u16(sorted.length)
    for (const g of sorted) o.u16(g)
  })

/** a class table (format 2): `classes[k]` are the glyphs of class k + 1 */
export const classDef = (classes: number[][]) =>
  table((o) => {
    const ranges = classes.flatMap((gs, k) => gs.map((g) => [g, k + 1])).sort((a, b) => a[0] - b[0])
    o.u16(2).u16(ranges.length)
    for (const [g, c] of ranges) o.u16(g).u16(g).u16(c)
  })

/** a subtable whose header is written by `head`, given where each of `parts` lands after it */
function withParts(headSize: number, parts: number[][], head: (o: Bytes, at: number[]) => void): number[] {
  const at: number[] = []
  let p = headSize
  for (const part of parts) {
    at.push(p)
    p += part.length
  }
  return table((o) => {
    head(o, at)
    if (o.b.length !== headSize) throw new Error(`header of ${o.b.length} bytes, ${headSize} announced`)
    for (const part of parts) o.add(part)
  })
}

/** GSUB 1, format 2: each glyph of `map` becomes another */
export function singleSubst(map: [number, number][]): number[] {
  const sorted = [...map].sort((a, b) => a[0] - b[0])
  return withParts(6 + sorted.length * 2, [coverage(sorted.map((m) => m[0]))], (o, at) => {
    o.u16(2).u16(at[0]).u16(sorted.length)
    for (const [, to] of sorted) o.u16(to)
  })
}

/** GSUB 4: `ligs` are [the glyphs in a row, the one glyph they become]; those of one first glyph are tried in the order given */
export function ligatureSubst(ligs: [number[], number][]): number[] {
  const firsts = [...new Set(ligs.map((l) => l[0][0]))].sort((a, b) => a - b)
  const sets = firsts.map((first) => {
    const mine = ligs.filter((l) => l[0][0] === first)
    const bodies = mine.map(([seq, to]) =>
      table((o) => {
        o.u16(to).u16(seq.length)
        for (const g of seq.slice(1)) o.u16(g)
      }),
    )
    return withParts(2 + mine.length * 2, bodies, (o, at) => {
      o.u16(mine.length)
      for (const a of at) o.u16(a)
    })
  })
  return withParts(6 + firsts.length * 2, [coverage(firsts), ...sets], (o, at) => {
    o.u16(1).u16(at[0]).u16(firsts.length)
    for (const a of at.slice(1)) o.u16(a)
  })
}

/**
 * GSUB 6 / GPOS 8, format 3: a rule for the glyphs `input` when `back` is before them (nearest first) and `ahead`
 * after them; `lookups`: [index in the input, lookup] applied when it matches (none: a rule that only stops others).
 */
export function chainContext(back: number[][], input: number[][], ahead: number[][], lookups: [number, number][] = []): number[] {
  const covs = [...back, ...input, ...ahead].map(coverage)
  const head = 2 + 2 + back.length * 2 + 2 + input.length * 2 + 2 + ahead.length * 2 + 2 + lookups.length * 4
  return withParts(head, covs, (o, at) => {
    o.u16(3)
    let k = 0
    for (const list of [back, input, ahead]) {
      o.u16(list.length)
      for (let i = 0; i < list.length; i++) o.u16(at[k++])
    }
    o.u16(lookups.length)
    for (const [seq, lookup] of lookups) o.u16(seq).u16(lookup)
  })
}

/** GPOS 2, format 1: `pairs` are [left glyph, right glyph, how much the left one's advance changes] */
export function pairPos(pairs: [number, number, number][]): number[] {
  const firsts = [...new Set(pairs.map((p) => p[0]))].sort((a, b) => a - b)
  const sets = firsts.map((first) => {
    const mine = pairs.filter((p) => p[0] === first).sort((a, b) => a[1] - b[1])
    return table((o) => {
      o.u16(mine.length)
      for (const [, second, value] of mine) o.u16(second).i16(value)
    })
  })
  return withParts(10 + firsts.length * 2, [coverage(firsts), ...sets], (o, at) => {
    // value format: the first glyph's advance (0x0004), nothing for the second
    o.u16(1).u16(at[0]).u16(0x0004).u16(0).u16(firsts.length)
    for (const a of at.slice(1)) o.u16(a)
  })
}

/** GPOS 2, format 2: `left[i]` and `right[k]` are classes of glyphs, `values[i][k]` what the pair's left advance changes by */
export function classPairPos(left: number[][], right: number[][], values: number[][]): number[] {
  const count1 = left.length + 1
  const count2 = right.length + 1
  return withParts(16 + count1 * count2 * 2, [coverage(left.flat()), classDef(left), classDef(right)], (o, at) => {
    o.u16(2).u16(at[0]).u16(0x0004).u16(0).u16(at[1]).u16(at[2]).u16(count1).u16(count2)
    // class 0 (every other glyph) on either side: nothing
    for (let i = 0; i < count1; i++) for (let k = 0; k < count2; k++) o.i16(i && k ? (values[i - 1]?.[k - 1] ?? 0) : 0)
  })
}

export interface LookupSpec {
  /** GSUB: 1 single, 4 ligature, 6 chained context; GPOS: 2 pair, 8 chained context */
  type: number
  flag?: number
  subtables: number[][]
  /** written through extension subtables (GSUB 7, GPOS 9) */
  extension?: boolean
}

/** A GSUB (`gpos` false) or GPOS table: `features` name lookups by their index; every feature is in the default language of `scripts`. */
export function layoutTable(gpos: boolean, features: [string, number[]][], lookups: LookupSpec[], scripts = ['DFLT', 'latn']): number[] {
  const feats = [...features].sort((a, b) => (a[0] < b[0] ? -1 : 1))
  const scriptTags = [...scripts].sort()
  const langSys = table((o) => {
    o.u16(0).u16(0xffff).u16(feats.length)
    feats.forEach((_, i) => o.u16(i))
  })
  const script = withParts(4, [langSys], (o, at) => o.u16(at[0]).u16(0))
  const scriptList = withParts(
    2 + scriptTags.length * 6,
    scriptTags.map(() => script),
    (o, at) => {
      o.u16(scriptTags.length)
      scriptTags.forEach((t, i) => o.tag(t).u16(at[i]))
    },
  )
  const featureList = withParts(
    2 + feats.length * 6,
    feats.map(([, ls]) =>
      table((o) => {
        o.u16(0).u16(ls.length)
        for (const l of ls) o.u16(l)
      }),
    ),
    (o, at) => {
      o.u16(feats.length)
      feats.forEach(([t], i) => o.tag(t).u16(at[i]))
    },
  )
  const ext = gpos ? 9 : 7
  const lookupList = withParts(
    2 + lookups.length * 2,
    lookups.map((l) => {
      // an extension subtable holds the type and a 32-bit offset to the real one, here put right after it
      const subs = l.extension ? l.subtables.map((s) => table((o) => o.u16(1).u16(l.type).u32(8).add(s))) : l.subtables
      return withParts(6 + subs.length * 2, subs, (o, at) => {
        o.u16(l.extension ? ext : l.type).u16(l.flag ?? 0).u16(subs.length)
        for (const a of at) o.u16(a)
      })
    }),
    (o, at) => {
      o.u16(lookups.length)
      for (const a of at) o.u16(a)
    },
  )
  return withParts(10, [scriptList, featureList, lookupList], (o, at) => o.u16(1).u16(0).u16(at[0]).u16(at[1]).u16(at[2]))
}

/** a kern table (version 0, one horizontal subtable of format 0): [left glyph, right glyph, value] */
export const kernTable = (pairs: [number, number, number][]) =>
  table((o) => {
    const sorted = [...pairs].sort((a, b) => a[0] - b[0] || a[1] - b[1])
    const n = sorted.length
    const pow = n ? Math.floor(Math.log2(n)) : 0
    o.u16(0).u16(1)
    o.u16(0).u16(14 + n * 6).u16(0x0001)
    o.u16(n).u16((1 << pow) * 6).u16(pow).u16((n - (1 << pow)) * 6)
    for (const [l, r, v] of sorted) o.u16(l).u16(r).i16(v)
  })

// ---------------------------------------------------------------- the font

export interface FontSpec {
  family?: string
  subfamily?: string
  /** units to the em (default 1000) */
  upem?: number
  /** how far the pen moves after each glyph; glyph 0 is the "missing glyph". Every glyph is a box of that width (an advance of 0: no outline) */
  advances: number[]
  /** character → glyph */
  chars: [number, number][]
  /** glyphs made of others, set side by side: glyph → its parts */
  composites?: Record<number, number[]>
  /** how many glyphs have a width of their own in hmtx (default all): the ones after them share the last */
  longMetrics?: number
  /** write the character map in format 12 as well (needed for characters past U+FFFF) */
  cmap12?: boolean
  kern?: number[]
  gsub?: number[]
  gpos?: number[]
  /** OS/2 weight (default 400) and italic */
  weight?: number
  italic?: boolean
  /** hhea ascender, descender (negative) and line gap; OS/2 typographic and Windows values (defaults: 800, -200, 0 for all) */
  hhea?: [number, number, number]
  typo?: [number, number, number]
  win?: [number, number]
  /** fsSelection bit 7 */
  useTypo?: boolean
  /** OS/2 fsType: 2 forbids embedding */
  fsType?: number
  /** other tables, as they are (an `fvar` to make the font a variable one, a `trak`…) */
  extra?: Record<string, number[]>
  /** PostScript outlines (a CFF table) instead of TrueType ones; 'cid': a CID-keyed one */
  cff?: boolean | 'cid'
}

function checksum(bytes: number[]): number {
  let sum = 0
  for (let i = 0; i < bytes.length; i += 4) sum = (sum + (((bytes[i] ?? 0) << 24) | ((bytes[i + 1] ?? 0) << 16) | ((bytes[i + 2] ?? 0) << 8) | (bytes[i + 3] ?? 0))) >>> 0
  return sum
}

/** a font file of these tables */
export function sfnt(tables: [string, number[]][], cff = false): Uint8Array<ArrayBuffer> {
  const sorted = [...tables].sort((a, b) => (a[0] < b[0] ? -1 : 1))
  const n = sorted.length
  const sel = Math.floor(Math.log2(n))
  const out = new Bytes()
  out.u32(cff ? 0x4f54544f : 0x00010000).u16(n).u16(16 << sel).u16(sel).u16(n * 16 - (16 << sel))
  let offset = 12 + 16 * n
  let headAt = -1
  for (const [tag, bytes] of sorted) {
    if (tag === 'head') headAt = offset + 8
    out.tag(tag).u32(checksum(bytes)).u32(offset).u32(bytes.length)
    offset += (bytes.length + 3) & ~3
  }
  for (const [, bytes] of sorted) out.add(bytes).zeros(((bytes.length + 3) & ~3) - bytes.length)
  if (headAt >= 0) {
    const adjust = (0xb1b0afba - checksum(out.b)) >>> 0
    for (let k = 0; k < 4; k++) out.b[headAt + k] = (adjust >>> (24 - 8 * k)) & 0xff
  }
  return Uint8Array.from(out.b)
}

const ASCENT = 800
const DESCENT = 200

/** a Type 2 charstring number */
const t2 = (o: Bytes, v: number) => (v >= -107 && v <= 107 ? o.u8(v + 139) : o.u8(28).i16(v))
/** a CFF DICT number, always five bytes */
const d5 = (o: Bytes, v: number) => o.u8(29).u32(v >>> 0)

function cffIndex(items: number[][]): number[] {
  return table((o) => {
    o.u16(items.length)
    if (!items.length) return
    const total = items.reduce((a, i) => a + i.length, 1)
    const size = total < 0x100 ? 1 : 2
    o.u8(size)
    let at = 1
    const off = (x: number) => (size === 1 ? o.u8(x) : o.u16(x))
    off(at)
    for (const it of items) off((at += it.length))
    for (const it of items) o.add(it)
  })
}

/** a CFF table of box glyphs: glyph n is named by the standard string n (or is CID n); every box is drawn through a local subroutine */
function cffTable(advances: number[], cid: boolean): number[] {
  const n = advances.length
  const charstrings = advances.map((w) => {
    const o = new Bytes()
    t2(o, w)
    if (w > 100) {
      t2(o, 50)
      t2(o, 0)
      o.u8(21) // rmoveto
      t2(o, w - 100)
      o.u8(6) // hlineto
      t2(o, 700)
      o.u8(7) // vlineto
      t2(o, -(w - 100))
      o.u8(6)
      // the first subroutine, which does nothing but return (number 0 is called as -107)
      t2(o, -107)
      o.u8(10)
    }
    o.u8(14) // endchar
    return o.b
  })
  const header = [1, 0, 4, 1]
  const name = cffIndex([[...'TestCFF'].map((c) => c.charCodeAt(0))])
  const stringIndex = cffIndex(cid ? [[...'Adobe'].map((c) => c.charCodeAt(0)), [...'Identity'].map((c) => c.charCodeAt(0))] : [])
  const gsubrs = cffIndex([])
  const charset = table((o) => {
    o.u8(0)
    for (let g = 1; g < n; g++) o.u16(g)
  })
  const strs = cffIndex(charstrings)
  const subrs = cffIndex([[11]])
  // Private: defaultWidthX 0, nominalWidthX 0, Subrs right after it
  const priv = (size: number) =>
    table((o) => {
      t2(o, 0).u8(20)
      t2(o, 0).u8(21)
      d5(o, size).u8(19)
    })
  const privSize = priv(0).length
  const privBytes = [...priv(privSize), ...subrs]
  const fdSelect = table((o) => {
    // format 3: one range, all glyphs in font dictionary 0
    o.u8(3).u16(1).u16(0).u8(0).u16(n)
  })
  const top = (at: { charset: number; strs: number; priv: number; fdArray: number; fdSelect: number }) =>
    table((o) => {
      if (cid) {
        // ROS: strings 391 and 392 are the first two of the font's own (after the 391 standard ones)
        t2(o, 391)
        t2(o, 392)
        t2(o, 0).u8(12).u8(30)
        d5(o, n).u8(12).u8(34)
      }
      d5(o, at.charset).u8(15)
      d5(o, at.strs).u8(17)
      if (cid) {
        d5(o, at.fdArray).u8(12).u8(36)
        d5(o, at.fdSelect).u8(12).u8(37)
      } else {
        d5(o, privSize)
        d5(o, at.priv).u8(18)
      }
    })
  const fdDict = (at: number) => cffIndex([table((o) => (d5(o, privSize), d5(o, at).u8(18)))])
  const topSize = cffIndex([top({ charset: 0, strs: 0, priv: 0, fdArray: 0, fdSelect: 0 })]).length
  let p = header.length + name.length + topSize + stringIndex.length + gsubrs.length
  const at = { charset: p, strs: 0, priv: 0, fdArray: 0, fdSelect: 0 }
  p += charset.length
  if (cid) {
    at.fdSelect = p
    p += fdSelect.length
  }
  at.strs = p
  p += strs.length
  if (cid) {
    at.fdArray = p
    p += fdDict(0).length
  }
  at.priv = p
  return [...header, ...name, ...cffIndex([top(at)]), ...stringIndex, ...gsubrs, ...charset, ...(cid ? fdSelect : []), ...strs, ...(cid ? fdDict(at.priv) : []), ...privBytes]
}

/** the font's bytes (a TTF, or an OTF with `cff`) */
export function buildFont(spec: FontSpec): Uint8Array<ArrayBuffer> {
  const upem = spec.upem ?? 1000
  const family = spec.family ?? 'Test Read'
  const n = spec.advances.length
  const [asc, desc, gap] = spec.hhea ?? [ASCENT, -DESCENT, 0]
  const [tAsc, tDesc, tGap] = spec.typo ?? [ASCENT, -DESCENT, 0]
  const [wAsc, wDesc] = spec.win ?? [ASCENT, DESCENT]
  const box = (w: number) =>
    table((o) => {
      o.i16(1).i16(50).i16(0).i16(w - 50).i16(700)
      o.u16(3).u16(0)
      for (let i = 0; i < 4; i++) o.u8(1)
      for (const dx of [50, 0, w - 100, 0]) o.i16(dx)
      for (const dy of [0, 700, 0, -700]) o.i16(dy)
      o.zeros(2)
    })
  const glyphs = spec.advances.map((w, g) => {
    const parts = spec.composites?.[g]
    if (!parts) return w > 100 ? box(w) : []
    return table((o) => {
      o.i16(-1).i16(50).i16(0).i16(w - 50).i16(700)
      let x = 0
      parts.forEach((part, i) => {
        // ARG_1_AND_2_ARE_WORDS, ARGS_ARE_XY_VALUES, MORE_COMPONENTS on all but the last
        o.u16(0x0003 | (i < parts.length - 1 ? 0x0020 : 0)).u16(part).i16(x).i16(0)
        x += spec.advances[part]
      })
      if (o.b.length & 1) o.u8(0)
    })
  })
  const loca = table((o) => {
    let at = 0
    o.u32(0)
    for (const g of glyphs) o.u32((at += g.length))
  })
  const widest = Math.max(...spec.advances)
  const head = table((o) => {
    o.u16(1).u16(0).u32(0x00010000).u32(0).u32(0x5f0f3cf5).u16(0x0003).u16(upem)
    o.zeros(16)
    o.i16(0).i16(-DESCENT).i16(widest).i16(ASCENT)
    o.u16((spec.weight ?? 400) >= 600 ? 1 : 0).u16(8).i16(2).i16(1).i16(0)
  })
  const long = Math.min(n, spec.longMetrics ?? n)
  const hhea = table((o) => {
    o.u16(1).u16(0).i16(asc).i16(desc).i16(gap).u16(widest)
    o.i16(0).i16(0).i16(widest).i16(1).i16(0).i16(0)
    o.zeros(8).i16(0).u16(long)
  })
  const maxp = spec.cff
    ? table((o) => o.u32(0x00005000).u16(n))
    : table((o) => {
        o.u32(0x00010000).u16(n).u16(4).u16(1).u16(8).u16(2).u16(2)
        o.zeros(12).u16(4).u16(1)
      })
  const codes = spec.chars.map((c) => c[0])
  const os2 = table((o) => {
    o.u16(4).i16(500).u16(spec.weight ?? 400).u16(5).u16(spec.fsType ?? 0)
    o.i16(650).i16(600).i16(0).i16(75).i16(650).i16(600).i16(0).i16(350).i16(50).i16(300).i16(0)
    o.zeros(10)
    o.u32(1).u32(0).u32(0).u32(0)
    o.tag('NONE').u16((spec.italic ? 0x0001 : (spec.weight ?? 400) >= 600 ? 0x0020 : 0x0040) | (spec.useTypo ? 0x0080 : 0))
    o.u16(Math.min(0xffff, Math.min(...codes))).u16(Math.min(0xffff, Math.max(...codes)))
    o.i16(tAsc).i16(tDesc).i16(tGap).u16(wAsc).u16(wDesc)
    o.u32(1).u32(0)
    o.i16(500).i16(700).u16(0).u16(0x20).u16(1)
  })
  const hmtx = table((o) => {
    for (let g = 0; g < long; g++) o.u16(spec.advances[g]).i16(spec.advances[g] > 100 ? 50 : 0)
    for (let g = long; g < n; g++) o.i16(50)
  })
  // format 4: one segment for each run of characters whose glyphs follow each other
  const bmp = [...spec.chars].filter((c) => c[0] < 0xffff).sort((a, b) => a[0] - b[0])
  const segs: { start: number; end: number; glyph: number }[] = []
  for (const [cp, g] of bmp) {
    const last = segs[segs.length - 1]
    if (last && cp === last.end + 1 && g === last.glyph + (cp - last.start)) last.end = cp
    else segs.push({ start: cp, end: cp, glyph: g })
  }
  segs.push({ start: 0xffff, end: 0xffff, glyph: 0 })
  const f4 = table((o) => {
    const c = segs.length
    const pow = Math.floor(Math.log2(c))
    o.u16(4).u16(16 + c * 8).u16(0).u16(c * 2).u16((1 << pow) * 2).u16(pow).u16(c * 2 - (1 << pow) * 2)
    for (const s of segs) o.u16(s.end)
    o.u16(0)
    for (const s of segs) o.u16(s.start)
    segs.forEach((s, i) => o.u16(i === c - 1 ? 1 : (s.glyph - s.start) & 0xffff))
    for (let i = 0; i < c; i++) o.u16(0)
  })
  const all = [...spec.chars].sort((a, b) => a[0] - b[0])
  const f12 = table((o) => {
    o.u16(12).u16(0).u32(16 + all.length * 12).u32(0).u32(all.length)
    for (const [cp, g] of all) o.u32(cp).u32(cp).u32(g)
  })
  const cmap = table((o) => {
    o.u16(0).u16(spec.cmap12 ? 2 : 1).u16(3).u16(1).u32(spec.cmap12 ? 20 : 12)
    if (spec.cmap12) o.u16(3).u16(10).u32(20 + f4.length)
    o.add(f4)
    if (spec.cmap12) o.add(f12)
  })
  const ps = family.replace(/[^A-Za-z0-9-]/g, '') || 'TestFont'
  const names: [number, string][] = [
    [1, family],
    [2, spec.subfamily ?? 'Regular'],
    [3, `${ps}-test`],
    [4, `${family} ${spec.subfamily ?? 'Regular'}`],
    [5, 'Version 1.000'],
    [6, `${ps}-${(spec.subfamily ?? 'Regular').replace(/\s/g, '')}`],
  ]
  const name = table((o) => {
    o.u16(0).u16(names.length).u16(6 + 12 * names.length)
    let at = 0
    for (const [id, text] of names) {
      o.u16(3).u16(1).u16(0x0409).u16(id).u16(text.length * 2).u16(at)
      at += text.length * 2
    }
    for (const [, text] of names) for (const c of text) o.u16(c.charCodeAt(0))
  })
  const post = table((o) => {
    o.u32(0x00030000).u32(spec.italic ? (-12 << 16) >>> 0 : 0).i16(-100).i16(50).u32(0)
    o.zeros(16)
  })
  const tables: [string, number[]][] = [
    ['OS/2', os2],
    ['cmap', cmap],
    ['head', head],
    ['hhea', hhea],
    ['hmtx', hmtx],
    ['maxp', maxp],
    ['name', name],
    ['post', post],
  ]
  if (spec.cff) tables.push(['CFF ', cffTable(spec.advances, spec.cff === 'cid')])
  else tables.push(['glyf', glyphs.flat()], ['loca', loca])
  if (spec.kern) tables.push(['kern', spec.kern])
  if (spec.gsub) tables.push(['GSUB', spec.gsub])
  if (spec.gpos) tables.push(['GPOS', spec.gpos])
  for (const [tag, bytes] of Object.entries(spec.extra ?? {})) tables.push([tag, bytes])
  return sfnt(tables, !!spec.cff)
}

/** a collection (TTC) of these font files, each with its own copy of every table */
export function collection(fonts: Uint8Array[]): Uint8Array<ArrayBuffer> {
  const out = new Bytes()
  out.tag('ttcf').u32(0x00010000).u32(fonts.length)
  let at = 12 + fonts.length * 4
  const starts: number[] = []
  for (const f of fonts) {
    starts.push(at)
    out.u32(at)
    at += (f.length + 3) & ~3
  }
  fonts.forEach((f, i) => {
    const copy = [...f]
    // the tables' offsets are from the start of the file
    const n = (copy[4] << 8) | copy[5]
    for (let k = 0; k < n; k++) {
      const rec = 12 + k * 16 + 8
      const off = ((copy[rec] << 24) | (copy[rec + 1] << 16) | (copy[rec + 2] << 8) | copy[rec + 3]) + starts[i]
      for (let b = 0; b < 4; b++) copy[rec + b] = (off >>> (24 - b * 8)) & 0xff
    }
    out.add(copy).zeros(((f.length + 3) & ~3) - f.length)
  })
  return Uint8Array.from(out.b)
}

async function pack(data: Uint8Array, format: string): Promise<Uint8Array> {
  const stream = new Blob([data as BlobPart]).stream().pipeThrough(new CompressionStream(format as CompressionFormat) as unknown as TransformStream<Uint8Array, Uint8Array>)
  return new Uint8Array(await new Response(stream).arrayBuffer())
}

/** the font file as WOFF: every table deflated (or as it is, when that is no shorter) */
export async function woff(font: Uint8Array): Promise<Uint8Array<ArrayBuffer>> {
  const v = new DataView(font.buffer, font.byteOffset, font.byteLength)
  const n = v.getUint16(4)
  const out = new Bytes()
  const bodies: { tag: number[]; data: Uint8Array; orig: number; sum: number }[] = []
  for (let i = 0; i < n; i++) {
    const rec = 12 + i * 16
    const data = font.subarray(v.getUint32(rec + 8), v.getUint32(rec + 8) + v.getUint32(rec + 12))
    const packed = await pack(data, 'deflate')
    bodies.push({ tag: [...font.subarray(rec, rec + 4)], data: packed.length < data.length ? packed : data, orig: data.length, sum: v.getUint32(rec + 4) })
  }
  let at = 44 + n * 20
  const total = bodies.reduce((a, b) => a + ((b.data.length + 3) & ~3), at)
  out.tag('wOFF').u32(v.getUint32(0)).u32(total).u16(n).u16(0).u32(font.length).u16(1).u16(0).zeros(20)
  for (const b of bodies) {
    out.add(b.tag).u32(at).u32(b.data.length).u32(b.orig).u32(b.sum)
    at += (b.data.length + 3) & ~3
  }
  for (const b of bodies) out.add(b.data).zeros(((b.data.length + 3) & ~3) - b.data.length)
  return Uint8Array.from(out.b)
}
