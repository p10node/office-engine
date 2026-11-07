/**
 * Writes small iWork '13+ packages (Pages) for tests and the playground sample: protobuf messages by field
 * number, IWA framing with literal-only Snappy blocks, and the package ZIP.
 */
import { deflateSync } from 'node:zlib'
import { crc32, writeZip, type PartData } from '../src/core/zipwrite'

// ---------------------------------------------------------------- protobuf encoding

export class F32 {
  constructor(readonly v: number) {}
}
export class Fixed32 {
  constructor(readonly v: number) {}
}
export class Raw {
  constructor(readonly bytes: Uint8Array) {}
}
export class M {
  constructor(readonly f: Record<number, V | undefined>) {}
}
export type V = number | boolean | string | Uint8Array | F32 | Fixed32 | Raw | M | V[]

export const msg = (f: Record<number, V | undefined>) => new M(f)
export const f32 = (v: number) => new F32(v)
export const fixed32 = (v: number) => new Fixed32(v)
export const ref = (id: number) => msg({ 1: id })

class Out {
  private parts: number[] = []
  varint(v: number) {
    if (v < 0) {
      // ten-byte two's complement
      let lo = v >>> 0
      let hi = Math.floor(v / 0x100000000) >>> 0
      for (let k = 0; k < 9; k++) {
        this.parts.push((lo & 0x7f) | 0x80)
        lo = (lo >>> 7) | ((hi & 0x7f) << 25)
        hi >>>= 7
      }
      this.parts.push(1)
      return
    }
    while (v >= 0x80) {
      this.parts.push((v % 0x80) | 0x80)
      v = Math.floor(v / 0x80)
    }
    this.parts.push(v)
  }
  bytes(b: Uint8Array) {
    for (const x of b) this.parts.push(x)
  }
  done(): Uint8Array {
    return Uint8Array.from(this.parts)
  }
}

export function encode(m: M): Uint8Array {
  const o = new Out()
  const field = (f: number, v: V) => {
    if (Array.isArray(v)) {
      for (const x of v) field(f, x)
      return
    }
    if (typeof v === 'number' || typeof v === 'boolean') {
      o.varint(f * 8)
      o.varint(typeof v === 'boolean' ? +v : v)
    } else if (v instanceof F32) {
      o.varint(f * 8 + 5)
      const b = new Uint8Array(4)
      new DataView(b.buffer).setFloat32(0, v.v, true)
      o.bytes(b)
    } else if (v instanceof Fixed32) {
      o.varint(f * 8 + 5)
      const b = new Uint8Array(4)
      new DataView(b.buffer).setUint32(0, v.v, true)
      o.bytes(b)
    } else {
      const b = typeof v === 'string' ? new TextEncoder().encode(v) : v instanceof M ? encode(v) : v instanceof Raw ? v.bytes : v
      o.varint(f * 8 + 2)
      o.varint(b.length)
      o.bytes(b)
    }
  }
  for (const [k, v] of Object.entries(m.f)) if (v !== undefined) field(Number(k), v)
  return o.done()
}

// ---------------------------------------------------------------- IWA

/** a Snappy block made of literals only (valid, uncompressed) */
export function snappyLiteral(data: Uint8Array): Uint8Array {
  const o = new Out()
  o.varint(data.length)
  for (let p = 0; p < data.length; p += 0x10000) {
    const n = Math.min(0x10000, data.length - p)
    if (n <= 60) o.bytes(Uint8Array.of((n - 1) << 2))
    else o.bytes(Uint8Array.of(62 << 2, (n - 1) & 0xff, ((n - 1) >> 8) & 0xff, ((n - 1) >> 16) & 0xff))
    o.bytes(data.subarray(p, p + n))
  }
  return o.done()
}

export function iwa(objects: { id: number; type: number; body: Uint8Array }[]): Uint8Array {
  const o = new Out()
  for (const { id, type, body } of objects) {
    const info = encode(msg({ 1: id, 2: msg({ 1: type, 2: [1, 0, 5], 3: body.length }) }))
    o.varint(info.length)
    o.bytes(info)
    o.bytes(body)
  }
  // chunks of at most 64 KB of data, each its own Snappy block, as iWork writes them
  const data = o.done()
  const chunks: Uint8Array[] = []
  for (let p = 0; p < data.length || !chunks.length; p += 0x10000) {
    const block = snappyLiteral(data.subarray(p, p + 0x10000))
    const chunk = new Uint8Array(4 + block.length)
    chunk[1] = block.length & 0xff
    chunk[2] = (block.length >> 8) & 0xff
    chunk[3] = (block.length >> 16) & 0xff
    chunk.set(block, 4)
    chunks.push(chunk)
  }
  const out = new Uint8Array(chunks.reduce((n, c) => n + c.length, 0))
  let q = 0
  for (const c of chunks) out.set(c, q), (q += c.length)
  return out
}

// ---------------------------------------------------------------- package

export class IworkBuilder {
  private objects: { id: number; type: number; body: Uint8Array }[] = []
  private datas: { id: number; name: string; bytes: Uint8Array }[] = []
  private next = 100
  extra = new Map<string, PartData>()

  add(type: number, m: M, id = this.next++): number {
    this.objects.push({ id, type, body: encode(m) })
    return id
  }

  /** a file in Data/ (image…): returns its data id */
  data(name: string, bytes: Uint8Array): number {
    const id = this.next++
    this.datas.push({ id, name, bytes })
    return id
  }

  async zip(): Promise<Uint8Array<ArrayBuffer>> {
    const files = new Map<string, PartData>()
    const fileName = (d: { id: number; name: string }) => d.name.replace(/(\.[^.]+)?$/, `-${d.id}$1`)
    // TSP.PackageMetadata: datas { identifier, preferred_file_name, file_name }
    const meta = msg({ 3: this.datas.map((d) => msg({ 1: d.id, 3: d.name, 4: fileName(d) })) })
    files.set('Index/Document.iwa', iwa(this.objects))
    files.set('Index/Metadata.iwa', iwa([{ id: 2, type: 11006, body: encode(meta) }]))
    for (const d of this.datas) files.set('Data/' + fileName(d), d.bytes)
    files.set('Metadata/DocumentIdentifier', '00000000-0000-4000-8000-000000000000')
    for (const [k, v] of this.extra) files.set(k, v)
    return new Uint8Array(await (await writeZip(files)).arrayBuffer())
  }
}

// ---------------------------------------------------------------- Pages content

export const T = {
  STORAGE: 2001,
  DRAWABLE_ATTACHMENT: 2003,
  FOOTNOTE_REF: 2008,
  SHAPE_INFO: 2011,
  CHAR_STYLE: 2021,
  PARA_STYLE: 2022,
  LIST_STYLE: 2023,
  HYPERLINK: 2032,
  NUMBER_ATTACHMENT: 2043,
  IMAGE: 3005,
  TABLE_INFO: 6000,
  TABLE_MODEL: 6001,
  TILE: 6002,
  DATA_LIST: 6005,
  HEADER_BUCKET: 6006,
  CELL_STYLE: 6004,
  DOCUMENT: 10000,
  FLOATING: 10010,
  SECTION: 10011,
  SETTINGS: 10012,
  PAGE_MASTER: 10143,
} as const

export const color = (hex: string) => {
  const n = parseInt(hex, 16)
  return msg({ 1: 1, 3: f32(((n >> 16) & 255) / 255), 4: f32(((n >> 8) & 255) / 255), 5: f32((n & 255) / 255), 6: f32(1) })
}

export interface CharSpec {
  bold?: boolean
  italic?: boolean
  size?: number
  font?: string
  color?: string
  underline?: number
}

export const charProps = (c: CharSpec) =>
  msg({ 1: c.bold, 2: c.italic, 3: c.size === undefined ? undefined : f32(c.size), 5: c.font, 7: c.color ? color(c.color) : undefined, 11: c.underline })

export interface ParaSpec {
  align?: number
  before?: number
  after?: number
  left?: number
  first?: number
  outline?: number
  list?: number
}

export const paraProps = (p: ParaSpec) =>
  msg({
    1: p.align,
    7: p.first === undefined ? undefined : f32(p.first),
    11: p.left === undefined ? undefined : f32(p.left),
    20: p.after === undefined ? undefined : f32(p.after),
    21: p.before === undefined ? undefined : f32(p.before),
    27: p.outline,
    40: p.list === undefined ? undefined : ref(p.list),
  })

/** TSS.StyleArchive: name (named style) or parent + variation flag */
export const styleInfo = (name: string | undefined, parent?: number) =>
  msg({ 1: name, 3: parent === undefined ? undefined : ref(parent), 4: name === undefined ? true : undefined })

export type Table = [number, number | undefined][]
const objTable = (t: Table | undefined) =>
  t && msg({ 1: t.map(([i, id]) => msg({ 1: i, 2: id === undefined ? undefined : ref(id) })) })

export interface StorageSpec {
  kind?: number
  text: string
  paras?: Table
  lists?: Table
  levels?: [number, number][]
  chars?: Table
  attachments?: Table
  links?: Table
  footnotes?: Table
}

export function storage(b: IworkBuilder, s: StorageSpec): number {
  return b.add(
    T.STORAGE,
    msg({
      1: s.kind ?? 0,
      3: s.text,
      5: objTable(s.paras),
      6: s.levels && msg({ 1: s.levels.map(([i, l]) => msg({ 1: i, 2: l, 3: 0 })) }),
      7: objTable(s.lists),
      8: objTable(s.chars),
      9: objTable(s.attachments),
      11: objTable(s.links),
      16: objTable(s.footnotes),
    }),
  )
}

/** TSD.DrawableArchive with its geometry (flags 3: width and height set) */
export const drawable = (w: number, h: number, x = 0, y = 0, extra: { flags?: number; angle?: number } = {}) =>
  msg({ 1: msg({ 1: msg({ 1: f32(x), 2: f32(y) }), 2: msg({ 1: f32(w), 2: f32(h) }), 3: extra.flags ?? 3, 4: f32(extra.angle ?? 0) }) })
const geometry = (w: number, h: number) => drawable(w, h)

/** inline image: the DrawableAttachment to put at a U+FFFC of the text */
export function image(b: IworkBuilder, name: string, bytes: Uint8Array, w: number, h: number): number {
  const data = b.data(name, bytes)
  const img = b.add(T.IMAGE, msg({ 1: geometry(w, h), 11: msg({ 1: data }), 4: msg({ 1: f32(w), 2: f32(h) }) }))
  return b.add(T.DRAWABLE_ATTACHMENT, msg({ 1: ref(img) }))
}

export interface CellSpec {
  text?: string
  num?: number
  /** seconds from 2001-01-01 UTC */
  date?: number
  bool?: boolean
  /** seconds */
  duration?: number
  /** TSK.FormatStructArchive (current storage only) */
  fmt?: M
}

/** TSK.FormatStructArchive helpers */
export const fmt = {
  currency: (code: string, decimals = 2) => msg({ 1: 257, 2: decimals, 3: code, 4: 2, 5: true }),
  percent: (decimals = 1) => msg({ 1: 258, 2: decimals }),
  decimal: (decimals: number, sep = true) => msg({ 1: 256, 2: decimals, 5: sep }),
  date: (icu: string) => msg({ 1: 261, 14: icu }),
  duration: (style: number) => msg({ 1: 268, 7: style, 15: 4, 16: 8 }),
  checkbox: () => msg({ 1: 263 }),
}

/** a table: cells in the pre-BNC (iWork '13–'14, version 3) or current (version 5) storage */
export function table(
  b: IworkBuilder,
  rows: CellSpec[][],
  opts: Parameters<typeof tableInfo>[2] = {},
): number {
  return b.add(T.DRAWABLE_ATTACHMENT, msg({ 1: ref(tableInfo(b, rows, opts)) }))
}

/** a date, boolean or duration cell in the current storage (version 5), with its format key */
function typedCell(cell: CellSpec, formats: M[]): Uint8Array {
  const [type, value] = cell.date !== undefined ? [5, cell.date] : cell.bool !== undefined ? [6, +cell.bool] : [7, cell.duration!]
  const fbit = type === 5 ? 0x8000 : type === 7 ? 0x10000 : 0x40000
  const fkey = cell.fmt ? formats.push(cell.fmt) : 0
  const out = new Uint8Array(12 + 8 + (fkey ? 4 : 0))
  const dv = new DataView(out.buffer)
  out[0] = 5
  out[1] = type
  dv.setUint32(8, (type === 5 ? 0x4 : 0x2) | (fkey ? fbit : 0), true)
  dv.setFloat64(12, value, true)
  if (fkey) dv.setUint32(20, fkey, true)
  return out
}

/** a table drawable (TST.TableInfoArchive); `frame` places it on a slide */
export function tableInfo(
  b: IworkBuilder,
  rows: CellSpec[][],
  opts: {
    headerRows?: number
    colWidths?: number[]
    format?: 'v5' | 'v3'
    headerFill?: string
    merges?: [number, number, number, number][]
    bodyStyle?: number
    headerStyle?: number
    /** table style with banded rows of this fill; `sidecar`: an empty stroke sidecar (a table without lines) */
    band?: string
    sidecar?: boolean
    frame?: [number, number]
    name?: string
    /** header rows stay in view */
    freeze?: boolean
  } = {},
): number {
  const nr = rows.length
  const nc = Math.max(...rows.map((r) => r.length))
  /** cell formats, keyed from 1 */
  const formats: M[] = []
  const strings: string[] = []
  const rowInfos: M[] = []
  rows.forEach((row, r) => {
    const cells: Uint8Array[] = []
    for (let c = 0; c < nc; c++) {
      const cell = row[c]
      if (!cell || (cell.text === undefined && cell.num === undefined && cell.date === undefined && cell.bool === undefined && cell.duration === undefined)) {
        cells.push(new Uint8Array(0))
        continue
      }
      if (opts.format !== 'v3' && cell.text === undefined && cell.num === undefined) {
        cells.push(typedCell(cell, formats))
        continue
      }
      const v3 = opts.format === 'v3'
      const isText = cell.text !== undefined
      let key = 0
      if (isText) {
        strings.push(cell.text!)
        key = strings.length
      }
      const head = new Uint8Array(12)
      const dv = new DataView(head.buffer)
      let tail: Uint8Array
      if (v3) {
        dv.setUint16(0, 3, true)
        head[2] = isText ? 3 : 2
        dv.setUint16(4, isText ? 0x10 : 0x20, true)
        tail = new Uint8Array(isText ? 4 : 8)
        if (isText) new DataView(tail.buffer).setUint32(0, key, true)
        else new DataView(tail.buffer).setFloat64(0, cell.num!, true)
      } else {
        head[0] = 5
        head[1] = isText ? 3 : cell.fmt?.f[1] === 257 ? 10 : 2
        const fkey = cell.fmt ? formats.push(cell.fmt) : 0
        const fbit = cell.fmt?.f[1] === 257 ? 0x4000 : 0x2000
        dv.setUint32(8, (isText ? 0x8 : 0x2) | (fkey ? fbit : 0), true)
        tail = new Uint8Array((isText ? 4 : 8) + (fkey ? 4 : 0))
        const td = new DataView(tail.buffer)
        if (isText) td.setUint32(0, key, true)
        else td.setFloat64(0, cell.num!, true)
        if (fkey) td.setUint32(isText ? 4 : 8, fkey, true)
      }
      const rec = new Uint8Array(head.length + tail.length)
      rec.set(head)
      rec.set(tail, head.length)
      cells.push(rec)
    }
    let size = 0
    const offs = new Uint8Array(nc * 2)
    const od = new DataView(offs.buffer)
    cells.forEach((c, k) => {
      od.setUint16(k * 2, c.length ? size : 0xffff, true)
      size += c.length
    })
    const buf = new Uint8Array(size)
    let p = 0
    for (const c of cells) buf.set(c, p), (p += c.length)
    rowInfos.push(
      opts.format === 'v3'
        ? msg({ 1: r, 2: nc, 3: buf, 4: offs, 5: 3 })
        : msg({ 1: r, 2: nc, 6: buf, 7: offs, 5: 5 }),
    )
  })
  const tile = b.add(T.TILE, msg({ 1: nc - 1, 2: nr - 1, 3: nr * nc, 4: nr, 5: rowInfos, 7: opts.format !== 'v3' }))
  const stringList = b.add(T.DATA_LIST, msg({ 1: 1, 2: strings.length + 1, 3: strings.map((s, k) => msg({ 1: k + 1, 2: 1, 3: s })) }))
  const formatList = formats.length ? b.add(T.DATA_LIST, msg({ 1: 3, 2: formats.length + 1, 3: formats.map((f, k) => msg({ 1: k + 1, 2: 1, 6: f })) })) : undefined
  const widths = opts.colWidths ?? Array(nc).fill(120)
  const colBucket = b.add(T.HEADER_BUCKET, msg({ 1: 1, 2: widths.map((w, k) => msg({ 1: k, 2: f32(w), 3: 0, 4: 0 })) }))
  const headerStyle = opts.headerFill ? b.add(T.CELL_STYLE, msg({ 1: styleInfo('Header Row'), 11: msg({ 1: msg({ 1: color(opts.headerFill) }) }) })) : undefined
  const mergeMap =
    opts.merges &&
    b.add(
      6144,
      msg({ 1: opts.merges.map(([r, c, rs, cs]) => msg({ 1: msg({ 1: fixed32((c << 16) | r) }), 2: msg({ 1: fixed32((cs << 16) | rs) }) })) }),
    )
  const store = msg({
    1: msg({ 1: 1 }),
    2: ref(colBucket),
    3: msg({ 1: [msg({ 1: 0, 2: ref(tile) })], 2: 256 }),
    4: ref(stringList),
    13: mergeMap === undefined ? undefined : ref(mergeMap),
    22: formatList === undefined ? undefined : ref(formatList),
  })
  const model = b.add(
    T.TABLE_MODEL,
    msg({
      1: 'table-1',
      4: store,
      6: nr,
      7: nc,
      8: opts.name ?? 'Table 1',
      12: opts.freeze,
      9: opts.headerRows ?? 0,
      19: headerStyle === undefined ? undefined : ref(headerStyle),
      24: opts.bodyStyle === undefined ? undefined : ref(opts.bodyStyle),
      25: opts.headerStyle === undefined ? undefined : ref(opts.headerStyle),
      3: opts.band ? ref(b.add(6003, msg({ 1: styleInfo('Table Style'), 11: msg({ 1: true, 2: msg({ 1: color(opts.band) }) }) }))) : undefined,
      49: opts.sidecar ? ref(b.add(6305, msg({ 1: 0, 2: nc, 3: nr }))) : undefined,
    }),
  )
  return b.add(T.TABLE_INFO, msg({ 1: drawable(widths.reduce((x, y) => x + y, 0), nr * (opts.frame ? 50 : 22), opts.frame?.[0], opts.frame?.[1]), 2: ref(model) }))
}

// ---------------------------------------------------------------- images

/** a horizontal-gradient PNG */
export function png(w: number, h: number, from = [37, 99, 235], to = [16, 185, 129]): Uint8Array {
  const raw = new Uint8Array((w * 3 + 1) * h)
  for (let y = 0; y < h; y++) {
    const row = y * (w * 3 + 1)
    for (let x = 0; x < w; x++) {
      const t = x / Math.max(1, w - 1)
      for (let k = 0; k < 3; k++) raw[row + 1 + x * 3 + k] = Math.round(from[k] + (to[k] - from[k]) * t)
    }
  }
  const chunk = (type: string, data: Uint8Array) => {
    const out = new Uint8Array(12 + data.length)
    const dv = new DataView(out.buffer)
    dv.setUint32(0, data.length)
    const td = new TextEncoder().encode(type)
    out.set(td, 4)
    out.set(data, 8)
    dv.setUint32(8 + data.length, crc32(out.subarray(4, 8 + data.length)) >>> 0)
    return out
  }
  const ihdr = new Uint8Array(13)
  const dv = new DataView(ihdr.buffer)
  dv.setUint32(0, w)
  dv.setUint32(4, h)
  ihdr.set([8, 2, 0, 0, 0], 8)
  const parts = [Uint8Array.of(137, 80, 78, 71, 13, 10, 26, 10), chunk('IHDR', ihdr), chunk('IDAT', deflateSync(raw)), chunk('IEND', new Uint8Array(0))]
  const out = new Uint8Array(parts.reduce((a, p) => a + p.length, 0))
  let p = 0
  for (const x of parts) out.set(x, p), (p += x.length)
  return out
}

// ---------------------------------------------------------------- sample document

/** The demo / test document: styles, inline formatting, a link, lists, an image, a table, header, footer and a footnote. */
export async function samplePages(): Promise<Uint8Array<ArrayBuffer>> {
  const b = new IworkBuilder()
  const sans = 'HelveticaNeue'
  // styles
  const body = b.add(T.PARA_STYLE, msg({ 1: styleInfo('Body'), 11: charProps({ size: 11, font: sans, color: '1F2937', bold: false, italic: false }), 12: paraProps({ after: 8 }) }))
  const title = b.add(T.PARA_STYLE, msg({ 1: styleInfo('Title', body), 11: charProps({ size: 26, font: 'HelveticaNeue-Bold', color: '111827' }), 12: paraProps({ after: 4, outline: 0 }) }))
  const subtitle = b.add(T.PARA_STYLE, msg({ 1: styleInfo('Subtitle', body), 11: charProps({ size: 13, color: '6B7280' }), 12: paraProps({ after: 18 }) }))
  const heading = b.add(T.PARA_STYLE, msg({ 1: styleInfo('Heading', body), 11: charProps({ size: 16, bold: true, color: '1D4ED8' }), 12: paraProps({ before: 14, after: 6, outline: 0 }) }))
  const caption = b.add(T.PARA_STYLE, msg({ 1: styleInfo('Caption', body), 11: charProps({ size: 9, italic: true, color: '6B7280' }), 12: paraProps({ align: 2, after: 12 }) }))
  const centered = b.add(T.PARA_STYLE, msg({ 1: styleInfo(undefined, body), 12: paraProps({ align: 2 }) }))
  const hf = b.add(T.PARA_STYLE, msg({ 1: styleInfo('Header & Footer', body), 11: charProps({ size: 9, color: '6B7280' }), 12: paraProps({ after: 0 }) }))
  const note = b.add(T.PARA_STYLE, msg({ 1: styleInfo('Footnote Text', body), 11: charProps({ size: 9 }), 12: paraProps({ after: 0 }) }))
  const cell = b.add(T.PARA_STYLE, msg({ 1: styleInfo('Table Body', body), 11: charProps({ size: 10 }), 12: paraProps({ after: 0 }) }))
  const cellHead = b.add(T.PARA_STYLE, msg({ 1: styleInfo('Table Header', cell), 11: charProps({ bold: true, color: 'FFFFFF' }) }))
  const strong = b.add(T.CHAR_STYLE, msg({ 1: styleInfo('Strong'), 11: charProps({ bold: true }) }))
  const emphasis = b.add(T.CHAR_STYLE, msg({ 1: styleInfo('Emphasis'), 11: charProps({ italic: true, color: 'B45309' }) }))
  const linkStyle = b.add(T.CHAR_STYLE, msg({ 1: styleInfo('Link'), 11: charProps({ color: '2563EB', underline: 1 }) }))
  const levels = (n: number, f: (l: number) => V) => Array.from({ length: n }, (_, l) => f(l))
  const bullets = b.add(
    T.LIST_STYLE,
    msg({ 1: styleInfo('Bullet'), 11: levels(9, () => 2), 12: levels(9, () => f32(1.2)), 13: levels(9, (l) => f32(l * 18)), 16: levels(9, (l) => ['•', '◦', '▪'][l % 3]) }),
  )
  const numbered = b.add(T.LIST_STYLE, msg({ 1: styleInfo('Numbered'), 11: levels(9, () => 3), 12: levels(9, () => f32(1.6)), 13: levels(9, (l) => f32(l * 18)), 15: levels(9, (l) => (l % 2 ? 12 : 0)) }))
  const none = b.add(T.LIST_STYLE, msg({ 1: styleInfo('None'), 11: levels(9, () => 0) }))

  // body text
  const parts: { text: string; para: number; char?: number; list?: number; level?: number; link?: string; att?: number; note?: number }[] = []
  const para = (para: number, ...runs: (string | [string, number] | { link: string; text: string } | { att: number } | { note: number })[]) => {
    for (const r of runs) {
      if (typeof r === 'string') parts.push({ text: r, para })
      else if (Array.isArray(r)) parts.push({ text: r[0], para, char: r[1] })
      else if ('link' in r) parts.push({ text: r.text, para, char: linkStyle, link: r.link })
      else if ('att' in r) parts.push({ text: '\ufffc', para, att: r.att })
      else parts.push({ text: '\u000e', para, note: r.note })
    }
    parts.push({ text: '\n', para })
  }
  const listItem = (list: number, level: number, text: string) => {
    parts.push({ text, para: body, list, level })
    parts.push({ text: '\n', para: body, list, level })
  }
  const noteText = storage(b, { kind: 2, text: 'Measured on the 2024 survey sample (n = 1,204).', paras: [[0, note]] })
  const noteRef = b.add(T.FOOTNOTE_REF, msg({ 1: msg({}), 2: ref(noteText) }))
  const link = b.add(T.HYPERLINK, msg({ 1: msg({}), 2: 'https://example.com/report' }))

  para(title, 'Quarterly Field Report')
  para(subtitle, 'Pages document opened in the browser, view only')
  para(body, 'This report sums up the season. Text keeps its ', ['bold', strong], ' and ', ['italic, colored', emphasis], ' formatting, and links such as ', { link: 'https://example.com/report', text: 'the full report' }, ' stay clickable.')
  para(heading, 'Highlights')
  listItem(bullets, 0, 'Coverage grew in every region')
  listItem(bullets, 1, 'Northern sites reported first')
  listItem(bullets, 0, 'Survey response time fell by a third')
  listItem(numbered, 0, 'Collect the remaining samples')
  listItem(numbered, 0, 'Publish the regional summaries')
  para(centered, { att: image(b, 'chart.png', png(480, 160), 360, 120) })
  para(caption, 'Figure 1. Readings across the season')
  para(heading, 'Results')
  para(
    body,
    {
      att: table(
        b,
        [
          [{ text: 'Region' }, { text: 'Sites' }, { text: 'Readings' }],
          [{ text: 'North' }, { num: 12 }, { num: 1480.5 }],
          [{ text: 'South' }, { num: 9 }, { num: 1102 }],
          [{ text: 'Total' }, { num: 21 }, { num: 2582.5 }],
        ],
        { headerRows: 1, colWidths: [180, 120, 140], headerFill: '1D4ED8', bodyStyle: cell, headerStyle: cellHead },
      ),
    },
  )
  para(body, 'Readings are averaged per site', { note: noteRef }, '.')

  let text = ''
  const paras: Table = []
  const chars: Table = []
  const lists: Table = []
  const lvls: [number, number][] = []
  const atts: Table = []
  const links: Table = []
  const notes: Table = []
  let paraStart = true
  for (const p of parts) {
    const i = text.length
    if (paraStart) {
      paras.push([i, p.para])
      lists.push([i, p.list ?? none])
      lvls.push([i, p.level ?? 0])
      paraStart = false
    }
    if (p.text !== '\n') {
      chars.push([i, p.char])
      links.push([i, p.link ? link : undefined])
    }
    if (p.att !== undefined) atts.push([i, p.att])
    if (p.note !== undefined) notes.push([i, p.note])
    text += p.text
    if (p.text === '\n') paraStart = true
  }
  const bodyStorage = storage(b, { text, paras, chars, lists, levels: lvls, attachments: atts, links, footnotes: notes })

  // header (left and right parts) and footer (center, with the page number)
  const empty = () => storage(b, { kind: 1, text: '', paras: [[0, hf]] })
  const pageNo = b.add(T.NUMBER_ATTACHMENT, msg({ 1: msg({ 2: 0 }), 2: 0, 3: '1' }))
  const headers = [storage(b, { kind: 1, text: 'Northwind Field Lab', paras: [[0, hf]] }), empty(), storage(b, { kind: 1, text: 'Q3 report', paras: [[0, hf]] })]
  const footers = [empty(), storage(b, { kind: 1, text: 'Page \ufffc', paras: [[0, hf]], attachments: [[5, pageNo]] }), empty()]
  const master = b.add(T.PAGE_MASTER, msg({ 1: headers.map(ref), 2: footers.map(ref) }))
  const section = b.add(T.SECTION, msg({ 17: true, 18: false, 23: ref(master), 24: ref(master), 25: ref(master), 26: 'Section 1' }))
  const settings = b.add(T.SETTINGS, msg({ 1: true, 2: true, 3: true, 31: 0 }))
  const floating = b.add(T.FLOATING, msg({}))
  b.add(
    T.DOCUMENT,
    msg({
      3: ref(floating),
      4: ref(bodyStorage),
      5: ref(section),
      7: ref(settings),
      30: f32(612),
      31: f32(792),
      32: f32(72),
      33: f32(72),
      34: f32(72),
      35: f32(72),
      36: f32(36),
      37: f32(36),
    }),
    1,
  )
  return b.zip()
}

// ---------------------------------------------------------------- Keynote

export const K = {
  DOCUMENT: 1,
  SHOW: 2,
  SLIDE_NODE: 4,
  SLIDE: 5,
  PLACEHOLDER: 7,
  SLIDE_STYLE: 9,
  SHAPE_STYLE: 2025,
  GROUP: 3008,
  MASK: 3006,
} as const

/** TSD.PathSourceArchive */
export const path = {
  rect: (w: number, h: number) =>
    msg({
      5: msg({
        2: msg({ 1: f32(w), 2: f32(h) }),
        3: msg({ 1: [[1, 0, 0], [2, w, 0], [2, w, h], [2, 0, h], [5]].map(([t, x, y]) => msg({ 1: t, 2: x === undefined ? undefined : msg({ 1: f32(x), 2: f32(y!) }) })) }),
      }),
    }),
  rounded: (radius: number, w: number, h: number) => msg({ 4: msg({ 1: 0, 2: f32(radius), 3: msg({ 1: f32(w), 2: f32(h) }) }) }),
  polygon: (sides: number, w: number, h: number) => msg({ 4: msg({ 1: 1, 2: f32(sides), 3: msg({ 1: f32(w), 2: f32(h) }) }) }),
  arrow: (w: number, h: number) => msg({ 3: msg({ 1: 1, 2: msg({ 1: f32(0.6), 2: f32(0.4) }), 3: msg({ 1: f32(w), 2: f32(h) }) }) }),
  /** a closed curve (an ellipse drawn with 4 cubic segments) */
  ellipse: (w: number, h: number) => {
    const k = 0.5523
    const [cx, cy, rx, ry] = [w / 2, h / 2, w / 2, h / 2]
    const pt = (x: number, y: number) => msg({ 1: f32(x), 2: f32(y) })
    const curve = (pts: number[][]) => msg({ 1: 4, 2: pts.map(([x, y]) => pt(x, y)) })
    return msg({
      5: msg({
        2: msg({ 1: f32(w), 2: f32(h) }),
        3: msg({
          1: [
            msg({ 1: 1, 2: [pt(cx + rx, cy)] }),
            curve([[cx + rx, cy + k * ry], [cx + k * rx, cy + ry], [cx, cy + ry]]),
            curve([[cx - k * rx, cy + ry], [cx - rx, cy + k * ry], [cx - rx, cy]]),
            curve([[cx - rx, cy - k * ry], [cx - k * rx, cy - ry], [cx, cy - ry]]),
            curve([[cx + k * rx, cy - ry], [cx + rx, cy - k * ry], [cx + rx, cy]]),
            msg({ 1: 5 }),
          ],
        }),
      }),
    })
  },
}

export interface ShapeLook {
  fill?: string
  stroke?: string
  strokeWidth?: number
  /** text: vertical alignment (0 top, 1 middle, 2 bottom), paragraph style */
  valign?: number
  paraStyle?: number
}

/** TSWP.ShapeStyleArchive */
export function shapeStyle(b: IworkBuilder, look: ShapeLook): number {
  const props = msg({
    1: look.fill ? msg({ 1: color(look.fill) }) : undefined,
    2: look.stroke ? msg({ 1: color(look.stroke), 2: f32(look.strokeWidth ?? 1), 6: msg({ 1: 1 }) }) : msg({ 1: color('000000'), 2: f32(1), 6: msg({ 1: 2 }) }),
  })
  return b.add(
    K.SHAPE_STYLE,
    msg({ 1: msg({ 1: styleInfo(undefined), 11: props }), 11: msg({ 2: look.valign ?? 0, 6: msg({ 1: f32(4), 2: f32(4), 3: f32(4), 4: f32(4) }), 10: look.paraStyle === undefined ? undefined : ref(look.paraStyle) }) }),
  )
}

/** a text box / shape (TSWP.ShapeInfoArchive); `placeholder`: wrapped in KN.PlaceholderArchive of that kind */
export function shape(
  b: IworkBuilder,
  o: { frame: [number, number, number, number]; path?: M; style: number; text?: number; flags?: number; angle?: number; placeholder?: number },
): number {
  const [x, y, w, h] = o.frame
  const info = msg({ 1: msg({ 1: drawable(w, h, x, y, { flags: o.flags, angle: o.angle }), 2: ref(o.style), 3: o.path ?? path.rect(w, h) }), 4: o.text === undefined ? undefined : ref(o.text) })
  if (o.placeholder !== undefined) return b.add(K.PLACEHOLDER, msg({ 1: info, 2: o.placeholder }))
  return b.add(T.SHAPE_INFO, info)
}

/** The demo / test presentation: a master with a band, title, bullets with a cropped image, a table, shapes and a group, a skipped slide. */
export async function sampleKeynote(): Promise<Uint8Array<ArrayBuffer>> {
  const b = new IworkBuilder()
  const W = 1024
  const H = 768
  const font = 'HelveticaNeue'
  const para = (name: string | undefined, c: CharSpec, p: ParaSpec = {}, parent?: number) => b.add(T.PARA_STYLE, msg({ 1: styleInfo(name, parent), 11: charProps({ font, color: '1F2937', ...c }), 12: paraProps(p) }))
  const titleStyle = para('Title', { size: 54, font: 'HelveticaNeue-Bold', color: '111827' }, { align: 2 })
  const subtitleStyle = para('Subtitle', { size: 26, color: '6B7280' }, { align: 2 })
  const headingStyle = para('Heading', { size: 40, bold: true, color: '1D4ED8' })
  const bodyStyle = para('Body', { size: 26 }, { after: 12 })
  const cellStyle = para('Table Body', { size: 20 })
  const cellHead = para('Table Header', { size: 20, bold: true, color: 'FFFFFF' })
  const caption = para('Caption', { size: 16, italic: true, color: '6B7280' }, { align: 2 })
  const label = para(undefined, { size: 18, bold: true, color: 'FFFFFF' }, { align: 2 }, bodyStyle)
  const bullets = b.add(T.LIST_STYLE, msg({ 1: styleInfo('Bullet'), 11: Array(9).fill(2), 12: Array(9).fill(f32(1)), 13: Array.from({ length: 9 }, (_, l) => f32(l * 36)), 16: ['\u2022', '\u2013', '\u2022', '\u2013', '\u2022', '\u2013', '\u2022', '\u2013', '\u2022'] }))
  const none = b.add(T.LIST_STYLE, msg({ 1: styleInfo('None'), 11: Array(9).fill(0) }))
  const textStyle = shapeStyle(b, {})
  const middle = shapeStyle(b, { valign: 1 })
  const text = (t: string, paraStyle: number, extra: Partial<Parameters<typeof storage>[1]> = {}) =>
    storage(b, { kind: 3, text: t, paras: [[0, paraStyle]], lists: [[0, none]], ...extra })

  // master: white background, a blue band at the bottom
  const slideBg = b.add(K.SLIDE_STYLE, msg({ 1: styleInfo('Master'), 11: msg({ 1: msg({ 1: color('FFFFFF') }) }) }))
  const band = shape(b, { frame: [0, H - 24, W, 24], style: shapeStyle(b, { fill: '1D4ED8' }) })
  const master = b.add(K.SLIDE, msg({ 1: ref(slideBg), 7: [ref(band)], 10: 'Title & Bullets' }))
  const slides: number[] = []
  const slide = (drawables: number[], opts: { bg?: string; hidden?: boolean } = {}) => {
    const style = opts.bg ? b.add(K.SLIDE_STYLE, msg({ 1: styleInfo(undefined, slideBg), 11: msg({ 1: msg({ 1: color(opts.bg) }) }) })) : slideBg
    const s = b.add(K.SLIDE, msg({ 1: ref(style), 7: drawables.map(ref), 17: ref(master), 42: drawables.map(ref) }))
    slides.push(b.add(K.SLIDE_NODE, msg({ 2: ref(s), 4: !!opts.hidden })))
  }

  // 1. title
  slide([
    shape(b, { frame: [80, 250, W - 160, 120], style: middle, text: text('Quarterly Review', titleStyle), placeholder: 2 }),
    shape(b, { frame: [80, 380, W - 160, 60], style: textStyle, text: text('Keynote presentation opened in the browser', subtitleStyle), placeholder: 3 }),
  ])
  // 2. bullets and a cropped image
  const items = ['Coverage grew in every region', 'Northern sites reported first', 'Response time fell by a third']
  let bodyText = ''
  const paras: Table = []
  const levels: [number, number][] = []
  items.forEach((t, k) => {
    paras.push([bodyText.length, bodyStyle])
    levels.push([bodyText.length, k === 1 ? 1 : 0])
    bodyText += t + '\n'
  })
  const data = b.data('photo.png', png(400, 300, [16, 185, 129], [37, 99, 235]))
  const mask = b.add(K.MASK, msg({ 1: drawable(300, 240, 50, 30) }))
  const photo = b.add(T.IMAGE, msg({ 1: drawable(400, 300, 600, 200), 11: msg({ 1: data }), 4: msg({ 1: f32(400), 2: f32(300) }), 5: ref(mask) }))
  slide([
    shape(b, { frame: [60, 50, W - 120, 90], style: textStyle, text: text('Highlights', headingStyle), placeholder: 2 }),
    shape(b, { frame: [60, 180, 560, 300], style: textStyle, text: storage(b, { kind: 3, text: bodyText, paras, lists: [[0, bullets]], levels }), placeholder: 3 }),
    photo,
    shape(b, { frame: [800, 490, 200, 30], style: textStyle, text: text('Field photo', caption), flags: 0 }),
  ])
  // 3. a table on a tinted background
  slide(
    [
      shape(b, { frame: [60, 50, W - 120, 90], style: textStyle, text: text('Results', headingStyle), placeholder: 2 }),
      tableInfo(
        b,
        [
          [{ text: 'Region' }, { text: 'Sites' }, { text: 'Readings' }],
          [{ text: 'North' }, { num: 12 }, { num: 1480.5 }],
          [{ text: 'South' }, { num: 9 }, { num: 1102 }],
          [{ text: 'Total' }, { num: 21 }, { num: 2582.5 }],
        ],
        { headerRows: 1, colWidths: [300, 220, 260], headerFill: '1D4ED8', bodyStyle: cellStyle, headerStyle: cellHead, band: 'E5E7EB', sidecar: true, frame: [122, 200] },
      ),
    ],
    { bg: 'F8FAFC' },
  )
  // 4. shapes and a group (children in the group's coordinates)
  const card = (fill: string, t: string, x: number, p: M) => shape(b, { frame: [x, 0, 180, 120], style: shapeStyle(b, { fill, valign: 1 }), path: p, text: text(t, label) })
  const group = b.add(K.GROUP, msg({ 1: drawable(640, 120, 190, 330), 2: [card('2563EB', 'Plan', 0, path.rounded(16, 180, 120)), card('10B981', 'Build', 230, path.ellipse(180, 120)), card('F59E0B', 'Ship', 460, path.polygon(6, 180, 120))].map(ref) }))
  slide([
    shape(b, { frame: [60, 50, W - 120, 90], style: textStyle, text: text('Process', headingStyle), placeholder: 2 }),
    group,
    shape(b, { frame: [440, 500, 140, 60], style: shapeStyle(b, { fill: '6B7280' }), path: path.arrow(140, 60) }),
  ])
  // 5. skipped
  slide([shape(b, { frame: [60, 300, W - 120, 90], style: textStyle, text: text('Backup slide', headingStyle) })], { hidden: true })

  const show = b.add(K.SHOW, msg({ 3: msg({ 2: slides.map(ref) }), 4: msg({ 1: f32(W), 2: f32(H) }) }))
  b.add(K.DOCUMENT, msg({ 2: ref(show) }), 1)
  return b.zip()
}

// ---------------------------------------------------------------- Numbers

/** seconds from 2001-01-01 UTC of a date */
export const appleDate = (y: number, m: number, d: number) => (Date.UTC(y, m - 1, d) - Date.UTC(2001, 0, 1)) / 1000

/** The demo / test spreadsheet: a budget (currency, percent, date, checkbox and duration cells), a sheet with two tables. */
export async function sampleNumbers(): Promise<Uint8Array<ArrayBuffer>> {
  const b = new IworkBuilder()
  const font = 'HelveticaNeue'
  const body = b.add(T.PARA_STYLE, msg({ 1: styleInfo('Table Body'), 11: charProps({ size: 11, font, color: '1F2937' }) }))
  const head = b.add(T.PARA_STYLE, msg({ 1: styleInfo('Table Header', body), 11: charProps({ bold: true, color: 'FFFFFF' }) }))
  const usd = fmt.currency('USD')
  const pct = fmt.percent(1)
  const day = fmt.date('MMM d, yyyy')
  const box = fmt.checkbox()
  const hm = fmt.duration(1)
  const row = (category: string, planned: number, actual: number, due: [number, number, number], paid: boolean, minutes: number): CellSpec[] => [
    { text: category },
    { num: planned, fmt: usd },
    { num: actual, fmt: usd },
    { num: actual / 2915, fmt: pct },
    { date: appleDate(...due), fmt: day },
    { bool: paid, fmt: box },
    { duration: minutes * 60, fmt: hm },
  ]
  const expenses = tableInfo(
    b,
    [
      [{ text: 'Category' }, { text: 'Planned' }, { text: 'Actual' }, { text: 'Share' }, { text: 'Due' }, { text: 'Paid' }, { text: 'Time spent' }],
      row('Rent', 1200, 1200, [2025, 1, 5], true, 30),
      row('Groceries', 450, 512.4, [2025, 1, 31], true, 210),
      row('Transport', 160, 142.6, [2025, 1, 31], false, 45),
      row('Utilities', 210, 228.15, [2025, 2, 10], false, 20),
      row('Savings', 800, 831.85, [2025, 2, 15], true, 15),
    ],
    { headerRows: 1, colWidths: [130, 100, 100, 70, 110, 60, 90], headerFill: '1D4ED8', bodyStyle: body, headerStyle: head, band: 'EFF6FF', name: 'Expenses', freeze: true },
  )
  const contacts = tableInfo(
    b,
    [
      [{ text: 'Name' }, { text: 'Role' }, { text: 'Email' }],
      [{ text: 'Ann Lee' }, { text: 'Treasurer' }, { text: 'ann@example.com' }],
      [{ text: 'Bob Ito' }, { text: 'Auditor' }, { text: 'bob@example.com' }],
    ],
    { headerRows: 1, colWidths: [120, 110, 170], headerFill: '0F766E', bodyStyle: body, headerStyle: head, name: 'Contacts' },
  )
  const totals = tableInfo(
    b,
    [
      [{ text: 'Quarter' }, { text: 'Total' }],
      [{ text: 'Q1' }, { num: 2915, fmt: usd }],
      [{ text: 'Q2' }, { num: 3120.5, fmt: usd }],
    ],
    { headerRows: 1, colWidths: [100, 120], headerFill: '0F766E', bodyStyle: body, headerStyle: head, name: 'Totals' },
  )
  const sheets = [b.add(2, msg({ 1: 'Budget', 2: [ref(expenses)] })), b.add(2, msg({ 1: 'Notes', 2: [ref(contacts), ref(totals)] }))]
  b.add(1, msg({ 1: sheets.map(ref) }), 1)
  return b.zip()
}
