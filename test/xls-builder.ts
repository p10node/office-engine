/**
 * Test helper: builds Excel 97-2003 workbooks (BIFF8 records in a compound file) record by record, for what the
 * committed fixtures (written by xlwt) do not have: shared and array formulas, comments, hyperlinks, names, 3D
 * references, rich text, CONTINUE records at chosen places, BIFF5, encryption.
 */
import { writeCfb } from './cfb-writer'

export type Bytes = number[] | Uint8Array

export const u8 = (n: number) => [n & 0xff]
export const u16 = (n: number) => [n & 0xff, (n >> 8) & 0xff]
export const u32 = (n: number) => [n & 0xff, (n >> 8) & 0xff, (n >> 16) & 0xff, (n >>> 24) & 0xff]
export const f64 = (n: number) => {
  const b = new Uint8Array(8)
  new DataView(b.buffer).setFloat64(0, n, true)
  return [...b]
}
const wide = (s: string) => !/^[\x00-\xff]*$/.test(s)
/** the characters of a string: 8-bit when they all fit, else 16-bit (after the flag byte) */
export const chars = (s: string, high = wide(s)) => [high ? 1 : 0, ...[...s].flatMap((_, i) => (high ? u16(s.charCodeAt(i)) : u8(s.charCodeAt(i))))]
/** XLUnicodeString: 16-bit length, flags, characters */
export const str16 = (s: string) => [...u16(s.length), ...chars(s)]
/** ShortXLUnicodeString: 8-bit length, flags, characters */
export const str8 = (s: string) => [...u8(s.length), ...chars(s)]
export const cat = (...parts: Bytes[]) => parts.flatMap((p) => [...p])

/** a record: id, length, data */
export const rec = (type: number, ...parts: Bytes[]) => {
  const data = cat(...parts)
  return [...u16(type), ...u16(data.length), ...data]
}

export const BOF = (dt: number, vers = 0x0600) => rec(0x0809, u16(vers), u16(dt), u16(0x0dbb), u16(0x07cc), u32(0), u32(0x0206))
export const EOF = () => rec(0x000a)

export interface FontSpec {
  name?: string
  /** pt */
  size?: number
  bold?: boolean
  italic?: boolean
  color?: number
}
export const font = (f: FontSpec = {}) =>
  rec(0x0031, u16((f.size ?? 10) * 20), u16(f.italic ? 2 : 0), u16(f.color ?? 0x7fff), u16(f.bold ? 700 : 400), u16(0), u8(0), u8(0), u8(0), u8(0), str8(f.name ?? 'Arial'))

export interface XfSpec {
  font?: number
  fmt?: number
  style?: boolean
  /** horizontal alignment 0-7, wrap */
  align?: number
  wrap?: boolean
  /** fill pattern, fore color */
  pattern?: number
  fore?: number
  /** thin border on all sides with this color */
  border?: number
}
export const xf = (x: XfSpec = {}) => {
  const lines = x.border !== undefined ? 1 | (1 << 4) | (1 << 8) | (1 << 12) | (x.border << 16) | (x.border << 23) : 0
  const lines2 = x.border !== undefined ? x.border | (x.border << 7) : 0
  return rec(
    0x00e0,
    u16(x.font ?? 0),
    u16(x.fmt ?? 0),
    u16(x.style ? 0xfff5 : 0x0001),
    u8((x.align ?? 0) | (x.wrap ? 8 : 0) | 0x20),
    u8(0),
    u8(0),
    u8(x.style ? 0 : 0xf8),
    u32(lines),
    u32(lines2 | ((x.pattern ?? 0) << 26)),
    u16((x.fore ?? 64) | (65 << 7)),
  )
}

export interface SheetSpec {
  name: string
  records: Bytes[]
  /** 0 visible, 1 hidden, 2 very hidden */
  state?: number
  /** 0 worksheet, 2 chart */
  type?: number
}

export interface BookSpec {
  sheets: SheetSpec[]
  /** shared strings; or the SST record's parts (first + CONTINUEs) as bytes */
  sst?: string[] | Bytes[]
  fonts?: FontSpec[]
  /** cell formats after the 16 default ones (index 16 on) */
  xfs?: XfSpec[]
  formats?: [number, string][]
  /** records after the sheet list (SUPBOOK, EXTERNSHEET, NAME, PALETTE…) */
  extra?: Bytes[]
  /** records right after BOF (FILEPASS) */
  head?: Bytes[]
  date1904?: boolean
}

/** the SST record of plain strings */
export const sstRecord = (strings: string[]) => rec(0x00fc, u32(strings.length), u32(strings.length), ...strings.map(str16))

/** the Workbook stream of a BIFF8 workbook */
export function workbookStream(b: BookSpec): Uint8Array {
  const fonts = b.fonts ?? [{}, {}, {}, {}]
  // BIFF skips font index 4: write four fonts before any others
  while (fonts.length < 4) fonts.push({})
  const xfs = [...Array.from({ length: 15 }, () => xf({ style: true })), xf(), ...(b.xfs ?? []).map(xf)]
  const head = cat(BOF(0x0005), ...(b.head ?? []), rec(0x0042, u16(1200)), rec(0x0022, u16(b.date1904 ? 1 : 0)), ...fonts.map(font), ...(b.formats ?? []).map(([id, code]) => rec(0x041e, u16(id), str16(code))), ...xfs)
  const sheetRec = (s: SheetSpec, pos: number) => rec(0x0085, u32(pos), u8(s.state ?? 0), u8(s.type ?? 0), str8(s.name))
  let sst: number[] = []
  if (b.sst?.length) sst = typeof b.sst[0] === 'string' ? sstRecord(b.sst as string[]) : cat(...(b.sst as Bytes[]))
  const tail = cat(...(b.extra ?? []), sst, EOF())
  const globalsLen = head.length + b.sheets.reduce((n, s) => n + sheetRec(s, 0).length, 0) + tail.length
  const bodies = b.sheets.map((s) => cat(BOF(s.type === 2 ? 0x0020 : 0x0010), ...s.records, EOF()))
  let pos = globalsLen
  const list = b.sheets.map((s, k) => {
    const r = sheetRec(s, pos)
    pos += bodies[k].length
    return r
  })
  return Uint8Array.from(cat(head, ...list, tail, ...bodies))
}

/** a compound file holding a workbook stream */
export function xlsFile(b: BookSpec): Uint8Array<ArrayBuffer> {
  return writeCfb({ Workbook: workbookStream(b) })
}

// cells and sheet records
export const number = (r: number, c: number, v: number, x = 15) => rec(0x0203, u16(r), u16(c), u16(x), f64(v))
export const labelSst = (r: number, c: number, i: number, x = 15) => rec(0x00fd, u16(r), u16(c), u16(x), u32(i))
export const label = (r: number, c: number, s: string, x = 15) => rec(0x0204, u16(r), u16(c), u16(x), str16(s))
export const rk = (r: number, c: number, v: number, x = 15) => rec(0x027e, u16(r), u16(c), u16(x), u32(((v << 2) | 2) >>> 0))
export const boolErr = (r: number, c: number, v: number, isErr: boolean, x = 15) => rec(0x0205, u16(r), u16(c), u16(x), u8(v), u8(isErr ? 1 : 0))
export const blank = (r: number, c: number, x: number) => rec(0x0201, u16(r), u16(c), u16(x))
export const mulBlank = (r: number, c0: number, xs: number[]) => rec(0x00be, u16(r), u16(c0), ...xs.map(u16), u16(c0 + xs.length - 1))
export const mulRk = (r: number, c0: number, vs: number[], x = 15) => rec(0x00bd, u16(r), u16(c0), vs.flatMap((v) => [...u16(x), ...u32(((v << 2) | 2) >>> 0)]), u16(c0 + vs.length - 1))
/** a FORMULA record: `value` a number, or [type, byte] for string (0), boolean (1), error (2), empty (3) results */
export const formula = (r: number, c: number, rgce: Bytes, value: number | [number, number] = 0, x = 15, extra: Bytes = []) =>
  rec(0x0006, u16(r), u16(c), u16(x), typeof value === 'number' ? f64(value) : [value[0], 0, value[1], 0, 0, 0, 0xff, 0xff], u16(0), u32(0), u16(rgce.length), rgce, extra)
export const stringRec = (s: string) => rec(0x0207, str16(s))
export const row = (r: number, heightPt: number, opts: { custom?: boolean; hidden?: boolean } = {}) =>
  rec(0x0208, u16(r), u16(0), u16(1), u16(heightPt * 20), u16(0), u16(0), u8((opts.custom ? 0x40 : 0) | (opts.hidden ? 0x20 : 0)), u8(1), u16(15))
export const colInfo = (c0: number, c1: number, widthChars: number, hidden = false) => rec(0x007d, u16(c0), u16(c1), u16(Math.round(widthChars * 256)), u16(15), u16(hidden ? 1 : 0), u16(0))
export const merge = (...refs: [number, number, number, number][]) => rec(0x00e5, u16(refs.length), ...refs.map(([r0, r1, c0, c1]) => cat(u16(r0), u16(r1), u16(c0), u16(c1))))
/** WINDOW2: gridlines on/off, frozen panes, selected tab */
export const window2 = (o: { grid?: boolean; frozen?: boolean; selected?: boolean } = {}) =>
  rec(0x023e, u16(0x04 | 0x10 | 0x20 | 0x80 | (o.grid === false ? 0 : 2) | (o.frozen ? 0x108 : 0) | (o.selected ? 0x600 : 0)), u16(0), u16(0), u32(64), u16(0), u16(0), u32(0))
export const pane = (cols: number, rows: number) => rec(0x0041, u16(cols), u16(rows), u16(rows), u16(cols), u8(0), u8(0))

// formula tokens
export const ptg = {
  ref: (r: number, c: number, rowRel = true, colRel = true) => [0x24, ...u16(r), ...u16(c | (colRel ? 0x4000 : 0) | (rowRel ? 0x8000 : 0))],
  area: (r0: number, r1: number, c0: number, c1: number, rel = true) => [0x25, ...u16(r0), ...u16(r1), ...u16(c0 | (rel ? 0xc000 : 0)), ...u16(c1 | (rel ? 0xc000 : 0))],
  refN: (dr: number, dc: number, rowRel = true, colRel = true) => [0x2c, ...u16(dr), ...u16((dc & 0xff) | (colRel ? 0x4000 : 0) | (rowRel ? 0x8000 : 0))],
  ref3d: (ixti: number, r: number, c: number) => [0x3a, ...u16(ixti), ...u16(r), ...u16(c | 0xc000)],
  area3d: (ixti: number, r0: number, r1: number, c0: number, c1: number, rel = true) => [0x3b, ...u16(ixti), ...u16(r0), ...u16(r1), ...u16(c0 | (rel ? 0xc000 : 0)), ...u16(c1 | (rel ? 0xc000 : 0))],
  int: (n: number) => [0x1e, ...u16(n)],
  num: (n: number) => [0x1f, ...f64(n)],
  str: (s: string) => [0x17, ...str8(s)],
  bool: (b: boolean) => [0x1d, b ? 1 : 0],
  err: (code: number) => [0x1c, code],
  name: (i: number) => [0x23, ...u16(i), 0, 0],
  nameX: (ixti: number, i: number) => [0x39, ...u16(ixti), ...u16(i), 0, 0],
  func: (i: number) => [0x41, ...u16(i)],
  funcVar: (i: number, argc: number) => [0x42, argc, ...u16(i)],
  exp: (r: number, c: number) => [0x01, ...u16(r), ...u16(c)],
  attrSum: () => [0x19, 0x10, 0, 0],
  attrSpace: (n: number) => [0x19, 0x40, 0, n],
  attrIf: (jump: number) => [0x19, 0x02, ...u16(jump)],
  attrChoose: (offsets: number[]) => [0x19, 0x04, ...u16(offsets.length - 1), ...offsets.flatMap(u16)],
  array: () => [0x60, 0, 0, 0, 0, 0, 0, 0],
  missArg: () => [0x16],
  op: (code: number) => [code],
  paren: () => [0x15],
}
