/**
 * Writes small PowerPoint 97-2003 files (MS-PPT) for tests and the playground sample: records and OfficeArt
 * drawings, text with its paragraph and character runs, master text styles, pictures in the Pictures stream, the
 * persist directory with one or more edits, the Current User stream, all in a compound file.
 */
import { writeCfb } from './cfb-writer'

// ---------------------------------------------------------------- bytes

export type Bytes = number[] | Uint8Array

export const u8 = (v: number) => [v & 0xff]
export const u16 = (v: number) => [v & 0xff, (v >>> 8) & 0xff]
export const u32 = (v: number) => [v & 0xff, (v >>> 8) & 0xff, (v >>> 16) & 0xff, (v >>> 24) & 0xff]

export function cat(...parts: Bytes[]): Uint8Array {
  const out = new Uint8Array(parts.reduce((n, p) => n + p.length, 0))
  let at = 0
  for (const p of parts) {
    out.set(p, at)
    at += p.length
  }
  return out
}

const chars16 = (s: string) => {
  const out: number[] = []
  for (let i = 0; i < s.length; i++) out.push(...u16(s.charCodeAt(i)))
  return out
}

/** a record: header (version, instance, type, length) and body */
export const rec = (type: number, body: Bytes = [], inst = 0, ver = 0) => cat(u16((ver & 0xf) | (inst << 4)), u16(type), u32(body.length), body)
/** a container record (version 0xF) */
export const box = (type: number, kids: Bytes[], inst = 0) => rec(type, cat(...kids), inst, 0xf)
/** a CString atom */
export const cstring = (s: string, inst = 0) => rec(0x0fba, chars16(s), inst)

// ---------------------------------------------------------------- colors

/** ColorIndexStruct of an RGB color (text) */
export const rgb = (hex: string) => {
  const v = parseInt(hex, 16)
  return (((v >> 16) & 0xff) | (((v >> 8) & 0xff) << 8) | ((v & 0xff) << 16) | (0xfe << 24)) >>> 0
}
/** ColorIndexStruct of a scheme color (text) */
export const schemeText = (i: number) => (i << 24) >>> 0
/** OfficeArtCOLORREF of an RGB color (shapes) */
export const oa = (hex: string) => {
  const v = parseInt(hex, 16)
  return ((v >> 16) & 0xff) | (((v >> 8) & 0xff) << 8) | ((v & 0xff) << 16)
}
/** OfficeArtCOLORREF of a scheme color (shapes) */
export const oaScheme = (i: number) => (0x08000000 | i) >>> 0

/** 16.16 fixed point */
export const fixed = (v: number) => Math.round(v * 65536) >>> 0

// ---------------------------------------------------------------- text properties

export interface PFSpec {
  bullet?: boolean
  bulletChar?: string
  bulletFont?: number
  /** percent */
  bulletSize?: number
  bulletColor?: number
  /** 0 left, 1 center, 2 right, 3 justify */
  align?: number
  lineSpacing?: number
  spaceBefore?: number
  spaceAfter?: number
  /** master units: where the text starts */
  leftMargin?: number
  /** master units: where the bullet starts */
  indent?: number
  tabs?: [number, number][]
}

export function pfBytes(p: PFSpec): Uint8Array {
  let m = 0
  const out: number[] = []
  const flags = (p.bullet !== undefined ? 1 : 0) | (p.bulletFont !== undefined ? 2 : 0) | (p.bulletColor !== undefined ? 4 : 0) | (p.bulletSize !== undefined ? 8 : 0)
  if (flags) {
    m |= flags
    out.push(...u16((p.bullet ? 1 : 0) | (flags & 0xe)))
  }
  if (p.bulletChar !== undefined) (m |= 0x80), out.push(...u16(p.bulletChar.charCodeAt(0)))
  if (p.bulletFont !== undefined) (m |= 0x10), out.push(...u16(p.bulletFont))
  if (p.bulletSize !== undefined) (m |= 0x40), out.push(...u16(p.bulletSize))
  if (p.bulletColor !== undefined) (m |= 0x20), out.push(...u32(p.bulletColor))
  if (p.align !== undefined) (m |= 0x800), out.push(...u16(p.align))
  if (p.lineSpacing !== undefined) (m |= 0x1000), out.push(...u16(p.lineSpacing))
  if (p.spaceBefore !== undefined) (m |= 0x2000), out.push(...u16(p.spaceBefore))
  if (p.spaceAfter !== undefined) (m |= 0x4000), out.push(...u16(p.spaceAfter))
  if (p.leftMargin !== undefined) (m |= 0x100), out.push(...u16(p.leftMargin))
  if (p.indent !== undefined) (m |= 0x400), out.push(...u16(p.indent))
  if (p.tabs) (m |= 0x100000), out.push(...u16(p.tabs.length), ...p.tabs.flatMap(([pos, type]) => [...u16(pos), ...u16(type)]))
  return cat(u32(m), out)
}

export interface CFSpec {
  bold?: boolean
  italic?: boolean
  underline?: boolean
  font?: number
  /** points */
  size?: number
  color?: number
  /** superscript / subscript percent */
  position?: number
  /** run group of the PowerPoint 2000 extras (numbering) */
  pp9rt?: number
}

export function cfBytes(c: CFSpec): Uint8Array {
  let m = 0
  let style = 0
  if (c.bold !== undefined) (m |= 1), (style |= c.bold ? 1 : 0)
  if (c.italic !== undefined) (m |= 2), (style |= c.italic ? 2 : 0)
  if (c.underline !== undefined) (m |= 4), (style |= c.underline ? 4 : 0)
  if (c.pp9rt !== undefined) (m |= 0x3c00), (style |= c.pp9rt << 10)
  const out: number[] = m ? u16(style) : []
  if (c.font !== undefined) (m |= 0x10000), out.push(...u16(c.font))
  if (c.size !== undefined) (m |= 0x20000), out.push(...u16(c.size))
  if (c.color !== undefined) (m |= 0x40000), out.push(...u32(c.color))
  if (c.position !== undefined) (m |= 0x80000), out.push(...u16(c.position))
  return cat(u32(m), out)
}

/** a TxMasterStyleAtom: the levels of a text type's master style */
export function masterStyle(inst: number, levels: { pf?: PFSpec; cf?: CFSpec }[]): Uint8Array {
  return rec(0x0fa3, cat(u16(levels.length), ...levels.flatMap((l, i) => [inst >= 5 ? u16(i) : [], pfBytes(l.pf ?? {}), cfBytes(l.cf ?? {})])), inst)
}

/** text types */
export const TX = { title: 0, body: 1, notes: 2, other: 4, centerBody: 5, centerTitle: 6 }

export interface RunSpec {
  text: string
  cf?: CFSpec
  /** an ExHyperlink id: the run is a link */
  link?: number
  /** a field: the run's single character is replaced by it */
  field?: 'slidenum' | 'footer' | 'date'
}

export interface ParaSpec {
  level?: number
  pf?: PFSpec
  runs: RunSpec[]
}

export interface TextSpec {
  type: number
  paras: (ParaSpec | string)[]
  /** 8-bit characters (TextBytesAtom) */
  bytes?: boolean
}

/** the records of a text: header, characters, style runs, link ranges, fields (`index`: of a slide's placeholder texts) */
export function textRecords(t: TextSpec, index = 0): Uint8Array[] {
  const paras = t.paras.map((p) => (typeof p === 'string' ? { runs: [{ text: p }] } : p))
  let text = ''
  const pf: Uint8Array[] = []
  const cf: Uint8Array[] = []
  const extra: Uint8Array[] = []
  paras.forEach((p, k) => {
    const last = k === paras.length - 1
    const len = p.runs.reduce((n, r) => n + r.text.length, 0)
    pf.push(cat(u32(len + 1), u16(p.level ?? 0), pfBytes(p.pf ?? {})))
    if (!p.runs.length) cf.push(cat(u32(1), cfBytes({})))
    p.runs.forEach((r, i) => {
      const start = text.length
      text += r.text
      cf.push(cat(u32(r.text.length + (i === p.runs.length - 1 ? 1 : 0)), cfBytes(r.cf ?? {})))
      if (r.link !== undefined) {
        extra.push(box(0x0ff2, [rec(0x0ff3, cat(u32(0), u32(r.link), u8(4), u8(0), u8(0), u8(0), u8(8), [0, 0, 0]))]))
        extra.push(rec(0x0fdf, cat(u32(start), u32(start + r.text.length))))
      }
      if (r.field === 'slidenum') extra.push(rec(0x0fd8, u32(start)))
      else if (r.field === 'footer') extra.push(rec(0x0ffa, u32(start)))
      else if (r.field === 'date') extra.push(rec(0x0ff8, u32(start)))
    })
    if (!last) text += '\r'
  })
  const chars = t.bytes ? rec(0x0fa8, [...text].map((c) => c.charCodeAt(0) & 0xff)) : rec(0x0fa0, chars16(text))
  return [rec(0x0f9f, u32(t.type), index), chars, rec(0x0fa1, cat(...pf, ...cf)), ...extra]
}

// ---------------------------------------------------------------- shapes

/** property ids */
export const P = {
  rotation: 0x0004,
  dxTextLeft: 0x0081,
  anchorText: 0x0087,
  pib: 0x0104,
  cropFromTop: 0x0100,
  cropFromLeft: 0x0102,
  geoRight: 0x0142,
  geoBottom: 0x0143,
  pVertices: 0x0145,
  pSegmentInfo: 0x0146,
  adjustValue: 0x0147,
  adjust2Value: 0x0148,
  fillType: 0x0180,
  fillColor: 0x0181,
  fillBackColor: 0x0183,
  fillBlip: 0x4186,
  fillAngle: 0x018b,
  fillFocus: 0x018c,
  fillBooleans: 0x01bf,
  lineColor: 0x01c0,
  lineWidth: 0x01cb,
  lineDashing: 0x01ce,
  lineEndArrowhead: 0x01d1,
  lineBooleans: 0x01ff,
  shadowBooleans: 0x023f,
  hspMaster: 0x0301,
  shapeName: 0x0380,
  groupBooleans: 0x03bf,
}

/** fill on / off (with its "use" bit) */
export const filled = (on: boolean) => (0x00100000 | (on ? 0x10 : 0)) >>> 0
/** line on / off */
export const lined = (on: boolean) => (0x00080000 | (on ? 0x08 : 0)) >>> 0

export interface ShapeSpec {
  /** shape type (1 rectangle, 3 ellipse, 202 text box, 75 picture…) */
  spt?: number
  /** x, y, width, height in master units (576 per inch) */
  at: [number, number, number, number]
  /** inside a group: a child anchor in the group's coordinates */
  child?: boolean
  props?: [number, number][]
  complex?: [number, Uint8Array][]
  /** FSP flags: 0x40 flip horizontally, 0x80 vertically */
  flags?: number
  /** OEPlaceholderAtom: placeholder type, placement id */
  ph?: [number, number]
  text?: TextSpec
  /** text of the slide's list (OutlineTextRefAtom index) */
  outline?: number
  /** click action: an ExHyperlink id */
  click?: number
  /** a StyleTextProp9Atom body (numbering) */
  ppt9?: Uint8Array
  /** the master shape it takes unset properties from (placeholders: the master's placeholder) */
  master?: number
}

/** a shape record with its shape id */
export type Shape = Uint8Array & { spid: number }

/** OfficeArtFOPT: simple then complex properties, sorted by id */
export function fopt(props: [number, number][], complex: [number, Uint8Array][] = [], type = 0xf00b): Uint8Array {
  const all = [...props.map(([id, v]) => ({ id, v, data: undefined as Uint8Array | undefined })), ...complex.map(([id, data]) => ({ id: id | 0x8000, v: data.length, data }))]
  all.sort((a, b) => (a.id & 0x3fff) - (b.id & 0x3fff))
  return rec(type, cat(...all.map((p) => cat(u16(p.id), u32(p.v))), ...all.flatMap((p) => (p.data ? [p.data] : []))), all.length, 3)
}

/** an IMsoArray of elements (vertices, segments) */
export const msoArray = (elems: Bytes[], size: number) => cat(u16(elems.length), u16(elems.length), u16(size), ...elems)

let nextSpid = 0x400

/** an OfficeArtSpContainer */
export function shape(s: ShapeSpec): Shape {
  const spid = nextSpid++
  const [x, y, w, h] = s.at
  const kids: Bytes[] = [rec(0xf00a, cat(u32(spid), u32(0xa00 | (s.child ? 2 : 0) | (s.master !== undefined ? 0x20 : 0) | (s.flags ?? 0))), s.spt ?? 1, 2)]
  const props = [...(s.props ?? []), ...(s.master !== undefined ? ([[P.hspMaster, s.master]] as [number, number][]) : [])]
  if (props.length || s.complex?.length) kids.push(fopt(props, s.complex ?? []))
  kids.push(s.child ? rec(0xf00f, cat(u32(x), u32(y), u32(x + w), u32(y + h))) : rec(0xf010, cat(u16(y), u16(x), u16(x + w), u16(y + h))))
  const data: Bytes[] = []
  if (s.ph) data.push(rec(0x0bc3, cat(u32(s.ph[1]), u8(s.ph[0]), u8(0), u16(0))))
  if (s.click !== undefined) data.push(box(0x0ff2, [rec(0x0ff3, cat(u32(0), u32(s.click), u8(4), u8(0), u8(0), u8(0), u8(8), [0, 0, 0]))]))
  if (s.ppt9) data.push(box(0x1388, [box(0x138a, [cstring('___PPT9'), rec(0x138b, rec(0x0fac, s.ppt9))])]))
  if (data.length) kids.push(box(0xf011, data))
  if (s.outline !== undefined) kids.push(box(0xf00d, [rec(0x0f9e, u32(s.outline))]))
  else if (s.text) kids.push(box(0xf00d, textRecords(s.text)))
  return Object.assign(box(0xf004, kids), { spid })
}

/** an OfficeArtSpgrContainer: a group at `at` whose children are placed in `coords` (x, y, width, height) */
export function group(at: [number, number, number, number], coords: [number, number, number, number], kids: Uint8Array[], child = false): Uint8Array {
  const [x, y, w, h] = at
  const [cx, cy, cw, ch] = coords
  const spid = nextSpid++
  const anchor = child ? rec(0xf00f, cat(u32(x), u32(y), u32(x + w), u32(y + h))) : rec(0xf010, cat(u16(y), u16(x), u16(x + w), u16(y + h)))
  const self = box(0xf004, [rec(0xf009, cat(u32(cx), u32(cy), u32(cx + cw), u32(cy + ch)), 0, 1), rec(0xf00a, cat(u32(spid), u32(0x201 | (child ? 2 : 0))), 0, 2), anchor])
  return box(0xf003, [self, ...kids])
}

/** a PPDrawing: the patriarch group with the shapes, and the background shape */
export function drawing(shapes: Uint8Array[], background: [number, number][] = [[P.fillColor, oaScheme(0)], [P.fillBooleans, filled(true)], [P.lineBooleans, lined(false)]]): Uint8Array {
  const patriarch = box(0xf004, [rec(0xf009, new Uint8Array(16), 0, 1), rec(0xf00a, cat(u32(nextSpid++), u32(5)), 0, 2)])
  const bg = box(0xf004, [rec(0xf00a, cat(u32(nextSpid++), u32(0xc00)), 1, 2), fopt(background)])
  return box(0x040c, [box(0xf002, [rec(0xf008, cat(u32(shapes.length + 2), u32(nextSpid)), nextDg++), box(0xf003, [patriarch, ...shapes]), bg])])
}

/** drawing ids (each drawing has its own) */
let nextDg = 1

/** a StyleTextProp9Atom body: per run group, numbering (scheme 3 = 1. 2. 3., start) or none */
export function style9(groups: ({ scheme: number; start: number } | null)[]): Uint8Array {
  return cat(...groups.map((g) => cat(g ? cat(u32(0x03800000), u16(0xffff), u16(1), u16(g.scheme), u16(g.start)) : u32(0), u32(0), u32(0))))
}

// ---------------------------------------------------------------- the file

/** default scheme: background, text, shadow, title, fill, accent, hyperlink, followed hyperlink */
export const SCHEME = ['FFFFFF', '000000', '808080', '1F2937', '2563EB', '10B981', '0563C1', '954F72']

const scheme = (inst: number, colors: string[]) => rec(0x07f0, cat(...colors.map((c) => cat(u8(parseInt(c.slice(0, 2), 16)), u8(parseInt(c.slice(2, 4), 16)), u8(parseInt(c.slice(4, 6), 16)), u8(0)))), inst)

interface Sheet {
  slideId: number
  persist: number
  body: Uint8Array
  texts: Uint8Array[][]
}

export interface SlideSpec {
  shapes: Uint8Array[]
  /** placeholder texts (OutlineTextRefAtom refers to them by index) */
  texts?: TextSpec[]
  /** the master's slide id (default: the first master) */
  master?: number
  /** SlideAtom flags: 1 master objects, 2 master scheme, 4 master background (default 7) */
  flags?: number
  /** own background (with flags & 4 cleared) */
  background?: [number, number][]
  /** own color scheme (with flags & 2 cleared) */
  scheme?: string[]
  hidden?: boolean
  /** layout geometry (SlideLayoutType): 0 title slide, 1 title and body, 7 title only, 16 blank */
  layout?: number
  /** placeholder types of the layout (default: those of the layout) */
  placeholders?: number[]
  notes?: TextSpec
  /** more records of the slide container (programmable tags…) */
  extra?: Uint8Array[]
}

/** a PowerPoint 2002 comment (in the slide's "___PPT10" tag): author, initials, text, date, position in master units */
export function comment10(author: string, initials: string, text: string, date: [number, number, number, number, number, number], at: [number, number], index = 1): Uint8Array {
  const [y, mo, d, h, mi, s] = date
  const time = cat(u16(y), u16(mo), u16(0), u16(d), u16(h), u16(mi), u16(s), u16(0))
  return box(0x2ee0, [cstring(author, 0), cstring(text, 1), cstring(initials, 2), rec(0x2ee1, cat(u32(index), time, u32(at[0]), u32(at[1])))])
}

/** a programmable binary tag container ("___PPT10" with comments…) */
export const progTag = (name: string, records: Uint8Array[]) => box(0x1388, [box(0x138a, [cstring(name), rec(0x138b, cat(...records))])])

export interface MasterSpec {
  shapes: Uint8Array[]
  styles: Uint8Array[]
  scheme?: string[]
  background?: [number, number][]
  /** a title master: its main master's slide id */
  titleOf?: number
  /** title master flags */
  flags?: number
}

/** Builds a PowerPoint 97-2003 file (shapes made after the builder get ids from a fresh start, so output is stable). */
export class PptBuilder {
  constructor() {
    nextSpid = 0x400
    nextDg = 1
  }

  fonts: string[] = ['Arial']
  private links: string[] = []
  private blips: Uint8Array[] = []
  private masters: MasterSpec[] = []
  private slides: SlideSpec[] = []
  /** document-level text defaults and the "other" text style */
  environmentStyles: Uint8Array[] = [masterStyle(TX.other, [{ pf: { align: 0, bullet: false }, cf: { font: 0, size: 18, color: schemeText(1), bold: false, italic: false, underline: false } }])]
  /** headers and footers of slides: flags (1 date, 4 user date, 8 slide number, 0x20 footer) and footer text */
  headersFooters?: { flags: number; footer?: string; date?: string }
  /** slide size in master units */
  size: [number, number] = [5760, 4320]
  /** keep only this slide (previews of one slide in viewers that show the first) */
  only?: number
  /** incremental edits appended after the first save: persist id → new record */
  private edits: Map<number, Uint8Array>[] = []

  font(name: string): number {
    const i = this.fonts.indexOf(name)
    return i >= 0 ? i : this.fonts.push(name) - 1
  }

  link(url: string): number {
    return this.links.push(url)
  }

  /** a picture (PNG, JPEG or DIB bytes): its BLIP index */
  picture(data: Uint8Array, kind: 'png' | 'jpg' | 'dib' = 'png'): number {
    const [type, inst] = kind === 'png' ? [0xf01e, 0x6e0] : kind === 'jpg' ? [0xf01d, 0x46a] : [0xf01f, 0x7a8]
    this.blips.push(rec(type, cat(new Uint8Array(16).fill(this.blips.length + 1), u8(0xff), data), inst))
    return this.blips.length
  }

  /** a BLIP record as is (EMF, WMF…) */
  rawBlip(record: Uint8Array): number {
    return this.blips.push(record)
  }

  master(m: MasterSpec): number {
    this.masters.push(m)
    return 0x80000000 + this.masters.length - 1
  }

  slide(s: SlideSpec): number {
    this.slides.push(s)
    return this.slides.length - 1
  }

  /** the persist ids: 1 the document, then masters, slides, notes, the notes master */
  get ids() {
    const masters = this.masters.map((_, i) => 2 + i)
    const slides = this.slides.map((_, i) => 2 + this.masters.length + i)
    let n = 2 + this.masters.length + this.slides.length
    const notes = this.slides.map((s) => (s.notes ? n++ : 0))
    const notesMaster = n++
    return { masters, slides, notes, notesMaster, next: n }
  }

  /** an incremental save: these records replace the ones of their persist ids */
  edit(records: Map<number, Uint8Array>) {
    this.edits.push(records)
  }

  slideRecord(i: number): Uint8Array {
    return this.sheetSlide(this.slides[i], i).body
  }

  private sheetSlide(s: SlideSpec, i: number): Sheet {
    const ids = this.ids
    const master = s.master ?? 0x80000000
    const flags = s.flags ?? 7
    // the layout's placeholder types: readers pick the master's look by them
    const layout = s.layout ?? 1
    const types = s.placeholders ?? (layout === 0 ? [0x0f, 0x10] : layout === 16 ? [] : layout === 7 ? [0x0d] : [0x0d, 0x0e])
    const kids: Bytes[] = [rec(0x03ef, cat(u32(layout), [...types, 0, 0, 0, 0, 0, 0, 0, 0].slice(0, 8), u32(master), u32(s.notes ? 0x200 + i : 0), u16(flags), u16(0)), 0, 2)]
    if (s.hidden) kids.push(rec(0x03f9, cat(u32(0), u32(0), u8(0), u8(0), u16(4), u8(0), [0, 0, 0])))
    kids.push(drawing(s.shapes, s.background), scheme(1, s.scheme ?? this.masters[0]?.scheme ?? SCHEME), ...(s.extra ?? []))
    return { slideId: 256 + i, persist: ids.slides[i], body: box(0x03ee, kids), texts: (s.texts ?? []).map((t, k) => textRecords(t, k)) }
  }

  build(): Uint8Array<ArrayBuffer> {
    if (this.only !== undefined) this.slides = [this.slides[this.only]]
    const ids = this.ids
    const sheets: Sheet[] = []
    this.masters.forEach((m, i) => {
      const id = 0x80000000 + i
      // layout: 1 for a main master, 2 (SL_MasterTitle) for a title master, with their placeholder types
      const title = m.titleOf !== undefined
      const atom = rec(0x03ef, cat(u32(title ? 2 : 1), title ? [3, 4, 7, 9, 8, 0, 0, 0] : [1, 2, 7, 9, 8, 0, 0, 0], u32(m.titleOf ?? 0), u32(0), u16(m.flags ?? 0), u16(0)), 0, 2)
      const kids: Bytes[] = [atom]
      if (m.titleOf === undefined) {
        kids.push(scheme(6, m.scheme ?? SCHEME))
        // every text type has a style (title, body, notes, center body / title, half / quarter body)
        const have = new Set(m.styles.map((r) => (r[0] | (r[1] << 8)) >> 4))
        const empty = (inst: number) =>
          masterStyle(inst, inst === TX.title || inst === TX.centerTitle ? [{}] : Array.from({ length: 5 }, () => (inst === TX.centerBody ? { pf: { bullet: false, align: 1, leftMargin: 0, indent: 0 } } : {})))
        kids.push(...m.styles, ...[0, 1, 2, 5, 6, 7, 8].filter((i) => !have.has(i)).map(empty))
      }
      kids.push(drawing(m.shapes, m.background), scheme(1, m.scheme ?? SCHEME))
      sheets.push({ slideId: id, persist: ids.masters[i], body: box(m.titleOf !== undefined ? 0x03ee : 0x03f8, kids), texts: [] })
    })
    const slides = this.slides.map((s, i) => this.sheetSlide(s, i))
    const notes: Sheet[] = []
    this.slides.forEach((s, i) => {
      if (!s.notes) return
      const body = box(0x03f0, [
        rec(0x03f1, cat(u32(256 + i), u16(7), u16(0)), 0, 1),
        drawing([
          shape({ spt: 1, at: [700, 500, 2900, 2200], ph: [0x0b, 0] }),
          shape({ spt: 1, at: [580, 2900, 3160, 2600], ph: [0x0c, 1], text: s.notes }),
        ]),
      ])
      notes.push({ slideId: 0x200 + i, persist: ids.notes[i], body, texts: [] })
    })

    const notesMaster = box(0x03f0, [
      rec(0x03f1, cat(u32(0), u16(0), u16(0)), 0, 1),
      drawing([shape({ spt: 1, at: [700, 500, 2900, 2200], ph: [5, 0] }), shape({ spt: 1, at: [580, 2900, 3160, 2600], ph: [6, 1], text: { type: TX.notes, paras: ['Click to edit Master text styles'] } })]),
      scheme(1, this.masters[0]?.scheme ?? SCHEME),
    ])
    const persistAtom = (sh: Sheet) => rec(0x03f3, cat(u32(sh.persist), u32(0), u32(sh.texts.length), u32(sh.slideId), u32(0)))
    const list = (inst: number, items: Sheet[]) => box(0x0ff0, items.flatMap((sh) => [persistAtom(sh), ...sh.texts.flat()]), inst)
    const font = (name: string, i: number) => rec(0x0fb7, cat(chars16(name.slice(0, 31)), new Uint8Array(64 - Math.min(31, name.length) * 2), u8(0), u8(0), u8(4), u8(0x22)), i)
    // pictures: BLIPs in the Pictures stream, FBSEs in the drawing group
    let at = 0
    const fbses = this.blips.map((b) => {
      const type = b[2] | (b[3] << 8)
      const bt = type === 0xf01e ? 6 : type === 0xf01d ? 5 : type === 0xf01f ? 7 : type === 0xf01a ? 2 : 3
      const r = rec(0xf007, cat(u8(bt), u8(bt), b.subarray(8, 24), u16(0xff), u32(b.length), u32(1), u32(at), u8(0), u8(0), u8(0), u8(0)), bt, 2)
      at += b.length
      return r
    })
    const dgg = box(0x040b, [box(0xf000, [rec(0xf006, cat(u32(nextSpid), u32(1), u32(0), u32(0))), ...(fbses.length ? [box(0xf001, fbses, fbses.length)] : [])])])
    const env = box(0x03f2, [
      box(0x07d5, this.fonts.map(font)),
      rec(0x0fa4, cfBytes({ font: 0, size: 18, color: schemeText(1) })),
      rec(0x0fa5, cat(u16(0), pfBytes({ align: 0 }))),
      rec(0x0fa9, [7, 0, 0, 0, 2, 0, 9, 4, 0, 0]),
      ...this.environmentStyles,
    ])
    const docKids: Bytes[] = [rec(0x03e9, cat(u32(this.size[0]), u32(this.size[1]), u32(4320), u32(5760), u32(1), u32(2), u32(ids.notesMaster), u32(0), u16(1), u16(0), u8(0), u8(0), u8(0), u8(1)), 0, 1)]
    if (this.links.length) {
      docKids.push(box(0x0409, [rec(0x040a, u32(this.links.length)), ...this.links.map((url, i) => box(0x0fd7, [rec(0x0fd3, u32(i + 1)), cstring(url, 0), cstring(url, 1)]))]))
    }
    // an empty document info list (view settings): some readers need the container
    docKids.push(env, dgg, list(1, sheets), box(0x07d0, []))
    if (this.headersFooters) {
      const hf = this.headersFooters
      docKids.push(box(0x0fd9, [rec(0x0fda, cat(u16(0), u16(hf.flags))), ...(hf.date ? [cstring(hf.date, 0)] : []), ...(hf.footer ? [cstring(hf.footer, 2)] : [])], 3))
    }
    docKids.push(list(0, slides))
    if (notes.length) docKids.push(list(2, notes))
    docKids.push(rec(0x03ea))
    const doc = box(0x03e8, docKids)

    // the stream: records, persist directory, edit
    const parts: Uint8Array[] = []
    let off = 0
    const offsets = new Map<number, number>()
    const put = (persist: number, body: Uint8Array) => {
      if (persist) offsets.set(persist, off)
      parts.push(body)
      off += body.length
    }
    put(1, doc)
    for (const sh of [...sheets, ...slides, ...notes]) put(sh.persist, sh.body)
    put(ids.notesMaster, notesMaster)
    const dirAt = off
    const entries = [...offsets].sort((a, b) => a[0] - b[0])
    put(0, rec(0x1772, cat(u32(1 | (entries.length << 20)), ...entries.map(([, o]) => u32(o)))))
    let editAt = off
    put(0, rec(0x0ff5, cat(u32(256), u16(0), u8(0), u8(3), u32(0), u32(dirAt), u32(1), u32(ids.next), u16(1), u16(0))))
    // incremental saves
    for (const e of this.edits) {
      const changed: [number, number][] = []
      for (const [persist, body] of e) {
        changed.push([persist, off])
        put(0, body)
      }
      const dir = off
      put(0, rec(0x1772, cat(...changed.flatMap(([p, o]) => [u32(p | (1 << 20)), u32(o)]))))
      const prev = editAt
      editAt = off
      put(0, rec(0x0ff5, cat(u32(256), u16(0), u8(0), u8(3), u32(prev), u32(dir), u32(1), u32(ids.next), u16(1), u16(0))))
    }
    const stream = cat(...parts)
    const user = 'Builder'
    const currentUser = rec(0x0ff6, cat(u32(0x14), u32(0xe391c05f), u32(editAt), u16(user.length), u16(0x03f4), u8(3), u8(0), u16(0), [...user].map((c) => c.charCodeAt(0)), u32(8)))
    return writeCfb({ 'Current User': currentUser, 'PowerPoint Document': stream, ...(this.blips.length ? { Pictures: cat(...this.blips) } : {}) })
  }
}

// ---------------------------------------------------------------- a default master

/** master text styles: title 44 pt centered, body levels with bullets */
export function defaultStyles(font = 0): Uint8Array[] {
  const bullets = ['•', '–', '•', '–', '»']
  return [
    masterStyle(TX.title, [{ pf: { align: 1, bullet: false, lineSpacing: 90 }, cf: { font, size: 44, color: schemeText(3), bold: false, italic: false, underline: false } }]),
    masterStyle(
      TX.body,
      [28, 24, 20, 18, 18].map((size, i) => ({
        pf: { bullet: true, bulletChar: bullets[i], align: 0, leftMargin: 270 + i * 288, indent: i * 288, spaceBefore: 20, lineSpacing: 100 },
        cf: { font, size, color: schemeText(1), bold: false, italic: false, underline: false },
      })),
    ),
  ]
}

/** a master with title and body placeholders, and footer placeholders (date, footer, slide number); `ph` their shape ids */
export function defaultMaster(extra: Uint8Array[] = []): MasterSpec & { ph: { title: number; body: number } } {
  const off: [number, number][] = [[P.fillBooleans, filled(false)], [P.lineBooleans, lined(false)]]
  const title = shape({ spt: 1, at: [432, 173, 4896, 720], ph: [1, 0], props: [...off, [P.anchorText, 1]], text: { type: TX.title, paras: ['Click to edit Master title style'] } })
  const body = shape({ spt: 1, at: [432, 1008, 4896, 2851], ph: [2, 1], props: off, text: { type: TX.body, paras: ['Click to edit Master text styles', { level: 1, runs: [{ text: 'Second level' }] }] } })
  return {
    styles: defaultStyles(),
    ph: { title: title.spid, body: body.spid },
    shapes: [
      ...extra,
      title,
      body,
      shape({ spt: 1, at: [432, 3936, 1344, 288], ph: [7, 2], props: off, text: { type: TX.other, paras: [{ runs: [{ text: '*', field: 'date', cf: { size: 12 } }] }] } }),
      shape({ spt: 1, at: [1968, 3936, 1824, 288], ph: [9, 3], props: off, text: { type: TX.other, paras: [{ pf: { align: 1 }, runs: [{ text: '*', field: 'footer', cf: { size: 12 } }] }] } }),
      shape({ spt: 1, at: [3984, 3936, 1344, 288], ph: [8, 4], props: off, text: { type: TX.other, paras: [{ pf: { align: 2 }, runs: [{ text: '*', field: 'slidenum', cf: { size: 12 } }] }] } }),
    ],
  }
}

// ---------------------------------------------------------------- the sample deck

/** a small PNG (a horizontal gradient) */
export { png } from './iwork-builder'
import { png } from './iwork-builder'

/**
 * The sample presentation (modeled on playground/public/samples/sample.pptx): a title slide on a title master,
 * bullets on levels with character formatting, a numbered list and a link, shapes (rounded rectangle, oval,
 * arrow with a gradient, line with an arrowhead, freeform, arc, rotated text), a picture, a table (a group of
 * cells, as PowerPoint stores tables), a hidden slide with its own background, notes, a comment, footers with slide
 * numbers.
 */
export function samplePpt(only?: number): Uint8Array<ArrayBuffer> {
  const b = new PptBuilder()
  b.only = only
  const georgia = b.font('Georgia')
  const wingdings = b.font('Wingdings')
  const band = shape({ spt: 1, at: [0, 4176, 5760, 144], props: [[P.fillColor, oaScheme(4)], [P.lineBooleans, lined(false)]] })
  const dm = defaultMaster([band])
  const main = b.master({ ...dm, scheme: SCHEME })
  b.headersFooters = { flags: 0x08 | 0x20, footer: 'office-engine sample' }
  const ctrTitle = shape({ spt: 1, at: [432, 1152, 4896, 864], ph: [3, 0], props: [[P.fillBooleans, filled(false)], [P.lineBooleans, lined(false)], [P.anchorText, 2]], text: { type: TX.centerTitle, paras: ['Click to edit Master title style'] } })
  const subTitle = shape({ spt: 1, at: [864, 2160, 4032, 1008], ph: [4, 1], props: [[P.fillBooleans, filled(false)], [P.lineBooleans, lined(false)]], text: { type: TX.centerBody, paras: ['Click to edit Master subtitle style'] } })
  const titleMaster = b.master({
    titleOf: main,
    flags: 0,
    styles: [],
    background: [[P.fillType, 7], [P.fillColor, oa('1D4ED8')], [P.fillBackColor, oa('0F172A')], [P.fillFocus, 100], [P.fillBooleans, filled(true)]],
    shapes: [ctrTitle, subTitle],
  })
  const title = (outline: number) => shape({ spt: 1, at: [432, 173, 4896, 720], ph: [0x0d, 0], master: dm.ph.title, outline })
  const white = rgb('FFFFFF')
  const link = b.link('https://example.com/office-engine')

  // 1. title slide
  b.slide({
    master: titleMaster,
    layout: 0,
    texts: [
      { type: TX.centerTitle, paras: [{ pf: { align: 1 }, runs: [{ text: 'office-engine', cf: { bold: true, size: 54, color: white } }] }] },
      { type: TX.centerBody, paras: [{ pf: { align: 1, bullet: false }, runs: [{ text: 'Progressive Office viewer for the browser', cf: { size: 24, color: rgb('CBD5E1') } }] }] },
    ],
    shapes: [
      shape({ spt: 1, at: [432, 1152, 4896, 864], ph: [0x0f, 0], master: ctrTitle.spid, outline: 0 }),
      shape({ spt: 1, at: [864, 2160, 4032, 1008], ph: [0x10, 1], master: subTitle.spid, outline: 1 }),
    ],
  })

  // 2. bullets
  b.slide({
    texts: [
      { type: TX.title, paras: ['Why it is fast'] },
      {
        type: TX.body,
        paras: [
          { runs: [{ text: 'Lazy ZIP: only the ' }, { text: 'central directory', cf: { bold: true } }, { text: ' is read up front' }] },
          { level: 1, runs: [{ text: 'Entries are inflated on demand', cf: { italic: true } }] },
          { runs: [{ text: 'Streaming XML parser in a ' }, { text: 'Web Worker', cf: { color: rgb('C00000'), underline: true } }] },
          { level: 1, pf: { bulletChar: 'n', bulletFont: wingdings, bulletColor: rgb('2563EB') }, runs: [{ text: 'Blocks, rows and slides arrive progressively' }] },
          { runs: [{ text: 'Read more at ' }, { text: 'example.com', link }] },
        ],
      },
    ],
    shapes: [
      title(0),
      shape({ spt: 1, at: [432, 1008, 2900, 2851], ph: [0x0e, 1], master: dm.ph.body, outline: 1 }),
      shape({
        spt: 202,
        at: [3456, 1008, 1872, 1440],
        props: [[P.fillBooleans, filled(false)], [P.lineBooleans, lined(false)]],
        ppt9: style9([{ scheme: 3, start: 1 }]),
        text: {
          type: TX.other,
          paras: ['Open', 'Parse', 'Render'].map((t) => ({ pf: { bullet: true, leftMargin: 216, indent: 0 }, runs: [{ text: t, cf: { size: 20, pp9rt: 0 } }] })),
        },
      }),
    ],
    notes: { type: TX.notes, paras: ['Speaker notes: mention the worker and the lazy ZIP reader.'] },
    extra: [progTag('___PPT10', [comment10('Ann Lee', 'AL', 'Add the startup time here?', [2026, 1, 15, 9, 30, 0], [3456, 2600])])],
  })

  // 3. shapes and a picture
  const pib = b.picture(png(160, 96))
  const triangle = msoArray([[0, 100], [50, 0], [100, 100]].map(([x, y]) => [...u16(x), ...u16(y)]), 4)
  const segments = msoArray([0x4000, 0x0002, 0x6001, 0x8000].map((v) => u16(v)), 2)
  b.slide({
    texts: [{ type: TX.title, paras: ['Shapes & images'] }],
    shapes: [
      title(0),
      shape({ spt: 2, at: [432, 1152, 1152, 720], props: [[P.fillColor, oa('2563EB')], [P.lineBooleans, lined(false)], [P.adjustValue, 3600], [P.anchorText, 1]], text: { type: TX.other, paras: [{ pf: { align: 1 }, runs: [{ text: 'Rounded', cf: { color: white, bold: true } }] }] } }),
      shape({ spt: 3, at: [1872, 1152, 1152, 720], props: [[P.fillColor, oa('10B981')], [P.lineColor, oa('065F46')], [P.lineWidth, 25400], [P.anchorText, 1]], text: { type: TX.other, paras: [{ pf: { align: 1 }, runs: [{ text: 'Oval', cf: { color: white } }] }] } }),
      shape({
        spt: 13,
        at: [3312, 1152, 1440, 720],
        props: [[P.fillType, 7], [P.fillColor, oa('F59E0B')], [P.fillBackColor, oa('FEF3C7')], [P.fillFocus, 100], [P.lineBooleans, lined(false)], [P.anchorText, 1]],
        text: { type: TX.other, paras: [{ pf: { align: 1 }, runs: [{ text: 'Arrow' }] }] },
      }),
      shape({ spt: 202, at: [432, 2160, 2304, 576], props: [[P.fillBooleans, filled(false)], [P.lineBooleans, lined(false)]], text: { type: TX.other, paras: [{ runs: [{ text: 'A free text box with ' }, { text: 'colored bold text', cf: { bold: true, color: rgb('7C3AED'), font: georgia } }] }] } }),
      shape({ spt: 75, at: [3312, 2160, 1440, 864], props: [[P.pib, pib], [P.cropFromLeft, fixed(0.1)]] }),
      shape({ spt: 20, at: [432, 3024, 1728, 0], props: [[P.lineColor, oa('DC2626')], [P.lineWidth, 19050], [P.lineDashing, 6], [P.lineEndArrowhead, 1]] }),
      shape({ spt: 0, at: [2304, 2880, 576, 576], props: [[P.fillColor, oa('A855F7')], [P.lineBooleans, lined(false)], [P.geoRight, 100], [P.geoBottom, 100]], complex: [[P.pVertices, triangle], [P.pSegmentInfo, segments]] }),
      shape({ spt: 19, at: [3024, 3168, 432, 432], props: [[P.fillBooleans, filled(false)], [P.lineColor, oa('0EA5E9')], [P.lineWidth, 28575]] }),
      shape({ spt: 202, at: [4176, 3312, 1152, 288], props: [[P.rotation, fixed(-15)], [P.fillBooleans, filled(false)], [P.lineBooleans, lined(false)]], text: { type: TX.other, paras: ['Rotated text'] } }),
    ],
  })

  // 4. a table: a group of cells
  const rows = [
    ['Metric', '2024', '2025', '2026'],
    ['Users', '380', '750', '424'],
    ['Revenue', '746', '55', '781'],
    ['Latency', '130', '209', '885'],
  ]
  const cells = rows.flatMap((r, ri) =>
    r.map((t, ci) =>
      shape({
        spt: 1,
        child: true,
        at: [ci * 1000, ri * 400, 1000, 400],
        props: [[P.fillColor, ri === 0 ? oa('1D4ED8') : ri % 2 ? oa('EFF6FF') : oa('FFFFFF')], [P.lineColor, oa('94A3B8')], [P.lineWidth, 9525], [P.anchorText, 1]],
        text: { type: TX.other, paras: [{ pf: { align: ci ? 2 : 0 }, runs: [{ text: t, cf: { size: 16, bold: ri === 0, color: ri === 0 ? white : rgb('111827') } }] }] },
      }),
    ),
  )
  b.slide({
    texts: [{ type: TX.title, paras: ['Table'] }],
    shapes: [title(0), group([720, 1152, 4320, 1728], [0, 0, 4000, 1600], cells)],
  })

  // 5. a hidden slide with its own background and colors
  b.slide({
    hidden: true,
    flags: 1,
    background: [[P.fillColor, oa('FEF3C7')], [P.fillBooleans, filled(true)]],
    scheme: ['FEF3C7', '78350F', '808080', '92400E', 'F59E0B', 'D97706', '0563C1', '954F72'],
    texts: [{ type: TX.title, paras: ['Hidden slide'] }],
    shapes: [title(0)],
  })
  return b.build()
}
