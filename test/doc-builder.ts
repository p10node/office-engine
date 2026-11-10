/**
 * Writes small Word 97 binary documents (.doc, MS-DOC) for tests and the sample fixture: a compound file with the
 * WordDocument stream (FIB, text, formatted disk pages, section properties), the 1Table stream (stylesheet, piece
 * table, fonts, lists, sections, headers, notes, fields, drawings) and the Data stream (inline pictures).
 */

// ---------------------------------------------------------------- bytes

class Buf {
  b: number[] = []
  get length() {
    return this.b.length
  }
  u8(v: number) {
    this.b.push(v & 0xff)
    return this
  }
  u16(v: number) {
    return this.u8(v).u8(v >> 8)
  }
  u32(v: number) {
    return this.u16(v & 0xffff).u16((v >>> 16) & 0xffff)
  }
  bytes(a: ArrayLike<number>) {
    for (let i = 0; i < a.length; i++) this.b.push(a[i] & 0xff)
    return this
  }
  utf16(s: string) {
    for (let i = 0; i < s.length; i++) this.u16(s.charCodeAt(i))
    return this
  }
  pad(n: number) {
    while (this.b.length % n) this.b.push(0)
    return this
  }
  zeros(n: number) {
    for (let i = 0; i < n; i++) this.b.push(0)
    return this
  }
  set32(at: number, v: number) {
    for (let k = 0; k < 4; k++) this.b[at + k] = (v >>> (8 * k)) & 0xff
  }
  done(): Uint8Array {
    return Uint8Array.from(this.b)
  }
}

// ---------------------------------------------------------------- sprms

/** a sprm with its operand (sized by the opcode's spra; variable operands get their size byte) */
export function sprm(op: number, v: number | number[] = 0): number[] {
  const out = [op & 0xff, op >> 8]
  const spra = op >> 13
  const n = typeof v === 'number' ? v : 0
  switch (spra) {
    case 0:
    case 1:
      return [...out, n & 0xff]
    case 2:
    case 4:
    case 5:
      return [...out, n & 0xff, (n >> 8) & 0xff]
    case 3:
      return [...out, n & 0xff, (n >> 8) & 0xff, (n >> 16) & 0xff, (n >>> 24) & 0xff]
    case 7:
      return [...out, n & 0xff, (n >> 8) & 0xff, (n >> 16) & 0xff]
    default: {
      const a = v as number[]
      if (op === 0xd608) return [...out, (a.length + 1) & 0xff, (a.length + 1) >> 8, ...a]
      return [...out, a.length, ...a]
    }
  }
}

const le16 = (v: number) => [v & 0xff, (v >> 8) & 0xff]
const le32 = (v: number) => [v & 0xff, (v >> 8) & 0xff, (v >> 16) & 0xff, (v >>> 24) & 0xff]
/** a COLORREF from hex */
export const cvOf = (hex: string) => [parseInt(hex.slice(0, 2), 16), parseInt(hex.slice(2, 4), 16), parseInt(hex.slice(4, 6), 16), 0]
/** a Brc80: width (eighths of a point), type, ico */
export const brc80 = (w: number, type = 1, ico = 1) => [w, type, ico, 0]

/** character sprms */
export const C = {
  bold: (on = true) => sprm(0x0835, on ? 1 : 0),
  italic: (on = true) => sprm(0x0836, on ? 1 : 0),
  strike: () => sprm(0x0837, 1),
  dstrike: () => sprm(0x2a53, 1),
  caps: () => sprm(0x083b, 1),
  smallCaps: () => sprm(0x083a, 1),
  vanish: () => sprm(0x083c, 1),
  underline: (kul = 1) => sprm(0x2a3e, kul),
  size: (pt: number) => sprm(0x4a43, Math.round(pt * 2)),
  font: (ftc: number) => [...sprm(0x4a4f, ftc), ...sprm(0x4a51, ftc)],
  ico: (n: number) => sprm(0x2a42, n),
  color: (hex: string) => sprm(0x6870, (cvOf(hex)[0] | (cvOf(hex)[1] << 8) | (cvOf(hex)[2] << 16)) >>> 0),
  highlight: (ico: number) => sprm(0x2a0c, ico),
  iss: (v: 1 | 2) => sprm(0x2a48, v),
  spacing: (tw: number) => sprm(0x8840, tw),
  shd: (fill: string) => sprm(0xca71, [...cvOf('000000').slice(0, 3), 0xff, ...cvOf(fill), 0, 0]),
  spec: () => sprm(0x0855, 1),
  pic: (fc: number) => sprm(0x6a03, fc),
  style: (istd: number) => sprm(0x4a30, istd),
  symbol: (ftc: number, ch: number) => sprm(0x6a09, (ftc | (ch << 16)) >>> 0),
  rmIns: (ibst: number) => [...sprm(0x0801, 1), ...sprm(0x4804, ibst)],
  rmDel: (ibst: number) => [...sprm(0x0800, 1), ...sprm(0x4863, ibst)],
}

/** paragraph sprms */
export const P = {
  jc: (v: number) => [...sprm(0x2403, v), ...sprm(0x2461, v)],
  left: (tw: number) => [...sprm(0x840f, tw), ...sprm(0x845e, tw)],
  right: (tw: number) => [...sprm(0x840e, tw), ...sprm(0x845d, tw)],
  first: (tw: number) => [...sprm(0x8411, tw), ...sprm(0x8460, tw)],
  before: (tw: number) => sprm(0xa413, tw),
  after: (tw: number) => sprm(0xa414, tw),
  line: (v: number, mult = true) => sprm(0x6412, ((v & 0xffff) | ((mult ? 1 : 0) << 16)) >>> 0),
  keep: () => sprm(0x2405, 1),
  keepNext: () => sprm(0x2406, 1),
  pageBreakBefore: () => sprm(0x2407, 1),
  outline: (lvl: number) => sprm(0x2640, lvl),
  list: (ilfo: number, ilvl = 0) => [...sprm(0x460b, ilfo), ...sprm(0x260a, ilvl)],
  tabs: (tabs: [number, number, number][]) => sprm(0xc60d, [0, tabs.length, ...tabs.flatMap(([pos]) => le16(pos)), ...tabs.map(([, jc, tlc]) => jc | (tlc << 3))]),
  border: (side: 'top' | 'left' | 'bottom' | 'right', b: number[]) => sprm({ top: 0x6424, left: 0x6425, bottom: 0x6426, right: 0x6427 }[side], (b[0] | (b[1] << 8) | (b[2] << 16) | (b[3] << 24)) >>> 0),
  shd: (fill: string) => sprm(0xc64d, [0, 0, 0, 0xff, ...cvOf(fill), 0, 0]),
  inTable: () => sprm(0x2416, 1),
  ttp: () => sprm(0x2417, 1),
  itap: (n: number) => sprm(0x6649, n),
  innerCell: () => sprm(0x244b, 1),
  innerTtp: () => sprm(0x244c, 1),
}

/** section sprms */
export const S = {
  bkc: (v: number) => sprm(0x3009, v),
  page: (w: number, h: number) => [...sprm(0xb01f, w), ...sprm(0xb020, h)],
  landscape: () => sprm(0x301d, 2),
  margins: (top: number, right: number, bottom: number, left: number) => [...sprm(0x9023, top), ...sprm(0xb022, right), ...sprm(0x9024, bottom), ...sprm(0xb021, left)],
  cols: (n: number, space = 720) => [...sprm(0x500b, n - 1), ...sprm(0x900c, space)],
  titlePage: () => sprm(0x300a, 1),
  headerDist: (top: number, bottom: number) => [...sprm(0xb017, top), ...sprm(0xb018, bottom)],
}

// ---------------------------------------------------------------- document model

/** a run: text and its character sprms (a footnote's text, a floating picture) */
export interface Run {
  text: string
  c?: number[]
  note?: string
  float?: { rect: [number, number, number, number]; pib: number }
  comment?: { text: string; author: string; initials: string; range: number }
}

export interface Cell {
  paras: (string | Run[])[]
  /** horizontal merge: 1 first cell, 2 merged into the one before */
  hm?: number
  /** vertical merge: 3 restart, 1 continue */
  vm?: number
  shd?: string
  /** nested table (written after the paragraphs) */
  table?: TableSpec
}

export interface TableSpec {
  widths: number[]
  rows: Cell[][]
  borders?: boolean
  header?: boolean
  rowHeight?: number
}

export interface StyleSpec {
  istd: number
  sti: number
  name: string
  /** 1 paragraph, 2 character */
  stk: 1 | 2
  base?: number
  next?: number
  papx?: number[]
  chpx?: number[]
}

export interface LevelSpec {
  nfc: number
  text: string
  /** level placeholders: positions (1-based) in `text` */
  nums?: number[]
  left: number
  hang: number
  font?: number
}

interface Para {
  /** CP after the paragraph mark */
  end: number
  istd: number
  papx: number[]
}

interface ChpRun {
  cp0: number
  cp1: number
  chpx: number[]
}

interface Story {
  text: string[]
  paras: Para[]
  chp: ChpRun[]
  fields: [number, number, number][]
}

const newStory = (): Story => ({ text: [], paras: [], chp: [], fields: [] })

export const STYLES = {
  normal: 0,
  heading1: 1,
  heading2: 2,
  heading3: 3,
  defaultFont: 10,
  title: 15,
  listBullet: 16,
  listNumber: 17,
  header: 18,
  footer: 19,
  footnoteText: 20,
  footnoteRef: 21,
  caption: 22,
  hyperlink: 23,
  strong: 24,
}

/** Builds a Word 97 document. Text goes into the main story; notes and headers into theirs. */
export class DocBuilder {
  fonts = ['Times New Roman', 'Symbol', 'Arial', 'Courier New', 'Wingdings']
  styles: StyleSpec[] = [
    { istd: 0, sti: 0, name: 'Normal', stk: 1, papx: [...P.after(120)], chpx: [...C.size(11), ...C.font(2)] },
    { istd: 1, sti: 1, name: 'heading 1', stk: 1, base: 0, next: 0, papx: [...P.keepNext(), ...P.before(240), ...P.after(120), ...P.outline(0)], chpx: [...C.bold(), ...C.size(16), ...C.color('1F3864')] },
    { istd: 2, sti: 2, name: 'heading 2', stk: 1, base: 0, next: 0, papx: [...P.keepNext(), ...P.before(200), ...P.after(80), ...P.outline(1)], chpx: [...C.bold(), ...C.size(13), ...C.color('2F5496')] },
    { istd: 3, sti: 3, name: 'heading 3', stk: 1, base: 0, next: 0, papx: [...P.keepNext(), ...P.outline(2)], chpx: [...C.bold(), ...C.italic()] },
    { istd: 10, sti: 65, name: 'Default Paragraph Font', stk: 2 },
    { istd: 15, sti: 62, name: 'Title', stk: 1, base: 0, next: 0, papx: [...P.after(240)], chpx: [...C.size(26), ...C.color('17365D')] },
    { istd: 16, sti: 48, name: 'List Bullet', stk: 1, base: 0, papx: [...P.after(60)] },
    { istd: 17, sti: 49, name: 'List Number', stk: 1, base: 0, papx: [...P.after(60)] },
    { istd: 18, sti: 31, name: 'header', stk: 1, base: 0, papx: [...P.after(0), ...P.tabs([[4680, 1, 0], [9360, 2, 0]])], chpx: [...C.size(9)] },
    { istd: 19, sti: 32, name: 'footer', stk: 1, base: 0, papx: [...P.after(0), ...P.tabs([[4680, 1, 0], [9360, 2, 0]])], chpx: [...C.size(9)] },
    { istd: 20, sti: 29, name: 'footnote text', stk: 1, base: 0, papx: [...P.after(0)], chpx: [...C.size(9)] },
    { istd: 21, sti: 38, name: 'footnote reference', stk: 2, base: 10, chpx: [...C.iss(1)] },
    { istd: 22, sti: 34, name: 'caption', stk: 1, base: 0, next: 0, papx: [...P.jc(1)], chpx: [...C.italic(), ...C.size(9)] },
    { istd: 23, sti: 85, name: 'Hyperlink', stk: 2, base: 10, chpx: [...C.color('0563C1'), ...C.underline(1)] },
    { istd: 24, sti: 87, name: 'Strong', stk: 2, base: 10, chpx: [...C.bold()] },
  ]
  lists: { lsid: number; levels: LevelSpec[] }[] = []
  /** section breaks: CP after each section's last character, with its sprms */
  sections: { end: number; sepx: number[] }[] = []
  private main = newStory()
  private ftn = newStory()
  private hdd = newStory()
  private atn = newStory()
  private comments: { cp: number; text: string; author: string; initials: string; range: number }[] = []
  /** header stories of the (single) header set: [even hdr, odd hdr, even ftr, odd ftr, first hdr, first ftr] */
  private hdrStories: (string | Run[] | undefined)[][] = []
  private footRefs: number[] = []
  private ftnStarts: number[] = []
  private data = new Buf()
  private blips: Uint8Array[] = []
  private floats: { cp: number; spid: number; rect: [number, number, number, number]; pib: number }[] = []
  private nextSpid = 1026
  /** write the main text as a 8-bit piece followed by a UTF-16 piece (with a piece sprm) */
  splitPieces = false
  /** set fEncrypted */
  encrypted = false
  /** FIB version written */
  nFib = 0xc1
  authors: string[] = []
  private cur: Story = this.main
  private pendingNotes: [number, string][] = []

  constructor() {
    this.data.zeros(16)
  }

  private cp(): number {
    return this.cur.text.length
  }

  /** appends characters with character sprms */
  private chars(s: string, chpx: number[] = []) {
    const cp0 = this.cp()
    for (const ch of s) this.cur.text.push(...ch.split(''))
    if (chpx.length) this.cur.chp.push({ cp0, cp1: this.cp(), chpx })
  }

  /** a paragraph of runs, with a style and paragraph sprms */
  p(runs: string | Run[], istd = 0, papx: number[] = [], mark = '\r', markChpx: number[] = []) {
    const list: Run[] = typeof runs === 'string' ? [{ text: runs }] : runs
    list.forEach((r, k) => {
      const cp = this.cp()
      if (r.note) this.pendingNotes.push([cp, r.note])
      if (r.float) this.floats.push({ cp, spid: this.nextSpid++, rect: r.float.rect, pib: r.float.pib })
      if (r.comment) this.comments.push({ cp, ...r.comment })
      const instr = list[k + 1]?.text ?? ''
      if (r.text === '\u0013') this.cur.fields.push([cp, 0x13, /HYPERLINK/.test(instr) ? 88 : /NUMPAGES/.test(instr) ? 26 : /PAGE/.test(instr) ? 33 : /DATE/.test(instr) ? 31 : 0])
      else if (r.text === '\u0014') this.cur.fields.push([cp, 0x14, 0xff])
      else if (r.text === '\u0015') this.cur.fields.push([cp, 0x15, 0x80])
      this.chars(r.text, r.c)
    })
    this.chars(mark, markChpx)
    this.cur.paras.push({ end: this.cp(), istd, papx })
  }

  /** a field: begin, instruction, separator, result, end */
  field(instr: string, result: string | Run[], c: number[] = []): Run[] {
    return [{ text: '\u0013', c: [...C.spec(), ...c] }, { text: ` ${instr} `, c }, { text: '\u0014', c: [...C.spec(), ...c] }, ...(typeof result === 'string' ? [{ text: result, c }] : result), { text: '\u0015', c: [...C.spec(), ...c] }]
  }

  hyperlink(url: string, text: string): Run[] {
    return this.field(`HYPERLINK "${url}"`, [{ text, c: C.style(STYLES.hyperlink) }])
  }

  /** a footnote: the reference run in the main text; the note's text goes into the footnote story */
  footnote(text: string): Run {
    return { text: '\u0002', c: [...C.spec(), ...C.style(STYLES.footnoteRef), ...C.iss(1)], note: text }
  }

  /** a comment on the `range` characters before it (0: a point) */
  comment(text: string, author: string, initials: string, range = 0): Run {
    return { text: '\u0005', c: C.spec(), comment: { text, author, initials, range } }
  }

  /** an inline PNG picture (size in twips) */
  picture(png: Uint8Array, w: number, h: number): Run {
    const fc = this.data.length
    const blipRec = new Buf().u16(0x6e0 << 4).u16(0xf01e).u32(17 + png.length).zeros(16).u8(0xff).bytes(png).done()
    const fbse = new Buf().u16(0x2 | (6 << 4)).u16(0xf007).u32(36 + blipRec.length).u8(6).u8(6).zeros(16).u16(0xff).u32(blipRec.length).u32(1).u32(0).u8(0).u8(0).u8(0).u8(0).bytes(blipRec).done()
    const fopt = new Buf().u16(0x3 | (1 << 4)).u16(0xf00b).u32(6).u16(0x4104).u32(1).done()
    const fsp = new Buf().u16(0x2 | (75 << 4)).u16(0xf00a).u32(8).u32(1025).u32(0xa00).done()
    const sp = new Buf().u16(0xf).u16(0xf004).u32(fsp.length + fopt.length).bytes(fsp).bytes(fopt).done()
    const total = 0x44 + sp.length + fbse.length
    const d = this.data
    d.u32(total).u16(0x44).u16(0x64).u16(0).u16(0).u16(0).zeros(14)
    d.u16(w).u16(h).u16(1000).u16(1000).zeros(8).u8(0).u8(0).zeros(16).u16(0).u16(0).u16(0)
    d.bytes(sp).bytes(fbse)
    return { text: '\u0001', c: [...C.spec(), ...C.pic(fc)] }
  }

  /** a floating PNG picture anchored at its 0x08 character (rect in twips from the column / paragraph) */
  floating(png: Uint8Array, rect: [number, number, number, number]): Run {
    this.blips.push(png)
    return { text: '\u0008', c: C.spec(), float: { rect, pib: this.blips.length } }
  }

  /** a table (nested tables in cells are written at depth + 1) */
  table(t: TableSpec, depth = 1) {
    const centers = [0]
    for (const w of t.widths) centers.push(centers[centers.length - 1] + w)
    const inner = depth > 1
    const tableP = inner ? [...P.inTable(), ...P.itap(depth)] : [...P.inTable()]
    for (const row of t.rows) {
      for (const cell of row) {
        cell.paras.forEach((para, k) => {
          const last = k === cell.paras.length - 1 && !cell.table
          const mark = last ? (inner ? '\r' : '\u0007') : '\r'
          this.p(para, 0, [...tableP, ...(last && inner ? P.innerCell() : [])], mark)
        })
        if (cell.table) {
          this.table(cell.table, depth + 1)
          this.p('', 0, [...tableP, ...(inner ? P.innerCell() : [])], inner ? '\r' : '\u0007')
        }
      }
      // the row end: its table properties
      const tc: number[] = []
      for (const cell of row) {
        const f = (cell.hm ?? 0) | ((cell.vm ?? 0) << 5)
        const b = t.borders === false ? [0, 0, 0, 0] : brc80(4)
        tc.push(...le16(f), ...le16(0), ...b, ...b, ...b, ...b)
      }
      const def = [row.length, ...centers.slice(0, row.length + 1).flatMap((x) => le16(x)), ...tc]
      let tap = [...sprm(0x9602, 108), ...sprm(0xd608, def)]
      if (t.borders !== false) tap.push(...sprm(0xd605, [...brc80(4), ...brc80(4), ...brc80(4), ...brc80(4), ...brc80(4), ...brc80(4)]))
      if (t.header && row === t.rows[0]) tap.push(...sprm(0x3404, 1))
      if (t.rowHeight) tap.push(...sprm(0x9407, t.rowHeight))
      const shd = row.map((c) => [...cvOf('000000').slice(0, 3), 0xff, ...(c.shd ? cvOf(c.shd) : [0, 0, 0, 0xff]), 0, 0]).flat()
      if (row.some((c) => c.shd)) tap.push(...sprm(0xd612, shd))
      tap = [...tableP, ...(inner ? P.innerTtp() : P.ttp()), ...tap]
      this.p('', 0, tap, inner ? '\r' : '\u0007')
    }
  }

  /** a list definition (levels); returns its list instance index (ilfo, 1-based) */
  list(levels: LevelSpec[]): number {
    // a list has 1 level (simple) or 9
    const all = levels.length === 1 ? levels : Array.from({ length: 9 }, (_, k) => levels[k] ?? { ...levels[levels.length - 1], left: levels[levels.length - 1].left + 360 * (k - levels.length + 1) })
    this.lists.push({ lsid: 0x1000 + this.lists.length, levels: all })
    return this.lists.length
  }

  /** ends the current section (the next one starts after a section break) */
  sectionBreak(sepx: number[]) {
    this.chars('\u000c')
    this.cur.paras.push({ end: this.cp(), istd: 0, papx: [] })
    this.sections.push({ end: this.cp(), sepx })
  }

  /** headers / footers of the first section: [even header, odd header, even footer, odd footer, first header, first footer] */
  headers(stories: (string | Run[] | undefined)[]) {
    this.hdrStories.push(stories)
  }

  // ---------------------------------------------------------------- writing

  build(lastSepx: number[] = []): Uint8Array {
    // footnotes: collect the note texts of the main text's references
    const mainRuns = this.main
    const notes: string[] = []
    // the main story was written run by run; notes were attached to runs
    this.collectNotes(notes)
    this.cur = this.ftn
    for (const n of notes) {
      this.ftnStarts.push(this.cp())
      this.p([{ text: '\u0002', c: [...C.spec(), ...C.style(STYLES.footnoteRef)] }, { text: ' ' + n }], STYLES.footnoteText)
    }
    if (notes.length) {
      this.ftnStarts.push(this.cp())
      this.p('')
    }
    // headers
    this.cur = this.hdd
    const hddCps: number[] = []
    const sets = this.hdrStories.length ? this.hdrStories : []
    if (sets.length) {
      for (let k = 0; k < 6; k++) hddCps.push(this.cp())
      for (const set of sets)
        for (let k = 0; k < 6; k++) {
          hddCps.push(this.cp())
          const s = set[k]
          if (s === undefined) continue
          this.p(s, k < 2 || k === 4 ? STYLES.header : STYLES.footer)
          this.p('', 0)
        }
      hddCps.push(this.cp())
      this.p('')
      hddCps.push(this.cp())
    }
    // comments
    this.cur = this.atn
    const atnStarts: number[] = []
    for (const c of this.comments) {
      atnStarts.push(this.cp())
      this.p([{ text: '\u0005', c: C.spec() }, { text: c.text }])
    }
    if (this.comments.length) {
      atnStarts.push(this.cp())
      this.p('')
    }
    this.cur = mainRuns

    const ccpText = this.main.text.length
    const ccpFtn = this.ftn.text.length
    const ccpHdd = this.hdd.text.length
    const ccpAtn = this.atn.text.length
    const all = [...this.main.text, ...this.ftn.text, ...this.hdd.text, ...this.atn.text]
    const offFtn = ccpText
    const offHdd = ccpText + ccpFtn
    const offAtn = offHdd + ccpHdd
    const shift = (st: Story, off: number) => ({
      paras: st.paras.map((p) => ({ ...p, end: p.end + off })),
      chp: st.chp.map((r) => ({ ...r, cp0: r.cp0 + off, cp1: r.cp1 + off })),
    })
    const subs = [shift(this.ftn, offFtn), shift(this.hdd, offHdd), shift(this.atn, offAtn)]
    const paras = [...this.main.paras, ...subs.flatMap((x) => x.paras)]
    const chps = [...this.main.chp, ...subs.flatMap((x) => x.chp)]

    // ---- WordDocument: FIB, text, FKPs, SEPXs
    const wd = new Buf()
    wd.zeros(1024)
    const textAt = 1024
    // pieces: one UTF-16 piece, or an 8-bit piece then a UTF-16 one
    const pieces: { cp0: number; cp1: number; fc: number; comp: boolean; prm: number }[] = []
    const fcOf: number[] = []
    if (this.splitPieces) {
      // a UTF-16 piece, then an 8-bit one right after it (with a piece sprm)
      const half = Math.floor(all.length / 2)
      if (!all.slice(half).every((ch) => ch.charCodeAt(0) < 0x80)) throw new Error('split pieces need ASCII in the second half')
      for (let i = 0; i < half; i++) {
        fcOf.push(wd.length)
        wd.u16(all[i].charCodeAt(0))
      }
      const fc2 = wd.length
      for (let i = half; i < all.length; i++) {
        fcOf.push(wd.length)
        wd.u8(all[i].charCodeAt(0))
      }
      pieces.push({ cp0: 0, cp1: half, fc: textAt, comp: false, prm: 0 }, { cp0: half, cp1: all.length, fc: fc2, comp: true, prm: 1 })
    } else {
      for (let i = 0; i < all.length; i++) {
        fcOf.push(wd.length)
        wd.u16(all[i].charCodeAt(0))
      }
      pieces.push({ cp0: 0, cp1: all.length, fc: textAt, comp: false, prm: 0 })
    }
    const fcEnd = wd.length
    const fcAt = (cp: number) => (cp < fcOf.length ? fcOf[cp] : fcEnd)
    wd.pad(512)

    // CHPX runs over every character (gaps get no properties)
    const sorted = [...chps].sort((a, b) => a.cp0 - b.cp0)
    const runs: { fc0: number; fc1: number; g: number[] }[] = []
    let at = 0
    for (const r of sorted) {
      if (r.cp0 > at) runs.push({ fc0: fcAt(at), fc1: fcAt(r.cp0), g: [] })
      if (r.cp0 >= at) runs.push({ fc0: fcAt(r.cp0), fc1: fcAt(r.cp1), g: r.chpx })
      at = Math.max(at, r.cp1)
    }
    if (at < all.length) runs.push({ fc0: fcAt(at), fc1: fcEnd, g: [] })
    // runs never cross the piece boundary (the pieces are back to back in the stream)
    const split = (list: { fc0: number; fc1: number; g: number[] }[]) => {
      if (!this.splitPieces) return list
      const b = pieces[1].fc
      return list.flatMap((r) => (r.fc0 < b && r.fc1 > b ? [{ ...r, fc1: b }, { ...r, fc0: b }] : [r]))
    }
    const chpPages = this.fkps(wd, split(runs), false)
    // PAPX runs: one per paragraph
    const pruns: { fc0: number; fc1: number; g: number[] }[] = []
    let start = 0
    for (const p of paras) {
      pruns.push({ fc0: fcAt(start), fc1: fcAt(p.end), g: [...le16(p.istd), ...p.papx] })
      start = p.end
    }
    const papPages = this.fkps(wd, split(pruns), true)
    // SEPXs
    const sects = [...this.sections, { end: ccpText, sepx: lastSepx }]
    const sepxAt = sects.map((s) => {
      wd.pad(2)
      const fc = wd.length
      wd.u16(s.sepx.length).bytes(s.sepx)
      return fc
    })

    // ---- 1Table
    const tb = new Buf()
    const fc: Record<number, [number, number]> = {}
    const put = (i: number, bytes: number[] | Uint8Array) => {
      tb.pad(2)
      fc[i] = [tb.length, bytes.length]
      tb.bytes(bytes)
    }
    put(1, this.stsh())
    // PlcBteChpx / PlcBtePapx
    const bte = (pages: { fc0: number; pn: number }[], last: number) => [...pages.flatMap((p) => le32(p.fc0)), ...le32(last), ...pages.flatMap((p) => le32(p.pn))]
    put(12, bte(chpPages, fcEnd))
    put(13, bte(papPages, fcEnd))
    // SttbfFfn
    const ffn = new Buf().u16(this.fonts.length).u16(0)
    for (const name of this.fonts) {
      const symbol = name === 'Symbol' || name === 'Wingdings'
      const ffid = symbol ? 0x56 : name === 'Courier New' ? 0x35 : name === 'Arial' ? 0x26 : 0x16
      const body = new Buf().u8(ffid).u16(400).u8(symbol ? 2 : 0).u8(0).zeros(10).zeros(24).utf16(name).u16(0).done()
      ffn.u8(body.length).bytes(body)
    }
    put(15, ffn.done())
    // PlcfSed
    const sed = new Buf()
    sed.u32(0)
    for (const s of sects) sed.u32(s.end)
    sects.forEach((_, k) => sed.u16(0).u32(sepxAt[k]).u16(0).u32(0xffffffff))
    put(6, sed.done())
    // PlcfHdd
    if (hddCps.length) put(11, hddCps.flatMap((c) => le32(c)))
    // footnotes
    if (this.footRefs.length) {
      put(2, [...this.footRefs.flatMap((c) => le32(c)), ...le32(ccpText), ...this.footRefs.flatMap(() => le16(1))])
      put(3, [...this.ftnStarts.flatMap((c) => le32(c)), ...le32(ccpFtn)])
    }
    // fields of the main text, the footnotes and the headers
    const plcFld = (i: number, flds: [number, number, number][], end: number) => {
      if (flds.length) put(i, [...flds.flatMap(([cp]) => le32(cp)), ...le32(end), ...flds.flatMap(([, ch, flt]) => [ch, flt])])
    }
    plcFld(16, this.main.fields, ccpText)
    plcFld(18, this.ftn.fields, ccpFtn)
    plcFld(17, this.hdd.fields, ccpHdd)
    // DOP
    const dop = new Buf().zeros(10).u16(720).zeros(488).done()
    put(31, dop)
    // Clx: one Prc (a piece's sprm), then the piece table
    const clx = new Buf()
    if (this.splitPieces) {
      const g = C.bold()
      clx.u8(1).u16(g.length).bytes(g)
    }
    clx.u8(2).u32((pieces.length + 1) * 4 + pieces.length * 8)
    for (const p of pieces) clx.u32(p.cp0)
    clx.u32(all.length)
    for (const p of pieces) clx.u16(0).u32(p.comp ? ((p.fc * 2) | 0x40000000) >>> 0 : p.fc).u16(p.prm ? 1 : 0)
    put(33, clx.done())
    // comments: references (ATRD), texts, authors, commented ranges (tagged annotation bookmarks)
    if (this.comments.length) {
      const authors = [...new Set(this.comments.map((c) => c.author))]
      const ref = new Buf()
      for (const c of this.comments) ref.u32(c.cp)
      ref.u32(ccpText)
      this.comments.forEach((c, k) => {
        const ini = c.initials.slice(0, 9)
        ref.u16(ini.length).utf16(ini.padEnd(9, '\0')).u16(authors.indexOf(c.author)).u16(0).u16(0).u32(c.range ? k + 1 : 0xffffffff)
      })
      put(4, ref.done())
      put(5, [...atnStarts.flatMap((c) => le32(c)), ...le32(ccpAtn)])
      const own = new Buf()
      for (const a of authors) own.u16(a.length).utf16(a)
      put(36, own.done())
      const ranged = this.comments.map((c, k) => ({ ...c, tag: k + 1 })).filter((c) => c.range)
      if (ranged.length) {
        const bkf = new Buf()
        for (const c of ranged) bkf.u32(c.cp - c.range)
        bkf.u32(ccpText + 1)
        ranged.forEach((_, j) => bkf.u16(j).u16(0))
        put(42, bkf.done())
        const bkl = new Buf()
        for (const c of ranged) bkl.u32(c.cp)
        bkl.u32(ccpText + 1)
        put(43, bkl.done())
        const names = new Buf().u16(0xffff).u16(ranged.length).u16(10)
        for (const c of ranged) names.u16(0).u16(0x100).u32(c.tag).u32(0xffffffff)
        put(37, names.done())
      }
    }
    // revision authors
    if (this.authors.length) {
      const st = new Buf().u16(0xffff).u16(this.authors.length).u16(0)
      for (const a of this.authors) st.u16(a.length).utf16(a)
      put(51, st.done())
    }
    // lists: PlfLst (the LVLs follow it), PlfLfo
    if (this.lists.length) {
      const lst = new Buf().u16(this.lists.length)
      for (const l of this.lists) {
        lst.u32(l.lsid).u32(l.lsid)
        for (let k = 0; k < 9; k++) lst.u16(0xfff)
        lst.u8(l.levels.length === 1 ? 1 : 0).u8(0)
      }
      tb.pad(2)
      fc[73] = [tb.length, lst.length]
      tb.bytes(lst.b)
      for (const l of this.lists)
        l.levels.forEach((v) => {
          const papx = [...P.left(v.left), ...P.first(-v.hang), ...sprm(0xc615, [0, 1, ...le16(v.left), 0])]
          const chpx = v.font !== undefined ? C.font(v.font) : []
          const nums = v.nums ?? []
          tb.u32(1).u8(v.nfc).u8(0)
          for (let k = 0; k < 9; k++) tb.u8(nums[k] ?? 0)
          tb.u8(0).u32(0).u32(0).u8(chpx.length).u8(papx.length).u8(0).u8(0)
          tb.bytes(papx).bytes(chpx).u16(v.text.length).utf16(v.text)
        })
      const lfo = new Buf().u32(this.lists.length)
      for (const l of this.lists) lfo.u32(l.lsid).u32(0).u32(0).u8(0).u8(0).u8(0).u8(0)
      for (let k = 0; k < this.lists.length; k++) lfo.u32(0xffffffff)
      put(74, lfo.done())
    }
    // floating pictures: anchors (PlcSpaMom) and the OfficeArt drawing with its BLIP store
    if (this.floats.length) {
      const spa = new Buf()
      for (const f of this.floats) spa.u32(f.cp)
      spa.u32(ccpText)
      for (const f of this.floats) spa.u32(f.spid).u32(f.rect[0]).u32(f.rect[1]).u32(f.rect[2]).u32(f.rect[3]).u16((2 << 1) | (2 << 3) | (2 << 5)).u32(0)
      put(40, spa.done())
      put(50, this.officeArt())
    }

    // ---- FIB
    const fib = new Buf()
    const flags = 0x0200 | 0x1000 | (this.encrypted ? 0x0100 : 0)
    fib.u16(0xa5ec).u16(this.nFib).u16(0).u16(0x409).u16(0).u16(flags).u16(0xbf).u32(0).u8(0).u8(0).u16(0).u16(0).u32(0).u32(0)
    fib.u16(14).zeros(28)
    fib.u16(22)
    const lw = [wd.length, 0, 0, ccpText, ccpFtn, ccpHdd, 0, ccpAtn, 0, 0, 0]
    for (let k = 0; k < 22; k++) fib.u32(lw[k] ?? 0)
    fib.u16(93)
    for (let k = 0; k < 93; k++) fib.u32(fc[k]?.[0] ?? 0).u32(fc[k]?.[1] ?? 0)
    fib.u16(0)
    for (let k = 0; k < fib.length; k++) wd.b[k] = fib.b[k]

    const streams: [string, Uint8Array][] = [
      ['WordDocument', wd.done()],
      ['1Table', tb.done()],
    ]
    if (this.data.length > 16) streams.push(['Data', this.data.done()])
    return cfb(streams)
  }

  /** moves the notes attached to runs into the footnote list (and records the reference CPs) */
  private collectNotes(notes: string[]) {
    for (const [cp, text] of this.pendingNotes) {
      this.footRefs.push(cp)
      notes.push(text)
    }
  }

  /** formatted disk pages for runs (CHPX or PAPX) */
  private fkps(wd: Buf, runs: { fc0: number; fc1: number; g: number[] }[], papx: boolean): { fc0: number; pn: number }[] {
    const pages: { fc0: number; pn: number }[] = []
    let i = 0
    while (i < runs.length) {
      const page = new Array(512).fill(0)
      const list: { fc0: number; fc1: number; off: number }[] = []
      let top = 511
      while (i < runs.length) {
        const r = runs[i]
        const data = papx ? papxBytes(r.g) : r.g.length ? [r.g.length, ...r.g] : []
        const n = list.length + 1
        const head = (n + 1) * 4 + n * (papx ? 13 : 1)
        let off = 0
        if (data.length) {
          off = (top - data.length) & ~1
          if (off < head) break
        } else if (head > top) break
        if (data.length) {
          for (let k = 0; k < data.length; k++) page[off + k] = data[k]
          top = off
        }
        list.push({ fc0: r.fc0, fc1: r.fc1, off })
        i++
      }
      if (!list.length) throw new Error('run too large for a page')
      list.forEach((r, k) => {
        page.splice(k * 4, 4, ...le32(r.fc0))
        if (papx) page[(list.length + 1) * 4 + k * 13] = r.off / 2
        else page[(list.length + 1) * 4 + k] = r.off / 2
      })
      page.splice(list.length * 4, 4, ...le32(list[list.length - 1].fc1))
      page[511] = list.length
      pages.push({ fc0: list[0].fc0, pn: wd.length / 512 })
      wd.bytes(page)
    }
    return pages
  }

  private stsh(): number[] {
    const byIstd = new Map(this.styles.map((s) => [s.istd, s]))
    const cstd = Math.max(...this.styles.map((s) => s.istd)) + 1
    const out = new Buf()
    out.u16(18)
    out.u16(cstd).u16(10).u16(1).u16(0x5b).u16(15).u16(0).u16(0).u16(0).u16(0)
    for (let istd = 0; istd < cstd; istd++) {
      const s = byIstd.get(istd)
      if (!s) {
        out.u16(0)
        continue
      }
      const std = new Buf()
      const cupx = s.stk === 1 ? 2 : 1
      std.u16(s.sti).u16(s.stk | ((s.base ?? 0xfff) << 4)).u16(cupx | ((s.next ?? istd) << 4)).u16(0).u16(0)
      std.u16(s.name.length).utf16(s.name).u16(0)
      std.pad(2)
      if (s.stk === 1) {
        const papx = [...le16(istd), ...(s.papx ?? [])]
        std.u16(papx.length).bytes(papx).pad(2)
      }
      const chpx = s.chpx ?? []
      std.u16(chpx.length).bytes(chpx).pad(2)
      // bchUpe: the size of the STD
      std.b[6] = std.length & 0xff
      std.b[7] = std.length >> 8
      out.u16(std.length).bytes(std.b)
    }
    return out.b
  }

  /** OfficeArtContent: the DggContainer (with the BLIP store) and the main text's drawing */
  private officeArt(): number[] {
    const rec = (ver: number, inst: number, type: number, body: number[]) => [...le16(ver | (inst << 4)), ...le16(type), ...le32(body.length), ...body]
    const fbse = this.blips.map((png) => {
      const blip = rec(0, 0x6e0, 0xf01e, [...new Array(16).fill(0), 0xff, ...png])
      return rec(2, 6, 0xf007, [6, 6, ...new Array(16).fill(0), ...le16(0xff), ...le32(blip.length), ...le32(1), ...le32(0), 0, 0, 0, 0, ...blip])
    })
    const dgg = rec(0xf, 0, 0xf000, [...rec(0, 0, 0xf006, [...le32(this.nextSpid), ...le32(2), ...le32(this.floats.length + 1), ...le32(1), ...le32(1), ...le32(this.floats.length + 1)]), ...rec(0xf, this.blips.length, 0xf001, fbse.flat())])
    const patriarch = rec(0xf, 0, 0xf004, [...rec(1, 0, 0xf009, new Array(16).fill(0)), ...rec(2, 0, 0xf00a, [...le32(1024), ...le32(5)])])
    const shapes = this.floats.map((f) => rec(0xf, 0, 0xf004, [...rec(2, 75, 0xf00a, [...le32(f.spid), ...le32(0xa00)]), ...rec(3, 1, 0xf00b, [...le16(0x4104), ...le32(f.pib)])]))
    const dg = rec(0xf, 1, 0xf002, [...rec(0, 1, 0xf008, [...le32(this.floats.length + 1), ...le32(this.nextSpid - 1)]), ...rec(0xf, 0, 0xf003, [...patriarch, ...shapes.flat()])])
    return [...dgg, 0, ...dg]
  }
}

/** PapxInFkp bytes of istd + grpprl: cb, then the data (cb 0 and a second count byte for even sizes) */
function papxBytes(g: number[]): number[] {
  if (g.length % 2) return [(g.length + 1) / 2, ...g]
  return [0, g.length / 2, ...g]
}

// ---------------------------------------------------------------- compound file

/** A compound file (512-byte sectors) holding the streams; every stream is padded past the 4096-byte mini cutoff. */
export function cfb(streams: [string, Uint8Array][]): Uint8Array {
  const sec = 512
  const padded = streams.map(([name, data]) => {
    const size = Math.max(4096, Math.ceil(data.length / sec) * sec)
    const out = new Uint8Array(size)
    out.set(data)
    return { name, data: out, size: Math.max(4096, data.length) }
  })
  // layout: stream sectors, then the directory, then the FAT
  const chains: number[][] = []
  let next = 0
  for (const s of padded) {
    const n = s.data.length / sec
    chains.push(Array.from({ length: n }, (_, k) => next + k))
    next += n
  }
  const entries = padded.length + 1
  const dirSectors = Math.ceil((entries * 128) / sec)
  const dirStart = next
  next += dirSectors
  let fatSectors = 1
  while (fatSectors * 128 < next + fatSectors) fatSectors++
  if (fatSectors > 109) throw new Error('compound file too large for the test writer')
  const fatStart = next
  const total = next + fatSectors
  const fat = new Uint32Array(fatSectors * 128).fill(0xffffffff)
  for (const ch of chains) ch.forEach((s, k) => (fat[s] = k + 1 < ch.length ? ch[k + 1] : 0xfffffffe))
  for (let k = 0; k < dirSectors; k++) fat[dirStart + k] = k + 1 < dirSectors ? dirStart + k + 1 : 0xfffffffe
  for (let k = 0; k < fatSectors; k++) fat[fatStart + k] = 0xfffffffd

  const out = new Uint8Array((total + 1) * sec)
  const v = new DataView(out.buffer)
  out.set([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1])
  v.setUint16(0x18, 0x3e, true)
  v.setUint16(0x1a, 3, true)
  v.setUint16(0x1c, 0xfffe, true)
  v.setUint16(0x1e, 9, true)
  v.setUint16(0x20, 6, true)
  v.setUint32(0x2c, fatSectors, true)
  v.setUint32(0x30, dirStart, true)
  v.setUint32(0x38, 4096, true)
  v.setUint32(0x3c, 0xfffffffe, true)
  v.setUint32(0x44, 0xfffffffe, true)
  for (let k = 0; k < 109; k++) v.setUint32(0x4c + k * 4, k < fatSectors ? fatStart + k : 0xffffffff, true)
  padded.forEach((s, k) => out.set(s.data, (chains[k][0] + 1) * sec))
  for (let k = 0; k < fat.length; k++) v.setUint32((fatStart + 1) * sec + k * 4, fat[k], true)
  // directory: the root, then the streams as a chain of right siblings (sorted the CFB way)
  const order = padded.map((_, k) => k).sort((a, b) => padded[a].name.length - padded[b].name.length || (padded[a].name.toUpperCase() < padded[b].name.toUpperCase() ? -1 : 1))
  const dirAt = (dirStart + 1) * sec
  const entry = (i: number, name: string, type: number, left: number, right: number, child: number, start: number, size: number) => {
    const o = dirAt + i * 128
    for (let k = 0; k < name.length; k++) v.setUint16(o + k * 2, name.charCodeAt(k), true)
    v.setUint16(o + 0x40, (name.length + 1) * 2, true)
    out[o + 0x42] = type
    out[o + 0x43] = 1
    v.setUint32(o + 0x44, left, true)
    v.setUint32(o + 0x48, right, true)
    v.setUint32(o + 0x4c, child, true)
    v.setUint32(o + 0x74, start, true)
    v.setUint32(o + 0x78, size, true)
  }
  entry(0, 'Root Entry', 5, 0xffffffff, 0xffffffff, order.length ? order[0] + 1 : 0xffffffff, 0xfffffffe, 0)
  order.forEach((k, j) => entry(k + 1, padded[k].name, 2, 0xffffffff, j + 1 < order.length ? order[j + 1] + 1 : 0xffffffff, 0xffffffff, chains[k][0], padded[k].size))
  for (let i = entries; i < dirSectors * 4; i++) {
    const o = dirAt + i * 128
    v.setUint32(o + 0x44, 0xffffffff, true)
    v.setUint32(o + 0x48, 0xffffffff, true)
    v.setUint32(o + 0x4c, 0xffffffff, true)
  }
  return out
}
