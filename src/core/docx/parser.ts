import { OpcPackage, relOfType, type Rels } from '../opc'
import { parseTheme, type Theme } from '../theme'
import { emuToPt } from '../units'
import { desc, descAll, kid, kids, num, parseXmlStream, path, TreeBuilder, type XEl } from '../xml'
import { Numbering } from './numbering'
import {
  borderCss, merge, readMargins, readPPr, readRPr, readTableBorders, shading, StyleSheet,
} from './styles'
import type {
  Block, CellBorders, DocxMeta, Inline, PageSetup, Paragraph, ParaStyle, RunStyle, Table, TableCell, TableRow,
} from './types'
import { DEFAULT_PAGE } from './types'

interface FieldState {
  instr: string
  phase: 'instr' | 'result'
  kind?: 'PAGE' | 'NUMPAGES'
  href?: string
}

interface Sink {
  push(x: Inline): void
  pageBreak(): void
}

function fieldKind(instr: string): 'PAGE' | 'NUMPAGES' | undefined {
  const m = /^\s*(PAGE|NUMPAGES|SECTIONPAGES)\b/i.exec(instr)
  if (!m) return undefined
  return m[1].toUpperCase() === 'PAGE' ? 'PAGE' : 'NUMPAGES'
}

function sameStyle(a: RunStyle, b: RunStyle): boolean {
  if (a === b) return true
  const ka = Object.keys(a)
  if (ka.length !== Object.keys(b).length) return false
  for (const k of ka) if ((a as Record<string, unknown>)[k] !== (b as Record<string, unknown>)[k]) return false
  return true
}

function cssLen(style: string | undefined, prop: string): number | undefined {
  if (!style) return undefined
  const m = new RegExp(`(?:^|;)\\s*${prop}\\s*:\\s*([\\d.]+)\\s*(pt|in|px|cm|mm)?`, 'i').exec(style)
  if (!m) return undefined
  const v = Number(m[1])
  switch ((m[2] ?? 'px').toLowerCase()) {
    case 'pt': return v
    case 'in': return v * 72
    case 'cm': return (v * 72) / 2.54
    case 'mm': return (v * 72) / 25.4
    default: return v * 0.75
  }
}

export function pageSetup(sectPr: XEl | undefined): PageSetup {
  if (!sectPr) return { ...DEFAULT_PAGE }
  const sz = kid(sectPr, 'pgSz')
  const mar = kid(sectPr, 'pgMar')
  const t = (v: string | undefined, d: number) => (v === undefined ? d : Math.abs(Number(v)) / 20)
  let w = t(sz?.a.w, DEFAULT_PAGE.w)
  let h = t(sz?.a.h, DEFAULT_PAGE.h)
  if (sz?.a.orient === 'landscape' && h > w) [w, h] = [h, w]
  return {
    w, h,
    mt: t(mar?.a.top, 72), mr: t(mar?.a.right ?? mar?.a.end, 72), mb: t(mar?.a.bottom, 72), ml: t(mar?.a.left ?? mar?.a.start, 72),
    header: t(mar?.a.header, 36), footer: t(mar?.a.footer, 36),
  }
}

class Converter {
  private fields: FieldState[] = []
  private prev?: { p: Paragraph; styleId: string }
  private tableLayer?: { ps: ParaStyle; rs: RunStyle }

  constructor(
    private styles: StyleSheet,
    private numbering: Numbering,
    private theme: Theme,
    private rels: Rels,
  ) {}

  blocks(el: XEl, out: Block[]): void {
    switch (el.n) {
      case 'p':
        this.paragraph(el, out)
        break
      case 'tbl':
        this.prev = undefined
        out.push(this.table(el))
        this.prev = undefined
        break
      case 'sdt':
        for (const c of kid(el, 'sdtContent')?.c ?? []) this.blocks(c, out)
        break
      case 'customXml':
      case 'ins':
      case 'smartTag':
      case 'sdtContent':
        for (const c of el.c) this.blocks(c, out)
        break
    }
  }

  paragraph(el: XEl, out: Block[]): void {
    const pPr = kid(el, 'pPr')
    const d = readPPr(pPr)
    const st = this.styles.para(d.styleId)
    const layer = this.tableLayer
    const numPr = d.numPr ?? st.numPr
    const listed = numPr && numPr.numId !== '0' ? this.numbering.next(numPr.numId, numPr.ilvl) : undefined
    const ps = merge<ParaStyle>(this.styles.docPs, layer?.ps, st.ps, listed?.level.ps, d.ps)
    const base = merge<RunStyle>(this.styles.docRs, layer?.rs, st.rs)
    const markRs = merge<RunStyle>(base, readRPr(kid(pPr, 'rPr'), this.theme))
    // contextual spacing: no space between paragraphs of the same style
    const styleId = d.styleId ?? this.styles.defaultPara ?? ''
    if (this.prev && this.prev.styleId === styleId) {
      if (ps.contextual) ps.before = 0
      if (this.prev.p.ps.contextual) this.prev.p.ps = { ...this.prev.p.ps, after: 0 }
    }
    let cur: Paragraph = { t: 'p', ps, runs: [], markSz: markRs.sz, markFont: markRs.font }
    if (ps.pageBreakBefore) cur.pb = true
    if (listed && listed.text !== undefined) {
      const ms = merge<RunStyle>(markRs, listed.level.rs)
      delete ms.u
      delete ms.hl
      cur.marker = { text: listed.level.suff === 'space' ? listed.text + ' ' : listed.text, s: ms, tab: listed.level.suff === 'tab' }
    }
    const parts: Paragraph[] = [cur]
    const sink: Sink = {
      push: (x) => {
        const runs = cur.runs
        const last = runs[runs.length - 1]
        if (x.t === 'text' && last?.t === 'text' && last.href === x.href && sameStyle(last.s, x.s)) last.text += x.text
        else runs.push(x)
      },
      pageBreak: () => {
        cur = { t: 'p', ps, runs: [], markSz: markRs.sz, markFont: markRs.font, pb: true }
        parts.push(cur)
      },
    }
    this.inlines(el.c, base, undefined, sink)
    // A trailing page break leaves an empty paragraph that only carries the break.
    for (const p of parts) out.push(p)
    this.prev = { p: parts[parts.length - 1], styleId }
    if (d.sectPr) out.push({ t: 'sect', page: pageSetup(d.sectPr), kind: kid(d.sectPr, 'type')?.a.val ?? 'nextPage' })
  }

  private inlines(children: XEl[], base: RunStyle, href: string | undefined, sink: Sink): void {
    for (const c of children) {
      switch (c.n) {
        case 'r':
          this.run(c, base, href, sink)
          break
        case 'hyperlink': {
          const id = c.a['r:id']
          const target = id ? this.rels.get(id)?.target : c.a.anchor ? '#' + c.a.anchor : undefined
          this.inlines(c.c, base, target ?? href, sink)
          break
        }
        case 'fldSimple': {
          const kind = fieldKind(c.a.instr ?? '')
          if (kind) {
            const r = kid(c, 'r')
            sink.push({ t: 'field', kind, s: merge(base, readRPr(kid(r, 'rPr'), this.theme)) })
          } else this.inlines(c.c, base, href, sink)
          break
        }
        case 'ins':
        case 'smartTag':
        case 'customXml':
        case 'bdo':
        case 'dir':
        case 'moveTo':
          this.inlines(c.c, base, href, sink)
          break
        case 'sdt':
          this.inlines(kid(c, 'sdtContent')?.c ?? [], base, href, sink)
          break
        case 'oMath':
        case 'oMathPara': {
          const text = descAll(c, 't').map((t) => t.t).join('')
          if (text) sink.push({ t: 'text', text, s: { ...base, i: true, font: '"Cambria Math", "STIX Two Math", serif' } })
          break
        }
      }
    }
  }

  private run(r: XEl, base: RunStyle, href: string | undefined, sink: Sink): void {
    const rPr = kid(r, 'rPr')
    const s = merge<RunStyle>(base, this.styles.char(kid(rPr, 'rStyle')?.a.val), readRPr(rPr, this.theme))
    if (s.vanish) return
    for (const c of r.c) {
      const f = this.fields[this.fields.length - 1]
      switch (c.n) {
        case 'fldChar': {
          const type = c.a.fldCharType
          if (type === 'begin') this.fields.push({ instr: '', phase: 'instr' })
          else if (type === 'separate' && f) {
            f.phase = 'result'
            f.kind = fieldKind(f.instr)
            const link = /HYPERLINK\s+(?:\\l\s+)?"([^"]+)"/.exec(f.instr)
            if (link) f.href = /\\l/.test(f.instr) ? '#' + link[1] : link[1]
            if (f.kind) sink.push({ t: 'field', kind: f.kind, s })
          } else if (type === 'end') {
            // Field without a result part (e.g. PAGE in some generators)
            if (f && f.phase === 'instr') {
              const kind = fieldKind(f.instr)
              if (kind) sink.push({ t: 'field', kind, s })
            }
            this.fields.pop()
          }
          break
        }
        case 'instrText':
          if (f) f.instr += c.t
          break
        case 't':
        case 'delText':
          if (c.n === 'delText') break
          if (f && (f.phase === 'instr' || f.kind)) break
          if (c.t) sink.push({ t: 'text', text: c.t, s, href: f?.href ?? href })
          break
        case 'tab':
        case 'ptab':
          if (!f || (f.phase === 'result' && !f.kind)) sink.push({ t: 'tab', s })
          break
        case 'br':
          if (c.a.type === 'page') sink.pageBreak()
          else sink.push({ t: 'br', s })
          break
        case 'cr':
          sink.push({ t: 'br', s })
          break
        case 'noBreakHyphen':
          sink.push({ t: 'text', text: '\u2011', s, href })
          break
        case 'softHyphen':
          sink.push({ t: 'text', text: '\u00AD', s, href })
          break
        case 'sym': {
          const code = parseInt(c.a.char ?? '', 16)
          if (code) sink.push({ t: 'text', text: String.fromCharCode(code >= 0xf000 ? code - 0xf000 : code), s: { ...s, font: c.a.font ? `"${c.a.font}", ${s.font ?? 'serif'}` : s.font }, href })
          break
        }
        case 'drawing':
          this.drawing(c, href, sink)
          break
        case 'pict':
        case 'object':
          this.vml(c, href, sink)
          break
        case 'footnoteReference':
        case 'endnoteReference':
          sink.push({ t: 'text', text: c.a.id ?? '*', s: { ...s, va: 'sup' } })
          break
      }
    }
  }

  private target(id: string | undefined): string | undefined {
    if (!id) return undefined
    const rel = this.rels.get(id)
    return rel && !rel.external ? rel.target : undefined
  }

  private drawing(d: XEl, href: string | undefined, sink: Sink): void {
    const box = kid(d, 'inline') ?? kid(d, 'anchor')
    if (!box) return
    const ext = kid(box, 'extent')
    const w = emuToPt(num(ext?.a.cx, 0))
    const h = emuToPt(num(ext?.a.cy, 0))
    const docPr = kid(box, 'docPr')
    const gd = path(box, 'graphic', 'graphicData')
    const txbx = descAll(gd, 'txbxContent')
    if (txbx.length) {
      const blocks: Block[] = []
      for (const t of txbx) for (const c of t.c) this.blocks(c, blocks)
      const spPr = desc(gd, 'spPr')
      const fill = path(spPr, 'solidFill', 'srgbClr')?.a.val
      const ln = path(spPr, 'ln', 'solidFill', 'srgbClr')?.a.val
      sink.push({ t: 'box', w, h, blocks, fill: fill ? '#' + fill : undefined, border: ln ? `1px solid #${ln}` : undefined })
      return
    }
    const blip = desc(gd, 'blip')
    const target = this.target(blip?.a['r:embed'])
    if (target) {
      const click = this.rels.get(kid(docPr, 'hlinkClick')?.a['r:id'] ?? '')?.target
      sink.push({ t: 'img', target, w, h, alt: docPr?.a.descr || docPr?.a.title || undefined, href: click ?? href })
      return
    }
    if (desc(gd, 'chart')) sink.push({ t: 'shape', w, h, label: 'Chart' })
    else if (desc(gd, 'relIds')) sink.push({ t: 'shape', w, h, label: 'Diagram' })
  }

  private vml(el: XEl, href: string | undefined, sink: Sink): void {
    const tb = desc(el, 'txbxContent')
    const shape = desc(el, 'shape') ?? desc(el, 'rect')
    const style = shape?.a.style
    const w = cssLen(style, 'width') ?? 0
    const h = cssLen(style, 'height') ?? 0
    if (tb) {
      const blocks: Block[] = []
      for (const c of tb.c) this.blocks(c, blocks)
      sink.push({ t: 'box', w, h, blocks })
      return
    }
    const target = this.target(desc(el, 'imagedata')?.a['r:id'])
    if (target && w && h) sink.push({ t: 'img', target, w, h, href })
  }

  table(el: XEl): Table {
    const tblPr = kid(el, 'tblPr')
    const ts = this.styles.table(kid(tblPr, 'tblStyle')?.a.val)
    const look = readLook(kid(tblPr, 'tblLook'))
    const borders = { ...ts.borders, ...readTableBorders(kid(tblPr, 'tblBorders')) }
    const mar = { top: 0, bottom: 0, left: 5.4, right: 5.4, ...ts.cellMar, ...readMargins(kid(tblPr, 'tblCellMar')) }
    const cols = kids(kid(el, 'tblGrid'), 'gridCol').map((g) => (num(g.a.w) ?? 0) / 20)
    const jc = kid(tblPr, 'jc')?.a.val
    const ind = kid(tblPr, 'tblInd')
    const rowEls = flatten(el, 'tr')

    interface Tmp extends TableCell { gs: number; vm?: 'restart' | 'continue'; tcB: ReturnType<typeof readTableBorders>; gone?: boolean; ri: number }
    const grid: { row: TableRow; cells: Tmp[] }[] = []
    let nCols = cols.length
    rowEls.forEach((tr, ri) => {
      const trPr = kid(tr, 'trPr')
      const th = kid(trPr, 'trHeight')
      const row: TableRow = { cells: [] }
      if (th?.a.val) {
        row.h = Number(th.a.val) / 20
        row.exact = th.a.hRule === 'exact'
      }
      if (kid(trPr, 'tblHeader')) row.header = true
      let col = num(kid(trPr, 'gridBefore')?.a.val, 0)!
      const cells: Tmp[] = []
      for (const tc of flatten(tr, 'tc')) {
        const tcPr = kid(tc, 'tcPr')
        const span = Math.max(1, num(kid(tcPr, 'gridSpan')?.a.val, 1)!)
        const vmEl = kid(tcPr, 'vMerge')
        const conds: string[] = []
        if (!look.noVBand) conds.push((col - (look.firstCol ? 1 : 0)) % 2 === 0 ? 'band1Vert' : 'band2Vert')
        if (!look.noHBand) {
          const bi = ri - (look.firstRow ? 1 : 0)
          if (bi >= 0) conds.push(bi % 2 === 0 ? 'band1Horz' : 'band2Horz')
        }
        if (look.firstCol && col === 0) conds.push('firstCol')
        if (look.lastCol && nCols && col + span >= nCols) conds.push('lastCol')
        if (look.firstRow && ri === 0) conds.push('firstRow')
        if (look.lastRow && ri === rowEls.length - 1) conds.push('lastRow')
        let shd = ts.shd
        let crs: RunStyle = {}
        for (const k of conds) {
          const cf = ts.cond[k]
          if (!cf) continue
          if (cf.shd) shd = cf.shd
          crs = merge(crs, cf.rs)
        }
        shd = shading(kid(tcPr, 'shd')) ?? shd
        const prev = this.tableLayer
        this.tableLayer = { ps: ts.ps, rs: merge(ts.rs, crs) }
        const blocks: Block[] = []
        for (const c of tc.c) if (c.n !== 'tcPr') this.blocks(c, blocks)
        this.tableLayer = prev
        const cm = readMargins(kid(tcPr, 'tcMar'))
        const va = kid(tcPr, 'vAlign')?.a.val
        const tcW = kid(tcPr, 'tcW')
        cells.push({
          blocks,
          colSpan: span,
          rowSpan: 1,
          w: tcW && (tcW.a.type ?? 'dxa') === 'dxa' ? Number(tcW.a.w) / 20 : undefined,
          shd,
          vAlign: va === 'center' ? 'middle' : va === 'bottom' ? 'bottom' : undefined,
          borders: {},
          pad: [cm.top ?? mar.top, cm.right ?? mar.right, cm.bottom ?? mar.bottom, cm.left ?? mar.left],
          gs: col,
          vm: vmEl ? (vmEl.a.val === 'restart' ? 'restart' : 'continue') : undefined,
          tcB: readTableBorders(kid(tcPr, 'tcBorders')),
          ri,
        })
        col += span
      }
      nCols = Math.max(nCols, col)
      grid.push({ row, cells })
    })

    // vertical merges
    for (let ri = 0; ri < grid.length; ri++) {
      for (const cell of grid[ri].cells) {
        if (cell.gone || !cell.vm) continue
        for (let rj = ri + 1; rj < grid.length; rj++) {
          const below = grid[rj].cells.find((c) => c.gs === cell.gs)
          if (!below || below.vm !== 'continue') break
          below.gone = true
          cell.rowSpan++
        }
      }
    }
    const nRows = grid.length
    const rows: TableRow[] = grid.map(({ row, cells }) => {
      for (const c of cells) {
        if (c.gone) continue
        const lastRow = c.ri + c.rowSpan - 1 === nRows - 1
        const b: CellBorders = {
          t: borderCss(c.tcB.top !== undefined ? c.tcB.top : c.ri === 0 ? borders.top : borders.insideH),
          b: borderCss(c.tcB.bottom !== undefined ? c.tcB.bottom : lastRow ? borders.bottom : borders.insideH),
          l: borderCss(c.tcB.left !== undefined ? c.tcB.left : c.gs === 0 ? borders.left : borders.insideV),
          r: borderCss(c.tcB.right !== undefined ? c.tcB.right : c.gs + c.colSpan >= nCols ? borders.right : borders.insideV),
        }
        const { gs: _gs, vm: _vm, tcB: _tcB, gone: _g, ri: _ri, ...cell } = c
        row.cells.push({ ...cell, borders: b })
      }
      return row
    })
    return {
      t: 'tbl',
      cols,
      rows,
      align: jc === 'center' ? 'center' : jc === 'right' || jc === 'end' ? 'right' : undefined,
      indL: ind && (ind.a.type ?? 'dxa') === 'dxa' ? Number(ind.a.w) / 20 : undefined,
    }
  }
}

function flatten(el: XEl, name: string): XEl[] {
  const out: XEl[] = []
  for (const c of el.c) {
    if (c.n === name) out.push(c)
    else if (c.n === 'sdt') out.push(...flatten(kid(c, 'sdtContent') ?? c, name))
    else if (c.n === 'customXml' || c.n === 'sdtContent') out.push(...flatten(c, name))
  }
  return out
}

function readLook(el: XEl | undefined) {
  const look = { firstRow: true, lastRow: false, firstCol: true, lastCol: false, noHBand: false, noVBand: true }
  if (!el) return look
  const hex = el.a.val
  if (hex && /^[0-9a-fA-F]+$/.test(hex) && el.a.firstRow === undefined) {
    const v = parseInt(hex, 16)
    return {
      firstRow: !!(v & 0x20), lastRow: !!(v & 0x40), firstCol: !!(v & 0x80), lastCol: !!(v & 0x100),
      noHBand: !!(v & 0x200), noVBand: !!(v & 0x400),
    }
  }
  const b = (k: string, d: boolean) => (el.a[k] === undefined ? d : el.a[k] === '1' || el.a[k] === 'true')
  return {
    firstRow: b('firstRow', true), lastRow: b('lastRow', false), firstCol: b('firstColumn', true),
    lastCol: b('lastColumn', false), noHBand: b('noHBand', false), noVBand: b('noVBand', true),
  }
}

// ---------------------------------------------------------------------------

export interface DocxStreamCallbacks {
  onBlocks(blocks: Block[]): void
  onMeta(meta: DocxMeta): void
  onProgress?(loaded: number, total: number): void
  signal?: AbortSignal
  /** Hold blocks until this much time passed so small docs render in one go (ms). */
  initialHoldMs?: number
}

export class DocxDocument {
  private constructor(
    private pkg: OpcPackage,
    private mainPath: string,
    private rels: Rels,
    private styles: StyleSheet,
    private numbering: Numbering,
    private theme: Theme,
  ) {}

  static async open(pkg: OpcPackage, mainPath: string): Promise<DocxDocument> {
    const rels = await pkg.rels(mainPath)
    const themeRel = relOfType(rels, 'theme')
    const theme = parseTheme(themeRel ? await pkg.xml(themeRel.target) : undefined)
    const stylesRel = relOfType(rels, 'styles')
    const numRel = relOfType(rels, 'numbering')
    const styles = new StyleSheet(stylesRel ? await pkg.xml(stylesRel.target) : undefined, theme)
    const numbering = new Numbering(numRel ? await pkg.xml(numRel.target) : undefined, theme)
    return new DocxDocument(pkg, mainPath, rels, styles, numbering, theme)
  }

  private async headerFooter(id: string | undefined): Promise<Block[] | undefined> {
    const rel = id ? this.rels.get(id) : undefined
    if (!rel) return undefined
    const root = await this.pkg.xml(rel.target)
    if (!root) return undefined
    const conv = new Converter(this.styles, this.numbering, this.theme, await this.pkg.rels(rel.target))
    const out: Block[] = []
    for (const c of root.c) conv.blocks(c, out)
    return out
  }

  async stream(cb: DocxStreamCallbacks): Promise<void> {
    const conv = new Converter(this.styles, this.numbering, this.theme, this.rels)
    const baseMeta = { font: this.styles.docRs.font ?? '"Times New Roman", serif', sz: this.styles.docRs.sz ?? 10 }
    let pending: Block[] = []
    let finalSect: XEl | undefined
    let firstSect: PageSetup | undefined
    const started = Date.now()
    const hold = cb.initialHoldMs ?? 350
    let lastFlush = 0
    let held = true
    const tb = new TreeBuilder({
      emitDepth: 2,
      onElement: (el) => {
        if (el.n === 'sectPr') {
          finalSect = el
          return
        }
        const before = pending.length
        conv.blocks(el, pending)
        if (!firstSect) for (let i = before; i < pending.length; i++) {
          const b = pending[i]
          if (b.t === 'sect') firstSect = b.page
        }
      },
    })
    const flush = () => {
      if (!pending.length) return
      cb.onBlocks(pending)
      pending = []
      lastFlush = Date.now()
    }
    const stream = await this.pkg.zip.stream(this.mainPath, cb.onProgress)
    await parseXmlStream(stream, tb, {
      signal: cb.signal,
      onChunk: async () => {
        const now = Date.now()
        if (held) {
          if (now - started < hold) return
          held = false
          cb.onMeta({ page: firstSect ?? { ...DEFAULT_PAGE }, ...baseMeta })
          flush()
          return
        }
        if (pending.length >= 400 || (pending.length && now - lastFlush > 120)) {
          flush()
          // yield so RPC requests (images, etc.) get served during long parses
          await new Promise((r) => setTimeout(r, 0))
        }
      },
    })
    const sect = finalSect
    const meta: DocxMeta = { page: pageSetup(sect), ...baseMeta, final: true }
    if (sect) {
      const ref = (tag: string, type: string) => kids(sect, tag).find((r) => (r.a.type ?? 'default') === type)?.a['r:id']
      meta.header = await this.headerFooter(ref('headerReference', 'default'))
      meta.footer = await this.headerFooter(ref('footerReference', 'default'))
      meta.titlePg = !!kid(sect, 'titlePg') && kid(sect, 'titlePg')?.a.val !== '0'
      if (meta.titlePg) {
        meta.firstHeader = await this.headerFooter(ref('headerReference', 'first'))
        meta.firstFooter = await this.headerFooter(ref('footerReference', 'first'))
      }
    }
    cb.onMeta(meta)
    flush()
  }
}

