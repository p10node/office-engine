import { OpcPackage, relOfType, type Rels } from '../opc'
import { parseTheme } from '../theme'
import { kid, kids, num, parseXmlStream, type Attrs, type SaxHandler } from '../xml'
import { formatValue, formatText } from './numfmt'
import { parseSheetStyles, type SheetStyles } from './styles'
import type { BlockOut, CellDetail, CellOut, ColInfo, Merge, RowOut, SheetInfo, SheetProgress, WorkbookMeta } from './types'

const T_NUM = 0
const T_SST = 1
const T_STR = 2
const T_BOOL = 3
const T_ERR = 4
const T_DATE = 5
const T_EMPTY = 6

const MAX_COL = 16384

const decodeX = (s: string) => (s.indexOf('_x') >= 0 ? s.replace(/_x([0-9A-Fa-f]{4})_/g, (_, h: string) => String.fromCharCode(parseInt(h, 16))) : s)

export function colIndex(letters: string): number {
  let c = 0
  for (let i = 0; i < letters.length; i++) {
    const code = letters.charCodeAt(i)
    if (code < 65) break
    c = c * 26 + ((code & ~32) - 64)
  }
  return c - 1
}

export function colName(c: number): string {
  let s = ''
  c++
  while (c > 0) {
    const m = (c - 1) % 26
    s = String.fromCharCode(65 + m) + s
    c = Math.floor((c - 1) / 26)
  }
  return s
}

function parseRef(ref: string): [number, number] {
  let i = 0
  while (i < ref.length && ref.charCodeAt(i) >= 65) i++
  return [Number(ref.slice(i)) - 1, colIndex(ref.slice(0, i))]
}

export function parseRange(ref: string): Merge | undefined {
  const [a, b] = ref.replace(/\$/g, '').split(':')
  if (!a) return undefined
  const [r0, c0] = parseRef(a)
  const [r1, c1] = b ? parseRef(b) : [r0, c0]
  return { r0, c0, r1, c1 }
}

class SheetData {
  rows: (unknown[] | undefined)[] = []
  maxRow = -1
  maxCol = -1
  formulas = new Map<number, string>()
  merges: Merge[] = []
  colInfo: ColInfo[] = []
  defaultColW = 64
  defaultRowH = 20
  frozen = { rows: 0, cols: 0 }
  gridLines = true
  done = false
  error?: string
  loading?: Promise<void>
}

class SheetHandler implements SaxHandler {
  private row = -1
  private rowCells: unknown[] | null = null
  private col = -1
  private cType = ''
  private cStyle = 0
  private val = ''
  private formula = ''
  private hasVal = false
  private capture = 0 // 1 v, 2 f, 3 inline t
  private inIs = false
  newSizes: [number, number][] = []
  headerReady = false
  private rowHasHt = false
  private rowMaxPt = 0

  /** fontPt: font size per cell format, used to auto-grow rows without explicit height */
  constructor(private d: SheetData, private fontPt: number[] = [], private basePt = 11) {}

  open(name: string, a: Attrs): boolean | void {
    switch (name) {
      case 'c': {
        const r = a.r
        this.col = r ? parseRef(r)[1] : this.col + 1
        this.cType = a.t ?? ''
        this.cStyle = a.s ? Number(a.s) : 0
        this.val = ''
        this.formula = ''
        this.hasVal = false
        return
      }
      case 'v':
        this.capture = 1
        this.hasVal = true
        return
      case 'f':
        this.capture = 2
        return
      case 'is':
        this.inIs = true
        this.hasVal = true
        return
      case 't':
        if (this.inIs) this.capture = 3
        return
      case 'rPh':
        return true
      case 'row': {
        this.headerReady = true
        this.row = a.r ? Number(a.r) - 1 : this.row + 1
        this.col = -1
        this.rowCells = []
        const hidden = a.hidden === '1' || a.hidden === 'true'
        this.rowHasHt = hidden || !!a.ht
        this.rowMaxPt = 0
        if (hidden) this.newSizes.push([this.row, 0])
        else if (a.ht) {
          const px = Math.round((Number(a.ht) * 4) / 3)
          if (px !== this.d.defaultRowH) this.newSizes.push([this.row, px])
        }
        return
      }
      case 'sheetFormatPr': {
        const h = num(a.defaultRowHeight)
        if (h) this.d.defaultRowH = Math.round((h * 4) / 3)
        const w = num(a.defaultColWidth)
        const base = num(a.baseColWidth)
        if (w) this.d.defaultColW = Math.round(w * 7)
        else if (base !== undefined) this.d.defaultColW = Math.round(base * 7 + 8)
        return
      }
      case 'col': {
        const min = Number(a.min) - 1
        const max = Math.min(Number(a.max) - 1, MAX_COL - 1)
        const hidden = a.hidden === '1' || a.hidden === 'true'
        const w = hidden ? 0 : a.width ? Math.round(Number(a.width) * 7) : this.d.defaultColW
        this.d.colInfo.push({ min, max, w })
        return
      }
      case 'pane':
        if (a.state === 'frozen' || a.state === 'frozenSplit') {
          this.d.frozen = { rows: Math.round(num(a.ySplit, 0)!), cols: Math.round(num(a.xSplit, 0)!) }
        }
        return
      case 'sheetView':
        if (a.showGridLines === '0' || a.showGridLines === 'false') this.d.gridLines = false
        return
      case 'sheetData':
        this.headerReady = true
        return
      case 'mergeCell': {
        const m = a.ref ? parseRange(a.ref) : undefined
        if (m) this.d.merges.push(m)
        return
      }
      case 'extLst':
      case 'conditionalFormatting':
      case 'dataValidations':
      case 'headerFooter':
        return true
    }
  }

  close(name: string): void {
    switch (name) {
      case 'v':
      case 'f':
      case 't':
        this.capture = 0
        return
      case 'is':
        this.inIs = false
        return
      case 'c': {
        const cells = this.rowCells
        if (!cells) return
        const c = this.col
        if (c > this.d.maxCol) this.d.maxCol = c
        if (this.formula) this.d.formulas.set(this.row * MAX_COL + c, this.formula)
        if (!this.hasVal) {
          if (this.cStyle) cells.push(c, T_EMPTY, '', this.cStyle)
          return
        }
        let type = T_NUM
        let v: unknown
        switch (this.cType) {
          case 's':
            type = T_SST
            v = Number(this.val)
            break
          case 'str':
          case 'inlineStr':
            type = T_STR
            v = decodeX(this.val)
            break
          case 'b':
            type = T_BOOL
            v = this.val === '1' || this.val === 'true' ? 1 : 0
            break
          case 'e':
            type = T_ERR
            v = this.val
            break
          case 'd':
            type = T_DATE
            v = this.val
            break
          default: {
            const n = Number(this.val)
            if (this.val === '' || Number.isNaN(n)) {
              type = this.val === '' ? T_EMPTY : T_STR
              v = this.val
            } else v = n
          }
        }
        cells.push(c, type, v, this.cStyle)
        if (!this.rowHasHt) {
          const pt = this.fontPt[this.cStyle] ?? this.basePt
          if (pt > this.rowMaxPt) this.rowMaxPt = pt
        }
        return
      }
      case 'row':
        if (!this.rowHasHt && this.rowMaxPt > this.basePt * 1.15) {
          const px = Math.round(this.rowMaxPt * (4 / 3) * 1.3 + 3)
          if (px > this.d.defaultRowH) this.newSizes.push([this.row, px])
        }
        if (this.rowCells && this.rowCells.length) this.d.rows[this.row] = this.rowCells
        if (this.row > this.d.maxRow) this.d.maxRow = this.row
        this.rowCells = null
        return
    }
  }

  text(t: string): void {
    if (this.capture === 1) this.val += t
    else if (this.capture === 2) this.formula += t
    else if (this.capture === 3) this.val += t
  }
}

class SstHandler implements SaxHandler {
  strings: string[] = []
  private cur = ''
  private capture = false
  open(name: string): boolean | void {
    if (name === 'si') this.cur = ''
    else if (name === 't') this.capture = true
    else if (name === 'rPh' || name === 'phoneticPr') return true
  }
  close(name: string): void {
    if (name === 't') this.capture = false
    else if (name === 'si') this.strings.push(decodeX(this.cur))
  }
  text(t: string): void {
    if (this.capture) this.cur += t
  }
}

export class XlsxWorkbook {
  private sheetsData = new Map<number, SheetData>()
  private sst: string[] = []
  private sstLoading?: Promise<void>

  private constructor(
    private pkg: OpcPackage,
    private rels: Rels,
    readonly sheets: SheetInfo[],
    private styles: SheetStyles,
    readonly date1904: boolean,
  ) {}

  static async open(pkg: OpcPackage, mainPath: string): Promise<XlsxWorkbook> {
    const root = await pkg.xml(mainPath)
    const rels = await pkg.rels(mainPath)
    const sheets: SheetInfo[] = []
    for (const s of kids(kid(root, 'sheets'), 'sheet')) {
      const rel = rels.get(s.a['r:id'])
      if (!rel || !rel.type.endsWith('/worksheet')) continue
      sheets.push({ name: s.a.name ?? `Sheet${sheets.length + 1}`, index: sheets.length, path: rel.target, hidden: !!s.a.state && s.a.state !== 'visible' })
    }
    const themeRel = relOfType(rels, 'theme')
    const theme = parseTheme(themeRel ? await pkg.xml(themeRel.target) : undefined)
    const stylesRel = relOfType(rels, 'styles')
    const styles = parseSheetStyles(stylesRel ? await pkg.xml(stylesRel.target) : undefined, theme)
    const pr = kid(root, 'workbookPr')
    const date1904 = pr?.a.date1904 === '1' || pr?.a.date1904 === 'true'
    return new XlsxWorkbook(pkg, rels, sheets, styles, date1904)
  }

  meta(): WorkbookMeta {
    return { sheets: this.sheets, styles: this.styles.xfs.map((x) => x.css), baseFont: this.styles.baseFont, date1904: this.date1904 }
  }

  private loadSst(onProgress?: (loaded: number, total: number) => void): Promise<void> {
    if (!this.sstLoading) {
      this.sstLoading = (async () => {
        const rel = relOfType(this.rels, 'sharedStrings')
        if (!rel || !this.pkg.has(rel.target)) return
        const h = new SstHandler()
        this.sst = h.strings
        await parseXmlStream(await this.pkg.zip.stream(rel.target, onProgress), h)
      })()
    }
    return this.sstLoading
  }

  /**
   * Streams a sheet into memory, reporting progress so the UI can grow while parsing.
   * Calling again for a loaded/in-flight sheet is a no-op.
   */
  loadSheet(index: number, emit: (p: SheetProgress) => void, onBytes?: (phase: string, loaded: number, total: number) => void): Promise<void> {
    const existing = this.sheetsData.get(index)
    if (existing) {
      emit(this.snapshot(index, existing, [], true))
      return existing.loading ?? Promise.resolve()
    }
    const d = new SheetData()
    this.sheetsData.set(index, d)
    d.loading = (async () => {
      const info = this.sheets[index]
      if (!info) throw new Error(`No sheet #${index}`)
      try {
        await this.loadSst((l, t) => onBytes?.('strings', l, t))
        const h = new SheetHandler(d, this.styles.xfs.map((x) => x.fontPt), this.styles.xfs[0]?.fontPt ?? 11)
        let sentHeader = false
        let last = 0
        await parseXmlStream(await this.pkg.zip.stream(info.path, (l, t) => onBytes?.('sheet', l, t)), h, {
          onChunk: async () => {
            const now = Date.now()
            if (!sentHeader && h.headerReady) {
              sentHeader = true
              emit(this.snapshot(index, d, h.newSizes.splice(0), true))
              last = now
              await new Promise((r) => setTimeout(r, 0))
            } else if (now - last > 150) {
              last = now
              emit(this.snapshot(index, d, h.newSizes.splice(0), false))
              await new Promise((r) => setTimeout(r, 0))
            }
          },
        })
        d.done = true
        emit({ ...this.snapshot(index, d, h.newSizes.splice(0), !sentHeader), merges: d.merges })
      } catch (e) {
        d.done = true
        d.error = (e as Error).message
        emit({ ...this.snapshot(index, d, [], true), error: d.error })
      }
    })()
    return d.loading
  }

  private snapshot(sheet: number, d: SheetData, rowSizes: [number, number][], header: boolean): SheetProgress {
    const p: SheetProgress = { sheet, rows: d.maxRow + 1, cols: d.maxCol + 1, rowSizes, done: d.done }
    if (header) {
      p.colInfo = d.colInfo
      p.defaultColW = d.defaultColW
      p.defaultRowH = d.defaultRowH
      p.frozen = d.frozen
      p.gridLines = d.gridLines
    }
    if (d.done) p.merges = d.merges
    if (d.error) p.error = d.error
    return p
  }

  /** Formatted cells for rows [r0, r1). */
  block(sheet: number, r0: number, r1: number): BlockOut {
    const d = this.sheetsData.get(sheet)
    const rows: RowOut[] = []
    if (!d) return { sheet, r0, r1, rows, complete: false }
    const xfs = this.styles.xfs
    const end = Math.min(r1, d.rows.length)
    for (let r = r0; r < end; r++) {
      const row = d.rows[r]
      if (!row) continue
      const cells: CellOut[] = []
      for (let k = 0; k < row.length; k += 4) {
        const c = row[k] as number
        const t = row[k + 1] as number
        const v = row[k + 2]
        const s = row[k + 3] as number
        const xf = xfs[s] ?? xfs[0]
        switch (t) {
          case T_NUM: {
            const f = formatValue(v as number, xf.numFmt, this.date1904)
            cells.push(f.color ? [c, f.text, s, 1, f.color] : [c, f.text, s, 1])
            break
          }
          case T_SST:
            cells.push([c, formatText(this.sst[v as number] ?? '', xf.numFmt), s, 0])
            break
          case T_STR:
            cells.push([c, formatText(v as string, xf.numFmt), s, 0])
            break
          case T_BOOL:
            cells.push([c, v ? 'TRUE' : 'FALSE', s, 2])
            break
          case T_ERR:
            cells.push([c, v as string, s, 2])
            break
          case T_DATE:
            cells.push([c, (v as string).replace('T00:00:00', '').replace('T', ' ').replace(/Z$/, ''), s, 1])
            break
          default:
            cells.push([c, '', s, 0])
        }
      }
      rows.push({ r, cells })
    }
    return { sheet, r0, r1, rows, complete: d.done || r1 <= d.maxRow }
  }

  cell(sheet: number, r: number, c: number): CellDetail | undefined {
    const d = this.sheetsData.get(sheet)
    const row = d?.rows[r]
    const formula = d?.formulas.get(r * MAX_COL + c)
    if (!row) return formula ? { value: '', formula: '=' + formula, type: 'formula' } : undefined
    for (let k = 0; k < row.length; k += 4) {
      if (row[k] !== c) continue
      const t = row[k + 1] as number
      const v = row[k + 2]
      const value = t === T_SST ? this.sst[v as number] ?? '' : t === T_BOOL ? (v ? 'TRUE' : 'FALSE') : String(v)
      const type = ['number', 'string', 'string', 'boolean', 'error', 'date', 'empty'][t]
      return { value, formula: formula ? '=' + formula : undefined, type }
    }
    return formula ? { value: '', formula: '=' + formula, type: 'formula' } : undefined
  }

  unload(sheet: number) {
    const d = this.sheetsData.get(sheet)
    if (d?.done) this.sheetsData.delete(sheet)
  }
}
