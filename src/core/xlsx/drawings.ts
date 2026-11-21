/**
 * The pictures and charts on a sheet, read from its drawing part (`xl/drawings/drawingN.xml`, linked from the sheet's
 * rels), each with the place it is anchored to: two cell corners (`twoCellAnchor`), one corner and a size
 * (`oneCellAnchor`), or a position on the sheet and a size (`absoluteAnchor`). A picture carries its media part, which
 * the view shows where the browser can draw it (not EMF / WMF); a chart carries its model with the values read from the
 * cells (chart.ts), or, of a kind not drawn, its kind and title for a placeholder the size of the chart. Shapes, text
 * boxes and the like are left out, as are drawings in a group the anchor's own frame does not hold.
 */
import { dict } from '../dict'
import { isWebImage } from '../opc'
import { desc, descAll, kid, textOf, type XEl } from '../xml'
import { MAX_COL, MAX_ROWS, MAX_SHEET_WIDTH_PX, MAX_SHEET_HEIGHT_PX, ooxmlInt } from './limits'
import { parseFormula, references } from './formula'
import { isoToSerial } from './formula/dates'
import { readChart, type ChartModel, type RangeReader, type RangeValues } from './chart'
import { T_BOOL, T_DATE, T_EMPTY, T_ERR, T_NUM, T_SST, type XlsxWorkbook } from './workbook'

/** A cell corner: 0-based column and row, and the offset into that cell in CSS pixels */
export interface DrawingCorner {
  c: number
  r: number
  dx: number
  dy: number
}

export interface SheetDrawing {
  kind: 'picture' | 'chart'
  from?: DrawingCorner
  to?: DrawingCorner
  /** position from the sheet's top left, px (absoluteAnchor) */
  at?: { x: number; y: number }
  /** px (oneCellAnchor, absoluteAnchor) */
  size?: { w: number; h: number }
  /** a picture's media part, where the browser can draw it */
  target?: string
  /** a chart's title, or a picture's description */
  title: string
  /** a chart's kind: 'column', 'bar', 'line', 'pie', … ('' where it is none the reader knows) */
  chart?: string
  /** a chart's model with its values, for drawing it (chart-svg.ts); absent where it is a kind not drawn (a placeholder) */
  model?: ChartModel
}

/** DrawingML measures in EMU: 914400 to the inch, 96 CSS pixels to the inch */
const EMU_PER_PX = 9525

const CHART_KINDS: Record<string, string> = dict({
  barChart: 'bar', bar3DChart: 'bar', lineChart: 'line', line3DChart: 'line', pieChart: 'pie', pie3DChart: 'pie', ofPieChart: 'pie',
  doughnutChart: 'doughnut', areaChart: 'area', area3DChart: 'area', scatterChart: 'scatter', bubbleChart: 'bubble', radarChart: 'radar',
  stockChart: 'stock', surfaceChart: 'surface', surface3DChart: 'surface',
})

/** an attribute whatever its namespace prefix: `r:id` is the usual, not the only one */
const attr = (el: XEl | undefined, local: string) => {
  if (!el) return undefined
  for (const [k, v] of Object.entries(el.a)) if (k === local || k.endsWith(':' + local)) return v
  return undefined
}

function emu(value: string | undefined, maxPx: number, signed = false): number | undefined {
  const n = ooxmlInt(value, signed ? -maxPx * EMU_PER_PX : 0, maxPx * EMU_PER_PX)
  return n === undefined ? undefined : n / EMU_PER_PX
}

function corner(el: XEl | undefined): DrawingCorner | undefined {
  if (!el) return undefined
  const c = ooxmlInt(textOf(kid(el, 'col')), 0, MAX_COL - 1)
  const r = ooxmlInt(textOf(kid(el, 'row')), 0, MAX_ROWS - 1)
  // an offset left out is none: the corner is at the cell's own (a wrong one still discards the anchor)
  const off = (name: string, max: number) => (kid(el, name) ? emu(textOf(kid(el, name)), max, true) : 0)
  const dx = off('colOff', MAX_SHEET_WIDTH_PX)
  const dy = off('rowOff', MAX_SHEET_HEIGHT_PX)
  if (c === undefined || r === undefined || dx === undefined || dy === undefined) return undefined
  return { c, r, dx, dy }
}

/** Reject an invalid anchor as a whole; do not move corrupt drawings to the sheet's last cell. */
function placement(anchor: XEl): Pick<SheetDrawing, 'from' | 'to' | 'at' | 'size'> | undefined {
  const from = corner(kid(anchor, 'from'))
  if (anchor.n === 'twoCellAnchor') {
    const to = corner(kid(anchor, 'to'))
    return from && to ? { from, to } : undefined
  }
  const ext = kid(anchor, 'ext')
  const w = emu(ext?.a.cx, MAX_SHEET_WIDTH_PX)
  const h = emu(ext?.a.cy, MAX_SHEET_HEIGHT_PX)
  if (w === undefined || h === undefined) return undefined
  const size = { w, h }
  if (anchor.n === 'oneCellAnchor') return from ? { from, size } : undefined
  const pos = kid(anchor, 'pos')
  const x = emu(pos?.a.x, MAX_SHEET_WIDTH_PX, true)
  const y = emu(pos?.a.y, MAX_SHEET_HEIGHT_PX, true)
  return x !== undefined && y !== undefined ? { at: { x, y }, size } : undefined
}

/** a chart's kind, title and model (null where it is a kind chart.ts does not draw), from its chart part */
async function chartOf(wb: XlsxWorkbook, sheet: number, part: string): Promise<{ chart: string; title: string; model: ChartModel | null }> {
  const root = await wb.pkg.xml(part)
  const plot = desc(root, 'plotArea')
  const type = plot?.c.find((e) => e.n.endsWith('Chart'))
  let chart = type ? (CHART_KINDS[type.n] ?? '') : ''
  // a bar chart is drawn with upright columns unless its direction says bars
  if (type?.n.startsWith('bar') && attr(kid(type, 'barDir'), 'val') !== 'bar') chart = 'column'
  const title = desc(kid(root, 'chart'), 'title')
  const model = await readChart(root, { theme: wb.theme, date1904: wb.date1904, read: workbookReader(wb, sheet) }).catch((e: Error) => {
    console.warn('[office-engine] chart not read:', e.message)
    return null
  })
  return { chart, title: model?.title?.text ?? (title ? descAll(title, 't').map((t) => t.t).join('').trim() : ''), model }
}

/**
 * Reads a chart's series references from the workbook's cells: `'Sheet 1'!$B$2:$B$9`, several areas, a sheet not shown
 * yet (loaded for it). Null for a defined name, another workbook or an unknown sheet: the chart's saved values stand in.
 */
export function workbookReader(wb: XlsxWorkbook, self: number): RangeReader {
  return async (ref) => {
    let refs
    try {
      refs = references(parseFormula(ref))
    } catch {
      return null
    }
    if (!refs.length) return null
    const out: RangeValues = { values: [], formats: [], hidden: [] }
    // the cells of all the reference's areas together, so a reference of many areas is no larger than one
    let cells = 0
    for (const r of refs) {
      if (r.sheet?.startsWith('[')) return null
      const s = r.sheet === undefined ? self : wb.sheets.findIndex((x) => x.name.toLowerCase() === r.sheet!.toLowerCase())
      if (s < 0) return null
      if (!wb.sheetsData.get(s)?.done) await wb.loadSheet(s, () => {})
      // formulas still being computed after the sheet was read (no cached value, a full calculation on load)
      await wb.settled()
      const d = wb.sheetsData.get(s)
      if (!d || d.error) return null
      // a whole column or row (A:A, 2:2) is read as far as the sheet is filled; a range given in full keeps its size, so its
      // empty cells at the end stay points and line up with categories read from elsewhere
      const [r0, c0, r1, c1] =
        'r' in r
          ? [r.r, r.c, r.r, r.c]
          : [r.r0, r.c0, r.r1 >= MAX_ROWS - 1 ? Math.min(r.r1, Math.max(r.r0, d.maxRow)) : r.r1, r.c1 >= MAX_COL - 1 ? Math.min(r.c1, Math.max(r.c0, d.maxCol)) : r.c1]
      cells += (r1 - r0 + 1) * (c1 - c0 + 1)
      if (cells > MAX_POINTS) return null
      await wb.ensure(d, r0, r1)
      const hiddenCols = new Set<number>()
      for (const ci of d.colInfo) if (ci.w === 0) for (let c = Math.max(ci.min, c0); c <= Math.min(ci.max, c1); c++) hiddenCols.add(c)
      for (let rr = r0; rr <= r1; rr++) {
        const row = d.rows[rr]
        const cells = new Map<number, [number, unknown, number]>()
        if (row) for (let k = 0; k < row.length; k += 4) cells.set(row[k] as number, [row[k + 1] as number, row[k + 2], row[k + 3] as number])
        const rowHidden = !!d.rowAttrs.get(rr)?.hidden
        for (let c = c0; c <= c1; c++) {
          const cell = cells.get(c)
          out.values.push(cell ? cellValue(wb, cell[0], cell[1]) : null)
          out.formats.push(cell ? wb.styles.xfs[cell[2]]?.numFmt : undefined)
          out.hidden.push(rowHidden || hiddenCols.has(c))
        }
      }
    }
    return out
  }
}

/** more cells than a chart reads from one reference, all its areas together (a whole column): its saved values are used instead */
const MAX_POINTS = 50_000

function cellValue(wb: XlsxWorkbook, t: number, v: unknown): number | string | null {
  switch (t) {
    case T_NUM:
      return typeof v === 'number' ? v : Number(v)
    // a date cell holds its date as ISO text; a chart plots its serial, and shows it in the cell's date format
    case T_DATE:
      return typeof v === 'number' ? v : isoToSerial(String(v), wb.date1904)
    case T_SST:
      return wb.sst[v as number] ?? ''
    case T_BOOL:
      return v ? 1 : 0
    case T_EMPTY:
    case T_ERR:
      return null
    default:
      return String(v)
  }
}

/** The pictures and charts on the workbook's sheet `sheet`, in the order they are drawn */
export async function readSheetDrawings(wb: XlsxWorkbook, sheet: number): Promise<SheetDrawing[]> {
  const pkg = wb.pkg
  const sheetPath = wb.sheets[sheet]?.path
  if (!sheetPath) return []
  const sheetRels = await pkg.rels(sheetPath)
  const out: SheetDrawing[] = []
  // a chart part is read once, however many anchors show it
  const charts = new Map<string, ReturnType<typeof chartOf>>()
  const chartFor = (part: string) => {
    let c = charts.get(part)
    if (!c) charts.set(part, (c = chartOf(wb, sheet, part)))
    return c
  }
  for (const rel of sheetRels.values()) {
    if (rel.external || !rel.type.endsWith('/drawing') || !pkg.has(rel.target)) continue
    const root = await pkg.xml(rel.target)
    const rels = await pkg.rels(rel.target)
    for (const anchor of root?.c ?? []) {
      if (anchor.n !== 'twoCellAnchor' && anchor.n !== 'oneCellAnchor' && anchor.n !== 'absoluteAnchor') continue
      const place = placement(anchor)
      if (!place) continue
      const chartRef = desc(anchor, 'chart')
      const blip = desc(anchor, 'blip')
      if (chartRef) {
        const target = rels.get(attr(chartRef, 'id') ?? '')
        const chart = target && !target.external && pkg.has(target.target) ? await chartFor(target.target) : { chart: '', title: '', model: null }
        out.push({ kind: 'chart', ...place, chart: chart.chart, title: chart.title, ...(chart.model ? { model: { ...chart.model } } : {}) })
      } else if (blip) {
        const target = rels.get(attr(blip, 'embed') ?? '')
        const pr = desc(anchor, 'cNvPr')
        out.push({
          kind: 'picture',
          ...place,
          target: target && !target.external && isWebImage(target.target) ? target.target : undefined,
          title: pr?.a.descr || pr?.a.title || '',
        })
      }
    }
  }
  return out
}
