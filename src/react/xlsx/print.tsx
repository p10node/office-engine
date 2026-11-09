import type { CSSProperties, ReactNode } from 'react'
import type { DocumentClient } from '../../core/worker/client'
import type { CellOut, Merge, WorkbookMeta } from '../../core/xlsx/types'
import type { Printable } from '../file/print'
import type { Axis } from './axis'
import { fitNumber, fontPx } from './fit'

export interface SheetLayout {
  rows: number
  cols: number
  rowAxis: Axis
  colAxis: Axis
  merges: Merge[]
  gridLines: boolean
  done: boolean
}

const CHUNK = 2000
const MAX_PRINT_ROWS = 20_000
/** Rows per `<table>`: Chrome paginates one long table in worse than linear time; short tables stay linear. */
const TABLE_ROWS = 100
/** printable width (px) of portrait / landscape paper after margins (A4 and Letter) */
const PORTRAIT_W = 700
const LANDSCAPE_W = 960

async function sheetRows(client: DocumentClient, sheet: number, rows: number): Promise<Map<number, CellOut[]>> {
  const out = new Map<number, CellOut[]>()
  for (let r0 = 0; r0 < rows; r0 += CHUNK) {
    const b = await client.call('xlsx.block', sheet, r0, Math.min(rows, r0 + CHUNK))
    for (const row of b.rows) out.set(row.r, row.cells)
  }
  return out
}

const loading = () => new Error('The sheet is still loading — try again in a moment')

/** Shown values of the sheet as CSV / TSV (UTF-8 BOM for CSV so Excel reads accents). */
const VALIGN: Record<string, CSSProperties['verticalAlign']> = { 'flex-start': 'top', center: 'middle', 'flex-end': 'bottom' }

/** The sheet as a table (used range, hidden rows / columns left out), fitted to the paper width. */
export async function sheetPrintable(client: DocumentClient, sheet: number, l: SheetLayout, meta: WorkbookMeta): Promise<Printable> {
  if (!l.done) throw loading()
  const data = await sheetRows(client, sheet, Math.min(l.rows, MAX_PRINT_ROWS))
  let lastR = -1
  let lastC = -1
  for (const [r, cells] of data) {
    for (const c of cells) {
      if (c[1] === '' && !c[2]) continue
      lastR = Math.max(lastR, r)
      lastC = Math.max(lastC, c[0])
    }
  }
  const merges = l.merges.filter((m) => m.r0 <= lastR && m.c0 <= lastC)
  for (const m of merges) {
    lastR = Math.max(lastR, Math.min(m.r1, MAX_PRINT_ROWS - 1))
    lastC = Math.max(lastC, m.c1)
  }
  const cols: number[] = []
  for (let c = 0; c <= lastC; c++) if (l.colAxis.size(c) > 0) cols.push(c)
  const anchors = new Map<number, Merge>()
  const covered = new Set<number>()
  const key = (r: number, c: number) => r * 16384 + c
  // a row may start a new table unless a merge runs across it
  const across = new Int32Array(lastR + 2)
  for (const m of merges) {
    anchors.set(key(m.r0, m.c0), m)
    const r1 = Math.min(m.r1, lastR)
    for (let r = m.r0; r <= r1; r++) for (let c = m.c0; c <= m.c1; c++) if (r !== m.r0 || c !== m.c0) covered.add(key(r, c))
    if (r1 > m.r0) {
      across[m.r0 + 1]++
      across[r1 + 1]--
    }
  }
  const visibleIn = (axis: Axis, a: number, b: number) => {
    let n = 0
    for (let i = a; i <= b; i++) if (axis.size(i) > 0) n++
    return Math.max(1, n)
  }

  const basePx = fontPx(meta.baseFont.fontSize, 14.67)
  const tables: ReactNode[][] = []
  let trs: ReactNode[] = []
  let inMerge = 0
  for (let r = 0; r <= lastR; r++) {
    inMerge += across[r]
    if (!inMerge && trs.length >= TABLE_ROWS) {
      tables.push(trs)
      trs = []
    }
    const h = l.rowAxis.size(r)
    if (!h) continue
    const cells = new Map((data.get(r) ?? []).map((c) => [c[0], c]))
    const tds: ReactNode[] = []
    cols.forEach((c, k) => {
      if (covered.has(key(r, c))) return
      const cell = cells.get(c)
      const m = anchors.get(key(r, c))
      const [, raw = '', si = 0, kind = 0, color] = cell ?? []
      const { textAlign, textRotate: _rot, alignItems, whiteSpace, ...rest } = (meta.styles[si] ?? {}) as Record<string, string | number>
      // like the grid: fewer decimals or ### when a number does not fit
      const w = m ? l.colAxis.pos(Math.min(m.c1, lastC) + 1) - l.colAxis.pos(m.c0) : l.colAxis.size(c)
      const text = kind === 1 && raw ? fitNumber(raw, w - 6, fontPx(rest.fontSize, basePx)) : raw
      const align = (textAlign as string) ?? (kind === 1 ? 'right' : kind === 2 ? 'center' : 'left')
      // like the grid: left-aligned text runs over empty neighbours (and past the last column)
      const next = cols[k + 1]
      const spill = !!text && !m && align === 'left' && !whiteSpace && (next === undefined || (!cells.get(next)?.[1] && !covered.has(key(r, next))))
      // only what differs from the `.de-print-table td` defaults (bottom, pre, left): keeps a big sheet's DOM small
      const style = { ...rest } as CSSProperties
      if (align !== 'left') style.textAlign = align as CSSProperties['textAlign']
      const valign = VALIGN[alignItems as string]
      if (valign && valign !== 'bottom') style.verticalAlign = valign
      if (whiteSpace) style.whiteSpace = whiteSpace as CSSProperties['whiteSpace']
      if (color) style.color = color
      tds.push(
        <td
          key={c}
          style={style}
          data-de-fit={text && !spill ? (kind === 1 ? '#' : '') : undefined}
          colSpan={m ? visibleIn(l.colAxis, m.c0, Math.min(m.c1, lastC)) : undefined}
          rowSpan={m ? visibleIn(l.rowAxis, m.r0, Math.min(m.r1, lastR)) : undefined}
        >
          {text}
        </td>,
      )
    })
    trs.push(
      <tr key={r} style={{ height: h }}>
        {tds}
      </tr>,
    )
  }
  if (trs.length || !tables.length) tables.push(trs)

  const width = cols.reduce((w, c) => w + l.colAxis.size(c), 0)
  const landscape = width > PORTRAIT_W
  const fit = Math.max(0.35, Math.min(1, (landscape ? LANDSCAPE_W : PORTRAIT_W) / Math.max(1, width)))
  const className = l.gridLines ? 'de-print-table de-print-grid' : 'de-print-table'
  const tableStyle = { width, ...(meta.baseFont as CSSProperties) }
  const colgroup = (
    <colgroup>
      {cols.map((c) => (
        <col key={c} style={{ width: l.colAxis.size(c) }} />
      ))}
    </colgroup>
  )
  const node = (
    <div className="de-print-sheet de-print-fit" style={{ '--de-fit': fit } as CSSProperties}>
      {tables.map((rows, i) => (
        <table key={i} className={className} style={tableStyle}>
          {colgroup}
          <tbody>{rows}</tbody>
        </table>
      ))}
    </div>
  )
  return {
    node,
    size: landscape ? 'landscape' : 'auto',
    margin: '12mm',
    notice: l.rows > MAX_PRINT_ROWS ? `Only the first ${MAX_PRINT_ROWS.toLocaleString()} rows are included` : undefined,
  }
}
