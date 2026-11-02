import type { Paragraph, Table, TableCell, TableRow } from '../../core/docx/types'

export interface GridCell {
  ri: number
  ci: number
  g: number
  cell: TableCell
}

/** Grid positions of every cell (row spans occupy following rows). */
export function gridOf(t: Table): { cells: GridCell[]; nCols: number; at: (ri: number, g: number) => GridCell | undefined } {
  const occ: boolean[][] = []
  const cells: GridCell[] = []
  t.rows.forEach((row, ri) => {
    let g = 0
    row.cells.forEach((cell, ci) => {
      while (occ[ri]?.[g]) g++
      for (let r = ri; r < ri + Math.max(1, cell.rowSpan); r++) for (let k = g; k < g + Math.max(1, cell.colSpan); k++) (occ[r] ??= [])[k] = true
      cells.push({ ri, ci, g, cell })
      g += Math.max(1, cell.colSpan)
    })
  })
  const nCols = Math.max(t.cols.length, ...cells.map((c) => c.g + c.cell.colSpan))
  const at = (ri: number, g: number) => cells.find((c) => ri >= c.ri && ri < c.ri + c.cell.rowSpan && g >= c.g && g < c.g + c.cell.colSpan)
  return { cells, nCols, at }
}

export type NewPara = (template: TableCell) => Paragraph

function blankCell(tpl: TableCell, mk: NewPara, colSpan = 1): TableCell {
  return { ...tpl, blocks: [mk(tpl)], colSpan, rowSpan: 1 }
}

function scaleCols(cols: number[], total: number): number[] {
  const sum = cols.reduce((a, b) => a + b, 0) || 1
  return cols.map((w) => +((w * total) / sum).toFixed(2))
}

export function insertRow(t: Table, ri: number, below: boolean, mk: NewPara): Table {
  const { cells, nCols, at } = gridOf(t)
  const target = below ? ri + 1 : ri
  const rows = t.rows.map((r) => ({ ...r, cells: [...r.cells] }))
  const newCells: TableCell[] = []
  for (let g = 0; g < nCols; ) {
    const above = at(below ? ri : ri, g)
    if (!above) {
      g++
      continue
    }
    const spansInto = above.ri < target && above.ri + above.cell.rowSpan > target
    if (spansInto) {
      const r = rows[above.ri]
      r.cells[above.ci] = { ...r.cells[above.ci], rowSpan: r.cells[above.ci].rowSpan + 1 }
    } else newCells.push(blankCell(above.cell, mk, above.cell.colSpan))
    g += above.cell.colSpan
  }
  void cells
  const tplRow = t.rows[ri]
  const row: TableRow = { ...tplRow, cells: newCells }
  rows.splice(target, 0, row)
  return { ...t, rows }
}

export function deleteRow(t: Table, ri: number): Table | null {
  if (t.rows.length <= 1) return null
  const { cells } = gridOf(t)
  const rows = t.rows.map((r) => ({ ...r, cells: [...r.cells] }))
  for (const c of cells) {
    if (c.ri < ri && c.ri + c.cell.rowSpan > ri) rows[c.ri].cells[c.ci] = { ...rows[c.ri].cells[c.ci], rowSpan: c.cell.rowSpan - 1 }
  }
  rows.splice(ri, 1)
  return { ...t, rows }
}

export function insertCol(t: Table, g0: number, right: boolean, mk: NewPara): Table {
  const { cells, nCols } = gridOf(t)
  const g = right ? g0 + 1 : g0
  const rows = t.rows.map((r) => ({ ...r, cells: [...r.cells] }))
  const total = t.cols.reduce((a, b) => a + b, 0)
  t.rows.forEach((_, ri) => {
    const rowCells = cells.filter((c) => c.ri === ri)
    const crossing = cells.find((c) => c.ri <= ri && ri < c.ri + c.cell.rowSpan && c.g < g && c.g + c.cell.colSpan > g)
    if (crossing) {
      if (crossing.ri === ri) rows[ri].cells[crossing.ci] = { ...rows[ri].cells[crossing.ci], colSpan: crossing.cell.colSpan + 1 }
      return
    }
    const covered = cells.find((c) => c.ri < ri && ri < c.ri + c.cell.rowSpan && c.g + c.cell.colSpan === g)
    if (covered && !rowCells.length) return
    const before = rowCells.filter((c) => c.g < g).length
    const tpl = rowCells[Math.max(0, before - 1)]?.cell ?? rowCells[0]?.cell
    if (tpl) rows[ri].cells.splice(before, 0, blankCell(tpl, mk))
  })
  const cols = [...(t.cols.length ? t.cols : new Array(nCols).fill(total / Math.max(1, nCols)))]
  cols.splice(g, 0, cols[Math.min(g0, cols.length - 1)] ?? 72)
  return { ...t, rows, cols: total ? scaleCols(cols, total) : cols }
}

export function deleteCol(t: Table, g: number): Table | null {
  const { cells, nCols } = gridOf(t)
  if (nCols <= 1) return null
  const rows = t.rows.map((r) => ({ ...r, cells: [...r.cells] }))
  const remove = new Set<string>()
  for (const c of cells) {
    if (g < c.g || g >= c.g + c.cell.colSpan) continue
    if (c.cell.colSpan > 1) rows[c.ri].cells[c.ci] = { ...rows[c.ri].cells[c.ci], colSpan: c.cell.colSpan - 1 }
    else remove.add(`${c.ri}:${c.ci}`)
  }
  rows.forEach((r, ri) => (r.cells = r.cells.filter((_, ci) => !remove.has(`${ri}:${ci}`))))
  const total = t.cols.reduce((a, b) => a + b, 0)
  const cols = [...t.cols]
  cols.splice(g, 1)
  return { ...t, rows: rows.filter((r) => r.cells.length), cols: total && cols.length ? scaleCols(cols, total) : cols }
}

/** New table model (rows × cols) with simple single borders. */
export function newTable(rowsN: number, colsN: number, width: number, mk: () => Paragraph): Table {
  const border = '0.5pt solid #000000'
  const w = +(width / colsN).toFixed(2)
  return {
    t: 'tbl',
    cols: new Array(colsN).fill(w),
    rows: Array.from({ length: rowsN }, () => ({
      cells: Array.from({ length: colsN }, () => ({ blocks: [mk()], colSpan: 1, rowSpan: 1, borders: { t: border, r: border, b: border, l: border }, pad: [0, 5.4, 0, 5.4] as [number, number, number, number] })),
    })),
  }
}
