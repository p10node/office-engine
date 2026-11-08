import { describe, expect, test } from 'bun:test'
import { Engine } from '../src/core/worker/engine'
import { tableSpec } from '../src/core/docx/spec'
import type { Block, Paragraph, Table, TableCell } from '../src/core/docx/types'
import { borderXml } from '../src/core/docx/writer'
import { cellsIn, distributeCols, gridOf, mergeCells, newTable, patchCells, rectOf, setBorders, unmergeCell } from '../src/react/docx/tables'

const S = (n: string) => `${import.meta.dir}/../playground/public/samples/${n}`

let pid = 5_000_000
const para = (text = ''): Paragraph => ({ t: 'p', ps: {}, runs: text ? [{ t: 'text', text, s: {} }] : [], pid: pid++, edited: true })
const mk = () => para()
const textOf = (b: Block | undefined) => (b && b.t === 'p' ? b.runs.map((r) => (r.t === 'text' ? r.text : '')).join('') : '')
const cellText = (c: TableCell) => c.blocks.map(textOf).join('|')

/** 3×3 table with "r,c" in every cell */
function grid3(): Table {
  const t = newTable(3, 3, 300, mk)
  t.rows.forEach((row, ri) => row.cells.forEach((c, ci) => (c.blocks = [para(`${ri},${ci}`)])))
  return t
}

describe('table cell ranges', () => {
  test('merge a 2×2 block keeps every text, in reading order', () => {
    const t = grid3()
    const r = rectOf(t, [0, 0], [1, 1])!
    expect(r).toEqual({ r0: 0, g0: 0, r1: 1, g1: 1 })
    const m = mergeCells(t, r)!
    expect(m.rows[0].cells.length).toBe(2)
    expect(m.rows[1].cells.length).toBe(1)
    expect(m.rows[0].cells[0].colSpan).toBe(2)
    expect(m.rows[0].cells[0].rowSpan).toBe(2)
    expect(cellText(m.rows[0].cells[0])).toBe('0,0|0,1|1,0|1,1')
    // grid positions still line up
    expect(gridOf(m).at(1, 2)?.cell.blocks.map(textOf)).toEqual(['1,2'])
  })

  test('a range touching a merged cell grows to cover it', () => {
    const t = mergeCells(grid3(), { r0: 0, g0: 1, r1: 1, g1: 1 })!
    // row 1 now has cells (1,0) and (1,2); selecting (1,0)→(0,0) must not cut the merged cell at column 1
    const r = rectOf(t, [1, 0], [0, 1])!
    expect(r).toEqual({ r0: 0, g0: 0, r1: 1, g1: 1 })
  })

  test('only blank cells: the top-left cell keeps its paragraph', () => {
    const t = newTable(2, 2, 200, mk)
    const m = mergeCells(t, { r0: 0, g0: 0, r1: 0, g1: 1 })!
    expect(m.rows[0].cells[0].blocks.length).toBe(1)
  })

  test('unmerge restores a full grid (text stays top-left)', () => {
    const m = mergeCells(grid3(), { r0: 0, g0: 0, r1: 1, g1: 1 })!
    const u = unmergeCell(m, 0, 0, mk)!
    expect(u.rows.map((r) => r.cells.length)).toEqual([3, 3, 3])
    expect(u.rows.every((r) => r.cells.every((c) => c.colSpan === 1 && c.rowSpan === 1))).toBe(true)
    expect(cellText(u.rows[0].cells[0])).toBe('0,0|0,1|1,0|1,1')
    expect(cellText(u.rows[1].cells[2])).toBe('1,2')
    expect(unmergeCell(u, 0, 0, mk)).toBeNull()
  })

  test('outer borders reach the neighbours; none clears; inner leaves the edge', () => {
    const t = grid3()
    const thick = '3pt solid #ff0000'
    const o = setBorders(t, { r0: 1, g0: 1, r1: 1, g1: 1 }, 'outer', thick)
    const c = o.rows[1].cells[1]
    expect([c.borders.t, c.borders.r, c.borders.b, c.borders.l]).toEqual([thick, thick, thick, thick])
    expect(o.rows[0].cells[1].borders.b).toBe(thick)
    expect(o.rows[2].cells[1].borders.t).toBe(thick)
    expect(o.rows[1].cells[0].borders.r).toBe(thick)
    expect(o.rows[1].cells[2].borders.l).toBe(thick)
    expect(o.rows[0].cells[0].set).toBeUndefined()
    const n = setBorders(t, { r0: 0, g0: 0, r1: 2, g1: 2 }, 'none', thick)
    expect(n.rows.flatMap((r) => r.cells.flatMap((x) => Object.values(x.borders))).every((v) => v === 'none')).toBe(true)
    const i = setBorders(t, { r0: 0, g0: 0, r1: 1, g1: 1 }, 'inner', thick)
    expect(i.rows[0].cells[0].borders.r).toBe(thick)
    expect(i.rows[0].cells[0].borders.t).not.toBe(thick)
  })

  test('shading, vertical alignment and distribution touch only the range', () => {
    const t = grid3()
    t.cols = [50, 100, 150]
    const s = patchCells(t, { r0: 0, g0: 0, r1: 0, g1: 1 }, (c) => ({ ...c, shd: '#ffcc00', set: { shd: true } }))
    expect(s.rows[0].cells.map((c) => c.shd)).toEqual(['#ffcc00', '#ffcc00', undefined])
    expect(distributeCols(t).cols).toEqual([100, 100, 100])
    expect(distributeCols(t, { r0: 0, g0: 0, r1: 0, g1: 1 }).cols).toEqual([75, 75, 150])
    expect(cellsIn(t, { r0: 2, g0: 2, r1: 2, g1: 2 }).length).toBe(1)
  })

  test('border lines are written the way Word reads them', () => {
    expect(borderXml('w:', 'top', '0.5pt solid #000000')).toBe('<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>')
    expect(borderXml('w:', 'left', '3pt dashed #FF0000')).toBe('<w:left w:val="dashed" w:sz="24" w:space="0" w:color="FF0000"/>')
    expect(borderXml('w:', 'bottom', 'none')).toBe('<w:bottom w:val="nil"/>')
  })
})

describe('table cell formatting survives save + reopen', () => {
  async function open(src: Blob) {
    const events: { event: string; data: any }[] = []
    const engine = new Engine((event, data) => events.push({ event, data }))
    await engine.open(src, {})
    await engine.call('docx.start', [])
    const blocks: Block[] = events.filter((e) => e.event === 'docx:blocks').flatMap((e) => e.data)
    return { engine, blocks }
  }

  test('merge (both ways), borders, fill, no fill and vertical alignment', async () => {
    const { engine, blocks } = await open(new Blob([await Bun.file(S('sample.docx')).arrayBuffer()]))
    const orig = blocks.find((b): b is Table => b.t === 'tbl')!
    expect(orig.rows.length).toBeGreaterThan(2)
    let t: Table = { ...orig, tid: 9_100_001 }
    // merge row 1 cells 0..1 (across) and rows 2..3 of column 2 (down)
    t = mergeCells(t, rectOf(t, [1, 0], [1, 1])!)!
    const g2 = gridOf(t)
    const down = rectOf(t, [2, g2.cells.find((c) => c.ri === 2 && c.g === 2)!.ci], [3, g2.cells.find((c) => c.ri === 3 && c.g === 2)!.ci])!
    t = mergeCells(t, down)!
    t = setBorders(t, { r0: 0, g0: 0, r1: 0, g1: 0 }, 'all', '3pt dashed #ff0000')
    t = patchCells(t, { r0: 0, g0: 1, r1: 0, g1: 1 }, (c) => ({ ...c, shd: '#00ff00', set: { ...c.set, shd: true } }))
    // the header row has a fill from the file: take it away on cell (0,2)
    t = patchCells(t, { r0: 0, g0: 2, r1: 0, g1: 2 }, (c) => ({ ...c, shd: undefined, set: { ...c.set, shd: true } }))
    t = patchCells(t, { r0: 1, g0: 0, r1: 1, g1: 1 }, (c) => ({ ...c, vAlign: 'bottom', set: { ...c.set, vAlign: true } }))
    await engine.call('docx.table', [{ tid: t.tid, item: t.item, spec: tableSpec(t) }])

    const re = await open((await engine.call('save', [])) as Blob)
    const r = re.blocks.find((b): b is Table => b.t === 'tbl')!
    expect(r.rows[1].cells[0].colSpan).toBe(2)
    expect(r.rows[1].cells[0].vAlign).toBe('bottom')
    const merged = gridOf(r).at(2, 2)!
    expect(merged.ri).toBe(2)
    expect(merged.cell.rowSpan).toBe(2)
    const c00 = r.rows[0].cells[0]
    expect([c00.borders.t, c00.borders.r, c00.borders.b, c00.borders.l]).toEqual(new Array(4).fill('3pt dashed #ff0000'))
    expect(r.rows[0].cells[1].shd).toBe('#00ff00')
    expect(r.rows[0].cells[2].shd).toBeUndefined()
    expect(r.rows.length).toBe(orig.rows.length)
  })
})
