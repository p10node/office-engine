import { describe, expect, test } from 'bun:test'
import type { Table } from '../src/core/docx/types'
import { gridOf } from '../src/react/docx/tables'
import { docxBlocks, useBrowser, viewer, type Viewer } from './harness'

useBrowser()

const open = async () => {
  const v = await viewer({ src: 'samples/sample.docx', editable: true, onSave: true })
  await v.page.locator('.de-ebar').waitFor()
  return v
}

/** a mouse drag from the middle of one table cell's text to another's */
async function dragCells(v: Viewer, from: string, to: string) {
  const td = (t: string) => v.page.locator('.de-pages td', { hasText: new RegExp(`^${t}$`) }).first()
  await td(from).scrollIntoViewIfNeeded()
  const a = (await td(from).boundingBox())!
  const b = (await td(to).boundingBox())!
  await v.page.mouse.move(a.x + 8, a.y + a.height / 2)
  await v.page.mouse.down()
  await v.page.mouse.move(b.x + b.width / 2, b.y + b.height / 2, { steps: 8 })
  await v.page.mouse.up()
  await v.page.waitForTimeout(150)
}

const savedTable = async (v: Viewer) => (await docxBlocks((await v.save()).bytes)).find((b): b is Table => b.t === 'tbl')!

describe('docx table formatting', () => {
  test('select cells by dragging, merge them, undo, unmerge', async () => {
    const v = await open()
    await dragCells(v, 'North', '140')
    expect(await v.page.locator('.de-pages td.de-cell-sel').count()).toBe(3)
    expect(await v.btn('Merge cells').isEnabled()).toBe(true)
    await v.btn('Merge cells').click()
    await v.page.waitForTimeout(200)
    expect(await v.page.locator('.de-pages td[colspan="3"]').count()).toBe(2) // + the file's merged footer
    let t = await savedTable(v)
    expect(t.rows[1].cells[0].colSpan).toBe(3)
    expect(t.rows[1].cells[0].blocks.map((p) => (p.t === 'p' ? p.runs.map((r) => (r.t === 'text' ? r.text : '')).join('') : ''))).toEqual(['North', '120', '140'])

    // caret in the merged cell: the same button splits it
    await v.selectText('North', 'end')
    expect(await v.btn('Unmerge cells').isEnabled()).toBe(true)
    await v.btn('Unmerge cells').click()
    await v.page.waitForTimeout(200)
    t = await savedTable(v)
    expect(t.rows[1].cells.length).toBe(4)

    await v.page.keyboard.press('Control+z')
    await v.page.keyboard.press('Control+z')
    await v.page.waitForTimeout(250)
    t = await savedTable(v)
    expect(t.rows[1].cells.map((c) => c.colSpan)).toEqual([1, 1, 1, 1])
    expect(v.errors).toEqual([])
  })

  test('vertical merge down a column', async () => {
    const v = await open()
    await dragCells(v, '120', '132')
    await v.btn('Merge cells').click()
    await v.page.waitForTimeout(200)
    const t = await savedTable(v)
    const g = gridOf(t).at(3, 1)!
    expect(g.ri).toBe(1)
    expect(g.cell.rowSpan).toBe(3)
    expect(v.errors).toEqual([])
  })

  test('cell color, borders and vertical alignment are saved', async () => {
    const v = await open()
    await dragCells(v, 'South', '110')
    await v.btn('Cell color').click()
    const sw = v.page.locator('.de-tb-color .de-tb-grid button[title^="#"]').nth(3)
    const color = (await sw.getAttribute('title'))!.toLowerCase()
    await sw.click()
    await v.page.waitForTimeout(150)

    await dragCells(v, 'South', '110')
    await v.btn('Cell borders').click()
    await v.page.locator('button[aria-label="Outer borders"]').click()
    await v.page.waitForTimeout(150)

    await v.selectText('South', 'end')
    await v.pick('Vertical align', 'Middle')
    await v.page.waitForTimeout(150)

    const t = await savedTable(v)
    const row = t.rows[2]
    expect(row.cells.slice(0, 3).map((c) => c.shd)).toEqual([color, color, color])
    expect(row.cells[3].shd).not.toBe(color)
    expect(row.cells[0].borders.l).toBe('0.5pt solid #000000')
    expect(row.cells[2].borders.r).toBe('0.5pt solid #000000')
    expect(row.cells[0].vAlign).toBe('middle')
    expect(v.errors).toEqual([])
  })

  test('dragging a column edge resizes the columns', async () => {
    const v = await open()
    await v.selectText('Region', 'end')
    const before = (await savedTable(v)).cols
    const h = v.page.locator('.de-col-resize').first()
    await h.scrollIntoViewIfNeeded()
    const box = (await h.boundingBox())!
    await v.page.mouse.move(box.x + box.width / 2, box.y + 20)
    await v.page.mouse.down()
    await v.page.mouse.move(box.x + box.width / 2 + 60, box.y + 20, { steps: 6 })
    await v.page.mouse.up()
    await v.page.waitForTimeout(200)
    const after = (await savedTable(v)).cols
    expect(after[0]).toBeGreaterThan(before[0] + 20)
    expect(after[1]).toBeLessThan(before[1] - 20)
    expect(after.reduce((a, b) => a + b, 0)).toBeCloseTo(before.reduce((a, b) => a + b, 0), 0)
    expect(v.errors).toEqual([])
  })
})
