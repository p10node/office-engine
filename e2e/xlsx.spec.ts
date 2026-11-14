import { describe, expect, test } from 'bun:test'
import { useBrowser, viewer, xlsxSheet, type Viewer } from './harness'

useBrowser()

const open = async () => {
  const v = await viewer({ src: 'samples/sample.xlsx', editable: true, onSave: true })
  await v.page.locator('.de-ebar').waitFor()
  return v
}
const saved = async (v: Viewer, sheet = 0) => xlsxSheet((await v.save()).bytes, sheet)

/** the rendered cell element under a reference: box + inline style */
async function cellBox(v: Viewer, ref: string) {
  const p = await v.cellPoint(ref)
  return v.page.evaluate(({ x, y }) => {
    const cell = document.elementsFromPoint(x, y).map((e) => e.closest<HTMLElement>('.de-cell')).find(Boolean)
    if (!cell) return null
    const r = cell.getBoundingClientRect()
    return { x: r.x, y: r.y, w: r.width, h: r.height, text: cell.innerText, style: cell.getAttribute('style') ?? '', weight: getComputedStyle(cell).fontWeight }
  }, p)
}

/** topmost element at the boundary between two cells (a gridline unless something covers it) */
async function topAt(v: Viewer, a: string, b: string) {
  const p = await v.cellPoint(a)
  const q = await v.cellPoint(b)
  const x = p.x === q.x ? p.x : (p.x + q.x) / 2
  const y = p.y === q.y ? p.y : (p.y + q.y) / 2
  return v.page.evaluate(
    ({ x, y }) => {
      // the gridline sits 1px before the boundary: probe a 3px band
      const hits = [-1, 0, 1].flatMap((d) => document.elementsFromPoint(x + d * 0, y + d)[0]?.className ?? '')
      const hitsX = [-1, 0, 1].flatMap((d) => document.elementsFromPoint(x + d, y)[0]?.className ?? '')
      return [...hits, ...hitsX].join(' ')
    },
    { x, y },
  )
}

describe('xlsx editing', () => {
  test('values, formulas and recalculation', async () => {
    const v = await open()
    await v.typeCell('J10', '5')
    await v.typeCell('J11', '7')
    await v.typeCell('J12', '=SUM(J10:J11)')
    expect(await v.shown('J12')).toBe('12')
    expect(await v.formula('J12')).toBe('=SUM(J10:J11)')
    await v.typeCell('J10', '10')
    expect(await v.shown('J12')).toBe('17')
    expect((await v.events('dirty')).at(-1)?.dirty).toBe(true)
    const s = await saved(v)
    expect(s.cell('J12')).toBe('17')
    expect(v.errors).toEqual([])
  })

  test('Escape discards the edit: a new value, a changed value, and a formula', async () => {
    const v = await open()
    await v.typeCell('J10', 'keep')
    // typing into an empty cell, then Escape: stays empty
    await v.clickCell('J11')
    await v.page.keyboard.type('dropped')
    await v.page.keyboard.press('Escape')
    await v.page.waitForTimeout(150)
    expect(await v.shown('J11')).toBe('')
    // editing an existing value (F2), then Escape: the old value stays
    await v.clickCell('J10')
    await v.page.keyboard.press('F2')
    await v.page.keyboard.type(' changed')
    await v.page.keyboard.press('Escape')
    await v.page.waitForTimeout(150)
    expect(await v.shown('J10')).toBe('keep')
    // a formula with autocomplete open: the first Escape closes the list, the second discards
    await v.clickCell('J12')
    await v.page.keyboard.type('=SU')
    await v.page.locator('.de-fassist-at').waitFor()
    await v.page.keyboard.press('Escape')
    await v.page.keyboard.press('Escape')
    await v.page.waitForTimeout(150)
    expect(await v.shown('J12')).toBe('')
    expect(await v.page.evaluate(() => document.activeElement?.classList.contains('de-grid'))).toBe(true)
    expect((await v.events('dirty')).at(-1)?.dirty).toBe(true)
    const s = await saved(v)
    expect(s.cell('J10')).toBe('keep')
    expect(s.cell('J11')).toBe('')
    expect(s.cell('J12')).toBe('')
    expect(v.errors).toEqual([])
  })

  test('Shift+Enter and Alt+Enter break the line, in the cell and in the formula bar: the cell wraps and its row grows', async () => {
    const v = await open()
    const rowH = async (n: number) => (await v.page.locator('.de-head-r').filter({ hasText: new RegExp(`^${n}$`) }).first().boundingBox())!.height
    const h0 = await rowH(10)
    const fxRef = () => v.page.locator('.de-fx-ref').innerText()
    // in the cell
    await v.clickCell('J10')
    await v.page.keyboard.type('one')
    await v.page.keyboard.press('Shift+Enter')
    await v.page.keyboard.type('two')
    await v.page.keyboard.press('Alt+Enter')
    await v.page.keyboard.type('three')
    expect(await v.page.locator('textarea.de-cell-edit').inputValue()).toBe('one\ntwo\nthree')
    // Enter alone still commits and moves down
    await v.page.keyboard.press('Enter')
    await v.page.waitForTimeout(200)
    expect(await fxRef()).toBe('J11')
    expect(await v.shown('J10')).toBe('one\ntwo\nthree')
    expect((await cellBox(v, 'J10'))!.style).toContain('pre-wrap')
    expect(await rowH(10)).toBe(h0 * 3)
    // in the formula bar: one line shows at a time, the caret's
    await v.clickCell('J12')
    const bar = v.page.locator('.de-fx-input')
    await bar.click()
    await v.page.keyboard.type('a')
    await v.page.keyboard.press('Shift+Enter')
    await v.page.keyboard.type('b')
    await v.page.keyboard.press('Alt+Enter')
    await v.page.keyboard.type('c')
    expect(await bar.inputValue()).toBe('a\nb\nc')
    expect(await bar.evaluate((el) => el.scrollTop / parseFloat(getComputedStyle(el).lineHeight))).toBe(2)
    await v.page.keyboard.press('Enter')
    await v.page.waitForTimeout(200)
    expect(await v.shown('J12')).toBe('a\nb\nc')
    expect(await rowH(12)).toBe(h0 * 3)
    // the bar of a cell with several lines keeps them, from the first
    expect(await v.formula('J10')).toBe('one\ntwo\nthree')
    expect(await bar.evaluate((el) => el.scrollTop)).toBe(0)
    // on the grid Shift+Enter still moves up
    await v.clickCell('L5')
    await v.page.keyboard.press('Shift+Enter')
    expect(await fxRef()).toBe('L4')
    const s = await saved(v)
    expect(s.cell('J10')).toBe('one\ntwo\nthree')
    expect(s.cell('J12')).toBe('a\nb\nc')
    // one undo per cell: the text, the wrap and the row height
    await v.clickCell('N20')
    await v.page.keyboard.press('Control+z')
    await v.page.keyboard.press('Control+z')
    await v.page.waitForTimeout(300)
    expect(await v.shown('J10')).toBe('')
    expect(await rowH(10)).toBe(h0)
    expect(await rowH(12)).toBe(h0)
    expect(v.errors).toEqual([])
  })

  test('merge all: one cell, no inner gridlines, other values cleared', async () => {
    const v = await open()
    await v.typeCell('J10', 'merged')
    await v.typeCell('K11', 'lost')
    const a = (await cellBox(v, 'J10'))!
    await v.selectRange('J10', 'K12')
    await v.pick('Merge cells', 'Merge all')
    await v.page.waitForTimeout(300)
    await v.clickCell('N20')

    // the merged cell spans 2 columns × 3 rows and covers the gridlines inside
    const m = (await cellBox(v, 'K11'))!
    expect(m.text).toBe('merged')
    expect(Math.round(m.x)).toBe(Math.round(a.x))
    expect(m.w).toBeGreaterThan(a.w * 1.8)
    expect(m.h).toBeGreaterThan(a.h * 2.5)
    expect(m.style).toContain('linear-gradient')
    expect(await topAt(v, 'J11', 'K11')).not.toContain('de-gl-v')
    expect(await topAt(v, 'K10', 'K11')).not.toContain('de-gl-h')
    // outside the merge gridlines stay
    expect(await topAt(v, 'L11', 'M11')).toContain('de-gl-v')

    const s = await saved(v)
    expect(s.merges).toContainEqual({ r0: 9, c0: 9, r1: 11, c1: 10 })
    expect(s.cell('J10')).toBe('merged')
    expect(s.cell('K11')).toBe('')

    // clicking the merged cell selects it whole: Unmerge is offered
    await v.clickCell('J10')
    await v.pick('Merge cells', 'Unmerge')
    await v.page.waitForTimeout(300)
    expect((await cellBox(v, 'J10'))!.w).toBeLessThan(a.w * 1.2)
    expect((await saved(v)).merges).not.toContainEqual({ r0: 9, c0: 9, r1: 11, c1: 10 })
  })

  test('merge horizontally and vertically', async () => {
    const v = await open()
    await v.selectRange('J10', 'K11')
    await v.pick('Merge cells', 'Merge horizontally')
    await v.selectRange('M10', 'N11')
    await v.pick('Merge cells', 'Merge vertically')
    await v.page.waitForTimeout(300)
    const s = await saved(v)
    expect(s.merges).toEqual(
      expect.arrayContaining([
        { r0: 9, c0: 9, r1: 9, c1: 10 },
        { r0: 10, c0: 9, r1: 10, c1: 10 },
        { r0: 9, c0: 12, r1: 10, c1: 12 },
        { r0: 9, c0: 13, r1: 10, c1: 13 },
      ]),
    )
  })

  test('bold, fill, currency format, borders', async () => {
    const v = await open()
    await v.typeCell('J10', '12.5')
    await v.clickCell('J10')
    await v.page.keyboard.press('Control+b')
    await v.btn('Format as currency').click()
    await v.btn('Fill color').click()
    await v.page.locator('.de-tb-pop .de-tb-grid button[title="#ffff00"]').click()
    await v.page.waitForTimeout(300)
    await v.clickCell('N20')
    const c = (await cellBox(v, 'J10'))!
    expect(c.text).toBe('$12.50')
    expect(Number(c.weight)).toBeGreaterThanOrEqual(700)
    expect(c.style).toMatch(/background: (rgb\(255, 255, 0\)|#ffff00)/)

    await v.selectRange('J12', 'K13')
    await v.btn('Borders').click()
    await v.page.locator('button[aria-label="All borders"]').click()
    await v.page.waitForTimeout(300)
    expect((await cellBox(v, 'K13'))?.style ?? '').toMatch(/border/)
    expect((await saved(v)).cell('J10')).toBe('$12.50')
  })

  test('insert / delete rows and columns, hide / unhide', async () => {
    const v = await open()
    await v.typeCell('J10', 'here')
    await v.clickCell('J10')
    await v.pick('Insert, delete or hide rows & columns', 'Insert row above')
    await v.page.waitForTimeout(300)
    expect(await v.shown('J11')).toBe('here')
    await v.clickCell('J11')
    await v.pick('Insert, delete or hide rows & columns', 'Insert column left')
    await v.page.waitForTimeout(300)
    expect(await v.shown('K11')).toBe('here')
    await v.clickCell('K11')
    await v.pick('Insert, delete or hide rows & columns', 'Delete column(s)')
    await v.page.waitForTimeout(300)
    expect(await v.shown('J11')).toBe('')

    await v.clickCell('C5')
    await v.pick('Insert, delete or hide rows & columns', 'Hide column(s)')
    await v.page.waitForTimeout(300)
    expect(await v.page.locator('.de-head-c').filter({ hasText: /^C$/ }).count()).toBe(0)
    await v.selectRange('B5', 'D5')
    await v.pick('Insert, delete or hide rows & columns', 'Unhide columns')
    await v.page.waitForTimeout(300)
    expect(await v.page.locator('.de-head-c').filter({ hasText: /^C$/ }).count()).toBe(1)
  })

  test('sheets: add, rename, switch', async () => {
    const v = await open()
    await v.page.locator('.de-tab-add').click()
    await v.page.waitForTimeout(400)
    const tab = v.page.locator('.de-tab[role=tab][aria-selected=true]')
    expect(await tab.innerText()).toMatch(/^Sheet\d+$/)
    await tab.dblclick()
    await v.page.locator('.de-tab-input').fill('Budget')
    await v.page.keyboard.press('Enter')
    await v.page.waitForTimeout(300)
    expect(await v.page.locator('.de-tab[role=tab]').allInnerTexts()).toEqual(['Sales', 'Budget', 'Notes', 'Data'])
    await v.typeCell('A1', 'new sheet')
    const s = await saved(v, 1)
    expect(s.names).toEqual(['Sales', 'Budget', 'Notes', 'Data'])
    expect(s.cell('A1')).toBe('new sheet')
  })

  test('freeze panes and gridlines', async () => {
    const v = await open()
    await v.pick('View: freeze, gridlines', 'Unfreeze')
    await v.pick('View: freeze, gridlines', 'Hide gridlines')
    await v.page.waitForTimeout(300)
    expect(await v.page.locator('.de-gl-h').count()).toBe(0)
    let s = await saved(v)
    expect(s.frozen ?? { rows: 0, cols: 0 }).toEqual({ rows: 0, cols: 0 })
    expect(s.gridLines).toBe(false)
    await v.clickCell('C4')
    await v.pick('View: freeze, gridlines', 'Freeze up to current cell')
    await v.page.waitForTimeout(300)
    s = await saved(v)
    expect(s.frozen).toEqual({ rows: 3, cols: 2 })
  })

  test('sort a range', async () => {
    const v = await open()
    await v.typeCell('J10', '3')
    await v.typeCell('J11', '1')
    await v.typeCell('J12', '2')
    await v.selectRange('J10', 'J12')
    await v.pick('Sort range', 'Sort range A → Z')
    await v.page.waitForTimeout(300)
    expect([await v.shown('J10'), await v.shown('J11'), await v.shown('J12')]).toEqual(['1', '2', '3'])
  })

  test('functions menu (AutoSum)', async () => {
    const v = await open()
    await v.typeCell('J10', '4')
    await v.typeCell('J11', '6')
    await v.clickCell('J12')
    await v.pick('Functions', 'SUM')
    await v.page.waitForTimeout(300)
    expect(await v.formula('J12')).toBe('=SUM(J10:J11)')
    expect(await v.shown('J12')).toBe('10')
  })

  test('find & replace', async () => {
    const v = await open()
    await v.clickCell('A3')
    await v.page.keyboard.press('Control+f')
    await v.page.locator('.de-find input[aria-label="Find"]').fill('North')
    await v.page.waitForTimeout(400)
    expect(await v.page.locator('.de-find-count').innerText()).toMatch(/^1 \/ \d+$/)
    await v.btn('Replace').click()
    await v.page.locator('.de-find input[aria-label="Replace with"]').fill('Nord')
    await v.page.locator('.de-find button', { hasText: 'All' }).click()
    await v.page.waitForTimeout(400)
    expect(await v.shown('A3')).toBe('Nord')
  })

  test('undo / redo', async () => {
    const v = await open()
    await v.typeCell('J10', 'first')
    await v.typeCell('J10', 'second')
    await v.btn('Undo').click()
    await v.page.waitForTimeout(300)
    expect(await v.shown('J10')).toBe('first')
    await v.btn('Redo').click()
    await v.page.waitForTimeout(300)
    expect(await v.shown('J10')).toBe('second')
    // keyboard
    await v.typeCell('J10', 'third')
    await v.clickCell('N20')
    await v.page.keyboard.press('Control+z')
    await v.page.waitForTimeout(300)
    expect(await v.shown('J10')).toBe('second')
    await v.page.keyboard.press('Control+y')
    await v.page.waitForTimeout(300)
    expect(await v.shown('J10')).toBe('third')
  })
})

describe('xlsx select all', () => {
  test('selects the whole sheet', async () => {
    const v = await open()
    await v.clickCell('B2')
    await v.page.keyboard.press('ControlOrMeta+a')
    await v.page.waitForFunction(() => document.querySelector('.de-fx-ref')?.textContent === 'A1:H62')
    expect(await v.page.locator('.de-head-on').count()).toBeGreaterThan(8)
    expect(await v.page.evaluate(() => getSelection()!.toString())).not.toContain('sample.xlsx')
    expect(v.errors).toEqual([])
  })
})

describe('xlsx keyboard focus', () => {
  test('the grid has the focus on load; Ctrl/⌘+A from the toolbar selects the sheet, in a cell being edited its text', async () => {
    const v = await open()
    expect(await v.page.evaluate(() => document.activeElement?.classList.contains('de-grid'))).toBe(true)
    await v.barButton('View').focus()
    await v.page.keyboard.press('ControlOrMeta+a')
    await v.page.waitForFunction(() => document.querySelector('.de-fx-ref')?.textContent === 'A1:H62')
    await v.clickCell('B2')
    await v.page.keyboard.type('abc')
    await v.page.keyboard.press('ControlOrMeta+a')
    await v.page.keyboard.type('x')
    await v.page.keyboard.press('Enter')
    await v.page.waitForTimeout(150)
    expect(await v.shown('B2')).toBe('x')
    expect(v.errors).toEqual([])
  })
})
