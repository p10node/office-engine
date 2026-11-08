import { describe, expect, test } from 'bun:test'
import { Engine } from '../src/core/worker/engine'
import type { Block, DocxMeta } from '../src/core/docx/types'
import { useBrowser, viewer } from './harness'

useBrowser()

const SHOTS = process.env.E2E_SHOTS

async function metaOf(bytes: Uint8Array): Promise<DocxMeta> {
  const events: { event: string; data: any }[] = []
  const engine = new Engine((event, data) => events.push({ event, data }))
  await engine.open(new Blob([bytes as BlobPart]), {})
  await engine.call('docx.start', [])
  return events.filter((e) => e.event === 'docx:meta').at(-1)!.data
}
const lineOf = (b: Block | undefined) => (b?.t === 'p' ? b.runs.map((r) => (r.t === 'text' ? r.text : r.t === 'field' ? `{${r.kind}}` : '')).join('') : '')

/** scrolls the document until page `n` is rendered */
async function showPage(v: Awaited<ReturnType<typeof viewer>>, n: number) {
  for (let i = 0; i < 40 && !(await v.page.locator(`.de-page[data-page="${n}"]`).count()); i++) {
    await v.page.locator('.de-scroll').evaluate((el) => el.scrollBy(0, 400))
    await v.page.waitForTimeout(60)
  }
  const p = v.page.locator(`.de-page[data-page="${n}"]`)
  await p.scrollIntoViewIfNeeded()
  return p
}

const open = async () => {
  const v = await viewer({ src: 'samples/sample.docx', editable: true, onSave: true })
  await v.page.locator('.de-ebar').waitFor()
  await v.page.waitForFunction(() => window.__events.some((e) => e.type === 'load'))
  await v.page.waitForTimeout(500)
  return v
}

describe('docx header & footer', () => {
  test('double-click the footer, page numbers "Trang 1/5", shown on every page, saved, undone', async () => {
    const v = await open()
    const page2 = await showPage(v, 2)
    // its bottom margin in view
    await page2.evaluate((el) => el.scrollIntoView({ block: 'end' }))
    await v.page.waitForTimeout(100)
    const box = (await page2.boundingBox())!
    await v.page.mouse.dblclick(box.x + box.width / 2, box.y + box.height - 20)
    const dlg = v.page.locator('form[aria-label="Header and footer"]')
    await dlg.waitFor()
    expect(await dlg.locator('[role=tab][aria-selected=true]').innerText()).toBe('Footer')
    await dlg.locator('select[aria-label="Page number preset"]').selectOption({ label: 'Bottom center: Trang 1/5' })
    await dlg.getByRole('button', { name: 'Apply' }).click()
    await v.page.waitForTimeout(500)

    const total = await v.page.locator('.de-page').count()
    const footer2 = (await page2.locator('.de-hf').last().innerText()).trim()
    expect(footer2).toMatch(/^Trang 2\/\d+$/)
    expect(Number(footer2.split('/')[1])).toBeGreaterThanOrEqual(total)
    if (SHOTS) await v.page.screenshot({ path: `${SHOTS}/footer-page-numbers.png` })

    const m = await metaOf((await v.save()).bytes)
    expect(lineOf(m.footer?.[0])).toBe('Trang {PAGE}/{NUMPAGES}')

    await v.btn('Undo').click()
    await v.page.waitForTimeout(500)
    expect((await page2.locator('.de-hf').last().innerText()).trim()).not.toMatch(/^Trang/)
    expect(v.errors).toEqual([])
  })

  test('header text from the toolbar, different first page', async () => {
    const v = await open()
    await v.btn('Header & footer').click()
    const dlg = v.page.locator('form[aria-label="Header and footer"]')
    await dlg.locator('textarea').fill('CÔNG TY ABC — Báo cáo quý')
    await dlg.getByText('Right', { exact: true }).click()
    await dlg.getByText('Different first page').click()
    await dlg.getByRole('tab', { name: 'First page header' }).click()
    await dlg.locator('textarea').fill('')
    await dlg.getByRole('button', { name: 'Apply' }).click()
    await v.page.waitForTimeout(500)
    const p1 = await v.page.locator('.de-page[data-page="1"] .de-hf').allInnerTexts()
    expect(p1.join(' ')).not.toContain('CÔNG TY ABC')
    const p2 = await showPage(v, 2)
    expect(await p2.locator('.de-hf').first().innerText()).toContain('CÔNG TY ABC — Báo cáo quý')
    const m = await metaOf((await v.save()).bytes)
    expect(m.titlePg).toBe(true)
    expect(lineOf(m.header?.[0])).toBe('CÔNG TY ABC — Báo cáo quý')
    expect(m.header?.[0].t === 'p' && m.header[0].ps.align).toBe('right')
    expect(v.errors).toEqual([])
  })
})
