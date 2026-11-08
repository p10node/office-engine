import { describe, expect, test } from 'bun:test'
import { Engine } from '../src/core/worker/engine'
import type { DocxMeta } from '../src/core/docx/types'
import { useBrowser, viewer, type Viewer } from './harness'

useBrowser()

const SHOTS = process.env.E2E_SHOTS

async function metaOf(bytes: Uint8Array): Promise<DocxMeta> {
  const events: { event: string; data: any }[] = []
  const engine = new Engine((event, data) => events.push({ event, data }))
  await engine.open(new Blob([bytes as BlobPart]), {})
  await engine.call('docx.start', [])
  return events.filter((e) => e.event === 'docx:meta').at(-1)!.data
}

const open = async () => {
  const v = await viewer({ src: 'samples/sample.docx', editable: true, onSave: true })
  await v.page.locator('.de-ebar').waitFor()
  await v.page.waitForFunction(() => window.__events.some((e) => e.type === 'load'))
  return v
}

const pageBox = async (v: Viewer) => (await v.page.locator('.de-page').first().boundingBox())!

describe('docx page setup', () => {
  test('A4 landscape, narrow margins, two columns: laid out, saved, undone', async () => {
    const v = await open()
    const before = await pageBox(v)
    expect(before.height).toBeGreaterThan(before.width)
    await v.btn('Page setup').click()
    const dlg = v.page.locator('form[aria-label="Page setup"]')
    await dlg.waitFor()
    await dlg.locator('select[aria-label="Paper size"]').selectOption('a4')
    await dlg.getByText('Landscape').click()
    await dlg.locator('select[aria-label="Margin preset"]').selectOption({ label: 'Narrow' })
    await dlg.locator('input[name="de-ps-cols"]').nth(1).check()
    await dlg.getByRole('button', { name: 'Apply' }).click()
    await v.page.waitForTimeout(600)

    const after = await pageBox(v)
    expect(after.width / after.height).toBeCloseTo(841.9 / 595.3, 1)
    const colsCss = await v.page.locator('.de-page-cols').first().evaluate((el) => getComputedStyle(el).columnCount)
    expect(colsCss).toBe('2')
    if (SHOTS) await v.page.screenshot({ path: `${SHOTS}/page-setup-2cols.png` })

    const m = await metaOf((await v.save()).bytes)
    expect(m.page.w).toBeCloseTo(841.9, 0)
    expect(m.page.h).toBeCloseTo(595.3, 0)
    expect(m.page.mt).toBeCloseTo(36, 0)
    expect(m.page.ml).toBeCloseTo(36, 0)
    expect(m.page.cols).toBe(2)

    // one undo puts the whole setup back
    await v.btn('Undo').click()
    await v.page.waitForTimeout(600)
    const back = await pageBox(v)
    expect(back.height).toBeGreaterThan(back.width)
    expect(await v.page.locator('.de-page-cols').count()).toBe(0)
    const m2 = await metaOf((await v.save()).bytes)
    expect(m2.page.w).toBeCloseTo(612, 0)
    expect(m2.page.cols).toBeUndefined()
    expect(v.errors).toEqual([])
  })

  test('margins that leave no room are refused in the dialog', async () => {
    const v = await open()
    await v.btn('Page setup').click()
    const dlg = v.page.locator('form[aria-label="Page setup"]')
    await dlg.locator('input[aria-label="Left margin (cm)"]').fill('12')
    await dlg.locator('input[aria-label="Right margin (cm)"]').fill('12')
    expect(await dlg.getByRole('button', { name: 'Apply' }).isDisabled()).toBe(true)
    await v.page.keyboard.press('Escape')
    expect(await dlg.count()).toBe(0)
  })
})
