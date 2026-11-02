import { describe, expect, test } from 'bun:test'
import { docxParas, openFile, paraText, runsOf, useBrowser, viewer, type Viewer } from './harness'

useBrowser()

/** picks a color through the button's system color picker (the OS dialog is simulated) */
async function pickCustom(v: Viewer, button: string, hex: string) {
  await v.btn(button).click()
  await v.page.locator('.de-tb-pop .de-tb-add').click()
  // "+" closes the popover and opens the OS picker
  await v.page.locator('.de-tb-pop').waitFor({ state: 'detached' })
  await v.page.locator(`button[aria-label^="${button}"] ~ input.de-tb-native`).first().evaluate((el: HTMLInputElement, c) => {
    el.value = c
    el.dispatchEvent(new Event('change'))
  }, hex)
  await v.page.waitForTimeout(300)
}

const customRow = (v: Viewer) => v.page.locator('.de-tb-pop .de-tb-grid').nth(1).locator('button:not(.de-tb-add)')

describe('custom colors', () => {
  test('docx text color: applied, remembered, reused, kept after reload', async () => {
    const v = await viewer({ src: 'samples/sample.docx', editable: true, onSave: true })
    await v.selectText('document')
    await pickCustom(v, 'Text color', '#3a7d44')
    const p = (await docxParas((await v.save()).bytes)).find((x) => paraText(x).startsWith('This document'))!
    expect(runsOf(p).find((r) => r.text.includes('document'))!.s.color).toBe('#3a7d44')

    // the custom row offers it (highlighted for the current selection) on every color button
    await v.selectText('document')
    await v.btn('Text color').click()
    expect(await customRow(v).evaluateAll((b) => b.map((x) => x.getAttribute('title')))).toEqual(['#3a7d44'])
    expect(await customRow(v).first().getAttribute('class')).toBe('on')
    await v.page.keyboard.press('Escape')
    await v.btn('Highlight color').click()
    expect(await customRow(v).count()).toBe(1)

    // one click reuses it
    await v.page.mouse.click(5, 5)
    await v.selectText('shows')
    await v.btn('Text color').click()
    await customRow(v).first().click()
    await v.page.waitForTimeout(200)
    const q = (await docxParas((await v.save()).bytes)).find((x) => paraText(x).startsWith('This document'))!
    expect(runsOf(q).find((r) => r.text.includes('shows'))!.s.color).toBe('#3a7d44')

    // stored per browser
    const w = await v.ctx.newPage()
    await w.goto(v.page.url())
    expect(await w.evaluate(() => localStorage.getItem('office-engine:custom-colors'))).toBe('["#3a7d44"]')
  })

  test('recent list: newest first, no duplicates, at most 7', async () => {
    const v = await viewer({ src: 'samples/sample.docx', editable: true })
    await v.selectText('document')
    for (const c of ['#000001', '#000002', '#000003', '#000004', '#000005', '#000006', '#000007', '#000008', '#000003']) await pickCustom(v, 'Text color', c)
    await v.btn('Text color').click()
    expect(await customRow(v).evaluateAll((b) => b.map((x) => x.getAttribute('title')))).toEqual(['#000003', '#000008', '#000007', '#000006', '#000005', '#000004', '#000002'])
  })

  test('xlsx fill and pptx text use the same picker', async () => {
    const x = await viewer({ src: 'samples/sample.xlsx', editable: true })
    await x.page.locator('.de-ebar').waitFor()
    await x.clickCell('J10')
    await pickCustom(x, 'Fill color', '#123456')
    await x.clickCell('N20')
    const style = await x.page.evaluate(async () => {
      const cells = [...document.querySelectorAll<HTMLElement>('.de-cell')]
      return cells.map((c) => c.getAttribute('style') ?? '').find((s) => /rgb\(18, 52, 86\)|#123456/.test(s)) ?? ''
    })
    expect(style).not.toBe('')

    const p = await viewer({ src: 'samples/sample.pptx', editable: true, onSave: true })
    await p.page.locator('.de-ebar').waitFor()
    await p.page.locator('.de-slide-frame').first().getByText('office-engine').first().click()
    await p.page.locator('.de-txt-edit').waitFor()
    await p.page.keyboard.press('ControlOrMeta+a')
    await pickCustom(p, 'Text color', '#654321')
    await p.page.keyboard.press('Escape')
    await p.page.waitForTimeout(300)
    const { engine } = await openFile((await p.save()).bytes)
    const slide: any = await engine.call('pptx.slide', [0])
    expect(JSON.stringify(slide.els)).toContain('#654321')
  })
})
