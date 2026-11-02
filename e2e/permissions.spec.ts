import { describe, expect, test } from 'bun:test'
import { useBrowser, viewer, type Viewer } from './harness'

useBrowser()

/** whether the last Ctrl+P keydown was claimed by the viewer (the browser does not print the page) */
async function ctrlPBlocked(v: Viewer) {
  await v.page.evaluate(() => window.addEventListener('keydown', (e) => e.key === 'p' && ((window as any).__ctrlP = e.defaultPrevented)))
  await v.page.locator('.de-page').first().click()
  await v.page.keyboard.press('Control+p')
  await v.page.waitForTimeout(200)
  return v.page.evaluate(() => (window as any).__ctrlP as boolean)
}

/** selects the first picture of the document (image tools appear) */
async function selectImage(v: Viewer) {
  const img = v.page.locator('.de-pages img').first()
  await img.scrollIntoViewIfNeeded()
  await img.click()
  await v.page.waitForTimeout(300)
}

describe('allowSave / allowDownload', () => {
  test('defaults: Download button, image download, all File entries', async () => {
    const v = await viewer({ src: 'samples/sample.docx', editable: true })
    expect(await v.barButton('Download').count()).toBe(1)
    const dl = await v.download(() => v.barButton('Download').click())
    expect(dl.name).toBe('sample.docx')
    await selectImage(v)
    expect(await v.btn('Download image').count()).toBe(1)
    expect(await ctrlPBlocked(v)).toBe(true)
  })

  test('allowSave=false: no Save button, no Save entry, Ctrl+S does nothing', async () => {
    const v = await viewer({ src: 'samples/sample.docx', editable: true, allowSave: false, onSave: true })
    await v.selectText('Quarterly Report', 'end')
    await v.page.keyboard.type('!')
    await v.page.waitForTimeout(200)
    expect(await v.page.locator('.de-bar button.de-mode', { hasText: /^(Save|Download|Saving…)$/ }).count()).toBe(0)
    await v.page.keyboard.press('Control+s')
    await v.page.waitForTimeout(500)
    expect(await v.events('save')).toEqual([])
    expect(await v.menuLabels()).not.toContain('Save')
    // downloading a copy is still allowed
    expect(await v.menuLabels()).toContain('Word document')
  })

  test('allowDownload=false: nothing reaches the device', async () => {
    const v = await viewer({ src: 'samples/sample.docx', editable: true, allowDownload: false })
    expect(await v.page.locator('.de-bar button.de-mode', { hasText: /^(Save|Download)$/ }).count()).toBe(0)
    const labels = await v.menuLabels()
    for (const l of ['Word document', 'PDF document', 'Web page', 'Plain text', 'Print', 'Save']) expect(labels).not.toContain(l)
    expect(labels).toContain('Open…')
    await selectImage(v)
    expect(await v.btn('Download image').count()).toBe(0)
    const box = (await v.page.locator('.de-pages img').first().boundingBox())!
    await v.page.mouse.click(box.x + box.width / 2, box.y + box.height / 2, { button: 'right' })
    const ctx = await v.page.locator('.de-ctx .de-ctx-label').allInnerTexts()
    expect(ctx).toContain('Replace image…')
    expect(ctx).not.toContain('Download image')
    await v.page.keyboard.press('Escape')
    // Ctrl+P is claimed so the browser cannot print (save as PDF) the page either
    expect(await ctrlPBlocked(v)).toBe(true)
    const printed = await v.page.evaluate(() => document.querySelectorAll('iframe').length)
    expect(printed).toBe(0)
  })

  test('allowDownload=false with onSave: saving to the host still works', async () => {
    const v = await viewer({ src: 'samples/sample.docx', editable: true, allowDownload: false, onSave: true })
    expect(await v.barButton('Save').count()).toBe(1)
    const saved = await v.save()
    expect(saved.name).toBe('sample.docx')
    expect(await v.menuLabels()).toContain('Save')
  })

  test('both off: no way to export from the UI', async () => {
    const v = await viewer({ src: 'samples/sample.xlsx', editable: true, allowSave: false, allowDownload: false })
    expect(await v.page.locator('.de-bar button.de-mode', { hasText: /^(Save|Download)$/ }).count()).toBe(0)
    const labels = await v.menuLabels()
    expect(labels.filter((l) => /CSV|TSV|PDF|Web page|Excel|Print|Save/.test(l) && !l.startsWith('Import'))).toEqual([])
  })
})
