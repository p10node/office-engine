import { describe, expect, test } from 'bun:test'
import { useBrowser, viewer } from './harness'

useBrowser()

describe('viewing', () => {
  test('docx: pages, status, outline, zoom', async () => {
    const v = await viewer({ src: 'samples/sample.docx' })
    const { page } = v
    expect(await v.title()).toBe('sample.docx')
    await page.waitForFunction(() => /Page 1 of \d+ · [\d,]+ words/.test(document.querySelector('.de-bar')?.textContent ?? ''))
    expect(await page.locator('.de-page').count()).toBeGreaterThan(0)
    expect(await page.locator('.de-badge').innerText()).toBe('DOCX')

    // outline lists the headings
    await page.locator('.de-outline-toggle').click()
    const items = await page.locator('.de-outline-item').allInnerTexts()
    expect(items).toContain('1. Highlights')
    await page.locator('.de-outline-item', { hasText: 'Table' }).click()

    // zoom in widens the page
    const w0 = (await page.locator('.de-page').first().boundingBox())!.width
    await page.locator('button[aria-label="Zoom in"]').click()
    await page.waitForTimeout(200)
    const w1 = (await page.locator('.de-page').first().boundingBox())!.width
    expect(w1).toBeGreaterThan(w0)
    expect(v.errors).toEqual([])
  })

  test('xlsx: cells, number formats, sheet tabs', async () => {
    const v = await viewer({ src: 'samples/sample.xlsx' })
    expect(await v.shown('B1')).toBe('Sales by region — FY2026')
    expect(await v.shown('A2')).toBe('Region')
    expect(await v.shown('D3')).toMatch(/^\$\d+\.\d\d$/)
    expect(await v.status()).toMatch(/\d+ rows × \d+ cols/)
    const tabs = await v.page.locator('.de-tab[role=tab]').allInnerTexts()
    expect(tabs).toEqual(['Sales', 'Notes', 'Data'])
    await v.page.locator('.de-tab[role=tab]', { hasText: 'Notes' }).click()
    await v.page.waitForTimeout(500)
    expect(await v.page.locator('.de-tab[role=tab][aria-selected=true]').innerText()).toBe('Notes')
    expect(v.errors).toEqual([])
  })

  test('pptx: slides, thumbnails, slideshow', async () => {
    const v = await viewer({ src: 'samples/sample.pptx' })
    const { page } = v
    await page.waitForFunction(() => /Slide 1 of 5/.test(document.querySelector('.de-bar')?.textContent ?? ''))
    expect(await page.locator('.de-thumb').count()).toBe(5)
    await page.locator('.de-thumb').nth(2).click()
    await page.waitForFunction(() => /Slide 3 of 5/.test(document.querySelector('.de-bar')?.textContent ?? ''))

    await page.locator('.de-present-btn').click()
    await page.locator('.de-present').waitFor()
    const n0 = await page.locator('.de-present-n').innerText()
    await page.keyboard.press('ArrowRight')
    await page.waitForTimeout(200)
    expect(await page.locator('.de-present-n').innerText()).not.toBe(n0)
    await page.keyboard.press('Escape')
    await page.locator('.de-present').waitFor({ state: 'detached' })
    expect(v.errors).toEqual([])
  })

  test('thumbnails=false hides the thumbnail column', async () => {
    const v = await viewer({ src: 'samples/sample.pptx', thumbnails: false })
    expect(await v.page.locator('.de-thumb').count()).toBe(0)
  })

  test('toolbar=false hides the bar', async () => {
    const v = await viewer({ src: 'samples/sample.docx', toolbar: false })
    expect(await v.page.locator('.de-bar').count()).toBe(0)
    expect(await v.page.locator('.de-page').count()).toBeGreaterThan(0)
  })

  test('allowEdit=false hides the Edit toggle', async () => {
    const v = await viewer({ src: 'samples/sample.docx', allowEdit: false })
    expect(await v.barButton('Edit').count()).toBe(0)
    const w = await viewer({ src: 'samples/sample.docx' })
    expect(await w.barButton('Edit').count()).toBe(1)
  })

  test('edit / view toggle shows the editor bar', async () => {
    const v = await viewer({ src: 'samples/sample.docx' })
    expect(await v.page.locator('.de-ebar').count()).toBe(0)
    await v.barButton('Edit').click()
    await v.page.locator('.de-ebar').waitFor()
    await v.barButton('View').click()
    await v.page.locator('.de-ebar').waitFor({ state: 'detached' })
  })

  test('broken file: error message and onError', async () => {
    const v = await viewer({ src: 'samples/missing.docx' }, { wait: false })
    await v.page.locator('.de-error').waitFor()
    expect(await v.page.locator('.de-error').innerText()).toContain('Cannot open document')
    const err = await v.waitEvent('error')
    expect(err.message).toBeTruthy()
  })

  test('theme: dark chrome, white pages; auto follows prefers-color-scheme', async () => {
    const v = await viewer({ src: 'samples/sample.docx', theme: 'dark' })
    const { page } = v
    const isDark = () => page.evaluate(() => document.querySelector('.de-root')!.classList.contains('de-dark'))
    const bg = (sel: string) => page.locator(sel).first().evaluate((el) => getComputedStyle(el).backgroundColor)
    expect(await isDark()).toBe(true)
    expect(await bg('.de-bar')).toBe('rgb(35, 38, 42)')
    expect(await bg('.de-page')).toBe('rgb(255, 255, 255)')
    // menus pick up the theme too
    await v.menu()
    expect(await bg('.de-dd-pop')).toBe('rgb(46, 49, 54)')
    await page.keyboard.press('Escape')

    // default: light, whatever the OS prefers
    await page.emulateMedia({ colorScheme: 'dark' })
    await page.evaluate(() => window.__setProps({ theme: undefined }))
    await page.waitForFunction(() => !document.querySelector('.de-root')!.classList.contains('de-dark'))
    expect(await bg('.de-bar')).toBe('rgb(255, 255, 255)')

    // auto: follows prefers-color-scheme, live
    await page.evaluate(() => window.__setProps({ theme: 'auto' }))
    await page.waitForFunction(() => document.querySelector('.de-root')!.classList.contains('de-dark'))
    await page.emulateMedia({ colorScheme: 'light' })
    await page.waitForFunction(() => !document.querySelector('.de-root')!.classList.contains('de-dark'))
    expect(v.errors).toEqual([])
  })

  test('new src from the host replaces the document', async () => {
    const v = await viewer({ src: 'samples/sample.docx' })
    await v.page.evaluate(() => window.__setProps({ src: 'samples/sample.xlsx' }))
    await v.page.locator('.de-grid').waitFor()
    expect(await v.title()).toBe('sample.xlsx')
  })
})
