import { describe, expect, test } from 'bun:test'
import { parasToDocx } from '../src/core/convert'
import { BlobSource } from '../src/core/source'
import { ZipArchive } from '../src/core/zip'
import { rebuildZip } from '../src/core/zipwrite'
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

  test('pages: rendered read-only, downloads as Word', async () => {
    // even when asked to start in edit mode: a converted document is view only
    const v = await viewer({ src: 'samples/sample.pages', editable: true })
    const { page } = v
    expect(await v.title()).toBe('sample.pages')
    expect(await page.locator('.de-badge').innerText()).toBe('PAGES')
    expect(await v.barButton('Edit').count()).toBe(0)
    expect(await page.locator('.de-ebar').count()).toBe(0)
    const text = await v.pagesText()
    for (const t of ['Quarterly Field Report', 'Northern sites reported first', 'Region', '1480.5', 'Northwind Field Lab']) expect(text).toContain(t)
    expect(await page.locator('.de-pages img').count()).toBeGreaterThan(0)
    expect(await page.locator('.de-pages [contenteditable=true]').count()).toBe(0)

    const dl = await v.downloadAs('Word document')
    expect(dl.name).toBe('sample.docx')
    expect(dl.bytes.subarray(0, 2).toString()).toBe('PK')
    expect(v.errors).toEqual([])
  })

  test('keynote: slides rendered read-only, downloads as PowerPoint', async () => {
    const v = await viewer({ src: 'samples/sample.key', editable: true })
    const { page } = v
    expect(await v.title()).toBe('sample.key')
    expect(await page.locator('.de-badge').innerText()).toBe('KEYNOTE')
    expect(await v.barButton('Edit').count()).toBe(0)
    expect(await page.locator('.de-ebar').count()).toBe(0)
    // the skipped slide is listed but marked hidden, like a hidden PowerPoint slide
    await page.waitForFunction(() => /Slide 1 of 5/.test(document.querySelector('.de-bar')?.textContent ?? ''))
    expect(await page.locator('.de-thumb').count()).toBe(5)
    await page.locator('.de-thumb').nth(2).click()
    await page.waitForFunction(() => /Slide 3 of 5/.test(document.querySelector('.de-bar')?.textContent ?? ''))
    const text = await page.locator('.de-slide-inner').nth(2).innerText()
    for (const t of ['Results', 'Region', '1480.5']) expect(text).toContain(t)
    expect(await page.locator('.de-slide-inner').nth(1).locator('img').count()).toBeGreaterThan(0)

    const dl = await v.downloadAs('PowerPoint presentation')
    expect(dl.name).toBe('sample.pptx')
    expect(dl.bytes.subarray(0, 2).toString()).toBe('PK')
    expect(v.errors).toEqual([])
  })

  test('numbers: tables as sheets, read-only, downloads as Excel', async () => {
    const v = await viewer({ src: 'samples/sample.numbers', editable: true })
    const { page } = v
    expect(await v.title()).toBe('sample.numbers')
    expect(await page.locator('.de-badge').innerText()).toBe('NUMBERS')
    expect(await v.barButton('Edit').count()).toBe(0)
    expect(await page.locator('.de-ebar').count()).toBe(0)
    expect(await page.locator('.de-tab[role=tab]').allInnerTexts()).toEqual(['Budget', 'Notes - Contacts', 'Notes - Totals'])
    expect(await v.shown('A2')).toBe('Rent')
    expect(await v.shown('B3')).toBe('$450.00')
    expect(await v.shown('D3')).toBe('17.6%')
    expect(await v.shown('E3')).toBe('Jan 31, 2025')
    expect(await v.shown('G3')).toBe('3h 30m')
    // typing does not edit the cell
    await v.clickCell('A2')
    await page.keyboard.type('x')
    expect(await v.shown('A2')).toBe('Rent')

    const dl = await v.downloadAs('Excel workbook')
    expect(dl.name).toBe('sample.xlsx')
    expect(dl.bytes.subarray(0, 2).toString()).toBe('PK')
    expect(v.errors).toEqual([])
  })

  test('csv: opens as a workbook, editable, downloads as Excel', async () => {
    const v = await viewer({ src: 'samples/sample.csv' })
    const { page } = v
    expect(await v.title()).toBe('sample.csv')
    expect(await page.locator('.de-badge').innerText()).toBe('CSV')
    expect(await page.locator('.de-tab[role=tab]').allInnerTexts()).toEqual(['sample'])
    expect(await v.shown('A1')).toBe('Region')
    expect(await v.shown('B3')).toBe('Hai Phong, port')
    expect(await v.shown('B4')).toBe('Huế')
    expect(await v.shown('D2')).toBe('1280')
    expect(await v.shown('F3')).toBe('FALSE')
    // unlike Pages / Keynote / Numbers files, editable
    expect(await v.barButton('Edit').count()).toBe(1)
    await v.edit()
    await v.typeCell('A2', 'West')
    expect(await v.shown('A2')).toBe('West')

    const dl = await v.downloadAs('Excel workbook')
    expect(dl.name).toBe('sample.xlsx')
    expect(dl.bytes.subarray(0, 2).toString()).toBe('PK')
    expect(v.errors).toEqual([])
  })

  test('markdown: opens as a document', async () => {
    const v = await viewer({ src: 'samples/sample.md' })
    expect(await v.title()).toBe('sample.md')
    expect(await v.page.locator('.de-badge').innerText()).toBe('MD')
    const text = await v.pagesText()
    for (const t of ['Field notes', 'Checklist', 'Survey the northern sites', 'Export the readings as CSV', 'Numbers are preliminary']) expect(text).toContain(t)
    expect(text).not.toContain('**')
    expect(v.errors).toEqual([])
  })

  test('a Blob without a name: its type, else its content, tells the format', async () => {
    const v = await viewer({ src: 'samples/sample.docx' })
    const { page } = v
    await page.evaluate(() => window.__setProps({ src: new Blob(['city,pop\nHanoi,8.4\nHue,0.65\n'], { type: 'text/csv' }) }))
    await page.locator('.de-grid').waitFor()
    expect(await v.shown('A2')).toBe('Hanoi')
    await page.evaluate(() => window.__setProps({ src: new Blob(['<h1>Hello</h1><p>from <b>HTML</b></p>']) }))
    await page.waitForFunction(() => document.querySelector('.de-badge')?.textContent === 'html')
    await page.locator('.de-page-inner').first().waitFor()
    expect(await v.pagesText()).toContain('from HTML')
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

  test('links: a script or data link is shown as text, not followed', async () => {
    // the builder refuses script links: write safe placeholders, then put the bad targets in the relationships
    const docx = await parasToDocx([
      {
        runs: [
          { text: 'web', href: 'https://example.org/' },
          { text: ' ' },
          { text: 'js', href: 'https://js.invalid/' },
          { text: ' ' },
          { text: 'tab', href: 'https://tab.invalid/' },
          { text: ' ' },
          { text: 'data', href: 'https://data.invalid/' },
        ],
      },
    ])
    const zip = await ZipArchive.open(new BlobSource(docx))
    const rels = (await zip.text('word/_rels/document.xml.rels'))
      .replace('https://js.invalid/', 'javascript:alert(1)')
      .replace('https://tab.invalid/', 'java&#9;script:alert(2)')
      .replace('https://data.invalid/', 'data:text/html,&lt;script&gt;alert(3)&lt;/script&gt;')
    const bytes = new Uint8Array(await (await rebuildZip(zip, new Map([['word/_rels/document.xml.rels', rels]]))).arrayBuffer())

    const v = await viewer({ src: 'samples/sample.docx' })
    const { page } = v
    await page.evaluate((b) => window.__setProps({ src: new Uint8Array(b), name: 'links.docx' }), [...bytes])
    await page.waitForFunction(() => document.querySelector('.de-pages')?.textContent?.includes('web js tab data'))
    const links = await page.evaluate(() => [...document.querySelectorAll('.de-pages a')].map((a) => [a.textContent, a.getAttribute('href')]))
    expect(links).toEqual([['web', 'https://example.org/']])
    const md = await v.downloadAs('Markdown')
    expect(md.text).toContain('[web](https://example.org/) js tab data')
    expect(v.errors).toEqual([])
  })

  test('new src from the host replaces the document', async () => {
    const v = await viewer({ src: 'samples/sample.docx' })
    await v.page.evaluate(() => window.__setProps({ src: 'samples/sample.xlsx' }))
    await v.page.locator('.de-grid').waitFor()
    expect(await v.title()).toBe('sample.xlsx')
  })
})
