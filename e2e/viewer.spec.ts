import { describe, expect, test } from 'bun:test'
import { parasToDocx } from '../src/core/convert'
import { BlobSource } from '../src/core/source'
import { ZipArchive } from '../src/core/zip'
import { rebuildZip } from '../src/core/zipwrite'
import { docxParas, useBrowser, viewer } from './harness'

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
    // the ```mermaid block is a drawing, not code
    expect(text).not.toContain('flowchart LR')
    const img = v.page.locator('.de-page-inner img')
    await img.first().waitFor()
    expect(await img.count()).toBe(1)
    expect(await img.first().evaluate((el: HTMLImageElement) => el.complete && el.naturalWidth > 100)).toBe(true)
    expect(v.errors).toEqual([])
  })

  test('diagrams: Mermaid, PlantUML and Graphviz files open as drawings', async () => {
    const files: [string, string, string, number][] = [
      ['flowchart.mmd', 'MMD', 'Mermaid diagram', 1],
      ['sequence.puml', 'PUML', 'PlantUML diagram', 1],
      ['graph.dot', 'DOT', 'Graphviz diagram', 1],
    ]
    for (const [name, badge, , count] of files) {
      const v = await viewer({ src: `samples/${name}` })
      const { page } = v
      expect(await v.title()).toBe(name)
      expect(await page.locator('.de-badge').innerText()).toBe(badge)
      const img = page.locator('.de-page-inner img')
      await img.first().waitFor()
      expect(await img.count()).toBe(count)
      await page.waitForFunction(() => {
        const el = document.querySelector('.de-page-inner img') as HTMLImageElement | null
        return !!el && el.complete && el.naturalWidth > 0
      })
      const box = await img.first().boundingBox()
      expect(box!.width).toBeGreaterThan(150)
      expect(box!.height).toBeGreaterThan(100)
      // the picture is a self-contained SVG
      const src = await img.first().getAttribute('src')
      const svg = await page.evaluate(async (u) => (await fetch(u!)).text(), src)
      expect(svg.startsWith('<svg')).toBe(true)
      expect(svg).not.toContain('<script')
      if (process.env.DIAGRAM_SHOTS) await page.screenshot({ path: `${process.env.DIAGRAM_SHOTS}/${name}.png` })
      const dl = await v.downloadAs('Word document')
      expect(dl.name).toBe(name.replace(/\.\w+$/, '.docx'))
      expect(dl.bytes.subarray(0, 2).toString()).toBe('PK')
      expect(v.errors).toEqual([])
    }
  })

  test('source side by side: Markdown edited, the document follows', async () => {
    const v = await viewer({ src: 'samples/sample.md', onSave: true })
    const { page } = v
    await page.locator('.de-page-inner').first().waitFor()
    await page.getByTestId('de-code-toggle').click()
    const input = page.getByTestId('de-source-input')
    await input.waitFor()
    expect(await input.inputValue()).toContain('# Field notes')
    // the document is a preview while its source is shown
    expect(await v.barButton('Edit').count()).toBe(0)
    if (process.env.DIAGRAM_SHOTS) await page.screenshot({ path: `${process.env.DIAGRAM_SHOTS}/split-md.png` })
    await input.click()
    await page.keyboard.press(process.platform === 'darwin' ? 'Meta+ArrowUp' : 'Control+Home')
    await page.keyboard.type('# Edited title\n\n')
    await page.waitForFunction(() => document.querySelector('.de-page-inner')?.textContent?.includes('Edited title'))
    // the source keeps what was typed; the bar keeps the name
    expect(await input.inputValue()).toMatch(/^# Edited title\n\n# Field notes/)
    expect(await v.title()).toBe('sample.md')
    expect(await page.locator('.de-badge').innerText()).toBe('MD')
    const saved = await v.save()
    expect(saved.name).toBe('sample.docx')
    expect((await docxParas(saved.bytes)).map((p) => p.runs.map((r) => (r.t === 'text' ? r.text : '')).join(''))[0]).toBe('Edited title')
    // closing the source: the document can be edited again
    await page.getByTestId('de-code-toggle').click()
    await input.waitFor({ state: 'detached' })
    expect(await v.barButton('Edit').count()).toBe(1)
    expect(v.errors).toEqual([])
  })

  test('source side by side: a diagram redrawn as its code changes; stacked when narrow', async () => {
    const v = await viewer({ src: 'samples/flowchart.mmd' }, { viewport: { width: 600, height: 900 } })
    const { page } = v
    await page.locator('.de-page-inner img').first().waitFor()
    await page.getByTestId('de-code-toggle').click()
    const input = page.getByTestId('de-source-input')
    await input.waitFor()
    // narrow: the source above the drawing
    const a = (await page.getByTestId('de-source').boundingBox())!
    const b = (await page.locator('.de-main').boundingBox())!
    expect(a.y + a.height).toBeLessThanOrEqual(b.y + 1)
    const svgText = () => page.evaluate(async () => {
      const img = document.querySelector('.de-page-inner img') as HTMLImageElement
      return img ? (await fetch(img.src)).text() : ''
    })
    expect(await svgText()).toContain('>Take payment<')
    const before = await page.locator('.de-page-inner img').first().getAttribute('src')
    await input.fill((await input.inputValue()).replace('Take payment', 'Charge the card'))
    // a new picture, drawn from the new code
    await page.waitForFunction((was) => {
      const img = document.querySelector('.de-page-inner img') as HTMLImageElement | null
      return !!img && img.src !== was && img.complete && img.naturalWidth > 0
    }, before)
    expect(await svgText()).toContain('>Charge the card<')
    expect(await page.locator('.de-badge').innerText()).toBe('MMD')
    if (process.env.DIAGRAM_SHOTS) await page.screenshot({ path: `${process.env.DIAGRAM_SHOTS}/split-mmd.png` })
    // a broken diagram: why, and the source, in the document
    await input.fill('flowchart TD\n  A --> B\ngitGraph')
    await page.waitForFunction(() => document.querySelector('.de-page-inner img') !== null)
    await input.fill('gitGraph\n  commit')
    await page.waitForFunction(() => document.querySelector('.de-page-inner')?.textContent?.includes('not drawn'))
    expect(v.errors).toEqual([])
  })

  test('a diagram without a name: told by its content', async () => {
    const v = await viewer({ src: 'samples/sample.docx' })
    const { page } = v
    await page.evaluate(() => window.__setProps({ src: new Blob(['sequenceDiagram\n  Alice->>Bob: Hello\n  Bob-->>Alice: Hi'], { type: 'text/plain' }) }))
    await page.waitForFunction(() => document.querySelector('.de-badge')?.textContent === 'mmd')
    await page.locator('.de-page-inner img').first().waitFor()
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
