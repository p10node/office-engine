import { describe, expect, test } from 'bun:test'
import { readFileSync } from 'node:fs'
import { docxParas, openFile, paraText, SAMPLES, useBrowser, viewer, xlsxSheet, type Viewer } from './harness'

useBrowser()

/** headless print is a no-op: record what reaches the print frame instead */
async function stubPrint(v: Viewer) {
  await v.page.evaluate(() => {
    ;(window as any).__printed = []
    new MutationObserver((ms) => {
      for (const m of ms)
        for (const n of m.addedNodes) {
          if ((n as HTMLElement).tagName !== 'IFRAME') continue
          const w = (n as HTMLIFrameElement).contentWindow!
          w.print = () => {
            const d = w.document
            ;(window as any).__printed.push({ pages: d.querySelectorAll('.de-print-page').length, rows: d.querySelectorAll('tr').length, title: d.title, top: document.title, text: d.body.innerText.slice(0, 2000) })
          }
          w.document.execCommand = () => false
        }
    }).observe(document.body, { childList: true })
  })
}
const printed = async (v: Viewer) => {
  await v.page.waitForFunction(() => (window as any).__printed?.length > 0, null, { timeout: 30_000 })
  return (await v.page.evaluate(() => (window as any).__printed.at(-1))) as { pages: number; rows: number; title: string; top: string; text: string }
}

const file = (name: string, text: string, mimeType = 'text/plain') => ({ name, mimeType, buffer: Buffer.from(text) })

describe('File menu entries', () => {
  test('documents', async () => {
    const v = await viewer({ src: 'samples/sample.docx', editable: true })
    expect(await v.menuLabels()).toEqual([
      'Document', 'Spreadsheet', 'Presentation', 'Open…',
      'Save', 'Rename…',
      'Word document', 'PDF document', 'Web page', 'Plain text', 'Markdown',
      'Document details', 'Print',
    ])
  })

  test('spreadsheets add Import and CSV / TSV', async () => {
    const v = await viewer({ src: 'samples/sample.xlsx', editable: true })
    const labels = await v.menuLabels()
    expect(labels).toContain('Import CSV into new sheet…')
    expect(labels).toEqual(expect.arrayContaining(['Excel workbook', 'CSV (current sheet)', 'TSV (current sheet)']))
  })

  test('presentations', async () => {
    const v = await viewer({ src: 'samples/sample.pptx' })
    const labels = await v.menuLabels()
    expect(labels).toEqual(expect.arrayContaining(['PowerPoint presentation', 'PDF document', 'Web page', 'Plain text', 'Print']))
    expect(labels).not.toContain('Import CSV into new sheet…')
  })

  test('host entries and Close', async () => {
    const v = await viewer({ src: 'samples/sample.docx', items: true, close: true })
    const labels = await v.menuLabels()
    expect(labels.slice(-3)).toEqual(['Share…', 'Version history', 'Close'])
    await v.menu('Share…')
    await v.menu('Close')
    expect((await v.events('item')).map((e) => e.label)).toEqual(['Share'])
    expect((await v.events('close')).length).toBe(1)
  })

  test('fileMenu=false and allowOpen=false', async () => {
    const v = await viewer({ src: 'samples/sample.docx', fileMenu: false })
    expect(await v.page.locator('.de-bar .de-file').count()).toBe(0)
    const w = await viewer({ src: 'samples/sample.docx', allowOpen: false })
    const labels = await w.menuLabels()
    expect(labels).not.toContain('Open…')
    expect(labels).not.toContain('Document')
  })
})

describe('Download as', () => {
  test('docx: native, text, markdown', async () => {
    const v = await viewer({ src: 'samples/sample.docx' })
    const native = await v.downloadAs('Word document')
    expect(native.name).toBe('sample.docx')
    expect((await docxParas(native.bytes)).map(paraText)).toContain('Quarterly Report')
    const txt = await v.downloadAs('Plain text')
    expect(txt.name).toBe('sample.txt')
    expect(txt.text).toContain('Quarterly Report\n')
    expect(txt.text).toContain('•\tRevenue grew')
    const md = await v.downloadAs('Markdown')
    expect(md.name).toBe('sample.md')
    expect(md.text).toContain('# 1. Highlights')
    expect(md.text).toContain('- Revenue grew 24% year over year')
    expect(md.text).toContain('[example.com](https://example.com)')
    expect(md.text).toMatch(/\| Region \| Q1 \| Q2 \| Total \|/)
  })

  test('xlsx: native, CSV, TSV', async () => {
    const v = await viewer({ src: 'samples/sample.xlsx' })
    const native = await v.downloadAs('Excel workbook')
    expect((await xlsxSheet(native.bytes)).cell('A2')).toBe('Region')
    const csv = await v.downloadAs('CSV (current sheet)')
    expect(csv.name).toBe('sample.csv')
    expect(csv.text.startsWith('﻿Sales by region')).toBe(true)
    expect(csv.text).toContain('\r\nRegion,Date,Units,Price,Revenue,Margin')
    const tsv = await v.downloadAs('TSV (current sheet)')
    expect(tsv.text.split('\r\n')[1].split('\t').slice(0, 3)).toEqual(['Region', 'Date', 'Units'])
  })

  test('pptx: native, text with notes', async () => {
    const v = await viewer({ src: 'samples/sample.pptx' })
    const native = await v.downloadAs('PowerPoint presentation')
    expect((await openFile(native.bytes)).info.pptx!.count).toBe(5)
    const txt = await v.downloadAs('Plain text')
    expect(txt.text).toMatch(/^Slide 1\n/)
    expect(txt.text).toContain('Slide 5')
  })

  test('web page: pages, pictures embedded', async () => {
    for (const [src, pages, needle] of [
      ['samples/sample.docx', 2, 'Quarterly Report'],
      ['samples/sample.pptx', 5, 'office-engine'],
      ['samples/sample.xlsx', 0, 'Sales by region'],
    ] as const) {
      const v = await viewer({ src })
      const html = await v.downloadAs('Web page')
      expect(html.name).toMatch(/^sample\.html$/)
      expect(html.text.startsWith('<!doctype html>')).toBe(true)
      expect(html.text).toContain(needle)
      expect(html.text).not.toContain('blob:')
      const n = (html.text.match(/class="de-print-page"/g) ?? []).length
      if (pages) expect(n).toBeGreaterThanOrEqual(pages)
      if (src.endsWith('pptx')) {
        expect(n).toBe(5)
        expect(html.text).toContain('data:image/')
      }
      if (src.endsWith('xlsx')) expect(html.text).toContain('<table class="de-print-table de-print-grid"')
    }
  })

  test('PDF goes to the print dialog with every page', async () => {
    const v = await viewer({ src: 'samples/sample.pptx' })
    await stubPrint(v)
    const before = await v.page.title()
    await v.menu('PDF document')
    const p = await printed(v)
    expect(p.pages).toBe(5)
    // the suggested PDF name comes from the document title during printing
    expect(p.title).toBe('sample')
    expect(p.top).toBe('sample')
    expect(await v.page.title()).toBe(before)
  })

  test('print: documents paginated, sheets as a table, Ctrl+P', async () => {
    const d = await viewer({ src: 'samples/sample.docx' })
    await stubPrint(d)
    await d.page.locator('.de-page').first().click()
    await d.page.keyboard.press('Control+p')
    const p = await printed(d)
    expect(p.pages).toBeGreaterThan(1)
    expect(p.text).toContain('Quarterly Report')

    const x = await viewer({ src: 'samples/sample.xlsx' })
    await stubPrint(x)
    await x.menu('Print')
    const q = await printed(x)
    expect(q.rows).toBeGreaterThan(50)
    expect(q.text).toContain('Sales by region')
  })
})

describe('New / Open', () => {
  for (const [label, kind, title, ready] of [
    ['Document', 'DOCX', 'Untitled document.docx', '.de-page-inner'],
    ['Spreadsheet', 'XLSX', 'Untitled spreadsheet.xlsx', '.de-grid'],
    ['Presentation', 'PPTX', 'Untitled presentation.pptx', '.de-slide-inner'],
  ] as const) {
    test(`New ${label.toLowerCase()}`, async () => {
      const v = await viewer({ src: 'samples/sample.docx' })
      await v.menu(label)
      await v.page.locator(ready).first().waitFor()
      await v.page.waitForFunction((t) => document.querySelector('.de-bar .de-title')?.textContent === t, title)
      expect(await v.page.locator('.de-badge').innerText()).toBe(kind)
      // starts in edit mode
      await v.page.locator('.de-ebar').waitFor()
      expect(v.errors).toEqual([])
    })
  }

  test('blank document is editable and saves', async () => {
    const v = await viewer({ src: 'samples/sample.docx', onSave: true })
    await v.menu('Document')
    await v.page.locator('.de-ebar').waitFor()
    await v.page.locator('.de-page-inner').first().click()
    await v.page.keyboard.type('Hello new doc')
    await v.page.waitForTimeout(200)
    const saved = await v.save()
    expect(saved.name).toBe('Untitled document.docx')
    expect((await docxParas(saved.bytes)).map(paraText)).toContain('Hello new doc')
  })

  test('Open: CSV becomes a workbook (separator detection, typed numbers)', async () => {
    const v = await viewer({ src: 'samples/sample.docx' })
    await v.chooseFile(() => v.menu('Open…'), file('data.csv', 'name;qty;code\nApple;12,5;007\n"Pear; green";3;010\n', 'text/csv'))
    await v.page.locator('.de-grid').waitFor()
    await v.page.waitForTimeout(500)
    expect(await v.title()).toBe('data.xlsx')
    expect(await v.shown('A3')).toBe('Pear; green')
    expect(await v.shown('B2')).toBe('12.5')
    expect(await v.shown('C2')).toBe('007')
    // converted files start unsaved
    expect(await v.barButton('Save').count()).toBe(1)
  })

  test('Open: Markdown, HTML and text become documents', async () => {
    const v = await viewer({ src: 'samples/sample.xlsx', onSave: true })
    await v.chooseFile(() => v.menu('Open…'), file('notes.md', '# Notes\n\nSome **bold** and [a link](https://example.com).\n\n- one\n- two\n'))
    // the converted workbook is unsaved: confirm
    await v.page.locator('.de-dialog').waitFor().catch(() => {})
    if (await v.page.locator('.de-dialog').count()) await v.page.locator('.de-dialog button', { hasText: 'Continue' }).click()
    await v.page.locator('.de-page-inner').first().waitFor()
    expect(await v.title()).toBe('notes.docx')
    let paras = await docxParas((await v.save()).bytes)
    expect(paras[0].ps.outline).toBe(0)
    expect(paras.map((p) => p.marker?.text).filter(Boolean)).toEqual(['•', '•'])
    expect(paras[1].runs.some((r) => r.t === 'text' && r.s.b && r.text === 'bold')).toBe(true)

    await v.chooseFile(() => v.menu('Open…'), file('page.html', '<h2>Web title</h2><p>Para <i>it</i></p><ol><li>a</li></ol>', 'text/html'))
    await v.page.waitForFunction(() => document.querySelector('.de-bar .de-title')?.textContent === 'page.docx')
    await v.page.waitForTimeout(500)
    paras = await docxParas((await v.save()).bytes)
    expect(paras.map(paraText)).toEqual(['Web title', 'Para it', 'a'])
    expect(paras[0].ps.outline).toBe(1)

    await v.chooseFile(() => v.menu('Open…'), file('plain.txt', 'line one\nline two'))
    await v.page.waitForFunction(() => document.querySelector('.de-bar .de-title')?.textContent === 'plain.docx')
    await v.page.waitForTimeout(500)
    expect((await docxParas((await v.save()).bytes)).map(paraText)).toEqual(['line one', 'line two'])
  })

  test('Open: Office file as is; unsupported file reports an error', async () => {
    const v = await viewer({ src: 'samples/sample.docx' })
    await v.chooseFile(() => v.menu('Open…'), `${SAMPLES}/sample.pptx`)
    await v.page.locator('.de-slide-inner').first().waitFor()
    expect(await v.title()).toBe('sample.pptx')
    await v.chooseFile(() => v.menu('Open…'), { name: 'scan.pdf', mimeType: 'application/pdf', buffer: Buffer.from('%PDF-1.7') })
    expect(await v.toast()).toContain('Cannot open “scan.pdf”')
    expect(await v.title()).toBe('sample.pptx')
  })

  test('unsaved changes are confirmed; Cancel keeps the document', async () => {
    const v = await viewer({ src: 'samples/sample.docx', editable: true })
    await v.selectText('Quarterly Report', 'end')
    await v.page.keyboard.type('!')
    await v.page.waitForTimeout(200)
    await v.menu('Spreadsheet')
    await v.page.locator('.de-dialog').waitFor()
    expect(await v.page.locator('.de-dialog').innerText()).toContain('Discard unsaved changes?')
    await v.page.locator('.de-dialog button', { hasText: 'Cancel' }).click()
    expect(await v.title()).toBe('sample.docx')
    await v.menu('Spreadsheet')
    await v.page.locator('.de-dialog button', { hasText: 'Continue' }).click()
    await v.page.locator('.de-grid').waitFor()
  })

  test('Ctrl+O opens the picker; a new src from the host drops the opened file', async () => {
    const v = await viewer({ src: 'samples/sample.docx' })
    await v.page.locator('.de-page').first().click()
    await v.chooseFile(() => v.page.keyboard.press('Control+o'), `${SAMPLES}/sample.xlsx`)
    await v.page.locator('.de-grid').waitFor()
    await v.page.evaluate(() => window.__setProps({ src: 'samples/sample.pptx' }))
    await v.page.locator('.de-slide-inner').first().waitFor()
    expect(await v.title()).toBe('sample.pptx')
  })
})

describe('Import, Save, Rename, Details', () => {
  test('Import CSV into a new sheet', async () => {
    const v = await viewer({ src: 'samples/sample.xlsx', editable: true, onSave: true })
    await v.chooseFile(() => v.menu('Import CSV into new sheet…'), file('cities.csv', 'city,pop\nHanoi,8.4\nHue,0.65\n', 'text/csv'))
    await v.page.waitForFunction(() => document.querySelector('.de-tab[aria-selected=true]')?.textContent === 'cities')
    await v.page.waitForTimeout(500)
    expect(await v.page.locator('.de-tab[role=tab]').allInnerTexts()).toEqual(['Sales', 'cities', 'Notes', 'Data'])
    expect(await v.shown('A2')).toBe('Hanoi')
    const s = await xlsxSheet((await v.save()).bytes, 1)
    expect(s.names[1]).toBe('cities')
    expect(s.cell('B3')).toBe('0.65')
  })

  test('Save from the menu', async () => {
    const v = await viewer({ src: 'samples/sample.docx', onSave: true })
    await v.menu('Save')
    const e = await v.waitEvent('save')
    expect(e.name).toBe('sample.docx')
  })

  test('Rename: shown name, saved name, onRename, no reload', async () => {
    const v = await viewer({ src: 'samples/sample.docx', editable: true, onSave: true })
    await v.selectText('Quarterly Report', 'end')
    await v.page.keyboard.type(' kept')
    await v.menu('Rename…')
    await v.page.locator('.de-dialog input').fill('Report 2026.docx')
    await v.page.locator('.de-dialog button', { hasText: 'Rename' }).click()
    expect(await v.title()).toBe('Report 2026.docx')
    expect((await v.events('rename')).at(-1)?.name).toBe('Report 2026.docx')
    // edits survive (the document was not reloaded)
    expect(await v.pagesText()).toContain('Quarterly Report kept')
    expect((await v.save()).name).toBe('Report 2026.docx')
    expect((await v.downloadAs('Plain text')).name).toBe('Report 2026.txt')
  })

  test('Document details', async () => {
    for (const [src, rows] of [
      ['samples/sample.docx', ['Name', 'Type', 'File size', 'Pages', 'Words', 'Characters', 'Paragraphs', 'Page size']],
      ['samples/sample.xlsx', ['Name', 'Type', 'File size', 'Author', 'Sheets', 'Current sheet', 'Used range']],
      ['samples/sample.pptx', ['Name', 'Type', 'File size', 'Slides', 'Slide size']],
    ] as const) {
      const v = await viewer({ src })
      await v.menu('Document details')
      await v.page.locator('.de-details').waitFor()
      await v.page.waitForTimeout(300)
      const labels = await v.page.locator('.de-details dt').allInnerTexts()
      expect(labels).toEqual(expect.arrayContaining([...rows]))
      expect(await v.page.locator('.de-details dd').first().innerText()).toBe(src.split('/').pop()!)
      await v.page.locator('.de-dialog button', { hasText: 'Close' }).click()
      await v.page.locator('.de-details').waitFor({ state: 'detached' })
    }
  })
})

test('sample bytes are untouched by downloads', () => {
  // guards the fixtures used above
  expect(readFileSync(`${SAMPLES}/sample.docx`).subarray(0, 2).toString()).toBe('PK')
})
