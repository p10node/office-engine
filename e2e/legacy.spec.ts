import { describe, expect, test } from 'bun:test'
import { BlobSource } from '../src/core/source'
import { ZipArchive } from '../src/core/zip'
import { useBrowser, viewer } from './harness'

useBrowser()

/** the parts of a saved Office file */
const partsOf = async (bytes: Buffer) => (await ZipArchive.open(new BlobSource(new Blob([new Uint8Array(bytes)])))).list().map((e) => e.name)

describe('Word / Excel / PowerPoint 97-2003, RTF and OpenDocument files', () => {
  for (const { file, badge, saved, texts } of [
    { file: 'sample.doc', badge: 'DOC', saved: 'sample.docx', texts: ['Revenue grew 24% year over year', 'Region', 'Q1', 'Prepared by the finance team'] },
    // the table is on page 2 (only the pages in view render)
    { file: 'sample.odt', badge: 'ODT', saved: 'sample.docx', texts: ['Revenue grew 24% year over year', '1. Highlights', 'a)Text documents', 'Quarterly Report — Draft'] },
    { file: 'sample.rtf', badge: 'RTF', saved: 'sample.docx', texts: ['Revenue grew 24% year over year', 'Highlights', 'café €', 'Hire the platform team'] },
  ]) {
    test(`${file}: converted to a document, editable, saved as Word`, async () => {
      const v = await viewer({ src: `samples/${file}`, onSave: true })
      const { page } = v
      expect(await v.title()).toBe(file)
      expect(await page.locator('.de-badge').innerText()).toBe(badge)
      const text = await v.pagesText()
      for (const t of ['Quarterly Report', ...texts]) expect(text).toContain(t)
      // unlike Pages files, editable; saving writes a .docx
      await v.edit()
      await v.selectText('Quarterly Report', 'end')
      await page.keyboard.type(' (edited)')
      expect(await v.pagesText()).toContain('Quarterly Report (edited)')
      const out = await v.save()
      expect(out.name).toBe(saved)
      expect(await partsOf(out.bytes)).toContain('word/document.xml')
      expect(v.errors).toEqual([])
    })
  }

  for (const { file, badge, tabs } of [
    { file: 'sample.xls', badge: 'XLS', tabs: ['Sales', 'Notes', 'Data'] },
    // (its Lookup sheet is hidden)
    { file: 'sample.ods', badge: 'ODS', tabs: ['Sales', 'Notes', 'Data'] },
  ]) {
    test(`${file}: converted to a workbook, formulas calculated, editable, saved as Excel`, async () => {
      const v = await viewer({ src: `samples/${file}`, onSave: true })
      const { page } = v
      expect(await v.title()).toBe(file)
      expect(await page.locator('.de-badge').innerText()).toBe(badge)
      expect(await page.locator('.de-tab[role=tab]').allInnerTexts()).toEqual(tabs)
      expect(await v.shown('A2')).toBe('Region')
      expect(await v.shown('E2')).toBe('Revenue')
      // a formula (Units × Price): shown with its value
      expect(await v.shown('E3')).not.toBe('')
      expect(await v.shown('E3')).not.toMatch(/^#/)
      await v.edit()
      await v.typeCell('A3', 'West')
      expect(await v.shown('A3')).toBe('West')
      const out = await v.save()
      expect(out.name).toBe(file.replace(/\.\w+$/, '.xlsx'))
      expect(await partsOf(out.bytes)).toContain('xl/workbook.xml')
      expect(v.errors).toEqual([])
    })
  }

  test('sample.ppt: converted to a presentation, hidden slide kept, saved as PowerPoint', async () => {
    const v = await viewer({ src: 'samples/sample.ppt', onSave: true })
    const { page } = v
    expect(await v.title()).toBe('sample.ppt')
    expect(await page.locator('.de-badge').innerText()).toBe('PPT')
    await page.waitForFunction(() => /Slide 1 of 5/.test(document.querySelector('.de-bar')?.textContent ?? ''))
    expect(await page.locator('.de-thumb').count()).toBe(5)
    expect(await page.locator('.de-slide-inner').first().innerText()).toContain('Progressive Office viewer for the browser')
    await page.locator('.de-thumb').nth(2).click()
    await page.waitForFunction(() => /Slide 3 of 5/.test(document.querySelector('.de-bar')?.textContent ?? ''))
    const shapes = await page.locator('.de-slide-inner').nth(2).innerText()
    for (const t of ['Rounded', 'Oval', 'colored bold text']) expect(shapes).toContain(t)
    await v.edit()
    const out = await v.save()
    expect(out.name).toBe('sample.pptx')
    expect(await partsOf(out.bytes)).toContain('ppt/presentation.xml')
    expect(v.errors).toEqual([])
  })

  test('sample.odp: converted to a presentation, slides and thumbnails, saved as PowerPoint', async () => {
    const v = await viewer({ src: 'samples/sample.odp', onSave: true })
    const { page } = v
    expect(await v.title()).toBe('sample.odp')
    expect(await page.locator('.de-badge').innerText()).toBe('ODP')
    await page.waitForFunction(() => /Slide 1 of 6/.test(document.querySelector('.de-bar')?.textContent ?? ''))
    expect(await page.locator('.de-thumb').count()).toBe(6)
    expect(await page.locator('.de-slide-inner').first().innerText()).toContain('office-engine')
    await page.locator('.de-thumb').nth(3).click()
    await page.waitForFunction(() => /Slide 4 of 6/.test(document.querySelector('.de-bar')?.textContent ?? ''))
    const table = await page.locator('.de-slide-inner').nth(3).innerText()
    for (const t of ['Metric', '2026', 'Users']) expect(table).toContain(t)
    expect(await v.barButton('Edit').count()).toBe(1)
    await v.edit()
    const out = await v.save()
    expect(out.name).toBe('sample.pptx')
    expect(await partsOf(out.bytes)).toContain('ppt/presentation.xml')
    expect(v.errors).toEqual([])
  })

  test('Document details name the source format', async () => {
    const v = await viewer({ src: 'samples/sample.doc' })
    await v.page.locator('.de-page-inner').first().waitFor()
    await v.menu('Document details')
    const dialog = await v.page.locator('.de-dialog').innerText()
    expect(dialog).toContain('Word 97-2003 document, opened as Word document')
    expect(v.errors).toEqual([])
  })
})
