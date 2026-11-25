import { describe, expect, test } from 'bun:test'
import { useBrowser, viewer } from './harness'

useBrowser()

describe('mail messages', () => {
  test('eml: header, HTML body with its own picture and table, read-only; a picture from the web is not fetched; the attachment is saved from the File menu', async () => {
    const fetched: string[] = []
    const v = await viewer({ src: 'samples/sample.eml', editable: true }, { setup: (v) => v.page.route('https://**/*', (r) => (fetched.push(r.request().url()), r.abort())) })
    const { page } = v
    expect(await v.title()).toBe('sample.eml')
    expect(await page.locator('.de-badge').innerText()).toBe('EML')
    expect(await v.barButton('Edit').count()).toBe(0)
    expect(await page.locator('.de-ebar').count()).toBe(0)
    const text = await page.locator('.de-page-inner').first().innerText()
    for (const t of ['Quarterly review – agenda', 'Ana Martínez', 'ana@example.org', 'Lee, Mei', 'q3-figures.csv', 'Revenue and margins', 'Bring the updated forecast', 'Pictures from the web are not shown']) expect(text).toContain(t)
    // the logo the message carries is shown, as a picture of the document
    const logo = page.locator('.de-page-inner img').first()
    await logo.waitFor()
    expect(await logo.evaluate((img: HTMLImageElement) => img.complete && img.naturalWidth)).toBe(120)
    // the agenda is a table, its header row shaded
    expect(await page.locator('.de-page-inner table').count()).toBeGreaterThanOrEqual(2)
    expect(fetched).toEqual([])

    expect(await v.menuLabels()).toContain('q3-figures.csv')
    const csv = await v.downloadAs('q3-figures.csv')
    expect(csv.name).toBe('q3-figures.csv')
    expect(csv.text).toStartWith('Quarter,Revenue,Margin')
    const docx = await v.downloadAs('Word document')
    expect(docx.name).toBe('sample.docx')
    expect(docx.bytes.subarray(0, 2).toString()).toBe('PK')
    expect(v.errors).toEqual([])
  })

  test('msg: Outlook item with importance, a hidden signature picture shown in the body, an attached message; File → Open takes a .msg', async () => {
    const v = await viewer({ src: 'samples/sample.msg' })
    const { page } = v
    expect(await page.locator('.de-badge').innerText()).toBe('MSG')
    const text = await page.locator('.de-page-inner').first().innerText()
    for (const t of ['Contract renewal – signed copy', 'High importance', 'Jonas Weber', 'Legal', 'contract-signed.pdf', 'signed contract', 'Attached message', 'could you send the signed renewal by Friday?']) expect(text).toContain(t)
    // the signature is in the body, not among the attachments
    expect(await page.locator('.de-page-inner img').count()).toBe(1)
    expect(await v.menuLabels()).not.toContain('signature.png')
    const pdf = await v.downloadAs('contract-signed.pdf')
    expect(pdf.text).toStartWith('%PDF')

    await v.chooseFile(() => v.menu('Open…'), `${import.meta.dir}/../playground/public/samples/sample.eml`)
    await page.waitForFunction(() => (document.querySelector('.de-badge') as HTMLElement | null)?.innerText === 'EML')
    expect(await v.title()).toBe('sample.eml')
    expect(v.errors).toEqual([])
  })

  test('the header is written in the viewer’s language', async () => {
    const v = await viewer({ src: 'samples/sample.msg', texts: 'de', locale: 'de-DE' })
    const text = await v.page.locator('.de-page-inner').first().innerText()
    for (const t of ['Von', 'Gesendet', 'An', 'Anlagen', 'Wichtigkeit: hoch', 'Angehängte Nachricht']) expect(text).toContain(t)
    expect(text).toContain('November 2026')
    expect(v.errors).toEqual([])
  })
})
