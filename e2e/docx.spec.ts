import { describe, expect, test } from 'bun:test'
import { readFileSync } from 'node:fs'
import { BlobSource } from '../src/core/source'
import type { Paragraph } from '../src/core/docx/types'
import { ZipArchive } from '../src/core/zip'
import { docxBlocks, docxParas, paraText, PNG, runsOf, SAMPLES, useBrowser, viewer } from './harness'

useBrowser()

const open = async () => {
  const v = await viewer({ src: 'samples/sample.docx', editable: true, onSave: true })
  await v.page.locator('.de-ebar').waitFor()
  return v
}
const INTRO = 'This document shows '

describe('docx editing', () => {
  test('typing, dirty state and save', async () => {
    const v = await open()
    await v.selectText('Quarterly Report', 'end')
    await v.page.keyboard.type(' 2026')
    await v.page.waitForTimeout(200)
    expect(await v.pagesText()).toContain('Quarterly Report 2026')
    expect((await v.events('dirty')).at(-1)?.dirty).toBe(true)
    expect(await v.barButton('Save').getAttribute('class')).toContain('de-save-dirty')

    const saved = await v.save()
    expect(saved.name).toBe('sample.docx')
    const paras = await docxParas(saved.bytes)
    expect(paras.some((p) => paraText(p) === 'Quarterly Report 2026')).toBe(true)
    expect((await v.events('dirty')).at(-1)?.dirty).toBe(false)
    expect(v.errors).toEqual([])
  })

  test('Enter splits a paragraph, Backspace joins it back', async () => {
    const v = await open()
    await v.selectText('Quarterly', 'end')
    await v.page.keyboard.press('Enter')
    await v.page.waitForTimeout(200)
    let paras = await docxParas((await v.save()).bytes)
    expect(paras.some((p) => paraText(p) === 'Quarterly')).toBe(true)
    // saving moved the focus to the Save button: caret back to the start of the second half
    await v.selectText(' Report', 'start')
    await v.page.keyboard.press('Backspace')
    await v.page.waitForTimeout(200)
    paras = await docxParas((await v.save()).bytes)
    expect(paras.some((p) => paraText(p) === 'Quarterly Report')).toBe(true)
  })

  test('bold, italic, underline, strike from the toolbar and shortcuts', async () => {
    const v = await open()
    await v.selectText('document')
    await v.btn('Bold').click()
    await v.selectText('shows')
    await v.page.keyboard.press('Control+i')
    await v.selectText('Visit')
    await v.btn('Underline').click()
    await v.selectText('This')
    await v.btn('Strikethrough').click()
    await v.page.waitForTimeout(200)
    const p = (await docxParas((await v.save()).bytes)).find((x) => paraText(x).startsWith('This document'))!
    const look = (t: string) => runsOf(p).find((r) => r.text.includes(t))!.s
    expect(look('document').b).toBe(true)
    expect(look('shows').i).toBe(true)
    expect(look('Visit').u).toBe('single')
    expect(look('This').strike).toBe(true)
  })

  test('undo / redo', async () => {
    const v = await open()
    await v.selectText('Quarterly Report', 'end')
    await v.page.keyboard.type('!!')
    await v.page.waitForTimeout(150)
    expect(await v.pagesText()).toContain('Quarterly Report!!')
    await v.btn('Undo').click()
    await v.page.waitForTimeout(150)
    expect(await v.pagesText()).not.toContain('Quarterly Report!!')
    await v.btn('Redo').click()
    await v.page.waitForTimeout(150)
    expect(await v.pagesText()).toContain('Quarterly Report!!')
  })

  test('paragraph style, lists, alignment', async () => {
    const v = await open()
    await v.selectText(INTRO, 'start')
    await v.pick('Paragraph style', 'Heading 2')
    await v.page.waitForTimeout(200)
    let p = (await docxParas((await v.save()).bytes)).find((x) => paraText(x).startsWith('This document'))!
    expect(p.ps.outline).toBe(1)

    await v.selectText(INTRO, 'start')
    await v.btn('Bulleted list').click()
    await v.btn('Center').click()
    await v.page.waitForTimeout(200)
    p = (await docxParas((await v.save()).bytes)).find((x) => paraText(x).startsWith('This document'))!
    expect(p.marker?.text).toBe('•')
    expect(p.ps.align).toBe('center')

    await v.btn('Numbered list').click()
    await v.page.waitForTimeout(200)
    p = (await docxParas((await v.save()).bytes)).find((x) => paraText(x).startsWith('This document'))!
    expect(p.marker?.text).toMatch(/^1[.)]$/)
  })

  test('text and highlight colors', async () => {
    const v = await open()
    await v.selectText('document')
    await v.btn('Text color').click()
    await v.page.locator('.de-tb-pop .de-tb-grid button[title="#ff0000"]').click()
    await v.selectText('shows')
    await v.btn('Highlight color').click()
    await v.page.locator('.de-tb-pop .de-tb-grid button[title="#ffff00"]').click()
    await v.page.waitForTimeout(200)
    const p = (await docxParas((await v.save()).bytes)).find((x) => paraText(x).startsWith('This document'))!
    expect(runsOf(p).find((r) => r.text.includes('document'))!.s.color).toBe('#ff0000')
    expect(runsOf(p).find((r) => r.text.includes('shows'))!.s.hl).toBe('#ffff00')
  })

  test('font and size', async () => {
    const v = await open()
    const orig = runsOf((await docxParas(readFileSync(`${SAMPLES}/sample.docx`))).find((x) => paraText(x).startsWith('This document'))!).find((x) => x.text.includes('document'))!.s.sz ?? 11
    await v.selectText('document')
    await v.pick('Font', 'Georgia')
    await v.selectText('document')
    await v.btn('Increase font size').click()
    await v.page.waitForTimeout(200)
    const p = (await docxParas((await v.save()).bytes)).find((x) => paraText(x).startsWith('This document'))!
    const r = runsOf(p).find((x) => x.text.includes('document'))!
    expect(r.s.font).toContain('Georgia')
    expect(r.s.sz).toBe(Math.floor(orig) + 1)
  })

  test('insert link (dialog)', async () => {
    const v = await open()
    await v.selectText('document')
    await v.page.keyboard.press('Control+k')
    await v.page.locator('.de-dialog input').fill('https://example.org/docs')
    await v.page.keyboard.press('Enter')
    await v.page.waitForTimeout(200)
    const p = (await docxParas((await v.save()).bytes)).find((x) => paraText(x).startsWith('This document'))!
    expect(runsOf(p).find((r) => r.text.includes('document'))!.href).toBe('https://example.org/docs')
  })

  test('insert table', async () => {
    const v = await open()
    const before = (await docxBlocks((await v.save()).bytes)).filter((b) => b.t === 'tbl').length
    await v.selectText('Quarterly Report', 'end')
    await v.btn('Insert table').click()
    // 3 columns × 2 rows: 8 cells per picker row
    await v.page.locator('.de-gridpick button').nth(8 + 2).click()
    await v.page.waitForTimeout(300)
    const tables = (await docxBlocks((await v.save()).bytes)).filter((b) => b.t === 'tbl')
    expect(tables.length).toBe(before + 1)
    const added = tables.find((t) => t.t === 'tbl' && t.rows.length === 2 && t.rows[0].cells.length === 3)
    expect(added).toBeTruthy()
  })

  test('insert image from a file', async () => {
    const v = await open()
    const imgs = await v.page.locator('.de-pages img').count()
    await v.selectText('Quarterly Report', 'end')
    await v.chooseFile(() => v.btn('Insert image').click(), { name: 'dot.png', mimeType: 'image/png', buffer: PNG })
    await v.page.waitForFunction((n) => document.querySelectorAll('.de-pages img').length > n, imgs)
    const p = (await docxParas((await v.save()).bytes)).find((x) => paraText(x).startsWith('Quarterly Report'))!
    expect(p.runs.some((r) => r.t === 'img')).toBe(true)
  })

  test('find & replace', async () => {
    const v = await open()
    await v.btn('Find & replace').click()
    await v.page.locator('.de-find input[aria-label="Find"]').fill('region')
    await v.page.waitForTimeout(300)
    expect(await v.page.locator('.de-find-count').innerText()).toMatch(/^1 \/ \d+$/)
    await v.btn('Replace').click()
    await v.page.locator('.de-find input[aria-label="Replace with"]').fill('area')
    await v.page.locator('.de-find button', { hasText: 'All' }).click()
    await v.page.waitForTimeout(300)
    const text = (await docxParas((await v.save()).bytes)).map(paraText).join('\n')
    expect(text).not.toMatch(/\bregion\b/)
    expect(text).toContain('area')
  })

  test('context menu offers clipboard and insert actions', async () => {
    const v = await open()
    await v.selectText('document')
    const box = (await v.page.getByText('Quarterly Report').first().boundingBox())!
    await v.page.mouse.click(box.x + 10, box.y + box.height / 2, { button: 'right' })
    const labels = await v.page.locator('.de-ctx .de-ctx-label').allInnerTexts()
    expect(labels).toEqual(expect.arrayContaining(['Cut', 'Copy', 'Paste']))
    await v.page.keyboard.press('Escape')
  })

  test('update style to match: the style takes the look of the selection, its other paragraphs follow, undo takes it back', async () => {
    const v = await open()
    const look = (text: string) =>
      v.page.evaluate((x) => {
        const el = [...document.querySelectorAll('.de-pages *')].find((e) => e.childElementCount === 0 && (e.textContent ?? '').includes(x))
        return el ? `${getComputedStyle(el).fontStyle} ${getComputedStyle(el).textDecorationLine}` : 'not found'
      }, text)
    const rightClickSelection = async () => {
      const r = await v.page.evaluate(() => {
        const b = getSelection()!.getRangeAt(0).getBoundingClientRect()
        return { x: b.x + b.width / 2, y: b.y + b.height / 2 }
      })
      await v.page.mouse.click(r.x, r.y, { button: 'right' })
      await v.page.locator('.de-ctx').first().waitFor()
    }
    expect(await look('2. Next steps')).toBe('normal none')
    await v.selectText('1. Highlights')
    await v.btn('Italic').click()
    await v.btn('Underline').click()
    await v.page.waitForTimeout(200)
    // formatting the text alone leaves the other headings as they are
    expect(await look('2. Next steps')).toBe('normal none')

    // keyboard: up from the top wraps to the end of the menu; → opens the styles, Esc closes them, then the menu
    await rightClickSelection()
    for (let i = 0; i < 3; i++) await v.page.keyboard.press('ArrowUp')
    expect(await v.page.locator('.de-ctx > .de-ctx-row > .de-ctx-hi .de-ctx-label').innerText()).toBe('Update style to match')
    await v.page.keyboard.press('ArrowRight')
    expect(await v.page.locator('.de-ctx-sub .de-ctx-hi .de-ctx-label').innerText()).toBe('Normal text')
    expect(await v.page.locator('.de-ctx-sub .de-ctx-row', { hasText: 'current' }).locator('.de-ctx-label').innerText()).toBe('Heading 1')
    await v.page.keyboard.press('Escape')
    expect(await v.page.locator('.de-ctx-sub').count()).toBe(0)
    expect(await v.page.locator('.de-ctx').count()).toBe(1)
    await v.page.keyboard.press('Escape')
    expect(await v.page.locator('.de-ctx').count()).toBe(0)

    // mouse: Update style to match → Heading 1
    await rightClickSelection()
    await v.page.locator('.de-ctx .de-ctx-label', { hasText: 'Update style to match' }).hover()
    await v.page.locator('.de-ctx-sub .de-ctx-label', { hasText: /^Heading 1$/ }).click()
    expect(await v.toast()).toBe('“Heading 1” updated to match the selection')
    await v.page.waitForFunction(() => [...document.querySelectorAll('.de-pages span')].some((e) => e.textContent === '2. Next steps' && getComputedStyle(e).fontStyle === 'italic'))
    expect(await look('2. Next steps')).toBe('italic underline')
    expect(await look('1. Highlights')).toBe('italic underline')
    // other styles keep their look
    expect(await look('Figures')).toBe('normal none')
    expect(await look(INTRO)).toBe('normal none')

    const saved = (await v.save()).bytes
    const paras = await docxParas(saved)
    const heads = paras.filter((p) => p.sid === 'Heading1')
    expect(heads.length).toBeGreaterThan(2)
    for (const p of heads) expect(runsOf(p).every((r) => r.s.i === true && r.s.u === 'single')).toBe(true)
    expect(runsOf(paras.find((p) => paraText(p) === 'Figures')!).some((r) => r.s.i)).toBe(false)
    // the look lives in the style of this file; the paragraph picked no longer carries it itself
    const zip = await ZipArchive.open(new BlobSource(new Blob([saved as Uint8Array<ArrayBuffer>])))
    expect(/<w:style [^>]*w:styleId="Heading1">.*?<\/w:style>/.exec(await zip.text('word/styles.xml'))![0]).toContain('<w:i/><w:iCs/>')
    expect(/<w:p[ >](?:(?!<\/w:p>)[\s\S])*?1\. Highlights[\s\S]*?<\/w:p>/.exec(await zip.text('word/document.xml'))![0]).not.toContain('<w:i/>')

    // again: nothing to update
    await v.selectText('1. Highlights')
    await rightClickSelection()
    await v.page.locator('.de-ctx .de-ctx-label', { hasText: 'Update style to match' }).hover()
    await v.page.locator('.de-ctx-sub .de-ctx-label', { hasText: /^Heading 1$/ }).click()
    await v.page.locator('.de-toast', { hasText: 'already matches' }).waitFor()

    // undo: the style as it was, the text with the formatting it had been given
    await v.btn('Undo').click()
    await v.page.waitForFunction(() => [...document.querySelectorAll('.de-pages span')].some((e) => e.textContent === '2. Next steps' && getComputedStyle(e).fontStyle === 'normal'))
    expect(await look('2. Next steps')).toBe('normal none')
    expect(await look('1. Highlights')).toBe('italic underline')
    expect(v.errors).toEqual([])
  })

  test('typing right after a style change: the caret stays, old and new text get the look of the style', async () => {
    const v = await open()
    const look = (text: string) =>
      v.page.evaluate((x) => {
        const el = [...document.querySelectorAll('.de-pages *')].find((e) => e.childElementCount === 0 && (e.textContent ?? '').includes(x))
        return el ? `${getComputedStyle(el).color} ${getComputedStyle(el).fontWeight}` : 'not found'
      }, text)
    const before = await look(INTRO)
    await v.selectText(INTRO, 'start')
    await v.pick('Paragraph style', 'Heading 2')
    // no pause: this edit reaches the worker before its answer to the style change is back
    await v.page.keyboard.type('XY')
    await v.page.waitForTimeout(600)
    await v.page.keyboard.type('Z ')
    await v.page.waitForTimeout(300)
    expect(await v.pagesText()).toContain(`XYZ ${INTRO}`)
    const heading = await look(INTRO)
    expect(heading).not.toBe(before)
    expect(await look('XYZ')).toBe(heading)
    const p = (await docxParas((await v.save()).bytes)).find((x) => paraText(x).startsWith('XYZ This document'))!
    expect(p.ps.outline).toBe(1)
    expect(v.errors).toEqual([])
  })

  test('outline panel follows new headings', async () => {
    const v = await open()
    await v.btn('Document outline').click()
    const before = await v.page.locator('.de-outline-item').count()
    await v.selectText(INTRO, 'start')
    await v.pick('Paragraph style', 'Heading 1')
    await v.page.waitForTimeout(800)
    expect(await v.page.locator('.de-outline-item').count()).toBe(before + 1)
  })
})

describe('docx select all', () => {
  test('editing: the whole document, replaced by typing across the table, undone in two steps', async () => {
    const v = await open()
    await v.selectText(INTRO, 'end')
    await v.page.keyboard.press('ControlOrMeta+a')
    await v.page.waitForTimeout(150)
    // the rendered pages are highlighted, the viewer around them is not
    const shown = await v.page.evaluate(() => getSelection()!.toString())
    expect(shown).toContain(INTRO)
    expect(shown).not.toContain('Page 1 of')
    await v.page.keyboard.type('Only this')
    await v.page.waitForTimeout(300)
    const text = await v.pagesText()
    expect(text).toContain('Only this')
    expect(text).not.toContain(INTRO)
    const saved = await v.save()
    const blocks = await docxBlocks(saved.bytes)
    expect(blocks.some((b) => b.t === 'tbl')).toBe(false)
    expect((await docxParas(saved.bytes)).map(paraText).filter(Boolean)).toEqual(['Only this'])

    await v.selectText('Only this', 'end')
    await v.page.keyboard.press('Control+z')
    await v.page.keyboard.press('Control+z')
    await v.page.waitForTimeout(300)
    expect(await v.pagesText()).toContain(INTRO)
    expect(await v.page.locator('.de-pages table').count()).toBeGreaterThan(0)
    const undone = await docxBlocks((await v.save()).bytes)
    expect(undone.some((b) => b.t === 'tbl')).toBe(true)
    expect(undone.filter((b): b is Paragraph => b.t === 'p').map(paraText).join('\n')).toContain(INTRO)
    expect(v.errors).toEqual([])
  })

  test('viewing: the document, not the viewer around it', async () => {
    const v = await viewer({ src: 'samples/sample.docx' })
    await v.page.locator('.de-page-inner').first().click({ position: { x: 200, y: 200 } })
    await v.page.keyboard.press('ControlOrMeta+a')
    const text = await v.page.evaluate(() => getSelection()!.toString())
    expect(text).toContain(INTRO)
    expect(text).not.toContain('Page 1 of')
    expect(v.errors).toEqual([])
  })
})

describe('docx keyboard focus', () => {
  test('viewing: the document has the focus on load; Ctrl/⌘+A from the toolbar or the page selects it', async () => {
    const v = await viewer({ src: 'samples/sample.docx' })
    expect(await v.page.evaluate(() => document.activeElement?.className)).toBe('de-scroll')
    await v.page.locator('button[aria-label="Zoom in"]').focus()
    await v.page.keyboard.press('ControlOrMeta+a')
    let text = await v.page.evaluate(() => getSelection()!.toString())
    expect(text).toContain(INTRO)
    expect(text).not.toContain('Page 1 of')
    // from outside the viewer (captureSelectAll, the default)
    await v.page.evaluate(() => {
      getSelection()!.removeAllRanges()
      ;(document.activeElement as HTMLElement).blur()
    })
    await v.page.keyboard.press('ControlOrMeta+a')
    text = await v.page.evaluate(() => getSelection()!.toString())
    expect(text).toContain(INTRO)
    expect(text).not.toContain('Page 1 of')
    expect(v.errors).toEqual([])
  })

  test('captureSelectAll={false}: outside the viewer the page keeps Ctrl/⌘+A, the viewer chrome still selects the document', async () => {
    const v = await viewer({ src: 'samples/sample.docx', captureSelectAll: false })
    await v.page.evaluate(() => (document.activeElement as HTMLElement).blur())
    await v.page.keyboard.press('ControlOrMeta+a')
    expect(await v.page.evaluate(() => getSelection()!.toString())).toContain('Page 1 of')
    await v.page.locator('button[aria-label="Zoom in"]').focus()
    await v.page.keyboard.press('ControlOrMeta+a')
    const text = await v.page.evaluate(() => getSelection()!.toString())
    expect(text).toContain(INTRO)
    expect(text).not.toContain('Page 1 of')
    expect(v.errors).toEqual([])
  })

  test('editing: the caret starts in the document, typing needs no click', async () => {
    const v = await open()
    await v.page.waitForFunction(() => document.activeElement?.classList.contains('de-pages'))
    await v.page.keyboard.type('Hi ')
    await v.page.waitForTimeout(200)
    expect(await v.pagesText()).toContain('Hi ')
    expect((await v.events('dirty')).at(-1)?.dirty).toBe(true)
    expect(v.errors).toEqual([])
  })
})
