import { describe, expect, test } from 'bun:test'
import type { Paragraph } from '../src/core/docx/types'
import { docxParas, useBrowser, viewer } from './harness'

useBrowser()

const SHOTS = process.env.E2E_SHOTS
const show = (p: Paragraph) => p.runs.map((r) => (r.t !== 'text' ? '￼' : r.tc ? `[${r.tc.k === 'ins' ? '+' : '-'}${r.text}]` : r.text)).join('')

const open = async (src = 'samples/tracked.docx') => {
  const v = await viewer({ src, editable: true, onSave: true })
  await v.page.locator('.de-ebar').waitFor()
  await v.page.waitForFunction(() => window.__events.some((e) => e.type === 'load'))
  await v.page.waitForTimeout(300)
  return v
}

describe('docx tracked changes', () => {
  test('changes are shown, stepped through, accepted one by one and all at once', async () => {
    const v = await open()
    expect(await v.page.locator('.de-pages [data-tc="del"]').count()).toBeGreaterThanOrEqual(2)
    expect(await v.page.locator('.de-pages [data-tc="ins"]').count()).toBeGreaterThanOrEqual(3)
    const del = v.page.locator('.de-pages [data-tc="del"]').first()
    expect(await del.evaluate((el) => getComputedStyle(el).textDecorationLine)).toContain('line-through')
    expect(await del.getAttribute('title')).toContain('Deleted by Lan Nguyen')
    if (SHOTS) {
      await del.scrollIntoViewIfNeeded()
      await v.page.screenshot({ path: `${SHOTS}/tracked-changes.png` })
    }

    await v.selectText('Quarterly Report', 'start')
    await v.pick('Review changes', 'Next change')
    await v.page.waitForTimeout(150)
    expect(await v.page.evaluate(() => getSelection()!.toString())).toBe('24%')
    await v.pick('Review changes', 'Accept change')
    await v.page.waitForTimeout(250)
    expect(await v.toast()).toContain('1 change accepted')
    let ps = await docxParas((await v.save()).bytes)
    expect(show(ps.find((p) => show(p).startsWith('Revenue'))!)).toBe('Revenue grew [+31%] year over year')

    await v.pick('Review changes', 'Accept all changes')
    await v.page.waitForTimeout(300)
    ps = await docxParas((await v.save()).bytes)
    expect(ps.some((p) => p.runs.some((r) => r.t === 'text' && r.tc))).toBe(false)
    expect(show(ps.find((p) => show(p).startsWith('Expand'))!)).toBe('Expand to Vietnam')
    expect(v.errors).toEqual([])
  })

  test('track my changes: typing is underlined, Backspace strikes through; reject all puts the text back', async () => {
    const v = await open('samples/sample.docx')
    await v.btn('Track changes').click()
    expect(await v.btn('Track changes').getAttribute('class')).toContain('de-tb-on')
    await v.selectText('Hire the platform team', 'end')
    await v.page.keyboard.type(' now')
    await v.page.keyboard.press('Backspace')
    await v.page.keyboard.press('Backspace')
    await v.page.waitForTimeout(200)
    // " now" was mine: Backspace removes it for real (2 + 2 presses); the 5th takes "m" of "team", struck through
    await v.page.keyboard.press('Backspace')
    await v.page.keyboard.press('Backspace')
    await v.page.keyboard.press('Backspace')
    await v.page.waitForTimeout(250)
    let ps = await docxParas((await v.save()).bytes)
    expect(show(ps.find((p) => show(p).startsWith('Hire'))!)).toBe('Hire the platform tea[-m]')
    // typing at the caret (Save took the focus: put it back after "tea") goes in as inserted text, before the deletion
    await v.selectText('platform tea', 'end')
    await v.page.keyboard.type('ms')
    await v.page.waitForTimeout(250)
    ps = await docxParas((await v.save()).bytes)
    expect(show(ps.find((p) => show(p).startsWith('Hire'))!)).toBe('Hire the platform tea[+ms][-m]')
    expect(await v.page.locator('.de-pages [data-tc="ins"]').first().evaluate((el) => getComputedStyle(el).textDecorationLine)).toContain('underline')

    await v.pick('Review changes', 'Reject all changes')
    await v.page.waitForTimeout(300)
    ps = await docxParas((await v.save()).bytes)
    expect(show(ps.find((p) => show(p).startsWith('Hire'))!)).toBe('Hire the platform team')
    expect(v.errors).toEqual([])
  })
})
