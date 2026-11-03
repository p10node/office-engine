import { describe, expect, test } from 'bun:test'
import { openFile, PNG, useBrowser, viewer, type Viewer } from './harness'

useBrowser()

const open = async () => {
  const v = await viewer({ src: 'samples/sample.pptx', editable: true, onSave: true })
  await v.page.locator('.de-ebar').waitFor()
  return v
}

/** the saved presentation: slide count, hidden flags and a slide model reader */
async function saved(v: Viewer) {
  const { engine, info } = await openFile((await v.save()).bytes)
  return { meta: info.pptx!, slide: (i: number): Promise<any> => engine.call('pptx.slide', [i]) }
}

const slide1 = (v: Viewer) => v.page.locator('.de-slide-frame').first()
const els = (v: Viewer) => slide1(v).locator('[data-top]')

/** starts text editing in the first slide's text containing `text` and selects all of it */
async function editText(v: Viewer, text: string) {
  await slide1(v).getByText(text).first().click()
  await v.page.locator('.de-txt-edit').waitFor()
  await v.page.keyboard.press('ControlOrMeta+a')
}

const stopEditing = (v: Viewer) => v.page.keyboard.press('Escape')

describe('pptx editing', () => {
  test('edit text in a shape and save', async () => {
    const v = await open()
    await editText(v, 'office-engine')
    await v.page.keyboard.type('Quarterly deck')
    await stopEditing(v)
    await v.page.waitForTimeout(300)
    expect(await slide1(v).innerText()).toContain('Quarterly deck')
    expect((await v.events('dirty')).at(-1)?.dirty).toBe(true)
    const s = await saved(v)
    expect(JSON.stringify((await s.slide(0)).els)).toContain('Quarterly deck')
    expect(v.errors).toEqual([])
  })

  test('text formatting: bold, color, size', async () => {
    const v = await open()
    await editText(v, 'office-engine')
    await v.btn('Bold').click()
    await v.btn('Text color').click()
    await v.page.locator('.de-tb-pop .de-tb-grid button[title="#ff0000"]').click()
    await v.btn('Increase font size').click()
    await stopEditing(v)
    await v.page.waitForTimeout(300)
    const s = await saved(v)
    const runs = (await s.slide(0)).els.flatMap((e: any) => e.text?.paras ?? []).flatMap((p: any) => p.runs)
    const r = runs.find((x: any) => x.text === 'office-engine')
    expect(r.b).toBe(true)
    expect(r.color).toBe('#ff0000')
  })

  test('insert, move, fill and delete a shape; undo', async () => {
    const v = await open()
    const n = await els(v).count()
    await v.pick('Insert', 'Rectangle')
    await v.page.waitForFunction(([sel, k]) => document.querySelectorAll(sel as string).length > (k as number), ['.de-slide-frame:first-child [data-top]', n] as const).catch(() => {})
    expect(await els(v).count()).toBe(n + 1)

    // the new shape is selected: arrow keys nudge it
    const x0 = (await v.page.locator('.de-shape-sel').boundingBox())!.x
    await v.page.keyboard.press('ArrowRight')
    await v.page.keyboard.press('ArrowRight')
    await v.page.waitForTimeout(300)
    expect((await v.page.locator('.de-shape-sel').boundingBox())!.x).toBeGreaterThan(x0)
    await v.btn('Fill color').click()
    await v.page.locator('.de-tb-pop .de-tb-grid button[title="#00ff00"]').click()
    await v.page.waitForTimeout(300)
    const s1 = await saved(v)
    expect(s1.meta.count).toBe(5)
    expect((await s1.slide(0)).els.some((e: any) => e.geom === 'rect' && e.fill?.c === '#00ff00')).toBe(true)

    // Save took the focus: select the shape again (it is the topmost at the slide center)
    const f = (await slide1(v).boundingBox())!
    await v.page.mouse.click(f.x + f.width / 2, f.y + f.height / 2)
    // a click on a shape with a text body edits its text: Esc keeps the shape selected
    if (await v.page.locator('.de-txt-edit').count()) await stopEditing(v)
    await v.btn('Delete').waitFor()
    await v.btn('Delete').click()
    await v.page.waitForTimeout(300)
    expect(await els(v).count()).toBe(n)
    await v.btn('Undo').click()
    await v.page.waitForTimeout(300)
    expect(await els(v).count()).toBe(n + 1)
  })

  test('arrange, rotate, duplicate', async () => {
    const v = await open()
    const n = await els(v).count()
    await v.pick('Insert', 'Ellipse')
    await v.page.waitForTimeout(300)
    await v.pick('Rotate', 'Rotate 90° clockwise')
    await v.pick('Arrange', 'Send to back')
    await v.page.keyboard.press('Control+d')
    await v.page.waitForTimeout(300)
    expect(await els(v).count()).toBe(n + 2)
    const s = await saved(v)
    const ellipses = (await s.slide(0)).els.filter((e: any) => e.geom === 'ellipse')
    expect(ellipses.length).toBeGreaterThanOrEqual(2)
    expect(ellipses.some((e: any) => e.box.rot === 90)).toBe(true)
  })

  test('insert a text box and an image', async () => {
    const v = await open()
    const n = await els(v).count()
    await v.pick('Insert', 'Text box')
    await v.page.waitForTimeout(300)
    await v.page.keyboard.type('Hello box')
    await stopEditing(v)
    await v.page.keyboard.press('Escape')
    await v.chooseFile(() => v.pick('Insert', 'Image…'), { name: 'dot.png', mimeType: 'image/png', buffer: PNG })
    await v.page.waitForTimeout(600)
    expect(await els(v).count()).toBe(n + 2)
    const s = await saved(v)
    const slide = await s.slide(0)
    expect(slide.els.some((e: any) => e.t === 'pic')).toBe(true)
  })

  test('slides: new, duplicate, hide, move, delete', async () => {
    const v = await open()
    const status = () => v.page.locator('.de-bar').innerText()
    await v.pick('Slide', 'New slide')
    await v.page.waitForFunction(() => /of 6/.test(document.querySelector('.de-bar')?.textContent ?? ''))
    await v.pick('Slide', 'Duplicate slide')
    await v.page.waitForFunction(() => /of 7/.test(document.querySelector('.de-bar')?.textContent ?? ''))
    await v.pick('Slide', 'Hide / show slide')
    await v.page.waitForTimeout(300)
    let s = await saved(v)
    expect(s.meta.count).toBe(7)
    // hidden flags live in the slides (the presentation list is read lazily)
    const hidden = await Promise.all(s.meta.slides.map(async (x) => (await s.slide(x.index)).hidden))
    expect(hidden.filter(Boolean).length).toBe(1)
    await v.pick('Slide', 'Delete slide')
    await v.page.waitForFunction(() => /of 6/.test(document.querySelector('.de-bar')?.textContent ?? ''))
    await v.page.locator('.de-thumb').nth(0).click()
    await v.pick('Slide', 'Move slide down')
    await v.page.waitForTimeout(300)
    s = await saved(v)
    expect(s.meta.count).toBe(6)
    // the title slide moved to position 2
    expect(JSON.stringify((await s.slide(1)).els)).toContain('office-engine')
    expect(await status()).toContain('of 6')
  })

  test('slide background', async () => {
    const v = await open()
    await v.page.keyboard.press('Escape')
    await v.btn('Slide background').click()
    await v.page.locator('.de-tb-pop .de-tb-grid button[title="#cfe2f3"]').click()
    await v.page.waitForTimeout(300)
    const s = await saved(v)
    expect((await s.slide(0)).bg).toEqual({ t: 'solid', c: '#cfe2f3' })
  })

  test('present from the beginning', async () => {
    const v = await open()
    await v.page.locator('.de-thumb').nth(3).click()
    await v.pick('Present', 'From beginning')
    await v.page.locator('.de-present').waitFor()
    expect(await v.page.locator('.de-present-n').innerText()).toMatch(/^1\s*\/\s*5$/)
    await v.page.keyboard.press('Escape')
    await v.page.locator('.de-present').waitFor({ state: 'detached' })
  })
})

describe('pptx thumbnails', () => {
  test('a long deck: the thumbnail panel renders the slides it is scrolled to', async () => {
    const v = await viewer({ src: 'samples/large-deck.pptx' })
    const panel = v.page.locator('.de-thumbs')
    await panel.waitFor()
    const num = (n: number) => v.page.locator('.de-thumb-n', { hasText: new RegExp(`^${n}$`) })
    await num(1).waitFor()
    // the panel is mounted after the first render (it needs the root width): scrolling it must still move the window
    await panel.evaluate((el) => {
      el.scrollTop = el.scrollHeight
    })
    await num(305).waitFor({ timeout: 10_000 })
    expect(await num(1).count()).toBe(0)
    await panel.evaluate((el) => {
      el.scrollTop = el.scrollHeight / 2
    })
    await num(153).waitFor({ timeout: 10_000 })
    expect(v.errors).toEqual([])
  })
})

describe('pptx keyboard focus', () => {
  test('the slides have the focus on load; Ctrl/⌘+A selects the slides, not the viewer', async () => {
    const v = await viewer({ src: 'samples/sample.pptx' })
    expect(await v.page.evaluate(() => document.activeElement?.className)).toBe('de-scroll')
    await v.page.keyboard.press('ControlOrMeta+a')
    const text = await v.page.evaluate(() => getSelection()!.toString())
    expect(text).toContain('office-engine')
    expect(text).not.toContain('Slide 1 of')
    expect(v.errors).toEqual([])
  })
})
