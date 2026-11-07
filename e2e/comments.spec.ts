import { describe, expect, test } from 'bun:test'
import { threadsOf, type DocComment } from '../src/core/comments'
import type { Block, Paragraph } from '../src/core/docx/types'
import { Engine } from '../src/core/worker/engine'
import { useBrowser, viewer, type Viewer } from './harness'

useBrowser()

async function openBytes(bytes: Buffer) {
  const events: { event: string; data: any }[] = []
  const engine = new Engine((event, data) => events.push({ event, data }))
  const info = await engine.open(new Blob([new Uint8Array(bytes)]), {})
  if (info.kind === 'docx') await engine.call('docx.start', [])
  const blocks: Block[] = events.filter((e) => e.event === 'docx:blocks').flatMap((e) => e.data)
  const list = (await engine.call('comments.list', [])) as DocComment[]
  return { engine, blocks, list, threads: threadsOf(list) }
}

/** the text marked by thread `id`, paragraph by paragraph */
const covered = (blocks: Block[], id: string) =>
  blocks
    .filter((b): b is Paragraph => b.t === 'p')
    .map((p) => p.runs.filter((r) => r.t === 'text' && r.cm?.includes(id)).map((r) => (r as { text: string }).text).join(''))
    .filter(Boolean)

const panel = (v: Viewer) => v.page.locator('.de-cm-panel')
const cards = (v: Viewer) => v.page.locator('.de-cm-panel .de-cm-card[data-thread]')

/** types into the focused comment box and posts it */
async function write(v: Viewer, text: string) {
  await v.page.keyboard.type(text)
  await v.page.keyboard.press('Enter')
}

describe('docx comments', () => {
  test('comment on a selection with a mention, reply, resolve, save', async () => {
    const v = await viewer({ src: 'samples/sample.docx', editable: true, onSave: true, me: 'Ann', meId: 'u1', people: true })
    await v.page.locator('.de-ebar').waitFor()
    await v.selectText('Quarterly Report')
    await v.page.keyboard.press(process.platform === 'darwin' ? 'Meta+Alt+KeyM' : 'Control+Alt+KeyM')
    await panel(v).waitFor()
    // the draft shows what it covers
    expect(await v.page.locator('.de-cm-draft .de-cm-quote').innerText()).toBe('Quarterly Report')
    await v.page.keyboard.type('Check this @Bo')
    await v.page.locator('.de-cm-pop button', { hasText: 'Bob Stone' }).waitFor()
    await v.page.keyboard.press('Enter')
    await write(v, 'please')
    await cards(v).first().waitFor()
    expect(await cards(v).count()).toBe(1)
    expect(await cards(v).first().locator('.de-cm-text').innerText()).toBe('Check this @Bob Stone please')
    expect(await cards(v).first().locator('.de-cm-at').innerText()).toBe('@Bob Stone')
    const add = (await v.events('comment')).at(-1)!
    expect(add).toMatchObject({ action: 'add', quote: 'Quarterly Report', mentioned: ['u2'] })
    // the text is highlighted
    expect(await v.page.locator(`.de-pages [data-cm~="${add.id}"]`).allInnerTexts()).toEqual(['Quarterly Report'])
    // the bar counts open threads
    expect(await v.page.locator('[data-testid="de-comments-toggle"] .de-cm-n').innerText()).toBe('1')

    // reply in the active card
    await cards(v).first().locator('textarea').click()
    await write(v, 'Done')
    await v.waitEvent('comment', 2)
    expect((await v.events('comment')).at(-1)).toMatchObject({ action: 'reply', thread: add.id, text: 'Done' })
    expect(await cards(v).first().locator('.de-cm-item').count()).toBe(2)

    // typing inside the commented text keeps it covered; right after it does not
    await v.selectText('Quarterly', 'end')
    await v.page.keyboard.type(' Sales')
    await v.selectText('Report', 'end')
    await v.page.keyboard.type(' 2026')
    await v.page.waitForTimeout(300)
    expect((await v.page.locator(`.de-pages [data-cm~="${add.id}"]`).allInnerTexts()).join('')).toBe('Quarterly Sales Report')

    // resolve: hidden from the document and the list
    await cards(v).first().hover()
    await cards(v).first().locator('button[aria-label="Resolve thread"]').click()
    await v.waitEvent('comment', 3)
    expect((await v.events('comment')).at(-1)).toMatchObject({ action: 'resolve', id: add.id })
    expect(await cards(v).count()).toBe(0)
    expect(await v.page.locator('.de-cm-filter').innerText()).toBe('Resolved (1)')

    const saved = await v.save()
    const re = await openBytes(saved.bytes)
    expect(re.threads.length).toBe(1)
    const t = re.threads[0]
    expect(t.root).toMatchObject({ id: add.id, text: 'Check this @Bob Stone please', done: true, author: { name: 'Ann', id: 'u1' } })
    expect(t.replies.map((r) => r.text)).toEqual(['Done'])
    expect(covered(re.blocks, add.id)).toEqual(['Quarterly Sales Report'])
    expect(v.errors).toEqual([])
  })

  test('undo takes a new comment away, redo brings it back; deleting a thread', async () => {
    const v = await viewer({ src: 'samples/sample.docx', editable: true, onSave: true })
    await v.page.locator('.de-ebar').waitFor()
    await v.selectText('Quarterly Report')
    await v.page.locator('[data-testid="de-comments-toggle"]').click()
    await v.page.locator('.de-cm-new').click()
    await write(v, 'First')
    await cards(v).first().waitFor()
    // back to the document: undo removes the comment's range, so the thread is gone
    await v.selectText('Quarterly', 'end')
    await v.page.keyboard.press(process.platform === 'darwin' ? 'Meta+z' : 'Control+z')
    await v.page.waitForTimeout(400)
    expect(await cards(v).count()).toBe(0)
    await v.page.keyboard.press(process.platform === 'darwin' ? 'Meta+Shift+z' : 'Control+y')
    await v.page.waitForTimeout(400)
    expect(await cards(v).count()).toBe(1)
    // delete the thread from its menu
    await cards(v).first().click()
    await cards(v).first().locator('button[aria-label="More actions"]').click()
    await v.page.locator('.de-ctx button', { hasText: 'Delete thread' }).click()
    await v.page.locator('.de-dialog button', { hasText: 'Delete' }).click()
    await v.page.waitForTimeout(300)
    expect(await cards(v).count()).toBe(0)
    expect(await v.page.locator('.de-pages [data-cm]').count()).toBeGreaterThan(0)
    const re = await openBytes((await v.save()).bytes)
    expect(re.list).toEqual([])
    expect(v.errors).toEqual([])
  })

  test('comments.focus opens the panel on a thread and shows its place', async () => {
    const v = await viewer({ src: 'samples/sample.docx', editable: true, onSave: true })
    await v.page.locator('.de-ebar').waitFor()
    for (const t of ['Quarterly Report', 'This document shows']) {
      await v.selectText(t)
      await v.page.keyboard.press(`${process.platform === 'darwin' ? 'Meta' : 'Control'}+Alt+KeyM`)
      await write(v, 'on ' + t)
      await v.page.waitForTimeout(200)
    }
    const first = (await v.events('comment'))[0].id as string
    await v.page.locator('button[aria-label="Close comments"]').click()
    await v.page.evaluate((id) => window.__setProps({ comments: { focus: id } }), first)
    await panel(v).waitFor()
    await v.page.waitForTimeout(200)
    expect(await v.page.locator('.de-cm-card.on').getAttribute('data-thread')).toBe(first)
  })

  test('view mode shows comments read-only', async () => {
    const v = await viewer({ src: 'samples/sample.docx', editable: true, onSave: true })
    await v.page.locator('.de-ebar').waitFor()
    await v.selectText('Quarterly Report')
    await v.page.locator('[data-testid="de-comments-toggle"]').click()
    await v.page.locator('.de-cm-new').click()
    await write(v, 'Seen by all')
    await cards(v).first().waitFor()
    await v.barButton('View').click()
    await v.page.waitForTimeout(300)
    expect(await v.page.locator('.de-cm-new').count()).toBe(0)
    expect(await cards(v).first().locator('button[aria-label="Resolve thread"]').count()).toBe(0)
    // clicking the highlighted text shows the thread
    await v.page.locator('.de-pages [data-cm]').first().click()
    expect(await cards(v).first().getAttribute('class')).toContain(' on')
  })
})

const MODK = process.platform === 'darwin' ? 'Meta' : 'Control'

describe('xlsx comments', () => {
  test('comment on a cell, mark, answer from the cell, rows inserted above move it, save', async () => {
    const v = await viewer({ src: 'samples/sample.xlsx', editable: true, onSave: true, me: 'Ann', people: true })
    await v.page.locator('.de-ebar').waitFor()
    await v.clickCell('B3')
    await v.page.keyboard.press(`${MODK}+Alt+KeyM`)
    await panel(v).waitFor()
    const quote = await v.page.locator('.de-cm-draft .de-cm-quote').innerText()
    expect(quote).toMatch(/!B3$/)
    await write(v, 'Check this value')
    await cards(v).first().waitFor()
    expect(await v.page.locator('.de-cm-mark').count()).toBe(1)
    // the cell did not get typed into
    expect(await v.shown('B3')).not.toContain('µ')
    // commenting on the same cell again answers its thread
    await v.clickCell('A1')
    await v.clickCell('B3')
    expect(await cards(v).first().getAttribute('class')).toContain(' on')
    await v.page.keyboard.press(`${MODK}+Alt+KeyM`)
    expect(await v.page.locator('.de-cm-draft').count()).toBe(0)
    await cards(v).first().locator('textarea').click()
    await write(v, 'Agreed')
    await v.waitEvent('comment', 2)
    // a row inserted above: the comment follows its cell
    const p = await v.cellPoint('B1')
    await v.page.mouse.click(p.x, p.y, { button: 'right' })
    await v.page.locator('.de-ctx button', { hasText: 'Insert 1 row above' }).click()
    await v.page.waitForTimeout(300)
    await v.clickCell('B4')
    expect(await cards(v).first().getAttribute('class')).toContain(' on')
    expect(await cards(v).first().locator('.de-cm-quote').innerText()).toMatch(/!B4$/)
    const { bytes } = await v.save()
    const re = await openBytes(bytes)
    expect(re.threads.length).toBe(1)
    expect(re.threads[0].root).toMatchObject({ text: 'Check this value', at: { r: 3, c: 1 } })
    expect(re.threads[0].replies.map((r) => r.text)).toEqual(['Agreed'])
    expect(v.errors).toEqual([])
  })
})

describe('pptx comments', () => {
  test('comment at a point of a slide, pin, resolve, save', async () => {
    const v = await viewer({ src: 'samples/sample.pptx', editable: true, onSave: true, me: 'Ann' })
    await v.page.locator('.de-ebar').waitFor()
    const frame = v.page.locator('.de-slide-frame').first()
    const b = (await frame.boundingBox())!
    await v.page.mouse.click(b.x + b.width - 30, b.y + b.height - 30, { button: 'right' })
    await v.page.locator('.de-ctx button', { hasText: /^Comment/ }).click()
    expect(await v.page.locator('.de-cm-draft .de-cm-quote').innerText()).toBe('Slide 1')
    await write(v, 'Too busy')
    await cards(v).first().waitFor()
    const pin = v.page.locator('.de-cm-pin')
    expect(await pin.count()).toBe(1)
    expect(await pin.innerText()).toBe('A')
    const pb = (await pin.boundingBox())!
    // the pin's tip is where the menu was opened (the slide got smaller when the panel opened)
    const b2 = (await frame.boundingBox())!
    expect(Math.abs((pb.x - b2.x) / b2.width - (b.width - 30) / b.width)).toBeLessThan(0.01)
    expect(Math.abs((pb.y + pb.height - b2.y) / b2.height - (b.height - 30) / b.height)).toBeLessThan(0.02)
    // the selected shape: the comment goes next to it, Ctrl/⌘+Alt+M
    await v.page.locator('[data-top]').first().click()
    await v.page.keyboard.press(`${MODK}+Alt+KeyM`)
    await write(v, 'Title too long')
    await v.waitEvent('comment', 2)
    expect(await pin.count()).toBe(2)
    expect(await v.page.locator('.de-thumb').count()).toBeGreaterThan(1)
    const { bytes } = await v.save()
    const re = await openBytes(bytes)
    expect(re.threads.map((t) => t.root.text).sort()).toEqual(['Title too long', 'Too busy'])
    expect(re.threads.every((t) => t.root.at && 'slide' in t.root.at && t.root.at.slide === 'ppt/slides/slide1.xml')).toBe(true)
    expect(v.errors).toEqual([])
  })
})
