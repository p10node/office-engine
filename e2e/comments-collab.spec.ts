import { afterAll, beforeAll, describe, expect, test } from 'bun:test'
import { readFileSync } from 'node:fs'
import { SAMPLES, useBrowser, viewer, type Viewer } from './harness'

useBrowser()

const PORT = 19_000 + Math.floor(Math.random() * 1000)
const HTTP = `http://localhost:${PORT}`
let server: ReturnType<typeof Bun.spawn> | null = null

beforeAll(async () => {
  // frequent snapshots: newcomers join from a file the comments were written into
  server = Bun.spawn(['bun', 'scripts/collab-server.ts'], { env: { ...process.env, PORT: String(PORT), SNAPSHOT_EVERY: '2' }, stdout: 'ignore', stderr: 'inherit' })
  for (let i = 0; i < 100; i++) {
    try {
      if ((await fetch(HTTP)).ok) return
    } catch {
      // not listening yet
    }
    await Bun.sleep(50)
  }
  throw new Error('collab server did not start')
})
afterAll(() => server?.kill())

async function room(sample: string) {
  const id = 'cm-' + Math.random().toString(36).slice(2, 8)
  await fetch(`${HTTP}/rooms/${id}/file`, { method: 'PUT', body: readFileSync(`${SAMPLES}/${sample}`), headers: { 'X-File-Name': sample } })
  return id
}

async function join(id: string, user: string, name: string) {
  const { snapshot } = (await (await fetch(`${HTTP}/rooms/${id}`)).json()) as { snapshot: number }
  const v = await viewer({ src: `${HTTP}/rooms/${id}/file?seq=${snapshot}`, name, collab: `ws://localhost:${PORT}/collab`, room: id, user, me: user, base: snapshot, editable: true, cmOpen: true })
  await v.page.locator('.de-collab-live').waitFor({ timeout: 20_000 })
  await v.page.locator('.de-cm-panel').waitFor()
  return v
}

const MODK = process.platform === 'darwin' ? 'Meta' : 'Control'
const cards = (v: Viewer) => v.page.locator('.de-cm-panel .de-cm-card[data-thread]')
const texts = (v: Viewer) => v.page.locator('.de-cm-panel .de-cm-card[data-thread] .de-cm-text').allInnerTexts()
const until = (v: Viewer, fn: () => Promise<boolean>, what: string) =>
  (async () => {
    for (let i = 0; i < 200; i++) {
      if (await fn()) return
      await v.page.waitForTimeout(100)
    }
    throw new Error(`timed out: ${what}`)
  })()

async function write(v: Viewer, text: string) {
  await v.page.keyboard.type(text)
  await v.page.keyboard.press('Enter')
}

describe('comments in a shared room', () => {
  test('docx: a thread, its range, a reply and resolving reach every editor and newcomers', async () => {
    const id = await room('sample.docx')
    const a = await join(id, 'Ann', 'shared.docx')
    const b = await join(id, 'Bob', 'shared.docx')
    await a.page.locator('.de-ebar').waitFor()
    await a.selectText('Quarterly Report')
    await a.page.keyboard.press(`${MODK}+Alt+KeyM`)
    await write(a, 'Shorter title?')
    await until(b, async () => (await texts(b)).includes('Shorter title?'), 'thread at Bob')
    const tid = (await cards(b).first().getAttribute('data-thread'))!
    // the range came along: Bob's copy highlights the same text
    await until(b, async () => (await b.page.locator(`.de-pages [data-cm~="${tid}"]`).allInnerTexts()).join('') === 'Quarterly Report', 'range at Bob')
    await cards(b).first().click()
    await cards(b).first().locator('textarea').click()
    await write(b, 'Fine as is')
    await until(a, async () => (await texts(a)).includes('Fine as is'), 'reply at Ann')
    // Bob's host hears who else wrote in the thread (to tell them about the reply)
    expect((await b.events('comment')).at(-1)).toMatchObject({ action: 'reply', participants: ['Ann'] })
    expect(await cards(a).first().locator('.de-cm-name').allInnerTexts()).toEqual(['Ann', 'Bob'])
    // Bob types inside the commented text: the range grows for both
    await b.selectText('Quarterly', 'end')
    await b.page.keyboard.type(' Sales')
    await until(a, async () => (await a.page.locator(`.de-pages [data-cm~="${tid}"]`).allInnerTexts()).join('') === 'Quarterly Sales Report', 'range grew at Ann')
    // a newcomer replays the room
    const c = await join(id, 'Cid', 'shared.docx')
    await until(c, async () => (await texts(c)).join('|') === 'Shorter title?|Fine as is', 'thread at Cid')
    await cards(a).first().hover()
    await cards(a).first().locator('button[aria-label="Resolve thread"]').click()
    await until(b, async () => (await cards(b).count()) === 0, 'resolved at Bob')
    expect(await c.page.locator('.de-cm-filter').innerText().catch(() => '')).toMatch(/Resolved \(1\)|^$/)
    for (const v of [a, b, c]) expect(v.errors).toEqual([])
  })

  test('xlsx: a cell thread moves with a row a collaborator inserts; newcomers see it where it went', async () => {
    const id = await room('sample.xlsx')
    const a = await join(id, 'Ann', 'shared.xlsx')
    const b = await join(id, 'Bob', 'shared.xlsx')
    await a.clickCell('B3')
    await a.page.keyboard.press(`${MODK}+Alt+KeyM`)
    await write(a, 'Check')
    await until(b, async () => (await texts(b)).includes('Check'), 'thread at Bob')
    const p = await b.cellPoint('B1')
    await b.page.mouse.click(p.x, p.y, { button: 'right' })
    await b.page.locator('.de-ctx button', { hasText: 'Insert 1 row above' }).click()
    for (const v of [a, b]) await until(v, async () => /!B4$/.test(await cards(v).first().locator('.de-cm-quote').innerText()), 'thread moved')
    // a few more edits: the server takes a snapshot with the comment in the file
    await a.typeCell('J10', 'x')
    await a.typeCell('J11', 'y')
    await Bun.sleep(500)
    const c = await join(id, 'Cid', 'shared.xlsx')
    await until(c, async () => (await cards(c).count()) === 1 && /!B4$/.test(await cards(c).first().locator('.de-cm-quote').innerText()), 'thread at Cid')
    for (const v of [a, b, c]) expect(v.errors).toEqual([])
  })

  test('pptx: threads and replies reach the other editor', async () => {
    const id = await room('sample.pptx')
    const a = await join(id, 'Ann', 'shared.pptx')
    const b = await join(id, 'Bob', 'shared.pptx')
    await a.page.locator('[data-top]').first().click()
    await a.page.keyboard.press(`${MODK}+Alt+KeyM`)
    await write(a, 'Bigger')
    await until(b, async () => (await b.page.locator('.de-cm-pin').count()) === 1, 'pin at Bob')
    await b.page.locator('.de-cm-pin').click()
    await cards(b).first().locator('textarea').click()
    await write(b, 'Ok')
    await until(a, async () => (await texts(a)).join('|') === 'Bigger|Ok', 'reply at Ann')
    for (const v of [a, b]) expect(v.errors).toEqual([])
  })
})
