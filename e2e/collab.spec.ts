import { afterAll, beforeAll, describe, expect, test } from 'bun:test'
import { readFileSync } from 'node:fs'
import { SAMPLES, useBrowser, viewer } from './harness'

useBrowser()

const PORT = 18_000 + Math.floor(Math.random() * 1000)
const HTTP = `http://localhost:${PORT}`
let server: ReturnType<typeof Bun.spawn> | null = null

beforeAll(async () => {
  server = Bun.spawn(['bun', 'scripts/collab-server.ts'], { env: { ...process.env, PORT: String(PORT) }, stdout: 'ignore', stderr: 'inherit' })
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
  const id = 'e2e-' + Math.random().toString(36).slice(2, 8)
  await fetch(`${HTTP}/rooms/${id}/file`, { method: 'PUT', body: readFileSync(`${SAMPLES}/${sample}`), headers: { 'X-File-Name': sample } })
  return id
}

const join = (id: string, user: string) =>
  viewer({ src: `${HTTP}/rooms/${id}/file`, name: 'shared.xlsx', collab: `ws://localhost:${PORT}/collab`, room: id, user, editable: true })

describe('real-time collaboration', () => {
  test('edits reach the other editor; presence is shown', async () => {
    const id = await room('sample.xlsx')
    const a = await join(id, 'Ann')
    const b = await join(id, 'Bob')
    for (const v of [a, b]) await v.page.locator('.de-collab-live').waitFor({ timeout: 20_000 })
    await a.page.waitForFunction(() => /2 editing/.test(document.querySelector('.de-collab')?.textContent ?? ''), null, { timeout: 20_000 })
    expect(await a.page.locator('.de-peer').first().getAttribute('title')).toBe('Bob')

    await a.typeCell('J10', 'from Ann')
    await b.page.waitForFunction(() => document.body.innerText.includes('from Ann'), null, { timeout: 20_000 })
    expect(await b.shown('J10')).toBe('from Ann')
    await b.typeCell('J11', '=LEN(J10)')
    await a.page.waitForFunction(() => [...document.querySelectorAll('.de-cell')].some((c) => (c as HTMLElement).innerText === '8'), null, { timeout: 20_000 })
  })

  test('New / Open are not offered in a shared room', async () => {
    const id = await room('sample.xlsx')
    const a = await join(id, 'Ann')
    await a.page.locator('.de-collab-live').waitFor({ timeout: 20_000 })
    const labels = await a.menuLabels()
    expect(labels).not.toContain('Open…')
    expect(labels).not.toContain('Document')
    expect(labels).toContain('Excel workbook')
  })
})
