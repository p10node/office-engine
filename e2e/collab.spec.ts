import { afterAll, beforeAll, describe, expect, test } from 'bun:test'
import { readFileSync } from 'node:fs'
import { SAMPLES, useBrowser, viewer } from './harness'

useBrowser()

const PORT = 18_000 + Math.floor(Math.random() * 1000)
const HTTP = `http://localhost:${PORT}`
let server: ReturnType<typeof Bun.spawn> | null = null

beforeAll(async () => {
  // frequent snapshots: live editing must not notice them, newcomers join from them
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
  const id = 'e2e-' + Math.random().toString(36).slice(2, 8)
  await fetch(`${HTTP}/rooms/${id}/file`, { method: 'PUT', body: readFileSync(`${SAMPLES}/${sample}`), headers: { 'X-File-Name': sample } })
  return id
}

const join = (id: string, user: string, base = 0) =>
  viewer({ src: `${HTTP}/rooms/${id}/file?seq=${base}`, name: 'shared.xlsx', collab: `ws://localhost:${PORT}/collab`, room: id, user, base, editable: true })

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

  test('a newcomer opens the latest snapshot and only replays later edits', async () => {
    const id = await room('sample.xlsx')
    const a = await join(id, 'Ann')
    await a.page.locator('.de-collab-live').waitFor({ timeout: 20_000 })
    for (const [ref, text] of [['J10', 'one'], ['J11', 'two'], ['J12', 'three']]) await a.typeCell(ref, text)
    let snapshot = 0
    for (let i = 0; i < 100 && snapshot < 2; i++) {
      snapshot = ((await (await fetch(`${HTTP}/rooms/${id}`)).json()) as { snapshot: number }).snapshot
      await Bun.sleep(100)
    }
    expect(snapshot).toBeGreaterThanOrEqual(2)

    const b = await join(id, 'Bob', snapshot)
    await b.page.locator('.de-collab-live').waitFor({ timeout: 20_000 })
    await b.page.waitForFunction(() => document.body.innerText.includes('three'), null, { timeout: 20_000 })
    expect(await b.shown('J10')).toBe('one')
    await a.typeCell('J13', 'after')
    await b.page.waitForFunction(() => document.body.innerText.includes('after'), null, { timeout: 20_000 })
    await b.typeCell('J14', 'from Bob')
    await a.page.waitForFunction(() => document.body.innerText.includes('from Bob'), null, { timeout: 20_000 })
    expect(b.errors).toEqual([])
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
