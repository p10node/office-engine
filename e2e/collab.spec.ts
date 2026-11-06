import { afterAll, beforeAll, describe, expect, test } from 'bun:test'
import { readFileSync } from 'node:fs'
import { rebuildZip } from '../src/core/zipwrite'
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

describe('docx: a built-in style the document lacks', () => {
  /** a document without a styles part: every built-in style is added on first use */
  async function blankRoom() {
    const head = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    const blank = await rebuildZip(
      null,
      new Map([
        ['[Content_Types].xml', `${head}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`],
        ['_rels/.rels', `${head}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`],
        ['word/document.xml', `${head}<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p/><w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>`],
      ]),
    )
    const id = 'e2e-' + Math.random().toString(36).slice(2, 8)
    await fetch(`${HTTP}/rooms/${id}/file`, { method: 'PUT', body: await blank.arrayBuffer(), headers: { 'X-File-Name': 'blank.docx' } })
    return id
  }
  type V = Awaited<ReturnType<typeof viewer>>
  const text = async (v: V) => (await v.pagesText()).replace(/\s+/g, ' ').trim()
  const colour = (v: V, s: string) =>
    v.page.evaluate((x) => {
      const el = [...document.querySelectorAll('.de-pages *')].find((e) => e.childElementCount === 0 && (e.textContent ?? '').includes(x))
      return el ? getComputedStyle(el).color : 'not found'
    }, s)
  /** Heading 1, then typing at once: the style is being added to the file (in a room: ordered by the hub first) */
  async function headingThenType(v: V) {
    await v.page.locator('.de-page-inner').first().click()
    await v.page.keyboard.type('Hello', { delay: 15 })
    await v.page.waitForTimeout(300)
    await v.pick('Paragraph style', 'Heading 1')
    await v.page.keyboard.type(' world')
    await v.page.waitForTimeout(800)
    await v.page.keyboard.type('!')
    await v.page.waitForTimeout(500)
  }
  const HEADING = 'rgb(47, 84, 150)'

  test('in a room: what is typed right after picking it reaches everyone, the caret stays', async () => {
    const id = await blankRoom()
    const join = (user: string) => viewer({ src: `${HTTP}/rooms/${id}/file?seq=0`, name: 'shared.docx', collab: `ws://localhost:${PORT}/collab`, room: id, user, base: 0, editable: true })
    const a = await join('Ann')
    const b = await join('Bob')
    for (const v of [a, b]) await v.page.locator('.de-collab-live').waitFor({ timeout: 20_000 })
    await headingThenType(a)
    // replays the operation that adds the style, then the edits
    const c = await join('Cid')
    await c.page.locator('.de-collab-live').waitFor({ timeout: 20_000 })
    await c.page.waitForFunction(() => document.querySelector('.de-pages')?.textContent?.includes('!'), null, { timeout: 20_000 })
    await c.page.waitForTimeout(800)
    for (const v of [a, b, c]) {
      expect(await text(v)).toBe('Hello world!')
      expect(await colour(v, 'Hello')).toBe(HEADING)
      expect(v.errors).toEqual([])
    }
  })

  test('alone: the caret stays where the typing took it while the style was added', async () => {
    const id = await blankRoom()
    const v = await viewer({ src: `${HTTP}/rooms/${id}/file?seq=0`, name: 'blank.docx', editable: true })
    await v.page.locator('.de-ebar').waitFor()
    await headingThenType(v)
    expect(await text(v)).toBe('Hello world!')
    expect(await colour(v, 'Hello')).toBe(HEADING)
    expect(v.errors).toEqual([])
  })
})
