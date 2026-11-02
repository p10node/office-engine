/**
 * Browser test harness: the fixture app (e2e/app) on a Vite dev server, driven by Chromium (playwright-core).
 * Needs a Chromium for playwright-core: `bunx playwright install chromium`.
 */
import { afterAll, afterEach, beforeAll, expect, setDefaultTimeout } from 'bun:test'
import { existsSync, readFileSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { chromium, type Browser, type BrowserContext, type Download, type Locator, type Page } from 'playwright-core'
import { createServer, type ViteDevServer } from 'vite'
import type { Block, Paragraph } from '../src/core/docx/types'
import { Engine } from '../src/core/worker/engine'

export const SAMPLES = `${import.meta.dir}/../playground/public/samples`

/** the last harness failure, in full: a terminal may trim bun's failure block */
export const FAILURE_LOG = join(tmpdir(), 'docs-engine-e2e-last-failure.log')
const dumped = new WeakSet<Error>()
function dump(e: Error, extra = '') {
  if (dumped.has(e)) return
  dumped.add(e)
  writeFileSync(FAILURE_LOG, `${new Date().toISOString()}\n${e.stack ?? e.message}${extra}\n`)
}

interface Shared {
  base: string
  browser: Browser
}

let shared: Shared | null = null
let server: ViteDevServer | null = null
const open = new Set<BrowserContext>()

/** Starts the fixture server and the browser for the current test file. */
export function useBrowser() {
  // per file: the setting does not carry over to the files that run next
  setDefaultTimeout(90_000)
  beforeAll(async () => {
    server = await createServer({ configFile: 'vite.e2e.config.ts', server: { port: 0, strictPort: false } })
    await server.listen()
    const browser = await chromium.launch()
    shared = { base: server.resolvedUrls!.local[0], browser }
  })
  afterEach(async () => {
    for (const c of open) await c.close().catch(() => {})
    open.clear()
  })
  afterAll(async () => {
    await shared?.browser.close()
    await server?.close()
    shared = null
    server = null
  })
}

export type Params = Record<string, string | number | boolean>

/** Opens the fixture with DocViewer props from `params` and waits for the document. */
export async function viewer(params: Params, opts: { wait?: boolean; page?: string; viewport?: { width: number; height: number } } = {}) {
  if (!shared) throw new Error('useBrowser() missing')
  const ctx = await shared.browser.newContext({ viewport: opts.viewport ?? { width: 1600, height: 1000 }, acceptDownloads: true })
  open.add(ctx)
  const page = await ctx.newPage()
  const v = new Viewer(page, ctx)
  // a missing sample is served as the fixture's index.html and fails later as "not a ZIP": say so (tests of a bad file pass wait: false)
  const sample = typeof params.src === 'string' && params.src.startsWith('samples/') ? params.src.slice('samples/'.length) : ''
  if (opts.wait !== false && sample && !existsSync(`${SAMPLES}/${sample}`)) throw new Error(`${params.src} is missing: large-* samples are generated (python3 scripts/fixtures.py --large-deck)`)
  const q = new URLSearchParams(Object.entries(params).map(([k, x]) => [k, String(x === true ? 1 : x === false ? 0 : x)]))
  try {
    await page.goto(`${shared.base}${opts.page ?? ''}?${q}`)
    if (opts.wait !== false && params.src) await v.ready()
  } catch (e) {
    const events = await v.events().catch(() => [])
    dump(e as Error, `\nevents: ${JSON.stringify(events)}\nerrors: ${JSON.stringify(v.errors)}`)
    throw e
  }
  return v
}

export const READY = { docx: '.de-page-inner', xlsx: '.de-grid', pptx: '.de-slide-inner' } as const

export class Viewer {
  readonly errors: string[] = []

  constructor(
    readonly page: Page,
    readonly ctx: BrowserContext,
  ) {
    page.on('pageerror', (e) => this.errors.push(e.message))
    page.on('console', (m) => {
      if (m.type() === 'error' && !/Failed to load resource/.test(m.text())) this.errors.push(m.text())
    })
  }

  /** document rendered (and its first progressive work settled); a load error fails here, with its message */
  async ready(kind?: keyof typeof READY) {
    await this.page
      .waitForFunction(() => window.__events?.some((e) => e.type === 'load' || e.type === 'error'), null, { timeout: 30_000 })
      .catch((e: Error) => this.fail('document did not load', e))
    const err = (await this.events('error')).at(-1)
    if (err && !(await this.events('load')).length) await this.fail(`document failed to load: ${err.message}`)
    const k = kind ?? ((await this.events('load')).at(-1)?.kind as keyof typeof READY)
    if (k) await this.page.locator(READY[k]).first().waitFor({ timeout: 30_000 }).catch((e: Error) => this.fail(`${k} did not render`, e))
    await this.page.waitForTimeout(400)
  }

  /** throws `what` together with what the page recorded (callback events, console and page errors), so a timeout names its cause */
  private async fail(what: string, cause?: Error): Promise<never> {
    const events = (await this.events().catch(() => [])).map(({ b64, ...e }) => e)
    const err = new Error(`${what}${cause ? `: ${cause.message.split('\n')[0]}` : ''}\nevents: ${JSON.stringify(events)}\nerrors: ${JSON.stringify(this.errors)}`)
    dump(err)
    throw err
  }

  async events(type?: string): Promise<{ type: string; [k: string]: any }[]> {
    return this.page.evaluate((t) => window.__events.filter((e) => !t || e.type === t), type)
  }

  async waitEvent(type: string, count = 1) {
    await this.page.waitForFunction(([t, n]) => window.__events.filter((e) => e.type === t).length >= (n as number), [type, count] as const, { timeout: 30_000 })
    return (await this.events(type)).at(-1)!
  }

  /** toolbar / bar button by the start of its label ("Bold" matches "Bold (⌘B)") */
  btn(label: string): Locator {
    return this.page.locator(`button[aria-label^="${label}"]`).first()
  }

  barButton(text: string): Locator {
    return this.page.locator('.de-bar button.de-mode', { hasText: text })
  }

  async edit() {
    if (await this.barButton('Edit').count()) await this.barButton('Edit').click()
    await this.page.locator('.de-ebar').first().waitFor()
  }

  // ---------------------------------------------------------------- dropdowns / File menu

  /** option of the open dropdown with exactly this label */
  option(label: string | RegExp): Locator {
    const re = typeof label === 'string' ? new RegExp(`^${label.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}$`) : label
    return this.page.locator('.de-dd-pop .de-dd-opt').filter({ has: this.page.locator('.de-dd-text', { hasText: re }) }).first()
  }

  async pick(dropdownTitle: string, label: string | RegExp) {
    await this.page.locator(`button[aria-label^="${dropdownTitle}"]`).first().click()
    await this.option(label).click()
  }

  async menu(label?: string | RegExp) {
    await this.page.locator('.de-bar .de-file').click()
    if (label !== undefined) await this.option(label).click()
  }

  async menuLabels(): Promise<string[]> {
    await this.page.locator('.de-bar .de-file').click()
    const labels = await this.page.locator('.de-dd-pop .de-dd-opt .de-dd-text').allInnerTexts()
    await this.page.keyboard.press('Escape')
    return labels
  }

  async download(trigger: () => Promise<unknown>): Promise<{ name: string; bytes: Buffer; text: string }> {
    const [dl] = await Promise.all([this.page.waitForEvent('download', { timeout: 30_000 }), trigger()])
    return readDownload(dl)
  }

  async downloadAs(label: string) {
    return this.download(() => this.menu(label))
  }

  /** answers the OS file picker opened by `trigger` */
  async chooseFile(trigger: () => Promise<unknown>, file: string | { name: string; mimeType: string; buffer: Buffer }) {
    const [chooser] = await Promise.all([this.page.waitForEvent('filechooser'), trigger()])
    await chooser.setFiles(file)
  }

  async toast(): Promise<string> {
    const t = this.page.locator('.de-toast').last()
    await t.waitFor()
    return t.innerText()
  }

  async title() {
    return this.page.locator('.de-bar .de-title').innerText()
  }

  async status() {
    return (await this.page.locator('.de-bar .de-status').allInnerTexts()).join(' ')
  }

  // ---------------------------------------------------------------- document text

  /** selects `text` (inside one run) in the editable pages; `collapse` puts the caret before / after it */
  async selectText(text: string, collapse?: 'start' | 'end') {
    const ok = await this.page.evaluate(
      ([t, c]) => {
        const root = document.querySelector<HTMLElement>('.de-pages')!
        const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT)
        for (let n = walker.nextNode(); n; n = walker.nextNode()) {
          const i = n.textContent!.indexOf(t!)
          if (i < 0 || n.parentElement?.closest('[contenteditable=false]')) continue
          root.focus()
          const r = document.createRange()
          r.setStart(n, c === 'end' ? i + t!.length : i)
          r.setEnd(n, c === 'start' ? i : i + t!.length)
          const sel = getSelection()!
          sel.removeAllRanges()
          sel.addRange(r)
          return true
        }
        return false
      },
      [text, collapse] as const,
    )
    expect(ok).toBe(true)
    await this.page.waitForTimeout(120)
  }

  async pagesText() {
    return this.page.locator('.de-pages').innerText()
  }

  /** Save button → the file handed to onSave (fixture param `onSave`) */
  async save() {
    const n = (await this.events('save')).length
    await this.barButton('Save').click()
    const e = await this.waitEvent('save', n + 1)
    return { name: e.name as string, bytes: fromB64(e.b64 as string) }
  }

  // ---------------------------------------------------------------- spreadsheet grid

  async cellPoint(ref: string) {
    const [, col, row] = /^([A-Z]+)(\d+)$/.exec(ref)!
    const head = async (sel: string, text: string) => {
      const b = await this.page.locator(sel).filter({ hasText: new RegExp(`^${text}$`) }).first().boundingBox()
      if (!b) throw new Error(`header ${text} not visible`)
      return b
    }
    const c = await head('.de-head-c', col)
    const r = await head('.de-head-r', row)
    return { x: c.x + c.width / 2, y: r.y + r.height / 2 }
  }

  async clickCell(ref: string) {
    const p = await this.cellPoint(ref)
    await this.page.mouse.click(p.x, p.y)
  }

  async selectRange(a: string, b: string) {
    const p = await this.cellPoint(a)
    const q = await this.cellPoint(b)
    await this.page.mouse.move(p.x, p.y)
    await this.page.mouse.down()
    await this.page.mouse.move(q.x, q.y, { steps: 4 })
    await this.page.mouse.up()
    expect(await this.page.locator('.de-fx-ref').innerText()).toBe(`${a}:${b}`)
  }

  /** types into a cell and commits with Enter */
  async typeCell(ref: string, text: string) {
    await this.clickCell(ref)
    await this.page.keyboard.type(text)
    await this.page.keyboard.press('Enter')
    await this.page.waitForTimeout(150)
  }

  /** shown text of the cell under the given reference */
  async shown(ref: string): Promise<string> {
    const p = await this.cellPoint(ref)
    return this.page.evaluate(({ x, y }) => {
      for (const el of document.elementsFromPoint(x, y)) {
        const cell = el.closest('.de-cell')
        if (cell) return (cell as HTMLElement).innerText
      }
      return ''
    }, p)
  }

  /** formula bar content of a cell */
  async formula(ref: string) {
    await this.clickCell(ref)
    await this.page.waitForFunction((r) => document.querySelector('.de-fx-ref')?.textContent === r, ref)
    // the cell detail (formula) is fetched from the engine
    await this.page.waitForTimeout(300)
    return this.page.locator('.de-fx-input').inputValue()
  }
}

async function readDownload(dl: Download) {
  const path = await dl.path()
  const bytes = readFileSync(path!)
  return { name: dl.suggestedFilename(), bytes, text: bytes.toString('utf8') }
}

// ---------------------------------------------------------------- saved files

export async function openFile(bytes: Uint8Array | Blob) {
  const events: { event: string; data: any }[] = []
  const engine = new Engine((event, data) => events.push({ event, data }))
  const blob = bytes instanceof Blob ? bytes : new Blob([bytes as Uint8Array<ArrayBuffer>])
  const info = await engine.open(blob, {})
  return { engine, info, events }
}

export async function docxBlocks(bytes: Uint8Array | Blob): Promise<Block[]> {
  const { engine, events } = await openFile(bytes)
  await engine.call('docx.start', [])
  return events.filter((e) => e.event === 'docx:blocks').flatMap((e) => e.data)
}

export async function docxParas(bytes: Uint8Array | Blob): Promise<Paragraph[]> {
  return (await docxBlocks(bytes)).filter((b): b is Paragraph => b.t === 'p')
}

/** text runs of a paragraph with their look */
export const runsOf = (p: Paragraph) => p.runs.flatMap((r) => (r.t === 'text' ? [r] : []))

/** a 2×2 PNG */
export const PNG = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAIAAAD91JpzAAAAFklEQVR4nGP4z8DAwMDAxMDAwMDAAAANHQEDK+mmyAAAAABJRU5ErkJggg==', 'base64')

export const paraText = (p: Paragraph) => p.runs.map((r) => (r.t === 'text' ? r.text : '')).join('')

/** a saved workbook: shown values, merges, frozen panes, gridlines of one sheet */
export async function xlsxSheet(bytes: Uint8Array | Blob, sheet = 0) {
  const { engine, info, events } = await openFile(bytes)
  await engine.call('xlsx.load', [sheet])
  const prog = () => events.filter((e) => e.event === 'xlsx:progress' && e.data.sheet === sheet).map((e) => e.data)
  for (let i = 0; i < 200 && !prog().some((p) => p.done); i++) await new Promise((r) => setTimeout(r, 10))
  const all = prog()
  const block: any = await engine.call('xlsx.block', [sheet, 0, 500])
  const at = (ref: string) => {
    const [, col, row] = /^([A-Z]+)(\d+)$/.exec(ref)!
    const c = [...col].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1
    return block.rows.find((x: any) => x.r === Number(row) - 1)?.cells.find((x: any) => x[0] === c)
  }
  return {
    engine,
    names: info.xlsx!.sheets.map((s) => s.name),
    cell: (ref: string): string => at(ref)?.[1] ?? '',
    merges: ([...all].reverse().find((p: any) => p.merges)?.merges ?? []) as { r0: number; c0: number; r1: number; c1: number }[],
    frozen: all.find((p) => p.frozen)?.frozen as { rows: number; cols: number } | undefined,
    gridLines: all.find((p) => p.gridLines !== undefined)?.gridLines as boolean | undefined,
  }
}

export const fromB64 = (s: string) => Buffer.from(s, 'base64')
