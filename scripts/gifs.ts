// README GIFs -> docs/screenshots/<scene>.gif
// Scenes play in headless Chromium on the browser-test fixture page (DocViewer only); the frames are screenshots,
// stacked side by side and encoded by ffmpeg.
//   xlsx-300k  the 300k-row sheet opens: rows stream in, then a jump to the last row
//   collab     Ann and Bob edit one sheet side by side through the reference collab server, Cara joins from a third browser
// Usage: bun run gifs [xlsx-300k|collab]   (needs a Chromium for playwright-core: `bunx playwright install chromium`, ffmpeg,
//        and large-sheet.xlsx: `python3 scripts/fixtures.py --large`)
import { existsSync } from 'node:fs'
import { mkdir, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { chromium, type Page } from 'playwright-core'
import { createServer } from 'vite'

const OUT = 'docs/screenshots'
const SAMPLES = 'playground/public/samples'
const FPS = 8
const TYPE_DELAY = 70 // ms per key, so typing is visible
const GAP = 6 // px between side-by-side browsers

// ---------------------------------------------------------------- recording

class Recorder {
  frames: Buffer[][] = [] // per tick, one shot per page
  private running = false
  private loop: Promise<void> | null = null
  constructor(private pages: Page[]) {}
  start() {
    this.running = true
    this.loop = (async () => {
      while (this.running) {
        const t = Date.now()
        this.frames.push(await Promise.all(this.pages.map((p) => p.screenshot({ type: 'png' }))))
        const left = 1000 / FPS - (Date.now() - t)
        if (left > 0) await Bun.sleep(left)
      }
    })()
  }
  async stop() {
    this.running = false
    await this.loop
    return this.frames
  }
}

/** encodes the frames (pages side by side) as a looping GIF */
async function encode(frames: Buffer[][], out: string) {
  const n = frames[0].length
  const dir = join(tmpdir(), 'office-engine-gif')
  await rm(dir, { recursive: true, force: true })
  await Promise.all(Array.from({ length: n }, (_, i) => mkdir(join(dir, String(i)), { recursive: true })))
  await Promise.all(frames.flatMap((shots, f) => shots.map((png, i) => Bun.write(join(dir, String(i), `${String(f).padStart(4, '0')}.png`), png))))
  const inputs = Array.from({ length: n }, (_, i) => ['-framerate', String(FPS), '-i', join(dir, String(i), '%04d.png')]).flat()
  const pads = Array.from({ length: n - 1 }, (_, i) => `[${i}]pad=iw+${GAP}:ih:0:0:color=#dadce0[p${i}];`).join('')
  const stack = n > 1 ? `${pads}${Array.from({ length: n }, (_, i) => (i < n - 1 ? `[p${i}]` : `[${i}]`)).join('')}hstack=inputs=${n},` : '[0]'
  const filter = `${stack}split[s0][s1];[s0]palettegen=stats_mode=diff[pal];[s1][pal]paletteuse=dither=bayer:bayer_scale=5:diff_mode=rectangle`
  const ffmpeg = Bun.spawn(['ffmpeg', '-y', '-loglevel', 'error', ...inputs, '-filter_complex', filter, '-loop', '0', out], { stdout: 'inherit', stderr: 'inherit' })
  if ((await ffmpeg.exited) !== 0) throw new Error('ffmpeg failed')
  await rm(dir, { recursive: true, force: true })
  console.log(`${out}: ${frames.length} frames, ${(Bun.file(out).size / 1024).toFixed(0)} KB`)
}

// ---------------------------------------------------------------- the fixture page

const vite = await createServer({ configFile: 'vite.e2e.config.ts', server: { port: 0, strictPort: false } })
await vite.listen()
const base = vite.resolvedUrls!.local[0]
const browser = await chromium.launch()

const fixture = (params: Record<string, string | number>) => `${base}?${new URLSearchParams(Object.entries(params).map(([k, v]) => [k, String(v)]))}`

const screen = async (width: number, height: number) => (await browser.newContext({ viewport: { width, height }, deviceScaleFactor: 1 })).newPage()

/** a name strip above the viewer, so side-by-side screens are told apart */
const label = (page: Page, text: string, h = 28) =>
  page.evaluate(
    ({ text, h }) => {
      const el = document.createElement('div')
      el.textContent = text
      el.style.cssText = `height:${h}px;display:flex;align-items:center;padding:0 12px;font:600 13px system-ui,-apple-system,sans-serif;color:#fff;background:#202124`
      document.body.prepend(el)
      window.__setProps({ style: { height: `calc(100vh - ${h}px)` } })
    },
    { text, h },
  )

const cellPoint = async (page: Page, ref: string) => {
  const [, col, row] = /^([A-Z]+)(\d+)$/.exec(ref)!
  const head = async (sel: string, text: string) => {
    const b = await page.locator(sel).filter({ hasText: new RegExp(`^${text}$`) }).first().boundingBox()
    if (!b) throw new Error(`${ref}: header ${text} not visible`)
    return b
  }
  const c = await head('.de-head-c', col)
  const r = await head('.de-head-r', row)
  return { x: c.x + c.width / 2, y: r.y + r.height / 2 }
}
const click = async (page: Page, ref: string) => {
  const p = await cellPoint(page, ref)
  await page.mouse.click(p.x, p.y)
}
const type = async (page: Page, ref: string, text: string) => {
  await click(page, ref)
  await Bun.sleep(300)
  await page.keyboard.type(text, { delay: TYPE_DELAY })
  await Bun.sleep(250)
  await page.keyboard.press('Enter')
}
const drag = async (page: Page, a: string, b: string) => {
  const p = await cellPoint(page, a)
  const q = await cellPoint(page, b)
  await page.mouse.move(p.x, p.y)
  await page.mouse.down()
  await page.mouse.move(q.x, q.y, { steps: 8 })
  await page.mouse.up()
}

// ---------------------------------------------------------------- scenes

/** the 300k-row sheet opens: rows stream in while the status counts them, then Ctrl+Down lands on the last row */
async function xlsx300k() {
  if (!existsSync(`${SAMPLES}/large-sheet.xlsx`)) throw new Error('large-sheet.xlsx is missing: python3 scripts/fixtures.py --large')
  const page = await screen(1440, 811) // the README screenshots' size
  const rec = new Recorder([page])
  await page.goto(fixture({ src: 'samples/large-sheet.xlsx', editable: 1 }))
  await page.locator('.de-root').waitFor()
  rec.start()
  await page.locator('.de-grid').waitFor({ timeout: 60_000 })
  await page.locator('.de-bar').getByText(/loading/).waitFor({ state: 'hidden', timeout: 90_000 })
  await Bun.sleep(1500)
  await click(page, 'A2')
  await Bun.sleep(500)
  await page.keyboard.press('ControlOrMeta+ArrowDown')
  await Bun.sleep(2500)
  return rec.stop()
}

/** Ann and Bob share one sheet through the collab server, Cara joins from a third (unrecorded) browser */
async function collab() {
  const PORT = 18_000 + Math.floor(Math.random() * 1000)
  const HTTP = `http://localhost:${PORT}`
  const server = Bun.spawn(['bun', 'scripts/collab-server.ts'], { env: { ...process.env, PORT: String(PORT) }, stdout: 'ignore', stderr: 'inherit' })
  try {
    for (let i = 0; ; i++) {
      try {
        if ((await fetch(HTTP)).ok) break
      } catch {
        if (i > 100) throw new Error('collab server did not start')
      }
      await Bun.sleep(50)
    }
    const room = 'readme-' + Math.random().toString(36).slice(2, 8)
    await fetch(`${HTTP}/rooms/${room}/file`, { method: 'PUT', body: Bun.file(`${SAMPLES}/sample.xlsx`), headers: { 'X-File-Name': 'sales.xlsx' } })

    const join = async (user: string, recorded: boolean) => {
      const page = await screen(720, recorded ? 478 : 450)
      await page.goto(fixture({ src: `${HTTP}/rooms/${room}/file?seq=0`, name: 'sales.xlsx', collab: `ws://localhost:${PORT}/collab`, room, user, editable: 1 }))
      await page.locator('.de-grid').waitFor({ timeout: 30_000 })
      await page.locator('.de-collab-live').waitFor({ timeout: 30_000 })
      if (recorded) await label(page, `${user}'s browser`)
      return page
    }
    const peers = (page: Page, n: number) => page.waitForFunction((n) => new RegExp(`${n} editing`).test(document.querySelector('.de-collab')?.textContent ?? ''), n, { timeout: 20_000 })

    const ann = await join('Ann', true)
    const bob = await join('Bob', true)
    await peers(ann, 2)
    await peers(bob, 2)
    const rec = new Recorder([ann, bob])
    rec.start()

    await Bun.sleep(1200)
    await type(ann, 'G2', 'Owner')
    await click(ann, 'G2')
    await ann.keyboard.press('ControlOrMeta+b')
    await Bun.sleep(600)
    await type(bob, 'G3', 'Bob')
    await type(bob, 'G4', 'Bob')
    await type(ann, 'G5', 'Ann')
    await Bun.sleep(400)

    const cara = await join('Cara', false) // avatars and "3 editing" appear on both screens
    await peers(ann, 3)
    await Bun.sleep(800)
    await type(cara, 'G6', 'Cara')
    await click(cara, 'I12') // park Cara's cursor: its name flag would otherwise sit on top of G6
    await Bun.sleep(600)

    await drag(bob, 'E3', 'E8') // Bob's selection shows up on Ann's screen
    await Bun.sleep(1000)
    await type(ann, 'G10', '=COUNTA(G3:G9)')
    await Bun.sleep(2500)
    return rec.stop()
  } finally {
    server.kill()
  }
}

// ---------------------------------------------------------------- main

const SCENES: Record<string, () => Promise<Buffer[][]>> = { 'xlsx-300k': xlsx300k, collab }
const wanted = process.argv.length > 2 ? process.argv.slice(2) : Object.keys(SCENES)
try {
  for (const name of wanted) {
    const scene = SCENES[name]
    if (!scene) throw new Error(`unknown scene ${name} (${Object.keys(SCENES).join(', ')})`)
    await encode(await scene(), `${OUT}/${name}.gif`)
  }
} finally {
  await browser.close()
  await vite.close()
}
