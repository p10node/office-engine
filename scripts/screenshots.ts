// README screenshots of the playground -> docs/screenshots/*.png
// Each shot is the viewer only (`.de-root`: File menu row and below), the playground chrome is cropped out.
// Usage: bun run screenshots   (needs a Chromium for playwright-core: `bunx playwright install chromium`)
import { mkdir } from 'node:fs/promises'
import { chromium } from 'playwright-core'
import { createServer } from 'vite'

const OUT = 'docs/screenshots'
const SHOTS = [
  { name: 'docx', src: 'samples/sample.docx', ready: '.de-page-inner' },
  { name: 'xlsx', src: 'samples/sample.xlsx', ready: '.de-grid' },
  { name: 'pptx', src: 'samples/sample.pptx', ready: '.de-slide-inner' },
]

await mkdir(OUT, { recursive: true })
const server = await createServer({ configFile: 'vite.playground.config.ts', server: { port: 5199 }, logLevel: 'error' })
await server.listen()
const base = server.resolvedUrls!.local[0]
const browser = await chromium.launch()
try {
  const page = await browser.newPage({ viewport: { width: 1440, height: 900 }, deviceScaleFactor: 2 })
  await page.addInitScript(() => localStorage.setItem('de-user', 'Ann')) // stable collaborator name
  for (const s of SHOTS) {
    await page.goto(`${base}?src=${encodeURIComponent(s.src)}`)
    await page.getByText(/opened in \d+ ms/).waitFor({ timeout: 90_000 })
    await page.locator(s.ready).first().waitFor()
    await page.locator('.de-bar').getByText(/loading/).waitFor({ state: 'hidden', timeout: 90_000 }) // full row count in the status
    await page.waitForTimeout(1500) // images, fonts, progressive layout
    await page.locator('.de-root').screenshot({ path: `${OUT}/${s.name}.png` })
    console.log(`${OUT}/${s.name}.png`)
  }
} finally {
  await browser.close()
  await server.close()
}
