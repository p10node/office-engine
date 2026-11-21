import { describe, expect, test } from 'bun:test'
import { parasToDocx } from '../src/core/convert'
import { BlobSource } from '../src/core/source'
import { ZipArchive } from '../src/core/zip'
import { rebuildZip } from '../src/core/zipwrite'
import { docxParas, useBrowser, viewer } from './harness'

useBrowser()

describe('large plain imports in Chromium', () => {
  for (const kind of ['markdown', 'html', 'code'] as const) {
    test(`${kind}: imports beyond V8's function-argument limit`, async () => {
      const v = await viewer({}, { page: 'plain.html' })
      const result = await v.page.evaluate((kind) => {
        const { htmlParas, markdownParas } = (window as unknown as { __plain: typeof import('../src/core/plain') }).__plain
        const count = 150_000
        const ps = kind === 'html'
          ? htmlParas('<h1>First</h1>' + '<p>paragraph</p>'.repeat(count) + '<h1>Last</h1>', false)
          : markdownParas(kind === 'code' ? '```text\n' + 'line\n'.repeat(count) + '```\n\nLast\n' : '# First\n\n' + 'paragraph\n\n'.repeat(count) + '# Last\n', false)
        const text = (i: number) => ps[i].runs.map((r) => r.text).join('')
        return { count: ps.length, first: kind === 'code' ? text(0).split('line').length - 1 : text(0), last: text(ps.length - 1) }
      }, kind)
      expect(result).toEqual({ count: kind === 'code' ? 2 : 150_002, first: kind === 'code' ? 150_000 : 'First', last: 'Last' })
      expect(v.errors).toEqual([])
    })
  }
})

test('a document beyond the browser scroll-height limit reaches, finds and edits its last page', async () => {
  // Tall pages exercise the same scroll limit as a many-thousand-page Markdown
  // import, without making the regression test parse a 50 MB fixture.
  const blob = await parasToDocx(Array.from({ length: 200 }, (_, i) => ({ runs: [{ text: i === 199 ? 'END_MARKER' : `Page ${i + 1}` }] })))
  const zip = await ZipArchive.open(new BlobSource(blob))
  const xml = (await zip.text('word/document.xml'))
    .replace(/<w:pgSz[^>]*\/>/, '<w:pgSz w:w="12240" w:h="2000000"/>')
    .replaceAll('<w:p>', '<w:p><w:pPr><w:pageBreakBefore/></w:pPr>')
  const file = await rebuildZip(zip, new Map([['word/document.xml', xml]]), blob.type)
  const v = await viewer({ editable: true, onSave: true })
  await v.page.evaluate((bytes) => window.__setProps({ src: new Uint8Array(bytes), name: 'long.docx' }), [...new Uint8Array(await file.arrayBuffer())])
  await v.ready('docx')
  await v.page.waitForFunction(() => document.querySelector('.de-bar')?.textContent?.includes('of 200'))
  const height = await v.page.locator('.de-scroll').evaluate((el) => {
    el.scrollTop = el.scrollHeight
    return el.scrollHeight
  })
  expect(height).toBeLessThanOrEqual(12_000_000)
  await v.page.locator('.de-page[data-page="200"]').waitFor()
  expect(await v.page.locator('.de-page[data-page="200"]').textContent()).toContain('END_MARKER')

  await v.btn('Find & replace').click()
  const input = v.page.locator('.de-find input[aria-label="Find"]')
  await input.fill('END_MARKER')
  await input.press('Enter')
  await v.page.waitForFunction(() => {
    const p = document.querySelector('.de-page[data-page="200"] [data-pa]')?.getBoundingClientRect()
    const sc = document.querySelector('.de-scroll')?.getBoundingClientRect()
    return p && sc && p.top >= sc.top && p.bottom <= sc.bottom
  }, null, { timeout: 5000 }).catch(async (e) => {
    const state = await v.page.evaluate(() => {
      const s = document.querySelector('.de-scroll')!
      const p = document.querySelector('.de-page[data-page="200"] [data-pa]')
      return { top: s.scrollTop, height: s.scrollHeight, viewport: s.getBoundingClientRect().toJSON(), paragraph: p?.getBoundingClientRect().toJSON(), status: document.querySelector('.de-bar')?.textContent, matches: document.querySelector('.de-find-count')?.textContent, pages: [...document.querySelectorAll('.de-page')].map(p => ({ page: p.getAttribute('data-page'), top: (p as HTMLElement).style.top })) }
    })
    throw new Error(`${e}\n${JSON.stringify(state)}`)
  })
  await input.press('Escape')
  await v.selectText('END_MARKER', 'end')
  await v.page.keyboard.type('!')
  await v.page.waitForFunction(() => document.querySelector('.de-page[data-page="200"]')?.textContent?.includes('END_MARKER!'))
  const saved = await docxParas((await v.save()).bytes)
  expect(saved.at(-1)?.runs.filter((r) => r.t === 'text').map((r) => r.text).join('')).toBe('END_MARKER!')
  expect(v.errors).toEqual([])
})
