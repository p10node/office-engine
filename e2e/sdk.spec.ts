import { describe, expect, test } from 'bun:test'
import { useBrowser, viewer } from './harness'

useBrowser()

describe('SDK packages in a page', () => {
  test('diagram blocks drawn in place, one rendered into an element, a PNG, Markdown with a drawing', async () => {
    const v = await viewer({}, { page: 'sdk.html' })
    const { page } = v
    const res = await page.evaluate(() => (window as unknown as { __sdk: Promise<{ drawn: number; png: { type: string; size: number } }> }).__sdk)
    expect(res.drawn).toBe(4)
    expect(await page.locator('figure.diagram').evaluateAll((els) => els.map((e) => e.className))).toEqual([
      'diagram diagram-mermaid',
      'diagram diagram-plantuml',
      'diagram diagram-graphviz',
      'diagram diagram-mermaid',
    ])
    // what cannot be drawn stays, with the reason; other code blocks are left alone
    expect(await page.locator('#broken').getAttribute('data-diagram-error')).toContain('gitGraph')
    expect(await page.locator('code.language-js').count()).toBe(1)
    const texts = await page.locator('figure.diagram svg text').allTextContents()
    for (const t of ['Write', 'Draw', 'Alice', 'hi', 'x', 'y', 'Ann', 'detected']) expect(texts).toContain(t)
    expect(await page.locator('#one svg').count()).toBe(1)
    expect(res.png.type).toBe('image/png')
    expect(res.png.size).toBeGreaterThan(500)
    expect(await page.locator('#md h1#notes').innerText()).toBe('Notes')
    expect(await page.locator('#md input[type=checkbox][checked]').count()).toBe(1)
    expect(await page.locator('#md figure.md-diagram svg').count()).toBe(1)
    // HTML in the Markdown is text, not markup
    expect(await page.locator('#md img').count()).toBe(0)
    expect(await page.locator('#md').innerText()).toContain('<img src=x onerror=alert(1)>')
    expect(v.errors).toEqual([])
  })
})
