import { describe, expect, test } from 'bun:test'
import { useBrowser, viewer } from './harness'

useBrowser()

describe('embedding without React', () => {
  test('<office-viewer>: shadow DOM, attributes, load event', async () => {
    const v = await viewer({}, { page: 'embed.html' })
    const wc = v.page.locator('#wc')
    await wc.locator('.de-page-inner').first().waitFor()
    // rendered inside the element's shadow root, styles included there
    expect(await v.page.evaluate(() => !!document.getElementById('wc')!.shadowRoot?.querySelector('style#office-engine-styles'))).toBe(true)
    expect(await v.page.evaluate(() => (window as any).__wcEvents)).toEqual([{ kind: 'docx', name: 'sample.docx' }])
    // save="false" file-menu="false"
    expect(await wc.locator('.de-file').count()).toBe(0)
    await wc.locator('.de-bar button.de-mode', { hasText: 'Edit' }).click()
    expect(await wc.locator('.de-bar button.de-mode', { hasText: /^(Save|Download)$/ }).count()).toBe(0)
    // attributes are live
    await v.page.evaluate(() => document.getElementById('wc')!.setAttribute('toolbar', 'false'))
    await wc.locator('.de-bar').waitFor({ state: 'detached' })
    await v.page.evaluate(() => document.getElementById('wc')!.setAttribute('src', 'samples/sample.pptx'))
    await wc.locator('.de-slide-inner').first().waitFor()
    expect(v.errors).toEqual([])
  })

  test('<office-viewer>: file property', async () => {
    const v = await viewer({}, { page: 'embed.html' })
    await v.page.evaluate(async () => {
      const wc = document.getElementById('wc') as HTMLElement & { file?: Blob }
      wc.file = await (await fetch('samples/sample.xlsx')).blob()
    })
    await v.page.locator('#wc .de-grid').waitFor()
  })

  test('mount(): options, update(), destroy()', async () => {
    const v = await viewer({}, { page: 'embed.html' })
    const host = v.page.locator('#mounted')
    await host.locator('.de-grid').waitFor()
    expect(await host.locator('.de-bar').count()).toBe(0)
    await v.page.evaluate(() => (window as any).__mounted.update({ toolbar: true }))
    await host.locator('.de-bar').waitFor()
    expect(await host.locator('.de-title').innerText()).toBe('sample.xlsx')
    await v.page.evaluate(() => (window as any).__mounted.destroy())
    expect(await host.innerHTML()).toBe('')
  })
})
