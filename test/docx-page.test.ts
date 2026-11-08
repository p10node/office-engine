import { describe, expect, test } from 'bun:test'
import { Engine } from '../src/core/worker/engine'
import type { DocxMeta, PageSetResult } from '../src/core/docx/types'
import { ZipArchive } from '../src/core/zip'
import { BlobSource } from '../src/core/source'

const S = (n: string) => `${import.meta.dir}/../playground/public/samples/${n}`

async function open(src: Blob) {
  const events: { event: string; data: any }[] = []
  const engine = new Engine((event, data) => events.push({ event, data }))
  await engine.open(src, {})
  await engine.call('docx.start', [])
  const meta = events.filter((e) => e.event === 'docx:meta').at(-1)!.data as DocxMeta
  return { engine, meta }
}

const sample = async () => new Blob([await Bun.file(S('sample.docx')).arrayBuffer()])
const mainXml = async (b: Blob) => (await ZipArchive.open(new BlobSource(b))).text('word/document.xml')

describe('docx page setup', () => {
  test('A4 landscape, margins and two columns survive save + reopen', async () => {
    const { engine, meta } = await open(await sample())
    const was = meta.page
    const res = (await engine.call('docx.page', [{ w: 841.9, h: 595.3, mt: 36, ml: 54, cols: 2, colGap: 18 }])) as PageSetResult
    expect(res.page).toMatchObject({ w: 841.9, h: 595.3, mt: 36, ml: 54, mr: was.mr, cols: 2, colGap: 18 })
    expect(res.undo).toEqual({ w: was.w, h: was.h, mt: was.mt, ml: was.ml, cols: 1, colGap: 36 })

    const saved = (await engine.call('save', [])) as Blob
    const xml = await mainXml(saved)
    const tag = (name: string) => /<w:[^>]*>/.exec(xml.slice(xml.indexOf(`<w:${name} `)))?.[0] ?? ''
    expect(tag('pgSz')).toContain('w:w="16838"')
    expect(tag('pgSz')).toContain('w:h="11906"')
    expect(tag('pgSz')).toContain('w:orient="landscape"')
    expect(tag('cols')).toContain('w:num="2"')
    expect(tag('cols')).toContain('w:space="360"')
    const re = await open(saved)
    expect(re.meta.page.w).toBeCloseTo(841.9, 1)
    expect(re.meta.page.h).toBeCloseTo(595.3, 1)
    expect(re.meta.page.mt).toBe(36)
    expect(re.meta.page.ml).toBe(54)
    expect(re.meta.page.mr).toBe(was.mr)
    expect(re.meta.page.cols).toBe(2)
    expect(re.meta.page.colGap).toBe(18)
  })

  test('undoing the change gives the file back its first page setup; one column drops the count', async () => {
    const { engine, meta } = await open(await sample())
    const res = (await engine.call('docx.page', [{ w: 400, cols: 3 }])) as PageSetResult
    await engine.call('docx.page', [res.undo])
    const re = await open((await engine.call('save', [])) as Blob)
    expect(re.meta.page).toEqual(meta.page)
  })

  test('nonsense is refused', async () => {
    const { engine } = await open(await sample())
    await expect(engine.call('docx.page', [{ mt: -5 }])).rejects.toThrow('Invalid page setup')
    await expect(engine.call('docx.page', [{ w: Number.NaN }])).rejects.toThrow('Invalid page setup')
  })
})
