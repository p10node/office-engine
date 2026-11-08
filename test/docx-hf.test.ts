import { describe, expect, test } from 'bun:test'
import { Engine } from '../src/core/worker/engine'
import type { Block, DocxMeta, HFSetResult } from '../src/core/docx/types'
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

/** a paragraph's text with fields as {PAGE} / {NUMPAGES} */
const lineOf = (b: Block | undefined) => (b?.t === 'p' ? b.runs.map((r) => (r.t === 'text' ? r.text : r.t === 'field' ? `{${r.kind}}` : '')).join('') : '')

async function blank(): Promise<Blob> {
  // the sample without its header / footer references and parts
  const zip = await ZipArchive.open(new BlobSource(Bun.file(S('sample.docx')) as unknown as Blob))
  const { rebuildZip } = await import('../src/core/zipwrite')
  const doc = (await zip.text('word/document.xml')).replace(/<w:(header|footer)Reference [^>]*\/>/g, '')
  return rebuildZip(zip, new Map([['word/document.xml', doc]]), 'application/vnd.openxmlformats-officedocument.wordprocessingml.document')
}

describe('docx headers / footers', () => {
  test('existing footer replaced by "Page {PAGE} of {NUMPAGES}", centred; survives save + reopen', async () => {
    const { engine, meta } = await open(new Blob([await Bun.file(S('sample.docx')).arrayBuffer()]))
    expect(meta.footer?.length).toBeGreaterThan(0)
    const res = (await engine.call('docx.hf', [{ parts: { footer: [{ text: 'Page {PAGE} of {NUMPAGES}', align: 'center', b: true, sz: 9 }] } }])) as HFSetResult
    expect(lineOf(res.headers.footer?.[0])).toBe('Page {PAGE} of {NUMPAGES}')
    expect(res.undo.raw?.footer).toContain('<w:ftr')
    const re = await open((await engine.call('save', [])) as Blob)
    const p = re.meta.footer![0]
    expect(lineOf(p)).toBe('Page {PAGE} of {NUMPAGES}')
    expect(p.t === 'p' && p.ps.align).toBe('center')
    expect(p.t === 'p' && p.runs[0].t === 'text' && p.runs[0].s.b && p.runs[0].s.sz).toBe(9)
    // the header was left alone
    expect(lineOf(re.meta.header?.[0])).toBe(lineOf(meta.header?.[0]))
  })

  test('a document without headers gets new parts, references and "different first page"', async () => {
    const { engine, meta } = await open(await blank())
    expect(meta.header).toBeUndefined()
    await engine.call('docx.hf', [{ parts: { header: [{ text: 'Công ty ABC', align: 'right' }], footer: [{ text: '{PAGE}', align: 'center' }], firstHeader: [{ text: 'Trang bìa' }] }, titlePg: true }])
    const saved = (await engine.call('save', [])) as Blob
    const zip = await ZipArchive.open(new BlobSource(saved))
    const doc = await zip.text('word/document.xml')
    expect(doc).toMatch(/<w:headerReference w:type="default" r:id="rId\d+"\/>/)
    expect(doc).toMatch(/<w:headerReference w:type="first" r:id="rId\d+"\/>/)
    expect(doc).toContain('<w:titlePg/>')
    expect(await zip.text('[Content_Types].xml')).toContain('wordprocessingml.footer+xml')
    const re = await open(saved)
    expect(re.meta.titlePg).toBe(true)
    expect(lineOf(re.meta.header?.[0])).toBe('Công ty ABC')
    expect(lineOf(re.meta.firstHeader?.[0])).toBe('Trang bìa')
    expect(lineOf(re.meta.footer?.[0])).toBe('{PAGE}')
  })

  test('undo gives back the part as it was', async () => {
    const { engine, meta } = await open(new Blob([await Bun.file(S('sample.docx')).arrayBuffer()]))
    const res = (await engine.call('docx.hf', [{ parts: { header: [{ text: 'x' }] }, titlePg: true }])) as HFSetResult
    const back = (await engine.call('docx.hf', [res.undo])) as HFSetResult
    expect(back.titlePg).toBe(false)
    expect(lineOf(back.headers.header?.[0])).toBe(lineOf(meta.header?.[0]))
  })

  test('nonsense is refused', async () => {
    const { engine } = await open(new Blob([await Bun.file(S('sample.docx')).arrayBuffer()]))
    await expect(engine.call('docx.hf', [{ parts: { sidebar: [] } }])).rejects.toThrow('Invalid header / footer')
    await expect(engine.call('docx.hf', [{ parts: { header: [{ text: 'a', sz: 500 }] } }])).rejects.toThrow('Invalid header / footer')
  })
})
