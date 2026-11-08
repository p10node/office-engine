import { describe, expect, test } from 'bun:test'
import { Engine } from '../src/core/worker/engine'
import { paraSpec } from '../src/core/docx/spec'
import type { Block, Inline, Paragraph } from '../src/core/docx/types'
import { DocumentClient } from '../src/core/worker/client'
import { DocxEditor, type DocState } from '../src/react/docx/DocxEditor'
import { opBackspace, opInsert } from '../src/react/edit/input'
import { ZipArchive } from '../src/core/zip'
import { BlobSource } from '../src/core/source'

const S = (n: string) => `${import.meta.dir}/../playground/public/samples/${n}`
const file = async (n: string) => new Blob([await Bun.file(S(n)).arrayBuffer()])

async function open(src: Blob) {
  const events: { event: string; data: any }[] = []
  const engine = new Engine((event, data) => events.push({ event, data }))
  await engine.open(src, {})
  await engine.call('docx.start', [])
  const blocks: Block[] = events.filter((e) => e.event === 'docx:blocks').flatMap((e) => e.data)
  return { engine, blocks }
}

/** text with changes as [+ins] / [-del] */
const show = (p: Paragraph) => p.runs.map((r) => (r.t !== 'text' ? '￼' : r.tc ? `[${r.tc.k === 'ins' ? '+' : '-'}${r.text}]` : r.text)).join('')
const paras = (blocks: Block[]) => blocks.filter((b): b is Paragraph => b.t === 'p')
const find = (blocks: Block[], start: string) => paras(blocks).find((p) => show(p).startsWith(start))!

describe('tracked changes in a Word file', () => {
  test('read: insertions and deletions with their authors', async () => {
    const { blocks } = await open(await file('tracked.docx'))
    const p = find(blocks, 'Revenue')
    expect(show(p)).toBe('Revenue grew [-24%][+31%] year over year')
    const del = p.runs.find((r) => r.t === 'text' && r.tc?.k === 'del') as Inline & { t: 'text' }
    expect(del.tc).toMatchObject({ author: 'Lan Nguyen', date: '2026-11-01T09:00:00Z' })
    expect(show(find(blocks, 'Two new'))).toBe('Two new regions launched[+ in Q3]')
  })

  test('a paragraph with changes keeps them when it is edited elsewhere and saved', async () => {
    const { engine, blocks } = await open(await file('tracked.docx'))
    const p = find(blocks, 'Revenue')
    const runs = [...p.runs]
    const last = runs[runs.length - 1] as Inline & { t: 'text' }
    runs[runs.length - 1] = { ...last, text: last.text + ' (audited)' }
    await engine.call('docx.edit', [{ anchor: p.pid, list: [paraSpec([{ ...p, runs, edited: true }])] }])
    const saved = (await engine.call('save', [])) as Blob
    const xml = await (await ZipArchive.open(new BlobSource(saved))).text('word/document.xml')
    expect(xml).toMatch(/<w:del w:id="\d+" w:author="Lan Nguyen" w:date="2026-11-01T09:00:00Z"><w:r><w:delText xml:space="preserve">24%<\/w:delText><\/w:r><\/w:del>/)
    const re = await open(saved)
    expect(show(find(re.blocks, 'Revenue'))).toBe('Revenue grew [-24%][+31%] year over year (audited)')
  })
})

describe('reviewing and suggesting in the editor', () => {
  async function editor(src: Blob) {
    const client = await DocumentClient.open(src, { worker: false })
    const st: DocState = { blocks: [], heights: new WeakMap(), dirty: new Set(), frontier: 0 }
    client.on('docx:blocks', (b) => st.blocks.push(...b))
    const done = new Promise<void>((r) => client.on('docx:done', () => r()))
    void client.call('docx.start')
    await done
    const ed = new DocxEditor(client, st, () => {})
    const at = (start: string) => String(st.blocks.findIndex((b) => b.t === 'p' && show(b).startsWith(start)))
    const settle = () => new Promise((r) => setTimeout(r, 30))
    const reopen = async () => (await open((await client.call('save')) as Blob)).blocks
    return { client, st, ed, at, settle, reopen }
  }

  test('accept the change at the caret, reject all the others; undo brings them back', async () => {
    const { ed, st, at, settle, reopen, client } = await editor(await file('tracked.docx'))
    expect(ed.changeCount()).toBe(5)
    const pa = at('Revenue')
    ed.input.lastSel = { a: { pa, o: 14 }, f: { pa, o: 14 } } // inside "24%" (deleted)
    expect(ed.review(true)).toBe(1)
    expect(show(st.blocks[Number(pa)] as Paragraph)).toBe('Revenue grew [+31%] year over year')
    expect(ed.review(false, true)).toBe(4)
    await settle()
    expect(ed.changeCount()).toBe(0)
    const saved = await reopen()
    expect(show(find(saved, 'Revenue'))).toBe('Revenue grew  year over year')
    expect(show(find(saved, 'Two new'))).toBe('Two new regions launched')
    expect(show(find(saved, 'Expand'))).toBe('Expand to APAC')
    for (let i = 0; i < 6; i++) ed.undo()
    await settle()
    expect(ed.changeCount()).toBe(5)
    client.destroy()
  })

  test('typing and Backspace while tracking are signed by the author and saved as Word changes', async () => {
    const { ed, st, settle, reopen, client } = await editor(await file('sample.docx'))
    ed.author = () => 'Vy'
    ed.setTracking(true)
    const pa = String(st.blocks.findIndex((b) => b.t === 'p' && show(b).startsWith('Hire the platform team')))
    const end = 'Hire the platform team'.length
    opInsert(ed, { a: { pa, o: end }, f: { pa, o: end } }, ' this month')
    ed.input.lastSel = { a: { pa, o: 4 }, f: { pa, o: 4 } }
    opBackspace(ed, { a: { pa, o: 4 }, f: { pa, o: 4 } })
    await settle()
    expect(show(st.blocks[Number(pa)] as Paragraph)).toBe('Hir[-e] the platform team[+ this month]')
    const ins = (st.blocks[Number(pa)] as Paragraph).runs.find((r) => r.t === 'text' && r.tc?.k === 'ins') as Inline & { t: 'text' }
    expect(ins.tc?.author).toBe('Vy')
    const saved = await reopen()
    expect(show(find(saved, 'Hir'))).toBe('Hir[-e] the platform team[+ this month]')
    // tracking off: plain edits again
    ed.setTracking(false)
    opInsert(ed, { a: { pa, o: 0 }, f: { pa, o: 0 } }, '» ')
    await settle()
    expect(show(st.blocks[Number(pa)] as Paragraph).startsWith('» Hir')).toBe(true)
    client.destroy()
  })

  test('next / previous change select the changes in order', async () => {
    const { ed, at, client } = await editor(await file('tracked.docx'))
    ed.input.lastSel = { a: { pa: '0', o: 0 }, f: { pa: '0', o: 0 } }
    expect(ed.gotoChange(1)).toBe(true)
    const rev = at('Revenue')
    expect(ed.input.lastSel).toEqual({ a: { pa: rev, o: 13 }, f: { pa: rev, o: 16 } })
    expect(ed.gotoChange(1)).toBe(true)
    expect(ed.input.lastSel).toEqual({ a: { pa: rev, o: 16 }, f: { pa: rev, o: 19 } })
    expect(ed.gotoChange(-1)).toBe(true)
    expect(ed.input.lastSel).toEqual({ a: { pa: rev, o: 13 }, f: { pa: rev, o: 16 } })
    client.destroy()
  })
})
