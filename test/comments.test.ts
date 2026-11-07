import { describe, expect, test } from 'bun:test'
import { Engine } from '../src/core/worker/engine'
import { paraSpec } from '../src/core/docx/spec'
import type { Block, Inline, Paragraph } from '../src/core/docx/types'
import type { DocComment } from '../src/core/comments'
import { threadsOf } from '../src/core/comments'
import { ZipArchive } from '../src/core/zip'
import { BlobSource } from '../src/core/source'
import { rebuildZip } from '../src/core/zipwrite'

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
const W14 = 'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"'

/** a document the way Word writes comments: a thread with a reply, a resolved one spanning two paragraphs, one without a range */
export function wordDocx(opts: { extended?: boolean } = {}): Promise<Blob> {
  const ext = opts.extended ?? true
  const files = new Map<string, string>()
  files.set(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>${ext ? '<Override PartName="/word/commentsExtended.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml"/><Override PartName="/word/commentsIds.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.commentsIds+xml"/>' : ''}</Types>`,
  )
  files.set('_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`)
  files.set(
    'word/_rels/document.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>${ext ? '<Relationship Id="rId10" Type="http://schemas.microsoft.com/office/2011/relationships/commentsExtended" Target="commentsExtended.xml"/><Relationship Id="rId11" Type="http://schemas.microsoft.com/office/2016/09/relationships/commentsIds" Target="commentsIds.xml"/>' : ''}</Relationships>`,
  )
  const ref = (id: number) => `<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="${id}"/></w:r>`
  files.set(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W} ${W14}><w:body>` +
      `<w:p w14:paraId="00000001"><w:r><w:t xml:space="preserve">Hello </w:t></w:r><w:commentRangeStart w:id="0"/><w:commentRangeStart w:id="1"/><w:r><w:t>brave new</w:t></w:r><w:commentRangeEnd w:id="0"/>${ref(0)}<w:commentRangeEnd w:id="1"/>${ref(1)}<w:r><w:t xml:space="preserve"> </w:t></w:r><w:commentRangeStart w:id="3"/><w:r><w:t>world</w:t></w:r>${ref(2)}</w:p>` +
      `<w:p w14:paraId="00000002"><w:r><w:t>Second</w:t></w:r><w:commentRangeEnd w:id="3"/>${ref(3)}<w:r><w:t xml:space="preserve"> line</w:t></w:r></w:p>` +
      `<w:p w14:paraId="00000003"><w:r><w:t>Third</w:t></w:r></w:p>` +
      `<w:sectPr/></w:body></w:document>`,
  )
  const cm = (id: number, author: string, para: string, text: string) =>
    `<w:comment w:id="${id}" w:author="${author}" w:date="2026-11-01T10:0${id}:00Z" w:initials="${author[0]}"><w:p w14:paraId="${para}" w14:textId="77777777"><w:pPr><w:pStyle w:val="CommentText"/></w:pPr><w:r><w:annotationRef/></w:r><w:r><w:t>${text}</w:t></w:r></w:p></w:comment>`
  files.set(
    'word/comments.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:comments ${W} ${W14}>${cm(0, 'Ann', 'AAAA0001', 'Is this right?')}${cm(1, 'Bob', 'AAAA0002', 'Yes')}${cm(2, 'Ann', 'AAAA0003', 'Point')}${cm(3, 'Cid', 'AAAA0004', 'Spans two')}</w:comments>`,
  )
  if (ext) {
    files.set(
      'word/commentsExtended.xml',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"><w15:commentEx w15:paraId="AAAA0001" w15:done="0"/><w15:commentEx w15:paraId="AAAA0002" w15:paraIdParent="AAAA0001" w15:done="0"/><w15:commentEx w15:paraId="AAAA0003" w15:done="0"/><w15:commentEx w15:paraId="AAAA0004" w15:done="1"/></w15:commentsEx>`,
    )
    files.set(
      'word/commentsIds.xml',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w16cid:commentsIds xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid"><w16cid:commentId w16cid:paraId="AAAA0001" w16cid:durableId="1A2B3C01"/><w16cid:commentId w16cid:paraId="AAAA0002" w16cid:durableId="1A2B3C02"/><w16cid:commentId w16cid:paraId="AAAA0003" w16cid:durableId="1A2B3C03"/><w16cid:commentId w16cid:paraId="AAAA0004" w16cid:durableId="1A2B3C04"/></w16cid:commentsIds>`,
    )
  }
  return rebuildZip(null, files as Map<string, string>, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document')
}

async function openDocx(src: Blob) {
  const events: { event: string; data: any }[] = []
  const engine = new Engine((event, data) => events.push({ event, data }))
  await engine.open(src, {})
  await engine.call('docx.start', [])
  const blocks: Block[] = events.filter((e) => e.event === 'docx:blocks').flatMap((e) => e.data)
  const list = (await engine.call('comments.list', [])) as DocComment[]
  return { engine, blocks, events, list }
}

const paras = (blocks: Block[]) => blocks.filter((b): b is Paragraph => b.t === 'p')
const textOf = (p: Paragraph) => p.runs.map((r) => (r.t === 'text' ? r.text : '￼')).join('')
/** the text a thread's marks cover, paragraph by paragraph */
const covered = (blocks: Block[], id: string) =>
  paras(blocks)
    .map((p) => p.runs.filter((r): r is Extract<Inline, { t: 'text' }> => r.t === 'text' && !!r.cm?.includes(id)).map((r) => r.text).join(''))
    .filter(Boolean)

async function partText(blob: Blob, path: string) {
  const zip = await ZipArchive.open(new BlobSource(blob))
  return zip.has(path) ? zip.text(path) : undefined
}

describe('docx comments', () => {
  test('threads, replies, resolved state and ranges are read', async () => {
    const { blocks, list } = await openDocx(await wordDocx())
    expect(list.map((c) => c.id).sort()).toEqual(['1A2B3C01', '1A2B3C02', '1A2B3C03', '1A2B3C04'])
    const th = threadsOf(list)
    expect(th.map((t) => t.root.id)).toEqual(['1A2B3C01', '1A2B3C03', '1A2B3C04'])
    expect(th[0].replies.map((r) => r.text)).toEqual(['Yes'])
    expect(th[0].root).toMatchObject({ text: 'Is this right?', author: { name: 'Ann', initials: 'A' }, time: '2026-11-01T10:00:00Z' })
    expect(th[2].root.done).toBe(true)
    // the reply marks nothing of its own: the thread's range is the root's
    expect(covered(blocks, '1A2B3C01')).toEqual(['brave new'])
    expect(covered(blocks, '1A2B3C02')).toEqual([])
    expect(covered(blocks, '1A2B3C04')).toEqual(['world', 'Second'])
    // no range: the character before the reference
    expect(covered(blocks, '1A2B3C03')).toEqual(['d'])
  })

  test('saving without changes keeps the file as it was', async () => {
    const src = await wordDocx()
    const { engine } = await openDocx(src)
    const out = (await engine.call('save', [])) as Blob
    expect(await partText(out, 'word/document.xml')).toBe(await partText(src, 'word/document.xml'))
    expect(await partText(out, 'word/comments.xml')).toBe(await partText(src, 'word/comments.xml'))
  })

  test('a reply, a resolved thread and a deleted one survive saving and reopening', async () => {
    const { engine, list } = await openDocx(await wordDocx())
    const root = list.find((c) => c.id === '1A2B3C01')!
    await engine.call('comments.set', ['ABCDEF01', { id: 'ABCDEF01', parent: root.id, author: { name: 'Dee', id: 'u-dee' }, time: '2026-11-07T09:00:00Z', text: 'Agreed\nsecond line' }])
    await engine.call('comments.set', [root.id, { ...root, done: true }])
    await engine.call('comments.set', ['1A2B3C03', null])
    const out = (await engine.call('save', [])) as Blob
    const doc = (await partText(out, 'word/document.xml'))!
    // the point comment's marks are gone, the thread's range now holds the new reply too
    expect(doc).not.toContain('w:id="2"')
    expect(doc.match(/commentRangeStart/g)?.length).toBe(4)
    const re = await openDocx(out)
    const th = threadsOf(re.list)
    expect(th.map((t) => t.root.id)).toEqual(['1A2B3C01', '1A2B3C04'])
    expect(th[0].root.done).toBe(true)
    expect(th[0].replies.map((r) => [r.id, r.text, r.author.name, r.author.id])).toEqual([
      ['1A2B3C02', 'Yes', 'Bob', undefined],
      ['ABCDEF01', 'Agreed\nsecond line', 'Dee', 'u-dee'],
    ])
    expect(covered(re.blocks, '1A2B3C01')).toEqual(['brave new'])
    expect(covered(re.blocks, '1A2B3C04')).toEqual(['world', 'Second'])
    // untouched comments keep their own XML
    expect(await partText(out, 'word/comments.xml')).toContain('<w:pStyle w:val="CommentText"/></w:pPr><w:r><w:annotationRef/></w:r><w:r><w:t>Yes</w:t>')
  })

  test('marked text stays marked through edits; a new comment on new text is written', async () => {
    const { engine, blocks } = await openDocx(await wordDocx())
    const [p1, p2] = paras(blocks)
    // type inside the first thread's range, put a new comment on "Third"…
    const runs = p1.runs.flatMap((r): Inline[] => (r.t === 'text' && r.text === 'brave new' ? [{ ...r, text: 'brave shiny new' }] : [r]))
    await engine.call('docx.edit', [{ anchor: p1.pid, list: [paraSpec([{ ...p1, runs, edited: true }])] }])
    const p3 = paras(blocks)[2]
    await engine.call('comments.set', ['0000BEEF', { id: '0000BEEF', author: { name: 'Eve' }, time: '2026-11-07T09:00:00Z', text: 'New' }])
    await engine.call('docx.edit', [{ anchor: p3.pid, list: [paraSpec([{ ...p3, runs: p3.runs.map((r) => ({ ...r, cm: ['0000BEEF'] })), edited: true }])] }])
    const out = (await engine.call('save', [])) as Blob
    const re = await openDocx(out)
    expect(covered(re.blocks, '1A2B3C01')).toEqual(['brave shiny new'])
    // the second paragraph was not touched: the range still ends there
    expect(covered(re.blocks, '1A2B3C04')).toEqual(['world', 'Second'])
    expect(covered(re.blocks, '0000BEEF')).toEqual(['Third'])
    expect(threadsOf(re.list).find((t) => t.root.id === '0000BEEF')?.root.text).toBe('New')
    expect(textOf(paras(re.blocks)[1])).toBe('Second line')
    void p2
  })

  test('a thread whose text is deleted leaves the file', async () => {
    const { engine, blocks } = await openDocx(await wordDocx())
    const p1 = paras(blocks)[0]
    const runs = p1.runs.filter((r) => !(r.t === 'text' && r.cm?.includes('1A2B3C01')))
    await engine.call('docx.edit', [{ anchor: p1.pid, list: [paraSpec([{ ...p1, runs, edited: true }])] }])
    const out = (await engine.call('save', [])) as Blob
    const re = await openDocx(out)
    expect(threadsOf(re.list).map((t) => t.root.id)).toEqual(['1A2B3C03', '1A2B3C04'])
    expect(await partText(out, 'word/comments.xml')).not.toContain('Is this right?')
  })

  test('documents without thread parts get them; saving twice keeps the new ids', async () => {
    const { engine, list } = await openDocx(await wordDocx({ extended: false }))
    // no durable ids in the file: derived from w:id, the same on every read
    expect(list.map((c) => c.id).sort()).toEqual(['10000000', '10000001', '10000002', '10000003'])
    await engine.call('comments.set', ['77777777', { id: '77777777', parent: '10000000', author: { name: 'Fay' }, time: '2026-11-07T09:00:00Z', text: 'Reply' }])
    await engine.call('save', [])
    const out = (await engine.call('save', [])) as Blob
    const re = await openDocx(out)
    const th = threadsOf(re.list)
    expect(th.find((t) => t.root.id === '10000000')?.replies.map((r) => r.id)).toEqual(['77777777'])
    expect((await partText(out, 'word/document.xml'))!.match(/commentReference/g)?.length).toBe(5)
  })
})

// ---------------------------------------------------------------- xlsx

import { rowsToXlsx } from '../src/core/convert'

async function openXlsx(src: Blob) {
  const events: { event: string; data: any }[] = []
  const engine = new Engine((event, data) => events.push({ event, data }))
  const info = await engine.open(src, {})
  const load = async (sheet: number) => {
    await engine.call('xlsx.load', [sheet])
    while (!events.some((e) => e.event === 'xlsx:progress' && e.data.sheet === sheet && e.data.done)) await new Promise((r) => setTimeout(r, 5))
  }
  await load(0)
  const list = (await engine.call('comments.list', [])) as DocComment[]
  return { engine, info: info as any, list, events, load }
}

/** a workbook with a note on B2 (legacy comment only, as Excel writes notes) */
async function withNote(): Promise<Blob> {
  const base = await rowsToXlsx([['a', 'b'], ['c', 'd'], ['e', 'f']], 'Data')
  const zip = await ZipArchive.open(new BlobSource(base))
  const files = new Map<string, string | Uint8Array>()
  const sheetRels = 'xl/worksheets/_rels/sheet1.xml.rels'
  files.set(sheetRels, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing1.vml"/></Relationships>`)
  files.set('xl/comments1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><authors><author>Zed</author></authors><commentList><comment ref="B2" authorId="0" shapeId="0"><text><r><rPr><b/></rPr><t>Zed:</t></r><r><t xml:space="preserve">\nCheck me</t></r></text></comment></commentList></comments>`)
  files.set('xl/drawings/vmlDrawing1.vml', `<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel"><v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"/><v:shape id="_x0000_s1025" type="#_x0000_t202"><x:ClientData ObjectType="Note"><x:Row>1</x:Row><x:Column>1</x:Column></x:ClientData></v:shape></xml>`)
  const sheet = (await zip.text('xl/worksheets/sheet1.xml')).replace('</worksheet>', '<legacyDrawing r:id="rId2"/></worksheet>').replace('<worksheet ', '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ')
  files.set('xl/worksheets/sheet1.xml', sheet)
  const ct = (await zip.text('[Content_Types].xml')).replace('</Types>', '<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/><Override PartName="/xl/comments1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/></Types>')
  files.set('[Content_Types].xml', ct)
  return rebuildZip(zip, files as Map<string, string>, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
}

describe('xlsx comments', () => {
  test('a new thread with a reply and a mention survives saving and reopening', async () => {
    const { engine, info } = await openXlsx(await rowsToXlsx([['Name', 'Qty'], ['Pen', '3']], 'Stock'))
    const sheet = info.xlsx.sheets[0].path
    await engine.call('comments.set', ['{11111111-1111-4111-8111-111111111111}', { id: '{11111111-1111-4111-8111-111111111111}', author: { name: 'Ann', id: 'u1' }, time: '2026-11-07T09:00:00Z', text: 'Too few @Bob', mentions: [{ name: 'Bob', id: 'u2' }], at: { sheet, r: 1, c: 1 } }])
    await engine.call('comments.set', ['{22222222-2222-4222-8222-222222222222}', { id: '{22222222-2222-4222-8222-222222222222}', parent: '{11111111-1111-4111-8111-111111111111}', author: { name: 'Bob', id: 'u2' }, time: '2026-11-07T09:05:00Z', text: 'Ordered more' }])
    const out = (await engine.call('save', [])) as Blob
    const sheetXml = (await partText(out, 'xl/worksheets/sheet1.xml'))!
    expect(sheetXml).toMatch(/<legacyDrawing r:id="rId\d+"\/><\/worksheet>/)
    expect(await partText(out, 'xl/comments1.xml')).toContain('[Threaded comment]')
    expect(await partText(out, 'xl/drawings/vmlDrawing1.vml')).toContain('<x:Row>1</x:Row><x:Column>1</x:Column>')
    const re = await openXlsx(out)
    const th = threadsOf(re.list)
    expect(th.length).toBe(1)
    expect(th[0].root).toMatchObject({ text: 'Too few @Bob', author: { name: 'Ann', id: 'u1' }, at: { sheet, r: 1, c: 1 }, time: '2026-11-07T09:00:00Z', mentions: [{ name: 'Bob', id: 'u2' }] })
    expect(th[0].replies.map((r) => r.text)).toEqual(['Ordered more'])
    // saving a cell edit after it keeps the link to the comments' drawing
    await re.engine.call('comments.set', [th[0].root.id, { ...th[0].root, done: true }])
    await re.engine.call('xlsx.set', [0, [{ r: 0, c: 0, input: 'Item' }]])
    const again = (await re.engine.call('save', [])) as Blob
    expect(await partText(again, 'xl/worksheets/sheet1.xml')).toMatch(/<legacyDrawing r:id="rId\d+"\/>/)
    expect((await openXlsx(again)).list.find((c) => !c.parent)?.done).toBe(true)
  })

  test('notes are read as threads; one left alone stays a note', async () => {
    const src = await withNote()
    const { engine, list, info } = await openXlsx(src)
    expect(list.length).toBe(1)
    expect(list[0]).toMatchObject({ author: { name: 'Zed' }, text: 'Check me', at: { sheet: info.xlsx.sheets[0].path, r: 1, c: 1 } })
    // a new comment elsewhere: the note is written back as it was
    await engine.call('comments.set', ['{33333333-3333-4333-8333-333333333333}', { id: '{33333333-3333-4333-8333-333333333333}', author: { name: 'Ann' }, time: '2026-11-07T09:00:00Z', text: 'New', at: { sheet: info.xlsx.sheets[0].path, r: 0, c: 0 } }])
    const out = (await engine.call('save', [])) as Blob
    const legacy = (await partText(out, 'xl/comments1.xml'))!
    expect(legacy).toContain('<t>Zed:</t>')
    expect(legacy.match(/<comment /g)?.length).toBe(2)
    const re = await openXlsx(out)
    expect(re.list.map((c) => c.text).sort()).toEqual(['Check me', 'New'])
    // the note keeps its id: answering it makes it a thread with that id
    const note = re.list.find((c) => c.text === 'Check me')!
    expect(note.id).toBe(list[0].id)
    await re.engine.call('comments.set', ['{44444444-4444-4444-8444-444444444444}', { id: '{44444444-4444-4444-8444-444444444444}', parent: note.id, author: { name: 'Ann' }, time: '2026-11-07T10:00:00Z', text: 'Done' }])
    const th = threadsOf((await openXlsx((await re.engine.call('save', [])) as Blob)).list)
    expect(th.find((t) => t.root.id === note.id)?.replies.map((r) => r.text)).toEqual(['Done'])
  })

  test('comments move with inserted / deleted rows, go with their row, come back on undo', async () => {
    const { engine, list, events } = await openXlsx(await withNote())
    const id = list[0].id
    await engine.call('xlsx.sheetOp', [0, { op: 'insertRows', at: 0, n: 2 }])
    let cur = ((await engine.call('comments.list', [])) as DocComment[])[0]
    expect(cur.at).toMatchObject({ r: 3, c: 1 })
    expect(events.filter((e) => e.event === 'comment').at(-1)?.data.id).toBe(id)
    await engine.call('xlsx.sheetOp', [0, { op: 'deleteCols', at: 1, n: 1 }])
    expect(((await engine.call('comments.list', [])) as DocComment[]).length).toBe(0)
    await engine.call('xlsx.undo', [])
    cur = ((await engine.call('comments.list', [])) as DocComment[])[0]
    expect(cur).toMatchObject({ id, at: { r: 3, c: 1 } })
    await engine.call('xlsx.undo', [])
    expect(((await engine.call('comments.list', [])) as DocComment[])[0].at).toMatchObject({ r: 1, c: 1 })
  })

  test('the last comment of a sheet takes its parts along', async () => {
    const { engine, list } = await openXlsx(await withNote())
    await engine.call('comments.set', [list[0].id, null])
    const out = (await engine.call('save', [])) as Blob
    expect(await partText(out, 'xl/comments1.xml')).toBeUndefined()
    expect(await partText(out, 'xl/worksheets/sheet1.xml')).not.toContain('legacyDrawing')
    expect((await openXlsx(out)).list).toEqual([])
  })
})

// ---------------------------------------------------------------- pptx

import { blankPptx } from '../src/core/convert'

async function openPptx(src: Blob) {
  const events: { event: string; data: any }[] = []
  const engine = new Engine((event, data) => events.push({ event, data }))
  const info = (await engine.open(src, {})) as any
  const list = (await engine.call('comments.list', [])) as DocComment[]
  return { engine, info, list }
}

describe('pptx comments', () => {
  test('modern comments: a thread with a reply, resolved, survives saving; a copy of the slide has none', async () => {
    const { engine, info } = await openPptx(await blankPptx())
    const slide = info.pptx.slides[0].path
    const root = { id: '{55555555-5555-4555-8555-555555555555}', author: { name: 'Ann', id: 'u1' }, time: '2026-11-07T09:00:00Z', text: 'Bigger title?\nMaybe', at: { slide, x: 120, y: 80 } }
    await engine.call('comments.set', [root.id, root])
    await engine.call('comments.set', ['{66666666-6666-4666-8666-666666666666}', { id: '{66666666-6666-4666-8666-666666666666}', parent: root.id, author: { name: 'Bob' }, time: '2026-11-07T09:01:00Z', text: 'Sure' }])
    await engine.call('comments.set', [root.id, { ...root, done: true }])
    await engine.call('pptx.slides', [{ op: 'duplicate', index: 0 }])
    const out = (await engine.call('save', [])) as Blob
    const slideXml = (await partText(out, slide))!
    expect(slideXml).toContain('uri="{6950BFC3-D8DA-4A85-94F7-54DA5524770B}"')
    expect(slideXml).toMatch(/<p188:commentRel xmlns:p188="[^"]+" r:id="rId\d+"\/>/)
    const re = await openPptx(out)
    expect(re.info.pptx.slides.length).toBe(2)
    const th = threadsOf(re.list)
    expect(th.length).toBe(1)
    expect(th[0].root).toMatchObject({ id: root.id, text: 'Bigger title?\nMaybe', done: true, author: { name: 'Ann', id: 'u1' }, at: { slide, x: 120, y: 80 }, time: '2026-11-07T09:00:00Z' })
    expect(th[0].replies.map((r) => [r.text, r.author.name])).toEqual([['Sure', 'Bob']])
    expect(await partText(out, 'ppt/authors.xml')).toContain('name="Bob"')
    // the copy: no comment part of its own, no link to the original's
    expect(await partText(out, re.info.pptx.slides[1].path)).not.toContain('6950BFC3')
  })

  test('deleting a slide deletes its comments; the last comment takes the link along', async () => {
    const { engine, info } = await openPptx(await blankPptx())
    await engine.call('pptx.slides', [{ op: 'add', after: 0 }])
    const meta = (await engine.call('pptx.slides', [{ op: 'add', after: 1 }])) as any
    const [s0, s1] = meta.slides.map((s: { path: string }) => s.path)
    const c = (id: string, slide: string) => ({ id, author: { name: 'Ann' }, time: '2026-11-07T09:00:00Z', text: id, at: { slide, x: 10, y: 10 } })
    await engine.call('comments.set', ['{77777777-7777-4777-8777-777777777777}', c('{77777777-7777-4777-8777-777777777777}', s0)])
    await engine.call('comments.set', ['{88888888-8888-4888-8888-888888888888}', c('{88888888-8888-4888-8888-888888888888}', s1)])
    const first = (await engine.call('save', [])) as Blob
    const re = await openPptx(first)
    expect(re.list.length).toBe(2)
    await re.engine.call('pptx.slides', [{ op: 'delete', index: 1 }])
    await re.engine.call('comments.set', ['{77777777-7777-4777-8777-777777777777}', null])
    const out = (await re.engine.call('save', [])) as Blob
    const again = await openPptx(out)
    expect(again.list).toEqual([])
    expect(await partText(out, s0)).not.toContain('6950BFC3')
    void info
  })
})

describe('mentions', () => {
  test('docx and pptx keep who was mentioned through saving (the people list), xlsx natively', async () => {
    const ann = { name: 'Ann Lee', id: 'u1' }
    const bob = { name: 'Bob', id: 'u2' }
    // docx
    const d = await openDocx(await wordDocx())
    const root = d.list.find((c) => c.id === '1A2B3C01')!
    await d.engine.call('comments.set', ['0000ABCD', { id: '0000ABCD', parent: root.id, author: ann, time: '2026-11-07T09:00:00Z', text: 'cc @Bob and @Bobby', mentions: [bob] }])
    const dr = await openDocx((await d.engine.call('save', [])) as Blob)
    expect(dr.list.find((c) => c.id === '0000ABCD')).toMatchObject({ author: ann, mentions: [bob] })
    // pptx
    const p = await openPptx(await blankPptx())
    const slide = p.info.pptx.slides[0].path
    await p.engine.call('comments.set', ['{99999999-9999-4999-8999-999999999999}', { id: '{99999999-9999-4999-8999-999999999999}', author: ann, time: '2026-11-07T09:00:00Z', text: '@Bob look', mentions: [bob], at: { slide, x: 1, y: 1 } }])
    const pr = await openPptx((await p.engine.call('save', [])) as Blob)
    expect(pr.list[0]).toMatchObject({ author: ann, mentions: [bob] })
  })
})
