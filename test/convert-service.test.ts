import { afterAll, describe, expect, test } from 'bun:test'
import { existsSync, mkdtempSync, readFileSync, rmSync, symlinkSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { gzipSync } from 'node:zlib'
import { BusyError, convert, ConvertError, Converter, createHandler, findFonts, pageRanges, TimeoutError, withGenerics } from '../packages/convert/src/index'
import { testFont } from './font-builder'
import { rowsToXlsx } from '../src/core/convert'
import { Engine } from '../src/core/worker/engine'
import { ZipArchive } from '../src/core/zip'
import { BlobSource } from '../src/core/source'

const S = (n: string) => new Uint8Array(readFileSync(`${import.meta.dir}/../playground/public/samples/${n}`))
const text = (d: Uint8Array) => new TextDecoder().decode(d)

/** a CSV that takes the engine a few seconds */
const bigCsv = () => {
  const rows: string[] = ['a,b,c,d']
  for (let i = 0; i < 600_000; i++) rows.push(`${i},name ${i},${i * 1.5},x`)
  return new TextEncoder().encode(rows.join('\n'))
}

describe('convert', () => {
  test('legacy and other formats to the Office file they open as', async () => {
    for (const [file, ext, from] of [
      ['sample.doc', 'docx', 'doc'],
      ['sample.xls', 'xlsx', 'xls'],
      ['sample.ppt', 'pptx', 'ppt'],
      ['sample.odt', 'docx', 'odt'],
      ['sample.key', 'pptx', 'keynote'],
      ['sample.md', 'docx', 'markdown'],
      ['sample.csv', 'xlsx', 'csv'],
    ] as const) {
      const r = await convert(S(file), { name: file })
      expect(r.name).toBe(file.replace(/\.\w+$/, `.${ext}`))
      expect(r.kind).toBe(ext)
      expect(r.from).toBe(from)
      // a real package of that kind: it opens again
      const again = await convert(r.data, { name: r.name })
      expect(again.kind).toBe(ext)
      expect(again.from).toBeUndefined()
    }
  })

  test('documents as text, Markdown and HTML', async () => {
    const md = text((await convert(S('sample.doc'), { name: 'sample.doc', to: 'md' })).data)
    expect(md).toContain('# Highlights')
    expect(md).toContain('[example.com](https://example.com/)')
    const html = await convert(S('sample.md'), { name: 'notes.md', to: 'html' })
    expect(html.type).toStartWith('text/html')
    expect(text(html.data)).toContain('<title>notes</title>')
    expect(text((await convert(S('sample.pptx'), { name: 'deck.pptx', to: 'txt' })).data)).toStartWith('Slide 1\n')
  })

  test('a workbook as CSV: formulas the file has no value for are computed', async () => {
    const r = await convert(S('sample.xlsx'), { name: 'sales.xlsx', to: 'csv' })
    expect(r.name).toBe('sales.csv')
    const rows = text(r.data).replace(/^﻿/, '').split('\r\n')
    expect(rows[2]).toStartWith('North,2026-01-16,195,$21.20,"4,134.00"')
    const tsv = await convert(S('sample.xlsx'), { name: 'sales.xlsx', to: 'tsv', sheet: 0 })
    expect(text(tsv.data)).toContain('North\t2026-01-16')
  })

  test('formats a kind cannot be written as, and PDF of documents and workbooks, are refused', async () => {
    const e = await convert(S('sample.pptx'), { to: 'csv' }).catch((e) => e)
    expect(e).toBeInstanceOf(ConvertError)
    expect(e.code).toBe('unsupported')
    const pdf = await convert(S('sample.docx'), { to: 'pdf' }).catch((e) => e)
    expect(pdf.code).toBe('not-implemented')
    expect(pdf.message).toContain('#48')
    expect(await convert(S('sample.xlsx'), { to: 'pdf' }).catch((e) => e.message)).toContain('#47')
    expect(await convert(S('sample.xlsx'), { to: 'csv', sheet: 'No such sheet' }).catch((e) => e.code)).toBe('unsupported')
    expect(await convert(S('sample.xlsx'), { to: 'csv', sheet: 0.5 }).catch((e) => e.code)).toBe('unsupported')
  })

  test('what is written, and what a gzip-compressed file unpacks to, is limited', async () => {
    expect(await convert(S('sample.xlsx'), { to: 'csv', maxOutput: 100 }).catch((e) => e.code)).toBe('too-large')
    // a blank cell formatted far away is no text: the CSV is as small as the cells in use
    const engine = new Engine(() => {})
    await engine.open(await rowsToXlsx([['x']], 'S'), {})
    await engine.call('xlsx.format', [0, { r0: 1048575, c0: 16383, r1: 1048575, c1: 16383 }, { b: true }])
    const sparse = new Uint8Array(await ((await engine.call('save', [])) as Blob).arrayBuffer())
    expect(text((await convert(sparse, { to: 'csv' })).data)).toBe('x\r\n')
    expect(await convert(S('sample.docx'), { to: 'html', maxOutput: 1000 }).catch((e) => e.code)).toBe('too-large')
    const gz = new Uint8Array(gzipSync(new Uint8Array(4 * 1024 * 1024).fill(0x41)))
    expect((await convert(gz, { name: 'a.csv.gz', maxUnpacked: 1024 * 1024 }).catch((e) => e)).message).toContain('too large')
    expect((await convert(gz, { name: 'a.csv.gz' })).kind).toBe('xlsx')
  })
})

describe('Converter', () => {
  const pool = new Converter({ workers: 2, timeout: 1500 })
  afterAll(() => pool.close())

  test('converts on worker threads', async () => {
    const [a, b, c] = await Promise.all([pool.convert(S('sample.doc'), { name: 'a.doc' }), pool.convert(S('sample.xls'), { name: 'b.xls', to: 'csv' }), pool.convert(S('sample.odp'), { name: 'c.odp' })])
    expect([a.name, b.name, c.name]).toEqual(['a.docx', 'b.csv', 'c.pptx'])
    expect(a.data).toBeInstanceOf(Uint8Array)
    expect((await convert(a.data, { name: a.name })).kind).toBe('docx')
    // errors keep their kind across the thread
    expect(await pool.convert(S('sample.pptx'), { to: 'csv' }).catch((e) => e)).toBeInstanceOf(ConvertError)
  })

  test('a conversion past the time limit ends its thread; the next one gets a new thread', async () => {
    const e = await pool.convert(bigCsv(), { name: 'big.csv' }).catch((e) => e)
    expect(e).toBeInstanceOf(TimeoutError)
    expect(pool.load.running).toBe(0)
    expect((await pool.convert(S('sample.doc'), { name: 'a.doc' })).name).toBe('a.docx')
  })

  test('a caller giving up ends its conversion', async () => {
    const ac = new AbortController()
    const p = pool.convert(bigCsv(), { name: 'big.csv', signal: ac.signal })
    setTimeout(() => ac.abort(new Error('gave up')), 100)
    expect((await p.catch((e) => e)).message).toBe('gave up')
    expect(pool.load.running).toBe(0)
  })

  test('a Node Buffer sharing its memory is converted, and left as it was', async () => {
    const doc = S('sample.doc')
    const shared = Buffer.alloc(doc.length + 64)
    shared.set(doc, 32)
    const view = shared.subarray(32, 32 + doc.length)
    expect((await pool.convert(view, { name: 'a.doc' })).name).toBe('a.docx')
    expect(Buffer.compare(view, doc)).toBe(0)
  })

  test('a full queue refuses more; so does more input than the byte budget', async () => {
    const budget = new Converter({ workers: 1, bytes: 1024, timeout: 0 })
    try {
      const ac = new AbortController()
      const first = budget.convert(bigCsv(), { signal: ac.signal }).catch(() => {})
      expect(await budget.convert(S('sample.doc')).catch((e) => e)).toBeInstanceOf(BusyError)
      ac.abort()
      await first
      // nothing running: one file is taken, whatever its size
      expect((await budget.convert(S('sample.doc'))).kind).toBe('docx')
    } finally {
      await budget.close()
    }
  })

  test('a full queue refuses more', async () => {
    const small = new Converter({ workers: 1, queue: 1, timeout: 0 })
    try {
      const ac = new AbortController()
      const first = small.convert(bigCsv(), { signal: ac.signal }).catch(() => {})
      const second = small.convert(S('sample.doc')).catch(() => {})
      expect(await small.convert(S('sample.doc')).catch((e) => e)).toBeInstanceOf(BusyError)
      ac.abort()
      await Promise.all([first, second])
    } finally {
      await small.close()
    }
  })
})

describe('HTTP service', () => {
  // (no font files: text is set in Helvetica, the same on every machine)
  const converter = new Converter({ workers: 0, fonts: [] })
  const handle = createHandler(converter, { maxBytes: 1024 * 1024, maxFiles: 3, version: '9.9.9' })
  const form = (files: [string, Uint8Array][], fields: Record<string, string> = {}) => {
    const fd = new FormData()
    for (const [name, data] of files) fd.append('files', new File([data as BlobPart], name))
    for (const [k, v] of Object.entries(fields)) fd.append(k, v)
    return fd
  }
  const post = (path: string, body: BodyInit, headers: Record<string, string> = {}) => handle(new Request(`http://x${path}`, { method: 'POST', body, headers }))

  test('health and version, as Gotenberg answers them', async () => {
    const health = await handle(new Request('http://x/health'))
    expect(((await health.json()) as { status: string }).status).toBe('up')
    expect(await (await handle(new Request('http://x/version'))).text()).toBe('9.9.9')
  })

  test('one file gives the file', async () => {
    const res = await post('/convert', form([['report.doc', S('sample.doc')]]))
    expect(res.status).toBe(200)
    expect(res.headers.get('content-type')).toBe('application/vnd.openxmlformats-officedocument.wordprocessingml.document')
    expect(res.headers.get('content-disposition')).toContain('filename="report.docx"')
    expect(res.headers.get('gotenberg-trace')).toBeTruthy()
  })

  test('several files give a ZIP; Gotenberg-Output-Filename names it, Gotenberg-Trace comes back', async () => {
    const res = await post('/convert', form([['a.odt', S('sample.odt')], ['a.rtf', S('sample.rtf')], ['b.ods', S('sample.ods')]]), { 'Gotenberg-Output-Filename': 'out', 'Gotenberg-Trace': 't-1' })
    expect(res.status).toBe(200)
    expect(res.headers.get('content-disposition')).toContain('filename="out.zip"')
    expect(res.headers.get('gotenberg-trace')).toBe('t-1')
    const zip = await ZipArchive.open(new BlobSource(await res.blob()))
    expect(zip.list().map((e) => e.name).sort()).toEqual(['a (2).docx', 'a.docx', 'b.xlsx'])
    // names that differ only in case are two files too
    const cased = await post('/convert', form([['Report.md', S('sample.md')], ['report.md', S('sample.md')]], { to: 'txt' }))
    const both = await ZipArchive.open(new BlobSource(await cased.blob()))
    expect(both.list().map((e) => e.name).sort()).toEqual(['Report.txt', 'report (2).txt'])
  })

  test('a file of a batch refused: the others are given up; PDF of a document is refused before the pool is asked', async () => {
    const one = new Converter({ workers: 1, queue: 0, timeout: 0 })
    try {
      const h = createHandler(one)
      const busy = h(new Request('http://x/convert', { method: 'POST', body: form([['big.csv', bigCsv()], ['a.doc', S('sample.doc')]]) }))
      expect((await busy).status).toBe(503)
      expect(one.load.running).toBe(0)
      const ac = new AbortController()
      const running = one.convert(bigCsv(), { signal: ac.signal }).catch(() => {})
      expect((await h(new Request('http://x/forms/libreoffice/convert', { method: 'POST', body: form([['a.docx', S('sample.docx')]]) }))).status).toBe(501)
      ac.abort()
      await running
    } finally {
      await one.close()
    }
  })

  test('the file as the body, its name and format in the query', async () => {
    const res = await post('/convert?to=csv&name=sales.xlsx', S('sample.xlsx') as BodyInit)
    expect(res.headers.get('content-disposition')).toContain('filename="sales.csv"')
    expect(await res.text()).toContain('North,2026-01-16')
  })

  test('errors', async () => {
    const status = async (res: Promise<Response>) => (await res).status
    expect(await status(post('/forms/libreoffice/convert', form([['a.docx', S('sample.docx')]])))).toBe(501)
    expect(await status(post('/convert', form([['a.docx', S('sample.docx')]], { to: 'nope' })))).toBe(400)
    expect(await status(post('/convert', form([['a.pptx', S('sample.pptx')]], { to: 'csv' })))).toBe(400)
    expect(await status(post('/convert', form([['a.docx', new TextEncoder().encode('PK\x03\x04 not a zip')]])))).toBe(422)
    expect(await status(post('/convert', form([['a.docx', new Uint8Array(2 * 1024 * 1024)]])))).toBe(413)
    expect(await status(post('/convert', form([['a', S('sample.md')], ['b', S('sample.md')], ['c', S('sample.md')], ['d', S('sample.md')]])))).toBe(400)
    expect(await status(post('/convert', new FormData()))).toBe(400)
    expect(await status(handle(new Request('http://x/convert')))).toBe(405)
    expect(await status(handle(new Request('http://x/nope')))).toBe(404)
  })
})

describe('office-convert', () => {
  const dir = mkdtempSync(`${tmpdir()}/office-convert-`)
  afterAll(() => rmSync(dir, { recursive: true, force: true }))
  const CLI = `${import.meta.dir}/../packages/convert/src/cli.ts`
  const run = (...args: string[]) => {
    const p = Bun.spawnSync(['bun', CLI, ...args], { cwd: dir })
    return { code: p.exitCode, err: p.stderr.toString(), out: p.stdout }
  }

  test('in out, and an output that exists is not taken for an input', () => {
    writeFileSync(`${dir}/in.doc`, S('sample.doc'))
    expect(run('in.doc', 'out.docx').code).toBe(0)
    expect(existsSync(`${dir}/out.docx`)).toBe(true)
    const again = run('--force', 'in.doc', 'out.docx')
    expect(again.code).toBe(2)
    expect(again.err).toContain('--out out.docx')
    expect(run('in.doc', '-o', 'out.docx').err).toContain('exists (--force overwrites it)')
    expect(run('in.doc', '-o', 'out.docx', '--force').code).toBe(0)
    // two files that exist are two inputs
    writeFileSync(`${dir}/second.md`, S('sample.md'))
    expect(run('in.doc', 'second.md', '-t', 'txt').code).toBe(0)
    expect(existsSync(`${dir}/in.txt`) && existsSync(`${dir}/second.txt`)).toBe(true)
  })

  test('an input is never overwritten, under any name', () => {
    writeFileSync(`${dir}/src.doc`, S('sample.doc'))
    symlinkSync(`${dir}/src.doc`, `${dir}/alias.txt`)
    const r = run('src.doc', '-t', 'txt', '-o', 'alias.txt', '--force')
    expect(r.code).toBe(1)
    expect(r.err).toContain('is one of the inputs')
    expect(Buffer.compare(readFileSync(`${dir}/src.doc`), S('sample.doc'))).toBe(0)
    // stdin read from the file being written
    const p = Bun.spawnSync(['bun', CLI, '-', '--name', 'src.doc', '-t', 'txt', '-o', 'src.doc', '--force'], { cwd: dir, stdin: Bun.file(`${dir}/src.doc`) })
    expect(p.exitCode).toBe(1)
    expect(p.stderr.toString()).toContain('is one of the inputs')
    expect(Buffer.compare(readFileSync(`${dir}/src.doc`), S('sample.doc'))).toBe(0)
  })

  test('to stdout, whole', () => {
    writeFileSync(`${dir}/big.csv`, bigCsv())
    const r = run('big.csv', '-o', '-', '--timeout', '0')
    expect(r.code).toBe(0)
    expect(r.out.length).toBeGreaterThan(1_000_000)
    expect(r.out.subarray(0, 2).toString('latin1')).toBe('PK')
  })
})

describe('PDF', () => {
  const dir = mkdtempSync(`${tmpdir()}/office-convert-pdf-`)
  afterAll(() => rmSync(dir, { recursive: true, force: true }))
  // a font named as the sample's text asks (Calibri), written to a file: fonts are declared by their paths
  writeFileSync(`${dir}/calibri.ttf`, testFont({ family: 'Calibri', advance: 500 }))
  const fonts = [{ family: 'Calibri', src: `${dir}/calibri.ttf` }]
  const latin1 = (b: Uint8Array) => new TextDecoder('latin1').decode(b)
  const count = (pdf: Uint8Array) => Number(/\/Count (\d+)/.exec(latin1(pdf))?.[1])

  test('a presentation (also a .ppt, .odp, Keynote or Visio file) as a PDF, its text in the fonts given', async () => {
    const r = await convert(S('sample.pptx'), { to: 'pdf', name: 'deck.pptx', fonts })
    expect(r).toMatchObject({ name: 'deck.pdf', type: 'application/pdf', ext: 'pdf', kind: 'pptx' })
    expect(latin1(r.data.subarray(0, 8))).toBe('%PDF-1.7')
    expect(count(r.data)).toBe(5)
    expect(latin1(r.data)).toMatch(/\/BaseFont\/[A-Z]{6}\+Calibri/)
    for (const f of ['sample.ppt', 'sample.odp', 'sample.key', 'sample.vsdx']) expect(count((await convert(S(f), { to: 'pdf', name: f, fonts })).data)).toBeGreaterThan(0)
  })

  test('the slides asked for; one that is not there is refused', async () => {
    expect(count((await convert(S('sample.pptx'), { to: 'pdf', fonts, slides: [1, 2] })).data)).toBe(2)
    expect(await convert(S('sample.pptx'), { to: 'pdf', fonts, slides: [9] }).catch((e) => e.code)).toBe('unsupported')
    expect(pageRanges('1-3, 5')).toEqual([0, 1, 2, 4])
    expect(pageRanges('3-1')).toBeNull()
    expect(pageRanges('1-1000000000')).toBeNull()
  })

  test('text no font file serves is set in Helvetica, and said so', async () => {
    const r = await convert(S('sample.pptx'), { to: 'pdf', fonts: [] })
    expect(latin1(r.data)).toContain('/BaseFont/Helvetica')
    expect(r.warnings.some((w) => /No font file for Calibri/.test(w))).toBe(true)
  })

  test('the font files of folders, by family, weight and style; the generic families after them', async () => {
    writeFileSync(`${dir}/sans.ttf`, testFont({ family: 'Liberation Sans' }))
    writeFileSync(`${dir}/not-a-font.ttf`, 'nothing')
    const found = await findFonts([dir])
    expect(found.filter((f) => f.family === 'Calibri')).toEqual([{ family: 'Calibri', src: `${dir}/calibri.ttf`, weight: 400, style: 'normal' }])
    expect(found.find((f) => f.family === 'sans-serif')?.src).toBe(`${dir}/sans.ttf`)
    expect(found.some((f) => String(f.src).endsWith('not-a-font.ttf'))).toBe(false)
    // a host's own generic family is kept
    expect(withGenerics([{ family: 'sans-serif', src: 'x' }, { family: 'Arial', src: 'y' }]).filter((f) => f.family === 'sans-serif')).toEqual([{ family: 'sans-serif', src: 'x' }])
  })

  test('on worker threads, the fonts read by their paths in the thread', async () => {
    const pool = new Converter({ workers: 1, fonts })
    try {
      const r = await pool.convert(S('sample.pptx'), { to: 'pdf' })
      expect(latin1(r.data)).toMatch(/\+Calibri/)
    } finally {
      await pool.close()
    }
  })

  test("Gotenberg's route: presentations as PDF, nativePageRanges; documents 501; what is not done here, 400", async () => {
    const handle = createHandler(new Converter({ workers: 0, fonts }))
    const form = (files: [string, Uint8Array][], fields: Record<string, string> = {}) => {
      const fd = new FormData()
      for (const [name, data] of files) fd.append('files', new File([data as BlobPart], name))
      for (const [k, v] of Object.entries(fields)) fd.append(k, v)
      return fd
    }
    const post = (body: FormData, path = '/forms/libreoffice/convert') => handle(new Request(`http://x${path}`, { method: 'POST', body, headers: { 'Gotenberg-Output-Filename': 'out' } }))
    const res = await post(form([['deck.pptx', S('sample.pptx')]], { nativePageRanges: '2-3' }))
    expect(res.status).toBe(200)
    expect(res.headers.get('content-type')).toBe('application/pdf')
    expect(res.headers.get('content-disposition')).toContain('filename="out.pdf"')
    expect(count(new Uint8Array(await res.arrayBuffer()))).toBe(2)
    expect((await post(form([['a.docx', S('sample.docx')]]))).status).toBe(501)
    expect((await post(form([['a.pptx', S('sample.pptx')]], { nativePageRanges: 'x' }))).status).toBe(400)
    expect((await post(form([['a.pptx', S('sample.pptx')], ['b.odp', S('sample.odp')]], { merge: 'true' }))).status).toBe(400)
    expect((await post(form([['a.pptx', S('sample.pptx')]], { pdfa: 'PDF/A-2b' }))).status).toBe(400)
    // two files: a ZIP of two PDFs
    const zip = await post(form([['a.pptx', S('sample.pptx')], ['b.odp', S('sample.odp')]]))
    expect(zip.headers.get('content-type')).toBe('application/zip')
    // and through /convert
    const one = await post(form([['a.pptx', S('sample.pptx')]], { to: 'pdf', slides: '1' }), '/convert')
    expect(count(new Uint8Array(await one.arrayBuffer()))).toBe(1)
  })

  test('office-convert deck.pptx deck.pdf, with --fonts and only those', () => {
    writeFileSync(`${dir}/deck.pptx`, S('sample.pptx'))
    const p = Bun.spawnSync(['bun', `${import.meta.dir}/../packages/convert/src/cli.ts`, '--fonts', `${dir}/calibri.ttf`, '--no-system-fonts', '--slides', '1,3', 'deck.pptx', 'deck.pdf'], { cwd: dir })
    expect(p.exitCode).toBe(0)
    const pdf = new Uint8Array(readFileSync(`${dir}/deck.pdf`))
    expect(count(pdf)).toBe(2)
    expect(latin1(pdf)).toMatch(/\+Calibri/)
  })
})

describe('PDF: review', () => {
  const dir = mkdtempSync(`${tmpdir()}/office-convert-pdf2-`)
  afterAll(() => rmSync(dir, { recursive: true, force: true }))
  writeFileSync(`${dir}/calibri.ttf`, testFont({ family: 'Calibri', advance: 500 }))
  const fonts = [{ family: 'Calibri', src: `${dir}/calibri.ttf` }]
  const latin1 = (b: Uint8Array) => new TextDecoder('latin1').decode(b)
  const count = (pdf: Uint8Array) => Number(/\/Count (\d+)/.exec(latin1(pdf))?.[1])
  const form = (files: [string, Uint8Array][], fields: Record<string, string> = {}) => {
    const fd = new FormData()
    for (const [name, data] of files) fd.append('files', new File([data as BlobPart], name))
    for (const [k, v] of Object.entries(fields)) fd.append(k, v)
    return fd
  }
  const handle = createHandler(new Converter({ workers: 0, fonts }))
  const post = (body: FormData, path = '/forms/libreoffice/convert') => handle(new Request(`http://x${path}`, { method: 'POST', body }))

  test('page ranges past what a presentation has are refused whole, never cut short', () => {
    expect(pageRanges('100001-200000,1')).toBeNull()
    expect(pageRanges('1-100000,1-100000')).toBeNull()
    expect(pageRanges('9'.repeat(30))).toBeNull()
    expect(pageRanges('1-100000')?.length).toBe(100000)
  })

  test("what Gotenberg's route is asked for and does not do is refused; its booleans are Go's", async () => {
    const deck = S('sample.pptx')
    expect((await post(form([['a.pptx', deck]], { userPassword: 'secret' }))).status).toBe(400)
    expect((await post(form([['a.pptx', deck]], { ownerPassword: 'secret' }))).status).toBe(400)
    expect((await post(form([['a.pptx', deck], ['b.pptx', deck]], { merge: '1' }))).status).toBe(400)
    expect((await post(form([['a.pptx', deck]], { merge: 'maybe' }))).status).toBe(400)
    const two = await post(form([['deck.pptx', deck], ['test.odp', S('sample.odp')]], { merge: 'false' }))
    expect(two.status).toBe(200)
    const zip = await ZipArchive.open(new BlobSource(new Blob([await two.arrayBuffer()])))
    expect(zip.list().map((e) => e.name).sort()).toEqual(['deck.pptx.pdf', 'test.odp.pdf'])
  })

  test('hidden slides only with exportHiddenSlides, also when nativePageRanges names them', async () => {
    const engine = new Engine(() => {})
    await engine.open(S('sample.pptx'), { name: 'sample.pptx' })
    await engine.call('pptx.slides', [{ op: 'hide', index: 0, hidden: true }])
    const hidden = new Uint8Array(await ((await engine.call('save', [])) as Blob).arrayBuffer())
    const pages = async (fields: Record<string, string>) => {
      const res = await post(form([['deck.pptx', hidden]], fields))
      return res.status === 200 ? count(new Uint8Array(await res.arrayBuffer())) : res.status
    }
    expect(await pages({})).toBe(4)
    expect(await pages({ nativePageRanges: '1' })).toBe(400)
    expect(await pages({ nativePageRanges: '1-2' })).toBe(1)
    expect(await pages({ nativePageRanges: '1', exportHiddenSlides: 'true' })).toBe(1)
    expect(await pages({ exportHiddenSlides: 'true' })).toBe(5)
  })

  test("a presentation's bytes are a presentation whatever its name; an unreadable file is 400 for Gotenberg, 422 for /convert", async () => {
    const res = await post(form([['deck.docx', S('sample.pptx')]]))
    expect(res.status).toBe(200)
    expect(count(new Uint8Array(await res.arrayBuffer()))).toBe(5)
    expect((await post(form([['a.docx', S('sample.docx')]]))).status).toBe(501)
    const broken = new TextEncoder().encode('PK\x03\x04 not a zip')
    expect((await post(form([['a.pptx', broken]]))).status).toBe(400)
    expect((await post(form([['a.pptx', broken]], { to: 'pdf' }), '/convert')).status).toBe(422)
  })

  test('bold and italic text falls back to the fonts of other scripts too', async () => {
    const { buildFont } = await import('./font-tables')
    const { FontBook } = await import('../src/core/font')
    const latin = (italic: boolean) => buildFont({ family: 'Liberation Sans', italic, advances: [500, 600], chars: [[0x58, 1]] })
    const cjk = buildFont({ family: 'Noto Sans CJK SC', advances: [500, 1000], chars: [[0x6f22, 1]] })
    const list = withGenerics([
      { family: 'Liberation Sans', src: latin(false), weight: 400, style: 'normal' },
      { family: 'Liberation Sans', src: latin(true), weight: 400, style: 'italic' },
      { family: 'Noto Sans CJK SC', src: cjk, weight: 400, style: 'normal' },
    ])
    const book = new FontBook(list)
    for (const italic of [false, true])
      for (const weight of [400, 700]) {
        const m = await book.measure({ text: 'X漢', font: '"Missing", sans-serif', size: 10, italic, weight })
        expect(m.pieces.map((p) => p.family)).toEqual(['sans-serif', 'sans-serif'])
        expect(m.pieces.every((p) => p.font)).toBe(true)
      }
  })

  test('conversions waiting for the fonts are queued work: counted, given up, closed; a failing font list is a rejection', async () => {
    let found!: (list: typeof fonts) => void
    const pool = new Converter({ workers: 1, fonts: () => new Promise((r) => (found = r)) })
    const ac = new AbortController()
    const waiting = pool.convert(S('sample.pptx'), { to: 'pdf', signal: ac.signal })
    const other = pool.convert(S('sample.pptx'), { to: 'pdf' })
    expect(pool.load.queued).toBe(2)
    ac.abort(new Error('gone'))
    expect(waiting).rejects.toThrow('gone')
    await pool.close()
    expect(other).rejects.toThrow('closed')
    found(fonts)
    const failing = new Converter({ workers: 1, fonts: () => { throw new Error('no folders') } })
    expect(failing.convert(S('sample.pptx'), { to: 'pdf' })).rejects.toThrow('no folders')
    await failing.close()
  })
})

describe('PDF: re-check', () => {
  const dir = mkdtempSync(`${tmpdir()}/office-convert-pdf3-`)
  afterAll(() => rmSync(dir, { recursive: true, force: true }))
  writeFileSync(`${dir}/calibri.ttf`, testFont({ family: 'Calibri', advance: 500 }))
  const fonts = [{ family: 'Calibri', src: `${dir}/calibri.ttf` }]
  const latin1 = (b: Uint8Array) => new TextDecoder('latin1').decode(b)
  const form = (files: [string, Uint8Array][], fields: [string, string][] = []) => {
    const fd = new FormData()
    for (const [name, data] of files) fd.append('files', new File([data as BlobPart], name))
    for (const [k, v] of fields) fd.append(k, v)
    return fd
  }
  const handle = createHandler(new Converter({ workers: 0, fonts }))
  const post = (body: FormData, path = '/forms/libreoffice/convert') => handle(new Request(`http://x${path}`, { method: 'POST', body }))
  const deck = S('sample.pptx')

  test("a field given twice is read by its first value; Gotenberg's booleans are read as they are", async () => {
    expect((await post(form([['a.pptx', deck]], [['userPassword', 'secret'], ['userPassword', '']]))).status).toBe(400)
    expect((await post(form([['a.pptx', deck]], [['exportHiddenSlides', ' true ']]))).status).toBe(400)
    expect((await post(form([['a.pptx', deck]], [['to', 'pdf'], ['hidden', 'TrUe']]), '/convert')).status).toBe(200)
  })

  test('one file keeps its own name; what the bytes are decides, whatever the name', async () => {
    const one = await post(form([['deck.pptx', deck]]))
    expect(one.headers.get('content-disposition')).toContain('filename="deck.pdf"')
    expect((await post(form([['renamed.docx', S('sample.key')]]))).status).toBe(200)
    expect((await post(form([['broken.docx', new TextEncoder().encode('PK\x03\x04 not a zip')]]))).status).toBe(400)
    expect((await post(form([['notes.pptx', S('sample.docx')]]))).status).toBe(501)
    expect(pageRanges('0000001')).toEqual([0])
  })

  test('a pool of no queue converts a PDF on an idle thread; a font list is taken as it was given', async () => {
    const list = [...fonts]
    const pool = new Converter({ workers: 1, queue: 0, fonts: list })
    try {
      list.length = 0
      const r = await pool.convert(deck, { to: 'pdf' })
      expect(latin1(r.data)).toMatch(/\+Calibri/)
    } finally {
      await pool.close()
    }
  })

  test('conversions waiting for the fonts count toward the bytes the pool takes', async () => {
    let found!: (l: typeof fonts) => void
    const pool = new Converter({ workers: 1, bytes: 1, fonts: () => new Promise((r) => (found = r)) })
    const first = pool.convert(new Uint8Array(2), { to: 'pdf' }).catch((e) => e)
    expect(pool.convert(new Uint8Array(2), { to: 'pdf' })).rejects.toBeInstanceOf(BusyError)
    // (the font list is asked for a moment later)
    await new Promise((r) => setTimeout(r, 0))
    found(fonts)
    expect(await first).toBeInstanceOf(Error)
    await pool.close()
  })

  test('every file of a family of other scripts serves its characters; a pipe named like a font does not hold the scan up', async () => {
    const { buildFont } = await import('./font-tables')
    const { FontBook } = await import('../src/core/font')
    const a = buildFont({ family: 'Noto Sans CJK SC', advances: [500, 1000], chars: [[0x6f22, 1]] })
    const b = buildFont({ family: 'Noto Sans CJK SC', advances: [500, 1000], chars: [[0x5b57, 1]] })
    const book = new FontBook(withGenerics([{ family: 'Noto Sans CJK SC', src: a, weight: 400, style: 'normal' }, { family: 'Noto Sans CJK SC', src: b, weight: 400, style: 'normal' }]))
    const m = await book.measure({ text: '漢字', font: 'sans-serif', size: 10 })
    expect(m.pieces.every((p) => p.font)).toBe(true)
    const pipes = mkdtempSync(`${tmpdir()}/office-convert-fifo-`)
    try {
      Bun.spawnSync(['mkfifo', `${pipes}/pipe.ttf`])
      writeFileSync(`${pipes}/calibri.ttf`, testFont({ family: 'Calibri' }))
      const found = await Promise.race([findFonts([pipes]), new Promise((r) => setTimeout(() => r('stuck'), 3000))])
      expect(Array.isArray(found) && found.some((f) => f.family === 'Calibri')).toBe(true)
    } finally {
      rmSync(pipes, { recursive: true, force: true })
    }
  })
})
