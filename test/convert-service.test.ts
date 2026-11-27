import { afterAll, describe, expect, test } from 'bun:test'
import { existsSync, mkdtempSync, readFileSync, rmSync, symlinkSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { gzipSync } from 'node:zlib'
import { BusyError, convert, ConvertError, Converter, createHandler, TimeoutError } from '../packages/convert/src/index'
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

  test('formats a kind cannot be written as, and PDF, are refused', async () => {
    const e = await convert(S('sample.pptx'), { to: 'csv' }).catch((e) => e)
    expect(e).toBeInstanceOf(ConvertError)
    expect(e.code).toBe('unsupported')
    const pdf = await convert(S('sample.docx'), { to: 'pdf' }).catch((e) => e)
    expect(pdf.code).toBe('not-implemented')
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
  const converter = new Converter({ workers: 0 })
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

  test('a file of a batch refused: the others are given up; PDF is refused before the pool is asked', async () => {
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

  // (the big file takes the engine 4 to 6 seconds, a test has 5 by default: the time is this test's own)
  test('to stdout, whole', () => {
    writeFileSync(`${dir}/big.csv`, bigCsv())
    const r = run('big.csv', '-o', '-', '--timeout', '0')
    expect(r.code).toBe(0)
    expect(r.out.length).toBeGreaterThan(1_000_000)
    expect(r.out.subarray(0, 2).toString('latin1')).toBe('PK')
  }, 60_000)
})
