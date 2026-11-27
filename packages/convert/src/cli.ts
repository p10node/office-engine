#!/usr/bin/env node
/** office-convert: converts files on the command line, or runs the HTTP service (`office-convert serve`). */
import { closeSync, constants, createReadStream, existsSync, fstatSync, ftruncateSync, openSync, statSync, writeSync, type Stats } from 'node:fs'
import { availableParallelism } from 'node:os'
import { basename, delimiter, dirname, extname, join, resolve } from 'node:path'
import { parseArgs } from 'node:util'
import pkg from '../package.json' with { type: 'json' }
import { pageRanges, targetOf, type Target } from './convert'
import { findFonts, systemFontDirs } from './fonts'
import { Converter } from './pool'
import { createHandler } from './server'
import { listen } from './serve'

const USAGE = `office-convert ${pkg.version}: Office files converted without LibreOffice

Usage:
  office-convert [options] <file…>     each file next to itself, or into --out
  office-convert <in> <out>            the format from the output's extension
  office-convert serve [options]       the HTTP service (POST /convert, Gotenberg's routes)

Options:
  -t, --to <format>     docx, xlsx, pptx, pdf, txt, md, html, csv, tsv
                        (default: the Office format the file opens as: .doc → .docx, .csv → .xlsx)
  -o, --out <path>      the output file (one input; - for stdout) or an existing directory
      --name <name>     the name of the file read from stdin (-): its extension tells text formats
      --sheet <n|name>  CSV / TSV: the sheet written (default: the one the file was saved on)
      --encoding <enc>  text files: the encoding they are read in (default: a guess)
      --slides <pages>  PDF: the slides written, as 1-3,5 (default all but the hidden ones)
      --hidden          PDF: hidden slides too
      --fonts <path>    PDF: a font file or folder to set text in, before the system's (again for more)
      --no-system-fonts PDF: only the fonts of --fonts
  -f, --force           overwrite files that exist
  -j, --jobs <n>        files converted at once (default: the number of CPUs)
      --timeout <s>     seconds a file may take (default 120; 0: no limit)
      --max-size <MB>   largest file read (default 100)

serve:
  -p, --port <n>        port (default $PORT or 3000)
      --host <addr>     address (default $HOST or 0.0.0.0)
  -j, --jobs <n>        conversions at once (default $WORKERS or the number of CPUs)
      --timeout <s>     seconds a file may take (default $TIMEOUT or 120)
      --max-size <MB>   largest request (default $MAX_SIZE or 100)
      --max-files <n>   files in one request (default $MAX_FILES or 20)
      --fonts <path>    font files or folders for PDF output (default $FONTS, separated by "${delimiter}")
      --no-system-fonts not the system's fonts (default: $SYSTEM_FONTS=0 says the same)

  -h, --help, -v, --version
`

const fail = (message: string, code = 2): never => {
  process.stderr.write(`office-convert: ${message}\n`)
  process.exit(code)
}

const { values: o, positionals } = (() => {
  try {
    return parseArgs({
      allowPositionals: true,
      options: {
        to: { type: 'string', short: 't' },
        out: { type: 'string', short: 'o' },
        name: { type: 'string' },
        sheet: { type: 'string' },
        encoding: { type: 'string' },
        force: { type: 'boolean', short: 'f' },
        jobs: { type: 'string', short: 'j' },
        timeout: { type: 'string' },
        'max-size': { type: 'string' },
        'max-files': { type: 'string' },
        port: { type: 'string', short: 'p' },
        slides: { type: 'string' },
        hidden: { type: 'boolean' },
        fonts: { type: 'string', multiple: true },
        'no-system-fonts': { type: 'boolean' },
        host: { type: 'string' },
        help: { type: 'boolean', short: 'h' },
        version: { type: 'boolean', short: 'v' },
      },
    })
  } catch (e) {
    return fail(`${(e as Error).message}\n\n${USAGE}`)
  }
})()

/** a number option (a flag, else an environment variable, else the default) */
const num = (flag: string | undefined, env: string | undefined, dflt: number, what: string): number => {
  const v = flag ?? env
  if (v === undefined || v === '') return dflt
  const n = Number(v)
  if (!Number.isFinite(n) || n < 0) fail(`${what} must be a number, not “${v}”`)
  return n
}

if (o.help) {
  process.stdout.write(USAGE)
  process.exit(0)
}
if (o.version) {
  process.stdout.write(`${pkg.version}\n`)
  process.exit(0)
}

const env = process.env

/** the fonts PDFs are set in: --fonts (or $FONTS) first, then the system's unless told not to */
const fontDirs = () => [...(o.fonts ?? env.FONTS?.split(delimiter).filter(Boolean) ?? []), ...(o['no-system-fonts'] || env.SYSTEM_FONTS === '0' ? [] : systemFontDirs())]

if (positionals[0] === 'serve') {
  if (positionals.length > 1) fail(`serve takes no files\n\n${USAGE}`)
  const maxBytes = num(o['max-size'], env.MAX_SIZE, 100, '--max-size') * 1024 * 1024
  // (found before the service listens: no request waits for the font folders to be read)
  const fonts = await findFonts(fontDirs())
  const converter = new Converter({ workers: num(o.jobs, env.WORKERS, NaN, '--jobs') || undefined, timeout: num(o.timeout, env.TIMEOUT, 120, '--timeout') * 1000, fonts })
  const handler = createHandler(converter, { maxBytes, maxFiles: num(o['max-files'], env.MAX_FILES, 20, '--max-files'), version: pkg.version })
  const server = await listen(handler, { port: num(o.port, env.PORT, 3000, '--port'), host: o.host ?? env.HOST ?? '0.0.0.0', maxBytes })
  const families = new Set(fonts.map((f) => f.family)).size
  process.stderr.write(`office-convert ${pkg.version} listening on ${server.url} (${converter.workers} threads, ${families} font families for PDF)\n`)
  const stop = () => {
    void server.close().then(() => converter.close()).then(() => process.exit(0))
  }
  process.on('SIGINT', stop)
  process.on('SIGTERM', stop)
} else await convertFiles()

async function convertFiles() {
  let inputs = positionals
  let out = o.out
  let to: Target | undefined
  if (o.to) to = targetOf(o.to) ?? fail(`Unknown format “${o.to}”: docx, xlsx, pptx, pdf, txt, md, html, csv, tsv`)
  // office-convert in.doc out.docx: a second name with a format's extension is the output, when it is no file yet; one that
  // is a file is an input, but not with --force, where writing it may have been meant
  if (!out && !to && inputs.length === 2 && targetOf(extname(inputs[1]))) {
    if (!existsSync(inputs[1])) {
      out = inputs[1]
      inputs = [inputs[0]]
    } else if (o.force) fail(`${inputs[1]} exists: give --out ${inputs[1]} to write it, or --to / --out <directory> to convert both files`)
  }
  if (!inputs.length) fail(`No file given\n\n${USAGE}`)
  const outDir = out && out !== '-' && existsSync(out) && statSync(out).isDirectory() ? out : undefined
  if (out && !outDir && inputs.length > 1) fail('--out names a file, but there are several inputs: give a directory')
  // the output's extension is the format, when --to does not say
  if (out && !outDir && out !== '-' && !to) to = targetOf(extname(out)) ?? undefined
  const maxBytes = num(o['max-size'], undefined, 100, '--max-size') * 1024 * 1024
  const tooBig = () => new Error(`larger than ${maxBytes / 1024 / 1024} MB (--max-size)`)
  // (no byte budget: as many files are read as there are threads, each within --max-size)
  const slides = o.slides !== undefined ? (pageRanges(o.slides) ?? fail(`--slides “${o.slides}” is not a list of pages such as 1-3,5`)) : undefined
  // (the font folders are read when a PDF is first asked for)
  const converter = new Converter({ workers: Math.min(inputs.length, num(o.jobs, undefined, 0, '--jobs') || availableParallelism()), timeout: num(o.timeout, undefined, 120, '--timeout') * 1000, bytes: Infinity, fonts: () => findFonts(fontDirs()) })

  // the inputs as files (device and inode): an output that is one of them, under any name (a link), is never written;
  // stdin too, when it is a file (< report.doc)
  const idOf = (st: Stats) => `${st.dev}:${st.ino}`
  const inputIds = new Set<string>()
  for (const input of inputs) {
    if (input !== '-' && existsSync(input)) inputIds.add(idOf(statSync(input)))
    else if (input === '-') {
      const st = fstatSync(0)
      if (st.isFile()) inputIds.add(idOf(st))
    }
  }
  // outputs this run writes: two inputs of the same name (report.doc, report.odt) would write one file
  const claimed = new Map<string, string>()

  const one = async (input: string) => {
    const stdin = input === '-'
    if (!stdin && statSync(input).size > maxBytes) throw tooBig()
    // (read up to the limit whatever the size said: a pipe or a device has none)
    const data = await readAtMost(stdin ? process.stdin : createReadStream(input), maxBytes, tooBig)
    const r = await converter.convert(data, { ...(to ? { to } : {}), name: stdin ? (o.name ?? 'document') : basename(input), ...(o.sheet !== undefined ? { sheet: o.sheet } : {}), ...(o.encoding ? { encoding: o.encoding } : {}), ...(slides ? { slides } : {}), ...(o.hidden ? { hidden: true } : {}) })
    for (const w of r.warnings) process.stderr.write(`${input}: ${w}\n`)
    if (out === '-') {
      await new Promise<void>((done, failed) => process.stdout.write(r.data, (e) => (e ? failed(e) : done())))
      return
    }
    const dest = out && !outDir ? out : join(outDir ?? (stdin ? '.' : dirname(input)), r.name)
    const other = claimed.get(resolve(dest))
    if (other) throw new Error(`${dest} is written from ${other} already`)
    claimed.set(resolve(dest), input)
    // without --force, the file is created only when there is none (a link left in its place included); with it, the file
    // opened is checked before anything in it is replaced, so it is not one of the inputs whatever its name
    const fd = openSync(dest, o.force ? constants.O_WRONLY | constants.O_CREAT : 'wx')
    try {
      if (inputIds.has(idOf(fstatSync(fd)))) throw new Error(`${dest} is one of the inputs: give --out or --to`)
      ftruncateSync(fd, 0)
      for (let at = 0; at < r.data.byteLength; ) at += writeSync(fd, r.data, at, r.data.byteLength - at, at)
    } finally {
      closeSync(fd)
    }
    process.stderr.write(`${input} → ${dest}\n`)
  }

  let failed = 0
  let next = 0
  // as many files at once as there are threads: the others are read once a thread is free
  const runner = async () => {
    while (next < inputs.length) {
      const input = inputs[next++]
      try {
        await one(input)
      } catch (e) {
        failed++
        const err = e as NodeJS.ErrnoException
        process.stderr.write(`${input}: ${err.code === 'EEXIST' ? `${err.path ?? 'the output'} exists (--force overwrites it)` : err.message}\n`)
      }
    }
  }
  await Promise.all(Array.from({ length: Math.max(1, converter.workers) }, runner))
  await converter.close()
  // (not process.exit: what is still going to stdout is written first)
  process.exitCode = failed ? 1 : 0
}

/** a stream read whole, refused past `max` bytes as soon as it gets there */
async function readAtMost(stream: AsyncIterable<Uint8Array>, max: number, tooBig: () => Error): Promise<Uint8Array> {
  const chunks: Uint8Array[] = []
  let size = 0
  for await (const chunk of stream) {
    size += chunk.byteLength
    if (size > max) throw tooBig()
    chunks.push(chunk)
  }
  return Buffer.concat(chunks)
}
