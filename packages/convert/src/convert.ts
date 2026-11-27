/**
 * One file in, one file out, on this thread: the engine opens the file (any format the viewer opens), waits until the
 * document is read and writes the format asked for. No DOM, no worker: Bun and Node run it as it is.
 */
import type { Block } from '../../../src/core/docx/types'
import type { SlideModel } from '../../../src/core/pptx/types'
import type { SheetProgress } from '../../../src/core/xlsx/types'
import { Engine } from '../../../src/core/worker/engine'
import type { Events, OpenResult } from '../../../src/core/worker/protocol'
import { docxToHtml, docxToMarkdown, docxToText, pictureData, slidesToText } from '../../../src/react/file/text'

/** the formats a file is converted to: an Office file, or the text of the document */
export type Target = 'docx' | 'xlsx' | 'pptx' | 'pdf' | 'txt' | 'md' | 'html' | 'csv' | 'tsv'

export const TARGETS: readonly Target[] = ['docx', 'xlsx', 'pptx', 'pdf', 'txt', 'md', 'html', 'csv', 'tsv']

/** what a file opens as: a document, a workbook or a presentation */
export type Kind = 'docx' | 'xlsx' | 'pptx'

/** what each kind of document can be written as; the first is what it opens as (`docx` for a .doc, an .odt, a .md…) */
export const WRITES: Record<Kind, readonly Target[]> = {
  docx: ['docx', 'txt', 'md', 'html'],
  xlsx: ['xlsx', 'csv', 'tsv'],
  pptx: ['pptx', 'txt'],
}

const MIME: Record<string, string> = {
  docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  pdf: 'application/pdf',
  txt: 'text/plain;charset=utf-8',
  md: 'text/markdown;charset=utf-8',
  html: 'text/html;charset=utf-8',
  csv: 'text/csv;charset=utf-8',
  tsv: 'text/tab-separated-values;charset=utf-8',
}

export interface ConvertOptions {
  /** the format to write (default: the Office format the file opens as: `docx` for a .doc, `xlsx` for a .csv…) */
  to?: Target
  /** the file's name: its extension helps tell the format of text files (.md, .csv, .puml…) and names the result */
  name?: string
  /** CSV / TSV output: the sheet written (its index or its name; default the sheet the file was saved on) */
  sheet?: number | string
  /** text files: the encoding the bytes are read in (a `TextDecoder` label), instead of a guess */
  encoding?: string
  /** the file is source code in this language ("typescript", "py"…), whatever its name says */
  language?: string
  /** bytes the file written may have (default 1 GB): a larger one fails with a `ConvertError` `too-large` */
  maxOutput?: number
  /** bytes a gzip-compressed input (.csv.gz, .svgz) may unpack to (default 1 GB) */
  maxUnpacked?: number
  /** gives up the conversion: the promise rejects with an `AbortError` */
  signal?: AbortSignal
}

export interface Converted {
  data: Uint8Array
  /** the name the file is given: the input's, with the extension of what was written */
  name: string
  /** the file's MIME type */
  type: string
  /** the format written; an Office file may keep its own variant (`docm`, `xltx`…) */
  ext: string
  /** what the file opened as */
  kind: Kind
  /** the format the file was read from when it is not an Office Open XML file ('doc', 'odt', 'pages', 'csv', 'markdown'…) */
  from?: string
  /** what the document lost on the way (more columns than a sheet holds…) */
  warnings: string[]
}

/** A request the converter cannot do (a format it does not write, from this kind of document): the HTTP service answers 400 / 501 */
export class ConvertError extends Error {
  constructor(
    message: string,
    readonly code: 'unsupported' | 'not-implemented' | 'too-large',
  ) {
    super(message)
    this.name = 'ConvertError'
  }
}

/** `to` read from a request or a command line: lower case, without a dot; null when it names no format */
export function targetOf(s: string | null | undefined): Target | null {
  const t = (s ?? '').trim().toLowerCase().replace(/^\./, '')
  const alias: Record<string, Target> = { markdown: 'md', text: 'txt', htm: 'html' }
  const v = alias[t] ?? t
  return (TARGETS as readonly string[]).includes(v) ? (v as Target) : null
}

const baseName = (name: string | undefined) => (name ?? '').replace(/^.*[\\/]/, '').replace(/\.[^.]*$/, '') || 'document'

const GB = 1024 * 1024 * 1024

const tooLarge = (max: number) => new ConvertError(`The converted file would be larger than ${Math.round(max / (1024 * 1024))} MB`, 'too-large')

const aborted = (signal?: AbortSignal) => {
  if (signal?.aborted) throw signal.reason instanceof Error ? signal.reason : new DOMException('The conversion was aborted', 'AbortError')
}

/** Converts one file: opens it, waits until it is read whole, writes it as `to`. */
export async function convert(input: Uint8Array | ArrayBuffer | Blob, opts: ConvertOptions = {}): Promise<Converted> {
  if (opts.to === 'pdf') throw new ConvertError('PDF output is not available yet: it needs the layout of p10node/office-engine#44', 'not-implemented')
  aborted(opts.signal)
  const maxOutput = opts.maxOutput ?? GB
  const warnings: string[] = []
  const listeners = new Set<(event: keyof Events, data: unknown) => void>()
  const engine = new Engine((event, data) => {
    if (event === 'warning') warnings.push((data as Events['warning']).text)
    for (const fn of listeners) fn(event, data)
  })
  const stop = () => engine.dispose()
  opts.signal?.addEventListener('abort', stop, { once: true })
  try {
    const name = opts.name ?? (typeof File !== 'undefined' && input instanceof File ? input.name : undefined)
    const info = await engine.open(input, { ...(name ? { name } : {}), ...(opts.encoding ? { encoding: opts.encoding } : {}), ...(opts.language ? { language: opts.language } : {}), maxUnpacked: opts.maxUnpacked ?? GB })
    aborted(opts.signal)
    const to = opts.to ?? info.kind
    if (!WRITES[info.kind].includes(to)) {
      throw new ConvertError(`A ${KIND_NAME[info.kind]} cannot be written as ${to.toUpperCase()}: ${WRITES[info.kind].map((t) => t.toUpperCase()).join(', ')}`, 'unsupported')
    }
    const call = (method: string, ...args: unknown[]) => engine.call(method as never, args)

    let out: Blob | string
    if (info.kind === 'docx') {
      // the blocks come as events while the body is read; `docx.start` resolves once it is read whole
      const blocks: Block[] = []
      let failed: string | undefined
      listeners.add((event, data) => {
        if (event === 'docx:blocks') blocks.push(...(data as Block[]))
        else if (event === 'docx:done') failed = (data as Events['docx:done']).error
      })
      await call('docx.start')
      aborted(opts.signal)
      if (failed) throw new Error(failed)
      if (to === 'docx') out = (await call('save')) as Blob
      else if (to === 'txt') out = docxToText(blocks)
      else {
        const images = await pictureData(blocks, (t) => call('part', t) as Promise<Blob>)
        out = to === 'md' ? docxToMarkdown(blocks, images) : docxToHtml(blocks, { title: baseName(name), images })
      }
    } else if (info.kind === 'xlsx') {
      if (to === 'xlsx') out = (await call('save')) as Blob
      else {
        const sheet = sheetIndex(info, opts.sheet)
        await loaded(engine, listeners, sheet, opts.signal)
        // (refused while it is written, or before when the cells in use cannot fit: one at XFD1048576)
        out = (await call('xlsx.text', sheet, to === 'tsv' ? '\t' : ',', maxOutput).catch((e: Error) => {
          throw e instanceof RangeError ? tooLarge(maxOutput) : e
        })) as Blob
      }
    } else {
      if (to === 'pptx') out = (await call('save')) as Blob
      else {
        const slides: SlideModel[] = []
        for (const s of info.pptx!.slides) {
          aborted(opts.signal)
          slides.push((await call('pptx.slide', s.index)) as SlideModel)
        }
        out = slidesToText(slides)
      }
    }
    aborted(opts.signal)
    if ((typeof out === 'string' ? out.length : out.size) > maxOutput) throw tooLarge(maxOutput)
    const data = typeof out === 'string' ? new TextEncoder().encode(out) : new Uint8Array(await out.arrayBuffer())
    if (data.byteLength > maxOutput) throw tooLarge(maxOutput)
    // an Office file keeps its variant (macros, template): the engine says which
    const ext = to === info.kind && info.ext ? info.ext : to
    const from = info.from ?? info.converted ?? info.imported
    return { data, name: `${baseName(name)}.${ext}`, type: (typeof out !== 'string' && out.type) || MIME[to], ext, kind: info.kind, ...(from ? { from } : {}), warnings }
  } finally {
    opts.signal?.removeEventListener('abort', stop)
    engine.dispose()
  }
}

const KIND_NAME: Record<Kind, string> = { docx: 'document', xlsx: 'workbook', pptx: 'presentation' }

function sheetIndex(info: OpenResult, sheet: number | string | undefined): number {
  const sheets = info.xlsx!.sheets
  if (sheet === undefined || sheet === '') return info.xlsx!.active ?? 0
  const n = typeof sheet === 'number' ? sheet : /^\d+$/.test(sheet) ? Number(sheet) : sheets.findIndex((s) => s.name === sheet)
  if (!Number.isInteger(n) || n < 0 || n >= sheets.length) throw new ConvertError(`No sheet “${sheet}”: ${sheets.map((s) => s.name).join(', ')}`, 'unsupported')
  return n
}

/** loads a sheet and resolves once it is read whole (or rejects with what stopped it) */
function loaded(engine: Engine, listeners: Set<(event: keyof Events, data: unknown) => void>, sheet: number, signal?: AbortSignal): Promise<void> {
  return new Promise<void>((resolve, reject) => {
    const done = (event: keyof Events, data: unknown) => {
      const p = data as SheetProgress & { error?: string }
      if (event !== 'xlsx:progress' || p.sheet !== sheet || !p.done) return
      listeners.delete(done)
      signal?.removeEventListener('abort', stop)
      if (p.error) reject(new Error(p.error))
      else resolve()
    }
    const stop = () => {
      listeners.delete(done)
      reject(signal!.reason instanceof Error ? signal!.reason : new DOMException('The conversion was aborted', 'AbortError'))
    }
    signal?.addEventListener('abort', stop, { once: true })
    listeners.add(done)
    engine.call('xlsx.load', [sheet]).catch(reject)
  })
}
