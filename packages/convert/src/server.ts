/**
 * The HTTP service, as a `fetch` handler (Request in, Response out): Bun.serve takes it as it is, serve.ts runs it on
 * Node's http module, and a host can mount it in its own router.
 *
 *   POST /convert                     multipart: the files (any field name), `to`, `sheet`, `encoding`, `language`,
 *                                     `slides` (PDF: "1-3,5"), `hidden`; or the file as the body: /convert?to=docx&name=report.doc
 *   POST /forms/libreoffice/convert   Gotenberg's route: the files to PDF, `nativePageRanges` the slides written,
 *                                     `exportHiddenSlides`. Presentations; documents and spreadsheets answer 501
 *                                     until their PDF is in; what it can be asked for and is not done here, 400
 *   GET  /health                      { status: 'up', details: … } as Gotenberg answers
 *   GET  /version                     the package version, as text
 *
 * One file gives the converted file, several a ZIP of them. Gotenberg's headers are read: `Gotenberg-Output-Filename`
 * names the result, `Gotenberg-Trace` is sent back (a new id when the request has none).
 */
import { writeZip } from '../../../src/core/zipwrite'
import { ConvertError, kindOfBytes, pageRanges, pdfLater, targetOf, type ConvertOptions, type Converted, type Target } from './convert'
import { BusyError, ThreadError, TimeoutError, type Converter } from './pool'

export interface ServiceOptions {
  /** largest request body, in bytes (default 100 MB): larger requests are answered 413 */
  maxBytes?: number
  /** files in one request (default 20) */
  maxFiles?: number
  /** bytes of request bodies held at once, being read or converted (default 1 GB): past it requests are answered 503 */
  maxPending?: number
  /** the version `/version` answers */
  version?: string
}

const MB = 1024 * 1024

class HttpError extends Error {
  constructor(
    readonly status: number,
    message: string,
  ) {
    super(message)
  }
}

/** A `fetch` handler serving the conversion routes; conversions run on `converter`. */
export function createHandler(converter: Converter, opts: ServiceOptions = {}): (req: Request) => Promise<Response> {
  const maxBytes = opts.maxBytes ?? 100 * MB
  const maxFiles = opts.maxFiles ?? 20
  const maxPending = opts.maxPending ?? 1024 * 1024 * 1024
  let pending = 0
  const version = opts.version ?? '0.0.0'

  return async (req) => {
    const url = new URL(req.url)
    const trace = req.headers.get('gotenberg-trace') || crypto.randomUUID()
    const answer = (res: Response) => {
      res.headers.set('Gotenberg-Trace', trace)
      return res
    }
    // the bytes of this request's body counted in `pending`, given back once it is answered
    let held = 0
    const hold = (n: number) => {
      held += n
      pending += n
      if (pending > maxPending) throw new HttpError(503, 'The converter is busy: try again later')
    }
    let gotenberg = false
    try {
      const path = url.pathname.replace(/\/+$/, '') || '/'
      if (path === '/health' && (req.method === 'GET' || req.method === 'HEAD')) {
        const { running, queued } = converter.load
        return answer(Response.json({ status: 'up', details: { convert: { status: 'up', timestamp: new Date().toISOString(), running, queued } } }))
      }
      if (path === '/version' && req.method === 'GET') return answer(new Response(version, { headers: { 'Content-Type': 'text/plain; charset=utf-8' } }))
      gotenberg = path === '/forms/libreoffice/convert'
      if (path !== '/convert' && !gotenberg) throw new HttpError(404, 'Not found')
      if (req.method !== 'POST') throw new HttpError(405, 'Method not allowed: POST the files')

      const { files, fields } = await readRequest(req, url, maxBytes, hold)
      if (!files.length) throw new HttpError(400, 'No file: POST the files as multipart/form-data, or one file as the body')
      if (files.length > maxFiles) throw new HttpError(400, `Too many files: at most ${maxFiles} in one request`)
      // Gotenberg's LibreOffice route writes PDF whatever else it is told; what it can be asked for and is not done here
      // is refused, not left out without a word (a file asked for with a password is not given without one)
      if (gotenberg) {
        if (flag(fields, 'merge') && files.length > 1) throw new HttpError(400, 'merge is not supported yet: convert the files one by one')
        if (fields.get('pdfa')) throw new HttpError(400, 'pdfa is not supported yet')
        if (flag(fields, 'pdfua')) throw new HttpError(400, 'pdfua is not supported yet')
        for (const f of ['userPassword', 'ownerPassword', 'splitMode', 'password']) if (fields.get(f)) throw new HttpError(400, `${f} is not supported yet`)
      }
      const field = gotenberg ? 'pdf' : fields.get('to')
      const to = field ? targetOf(field) : undefined
      if (to === null) throw new HttpError(400, `Unknown format “${field}”`)
      // a file whose bytes say it is a document or a workbook, which have no PDF output yet, is answered so before the
      // pool is asked (not 503 when it is busy); the engine decides for the others, whatever their names
      if (to === 'pdf')
        for (const f of files) {
          const later = pdfLater(kindOfBytes(f.data))
          if (later) throw later
        }
      const rangeField = gotenberg ? 'nativePageRanges' : 'slides'
      const range = fields.get(rangeField)?.trim()
      const slides = range ? pageRanges(range) : undefined
      if (slides === null) throw new HttpError(400, `${rangeField} “${range}” is not a list of pages such as 1-3,5`)
      const options: ConvertOptions = {
        ...(to ? { to: to as Target } : {}),
        ...(fields.get('sheet') ? { sheet: fields.get('sheet')! } : {}),
        ...(fields.get('encoding') ? { encoding: fields.get('encoding')! } : {}),
        ...(fields.get('language') ? { language: fields.get('language')! } : {}),
        ...(slides ? { slides } : {}),
        ...((gotenberg ? flag(fields, 'exportHiddenSlides') : /^(true|1|yes)$/i.test(fields.get('hidden')?.trim() ?? '')) ? { hidden: true } : {}),
      }
      // one file failing (busy, too large, damaged) fails the request: the others are given up with it
      const batch = new AbortController()
      const signal = AbortSignal.any([req.signal, batch.signal])
      const results = await Promise.all(
        files.map((f) =>
          converter.convert(f.data, { ...options, name: f.name, signal }).then(
            // Gotenberg names each PDF of a ZIP after its file, extension and all: deck.pptx.pdf
            (r) => (gotenberg && files.length > 1 ? { ...r, name: `${f.name.replace(/^.*[\\/]/, '') || 'document'}.pdf` } : r),
            (e: Error) => {
              batch.abort(e)
              throw e
            },
          ),
        ),
      )
      const outName = req.headers.get('gotenberg-output-filename')?.trim()
      return answer(results.length === 1 ? fileResponse(results[0], outName) : await zipResponse(results, outName))
    } catch (e) {
      return answer(errorResponse(e, gotenberg))
    } finally {
      pending -= held
    }
  }
}

interface Upload {
  name: string
  data: Uint8Array
}

/** the files and fields of a request: multipart, or the body as one file (`name` from the query or Content-Disposition) */
async function readRequest(req: Request, url: URL, maxBytes: number, hold: (n: number) => void): Promise<{ files: Upload[]; fields: Map<string, string> }> {
  const length = Number(req.headers.get('content-length') ?? NaN)
  if (length > maxBytes) throw new HttpError(413, `The request is larger than ${maxBytes / MB} MB`)
  const body = await readBody(req, maxBytes, hold)
  const fields = new Map<string, string>()
  for (const [k, v] of url.searchParams) fields.set(k, v)
  const type = req.headers.get('content-type') ?? ''
  if (/^multipart\/form-data/i.test(type)) {
    let form: FormData
    try {
      form = await new Response(body as BlobPart, { headers: { 'Content-Type': type } }).formData()
    } catch {
      throw new HttpError(400, 'The multipart body cannot be read')
    }
    const files: Upload[] = []
    // (a field given twice is read as Gotenberg reads it: its first value; the form's over the query's)
    const own = new Map<string, string>()
    for (const [k, v] of form) {
      if (typeof v !== 'string') files.push({ name: v.name || 'document', data: new Uint8Array(await v.arrayBuffer()) })
      else if (!own.has(k)) own.set(k, v)
    }
    for (const [k, v] of own) fields.set(k, v)
    return { files, fields }
  }
  if (!body.byteLength) return { files: [], fields }
  const disposition = /filename\*?=(?:UTF-8'')?"?([^";]+)"?/i.exec(req.headers.get('content-disposition') ?? '')
  const name = fields.get('name') || (disposition ? safeDecode(disposition[1]) : 'document')
  return { files: [{ name, data: body }], fields }
}

const safeDecode = (s: string) => {
  try {
    return decodeURIComponent(s)
  } catch {
    return s
  }
}

/** the whole body, refused past `maxBytes` (also when no Content-Length announced it) */
async function readBody(req: Request, maxBytes: number, hold: (n: number) => void): Promise<Uint8Array> {
  if (!req.body) return new Uint8Array(0)
  const chunks: Uint8Array[] = []
  let size = 0
  const reader = req.body.getReader()
  for (;;) {
    const { done, value } = await reader.read()
    if (done) break
    size += value.byteLength
    try {
      hold(value.byteLength)
    } catch (e) {
      void reader.cancel()
      throw e
    }
    if (size > maxBytes) {
      void reader.cancel()
      throw new HttpError(413, `The request is larger than ${maxBytes / MB} MB`)
    }
    chunks.push(value)
  }
  const out = new Uint8Array(size)
  let at = 0
  for (const c of chunks) {
    out.set(c, at)
    at += c.byteLength
  }
  return out
}

/** `attachment` with an ASCII name for old clients and the real one in `filename*` */
const disposition = (name: string) => `attachment; filename="${name.replace(/[^\x20-\x7e]|["\\]/g, '_')}"; filename*=UTF-8''${encodeURIComponent(name)}`

function fileResponse(r: Converted, outName?: string): Response {
  const name = outName ? `${outName}.${r.ext}` : r.name
  const res = new Response(r.data as BlobPart, { headers: { 'Content-Type': r.type, 'Content-Disposition': disposition(name) } })
  for (const w of r.warnings) res.headers.append('Office-Convert-Warning', encodeURIComponent(w))
  return res
}

async function zipResponse(results: Converted[], outName?: string): Promise<Response> {
  const entries = new Map<string, Uint8Array>()
  const taken = new Set<string>()
  for (const r of results) {
    // two inputs of the same name: the second is report (2).docx (the ZIP writer takes names without their case)
    let name = r.name
    for (let n = 2; taken.has(name.toLowerCase()); n++) name = r.name.replace(/(\.[^.]*)?$/, ` (${n})$1`)
    taken.add(name.toLowerCase())
    entries.set(name, r.data)
  }
  const zip = await writeZip(entries)
  const res = new Response(zip, { headers: { 'Content-Type': 'application/zip', 'Content-Disposition': disposition(`${outName || 'converted'}.zip`) } })
  for (const r of results) for (const w of r.warnings) res.headers.append('Office-Convert-Warning', encodeURIComponent(`${r.name}: ${w}`))
  return res
}

/**
 * A boolean field as Gotenberg reads one (Go's ParseBool: 1, t, true, TRUE, True and their opposites, as they are);
 * a field that is none of them is refused
 */
function flag(fields: Map<string, string>, name: string): boolean {
  const v = fields.get(name)
  if (v === undefined || v === '') return false
  if (/^(1|t|T|true|TRUE|True)$/.test(v)) return true
  if (/^(0|f|F|false|FALSE|False)$/.test(v)) return false
  throw new HttpError(400, `${name} “${v}” is not true or false`)
}

function errorResponse(e: unknown, gotenberg = false): Response {
  const err = e as Error
  const status =
    e instanceof HttpError ? e.status
    : e instanceof ConvertError ? (e.code === 'not-implemented' ? 501 : e.code === 'too-large' ? 413 : 400)
    : e instanceof BusyError || e instanceof TimeoutError ? 503
    : e instanceof ThreadError ? 500
    : err?.name === 'AbortError' ? 499
    // the file could not be read (damaged, encrypted, not a format the engine knows): Gotenberg says 400
    : gotenberg ? 400
    : 422
  const res = new Response(`${err?.message ?? String(e)}\n`, { status, headers: { 'Content-Type': 'text/plain; charset=utf-8' } })
  if (status === 503 && !(e instanceof TimeoutError)) res.headers.set('Retry-After', '5')
  return res
}
