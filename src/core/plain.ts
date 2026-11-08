/**
 * Plain formats opened as Office files: CSV / TSV become a workbook, text / Markdown / HTML a document.
 * Used by the engine (a source that is not an Office package) and by File → Open.
 */
import { detectSeparator, markdownToHtml, parasToDocx, parseCsv, rowsToXlsx, textToParas, type ImportPara } from './convert'
import { htmlToParas } from './html'

export type PlainFormat = 'csv' | 'tsv' | 'text' | 'markdown' | 'html'

export const PLAIN_EXT: Record<string, PlainFormat> = {
  csv: 'csv',
  tsv: 'tsv',
  tab: 'tsv',
  txt: 'text',
  text: 'text',
  log: 'text',
  md: 'markdown',
  markdown: 'markdown',
  html: 'html',
  htm: 'html',
  xhtml: 'html',
}

/** packages (Office, iWork) and the legacy binary formats: never read as text */
const PACKAGE_EXT = new Set([
  'docx', 'dotx', 'docm', 'dotm', 'xlsx', 'xltx', 'xlsm', 'xltm', 'pptx', 'potx', 'pptm', 'potm', 'ppsx', 'ppsm',
  'pages', 'key', 'numbers', 'doc', 'dot', 'xls', 'xlt', 'ppt', 'pot', 'pps',
])

const MIME: Record<string, PlainFormat> = {
  'text/csv': 'csv',
  'application/csv': 'csv',
  'text/comma-separated-values': 'csv',
  'text/tab-separated-values': 'tsv',
  'text/plain': 'text',
  'text/markdown': 'markdown',
  'text/x-markdown': 'markdown',
  'text/html': 'html',
  'application/xhtml+xml': 'html',
}

/** what a plain format opens as */
export const PLAIN_KIND: Record<PlainFormat, 'xlsx' | 'docx'> = { csv: 'xlsx', tsv: 'xlsx', text: 'docx', markdown: 'docx', html: 'docx' }

export const extOf = (name?: string) => /\.([^./\\]+)$/.exec(name ?? '')?.[1].toLowerCase() ?? ''

/** false for a name that says Office / iWork package: no need to look at its bytes */
export const mayBePlain = (name?: string) => !PACKAGE_EXT.has(extOf(name))

const charsetOf = (type?: string) => /;\s*charset=["']?([^;"'\s]+)/i.exec(type ?? '')?.[1]

/** no control characters a text file would not have (UTF-16 has a byte order mark) */
function isText(b: Uint8Array): boolean {
  if ((b[0] === 0xff && b[1] === 0xfe) || (b[0] === 0xfe && b[1] === 0xff)) return true
  if (b[0] === 0x25 && b[1] === 0x50 && b[2] === 0x44 && b[3] === 0x46) return false // %PDF
  for (const x of b) if (x < 9 || (x > 13 && x < 32 && x !== 27) || x === 0x7f) return false
  return true
}

/** Text of a file: byte order mark, then the declared charset, then UTF-8, then Windows-1252 (Excel's CSV on Windows). */
export function decodeText(b: Uint8Array, type?: string): string {
  if (b[0] === 0xef && b[1] === 0xbb && b[2] === 0xbf) return new TextDecoder('utf-8').decode(b)
  if (b[0] === 0xff && b[1] === 0xfe) return new TextDecoder('utf-16le').decode(b)
  if (b[0] === 0xfe && b[1] === 0xff) {
    const swapped = new Uint8Array(b.length - 2)
    for (let k = 2; k + 1 < b.length; k += 2) {
      swapped[k - 2] = b[k + 1]
      swapped[k - 1] = b[k]
    }
    return new TextDecoder('utf-16le').decode(swapped)
  }
  const charset = charsetOf(type)
  if (charset) {
    try {
      return new TextDecoder(charset).decode(b)
    } catch {
      // unknown label: guess below
    }
  }
  try {
    return new TextDecoder('utf-8', { fatal: true }).decode(b)
  } catch {
    return new TextDecoder('windows-1252').decode(b)
  }
}

/** a file with neither a known extension nor a MIME type: HTML, delimited rows, or text */
function sniff(text: string, whole: boolean): PlainFormat {
  // trimStart also drops a byte order mark
  const t = text.trimStart()
  if (/^<(!doctype\s+html|html|head|body)[\s>]/i.test(t) || (t[0] === '<' && /<\/(p|div|h[1-6]|table|ul|ol|li|span|a|b|i|em|strong)>/i.test(t))) return 'html'
  const lines = t.split(/\r\n|\r|\n/)
  // the head ends mid-line
  if (!whole) lines.pop()
  const rows = lines.filter((l) => l.trim()).slice(0, 20)
  if (rows.length >= 2) {
    const sep = detectSeparator(rows[0])
    const cols = (l: string) => parseCsv(l, sep)[0]?.length ?? 0
    const n = cols(rows[0])
    if (n >= 2 && rows.every((l) => cols(l) === n)) return sep === '\t' ? 'tsv' : 'csv'
  }
  return 'text'
}

/**
 * The plain format of a file from its name, MIME type and first bytes (`whole`: they are the whole file), or null
 * for an Office / iWork package or a binary file.
 */
export function plainFormat(name: string | undefined, type: string | undefined, head: Uint8Array, whole: boolean): PlainFormat | null {
  // ZIP (a package, whatever its name says); OLE: legacy binary Office or encrypted, the ZIP reader says so
  if ((head[0] === 0x50 && head[1] === 0x4b) || (head[0] === 0xd0 && head[1] === 0xcf)) return null
  const ext = extOf(name)
  if (PACKAGE_EXT.has(ext)) return null
  const byExt = PLAIN_EXT[ext]
  if (!isText(head)) {
    if (byExt) throw new Error(`“${name}” is not a text file`)
    return null
  }
  if (byExt) return byExt
  if (!head.length) return null
  const mime = MIME[(type ?? '').split(';')[0].trim().toLowerCase()]
  return mime ?? sniff(new TextDecoder().decode(head), whole)
}

const fromHtml = (html: string): ImportPara[] =>
  htmlToParas(html).map((p) => ({
    heading: p.heading,
    list: p.list,
    runs: p.runs.flatMap((r) => (r.t === 'text' ? [{ text: r.text, set: r.d, href: r.href }] : [])),
  }))

/** Builds the workbook (CSV, TSV) or document (text, Markdown, HTML) a plain file opens as; `name` names the sheet. */
export function plainToOffice(text: string, format: PlainFormat, name = ''): Promise<Blob> {
  switch (format) {
    case 'csv':
    case 'tsv': {
      const sep = format === 'tsv' ? '\t' : detectSeparator(text)
      return rowsToXlsx(parseCsv(text, sep), name.replace(/\.[^./\\]+$/, '') || 'Sheet1', sep === ';')
    }
    case 'text':
      return parasToDocx(textToParas(text))
    case 'markdown':
      return parasToDocx(fromHtml(markdownToHtml(text)))
    case 'html':
      return parasToDocx(fromHtml(text))
  }
}
