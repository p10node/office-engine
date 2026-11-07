import { detectSeparator, markdownToHtml, parasToDocx, parseCsv, rowsToXlsx, textToParas, type ImportPara } from '../../core/convert'
import { htmlToParas } from '../docx/html'

// .pages / .key / .numbers (Apple iWork) open as read-only Word / PowerPoint / Excel files
const OFFICE = new Set(['docx', 'dotx', 'xlsx', 'xltx', 'pptx', 'potx', 'pages', 'key', 'numbers'])
const TABLE = new Set(['csv', 'tsv', 'tab'])
const TEXT = new Set(['txt', 'text', 'log'])
const MARKDOWN = new Set(['md', 'markdown'])
const HTML = new Set(['html', 'htm', 'xhtml'])

/** `accept` of the File → Open picker */
export const OPEN_ACCEPT = [...OFFICE, ...TABLE, ...TEXT, ...MARKDOWN, ...HTML].map((e) => '.' + e).join(',')

export interface Imported {
  src: Blob
  name: string
  /** built from another format: not saved as an Office file yet */
  converted: boolean
}

const fromHtml = (html: string): ImportPara[] =>
  htmlToParas(html).map((p) => ({
    heading: p.heading,
    list: p.list,
    runs: p.runs.flatMap((r) => (r.t === 'text' ? [{ text: r.text, set: r.d, href: r.href }] : [])),
  }))

/** Office and iWork (Pages, Keynote, Numbers) files open as they are; CSV / TSV become a workbook, text / Markdown / HTML a document. */
export async function importFile(file: File): Promise<Imported> {
  const ext = /\.([^./\\]+)$/.exec(file.name)?.[1].toLowerCase() ?? ''
  const base = file.name.replace(/\.[^./\\]+$/, '') || 'document'
  if (OFFICE.has(ext)) return { src: file, name: file.name, converted: false }
  if (TABLE.has(ext)) {
    const text = await file.text()
    const sep = ext === 'csv' ? detectSeparator(text) : '\t'
    return { src: await rowsToXlsx(parseCsv(text, sep), base, sep === ';'), name: `${base}.xlsx`, converted: true }
  }
  let paras: ImportPara[] | undefined
  if (TEXT.has(ext)) paras = textToParas(await file.text())
  else if (MARKDOWN.has(ext)) paras = fromHtml(markdownToHtml(await file.text()))
  else if (HTML.has(ext)) paras = fromHtml(await file.text())
  if (paras) return { src: await parasToDocx(paras), name: `${base}.docx`, converted: true }
  // no / unknown extension: an Office package is a ZIP
  const sig = new Uint8Array(await file.slice(0, 4).arrayBuffer())
  if (sig[0] === 0x50 && sig[1] === 0x4b) return { src: file, name: file.name, converted: false }
  throw new Error(`Cannot open “${file.name}”: supported files are Word, Excel, PowerPoint, Pages, Keynote, Numbers, CSV, text, Markdown and HTML`)
}
