import { decodeText, extOf, PLAIN_EXT, PLAIN_KIND, plainToOffice } from '../../core/plain'

// .pages / .key / .numbers (Apple iWork) open as read-only Word / PowerPoint / Excel files
const OFFICE = new Set(['docx', 'dotx', 'xlsx', 'xltx', 'pptx', 'potx', 'pages', 'key', 'numbers'])

/** `accept` of the File → Open picker */
export const OPEN_ACCEPT = [...OFFICE, ...Object.keys(PLAIN_EXT)].map((e) => '.' + e).join(',')

export interface Imported {
  src: Blob
  name: string
  /** built from another format: not saved as an Office file yet */
  converted: boolean
}

/** Office and iWork (Pages, Keynote, Numbers) files open as they are; CSV / TSV become a workbook, text / Markdown / HTML a document. */
export async function importFile(file: File): Promise<Imported> {
  const ext = extOf(file.name)
  const base = file.name.replace(/\.[^./\\]+$/, '') || 'document'
  if (OFFICE.has(ext)) return { src: file, name: file.name, converted: false }
  const format = PLAIN_EXT[ext]
  if (format) {
    const text = decodeText(new Uint8Array(await file.arrayBuffer()), file.type)
    return { src: await plainToOffice(text, format, file.name), name: `${base}.${PLAIN_KIND[format]}`, converted: true }
  }
  // no / unknown extension: an Office package is a ZIP
  const sig = new Uint8Array(await file.slice(0, 4).arrayBuffer())
  if (sig[0] === 0x50 && sig[1] === 0x4b) return { src: file, name: file.name, converted: false }
  throw new Error(`Cannot open “${file.name}”: supported files are Word, Excel, PowerPoint, Pages, Keynote, Numbers, CSV, text, Markdown and HTML`)
}
