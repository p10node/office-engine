import { VISIO_EXT } from '../../core/visio'
import { MAIL_EXT } from '../../core/mail/types'
import { describeBinary } from '../../core/filetype'
import { extOf, PLAIN_EXT, plainFormat } from '../../core/plain'
import { LEGACY_EXT, legacyFormatOf } from '../../core/legacy'

// .pages / .key / .numbers (Apple iWork) open as read-only Word / PowerPoint / Excel files, Visio drawings as read-only presentations,
// mail messages (.eml, .emlx, .msg, .oft) as read-only documents
const OFFICE = new Set(['docx', 'dotx', 'xlsx', 'xltx', 'pptx', 'potx', 'pages', 'key', 'numbers', ...VISIO_EXT, ...Object.keys(MAIL_EXT)])

/** `accept` of the File → Open picker */
export const OPEN_ACCEPT = [...OFFICE, ...Object.keys(LEGACY_EXT), ...Object.keys(PLAIN_EXT), 'gz'].map((e) => '.' + e).join(',')

export interface Imported {
  src: Blob
  name: string
  /** built from another format: not saved as an Office file yet */
  converted: boolean
  /** a preview of the source being edited (the shown document stays until it is ready) */
  live?: boolean
}

/**
 * Office, iWork (Pages, Keynote, Numbers), Visio files and mail messages open as they are; Word / Excel / PowerPoint 97-2003, RTF and
 * OpenDocument files are converted to Office Open XML; CSV / TSV become a workbook, text / Markdown / HTML, SVG and
 * Mermaid / PlantUML / Graphviz diagrams a document, also gzip-compressed (.svgz, .csv.gz).
 */
export async function importFile(file: File, unsupported: (name: string, what?: string) => string): Promise<Imported> {
  const ext = extOf(file.name)
  if (OFFICE.has(ext)) return { src: file, name: file.name, converted: false }
  const head = new Uint8Array(await file.slice(0, 4096).arrayBuffer())
  // .doc / .xls / .ppt, RTF, OpenDocument (a ZIP), whatever the name says when the bytes are a compound file, RTF or flat XML
  if ((LEGACY_EXT[ext] && ext !== 'dot') || legacyFormatOf(head)) return { src: file, name: file.name, converted: true }
  // by the extension (also the one inside a .gz name), else the MIME type, else the content (a text file with no extension is
  // text, as it is when it comes through `src`); the bytes deciding (a binary .dot is a Word 97 template, not a Graphviz graph)
  const format = plainFormat(file.name, file.type, head, file.size <= 4096)
  // converted by the engine, like any src: the document keeps the file's name, format and source
  if (format) return { src: file, name: file.name, converted: true }
  // no / unknown extension: an Office package is a ZIP
  if (head[0] === 0x50 && head[1] === 0x4b) return { src: file, name: file.name, converted: false }
  throw new Error(unsupported(file.name, describeBinary(head) ?? undefined))
}
