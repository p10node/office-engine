/**
 * What a file is when it is not what its name says (a PDF renamed .docx, a cut-short download, a web page saved as
 * .xls): the words for the error a user reads, and the formats of text files the engine cannot open.
 */

/** said for every encrypted Office file (a compound file holding an encrypted package, a Word / Excel password flag) */
export const PASSWORD_MESSAGE = 'This file is password-protected (encrypted): open it in Office and save it without a password'

const at = (b: Uint8Array, offset: number, text: string) => [...text].every((ch, i) => b[offset + i] === ch.charCodeAt(0))

/** what a binary file is by its first bytes, as "a PDF file"; null when it is not a format known here */
export function describeBinary(b: Uint8Array): string | null {
  if (at(b, 0, '%PDF')) return 'a PDF file'
  if (b[0] === 0x89 && at(b, 1, 'PNG')) return 'a PNG image'
  if (b[0] === 0xff && b[1] === 0xd8 && b[2] === 0xff) return 'a JPEG image'
  if (at(b, 0, 'GIF8')) return 'a GIF image'
  if (at(b, 0, 'RIFF') && at(b, 8, 'WEBP')) return 'a WebP image'
  if (at(b, 0, 'RIFF') && at(b, 8, 'WAVE')) return 'an audio file'
  if (at(b, 0, 'BM') && b.length > 14 && b[6] === 0 && b[7] === 0) return 'a BMP image'
  if ((at(b, 0, 'II') && b[2] === 0x2a && b[3] === 0) || (at(b, 0, 'MM') && b[2] === 0 && b[3] === 0x2a)) return 'a TIFF image'
  if (at(b, 4, 'ftyp')) return 'a video or audio file'
  if (at(b, 0, 'ID3')) return 'an audio file'
  if (at(b, 0, 'Rar!')) return 'a RAR archive'
  if (b[0] === 0x37 && b[1] === 0x7a && b[2] === 0xbc && b[3] === 0xaf) return 'a 7-Zip archive'
  if (at(b, 0, 'BZh') || (b[0] === 0xfd && at(b, 1, '7zXZ'))) return 'a compressed archive'
  if (at(b, 0, 'MZ')) return 'a Windows program'
  if (b[0] === 0x7f && at(b, 1, 'ELF')) return 'a program'
  if (at(b, 0, 'SQLite format 3')) return 'a SQLite database'
  return null
}

/** the text formats of Office files that are not read: [pattern in the file's start, what it is, what to do] */
const UNREAD_TEXT: [RegExp, string, string][] = [
  [/<Workbook[\s>][^>]*urn:schemas-microsoft-com:office:spreadsheet|progid="Excel\.Sheet"/i, 'an Excel 2003 XML spreadsheet (SpreadsheetML)', 'open it in Excel and save it as .xlsx'],
  [/<w:wordDocument[\s>]|progid="Word\.Document"/i, 'a Word 2003 XML document', 'open it in Word and save it as .docx'],
  [/<pkg:package[\s>]/i, 'a flat OPC (Office XML) file', 'open it in Office and save it as .docx, .xlsx or .pptx'],
]

/** the error for a text file the engine cannot read as the Office file it is named, or null when it is not one of those */
export function unreadText(text: string): string | null {
  const start = text.slice(0, 4096)
  for (const [re, what, todo] of UNREAD_TEXT) if (re.test(start)) return `This is ${what}, which cannot be opened here: ${todo}`
  return null
}

/**
 * The message for a file that is not a ZIP package, an Office compound file or a text file, though its name says it is
 * a document: it says what the file is (or that it is empty or cut short) and what to do.
 */
export function notOffice(b: Uint8Array, size: number): string {
  if (!size) return 'The file is empty'
  if (b[0] === 0x50 && b[1] === 0x4b) return 'The file is damaged: it is a ZIP archive that ends too soon (an incomplete download or copy?)'
  const text = /^(?:\uFEFF|\s)*<(?:!doctype\s+html|html|head|body|\?xml|[a-z][\w:-]*[\s>])/i.test(new TextDecoder().decode(b.subarray(0, 512)))
  const unread = unreadText(new TextDecoder().decode(b))
  if (unread) return unread
  if (text && !/[\x00-\x08]/.test(new TextDecoder('latin1').decode(b.subarray(0, 512)))) {
    return 'This is a web page or an XML text, not a Word, Excel or PowerPoint document: the address may be wrong, or the server sent an error page instead of the file'
  }
  const what = describeBinary(b)
  if (what) return `This is ${what}, not a Word, Excel or PowerPoint document: open the file with a program for that format, or, if it should be a document, get the file again`
  return 'The file is not in a format that can be opened here (Word, Excel, PowerPoint, OpenDocument, RTF, Pages, Keynote, Numbers, Visio, e-mail, CSV, text, Markdown, HTML…): its content does not match its name, or it is damaged'
}
