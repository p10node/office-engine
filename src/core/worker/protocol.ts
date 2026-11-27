import type { DocKind } from '../opc'
import type { DocComment } from '../comments'
import type { PlainFormat } from '../plain'
import type { LegacyFormat } from '../legacy'
import type { XOp } from '../../collab/xlsx-ops'
import type { POp } from '../../collab/pptx-ops'
import type { PptxCollabResult } from '../../collab/pptx-worker'
import type { CanonStyle, DocxMeta, Block, EditOp, EditResult, HFSet, HFSetResult, PageSet, PageSetResult, ParaSpec, StyleInfo, StyleLook, StyleSetResult, TableOp } from '../docx/types'
import type { BlockOut, BookOp, CellDetail, CellInput, FindMatch, FindOptions, FmtPatch, Range, SetOptions, SheetOp, SheetProgress, WorkbookMeta, XlsxClip, XlsxEditResult } from '../xlsx/types'
import type { SheetDrawing } from '../xlsx/drawings'
import type { PParaSpec, PresentationMeta, ShapeOp, SlideModel, SlideOp } from '../pptx/types'
import type { FontSource } from '../fonts'
import type { TextMeasure, TextSpec } from '../font/book'

/** a text measured in the font files of `fonts.set`: `TextMeasure` as plain data (its pieces without their font and glyphs) */
export interface MeasuredText extends Omit<TextMeasure, 'pieces'> {
  pieces: { from: number; to: number; family: string; width: number }[]
}

export interface OpenParams {
  name?: string
  headers?: Record<string, string>
  credentials?: RequestCredentials
  /**
   * CSV / TSV: cells kept in memory at a time (default 12 million, about 300 MB); the rows past them are read from
   * the file as they are shown, edited, searched or saved
   */
  residentCells?: number
  /** CSV / TSV: rows read from the file at a time past `residentCells` (default 1024) */
  pageRows?: number
  /**
   * Text files (CSV, TSV, text, Markdown, HTML…): the encoding the bytes are read in (a `TextDecoder` label such as
   * `windows-1258`), instead of the byte order mark, the declared charset or a guess (UTF-8, else Windows-1258 for
   * Vietnamese text, else Windows-1252)
   */
  encoding?: string
  /**
   * The file is source code in this language ("typescript", "py", "c++"…): a text file opens in the code view,
   * colored in it, whatever its name and type say. Without it, the name (main.ts) or MIME type tells a source file.
   */
  language?: string
}

export interface OpenResult {
  kind: DocKind
  name?: string
  size: number
  /** the file was converted from another format (shown read-only): Apple Pages (as docx), Keynote (as pptx), Numbers (as xlsx), Microsoft Visio (as pptx) */
  from?: 'pages' | 'keynote' | 'numbers' | 'visio'
  /** built from a plain file, editable: CSV / TSV (as xlsx), text / Markdown / HTML (as docx); saving writes the Office format */
  imported?: PlainFormat
  /** a source file (`imported` 'code'): its language ("typescript", "python"…; a hint's own word when the language is not known) */
  language?: string
  /**
   * converted from another format, editable: Word / Excel / PowerPoint 97-2003 (doc, xls, ppt), RTF, OpenDocument
   * (odt, ods, odp); saving writes the Office Open XML file (docx, xlsx, pptx)
   */
  converted?: LegacyFormat
  /**
   * The file is not the plain document of its kind but a variant: with macros (docm, xlsm, pptm), a template (dotx,
   * xltx, potx and their macro forms), a slide show (ppsx, ppsm). Saving keeps it that: this is the extension of the
   * saved file (the macros, which the engine does not read, are kept in it).
   */
  ext?: string
  xlsx?: WorkbookMeta
  pptx?: PresentationMeta
}

export interface Methods {
  open(src: Blob | ArrayBuffer | string, params: OpenParams): OpenResult
  part(path: string): Blob
  /** the text of a plain file opened as a document (Markdown, HTML, diagrams, source code), null otherwise */
  'plain.source'(): string | null
  /** a source file's bytes as they were opened (saved unchanged when its code was not edited), null otherwise */
  'plain.original'(): Blob | null
  'docx.start'(): void
  'docx.edit'(op: EditOp): EditResult
  'docx.table'(op: TableOp): EditResult
  'docx.image'(bytes: ArrayBuffer, ext: string): { rid: string; path: string }
  'docx.link'(url: string): string
  /** id of a built-in paragraph style, created in styles.xml when missing */
  'docx.style'(kind: CanonStyle): { id: string; styles: StyleInfo[] }
  /** gives paragraph style `id` a look ("update style to match"); only the document's own styles part is written */
  'docx.styleSet'(id: string, look: StyleLook): StyleSetResult
  /** changes the page setup of every section (size, orientation, margins, columns) */
  'docx.page'(set: PageSet): PageSetResult
  /** changes the headers / footers (text, page numbers, different first page) */
  'docx.hf'(set: HFSet): HFSetResult
  /** the blocks each paragraph converts to with the styles as they are now; `bare`: as it looks outside its table */
  'docx.convert'(specs: ParaSpec[], bare?: boolean): Block[][]
  /** the document's comments (threads: roots and their replies) */
  'comments.list'(): DocComment[]
  /** sets comment `id` (null deletes it; a root's replies are left out of the file with it); returns what it was */
  'comments.set'(id: string, c: DocComment | null): DocComment | null
  /** builds the edited file */
  save(): Blob
  dirty(): boolean
  'xlsx.load'(sheet: number): void
  'xlsx.block'(sheet: number, r0: number, r1: number): BlockOut
  'xlsx.cell'(sheet: number, r: number, c: number): CellDetail | undefined
  /** the sheet's pictures and charts, where they are anchored */
  'xlsx.drawings'(sheet: number): SheetDrawing[]
  /** cells whose shown value (or formula) contains the text */
  'xlsx.find'(sheet: number, query: string, opts?: FindOptions): FindMatch[]
  /** the sheet as delimited text (the values as shown, a General number with its digits; texts a spreadsheet would run as formulas with a `'` in front), after a byte order mark: `sep` a comma, a semicolon or a tab */
  'xlsx.text'(sheet: number, sep: string): Blob
  /** replaces text in matching cells (or only `at`); null when nothing changed */
  'xlsx.replace'(sheet: number, query: string, by: string, opts?: FindOptions, at?: { r: number; c: number }): XlsxEditResult | null
  /** what was typed into cells; `opts.locale`: the BCP 47 tag numbers were typed in (`1,5` in `de-DE`), a decimal point without one */
  'xlsx.set'(sheet: number, cells: CellInput[], opts?: SetOptions): XlsxEditResult
  /** the block of filled cells a range is part of (the table around it), and whether its first row reads as a header */
  'xlsx.region'(sheet: number, range: Range): { range: Range; header: boolean }
  /** the cells of a range as they are (values, formulas, formats) for the `paste` sheet operation, and as text for the clipboard; rows not shown yet included */
  'xlsx.copy'(sheet: number, range: Range): XlsxClip
  'xlsx.format'(sheet: number, range: Range, patch: FmtPatch): XlsxEditResult
  'xlsx.sheetOp'(sheet: number, op: SheetOp): XlsxEditResult
  'xlsx.bookOp'(op: BookOp): XlsxEditResult
  /**
   * a new sheet after `after` holding the rows of a delimited text file (CSV / TSV), streamed in when the sheet is
   * loaded; `sep` is detected from the first line when missing, `name` defaults to the file's
   */
  'xlsx.import'(after: number, src: Blob | ArrayBuffer | string, params?: { name?: string; sep?: string; residentCells?: number; pageRows?: number; encoding?: string }): XlsxEditResult
  'xlsx.undo'(): XlsxEditResult | null
  'xlsx.redo'(): XlsxEditResult | null
  'pptx.slide'(index: number): SlideModel
  /** every slide's title placeholder text ('' where there is none), read without laying out the slides */
  'pptx.titles'(): string[]
  /** `cell`: [row, column] of a table cell (`sid` is then the table's graphic frame) */
  /** `step`: the edit is an undo step of its own (the text was just entered), not one with the edit before it of the same text */
  'pptx.setText'(index: number, sid: number, paras: PParaSpec[], cell?: [number, number], step?: boolean): SlideModel
  'pptx.shapes'(index: number, ops: ShapeOp[]): SlideModel
  'pptx.slides'(op: SlideOp): PresentationMeta
  'pptx.image'(index: number, bytes: ArrayBuffer, ext: string): string
  /** undo (or redo) the last edit; `meta` when it changed the slide list */
  'pptx.undo'(redo?: boolean): { index: number; model: SlideModel; meta?: PresentationMeta } | null
  /** how many edits `pptx.undo` can take back, and bring back */
  'pptx.history'(): { undo: number; redo: number }
  /**
   * The font files text is measured in (the list DocViewer's `fonts` takes), instead of any set before. Nothing is
   * fetched until text asks for a file. No document has to be open.
   */
  'fonts.set'(fonts: FontSource[]): void
  /**
   * Each text measured from the font files themselves (glyph widths, kerning, ligatures, line metrics): the same
   * numbers in every browser and on a server. `exact` is false, and `unsure` says why, where a browser would set the
   * text differently or no declared file serves it.
   */
  'fonts.measure'(texts: TextSpec[]): MeasuredText[]
  /** collaboration (src/collab): a shared XLSX operation, outside the undo history; `rec` reverts it */
  'collab.xlsx.apply'(op: XOp): { res: XlsxEditResult | null; rec: number | null }
  'collab.xlsx.revert'(rec: number): XlsxEditResult | null
  /** collaboration: a shared PPTX operation (recorded as slide snapshots) */
  'collab.pptx.apply'(op: POp): { res: PptxCollabResult; rec: number | null }
  'collab.pptx.revert'(rec: number): PptxCollabResult
}

export interface Events {
  progress: { phase: string; loaded: number; total: number }
  'docx:blocks': Block[]
  'docx:meta': DocxMeta
  'docx:done': { error?: string }
  'xlsx:progress': SheetProgress
  /** collaboration: remote (or rebased) XLSX changes applied on this client */
  'collab:xlsx': XlsxEditResult
  /** collaboration: a slide (or the slide list) changed remotely */
  'collab:pptx': PptxCollabResult
  /** the document opened, but not whole: what is missing from it (more columns in a CSV than a sheet holds…) */
  warning: { text: string }
  /** a font file of `fonts.set` could not be fetched or read: text asking for it is measured in the next font of its list */
  'font:error': { family: string; message: string }
  /** a comment was set or deleted (by this client or a collaborator) */
  comment: { id: string; c: DocComment | null }
}

export type Req = { id: number; method: keyof Methods; args: unknown[] }
export type Res = { id: number; ok: true; result: unknown } | { id: number; ok: false; error: string }
export type Evt = { event: keyof Events; data: unknown }
