import type { DocKind } from '../opc'
import type { DocComment } from '../comments'
import type { XOp } from '../../collab/xlsx-ops'
import type { POp } from '../../collab/pptx-ops'
import type { PptxCollabResult } from '../../collab/pptx-worker'
import type { CanonStyle, DocxMeta, Block, EditOp, EditResult, PageSet, PageSetResult, ParaSpec, StyleInfo, StyleLook, StyleSetResult, TableOp } from '../docx/types'
import type { BlockOut, BookOp, CellDetail, CellInput, FindMatch, FindOptions, FmtPatch, Range, SheetOp, SheetProgress, WorkbookMeta, XlsxEditResult } from '../xlsx/types'
import type { PParaSpec, PresentationMeta, ShapeOp, SlideModel, SlideOp } from '../pptx/types'

export interface OpenParams {
  name?: string
  headers?: Record<string, string>
  credentials?: RequestCredentials
}

export interface OpenResult {
  kind: DocKind
  name?: string
  size: number
  /** the file was converted from another format (shown read-only): Apple Pages (as docx), Keynote (as pptx), Numbers (as xlsx) */
  from?: 'pages' | 'keynote' | 'numbers'
  xlsx?: WorkbookMeta
  pptx?: PresentationMeta
}

export interface Methods {
  open(src: Blob | ArrayBuffer | string, params: OpenParams): OpenResult
  part(path: string): Blob
  'docx.start'(): void
  'docx.edit'(op: EditOp): EditResult
  'docx.table'(op: TableOp): void
  'docx.image'(bytes: ArrayBuffer, ext: string): { rid: string; path: string }
  'docx.link'(url: string): string
  /** id of a built-in paragraph style, created in styles.xml when missing */
  'docx.style'(kind: CanonStyle): { id: string; styles: StyleInfo[] }
  /** gives paragraph style `id` a look ("update style to match"); only the document's own styles part is written */
  'docx.styleSet'(id: string, look: StyleLook): StyleSetResult
  /** changes the page setup of every section (size, orientation, margins, columns) */
  'docx.page'(set: PageSet): PageSetResult
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
  /** cells whose shown value (or formula) contains the text */
  'xlsx.find'(sheet: number, query: string, opts?: FindOptions): FindMatch[]
  /** replaces text in matching cells (or only `at`); null when nothing changed */
  'xlsx.replace'(sheet: number, query: string, by: string, opts?: FindOptions, at?: { r: number; c: number }): XlsxEditResult | null
  'xlsx.set'(sheet: number, cells: CellInput[]): XlsxEditResult
  'xlsx.format'(sheet: number, range: Range, patch: FmtPatch): XlsxEditResult
  'xlsx.sheetOp'(sheet: number, op: SheetOp): XlsxEditResult
  'xlsx.bookOp'(op: BookOp): XlsxEditResult
  'xlsx.undo'(): XlsxEditResult | null
  'xlsx.redo'(): XlsxEditResult | null
  'pptx.slide'(index: number): SlideModel
  'pptx.setText'(index: number, sid: number, paras: PParaSpec[]): SlideModel
  'pptx.shapes'(index: number, ops: ShapeOp[]): SlideModel
  'pptx.slides'(op: SlideOp): PresentationMeta
  'pptx.image'(index: number, bytes: ArrayBuffer, ext: string): string
  /** undo (or redo) the last shape edit */
  'pptx.undo'(redo?: boolean): { index: number; model: SlideModel } | null
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
  /** a comment was set or deleted (by this client or a collaborator) */
  comment: { id: string; c: DocComment | null }
}

export type Req = { id: number; method: keyof Methods; args: unknown[] }
export type Res = { id: number; ok: true; result: unknown } | { id: number; ok: false; error: string }
export type Evt = { event: keyof Events; data: unknown }
