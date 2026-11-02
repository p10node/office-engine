import { DocxDocument } from '../docx/parser'
import { addImage, addLink, applyEdit, applyTable, ensureStyle, saveDocx } from '../docx/writer'
import type { CanonStyle, EditOp, TableOp } from '../docx/types'
import { OpcPackage, type DocKind } from '../opc'
import { PptxPresentation } from '../pptx/parser'
import { openSource, type DocSource } from '../source'
import { XlsxWorkbook } from '../xlsx/workbook'
import { XlsxEditor } from '../xlsx/editor'
import type { BookOp, CellInput, FindOptions, FmtPatch, Range, SheetOp } from '../xlsx/types'
import type { PParaSpec, ShapeOp, SlideOp } from '../pptx/types'
import { addSlideImage, shapeOps, slideOp, undoShapes } from '../pptx/edit'
import { ZipArchive } from '../zip'
import { XlsxCollabWorker } from '../../collab/xlsx-worker'
import type { XOp } from '../../collab/xlsx-ops'
import { PptxCollabWorker } from '../../collab/pptx-worker'
import type { POp } from '../../collab/pptx-ops'
import type { Events, Methods, OpenParams, OpenResult } from './protocol'

const WRITES = new Set<string>([
  'docx.edit', 'docx.table', 'docx.image', 'docx.link', 'docx.style', 'save',
  'xlsx.set', 'xlsx.format', 'xlsx.sheetOp', 'xlsx.bookOp', 'xlsx.replace', 'xlsx.undo', 'xlsx.redo', 'collab.xlsx.apply', 'collab.xlsx.revert',
  'pptx.setText', 'pptx.shapes', 'pptx.slides', 'pptx.image', 'pptx.undo', 'collab.pptx.apply', 'collab.pptx.revert',
])

type Emit = <K extends keyof Events>(event: K, data: Events[K]) => void

/** Format-agnostic engine; runs inside the worker (or directly, e.g. in tests). */
export class Engine {
  private pkg?: OpcPackage
  private kind?: DocKind
  private docx?: DocxDocument
  private xlsx?: XlsxWorkbook
  private xlsxEd?: XlsxEditor
  private pptx?: PptxPresentation
  private abort = new AbortController()

  constructor(private emit: Emit) {}

  async open(src: DocSource, params: OpenParams = {}): Promise<OpenResult> {
    const source = await openSource(src, {
      ...params,
      signal: this.abort.signal,
      onProgress: (loaded, total) => this.emit('progress', { phase: 'download', loaded, total }),
    })
    const zip = await ZipArchive.open(source)
    const pkg = (this.pkg = new OpcPackage(zip))
    const main = await pkg.mainPart()
    this.kind = main.kind
    const res: OpenResult = { kind: main.kind, name: params.name ?? source.name, size: source.size }
    if (main.kind === 'docx') this.docx = await DocxDocument.open(pkg, main.path)
    else if (main.kind === 'xlsx') res.xlsx = (this.xlsx = await XlsxWorkbook.open(pkg, main.path)).meta()
    else res.pptx = (this.pptx = await PptxPresentation.open(pkg, main.path)).meta
    return res
  }

  private writes: Promise<unknown> = Promise.resolve()

  call<M extends keyof Methods>(method: M, args: unknown[]): Promise<unknown> {
    // document writes run one at a time, in call order: they await internally (list / style
    // parts, sheet loading) and a later edit reads state (paragraph locations…) the previous one sets
    if (!WRITES.has(method)) return this.dispatch(method, args)
    const run = this.writes.then(() => this.dispatch(method, args))
    this.writes = run.catch(() => {})
    return run
  }

  private async dispatch<M extends keyof Methods>(method: M, args: unknown[]): Promise<unknown> {
    switch (method) {
      case 'open':
        return this.open(args[0] as DocSource, args[1] as OpenParams)
      case 'part':
        return this.need(this.pkg).blob(args[0] as string)
      case 'docx.start': {
        const doc = this.need(this.docx)
        try {
          await doc.stream({
            signal: this.abort.signal,
            onBlocks: (b) => this.emit('docx:blocks', b),
            onMeta: (m) => this.emit('docx:meta', m),
            onProgress: (loaded, total) => this.emit('progress', { phase: 'parse', loaded, total }),
          })
          this.emit('docx:done', {})
        } catch (e) {
          this.emit('docx:done', { error: (e as Error).message })
        }
        return
      }
      case 'docx.edit':
        return applyEdit(this.need(this.docx), args[0] as EditOp)
      case 'docx.table':
        return applyTable(this.need(this.docx), args[0] as TableOp)
      case 'docx.image':
        return addImage(this.need(this.docx), new Uint8Array(args[0] as ArrayBuffer), args[1] as string)
      case 'docx.link':
        return addLink(this.need(this.docx), args[0] as string)
      case 'docx.style':
        return ensureStyle(this.need(this.docx), args[0] as CanonStyle)
      case 'save':
        return this.save()
      case 'dirty':
        return this.docx ? this.docx.dirty : this.xlsxEd ? this.xlsxEd.dirty : this.pptx ? this.pptx.dirty : this.need(this.pkg).dirty
      case 'xlsx.load': {
        const wb = this.need(this.xlsx)
        void wb.loadSheet(
          args[0] as number,
          (p) => this.emit('xlsx:progress', p),
          (phase, loaded, total) => this.emit('progress', { phase, loaded, total }),
        )
        return
      }
      case 'xlsx.set':
        return this.xed().setCells(args[0] as number, args[1] as CellInput[])
      case 'xlsx.format':
        return this.xed().format(args[0] as number, args[1] as Range, args[2] as FmtPatch)
      case 'xlsx.sheetOp':
        return this.xed().sheetOp(args[0] as number, args[1] as SheetOp)
      case 'xlsx.replace':
        return this.xed().replace(args[0] as number, args[1] as string, args[2] as string, args[3] as FindOptions, args[4] as { r: number; c: number } | undefined)
      case 'xlsx.bookOp':
        return this.xed().bookOp(args[0] as BookOp)
      case 'xlsx.undo':
        return this.xed().undo()
      case 'xlsx.redo':
        return this.xed().redo()
      case 'collab.xlsx.apply':
        return this.xc().apply(args[0] as XOp)
      case 'collab.xlsx.revert':
        return this.xc().revert(args[0] as number)
      case 'collab.pptx.apply':
        return (this.pcollab ??= new PptxCollabWorker(this.need(this.pptx))).apply(args[0] as POp)
      case 'collab.pptx.revert':
        return (this.pcollab ??= new PptxCollabWorker(this.need(this.pptx))).revert(args[0] as number)
      case 'xlsx.block':
        return this.need(this.xlsx).block(args[0] as number, args[1] as number, args[2] as number)
      case 'xlsx.cell':
        return this.need(this.xlsx).cell(args[0] as number, args[1] as number, args[2] as number)
      case 'xlsx.find':
        return this.need(this.xlsx).find(args[0] as number, args[1] as string, args[2] as FindOptions)
      case 'pptx.setText':
        return this.need(this.pptx).setText(args[0] as number, args[1] as number, args[2] as PParaSpec[])
      case 'pptx.shapes':
        return shapeOps(this.need(this.pptx), args[0] as number, args[1] as ShapeOp[])
      case 'pptx.slides':
        return slideOp(this.need(this.pptx), args[0] as SlideOp)
      case 'pptx.image':
        return addSlideImage(this.need(this.pptx), args[0] as number, new Uint8Array(args[1] as ArrayBuffer), args[2] as string)
      case 'pptx.slide':
        return this.need(this.pptx).slide(args[0] as number)
      case 'pptx.undo':
        return undoShapes(this.need(this.pptx), !!args[0])
    }
    throw new Error(`Unknown method ${String(method)}`)
  }

  private xcollab?: XlsxCollabWorker
  private pcollab?: PptxCollabWorker
  private xc(): XlsxCollabWorker {
    return (this.xcollab ??= new XlsxCollabWorker(this.xed(), this.need(this.xlsx), (p) => this.emit('xlsx:progress', p)))
  }

  private xed(): XlsxEditor {
    return (this.xlsxEd ??= new XlsxEditor(this.need(this.xlsx)))
  }

  private async save(): Promise<Blob> {
    if (this.docx) return saveDocx(this.docx)
    if (this.xlsx) return this.xed().save()
    if (this.pptx) return this.pptx.save()
    throw new Error('Saving is not supported for this document yet')
  }

  dispose() {
    this.abort.abort()
  }

  private need<T>(v: T | undefined): T {
    if (!v) throw new Error(this.kind ? `Operation not available for ${this.kind}` : 'Document not opened')
    return v
  }
}
