import { DocxDocument } from '../docx/parser'
import { OpcPackage, type DocKind } from '../opc'
import { PptxPresentation } from '../pptx/parser'
import { openSource, type DocSource } from '../source'
import { XlsxWorkbook } from '../xlsx/workbook'
import { ZipArchive } from '../zip'
import type { Events, Methods, OpenParams, OpenResult } from './protocol'

type Emit = <K extends keyof Events>(event: K, data: Events[K]) => void

/** Format-agnostic engine; runs inside the worker (or directly, e.g. in tests). */
export class Engine {
  private pkg?: OpcPackage
  private kind?: DocKind
  private docx?: DocxDocument
  private xlsx?: XlsxWorkbook
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

  async call<M extends keyof Methods>(method: M, args: unknown[]): Promise<unknown> {
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
      case 'xlsx.load': {
        const wb = this.need(this.xlsx)
        void wb.loadSheet(
          args[0] as number,
          (p) => this.emit('xlsx:progress', p),
          (phase, loaded, total) => this.emit('progress', { phase, loaded, total }),
        )
        return
      }
      case 'xlsx.block':
        return this.need(this.xlsx).block(args[0] as number, args[1] as number, args[2] as number)
      case 'xlsx.cell':
        return this.need(this.xlsx).cell(args[0] as number, args[1] as number, args[2] as number)
      case 'pptx.slide':
        return this.need(this.pptx).slide(args[0] as number)
    }
    throw new Error(`Unknown method ${String(method)}`)
  }

  dispose() {
    this.abort.abort()
  }

  private need<T>(v: T | undefined): T {
    if (!v) throw new Error(this.kind ? `Operation not available for ${this.kind}` : 'Document not opened')
    return v
  }
}
