import type { DocKind } from '../opc'
import type { DocxMeta, Block } from '../docx/types'
import type { BlockOut, CellDetail, SheetProgress, WorkbookMeta } from '../xlsx/types'
import type { PresentationMeta, SlideModel } from '../pptx/types'

export interface OpenParams {
  name?: string
  headers?: Record<string, string>
  credentials?: RequestCredentials
}

export interface OpenResult {
  kind: DocKind
  name?: string
  size: number
  xlsx?: WorkbookMeta
  pptx?: PresentationMeta
}

export interface Methods {
  open(src: Blob | ArrayBuffer | string, params: OpenParams): OpenResult
  part(path: string): Blob
  'docx.start'(): void
  'xlsx.load'(sheet: number): void
  'xlsx.block'(sheet: number, r0: number, r1: number): BlockOut
  'xlsx.cell'(sheet: number, r: number, c: number): CellDetail | undefined
  'pptx.slide'(index: number): SlideModel
}

export interface Events {
  progress: { phase: string; loaded: number; total: number }
  'docx:blocks': Block[]
  'docx:meta': DocxMeta
  'docx:done': { error?: string }
  'xlsx:progress': SheetProgress
}

export type Req = { id: number; method: keyof Methods; args: unknown[] }
export type Res = { id: number; ok: true; result: unknown } | { id: number; ok: false; error: string }
export type Evt = { event: keyof Events; data: unknown }
