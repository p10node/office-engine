/**
 * XLSX collaboration (main thread): turns the view's worker calls into shared operations.
 * Cell edits, formatting, layout, fill / sort / replace are applied optimistically; inserting /
 * deleting rows or columns and sheet-list changes wait for the hub's order. Undo / redo only
 * touch the local user's own operations (shared as `revert` operations).
 */
import type { DocumentClient } from '../core/worker/client'
import type { Methods } from '../core/worker/protocol'
import type { BookOp, Range, SheetOp, WorkbookMeta } from '../core/xlsx/types'
import { onLocal } from './hosts'
import type { Adapter, ApplyResult } from './session'
import { conflictsX, isStructural, transformX, type PBookOp, type XOp } from './xlsx-ops'

const CAPTURED = new Set<keyof Methods>(['xlsx.set', 'xlsx.format', 'xlsx.sheetOp', 'xlsx.replace', 'xlsx.bookOp', 'xlsx.undo', 'xlsx.redo'])

type UndoEntry = { kind: 'op'; id: string } | { kind: 'inverse'; sheet: string; undo: SheetOp; redo: SheetOp }

export interface XlsxPresence {
  k: 'xlsx'
  /** sheet part path */
  sheet: string
  /** selected range and active cell */
  range: Range
  ar: number
  ac: number
}

/** what XlsxView publishes about the local selection */
export interface XlsxLocalSel {
  sheet: number
  range: Range
  ar: number
  ac: number
}

export class XlsxCollab implements Adapter<XOp> {
  submit?: (op: XOp) => Promise<unknown>
  nextId?: () => string
  onPresence?: (p: XlsxPresence) => void
  private undoStack: UndoEntry[] = []
  private redoStack: UndoEntry[] = []
  private offLocal: () => void

  constructor(private client: DocumentClient) {
    this.offLocal = onLocal(client, (d) => {
      const s = d as XlsxLocalSel
      const sheet = this.path(s.sheet)
      if (sheet) this.onPresence?.({ k: 'xlsx', sheet, range: s.range, ar: s.ar, ac: s.ac })
    })
  }

  dispose() {
    this.offLocal()
  }

  private get meta() {
    return this.client.info.xlsx as WorkbookMeta
  }

  private path(i: number): string {
    return this.meta.sheets[i]?.path ?? ''
  }

  private id() {
    return this.nextId!()
  }

  captures(method: keyof Methods): boolean {
    return CAPTURED.has(method)
  }

  intercept(method: keyof Methods, args: unknown[]): Promise<unknown> | undefined {
    if (!this.submit) return undefined
    const a = args as never[]
    switch (method) {
      case 'xlsx.set':
        return this.local({ t: 'set', id: this.id(), sheet: this.path(a[0]), cells: a[1] })
      case 'xlsx.format':
        return this.local({ t: 'fmt', id: this.id(), sheet: this.path(a[0]), range: a[1], patch: a[2] })
      case 'xlsx.sheetOp':
        return this.local({ t: 'sheet', id: this.id(), sheet: this.path(a[0]), op: a[1] })
      case 'xlsx.replace':
        return this.local({ t: 'replace', id: this.id(), sheet: this.path(a[0]), query: a[1], by: a[2], opts: a[3], at: a[4] })
      case 'xlsx.bookOp':
        return this.local({ t: 'book', id: this.id(), op: this.pbook(a[0]) })
      case 'xlsx.undo':
        return this.history(this.undoStack, this.redoStack, 'prev')
      case 'xlsx.redo':
        return this.history(this.redoStack, this.undoStack, 'next')
    }
    return undefined
  }

  private pbook(b: BookOp): PBookOp {
    switch (b.op) {
      case 'add':
        return { op: 'add', name: b.name, after: b.after !== undefined ? this.path(b.after) : undefined }
      case 'rename':
        return { op: 'rename', path: this.path(b.index), name: b.name }
      case 'delete':
        return { op: 'delete', path: this.path(b.index) }
      case 'move':
        return { op: 'move', path: this.path(b.from), to: b.to }
    }
  }

  private local(op: XOp): Promise<unknown> {
    if (op.t === 'sheet' && (op.op.op === 'insertRows' || op.op.op === 'insertCols')) {
      const del: SheetOp = { op: op.op.op === 'insertRows' ? 'deleteRows' : 'deleteCols', at: op.op.at, n: op.op.n }
      this.undoStack.push({ kind: 'inverse', sheet: op.sheet, undo: del, redo: op.op })
    } else if (!isStructural(op)) this.undoStack.push({ kind: 'op', id: op.id })
    if (this.undoStack.length > 200) this.undoStack.shift()
    this.redoStack = []
    return this.submit!(op)
  }

  private history(from: UndoEntry[], to: UndoEntry[], which: 'prev' | 'next'): Promise<unknown> {
    const e = from.pop()
    if (!e) return Promise.resolve(null)
    to.push(e)
    if (e.kind === 'op') return this.submit!({ t: 'revert', id: this.id(), target: e.id, which })
    return this.submit!({ t: 'sheet', id: this.id(), sheet: e.sheet, op: which === 'prev' ? e.undo : e.redo })
  }

  // ---------------------------------------------------------------- adapter

  transform(op: XOp, prior: XOp): XOp | null {
    return transformX(op, prior)
  }

  conflicts(pending: XOp[], remote: XOp[]): boolean {
    return conflictsX(pending, remote)
  }

  pessimistic(op: XOp): boolean {
    return isStructural(op)
  }

  async apply(op: XOp, kind: 'local' | 'replay' | 'remote'): Promise<ApplyResult> {
    const { res, rec } = await this.client.direct('collab.xlsx.apply', op)
    if (isStructural(op)) {
      // the worker dropped its records: older entries can't be reverted any more
      const keep = (e: UndoEntry) => e.kind === 'inverse' && kind === 'local'
      this.undoStack = this.undoStack.filter(keep)
      this.redoStack = this.redoStack.filter(keep)
    }
    if (kind !== 'local' && res) this.client.notify('collab:xlsx', res)
    return { result: res, undo: rec }
  }

  async revert(undo: unknown) {
    if (typeof undo !== 'number') return
    const res = await this.client.direct('collab.xlsx.revert', undo)
    if (res) this.client.notify('collab:xlsx', res)
  }
}

