/**
 * PPTX collaboration (main thread). Text and shape edits are applied optimistically (last writer
 * wins per shape); slide list changes wait for the hub's order. Undo / redo restore the slide
 * touched by the local user's own operation.
 */
import type { PresentationMeta, ShapeOp, SlideOp } from '../core/pptx/types'
import type { DocumentClient } from '../core/worker/client'
import type { Methods } from '../core/worker/protocol'
import { getHost, onLocal } from './hosts'
import { conflictsP, transformP, type POp, type PSlideOp } from './pptx-ops'
import type { PptxCollabResult } from './pptx-worker'
import type { Adapter, ApplyResult } from './session'
import { applyComment, revertComment, type CommentOp } from './comments'

const CAPTURED = new Set<keyof Methods>(['pptx.setText', 'pptx.shapes', 'pptx.slides', 'pptx.image', 'pptx.undo'])

export interface PptxPresence {
  k: 'pptx'
  /** slide part path */
  slide: string
  /** selected element id */
  top?: number
}

/** what PptxView publishes about the local selection */
export interface PptxLocalSel {
  slide: number
  top?: number
}

/** PptxView registers this so remote slide-list changes can close its text editing first */
export interface PptxHost {
  endEditing(): void
}

export class PptxCollab implements Adapter<POp> {
  submit?: (op: POp) => Promise<unknown>
  nextId?: () => string
  onPresence?: (p: PptxPresence) => void
  private undoStack: string[] = []
  private redoStack: string[] = []
  /** this client's image relationship id → image key */
  private ridKey = new Map<string, string>()
  private offLocal: () => void

  constructor(private client: DocumentClient) {
    this.offLocal = onLocal(client, (d) => {
      const s = d as PptxLocalSel
      const slide = this.path(s.slide)
      if (slide) this.onPresence?.({ k: 'pptx', slide, top: s.top })
    })
  }

  dispose() {
    this.offLocal()
  }

  private get meta() {
    return this.client.info.pptx as PresentationMeta
  }

  private path(i: number): string {
    return this.meta.slides[i]?.path ?? ''
  }

  private id() {
    return this.nextId!()
  }

  captures(method: keyof Methods): boolean {
    return method === 'comments.set' || CAPTURED.has(method)
  }

  intercept(method: keyof Methods, args: unknown[]): Promise<unknown> | undefined {
    if (!this.submit) return undefined
    const a = args as never[]
    switch (method) {
      case 'comments.set':
        return this.submit({ t: 'comment', id: this.id(), cid: a[0], c: a[1] as CommentOp['c'] })
      case 'pptx.setText':
        return this.local({ t: 'text', id: this.id(), slide: this.path(a[0]), sid: a[1], paras: a[2] }).then((r) => (r as PptxCollabResult).model)
      case 'pptx.shapes': {
        const ops = (a[1] as ShapeOp[]).map((o) => (o.op === 'insert' && o.rid && this.ridKey.has(o.rid) ? { ...o, rid: undefined, key: this.ridKey.get(o.rid) } : o))
        return this.local({ t: 'shapes', id: this.id(), slide: this.path(a[0]), ops }).then((r) => (r as PptxCollabResult).model)
      }
      case 'pptx.image': {
        const key = `${this.id()}`
        const bytes = new Uint8Array((a[1] as ArrayBuffer).slice(0))
        return this.submit({ t: 'image', id: key, key, slide: this.path(a[0]), bytes, ext: a[2] }).then((r) => {
          const rid = (r as PptxCollabResult).rid
          if (rid) this.ridKey.set(rid, key)
          return rid
        })
      }
      case 'pptx.slides':
        return this.submit({ t: 'slides', id: this.id(), op: this.pslide(a[0]) }).then((r) => (r as PptxCollabResult | null)?.meta ?? this.meta)
      case 'pptx.undo':
        return this.history(a[0] ? this.redoStack : this.undoStack, a[0] ? this.undoStack : this.redoStack, a[0] ? 'next' : 'prev')
    }
    return undefined
  }

  private pslide(o: SlideOp): PSlideOp {
    switch (o.op) {
      case 'add':
        return { op: 'add', after: o.after >= 0 ? this.path(o.after) || null : null, layout: o.layout }
      case 'duplicate':
        return { op: 'duplicate', path: this.path(o.index) }
      case 'delete':
        return { op: 'delete', path: this.path(o.index) }
      case 'move':
        return { op: 'move', path: this.path(o.from), to: o.to }
      case 'hide':
        return { op: 'hide', path: this.path(o.index), hidden: o.hidden }
    }
  }

  private local(op: POp): Promise<unknown> {
    this.undoStack.push(op.id)
    if (this.undoStack.length > 200) this.undoStack.shift()
    this.redoStack = []
    return this.submit!(op)
  }

  private async history(from: string[], to: string[], which: 'prev' | 'next') {
    const target = from.pop()
    if (!target) return null
    to.push(target)
    const r = (await this.submit!({ t: 'revert', id: this.id(), target, which })) as PptxCollabResult
    return r.model && r.slide !== undefined ? { index: r.slide, model: r.model } : null
  }

  // ---------------------------------------------------------------- adapter

  transform(op: POp, prior: POp): POp | null {
    return transformP(op, prior)
  }

  conflicts(pending: POp[], remote: POp[]): boolean {
    return conflictsP(pending, remote)
  }

  pessimistic(op: POp): boolean {
    return op.t === 'slides'
  }

  prepare(ops: POp[]) {
    // slide indexes are about to move: let the text editor commit with the current ones first
    if (ops.some((o) => o.t === 'slides')) getHost<PptxHost>(this.client, 'pptx')?.endEditing()
  }

  async apply(op: POp, kind: 'local' | 'replay' | 'remote'): Promise<ApplyResult> {
    if (op.t === 'comment') return applyComment(this.client, op)
    const { res, rec } = await this.client.direct('collab.pptx.apply', op)
    if (op.t === 'slides') {
      this.undoStack = []
      this.redoStack = []
      if (res.meta) Object.assign(this.meta, res.meta)
    }
    if (kind !== 'local') this.client.notify('collab:pptx', res)
    return { result: res, undo: rec }
  }

  async revert(undo: unknown) {
    const cm = revertComment(this.client, undo)
    if (cm) return cm
    if (typeof undo !== 'number') return
    this.client.notify('collab:pptx', await this.client.direct('collab.pptx.revert', undo))
  }
}
