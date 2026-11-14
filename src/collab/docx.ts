/**
 * DOCX collaboration: operations on the main-thread paragraph model (DocxEditor),
 * which mirrors every change to the worker itself.
 *
 *  - `text`   splice inside one paragraph (typing, deleting, inline formatting) — merges
 *             character-wise with concurrent edits of the same paragraph
 *  - `paras`  contiguous paragraphs replaced (Enter, join, paste, paragraph formatting…) — last writer wins
 *  - `blocks` top-level blocks replaced (tables)
 *  - `image` / `link`  package resources, shared by key (relationship ids differ per client)
 *  - `call`   other document-level worker calls (adding a built-in style, giving a style another look), run in hub order
 *  - `comment` a comment set / deleted (its range is marked in the text, shared with the text)
 */
import type { Block, HFSetResult, Inline, PageSetResult, Paragraph, StyleInfo } from '../core/docx/types'
import type { DocumentClient } from '../core/worker/client'
import type { Methods } from '../core/worker/protocol'
import type { DocState, DocxCommit, DocxEditor } from '../react/docx/DocxEditor'
import { domFromPos } from '../react/edit/input'
import { plen, runsText, splitRuns, type Pos, type RRun, type Sel } from '../react/edit/rich'
import { watchHost } from './hosts'
import type { Adapter, ApplyResult } from './session'
import { applyComment, revertComment, type CommentOp } from './comments'
import { dict } from '../core/dict'

export interface DocxHost {
  editor: DocxEditor
  st: DocState
}

/** top-level block identity: paragraph id, or table (regenerated id / original body item) */
export type BKey = { p: number } | { tid?: number; item?: number }

export type DocxOp =
  | { t: 'text'; pid: number; at: number; del: number; ins: Inline[] }
  | { t: 'paras'; before: number[]; next: Paragraph[] }
  | { t: 'blocks'; before: BKey[]; prev: BKey | null; next: Block[] }
  | { t: 'image'; key: string; bytes: Uint8Array; ext: string }
  | { t: 'link'; key: string; url: string }
  | { t: 'call'; method: string; args: unknown[] }
  | CommentOp

/** a point in the document that survives structural edits: paragraph id, page-break part, offset */
export interface DPos {
  pid: number
  k: number
  o: number
}

export interface DocxPresence {
  k: 'docx'
  a: DPos
  f: DPos
}

/** where a paragraph removed on its own stood: its index in the body (tables around it), or in a cell of table `tbl` */
interface Home {
  index: number
  cell?: { tbl: BKey; ri: number; ci: number }
  prevT?: BKey
  nextT?: BKey
}

/**
 * How to undo one model change; `paras` with an empty `next` re-inserts `before` next to a paragraph it had as a
 * neighbour, else at its `home`.
 */
type Step =
  | { t: 'paras'; before: Paragraph[]; next: Paragraph[]; prev?: { pid: number; k: number }; after?: { pid: number; k: number }; home?: Home }
  | { t: 'blocks'; before: Block[]; next: Block[]; prev: BKey | null }
/** the steps of an op, undone in reverse order */
type Undo = Step[] | null

/** worker methods that only read */
const READS = new Set<keyof Methods>(['open', 'part', 'docx.start', 'docx.convert', 'save', 'dirty'])

// ---------------------------------------------------------------- ops as received
//
// An op arrives from the hub as another client wrote it, and every client in the room runs it. What it asks the
// worker to do is therefore checked here, by the receiver: a `call` op names one of the document-level edits
// `intercept` shares and nothing else. `open` with an address to fetch, `save`, a method of another format or
// arguments of another shape are refused, and the op is dropped (by every client alike, so they still agree).

const isObj = (x: unknown): x is Record<string, unknown> => typeof x === 'object' && x !== null && !Array.isArray(x)
/** a plain object holding only `keys` */
const only = (x: unknown, keys: readonly string[]): x is Record<string, unknown> => isObj(x) && Object.keys(x).every((k) => keys.includes(k))
const text = (x: unknown, max: number): x is string => typeof x === 'string' && x.length <= max
const scalar = (x: unknown) => x === null || typeof x === 'boolean' || (typeof x === 'number' && Number.isFinite(x)) || text(x, 255)

const CANON_STYLES: readonly unknown[] = ['normal', 'title', 'subtitle', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6']
const LOOK_KEYS = ['b', 'i', 'u', 'strike', 'color', 'sz', 'font', 'align', 'line', 'lineRule', 'before', 'after']
const PAGE_KEYS = ['w', 'h', 'mt', 'mr', 'mb', 'ml', 'cols', 'colGap']
const HF_KINDS = ['header', 'footer', 'firstHeader', 'firstFooter']
/** a header / footer part's XML, as an undo carries it */
const MAX_HF_XML = 4 * 1024 * 1024

const hfLine = (l: unknown) =>
  only(l, ['text', 'align', 'b', 'sz']) && text(l.text, 2000) && (l.align === undefined || text(l.align, 10)) && (l.b === undefined || typeof l.b === 'boolean') && (l.sz === undefined || typeof l.sz === 'number')

function hfSet(x: unknown): boolean {
  if (!only(x, ['parts', 'raw', 'titlePg'])) return false
  if (x.titlePg !== undefined && typeof x.titlePg !== 'boolean') return false
  if (x.parts !== undefined && !(only(x.parts, HF_KINDS) && Object.values(x.parts).every((lines) => Array.isArray(lines) && lines.length <= 20 && lines.every(hfLine)))) return false
  return x.raw === undefined || (only(x.raw, HF_KINDS) && Object.values(x.raw).every((xml) => text(xml, MAX_HF_XML)))
}

/** the worker calls a `call` op may name, each with the check of its arguments (the worker checks their values) */
const SHARED_CALLS: Record<string, (args: unknown[]) => boolean> = dict({
  'docx.style': (a) => a.length === 1 && CANON_STYLES.includes(a[0]),
  'docx.styleSet': (a) => a.length === 2 && text(a[0], 255) && only(a[1], LOOK_KEYS) && Object.values(a[1]).every(scalar),
  'docx.page': (a) => a.length === 1 && only(a[0], PAGE_KEYS) && Object.values(a[0]).every((v) => typeof v === 'number' && Number.isFinite(v)),
  'docx.hf': (a) => a.length === 1 && hfSet(a[0]),
})

/** whether a `call` op is one the document-level edits make: the method is shared, the arguments are its kind */
export function isSharedCall(method: unknown, args: unknown): method is keyof Methods {
  const check = typeof method === 'string' ? SHARED_CALLS[method] : undefined
  return !!check && Array.isArray(args) && check(args)
}

const IMAGE_EXT = /^[a-z0-9]{1,8}$/i

/** throws for a resource / call op that is not of the shape `fetchRes` hands to the worker */
function checkOp(op: DocxOp): void {
  const bad = () => new Error(`Collaboration: operation refused (${String((op as { method?: unknown }).method ?? op.t).slice(0, 40)})`)
  if (op.t === 'call') {
    if (!isSharedCall(op.method, op.args)) throw bad()
  } else if (op.t === 'image') {
    if (!text(op.key, 200) || !(op.bytes instanceof Uint8Array) || typeof op.ext !== 'string' || !IMAGE_EXT.test(op.ext)) throw bad()
  } else if (op.t === 'link') {
    if (!text(op.key, 200) || !text(op.url, 8192)) throw bad()
  }
}

const bkey = (b: Block): BKey | null => (b.t === 'p' ? (b.pid !== undefined ? { p: b.pid } : null) : b.t === 'tbl' ? { tid: b.tid, item: b.item } : null)

function matches(b: Block | undefined, k: BKey): boolean {
  if (!b) return false
  if ('p' in k) return b.t === 'p' && b.pid === k.p
  return b.t === 'tbl' && ((k.tid !== undefined && b.tid === k.tid) || (k.item !== undefined && b.item === k.item))
}

const uniq = (xs: number[]) => [...new Set(xs)]

/** paragraph ids inside blocks (table cells included) */
function pidsOf(blocks: Block[], out: number[] = []): number[] {
  for (const b of blocks) {
    if (b.t === 'p' && b.pid !== undefined) out.push(b.pid)
    else if (b.t === 'tbl') for (const r of b.rows) for (const c of r.cells) pidsOf(c.blocks, out)
  }
  return out
}

// ---------------------------------------------------------------- run diff / splice

function atomKey(r: Inline): string {
  if (r.t !== 'text') return JSON.stringify(r)
  const { text: _t, ...rest } = r
  return JSON.stringify(rest)
}

/** single replaced range between two run lists (common prefix / suffix, formatting included) */
export function diffRuns(a: Inline[], b: Inline[]): { at: number; del: number; ins: Inline[] } | null {
  // one entry per UTF-16 unit (what offsets count), formatting included
  const units = (runs: Inline[]) => {
    const out: string[] = []
    for (const r of runs) {
      const k = atomKey(r)
      if (r.t === 'text') for (let i = 0; i < r.text.length; i++) out.push(k + '\u0000' + r.text[i])
      else out.push(k)
    }
    return out
  }
  const x = units(a)
  const y = units(b)
  let pre = 0
  while (pre < x.length && pre < y.length && x[pre] === y[pre]) pre++
  let suf = 0
  while (suf < x.length - pre && suf < y.length - pre && x[x.length - 1 - suf] === y[y.length - 1 - suf]) suf++
  if (pre === x.length && pre === y.length) return null
  const [, rest] = splitRuns(b as RRun[], pre)
  const [mid] = splitRuns(rest, y.length - pre - suf)
  return { at: pre, del: x.length - pre - suf, ins: mid as Inline[] }
}

export function spliceRuns(runs: Inline[], at: number, del: number, ins: Inline[]): Inline[] {
  const n = plen(runs as RRun[])
  const a = Math.max(0, Math.min(at, n))
  const d = Math.max(0, Math.min(del, n - a))
  const [left, rest] = splitRuns(runs as RRun[], a)
  const [, right] = splitRuns(rest, d)
  return [...(left as Inline[]), ...ins, ...(right as Inline[])]
}

/** `x` (a position in the text before `prior` replaced [at, at+del) by `ins` chars) after it */
function mapPos(x: number, at: number, del: number, ins: number, end: boolean): number {
  if (x < at || (x === at && (del > 0 || end))) return x
  if (x >= at + del) return x - del + ins
  return at + ins
}

/** maps an offset through an edit of the text, using the common prefix / suffix */
function mapByDiff(o: number, before: string, after: string): number {
  if (before === after) return o
  let pre = 0
  while (pre < before.length && pre < after.length && before[pre] === after[pre]) pre++
  let suf = 0
  while (suf < before.length - pre && suf < after.length - pre && before[before.length - 1 - suf] === after[after.length - 1 - suf]) suf++
  if (o <= pre) return o
  if (o >= before.length - suf) return after.length - (before.length - o)
  return after.length - suf
}

const sameProps = (a: Paragraph, b: Paragraph) =>
  a.pid === b.pid &&
  a.from === b.from &&
  a.sid === b.sid &&
  !!a.pb === !!b.pb &&
  a.keepSect === b.keepSect &&
  JSON.stringify(a.pd ?? null) === JSON.stringify(b.pd ?? null) &&
  JSON.stringify(a.ps) === JSON.stringify(b.ps) &&
  JSON.stringify(a.orig ?? null) === JSON.stringify(b.orig ?? null)

// ---------------------------------------------------------------- adapter

export class DocxCollab implements Adapter<DocxOp> {
  private host?: DocxHost
  /** > 0 while this adapter drives the editor (its commits are not local ops) */
  private suppress = 0
  /** resource key → this client's value ({ rid, path } for images, link key for links) */
  private res = new Map<string, { rid: string; path: string } | string>()
  /** this client's rid / link key → resource key */
  private resKey = new Map<string, string>()
  private results = new WeakMap<DocxOp, unknown>()
  private snap: { a: DPos; f: DPos; text: Map<string, string>; focused: boolean } | null = null
  /** the model was changed since the last `settled` */
  private touchedModel = false
  private site = 0
  private counter = 0
  private offHost?: () => void
  /** sends presence (set by the session owner) */
  onPresence?: (p: DocxPresence) => void
  private lastPresence = ''
  private raf = 0

  /** registers a local op the editor already applied */
  push?: (op: DocxOp, undo: Undo) => void
  /** shares a local op through the session (applied by `apply`) */
  submit?: (op: DocxOp) => Promise<unknown>

  constructor(private client: DocumentClient) {}

  // ---------------------------------------------------------------- wiring

  ready(): Promise<void> {
    return new Promise((resolve) => {
      this.offHost = watchHost<DocxHost>(this.client, 'docx', (h) => {
        if (this.host?.editor === h.editor) return
        if (this.host) this.host.editor.onCommit = undefined
        this.host = h
        h.editor.onCommit = (c) => this.onCommit(c)
        this.allocPids()
        resolve()
      })
      if (typeof document !== 'undefined') document.addEventListener('selectionchange', this.onSel)
    })
  }

  /** gives this client its own range of paragraph / table ids */
  setSite(site: number) {
    this.site = site
    this.allocPids()
  }

  private allocPids() {
    const ed = this.host?.editor
    if (!ed || !this.site) return
    const base = 2_000_000_000 + this.site * 10_000_000
    if (ed.nextPid < base || ed.nextPid >= base + 10_000_000) ed.nextPid = base
  }

  dispose() {
    this.offHost?.()
    if (this.host) this.host.editor.onCommit = undefined
    if (typeof document !== 'undefined') document.removeEventListener('selectionchange', this.onSel)
    if (this.raf) cancelAnimationFrame(this.raf)
  }

  captures(method: keyof Methods): boolean {
    if (method === 'comments.set') return true
    // paragraph / table writes mirror model changes (captured by DocxEditor.onCommit)
    return !this.suppress && !READS.has(method) && method !== 'docx.edit' && method !== 'docx.table' && method.startsWith('docx.')
  }

  /** client.call interceptor: package resources and document-level calls become ops */
  intercept(method: keyof Methods, args: unknown[]): Promise<unknown> | undefined {
    if (!this.submit || !this.captures(method)) return undefined
    switch (method) {
      case 'comments.set':
        return this.submit({ t: 'comment', cid: args[0] as string, c: args[1] as CommentOp['c'] })
      case 'docx.image': {
        const bytes = new Uint8Array((args[0] as ArrayBuffer).slice(0))
        return this.submit({ t: 'image', key: this.key(), bytes, ext: args[1] as string })
      }
      case 'docx.link':
        return this.submit({ t: 'link', key: this.key(), url: args[0] as string })
    }
    return this.submit({ t: 'call', method, args })
  }

  private key() {
    return `${this.site}:${++this.counter}:${Math.random().toString(36).slice(2, 7)}`
  }

  // ---------------------------------------------------------------- model helpers

  private get st() {
    return this.host!.st
  }

  private get ed() {
    return this.host!.editor
  }

  /** every block of paragraph `pid` (page-break parts) with its address */
  private locate(pid: number): { list: Block[]; index: number; pa: string; p: Paragraph }[] {
    const out: { list: Block[]; index: number; pa: string; p: Paragraph }[] = []
    this.st.blocks.forEach((b, bi) => {
      if (b.t === 'p') {
        if (b.pid === pid) out.push({ list: this.st.blocks, index: bi, pa: String(bi), p: b })
      } else if (b.t === 'tbl')
        b.rows.forEach((r, ri) =>
          r.cells.forEach((cell, ci) =>
            cell.blocks.forEach((x, pi) => {
              if (x.t === 'p' && x.pid === pid) out.push({ list: cell.blocks, index: pi, pa: `${bi}.${ri}.${ci}.${pi}`, p: x })
            }),
          ),
        )
    })
    return out
  }

  private topIndex(k: BKey): number {
    return this.st.blocks.findIndex((b) => matches(b, k))
  }

  /** contiguous paragraph addresses from the first block of pids[0], following `pids` */
  private region(pids: number[]): string[] | null {
    const want = new Set(pids)
    for (const pid of pids) {
      const loc = this.locate(pid)[0]
      if (!loc) continue
      const pas: string[] = []
      for (let i = loc.index; i < loc.list.length; i++) {
        const b = loc.list[i]
        if (b.t !== 'p' || b.pid === undefined || !want.has(b.pid)) break
        pas.push(loc.pa.replace(/\d+$/, String(i)))
      }
      return pas
    }
    return null
  }

  // ---------------------------------------------------------------- resources (ids differ per client)

  private remap<T extends Block | Inline>(x: T): T {
    if (x.t === 'img') {
      const ck = (x as { ck?: string }).ck
      const r = ck ? this.res.get(ck) : undefined
      return r && typeof r === 'object' ? { ...x, rid: r.rid, target: r.path } : x
    }
    if (x.t === 'text') {
      const ck = (x as { lck?: string }).lck
      const r = ck ? this.res.get(ck) : undefined
      return typeof r === 'string' ? { ...x, lk: r } : x
    }
    if (x.t === 'p') return { ...x, runs: x.runs.map((r) => this.remap(r)) }
    if (x.t === 'tbl') return { ...x, rows: x.rows.map((r) => ({ ...r, cells: r.cells.map((c) => ({ ...c, blocks: c.blocks.map((b) => this.remap(b)) })) })) }
    return x
  }

  /** tags runs referring to resources this client created, so others can map them */
  private tag<T extends Block | Inline>(x: T): T {
    if (x.t === 'img' && x.rid && this.resKey.has(x.rid) && !(x as { ck?: string }).ck) return { ...x, ck: this.resKey.get(x.rid) }
    if (x.t === 'text' && x.lk && this.resKey.has(x.lk) && !(x as { lck?: string }).lck) return { ...x, lck: this.resKey.get(x.lk) }
    if (x.t === 'p') return { ...x, runs: x.runs.map((r) => this.tag(r)), caret: undefined }
    if (x.t === 'tbl') return { ...x, rows: x.rows.map((r) => ({ ...r, cells: r.cells.map((c) => ({ ...c, blocks: c.blocks.map((b) => this.tag(b)) })) })) }
    return x
  }

  private async fetchRes(op: DocxOp) {
    checkOp(op)
    if (op.t === 'image' && !this.res.has(op.key)) {
      const r = await this.client.direct('docx.image', op.bytes.slice().buffer as ArrayBuffer, op.ext)
      this.res.set(op.key, r)
      this.resKey.set(r.rid, op.key)
    } else if (op.t === 'link' && !this.res.has(op.key)) {
      const lk = await this.client.direct('docx.link', op.url)
      this.res.set(op.key, lk)
      this.resKey.set(lk, op.key)
    } else if (op.t === 'call' && !this.results.has(op)) {
      // bound: `direct` reads the client's state
      const direct = this.client.direct.bind(this.client) as (m: string, ...a: unknown[]) => Promise<unknown>
      this.results.set(op, await direct(op.method, ...op.args))
    }
  }

  // ---------------------------------------------------------------- local capture

  private onCommit(c: DocxCommit) {
    if (this.suppress || !this.push) return
    if (c.kind === 'paras') {
      const b = c.before
      const n = c.next
      if (b.length === 1 && n.length === 1 && b[0].pid !== undefined && sameProps(b[0], n[0]) && this.locate(b[0].pid).length === 1) {
        const d = diffRuns(b[0].runs, n[0].runs)
        if (!d) return
        this.push({ t: 'text', pid: b[0].pid, at: d.at, del: d.del, ins: d.ins.map((r) => this.tag(r)) }, [{ t: 'paras', before: b, next: n }])
        return
      }
      this.push({ t: 'paras', before: b.map((p) => p.pid!), next: n.map((p) => this.tag(p)) }, [{ t: 'paras', before: b, next: n }])
      return
    }
    const prev = c.bi > 0 ? bkey(this.st.blocks[c.bi - 1]) : null
    this.push(
      { t: 'blocks', before: c.before.map(bkey).filter((k): k is BKey => !!k), prev, next: c.next.map((b) => this.tag(b)) },
      [{ t: 'blocks', before: c.before, next: c.next, prev }],
    )
  }

  // ---------------------------------------------------------------- adapter

  transform(op: DocxOp, prior: DocxOp): DocxOp | null {
    if (op.t !== 'text') return op
    if (prior.t === 'text' && prior.pid === op.pid) {
      const pins = plen(prior.ins as RRun[])
      const at = mapPos(op.at, prior.at, prior.del, pins, false)
      const end = op.del ? Math.max(at, mapPos(op.at + op.del, prior.at, prior.del, pins, true)) : at
      return at === op.at && end - at === op.del ? op : { ...op, at, del: end - at }
    }
    if (prior.t === 'paras' && (prior.before.includes(op.pid) || prior.next.some((p) => p.pid === op.pid))) {
      // the paragraph was rewritten: keep the splice (clamped) while it still exists
      return prior.next.some((p) => p.pid === op.pid) ? op : null
    }
    if (prior.t === 'blocks' && prior.before.some((k) => 'p' in k && k.p === op.pid) && !pidsOf(prior.next).includes(op.pid)) return null
    return op
  }

  private touched(op: DocxOp): string[] {
    switch (op.t) {
      case 'text':
        return ['p' + op.pid]
      case 'paras':
        return [...op.before, ...op.next.map((p) => p.pid!)].map((p) => 'p' + p)
      case 'comment':
        return ['c' + op.cid]
      case 'blocks': {
        const keys = [...op.before, ...(op.prev ? [op.prev] : []), ...op.next.map(bkey).filter((k): k is BKey => !!k)]
        const out = keys.map((k) => ('p' in k ? 'p' + k.p : 't' + (k.item ?? 'n' + k.tid)))
        for (const pid of pidsOf(op.next)) out.push('p' + pid)
        return out
      }
      default:
        return []
    }
  }

  conflicts(pending: DocxOp[], remote: DocxOp[]): boolean {
    const mine = new Set(pending.flatMap((o) => this.touched(o)))
    return remote.some((o) => this.touched(o).some((k) => mine.has(k)))
  }

  pessimistic(op: DocxOp): boolean {
    return op.t === 'call'
  }

  async prepare(ops: DocxOp[]) {
    // never rebuild paragraphs under an active IME composition
    while (this.host?.editor.input.isComposing) await new Promise((r) => setTimeout(r, 50))
    for (const op of ops) await this.fetchRes(op)
    this.capture()
  }

  async apply(op: DocxOp, kind: 'local' | 'replay' | 'remote'): Promise<ApplyResult> {
    switch (op.t) {
      case 'comment':
        return applyComment(this.client, op)
      case 'image':
      case 'link':
        await this.fetchRes(op)
        return { result: this.res.get(op.key) }
      case 'call': {
        await this.fetchRes(op)
        const result = this.results.get(op)
        // a style added to the document, or one given another look: the editor shows it (once, whoever asked)
        if ((op.method === 'docx.style' || op.method === 'docx.styleSet') && result && typeof result === 'object' && this.host) this.ed.styled(result as { styles: StyleInfo[] })
        // the page setup changed: every editor lays the pages out again
        if (op.method === 'docx.page' && result && typeof result === 'object' && this.host) this.ed.paged(result as PageSetResult)
        if (op.method === 'docx.hf' && result && typeof result === 'object' && this.host) this.ed.hfed(result as HFSetResult)
        return { result }
      }
    }
    if (!this.host) throw new Error('Document editor not ready')
    if (kind === 'remote' || kind === 'replay') this.capture()
    this.touchedModel = true
    this.suppress++
    try {
      return { undo: this.run(op) }
    } finally {
      this.suppress--
    }
  }

  /** where the paragraph at `loc` stands, for a revert that finds no paragraph next to it */
  private home(loc: { list: Block[]; index: number; pa: string }): Home {
    const tk = (i: number) => {
      const b = loc.list[i]
      return b?.t === 'tbl' ? bkey(b) ?? undefined : undefined
    }
    if (loc.list === this.st.blocks) return { index: loc.index, prevT: tk(loc.index - 1), nextT: tk(loc.index + 1) }
    const [bi, ri, ci] = loc.pa.split('.').map(Number)
    const tbl = bkey(this.st.blocks[bi])
    return tbl ? { index: loc.index, cell: { tbl, ri, ci } } : { index: loc.index }
  }

  /** the address a paragraph that stood at `h` gets now */
  private homePa(h: Home): string | undefined {
    if (h.cell) {
      const ti = this.topIndex(h.cell.tbl)
      const t = this.st.blocks[ti]
      const cell = t?.t === 'tbl' ? t.rows[h.cell.ri]?.cells[h.cell.ci] : undefined
      return cell ? `${ti}.${h.cell.ri}.${h.cell.ci}.${Math.min(h.index, cell.blocks.length)}` : undefined
    }
    const prev = h.prevT ? this.topIndex(h.prevT) : -1
    if (prev >= 0) return String(prev + 1)
    const next = h.nextT ? this.topIndex(h.nextT) : -1
    return String(next >= 0 ? next : Math.min(h.index, this.st.blocks.length))
  }

  /** applies a model op through the editor; returns how to revert it */
  private run(op: DocxOp): Undo {
    const ed = this.ed
    if (op.t === 'text') {
      const locs = this.locate(op.pid)
      if (locs.length !== 1 || !ed.para(locs[0].pa)) return null
      const p = locs[0].p
      const next: Paragraph = { ...p, runs: spliceRuns(p.runs, op.at, op.del, op.ins.map((r) => this.remap(r))), edited: true }
      ed.replace([locs[0].pa], [next], this.sel(), { history: false, label: 'collab' })
      return [{ t: 'paras', before: [p], next: [next] }]
    }
    if (op.t === 'paras') {
      const next = op.next.map((p) => this.remap(p))
      const nextPids = new Set(next.map((p) => p.pid!))
      const pas = this.region(op.before)
      if (!pas || !pas.length) return null
      const steps: Step[] = []
      // the same paragraphs elsewhere (moved by a concurrent edit) go away first
      const inRegion = new Set(pas.map((pa) => (ed.para(pa)?.pid ?? -1)))
      for (const pid of uniq([...op.before, ...nextPids])) {
        if (inRegion.has(pid)) continue
        for (const loc of this.locate(pid).reverse()) {
          if (!ed.para(loc.pa)) continue
          const nb = (i: number) => {
            const b = loc.list[i]
            return b?.t === 'p' && b.pid !== undefined ? { pid: b.pid, k: this.locate(b.pid).findIndex((x) => x.list === loc.list && x.index === i) } : undefined
          }
          const step: Step = { t: 'paras', before: [loc.p], next: [], prev: nb(loc.index - 1), after: nb(loc.index + 1), home: this.home(loc) }
          ed.replace([loc.pa], [], this.sel(), { history: false, label: 'collab' })
          steps.push(step)
        }
      }
      const region = this.region(op.before)
      if (!region?.length) return steps
      const before = region.map((pa) => ed.para(pa)!).filter(Boolean)
      if (before.length !== region.length) return steps
      ed.replace(region, next, this.sel(), { history: false, label: 'collab' })
      steps.push({ t: 'paras', before, next })
      return steps
    }
    if (op.t === 'blocks') {
      let next = op.next.map((b) => this.remap(b))
      let bi: number
      let count = 0
      if (op.before.length) {
        bi = this.topIndex(op.before[0])
        if (bi < 0) {
          // first block gone: try the others
          const k = op.before.find((x) => this.topIndex(x) >= 0)
          if (!k) return null
          bi = this.topIndex(k)
        }
        while (bi + count < this.st.blocks.length && op.before.some((k) => matches(this.st.blocks[bi + count], k))) count++
      } else {
        bi = op.prev ? this.topIndex(op.prev) + 1 : 0
        if (op.prev && bi === 0) return null
      }
      const before = this.st.blocks.slice(bi, bi + count)
      // paragraphs around a table keep their current text: the op is about the blocks, not their content. One that a
      // concurrent edit took away stays away (its text went where that edit put it; the file wouldn't get it back)
      const cur = new Map(before.flatMap((b) => (b.t === 'p' && b.pid !== undefined ? [[b.pid, b] as const] : [])))
      const had = new Set(op.before.flatMap((k) => ('p' in k ? [k.p] : [])))
      next = next.flatMap((b) => (b.t === 'p' && b.pid !== undefined && had.has(b.pid) ? (cur.has(b.pid) ? [cur.get(b.pid)!] : []) : [b]))
      ed.replaceBlocks(bi, count, next, null, false)
      return [{ t: 'blocks', before, next, prev: bi > 0 ? bkey(this.st.blocks[bi - 1]) : null }]
    }
    return null
  }

  revert(undo: unknown) {
    const cm = revertComment(this.client, undo)
    if (cm) return cm
    const u = undo as Undo
    if (!u || !this.host) return
    this.capture()
    this.touchedModel = true
    this.suppress++
    try {
      for (let i = u.length - 1; i >= 0; i--) this.undoStep(u[i])
    } finally {
      this.suppress--
    }
  }

  private undoStep(u: Step) {
    const ed = this.ed
    if (u.t === 'paras') {
      if (!u.next.length) {
        // re-insert next to the neighbour it had
        const at = (n: { pid: number; k: number } | undefined) => (n ? this.locate(n.pid)[n.k] : undefined)
        const prev = at(u.prev)
        if (prev && ed.para(prev.pa)) return ed.replace([prev.pa], [prev.p, ...u.before], this.sel(), { history: false, label: 'collab' })
        const after = at(u.after)
        if (after && ed.para(after.pa)) return ed.replace([after.pa], [...u.before, after.p], this.sel(), { history: false, label: 'collab' })
        // no paragraph to place it by (it stood between tables, or alone in a cell): back where it was
        const pa = u.home && this.homePa(u.home)
        if (pa !== undefined) for (const p of [...u.before].reverse()) ed.restore(pa, p)
        return
      }
      const first = this.locate(u.next[0].pid!)[0]
      if (!first) return
      const pas = u.next.map((_, i) => first.pa.replace(/\d+$/, String(first.index + i)))
      if (pas.some((pa, i) => ed.para(pa)?.pid !== u.next[i].pid)) return
      ed.replace(pas, u.before, this.sel(), { history: false, label: 'collab' })
      return
    }
    const k = u.next.length ? bkey(u.next[0]) : null
    const bi = k ? this.topIndex(k) : u.prev ? this.topIndex(u.prev) + 1 : 0
    if (bi < 0) return
    const cur = new Map(this.st.blocks.slice(bi, bi + u.next.length).flatMap((b) => (b.t === 'p' && b.pid !== undefined ? [[b.pid, b] as const] : [])))
    ed.replaceBlocks(bi, u.next.length, u.before.map((b) => (b.t === 'p' && b.pid !== undefined && cur.has(b.pid) ? cur.get(b.pid)! : b)), null, false)
  }

  // ---------------------------------------------------------------- selection across remote edits

  private toD(pos: Pos): DPos | null {
    const p = this.ed.para(pos.pa)
    if (!p || p.pid === undefined) return null
    const locs = this.locate(p.pid)
    return { pid: p.pid, k: Math.max(0, locs.findIndex((l) => l.pa === pos.pa)), o: pos.o }
  }

  private fromD(d: DPos): Pos | null {
    const locs = this.locate(d.pid)
    const l = locs[Math.min(d.k, locs.length - 1)]
    return l ? { pa: l.pa, o: Math.min(d.o, plen(l.p.runs as RRun[])) } : null
  }

  /** remembers the selection (by paragraph id + text) before the model changes */
  private capture() {
    if (this.snap || !this.host) return
    const focused = !!this.ed.input.readSel()
    const sel = this.ed.input.sel()
    const a = sel && this.toD(sel.a)
    const f = sel && this.toD(sel.f)
    if (!a || !f) return
    const text = new Map<string, string>()
    for (const d of [a, f]) {
      const l = this.locate(d.pid)[d.k]
      if (l) text.set(`${d.pid}:${d.k}`, runsText(l.p.runs as RRun[]))
    }
    this.snap = { a, f, text, focused }
  }

  /** the remembered selection in the current model */
  private sel(): Sel {
    const s = this.snap
    const fallback = { pa: '0', o: 0 }
    if (!s) return this.ed.input.lastSel ?? { a: fallback, f: fallback }
    const map = (d: DPos): Pos => {
      const l = this.locate(d.pid)[d.k] ?? this.locate(d.pid)[0]
      if (!l) return this.ed.input.lastSel?.f ?? fallback
      const was = s.text.get(`${d.pid}:${d.k}`)
      const now = runsText(l.p.runs as RRun[])
      return { pa: l.pa, o: was === undefined ? Math.min(d.o, now.length) : mapByDiff(d.o, was, now) }
    }
    return { a: map(s.a), f: map(s.f) }
  }

  settled() {
    const s = this.snap
    const touched = this.touchedModel
    this.touchedModel = false
    if (!this.host) return
    if (!s) {
      // nobody was editing: don't let the collaborators' edits focus the document
      if (touched) this.ed.pendingSel = null
      return
    }
    const sel = this.sel()
    this.snap = null
    // keep the caret where the user was; never steal focus from other controls
    this.ed.pendingSel = s.focused ? sel : null
    if (!s.focused) this.ed.input.lastSel = sel
  }

  // ---------------------------------------------------------------- presence

  private onSel = () => {
    if (this.raf || !this.host) return
    this.raf = requestAnimationFrame(() => {
      this.raf = 0
      const sel = this.host && this.ed.input.readSel()
      if (!sel) return
      const a = this.toD(sel.a)
      const f = this.toD(sel.f)
      if (!a || !f) return
      const p: DocxPresence = { k: 'docx', a, f }
      const key = JSON.stringify(p)
      if (key === this.lastPresence) return
      this.lastPresence = key
      this.onPresence?.(p)
    })
  }

  /** screen rectangles of a collaborator's selection (caret = last rect) inside `root` */
  rects(root: HTMLElement, p: DocxPresence): { caret: DOMRect | null; sel: DOMRect[] } {
    if (!this.host) return { caret: null, sel: [] }
    const pages = root.querySelector<HTMLElement>('.de-pages')
    const a = this.fromD(p.a)
    const f = this.fromD(p.f)
    if (!pages || !a || !f) return { caret: null, sel: [] }
    const fd = domFromPos(pages, f)
    let caret: DOMRect | null = null
    if (fd) {
      const r = document.createRange()
      try {
        r.setStart(fd[0], fd[1])
        r.collapse(true)
        caret = r.getClientRects()[0] ?? (fd[0].nodeType === 1 ? (fd[0] as Element).getBoundingClientRect() : r.getBoundingClientRect())
      } catch {
        caret = null
      }
    }
    const sel: DOMRect[] = []
    const ad = domFromPos(pages, a)
    if (ad && fd && (a.pa !== f.pa || a.o !== f.o)) {
      const r = document.createRange()
      try {
        const forward = this.ed.cmp(a.pa, f.pa) < 0 || (a.pa === f.pa && a.o < f.o)
        if (forward) {
          r.setStart(ad[0], ad[1])
          r.setEnd(fd[0], fd[1])
        } else {
          r.setStart(fd[0], fd[1])
          r.setEnd(ad[0], ad[1])
        }
        for (const x of Array.from(r.getClientRects())) if (x.width > 0.5) sel.push(x)
      } catch {
        /* nodes detached */
      }
    }
    return { caret, sel }
  }
}
