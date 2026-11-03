/**
 * Main-thread DOCX editing host: owns undo history and paragraph mutations,
 * mirrors every change to the worker (which regenerates the XML).
 */
import { paraSpec } from '../../core/docx/spec'
import type { Block, CanonStyle, EditOp, EditResult, Inline, Paragraph, ParaDirty, ParaStyle, RunSet, RunStyle, StyleInfo, Table, TableSpec } from '../../core/docx/types'
import { cssFont } from '../../core/fonts'
import type { DocumentClient } from '../../core/worker/client'
import { opInsert, opInsertRuns, ordered, RichInput, type ReplaceOpts, type RichHost } from '../edit/input'
import { htmlToParas } from './html'
import { deleteCol, deleteRow, gridOf, insertCol, insertRow, newTable } from './tables'
import { collapsed, formatRange, normalize, plen, runAt, shiftPa, splitRuns, type Pos, type RRun, type Sel } from '../edit/rich'

export interface DocState {
  blocks: Block[]
  heights: WeakMap<Block, number>
  dirty: Set<Block>
  /** leading blocks measured at least once */
  frontier: number
}

interface Entry {
  pas: string[]
  before: Paragraph[]
  after: Paragraph[]
  /** top-level block replacement (tables) */
  blocks?: { bi: number; before: Block[]; after: Block[] }
  selBefore: Sel
  selAfter: Sel
  label: string
  t: number
}

export interface FormatState {
  b: boolean
  i: boolean
  u: boolean
  strike: boolean
  color?: string
  hl?: string
  sz?: number
  font?: string
  align?: string
  style?: string
  lst?: 'bullet' | 'number'
  line?: number
  va?: 'sup' | 'sub'
  link?: string
  /** caret inside a table cell */
  inTable: boolean
  /** a single image is selected */
  img?: { w: number; h: number; alt?: string; target: string; pa: string; o: number }
  canUndo: boolean
  canRedo: boolean
}

const path = (pa: string) => pa.split('.').map(Number)

/** selection recorded around address `from`, re-based onto `to` (same container) */
function moveSel(sel: Sel, from: string, to: string): Sel {
  if (from === to) return sel
  const dir = (pa: string) => pa.replace(/\.?\d+$/, '')
  const d = Number(to.split('.').pop()) - Number(from.split('.').pop())
  const mv = (p: Sel['a']) => (dir(p.pa) === dir(from) ? { ...p, pa: dir(to) ? `${dir(to)}.${Number(p.pa.split('.').pop()) + d}` : String(Number(p.pa.split('.').pop()) + d) } : p)
  return { a: mv(sel.a), f: mv(sel.f) }
}

/** a model change: contiguous paragraphs replaced, or top-level blocks replaced */
export type DocxCommit = { kind: 'paras'; before: Paragraph[]; next: Paragraph[] } | { kind: 'blocks'; bi: number; before: Block[]; next: Block[] }

export class DocxEditor implements RichHost<Paragraph> {
  readonly input: RichInput
  pendingSel: Sel | null = null
  styles: StyleInfo[] = []
  private history: Entry[] = []
  private future: Entry[] = []
  private pend: { s: Partial<RunStyle>; d: RunSet; at: { pa: string; o: number } } | null = null
  /** next id for new paragraphs / tables (a collaboration session gives each client its own range) */
  nextPid = 1_000_000_000
  private inflight = 0
  onState?: (s: FormatState | null) => void
  /** image selection changed (overlay with resize handles) */
  onImage?: (img: FormatState['img'] | null) => void
  /** style list changed (a built-in style was added to the document) */
  onStyles?: (styles: StyleInfo[]) => void
  shortcut?: (e: KeyboardEvent) => boolean
  onDirty?: () => void
  /** paint format: run formatting copied from the caret, applied to the next selection */
  painter: RunSet | null = null
  onPainter?: (on: boolean) => void
  /** every model change, after it was mirrored to the worker (collaboration hook) */
  onCommit?: (c: DocxCommit) => void

  constructor(
    private client: DocumentClient,
    private st: DocState,
    private changed: () => void,
  ) {
    this.input = new RichInput(this)
  }

  // ---------------------------------------------------------------- addressing

  private container(pa: string): { list: Block[]; at: number; cell?: [number, number, number] } | undefined {
    const p = path(pa)
    if (p.length === 1) return { list: this.st.blocks, at: p[0] }
    const t = this.st.blocks[p[0]]
    if (t?.t !== 'tbl') return undefined
    const cell = t.rows[p[1]]?.cells[p[2]]
    return cell ? { list: cell.blocks, at: p[3], cell: [p[0], p[1], p[2]] } : undefined
  }

  para(pa: string): Paragraph | undefined {
    const c = this.container(pa)
    const b = c?.list[c.at]
    return b && b.t === 'p' && b.pid !== undefined && !b.ro ? b : undefined
  }

  span(a: string, b: string): string[] | null {
    const pa = path(a)
    const pb = path(b)
    if (pa.length !== pb.length || pa.slice(0, -1).join('.') !== pb.slice(0, -1).join('.')) return null
    const out: string[] = []
    for (let i = pa[pa.length - 1]; i <= pb[pb.length - 1]; i++) {
      const addr = shiftPa(a, i - pa[pa.length - 1])
      if (!this.para(addr)) return null
      out.push(addr)
    }
    return out
  }

  prev(pa: string) {
    const at = path(pa).pop()!
    const q = shiftPa(pa, -1)
    return at > 0 && this.para(q) ? q : undefined
  }

  next(pa: string) {
    const q = shiftPa(pa, 1)
    return this.para(q) ? q : undefined
  }

  cmp(a: string, b: string): number {
    const x = path(a)
    const y = path(b)
    for (let i = 0; i < Math.min(x.length, y.length); i++) if (x[i] !== y[i]) return x[i] - y[i]
    return x.length - y.length
  }

  /** the whole document: from the first to the last editable top-level paragraph */
  all(): Sel | null {
    let a = -1
    let f = -1
    this.st.blocks.forEach((_, i) => {
      if (!this.para(String(i))) return
      if (a < 0) a = i
      f = i
    })
    if (a < 0) return null
    return { a: { pa: String(a), o: 0 }, f: { pa: String(f), o: plen(this.para(String(f))!.runs as RRun[]) } }
  }

  selectAll() {
    this.input.selectAll()
  }

  /** Deletes a top-level range that spans tables: the end paragraphs are joined, everything between goes (section breaks stay). */
  deleteAcross(s: Pos, e: Pos): Pos | null {
    const a = path(s.pa)
    const b = path(e.pa)
    if (a.length !== 1 || b.length !== 1 || a[0] >= b[0]) return null
    const first = this.para(s.pa)
    const last = this.para(e.pa)
    if (!first || !last) return null
    // parts of a paragraph split around a page break share its pid: the worker regenerates whole paragraphs only
    const part = (i: number, p: Paragraph) => this.st.blocks[i]?.t === 'p' && (this.st.blocks[i] as Paragraph).pid === p.pid
    if (part(a[0] - 1, first) || part(b[0] + 1, last)) return null
    const runs = normalize([...splitRuns(first.runs as RRun[], s.o)[0], ...splitRuns(last.runs as RRun[], e.o)[1]])
    const keep = this.st.blocks.slice(a[0] + 1, b[0]).filter((x) => x.t === 'sect')
    this.replaceBlocks(a[0], b[0] - a[0] + 1, [this.join(first, last, runs), ...keep], { a: s, f: s })
    return s
  }

  // ---------------------------------------------------------------- paragraph factories

  withRuns(p: Paragraph, runs: RRun[]): Paragraph {
    return { ...p, runs: runs as Inline[], edited: true }
  }

  private origOf(p: Paragraph): number[] {
    return p.orig ?? (p.pid === (p.from ?? p.pid) ? [p.pid!] : [])
  }

  defaultRun(p: Paragraph): Partial<RRun> | undefined {
    if (p.caret) return { ...(p.caret as RRun), text: '' }
    const last = [...p.runs].reverse().find((r) => r.t === 'text') as RRun | undefined
    if (last) return { ...last, text: '' }
    return { t: 'text', s: { sz: p.markSz, font: p.markFont } } as Partial<RRun>
  }

  split(p: Paragraph, left: RRun[], right: RRun[], atEnd: boolean): [Paragraph, Paragraph] {
    const keep = p.keepSect ?? p.pid === (p.from ?? p.pid)
    const a: Paragraph = { ...p, runs: left as Inline[], edited: true, keepSect: false }
    const b: Paragraph = {
      ...p,
      runs: right as Inline[],
      pid: this.nextPid++,
      from: p.from ?? p.pid,
      orig: [],
      edited: true,
      keepSect: keep,
      pb: false,
      ver: 0,
      caret: right.length ? undefined : ((runAt(left, plen(left)) as Inline | undefined) ?? p.caret),
    }
    if (atEnd) {
      const next = this.styles.find((s) => s.id === p.sid)?.next
      if (next && next !== p.sid) {
        b.pd = { ...(b.pd ?? {}), style: next }
        b.sid = next
      }
    }
    return [a, b]
  }

  join(a: Paragraph, b: Paragraph, runs: RRun[]): Paragraph {
    return { ...a, runs: runs as Inline[], edited: true, orig: [...this.origOf(a), ...this.origOf(b)], keepSect: (a.keepSect ?? a.pid === (a.from ?? a.pid)) || (b.keepSect ?? b.pid === (b.from ?? b.pid)) }
  }

  // ---------------------------------------------------------------- mutation core

  private splice(pa: string, count: number, next: Block[]): Block[] {
    const c = this.container(pa)!
    if (!c.cell) {
      this.st.blocks.splice(c.at, count, ...next)
      return this.st.blocks
    }
    const [bi, ri, ci] = c.cell
    const t = this.st.blocks[bi] as Table
    const blocks = [...c.list]
    blocks.splice(c.at, count, ...next)
    const nt: Table = { ...t, rows: t.rows.map((r, i) => (i !== ri ? r : { ...r, cells: r.cells.map((cell, j) => (j !== ci ? cell : { ...cell, blocks })) })) }
    this.st.heights.set(nt, this.st.heights.get(t) ?? 0)
    this.st.dirty.add(nt)
    this.st.blocks[bi] = nt
    return blocks
  }

  replace(pas: string[], next: Paragraph[], sel: Sel, opts: ReplaceOpts & { history?: boolean; want?: boolean } = {}): void {
    const before = pas.map((pa) => this.para(pa)!)
    if (before.some((b) => !b)) return
    if (opts.remount) next = next.map((p) => ({ ...p, ver: (p.ver ?? 0) + 1 }))
    const c = this.container(pas[0])!
    const at = c.at
    const list = this.splice(pas[0], pas.length, next)
    // split paragraphs: trailing page-break parts follow the paragraph that now ends the range
    const lastB = before[before.length - 1]
    const lastN = next[next.length - 1]
    if (lastN && lastB.pid !== lastN.pid) {
      for (let k = at + next.length; k < list.length; k++) {
        const b = list[k]
        if (b.t !== 'p' || b.pid !== lastB.pid || !b.pb) break
        list[k] = { ...b, pid: lastN.pid, from: lastN.from ?? lastN.pid, edited: true }
      }
    }
    // heights: estimate from what was there, remeasure soon
    if (!c.cell) {
      next.forEach((p, i) => {
        this.st.heights.set(p, this.st.heights.get(before[i] ?? before[0]) ?? 0)
        this.st.dirty.add(p)
      })
      for (const k of [at - 1, at + next.length]) if (list[k]) this.st.dirty.add(list[k])
      if (at < this.st.frontier) this.st.frontier = Math.max(0, this.st.frontier + next.length - pas.length)
    }
    if (opts.history !== false) {
      const now = Date.now()
      const last = this.history[this.history.length - 1]
      const label = opts.label ?? 'edit'
      const selBefore = this.input.lastSel ?? sel
      if (label === 'type' && last?.label === 'type' && pas.length === 1 && next.length === 1 && last.after.length === 1 && last.pas[0] === pas[0] && now - last.t < 1500) {
        last.after = next
        last.selAfter = sel
        last.t = now
      } else {
        this.history.push({ pas, before, after: next, selBefore, selAfter: sel, label, t: now })
        if (this.history.length > 500) this.history.shift()
      }
      this.future = []
    }
    this.commit(pas[0], before, next, opts.want)
    this.onCommit?.({ kind: 'paras', before, next })
    this.pendingSel = sel
    this.onDirty?.()
    this.changed()
    // the caret may not move (formatting): refresh the toolbar from the new runs
    this.emitState(sel)
  }

  /** contiguous blocks sharing `pid` around index `near` */
  private partsOf(list: Block[], pid: number, near: number): Paragraph[] {
    let i = near
    while (i > 0 && (list[i - 1] as Paragraph).pid === pid && list[i - 1].t === 'p') i--
    const out: Paragraph[] = []
    for (; i < list.length && list[i].t === 'p' && (list[i] as Paragraph).pid === pid; i++) out.push(list[i] as Paragraph)
    return out
  }

  private commit(pa: string, before: Paragraph[], next: Paragraph[], want?: boolean) {
    const c = this.container(pa)!
    if (c.cell && (this.st.blocks[c.cell[0]] as Table).tid !== undefined) {
      this.commitTable(c.cell[0])
      return
    }
    const list = c.list
    const pids: number[] = []
    next.forEach((p) => p.pid !== undefined && !pids.includes(p.pid) && pids.push(p.pid))
    const specs = pids.map((pid) => {
      const idx = c.at + next.findIndex((p) => p.pid === pid)
      return paraSpec(this.partsOf(list, pid, idx))
    })
    const remove: number[] = []
    for (const b of before) {
      if (b.pid === undefined || pids.includes(b.pid) || remove.includes(b.pid)) continue
      const still = list.findIndex((x) => x.t === 'p' && (x as Paragraph).pid === b.pid)
      if (still >= 0) specs.push(paraSpec(this.partsOf(list, b.pid, still)))
      else remove.push(b.pid)
    }
    const anchor = before[0].pid!
    const structural = want || next.some((p) => p.pd !== before.find((b) => b.pid === p.pid)?.pd)
    const sent = new Set<Paragraph>(next)
    const op: EditOp = { anchor, list: specs, remove, want: structural }
    this.inflight++
    this.client
      .call('docx.edit', op)
      .then((res) => this.applyResult(res, sent))
      .catch((e) => console.error('[office-engine] edit failed', e))
      .finally(() => this.inflight--)
  }

  get busy() {
    return this.inflight > 0
  }

  private locate(pid: number): { list: Block[]; index: number; pa: string }[] {
    const out: { list: Block[]; index: number; pa: string }[] = []
    this.st.blocks.forEach((b, bi) => {
      if (b.t === 'p' && b.pid === pid) out.push({ list: this.st.blocks, index: bi, pa: String(bi) })
      else if (b.t === 'tbl')
        b.rows.forEach((r, ri) =>
          r.cells.forEach((cell, ci) =>
            cell.blocks.forEach((x, pi) => {
              if (x.t === 'p' && x.pid === pid) out.push({ list: cell.blocks, index: pi, pa: `${bi}.${ri}.${ci}.${pi}` })
            }),
          ),
        )
    })
    return out
  }

  private put(pa: string, p: Paragraph) {
    const c = this.container(pa)
    if (!c) return
    const old = c.list[c.at]
    this.splice(pa, 1, [p])
    if (!c.cell) {
      this.st.heights.set(p, this.st.heights.get(old) ?? 0)
      this.st.dirty.add(p)
    }
  }

  private applyResult(res: EditResult, sent: Set<Paragraph>) {
    if (this.input.isComposing) return
    let touched = false
    // re-converted paragraphs (style / list changes): only if not edited again meanwhile
    for (const [k, blocks] of Object.entries(res.blocks)) {
      const locs = this.locate(Number(k))
      const paras = blocks.filter((b): b is Paragraph => b.t === 'p')
      if (!locs.length || paras.length !== locs.length) continue
      if (!locs.every((l) => sent.has(l.list[l.index] as Paragraph))) continue
      locs.forEach((l, i) => {
        const cur = l.list[l.index] as Paragraph
        this.put(l.pa, { ...paras[i], from: cur.from, orig: cur.orig, pd: cur.pd, edited: cur.edited, keepSect: cur.keepSect, ver: cur.ver, pb: cur.pb })
      })
      touched = true
    }
    // list numbering changes anywhere in the document
    for (const [k, m] of Object.entries(res.markers)) {
      for (const loc of this.locate(Number(k))) {
        const p = loc.list[loc.index] as Paragraph
        if ((loc.list[loc.index - 1] as Paragraph | undefined)?.pid === p.pid) continue
        this.put(loc.pa, { ...p, marker: m ? { text: m.text, s: m.s, tab: m.tab } : undefined, lst: m?.lst })
        touched = true
      }
    }
    if (touched) this.changed()
  }

  // ---------------------------------------------------------------- history

  /** current address of the first block of `list` (paragraphs by pid, tables by identity), `empty` for an empty list; undefined when gone */
  private whereNow(list: Block[], empty: string): string | undefined {
    const b = list[0]
    if (!b) return empty
    if (b.t === 'p') return b.pid !== undefined ? this.locate(b.pid)[0]?.pa : undefined
    if (b.t === 'tbl') {
      const i = this.st.blocks.findIndex((x) => x.t === 'tbl' && ((b.tid !== undefined && x.tid === b.tid) || (b.item !== undefined && x.item === b.item)))
      return i >= 0 ? String(i) : undefined
    }
    return empty
  }

  /** table undo / redo: paragraphs still in place keep their current text (edited since, maybe by collaborators) */
  private keepCurrent(list: Block[], bi: number, count: number): Block[] {
    const cur = new Map(this.st.blocks.slice(bi, bi + count).flatMap((b) => (b.t === 'p' && b.pid !== undefined ? [[b.pid, b] as const] : [])))
    return list.map((b) => (b.t === 'p' && b.pid !== undefined && cur.has(b.pid) ? cur.get(b.pid)! : b))
  }

  undo() {
    const e = this.history.pop()
    if (!e) return
    if (e.blocks) {
      const bi = this.whereNow(e.blocks.after, String(e.blocks.bi))
      // the blocks were removed by a collaborator: nothing left to undo
      if (bi === undefined) return
      this.replaceBlocks(Number(bi), e.blocks.after.length, this.keepCurrent(e.blocks.before, Number(bi), e.blocks.after.length), e.selBefore, false)
      this.future.push(e)
      return
    }
    const at = this.whereNow(e.after, e.pas[0])
    if (at === undefined) return
    const pasAfter = e.after.map((_, i) => shiftPa(at, i))
    // a collaborator changed these paragraphs' structure meanwhile: drop the entry
    if (pasAfter.some((pa, i) => this.para(pa)?.pid !== e.after[i].pid)) return
    this.replace(pasAfter.length ? pasAfter : [at], e.before, moveSel(e.selBefore, e.pas[0], at), { history: false, label: 'undo', want: true })
    this.future.push(e)
  }

  redo() {
    const e = this.future.pop()
    if (!e) return
    if (e.blocks) {
      const bi = this.whereNow(e.blocks.before, String(e.blocks.bi))
      if (bi === undefined) return
      this.replaceBlocks(Number(bi), e.blocks.before.length, this.keepCurrent(e.blocks.after, Number(bi), e.blocks.before.length), e.selAfter, false)
      this.history.push(e)
      return
    }
    const at = this.whereNow(e.before, e.pas[0])
    if (at === undefined) return
    const pasNow = e.before.map((_, i) => shiftPa(at, i))
    if (pasNow.some((pa, i) => this.para(pa)?.pid !== e.before[i].pid)) return
    this.replace(pasNow, e.after, moveSel(e.selAfter, e.pas[0], at), { history: false, label: 'redo', want: true })
    this.history.push(e)
  }

  get canUndo() {
    return this.history.length > 0
  }
  get canRedo() {
    return this.future.length > 0
  }

  // ---------------------------------------------------------------- formatting

  pendingStyle(): Partial<RRun> | undefined {
    return this.pend ? ({ s: this.pend.s, d: this.pend.d } as Partial<RRun>) : undefined
  }

  clearPending() {
    this.pend = null
  }

  private selectedParas(sel: Sel): { pas: string[]; s: { pa: string; o: number }; e: { pa: string; o: number } } | null {
    const [s, e] = ordered(this, sel)
    const pas = s.pa === e.pa ? (this.para(s.pa) ? [s.pa] : null) : this.span(s.pa, e.pa)
    return pas ? { pas, s, e } : null
  }

  format(s: Partial<RunStyle>, d: RunSet, opts: { want?: boolean } = {}) {
    const sel = this.input.sel()
    if (!sel) return
    if (collapsed(sel)) {
      this.pend = { s: { ...(this.pend?.s ?? {}), ...s }, d: { ...(this.pend?.d ?? {}), ...d }, at: sel.a }
      this.input.restore(sel)
      this.emitState(sel)
      return
    }
    const r = this.selectedParas(sel)
    if (!r) return
    const next = r.pas.map((pa, i) => {
      const p = this.para(pa)!
      const a = i === 0 ? r.s.o : 0
      const b = i === r.pas.length - 1 ? r.e.o : plen(p.runs as RRun[])
      return this.withRuns(p, formatRange(p.runs as RRun[], a, b, s, d))
    })
    this.replace(r.pas, next, sel, { label: 'format', want: opts.want })
  }

  toggle(key: 'b' | 'i' | 'u' | 'strike') {
    const st = this.state()
    const on = !(st && st[key])
    if (key === 'u') this.format({ u: on ? 'single' : 'none' }, { u: on ? 'single' : 'none' })
    else this.format({ [key]: on }, { [key]: on })
  }

  setFont(name: string) {
    this.format({ font: cssFont(name) }, { font: name })
  }

  setSize(sz: number) {
    this.format({ sz }, { sz })
  }

  setColor(c: string | null) {
    this.format({ color: c ?? undefined }, { color: c })
  }

  setHighlight(c: string | null) {
    this.format({ hl: c ?? undefined, shd: undefined }, { hl: c })
  }

  clearFormat() {
    this.format({ b: false, i: false, u: 'none', strike: false, hl: undefined }, { b: false, i: false, u: 'none', strike: false, color: null, hl: null, va: null }, { want: true })
  }

  setPara(patch: ParaDirty, optimistic: Partial<ParaStyle> = {}, sel = this.input.sel()) {
    if (!sel) return
    const r = this.selectedParas(sel)
    if (!r) return
    const next = r.pas.map((pa) => {
      const p = this.para(pa)!
      const pd = { ...(p.pd ?? {}), ...patch }
      return { ...p, pd, ps: { ...p.ps, ...optimistic }, edited: true, sid: patch.style ?? p.sid }
    })
    this.replace(r.pas, next, sel, { label: 'para', want: true })
  }

  toggleList(kind: 'bullet' | 'number') {
    const sel = this.input.sel()
    if (!sel) return
    const r = this.selectedParas(sel)
    if (!r) return
    const all = r.pas.every((pa) => this.para(pa)?.lst === kind)
    this.setPara({ list: all ? 'none' : kind, ilvl: all ? undefined : 0 }, {}, sel)
  }

  indent(dir: 1 | -1, sel = this.input.sel()) {
    if (!sel) return
    const r = this.selectedParas(sel)
    if (!r) return
    const next = r.pas.map((pa) => {
      const p = this.para(pa)!
      const pd: ParaDirty = { ...(p.pd ?? {}) }
      if (p.lst) pd.ilvl = Math.max(0, Math.min(8, (p.ilvl ?? 0) + dir))
      else pd.indL = Math.max(0, (p.ps.indL ?? 0) + dir * 36)
      return { ...p, pd, ps: p.lst ? p.ps : { ...p.ps, indL: pd.indL }, edited: true }
    })
    this.replace(r.pas, next, sel, { label: 'para', want: true })
  }

  tab(sel: Sel, shift: boolean): boolean {
    if (sel.f.pa.includes('.')) {
      // move between table cells (Tab on the last cell adds a row)
      const [bi, ri, ci] = sel.f.pa.split('.').map(Number)
      const t = this.st.blocks[bi] as Table
      const cells = t.rows.flatMap((r, i) => r.cells.map((_, j) => [i, j] as [number, number]))
      const k = cells.findIndex(([i, j]) => i === ri && j === ci) + (shift ? -1 : 1)
      if (k >= cells.length && !shift) {
        this.tableOp('rowBelow')
        return true
      }
      const [nr, nc] = cells[Math.max(0, k)]
      const pa = `${bi}.${nr}.${nc}.0`
      const p = this.para(pa)
      this.input.restore({ a: { pa, o: 0 }, f: { pa, o: p ? plen(p.runs as RRun[]) : 0 } })
      return true
    }
    const p = this.para(sel.a.pa)
    if (!p?.lst) return false
    this.indent(shift ? -1 : 1, sel)
    return true
  }

  backspaceAtStart(pa: string, p: Paragraph): boolean {
    if (!p.lst) return false
    this.setPara({ list: 'none' }, {}, { a: { pa, o: 0 }, f: { pa, o: 0 } })
    return true
  }

  // ---------------------------------------------------------------- blocks (tables)

  /** content width (pt) of the page, for sizing new tables / images */
  contentW = 468
  /** document default font (CSS) / size, shown when a run inherits them */
  defaultFont?: string
  defaultSz?: number

  private prevParaPid(bi: number): number | undefined {
    for (let i = bi - 1; i >= 0; i--) {
      const b = this.st.blocks[i]
      if (b.t === 'p' && b.pid !== undefined) return b.pid
    }
    return undefined
  }

  private tableSpec(t: Table): TableSpec {
    return {
      tid: t.tid!,
      item: t.item,
      cols: t.cols,
      rows: t.rows.map((r) => ({
        src: r.src,
        cells: r.cells.map((c) => ({
          src: c.src,
          colSpan: c.colSpan,
          rowSpan: c.rowSpan,
          paras: c.blocks.filter((b): b is Paragraph => b.t === 'p' && b.pid !== undefined).map((p) => paraSpec([p])),
        })),
      })),
    }
  }

  private commitTable(bi: number) {
    const t = this.st.blocks[bi] as Table
    if (t?.t !== 'tbl' || t.tid === undefined) return
    this.send('docx.table', { tid: t.tid, item: t.item, anchor: t.item === undefined ? this.prevParaPid(bi) : undefined, spec: this.tableSpec(t) })
  }

  private send(method: 'docx.table' | 'docx.edit', op: unknown) {
    this.inflight++
    ;(this.client.call as (m: string, ...a: unknown[]) => Promise<unknown>)(method, op)
      .then((res) => method === 'docx.edit' && this.applyResult(res as EditResult, new Set()))
      .catch((e) => console.error('[office-engine] edit failed', e))
      .finally(() => this.inflight--)
  }

  /** Replaces top-level blocks (table insert/delete/restructure) with history + write-back. */
  replaceBlocks(bi: number, count: number, next: Block[], sel: Sel | null, history = true) {
    const before = this.st.blocks.slice(bi, bi + count)
    this.st.blocks.splice(bi, count, ...next)
    next.forEach((b, i) => {
      this.st.heights.set(b, this.st.heights.get(before[i] ?? before[0]) ?? 0)
      this.st.dirty.add(b)
    })
    if (bi < this.st.frontier) this.st.frontier = Math.max(0, this.st.frontier + next.length - count)
    const key = (b: Block) => (b.t === 'tbl' ? `t${b.tid ?? 'i' + b.item}` : b.t === 'p' ? `p${b.pid}` : '')
    const nextKeys = new Set(next.map(key))
    const beforeKeys = new Set(before.map(key))
    // paragraphs that appear / disappear
    const anchorP = [...before, ...next].find((b): b is Paragraph => b.t === 'p' && b.pid !== undefined && beforeKeys.has(key(b)) && nextKeys.has(key(b)))
    const added = next.filter((b): b is Paragraph => b.t === 'p' && b.pid !== undefined && !beforeKeys.has(key(b)))
    const removed = before.filter((b): b is Paragraph => b.t === 'p' && b.pid !== undefined && !nextKeys.has(key(b)))
    // paragraphs kept but replaced by another version (undo, collaborators) must reach the worker too
    const beforeObj = new Map(before.map((b) => [key(b), b]))
    const changed = next.some((b) => b.t === 'p' && b.pid !== undefined && beforeObj.has(key(b)) && beforeObj.get(key(b)) !== b)
    if (anchorP && (added.length || removed.length || changed)) {
      const list = next.filter((b): b is Paragraph => b.t === 'p' && b.pid !== undefined)
      this.send('docx.edit', { anchor: anchorP.pid, list: list.map((p) => paraSpec([p])), remove: removed.map((p) => p.pid) } as EditOp)
    }
    for (const b of before) if (b.t === 'tbl' && !nextKeys.has(key(b)) && (b.tid !== undefined || b.item !== undefined)) this.send('docx.table', { tid: b.tid ?? -1, item: b.item, spec: null })
    next.forEach((b, i) => {
      if (b.t !== 'tbl') return
      // an original table coming back (undo of its deletion) is regenerated from the model like a new one
      if (b.tid === undefined && !beforeKeys.has(key(b))) this.st.blocks[bi + i] = { ...b, tid: this.nextPid++ }
      this.commitTable(bi + i)
    })
    this.onCommit?.({ kind: 'blocks', bi, before, next })
    if (history) {
      const s = sel ?? this.input.lastSel ?? { a: { pa: String(bi), o: 0 }, f: { pa: String(bi), o: 0 } }
      this.history.push({ pas: [], before: [], after: [], blocks: { bi, before, after: next }, selBefore: this.input.lastSel ?? s, selAfter: s, label: 'table', t: Date.now() })
      this.future = []
    }
    if (sel) this.pendingSel = sel
    this.onDirty?.()
    this.changed()
  }

  private emptyPara(tpl: Paragraph): Paragraph {
    return {
      t: 'p', ps: { ...tpl.ps, indFirst: undefined, indL: undefined }, runs: [], markSz: tpl.markSz, markFont: tpl.markFont,
      pid: this.nextPid++, from: tpl.from ?? tpl.pid, orig: [], edited: true, keepSect: false, pd: { style: '', list: 'none' }, sid: undefined,
    }
  }

  insertTable(rowsN: number, colsN: number) {
    const sel = this.input.sel()
    const bi = sel ? Number(sel.f.pa.split('.')[0]) : -1
    const p = this.st.blocks[bi]
    if (!sel || sel.f.pa.includes('.') || p?.t !== 'p' || p.pid === undefined) return
    const t: Table = { ...newTable(rowsN, colsN, this.contentW, () => this.emptyPara(p)), tid: this.nextPid++ }
    const after: Block[] = [p, t]
    const nextBlock = this.st.blocks[bi + 1]
    if (!nextBlock || nextBlock.t !== 'p') after.push(this.emptyPara(p))
    this.replaceBlocks(bi, 1, after, { a: { pa: `${bi + 1}.0.0.0`, o: 0 }, f: { pa: `${bi + 1}.0.0.0`, o: 0 } })
  }

  /** Structural table edit at the caret's cell. */
  tableOp(kind: 'rowAbove' | 'rowBelow' | 'colLeft' | 'colRight' | 'delRow' | 'delCol' | 'delTable') {
    const sel = this.input.sel()
    if (!sel || !sel.f.pa.includes('.')) return
    const [bi, ri, ci] = sel.f.pa.split('.').map(Number)
    const t = this.st.blocks[bi] as Table
    if (t?.t !== 'tbl') return
    const g = gridOf(t).cells.find((c) => c.ri === ri && c.ci === ci)?.g ?? 0
    const mk = (tpl: { blocks: Block[] }) => {
      const p = tpl.blocks.find((b): b is Paragraph => b.t === 'p') ?? (this.st.blocks.find((b) => b.t === 'p' && b.pid !== undefined) as Paragraph)
      return { ...this.emptyPara(p), pd: {}, ps: p.ps, sid: p.sid }
    }
    let next: Table | null = t
    let caret = `${bi}.${ri}.${ci}.0`
    switch (kind) {
      case 'rowAbove':
        next = insertRow(t, ri, false, mk)
        caret = `${bi}.${ri}.0.0`
        break
      case 'rowBelow':
        next = insertRow(t, ri, true, mk)
        caret = `${bi}.${ri + 1}.0.0`
        break
      case 'colLeft':
        next = insertCol(t, g, false, mk)
        break
      case 'colRight':
        next = insertCol(t, g, true, mk)
        caret = `${bi}.${ri}.${ci + 1}.0`
        break
      case 'delRow':
        next = deleteRow(t, ri)
        caret = `${bi}.${Math.max(0, Math.min(ri, (next?.rows.length ?? 1) - 1))}.0.0`
        break
      case 'delCol':
        next = deleteCol(t, g)
        caret = `${bi}.${ri}.${Math.max(0, ci - 1)}.0`
        break
      case 'delTable':
        next = null
        break
    }
    if (!next) {
      const after = this.st.blocks[bi + 1]
      this.replaceBlocks(bi, 1, [], after?.t === 'p' ? { a: { pa: String(bi), o: 0 }, f: { pa: String(bi), o: 0 } } : null)
      return
    }
    next = { ...next, tid: t.tid ?? this.nextPid++ }
    this.replaceBlocks(bi, 1, [next], { a: { pa: caret, o: 0 }, f: { pa: caret, o: 0 } })
  }

  // ---------------------------------------------------------------- inserts

  async insertImages(files: Blob[], at?: Sel) {
    const sel = at ?? this.input.sel()
    if (!sel) return
    const runs: Inline[] = []
    for (const f of files) {
      const ext = (f.type.split('/')[1] ?? 'png').replace('svg+xml', 'svg').replace('jpeg', 'jpg')
      const { rid, path } = await this.client.call('docx.image', await f.arrayBuffer(), ext)
      const dims = await imageSize(f)
      const scale = Math.min(1, this.contentW / Math.max(1, dims.w * 0.75))
      runs.push({ t: 'img', target: path, w: +(dims.w * 0.75 * scale).toFixed(1), h: +(dims.h * 0.75 * scale).toFixed(1), rid })
    }
    opInsertRuns(this, sel, [runs as RRun[]], { label: 'image' })
  }

  async setLink(url: string | null) {
    const sel = this.input.sel()
    if (!sel) return
    if (!url) {
      this.patchRuns(sel, (r) => ({ ...r, href: undefined, lk: undefined }))
      return
    }
    const href = /^[a-z][\w+.-]*:/i.test(url) || url.startsWith('#') ? url : 'https://' + url
    const lk = await this.client.call('docx.link', href)
    const style = { color: '#0563c1', u: 'single' }
    if (collapsed(sel)) {
      opInsertRuns(this, sel, [[{ t: 'text', text: url, href, lk, s: { ...(runAt(this.para(sel.a.pa)?.runs as RRun[] ?? [], sel.a.o)?.s ?? {}), ...style }, d: style } as RRun]], { label: 'link' })
      return
    }
    this.patchRuns(sel, (r) => ({ ...r, href, lk, s: { ...(r.s ?? {}), ...style }, d: { ...(r.d ?? {}), ...style } }))
  }

  /** applies `fn` to the text runs inside the selection */
  private patchRuns(sel: Sel, fn: (r: RRun) => RRun) {
    const r = this.selectedParas(sel)
    if (!r) return
    const next = r.pas.map((pa, i) => {
      const p = this.para(pa)!
      const a = i === 0 ? r.s.o : 0
      const b = i === r.pas.length - 1 ? r.e.o : plen(p.runs as RRun[])
      const [left, rest] = splitRuns(p.runs as RRun[], a)
      const [mid, right] = splitRuns(rest, b - a)
      return this.withRuns(p, [...left, ...mid.map((x) => (x.t === 'text' ? fn(x) : x)), ...right])
    })
    this.replace(r.pas, next, sel, { label: 'format' })
  }

  insertPageBreak() {
    const sel = this.input.sel()
    if (!sel) return
    const p = this.para(sel.f.pa)
    if (!p || sel.f.pa.includes('.')) return
    const [l, r] = splitRuns(p.runs as RRun[], sel.f.o)
    const [a, b] = this.split(p, l, r, false)
    const b2 = { ...b, pb: true, pd: { ...(b.pd ?? {}), pageBreak: true } }
    this.replace([sel.f.pa], [a, b2], { a: { pa: shiftPa(sel.f.pa, 1), o: 0 }, f: { pa: shiftPa(sel.f.pa, 1), o: 0 } }, { label: 'pagebreak', want: true })
  }

  /** scales the selected image (factor) or fits it to the page width */
  resizeImage(f: number | 'fit') {
    const sel = this.input.sel()
    if (!sel) return
    const [s] = ordered(this, sel)
    this.patchObjects(sel, (x) => {
      if (x.t !== 'img') return x
      const k = f === 'fit' ? this.contentW / x.w : f
      return { ...x, w: +(x.w * k).toFixed(1), h: +(x.h * k).toFixed(1), resized: true }
    })
    this.pendingSel = { a: s, f: { ...s, o: s.o + 1 } }
  }

  private patchObjects(sel: Sel, fn: (x: Inline) => Inline) {
    const [s, e] = ordered(this, sel)
    const p = this.para(s.pa)
    if (!p || s.pa !== e.pa) return
    const [left, rest] = splitRuns(p.runs as RRun[], s.o)
    const [mid, right] = splitRuns(rest, e.o - s.o)
    this.replace([s.pa], [this.withRuns(p, [...left, ...(mid as Inline[]).map(fn), ...right] as RRun[])], sel, { label: 'image' })
  }

  pasteRich(sel: Sel, data: DataTransfer): boolean {
    const files = Array.from(data.files ?? []).filter((f) => f.type.startsWith('image/'))
    if (files.length) {
      void this.insertImages(files, sel)
      return true
    }
    const html = data.getData('text/html')
    if (!html) return false
    const paras = htmlToParas(html)
    if (!paras.length) return false
    void (async () => {
      const links = new Map<string, string>()
      for (const p of paras) for (const r of p.runs) if (r.t === 'text' && r.href && !links.has(r.href)) links.set(r.href, await this.client.call('docx.link', r.href))
      const lines = paras.map((p) => p.runs.map((r) => (r.t === 'text' && r.href ? { ...r, lk: links.get(r.href) } : r)) as RRun[])
      const headingId = (n: number) => this.styles.find((s) => s.name.toLowerCase() === `heading ${n}`)?.id
      const props = paras.map((p) => {
        if (!p.heading && !p.list) return undefined
        return (q: Paragraph) => {
          const pd: ParaDirty = { ...(q.pd ?? {}) }
          const h = p.heading ? headingId(p.heading) : undefined
          if (h) pd.style = h
          if (p.list) pd.list = p.list
          return { ...q, pd, sid: h ?? q.sid, edited: true }
        }
      })
      opInsertRuns(this, sel, lines, { label: 'paste' }, props as unknown as ((p: RichParaLike) => RichParaLike)[])
    })()
    return true
  }

  // ---------------------------------------------------------------- styles

  /** applies a built-in style (Title, Heading 1…), creating it in the document when missing */
  async applyCanon(kind: CanonStyle) {
    const sel = this.input.sel()
    if (!sel) return
    let id = this.styles.find((s) => s.canon === kind)?.id
    if (!id) {
      const r = await this.client.call('docx.style', kind)
      this.styles = r.styles
      this.onStyles?.(r.styles)
      id = r.id
    }
    if (id) this.setPara({ style: id }, {}, sel)
  }

  // ---------------------------------------------------------------- images

  /** selects the inline object at (pa, o) */
  selectObject(pa: string, o: number) {
    const sel = { a: { pa, o }, f: { pa, o: o + 1 } }
    this.input.restore(sel)
    this.onSelection(sel)
  }

  private imageSel(): Sel | null {
    const sel = this.input.sel()
    const img = sel && this.state(sel)?.img
    return img ? { a: { pa: img.pa, o: img.o }, f: { pa: img.pa, o: img.o + 1 } } : null
  }

  private patchImage(fn: (x: Inline & { t: 'img' }) => Inline) {
    const sel = this.imageSel()
    if (!sel) return
    this.patchObjects(sel, (x) => (x.t === 'img' ? fn(x) : x))
    this.pendingSel = sel
  }

  /** exact size in pt */
  setImageSize(w: number, h: number) {
    this.patchImage((x) => ({ ...x, w: +Math.max(4, w).toFixed(1), h: +Math.max(4, h).toFixed(1), resized: true }))
  }

  /** natural pixel size of the picture → pt, limited to the text width */
  imageOriginal(nw: number, nh: number) {
    const k = Math.min(1, this.contentW / Math.max(1, nw * 0.75))
    this.setImageSize(nw * 0.75 * k, nh * 0.75 * k)
  }

  setAltText(alt: string) {
    this.patchImage((x) => ({ ...x, alt: alt || undefined, altSet: true }))
  }

  async replaceImage(file: Blob) {
    const sel = this.imageSel()
    if (!sel) return
    const ext = (file.type.split('/')[1] ?? 'png').replace('svg+xml', 'svg').replace('jpeg', 'jpg')
    const { rid, path } = await this.client.call('docx.image', await file.arrayBuffer(), ext)
    const dims = await imageSize(file)
    this.input.restore(sel)
    this.patchImage((x) => {
      // keep the current width, follow the new picture's proportions
      const h = +(x.w * (dims.h / Math.max(1, dims.w))).toFixed(1)
      return { ...x, target: path, rid, repl: !!x.src, h, resized: true }
    })
  }

  /** deletes the selection (or the selected object) */
  deleteSelection() {
    const sel = this.input.sel()
    if (sel && !collapsed(sel)) opInsert(this, sel, '')
  }

  insertText(text: string) {
    const sel = this.input.sel()
    if (sel) opInsert(this, sel, text)
  }

  // ---------------------------------------------------------------- paint format

  /** copies the caret's character formatting; the next selection receives it */
  copyFormat() {
    if (this.painter) {
      this.painter = null
      this.onPainter?.(false)
      return
    }
    const st = this.state()
    if (!st) return
    this.painter = {
      b: st.b,
      i: st.i,
      u: st.u ? 'single' : 'none',
      strike: st.strike,
      color: st.color ?? null,
      hl: st.hl ?? null,
      va: st.va ?? null,
      sz: st.sz,
      font: st.font ? (/^\s*"?([^",]+)/.exec(st.font)?.[1] ?? st.font) : undefined,
    } as RunSet
    this.onPainter?.(true)
  }

  /** applies the copied formatting to the current selection (after the mouse is released) */
  applyPainter(): boolean {
    const set = this.painter
    const sel = this.input.readSel()
    if (!set || !sel || collapsed(sel)) return false
    this.painter = null
    this.onPainter?.(false)
    const s: Partial<RunStyle> = {}
    for (const [k, v] of Object.entries(set)) {
      if (k === 'font') (s as Record<string, unknown>).font = v ? cssFont(v as string) : undefined
      else if (k === 'u') (s as Record<string, unknown>).u = v
      else (s as Record<string, unknown>)[k] = v === null ? undefined : v
    }
    this.format(s, set)
    return true
  }

  // ---------------------------------------------------------------- toolbar state

  state(sel = this.input.sel()): FormatState | null {
    if (!sel) return null
    const [s] = ordered(this, sel)
    const p = this.para(s.pa)
    if (!p) return null
    const run = runAt(p.runs as RRun[], collapsed(sel) ? s.o : s.o + (plen(p.runs as RRun[]) > s.o ? 1 : 0)) as (Inline & { t: 'text' }) | undefined
    const rs: RunStyle = { ...(run?.s ?? {}), ...(collapsed(sel) ? this.pend?.s ?? {} : {}) }
    const [, e] = ordered(this, sel)
    let img: FormatState['img']
    if (s.pa === e.pa && e.o === s.o + 1) {
      const [, right] = splitRuns(p.runs as RRun[], s.o)
      const x = right[0] as Inline | undefined
      if (x?.t === 'img') img = { w: x.w, h: x.h, alt: x.alt, target: x.target, pa: s.pa, o: s.o }
    }
    return {
      line: p.pd?.line ?? (p.ps.lineRule === 'auto' || !p.ps.lineRule ? p.ps.line : undefined),
      va: rs.va,
      link: (run as { href?: string } | undefined)?.href,
      inTable: s.pa.includes('.'),
      img,
      b: !!rs.b,
      i: !!rs.i,
      u: !!rs.u && rs.u !== 'none',
      strike: !!rs.strike,
      color: rs.color,
      hl: rs.hl,
      sz: rs.sz ?? p.markSz ?? this.defaultSz,
      font: rs.font ?? p.markFont ?? this.defaultFont,
      align: p.ps.align ?? 'left',
      style: p.sid,
      lst: p.lst,
      canUndo: this.canUndo,
      canRedo: this.canRedo,
    }
  }

  private emitState(sel: Sel | null) {
    const st = sel ? this.state(sel) : null
    this.onState?.(st)
    this.onImage?.(st?.img ?? null)
  }

  onSelection(sel: Sel | null) {
    if (sel && this.pend && (sel.a.pa !== this.pend.at.pa || sel.a.o !== this.pend.at.o)) this.pend = null
    this.emitState(sel)
  }
}

type RichParaLike = Parameters<NonNullable<Parameters<typeof opInsertRuns>[4]>[number]>[0]

function imageSize(blob: Blob): Promise<{ w: number; h: number }> {
  return new Promise((resolve) => {
    const url = URL.createObjectURL(blob)
    const img = new Image()
    img.onload = () => {
      resolve({ w: img.naturalWidth || 300, h: img.naturalHeight || 200 })
      URL.revokeObjectURL(url)
    }
    img.onerror = () => {
      resolve({ w: 300, h: 200 })
      URL.revokeObjectURL(url)
    }
    img.src = url
  })
}
