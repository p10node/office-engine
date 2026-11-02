import { useEffect, useMemo, useState, type CSSProperties, type ReactNode } from 'react'
import type { CellDetail, CellOut, Merge, SheetProgress, WorkbookMeta } from '../../core/xlsx/types'
import { colName } from '../../core/xlsx/workbook'
import type { DocumentClient } from '../../core/worker/client'
import { useForceUpdate, useScroll, useSize } from '../common/hooks'
import { Status } from '../common/Status'
import { Axis } from './axis'

const BLOCK = 128
const HH = 22
const MAX_SCROLL_PX = 12_000_000

interface CachedBlock {
  rows: Map<number, CellOut[]>
  complete: boolean
  rowsAt: number
}

class SheetState {
  rows = 0
  cols = 0
  done = false
  error?: string
  header = false
  rowAxis = new Axis(20)
  colAxis = new Axis(64)
  merges: Merge[] = []
  frozen = { rows: 0, cols: 0 }
  gridLines = true
  blocks = new Map<number, CachedBlock>()
  inflight = new Set<number>()
}

class XlsxStore {
  sheets = new Map<number, SheetState>()
  private listeners = new Set<() => void>()
  private off: () => void

  constructor(private client: DocumentClient) {
    this.off = client.on('xlsx:progress', (p) => this.onProgress(p))
  }

  subscribe(fn: () => void) {
    this.listeners.add(fn)
    return () => {
      this.listeners.delete(fn)
    }
  }

  private emit() {
    for (const fn of this.listeners) fn()
  }

  sheet(i: number): SheetState {
    let s = this.sheets.get(i)
    if (!s) {
      s = new SheetState()
      this.sheets.set(i, s)
      this.client.call('xlsx.load', i).catch((e) => {
        s!.error = (e as Error).message
        s!.done = true
        this.emit()
      })
    }
    return s
  }

  private onProgress(p: SheetProgress) {
    const s = this.sheet(p.sheet)
    if (p.defaultRowH !== undefined && !s.header) {
      s.header = true
      s.rowAxis.def = p.defaultRowH || 20
      s.colAxis.def = p.defaultColW || 64
      for (const ci of p.colInfo ?? []) s.colAxis.setRange(ci.min, ci.max, ci.w)
      s.frozen = p.frozen ?? s.frozen
      s.gridLines = p.gridLines ?? true
    }
    for (const [r, px] of p.rowSizes) s.rowAxis.set(r, px)
    s.rows = Math.max(s.rows, p.rows)
    s.cols = Math.max(s.cols, p.cols)
    if (p.merges) s.merges = p.merges
    if (p.error) s.error = p.error
    s.done = p.done
    this.emit()
  }

  ensure(sheet: number, r0: number, r1: number) {
    const s = this.sheet(sheet)
    if (!s.header && !s.done) return
    for (let b = Math.floor(r0 / BLOCK); b <= Math.floor(r1 / BLOCK); b++) {
      const cached = s.blocks.get(b)
      if (s.inflight.has(b)) continue
      if (cached && (cached.complete || (cached.rowsAt === s.rows && !s.done))) continue
      if (b * BLOCK > s.rows && !s.done && cached) continue
      s.inflight.add(b)
      const rowsAt = s.rows
      this.client
        .call('xlsx.block', sheet, b * BLOCK, (b + 1) * BLOCK)
        .then((res) => {
          const rows = new Map<number, CellOut[]>()
          for (const r of res.rows) rows.set(r.r, r.cells)
          s.blocks.set(b, { rows, complete: res.complete, rowsAt })
          if (s.blocks.size > 400) {
            const firstKey = s.blocks.keys().next().value as number
            s.blocks.delete(firstKey)
          }
        })
        .catch(() => {})
        .finally(() => {
          s.inflight.delete(b)
          this.emit()
        })
    }
  }

  row(sheet: number, r: number): CellOut[] | undefined {
    return this.sheets.get(sheet)?.blocks.get(Math.floor(r / BLOCK))?.rows.get(r)
  }

  dispose() {
    this.off()
    this.listeners.clear()
  }
}

const stores = new WeakMap<DocumentClient, XlsxStore>()
function storeFor(client: DocumentClient) {
  let s = stores.get(client)
  if (!s) stores.set(client, (s = new XlsxStore(client)))
  return s
}

export interface XlsxViewProps {
  client: DocumentClient
  onStatus?: (s: string) => void
  initialSheet?: number
}

export function XlsxView({ client, onStatus, initialSheet }: XlsxViewProps) {
  const meta = client.info.xlsx as WorkbookMeta
  const visible = meta.sheets.filter((s) => !s.hidden)
  const [active, setActive] = useState(initialSheet ?? visible[0]?.index ?? 0)
  const store = useMemo(() => storeFor(client), [client])
  const [, bump] = useForceUpdate()
  useEffect(() => store.subscribe(bump), [store, bump])
  const state = store.sheet(active)

  const status = state.error
    ? `Error: ${state.error}`
    : `${state.rows.toLocaleString()} rows × ${state.cols.toLocaleString()} cols${state.done ? '' : ' · loading…'}`

  return (
    <div className="de-xlsx">
      <Status text={status} onStatus={onStatus} />
      <SheetGrid key={active} store={store} sheet={active} meta={meta} client={client} />
      <div className="de-tabs" role="tablist">
        {visible.map((s) => (
          <button key={s.index} type="button" role="tab" className="de-tab" aria-selected={s.index === active} onClick={() => setActive(s.index)}>
            {s.name}
          </button>
        ))}
      </div>
    </div>
  )
}

interface GridProps {
  store: XlsxStore
  sheet: number
  meta: WorkbookMeta
  client: DocumentClient
}

const fontPx = (v: string | number | undefined, base: number) =>
  typeof v === 'string' && v.endsWith('pt') ? (parseFloat(v) * 4) / 3 : typeof v === 'number' ? v : base

/** Excel-like: shrink General decimals to fit, otherwise show ### when a number is too wide. */
function fitNumber(text: string, width: number, px: number): string {
  const cw = px * 0.56
  if (text.length * cw <= width) return text
  const m = /^(-?\d+)\.(\d+)$/.exec(text)
  if (m && (m[1].length + (text.startsWith('-') ? 0 : 0)) * cw <= width) {
    const keep = Math.max(0, Math.floor(width / cw) - m[1].length - 1)
    const v = Number(text).toFixed(keep)
    if (v.length * cw <= width) return v
  }
  return '#'.repeat(Math.max(1, Math.floor(width / cw)))
}

const JUSTIFY: Record<string, string> = { center: 'center', right: 'flex-end', left: 'flex-start', justify: 'flex-start' }

function SheetGrid({ store, sheet, meta, client }: GridProps) {
  const s = store.sheet(sheet)
  const [scrollRef, size] = useSize<HTMLDivElement>()
  const scroll = useScroll(scrollRef)
  const [sel, setSel] = useState<{ r: number; c: number }>({ r: 0, c: 0 })
  const [detail, setDetail] = useState<CellDetail | undefined>()

  const { rowAxis, colAxis, frozen } = s
  const HW = Math.max(40, 12 + 7 * String(Math.max(s.rows, 1000)).length)
  const viewW = Math.max(0, size.w - HW)
  const viewH = Math.max(0, size.h - HH)
  rowAxis.count = Math.max(s.rows + (s.done ? 20 : 1), Math.ceil(viewH / rowAxis.def) + 1, frozen.rows + 1)
  colAxis.count = Math.max(s.cols + 3, 26, Math.ceil(viewW / Math.max(1, colAxis.def)) + 1, frozen.cols + 1)
  const frozenW = colAxis.pos(frozen.cols)
  const frozenH = rowAxis.pos(frozen.rows)
  const totalW = colAxis.total()
  const totalH = rowAxis.total()

  // virtual scroll with clamped spacer height (browsers cap element size)
  const spacerW = HW + totalW
  const spacerH = Math.min(HH + totalH, MAX_SCROLL_PX)
  const maxLogicalY = Math.max(0, HH + totalH - size.h)
  const maxScrollY = Math.max(0, spacerH - size.h)
  const scrollY = maxScrollY > 0 ? Math.min(maxLogicalY, (scroll.y / maxScrollY) * maxLogicalY) : 0
  const scrollX = scroll.x

  const mainOx = frozenW + scrollX
  const mainOy = frozenH + scrollY
  const mainW = Math.max(0, viewW - frozenW)
  const mainH = Math.max(0, viewH - frozenH)
  const r0 = rowAxis.index(mainOy)
  const r1 = Math.min(rowAxis.count - 1, rowAxis.index(mainOy + mainH))
  const c0 = colAxis.index(mainOx)
  const c1 = Math.min(colAxis.count - 1, colAxis.index(mainOx + mainW))

  // merges touching the visible area (+ frozen panes)
  const visMerges = useMemo(() => {
    const inRange = (m: Merge) =>
      (m.r1 >= r0 && m.r0 <= r1) || m.r0 < frozen.rows ? (m.c1 >= c0 && m.c0 <= c1) || m.c0 < frozen.cols : false
    return s.merges.filter(inRange)
  }, [s.merges, r0, r1, c0, c1, frozen.rows, frozen.cols])

  useEffect(() => {
    if (frozen.rows) store.ensure(sheet, 0, frozen.rows - 1)
    store.ensure(sheet, Math.max(0, r0 - 20), r1 + 40)
    for (const m of visMerges) if (m.r0 < r0) store.ensure(sheet, m.r0, m.r0)
  })

  useEffect(() => {
    let live = true
    client.call('xlsx.cell', sheet, sel.r, sel.c).then((d) => live && setDetail(d), () => {})
    return () => {
      live = false
    }
  }, [client, sheet, sel.r, sel.c, s.done])

  const styles = meta.styles
  const basePx = fontPx(meta.baseFont.fontSize, 14.67)
  const mergeAt = (r: number, c: number) => visMerges.find((m) => r >= m.r0 && r <= m.r1 && c >= m.c0 && c <= m.c1)

  function region(key: string, rr0: number, rr1: number, cc0: number, cc1: number, ox: number, oy: number, left: number, top: number, w: number, h: number) {
    if (w <= 0 || h <= 0 || rr1 < rr0 || cc1 < cc0) return null
    const out: ReactNode[] = []
    if (s.gridLines) {
      for (let r = rr0; r <= rr1; r++) {
        const y = rowAxis.pos(r + 1) - oy - 1
        out.push(<div key={'h' + r} className="de-gl-h" style={{ top: y }} />)
      }
      for (let c = cc0; c <= cc1; c++) {
        const x = colAxis.pos(c + 1) - ox - 1
        out.push(<div key={'v' + c} className="de-gl-v" style={{ left: x }} />)
      }
    }
    const covered = new Set<number>()
    const anchors: Merge[] = []
    for (const m of visMerges) {
      if (m.r1 < rr0 || m.r0 > rr1 || m.c1 < cc0 || m.c0 > cc1) continue
      anchors.push(m)
      for (let r = Math.max(m.r0, rr0); r <= Math.min(m.r1, rr1); r++)
        for (let c = Math.max(m.c0, cc0); c <= Math.min(m.c1, cc1); c++) covered.add(r * 16384 + c)
    }
    const cellEl = (r: number, cell: CellOut, wOverride?: number, hOverride?: number, filled?: Set<number>, merged = false) => {
      const [c, raw, si, kind, color] = cell
      const cw = wOverride ?? colAxis.size(c)
      const ch = hOverride ?? rowAxis.size(r)
      if (!cw || !ch) return null
      const css = (styles[si] ?? {}) as Record<string, string | number>
      let text = raw
      if (kind === 1 && raw) text = fitNumber(raw, cw - 6, fontPx(css.fontSize, basePx))
      const { textAlign, textRotate: _rot, alignItems, whiteSpace, ...rest } = css
      const justify = (textAlign && JUSTIFY[textAlign as string]) || (kind === 1 ? 'flex-end' : kind === 2 ? 'center' : 'flex-start')
      let ov = 0
      if (text && !whiteSpace && justify === 'flex-start' && !merged && filled) {
        // let text overflow into following empty cells (Excel behaviour)
        let next = c + 1
        while (next <= cc1 + 8 && !filled.has(next) && !covered.has(r * 16384 + next)) next++
        ov = colAxis.pos(next) - colAxis.pos(c)
      }
      const style: CSSProperties = {
        ...(rest as CSSProperties),
        left: colAxis.pos(c) - ox,
        top: rowAxis.pos(r) - oy,
        width: cw,
        height: ch,
        justifyContent: justify,
        alignItems: (alignItems as string) ?? 'flex-end',
        textAlign: (textAlign as CSSProperties['textAlign']) ?? undefined,
      }
      if (color) style.color = color
      if (whiteSpace) style.whiteSpace = whiteSpace as CSSProperties['whiteSpace']
      const overflow = ov > cw
      return (
        <div key={`${r}:${c}`} className={overflow ? 'de-cell de-ov' : 'de-cell'} style={style}>
          {text && <span style={overflow ? { maxWidth: ov - 6, overflow: 'hidden', flex: 'none' } : undefined}>{text}</span>}
        </div>
      )
    }
    for (const m of anchors) {
      const row = store.row(sheet, m.r0)
      const cell = row?.find((x) => x[0] === m.c0) ?? ([m.c0, '', 0, 0] as CellOut)
      out.push(cellEl(m.r0, cell, colAxis.pos(m.c1 + 1) - colAxis.pos(m.c0), rowAxis.pos(m.r1 + 1) - rowAxis.pos(m.r0), undefined, true))
    }
    for (let r = rr0; r <= rr1; r++) {
      const row = store.row(sheet, r)
      if (!row) continue
      let filled: Set<number> | undefined
      for (const cell of row) {
        const c = cell[0]
        if (c < cc0 - 8 || c > cc1) continue
        if (covered.has(r * 16384 + c)) continue
        if (c < cc0 && (!cell[1] || cell[3] !== 0)) continue
        if (anchors.some((m) => m.r0 === r && m.c0 === c)) continue
        if (!filled && cell[3] === 0 && cell[1]) filled = new Set(row.filter((x) => x[1] !== '').map((x) => x[0]))
        out.push(cellEl(r, cell, undefined, undefined, filled))
      }
    }
    // selection
    if (sel.r >= rr0 - 1 && sel.r <= rr1 && sel.c >= cc0 - 1 && sel.c <= cc1) {
      const m = mergeAt(sel.r, sel.c)
      const sr = m ? m.r0 : sel.r
      const sc = m ? m.c0 : sel.c
      const sw = m ? colAxis.pos(m.c1 + 1) - colAxis.pos(m.c0) : colAxis.size(sel.c)
      const sh = m ? rowAxis.pos(m.r1 + 1) - rowAxis.pos(m.r0) : rowAxis.size(sel.r)
      out.push(<div key="sel" className="de-sel" style={{ left: colAxis.pos(sc) - ox - 1, top: rowAxis.pos(sr) - oy - 1, width: sw + 1, height: sh + 1 }} />)
    }
    return (
      <div key={key} className="de-region" style={{ left, top, width: w, height: h }}>
        {out}
      </div>
    )
  }

  function colHeaders(key: string, cc0: number, cc1: number, ox: number, left: number, w: number) {
    if (w <= 0) return null
    const out: ReactNode[] = []
    for (let c = cc0; c <= cc1; c++) {
      const cw = colAxis.size(c)
      if (!cw) continue
      out.push(
        <div key={c} className={c === sel.c ? 'de-head-c de-head-on' : 'de-head-c'} style={{ left: colAxis.pos(c) - ox, top: 0, width: cw, height: HH }}>
          {colName(c)}
        </div>,
      )
    }
    return (
      <div key={key} className="de-region" style={{ left, top: 0, width: w, height: HH }}>
        {out}
      </div>
    )
  }

  function rowHeaders(key: string, rr0: number, rr1: number, oy: number, top: number, h: number) {
    if (h <= 0) return null
    const out: ReactNode[] = []
    for (let r = rr0; r <= rr1; r++) {
      const rh = rowAxis.size(r)
      if (!rh) continue
      out.push(
        <div key={r} className={r === sel.r ? 'de-head-r de-head-on' : 'de-head-r'} style={{ left: 0, top: rowAxis.pos(r) - oy, width: HW, height: rh }}>
          {r + 1}
        </div>,
      )
    }
    return (
      <div key={key} className="de-region" style={{ left: 0, top, width: HW, height: h }}>
        {out}
      </div>
    )
  }

  const fr = frozen.rows
  const fc = frozen.cols
  const fw = Math.min(frozenW, viewW)
  const fh = Math.min(frozenH, viewH)

  const scrollIntoView = (r: number, c: number) => {
    const el = scrollRef.current
    if (!el) return
    const x = colAxis.pos(c)
    const y = rowAxis.pos(r)
    if (c >= fc) {
      if (x - frozenW < scrollX) el.scrollLeft = x - frozenW
      else if (x + colAxis.size(c) > mainOx + mainW) el.scrollLeft = x + colAxis.size(c) - frozenW - mainW
    }
    if (r >= fr && maxLogicalY > 0) {
      let ty = scrollY
      if (y - frozenH < scrollY) ty = y - frozenH
      else if (y + rowAxis.size(r) > mainOy + mainH) ty = y + rowAxis.size(r) - frozenH - mainH
      if (ty !== scrollY) el.scrollTop = (ty / maxLogicalY) * maxScrollY
    }
  }

  const onMouseDown = (e: React.MouseEvent) => {
    const rect = (e.currentTarget as HTMLElement).getBoundingClientRect()
    const x = e.clientX - rect.left - HW
    const y = e.clientY - rect.top - HH
    if (x < 0 && y < 0) return
    const lx = x < frozenW ? x : x + scrollX
    const ly = y < frozenH ? y : y + scrollY
    const c = x < 0 ? sel.c : colAxis.index(Math.max(0, lx))
    const r = y < 0 ? sel.r : rowAxis.index(Math.max(0, ly))
    setSel({ r, c })
    scrollRef.current?.focus({ preventScroll: true })
  }

  const onKeyDown = (e: React.KeyboardEvent) => {
    const d: Record<string, [number, number]> = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1], Enter: [1, 0], Tab: [0, 1] }
    let mv = d[e.key]
    if (e.key === 'PageDown') mv = [Math.max(1, Math.floor(mainH / rowAxis.def) - 1), 0]
    if (e.key === 'PageUp') mv = [-Math.max(1, Math.floor(mainH / rowAxis.def) - 1), 0]
    if (!mv) return
    e.preventDefault()
    const m = mergeAt(sel.r, sel.c)
    let r = sel.r + mv[0]
    let c = sel.c + mv[1]
    if (m && mv[0] > 0) r = m.r1 + mv[0]
    if (m && mv[1] > 0) c = m.c1 + 1
    r = Math.max(0, Math.min(rowAxis.count - 1, r))
    c = Math.max(0, Math.min(colAxis.count - 1, c))
    while (r > 0 && rowAxis.size(r) === 0 && mv[0] < 0) r--
    while (rowAxis.size(r) === 0 && r < rowAxis.count - 1) r++
    while (c > 0 && colAxis.size(c) === 0 && mv[1] < 0) c--
    while (colAxis.size(c) === 0 && c < colAxis.count - 1) c++
    setSel({ r, c })
    scrollIntoView(r, c)
  }

  const fx = detail?.formula ?? detail?.value ?? ''
  return (
    <>
      <div className="de-fx">
        <div className="de-fx-ref">
          {colName(sel.c)}
          {sel.r + 1}
        </div>
        <div className="de-fx-val" title={fx}>
          {fx}
        </div>
      </div>
      <div ref={scrollRef} className="de-grid" tabIndex={0} onKeyDown={onKeyDown}>
        <div style={{ position: 'relative', width: Math.max(spacerW, size.w), height: Math.max(spacerH, size.h) }}>
        <div className="de-grid-view" style={{ width: size.w, height: size.h, ...(meta.baseFont as CSSProperties) }} onMouseDown={onMouseDown}>
          {region('main', r0, r1, c0, c1, mainOx, mainOy, HW + fw, HH + fh, mainW, mainH)}
          {fr > 0 && region('top', 0, fr - 1, c0, c1, mainOx, 0, HW + fw, HH, mainW, fh)}
          {fc > 0 && region('left', r0, r1, 0, fc - 1, 0, mainOy, HW, HH + fh, fw, mainH)}
          {fr > 0 && fc > 0 && region('tl', 0, fr - 1, 0, fc - 1, 0, 0, HW, HH, fw, fh)}
          {colHeaders('ch', c0, c1, mainOx, HW + fw, mainW)}
          {fc > 0 && colHeaders('chf', 0, fc - 1, 0, HW, fw)}
          {rowHeaders('rh', r0, r1, mainOy, HH + fh, mainH)}
          {fr > 0 && rowHeaders('rhf', 0, fr - 1, 0, HH, fh)}
          {fr > 0 && <div className="de-freeze-h" style={{ left: 0, top: HH + fh - 1, width: size.w, height: 2 }} />}
          {fc > 0 && <div className="de-freeze-v" style={{ left: HW + fw - 1, top: 0, width: 2, height: size.h }} />}
          <div className="de-corner" style={{ width: HW, height: HH }} />
        </div>
        </div>
      </div>
    </>
  )
}
