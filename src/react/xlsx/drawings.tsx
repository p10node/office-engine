import { useEffect, useRef, useState, type CSSProperties, type ReactNode } from 'react'
import type { DocumentClient } from '../../core/worker/client'
import type { ChartModel } from '../../core/xlsx/chart'
import { renderChart } from '../../core/xlsx/chart-svg'
import type { SheetDrawing } from '../../core/xlsx/drawings'
import { MAX_COL, MAX_ROWS } from '../../core/xlsx/limits'
import { Icon } from '../edit/Toolbar'
import type { Strings } from '../strings'
import type { Axis } from './axis'

/** a sheet's drawing with its picture's object URL, where the browser can draw it */
export interface ShownDrawing extends SheetDrawing {
  url?: string
}

/** Edits re-read the charts this long after they rest: the values come from the edited cells. */
const REREAD_MS = 300

/**
 * The pictures and charts of the sheet (`xlsx.drawings`), read again when another sheet is shown and after edits
 * (`edits` counts them), so charts follow the cells; the charts' pictures of the list replaced are released.
 */
export function useSheetDrawings(client: DocumentClient, sheet: number, edits = 0): ShownDrawing[] {
  const [list, setList] = useState<ShownDrawing[]>([])
  const shown = useRef<ShownDrawing[]>([])
  const read = (live: () => boolean) =>
    client
      .call('xlsx.drawings', sheet)
      .then((drawings) => Promise.all(drawings.map(async (d) => ({ ...d, url: d.target ? await client.mediaUrl(d.target).catch(() => undefined) : undefined }))))
      .then((next) => {
        if (!live()) return
        releaseCharts(shown.current)
        shown.current = next
        setList(next)
      })
      .catch((e: Error) => console.warn('[office-engine] drawings not read:', e.message))
  useEffect(() => {
    let live = true
    setList([])
    void read(() => live)
    return () => {
      live = false
    }
  }, [client, sheet]) // eslint-disable-line react-hooks/exhaustive-deps
  useEffect(() => {
    if (!edits) return
    let live = true
    const timer = setTimeout(() => void read(() => live), REREAD_MS)
    return () => {
      live = false
      clearTimeout(timer)
    }
  }, [edits]) // eslint-disable-line react-hooks/exhaustive-deps
  useEffect(() => () => releaseCharts(shown.current), [])
  return list
}

/** each chart's picture, drawn for the size it was last shown at ('' when it could not be drawn at that size) */
const chartUrls = new WeakMap<ChartModel, { key: string; url: string }>()

/**
 * The chart drawn at `w` × `h` as an SVG picture's object URL (in an img: nothing in it runs), or '' when drawing it
 * failed: it is drawn in the grid's render, so a chart a crafted file breaks shows its placeholder box, not a blank view.
 */
function chartUrl(m: ChartModel, w: number, h: number): string {
  const key = `${Math.round(w)}x${Math.round(h)}`
  const had = chartUrls.get(m)
  if (had?.key === key) return had.url
  if (had?.url) URL.revokeObjectURL(had.url)
  let url = ''
  try {
    url = URL.createObjectURL(new Blob([renderChart(m, Math.round(w), Math.round(h))], { type: 'image/svg+xml' }))
  } catch (e) {
    console.warn('[office-engine] chart not drawn:', (e as Error).message)
  }
  chartUrls.set(m, { key, url })
  return url
}

function releaseCharts(list: ShownDrawing[]) {
  for (const d of list) {
    const had = d.model && chartUrls.get(d.model)
    if (!had) continue
    if (had.url) URL.revokeObjectURL(had.url)
    chartUrls.delete(d.model!)
  }
}

/**
 * Where a drawing lies on the sheet, px from its top left: anchored to cells it moves and grows with their rows and
 * columns (a hidden one is 0 wide); given by a size it keeps that size.
 */
export function drawingBox(d: SheetDrawing, cols: Axis, rows: Axis): { x: number; y: number; w: number; h: number } | null {
  let x: number
  let y: number
  if (d.from) {
    x = cols.pos(d.from.c) + d.from.dx
    y = rows.pos(d.from.r) + d.from.dy
  } else if (d.at) {
    x = d.at.x
    y = d.at.y
  } else return null
  const w = d.to ? cols.pos(d.to.c) + d.to.dx - x : (d.size?.w ?? 0)
  const h = d.to ? rows.pos(d.to.r) + d.to.dy - y : (d.size?.h ?? 0)
  return w >= 2 && h >= 2 ? { x, y, w, h } : null
}

/** How many rows and columns the sheet needs to reach every drawing (a drawing given by a size taken over default-sized cells) */
export function drawingExtent(list: SheetDrawing[], cols: Axis, rows: Axis): { rows: number; cols: number } {
  let r = 0
  let c = 0
  for (const d of list) {
    const end = d.to ?? (d.from && d.size ? { c: d.from.c + Math.ceil(d.size.w / cols.def), r: d.from.r + Math.ceil(d.size.h / rows.def) } : undefined)
    const fromAt = !end && d.at && d.size ? { c: Math.ceil((d.at.x + d.size.w) / cols.def), r: Math.ceil((d.at.y + d.size.h) / rows.def) } : undefined
    const e = end ?? fromAt
    if (e && Number.isFinite(e.r) && Number.isFinite(e.c)) {
      // Sizes/absolute positions are in pixels and can span many default-sized cells. Only this
      // OOXML drawing contribution is bounded; the sheet's CSV row count and Axis stay unrestricted.
      r = Math.max(r, Math.min(MAX_ROWS, e.r + 1))
      c = Math.max(c, Math.min(MAX_COL, e.c + 1))
    }
  }
  return { rows: r, cols: c }
}

/**
 * A drawing as shown: the picture itself, a chart drawn at its size, or a box of its size saying what is there (a chart
 * of a kind not drawn or that could not be drawn, a picture the browser cannot draw).
 */
export function drawingEl(key: string, d: ShownDrawing, style: CSSProperties & { width: number; height: number }, strings: Strings): ReactNode {
  if (d.url) return <img key={key} className="de-xl-pic" src={d.url} alt={d.title} style={style} draggable={false} />
  const t = strings.xlsx.drawings
  const drawn = d.model && chartUrl(d.model, style.width, style.height)
  if (drawn) return <img key={key} className="de-xl-pic de-xl-chart" src={drawn} alt={d.title || t.chart} style={style} draggable={false} />
  const chart = d.chart ?? ''
  const what = d.kind === 'chart' ? (Object.hasOwn(t.charts, chart) ? t.charts[chart as keyof typeof t.charts] : t.chart) : t.picture
  return (
    <div key={key} className="de-xl-drawing" style={style} title={[what, d.title].filter(Boolean).join(': ')}>
      <Icon name={d.kind === 'chart' ? 'chart' : 'image'} size={24} />
      <strong>{what}</strong>
      {d.title && <span>{d.title}</span>}
    </div>
  )
}
