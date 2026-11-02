import { useEffect, useMemo, useRef, useState, type CSSProperties } from 'react'
import type { Block, DocxMeta, PageSetup } from '../../core/docx/types'
import { DEFAULT_PAGE } from '../../core/docx/types'
import type { DocumentClient } from '../../core/worker/client'
import { lowerIndex, useForceUpdate, useIsoLayoutEffect, useScroll, useSize } from '../common/hooks'
import { Status } from '../common/Status'
import { ZoomControl, type Zoom } from '../common/ZoomControl'
import { BlockView, type RenderCtx } from './render'

const PX = 4 / 3
const GAP = 16
const PAD = 24

interface PageSlice {
  start: number
  end: number
  used: number
}

function paginate(blocks: Block[], heights: number[], count: number, avail: number): PageSlice[] {
  const pages: PageSlice[] = []
  let cur: PageSlice = { start: 0, end: 0, used: 0 }
  for (let i = 0; i < count; i++) {
    const b = blocks[i]
    const h = heights[i] ?? 0
    if (b.t === 'sect') {
      cur.end = i + 1
      if (b.kind !== 'continuous') {
        pages.push(cur)
        cur = { start: i + 1, end: i + 1, used: 0 }
      }
      continue
    }
    const forced = b.t === 'p' && b.pb && cur.end > cur.start
    if (forced || (cur.used > 0 && cur.used + h > avail)) {
      pages.push(cur)
      cur = { start: i, end: i, used: 0 }
    }
    cur.end = i + 1
    cur.used += h
  }
  if (cur.end > cur.start || !pages.length) pages.push(cur)
  return pages
}

const started = new WeakSet<DocumentClient>()

export interface DocxViewProps {
  client: DocumentClient
  onStatus?: (s: string) => void
  initialZoom?: Zoom
}

export function DocxView({ client, onStatus, initialZoom = 'fit' }: DocxViewProps) {
  const blocks = useRef<Block[]>([])
  const [version, bump] = useForceUpdate()
  const [meta, setMeta] = useState<DocxMeta | null>(null)
  const [done, setDone] = useState<{ error?: string } | null>(null)
  const [zoom, setZoom] = useState<Zoom>(initialZoom)
  const [scrollRef, size] = useSize<HTMLDivElement>()
  const scroll = useScroll(scrollRef)

  useEffect(() => {
    const offs = [
      client.on('docx:blocks', (b) => {
        for (const x of b) blocks.current.push(x)
        bump()
      }),
      client.on('docx:meta', setMeta),
      client.on('docx:done', setDone),
    ]
    if (!started.has(client)) {
      started.add(client)
      client.call('docx.start').catch(() => {})
    }
    return () => offs.forEach((f) => f())
  }, [client, bump])

  const page: PageSetup = meta?.page ?? DEFAULT_PAGE
  const contentW = page.w - page.ml - page.mr
  const availPx = (page.h - page.mt - page.mb) * PX

  // ---- measurement (offscreen, incremental, time-sliced) ----
  const heights = useRef<number[]>([])
  const [measured, setMeasured] = useState(0)
  const [batch, setBatch] = useState(40)
  const measureRef = useRef<HTMLDivElement>(null)
  const widthKey = `${contentW}|${meta?.font}|${meta?.sz}`
  const lastKey = useRef(widthKey)
  if (lastKey.current !== widthKey) {
    lastKey.current = widthKey
    heights.current = []
    if (measured !== 0) setMeasured(0)
  }
  const total = blocks.current.length
  const mEnd = Math.min(total, measured + batch)
  const renderStart = performance.now()
  const measureCtx = useMemo<RenderCtx>(() => ({ client, page: -1, pages: 0, measuring: true }), [client])

  useIsoLayoutEffect(() => {
    const el = measureRef.current
    if (!el || mEnd <= measured || !meta) return
    const hs: number[] = []
    for (const child of Array.from(el.children)) hs.push((child as HTMLElement).getBoundingClientRect().height)
    const dt = performance.now() - renderStart
    const id = setTimeout(() => {
      for (let i = 0; i < hs.length; i++) heights.current[measured + i] = hs[i]
      setMeasured(measured + hs.length)
      setBatch((b) => (dt < 10 ? Math.min(b * 2, 1500) : dt > 30 ? Math.max(10, b >> 1) : b))
    }, 0)
    return () => clearTimeout(id)
  })

  const pages = useMemo(
    () => paginate(blocks.current, heights.current, measured, availPx),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [measured, availPx, version],
  )

  // ---- layout ----
  const pageW = page.w * PX
  const pageH = page.h * PX
  const fit = size.w ? Math.max(0.2, Math.min(2, (size.w - PAD * 2) / pageW)) : 1
  const z = zoom === 'fit' ? Math.min(fit, 1.25) : zoom
  const complete = !!done && measured >= total
  const pageHeights = pages.map((p) => Math.max(pageH, (page.mt + page.mb) * PX + p.used) * z)
  const tops: number[] = []
  let acc = PAD
  for (const h of pageHeights) {
    tops.push(acc)
    acc += h + GAP
  }
  const skeleton = complete ? 0 : 1
  const totalH = acc + skeleton * (pageH * z + GAP) + PAD - GAP
  const totalW = Math.max(size.w, pageW * z + PAD * 2)
  const left = Math.max(PAD, (totalW - pageW * z) / 2)

  const first = tops.length ? lowerIndex(tops, scroll.y - pageH * z) : 0
  const last = tops.length ? lowerIndex(tops, scroll.y + size.h + 200) : -1
  const pageCount = complete ? pages.length : 0

  const cur = tops.length ? lowerIndex(tops, scroll.y + size.h / 3) + 1 : 0
  const status = done?.error ? `Error: ${done.error}` : complete ? `Page ${cur} of ${pages.length}` : `Page ${cur} of ${pages.length}+ · loading…`

  const docStyle: CSSProperties = { fontFamily: meta?.font, fontSize: meta ? `${meta.sz}pt` : undefined }

  return (
    <div style={{ position: 'relative', flex: 1, minWidth: 0, display: 'flex' }}>
    <div ref={scrollRef} className="de-scroll">
      {meta && mEnd > measured && (
        <div ref={measureRef} className="de-measure de-doc" style={{ ...docStyle, width: `${contentW}pt` }}>
          {blocks.current.slice(measured, mEnd).map((b, i) => (
            <BlockView key={measured + i} block={b} ctx={measureCtx} />
          ))}
        </div>
      )}
      <div className="de-pages" style={{ height: totalH, width: totalW }}>
        {pages.map((p, i) =>
          i < first || i > last ? null : (
            <PageView
              key={i}
              index={i}
              count={pageCount}
              slice={p}
              blocks={blocks.current}
              meta={meta}
              page={page}
              top={tops[i]}
              left={left}
              height={pageHeights[i]}
              zoom={z}
              client={client}
              docStyle={docStyle}
            />
          ),
        )}
        {skeleton > 0 && (
          <div className="de-page-skeleton" style={{ top: acc, left, width: pageW * z, height: pageH * z }} />
        )}
      </div>
    </div>
      <ZoomControl zoom={zoom} effective={z} onChange={setZoom} />
      <Status text={status} onStatus={onStatus} />
    </div>
  )
}

interface PageViewProps {
  index: number
  count: number
  slice: PageSlice
  blocks: Block[]
  meta: DocxMeta | null
  page: PageSetup
  top: number
  left: number
  height: number
  zoom: number
  client: DocumentClient
  docStyle: CSSProperties
}

function PageView({ index, count, slice, blocks, meta, page, top, left, height, zoom, client, docStyle }: PageViewProps) {
  const ctx = useMemo<RenderCtx>(() => ({ client, page: index, pages: count, measuring: false }), [client, index, count])
  const first = index === 0 && meta?.titlePg
  const header = first ? meta?.firstHeader : meta?.header
  const footer = first ? meta?.firstFooter : meta?.footer
  const items: Block[] = []
  for (let i = slice.start; i < slice.end; i++) items.push(blocks[i])
  return (
    <div className="de-page" style={{ top, left, width: page.w * PX * zoom, height }} data-page={index + 1}>
      <div className="de-page-inner de-doc" style={{ ...docStyle, width: `${page.w}pt`, height: height / zoom, transform: `scale(${zoom})` }}>
        {header && (
          <div className="de-hf" style={{ top: `${page.header}pt`, padding: `0 ${page.mr}pt 0 ${page.ml}pt` }}>
            {header.map((b, i) => (
              <BlockView key={i} block={b} ctx={ctx} />
            ))}
          </div>
        )}
        <div style={{ position: 'absolute', top: `${page.mt}pt`, left: `${page.ml}pt`, width: `${page.w - page.ml - page.mr}pt` }}>
          {items.map((b, i) => (
            <BlockView key={slice.start + i} block={b} ctx={ctx} />
          ))}
        </div>
        {footer && (
          <div className="de-hf" style={{ bottom: `${page.footer}pt`, padding: `0 ${page.mr}pt 0 ${page.ml}pt` }}>
            {footer.map((b, i) => (
              <BlockView key={i} block={b} ctx={ctx} />
            ))}
          </div>
        )}
      </div>
    </div>
  )
}
