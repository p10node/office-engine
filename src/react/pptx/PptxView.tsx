import { useEffect, useRef, useState } from 'react'
import type { PresentationMeta, SlideModel } from '../../core/pptx/types'
import type { DocumentClient } from '../../core/worker/client'
import { useScroll, useSize } from '../common/hooks'
import { Status } from '../common/Status'
import { ZoomControl, type Zoom } from '../common/ZoomControl'
import { SlideView } from './Slide'

const caches = new WeakMap<DocumentClient, { models: Map<number, SlideModel>; pending: Map<number, Promise<SlideModel>> }>()

function cacheFor(client: DocumentClient) {
  let c = caches.get(client)
  if (!c) caches.set(client, (c = { models: new Map(), pending: new Map() }))
  return c
}

function useSlide(client: DocumentClient, index: number): SlideModel | { error: string } | undefined {
  const cache = cacheFor(client)
  const [model, setModel] = useState<SlideModel | { error: string } | undefined>(() => cache.models.get(index))
  useEffect(() => {
    const hit = cache.models.get(index)
    if (hit) return setModel(hit)
    let live = true
    let p = cache.pending.get(index)
    if (!p) {
      p = client.call('pptx.slide', index)
      cache.pending.set(index, p)
      p.then((m) => {
        cache.models.set(index, m)
        cache.pending.delete(index)
        if (cache.models.size > 120) cache.models.delete(cache.models.keys().next().value as number)
      }, () => cache.pending.delete(index))
    }
    p.then(
      (m) => live && setModel(m),
      (e) => live && setModel({ error: (e as Error).message }),
    )
    return () => {
      live = false
    }
  }, [client, cache, index])
  return model
}

function SlideSlot({ client, index, meta, scale, live }: { client: DocumentClient; index: number; meta: PresentationMeta; scale: number; live: boolean }) {
  const model = useSlide(client, index)
  if (!model) return <div className="de-slide" style={{ width: meta.w * scale, height: meta.h * scale, background: '#fff' }} />
  if ('error' in model) {
    return (
      <div className="de-slide de-ph" style={{ width: meta.w * scale, height: meta.h * scale, display: 'flex' }}>
        {model.error}
      </div>
    )
  }
  return <SlideView model={model} client={client} scale={scale} live={live} />
}

export interface PptxViewProps {
  client: DocumentClient
  onStatus?: (s: string) => void
  initialZoom?: Zoom
  thumbnails?: boolean
}

const PAD = 24
const GAP = 24
const THUMB_W = 188

export function PptxView({ client, onStatus, initialZoom = 'fit', thumbnails = true }: PptxViewProps) {
  const meta = client.info.pptx as PresentationMeta
  const [rootRef, rootSize] = useSize<HTMLDivElement>()
  const [scrollRef, size] = useSize<HTMLDivElement>()
  const scroll = useScroll(scrollRef)
  const thumbRef = useRef<HTMLDivElement>(null)
  const thumbScroll = useScroll(thumbRef)
  const [zoom, setZoom] = useState<Zoom>(initialZoom)
  const showThumbs = thumbnails && rootSize.w >= 720 && meta.count > 1

  const fitW = size.w ? (size.w - PAD * 2) / meta.w : 1
  const fitH = size.h ? (size.h - PAD * 2) / meta.h : 1
  const fit = Math.max(0.1, Math.min(fitW, Math.max(fitH, fitW * 0.6), 2))
  const scale = zoom === 'fit' ? fit : zoom
  const slideH = meta.h * scale
  const slideW = meta.w * scale
  const pitch = slideH + GAP
  const totalH = PAD * 2 + meta.count * pitch - GAP
  const totalW = Math.max(size.w, slideW + PAD * 2)
  const left = Math.max(PAD, (totalW - slideW) / 2)
  const first = Math.max(0, Math.floor((scroll.y - PAD) / pitch) - 1)
  const last = Math.min(meta.count - 1, Math.floor((scroll.y + size.h) / pitch) + 1)
  const current = Math.min(meta.count - 1, Math.max(0, Math.floor((scroll.y + size.h / 3 - PAD) / pitch)))

  // thumbnails
  const tScale = (THUMB_W - 46) / meta.w
  const tH = meta.h * tScale
  const tPitch = tH + 14
  const tFirst = Math.max(0, Math.floor(thumbScroll.y / tPitch) - 1)
  const tLast = Math.min(meta.count - 1, Math.floor((thumbScroll.y + size.h) / tPitch) + 1)


  // keep the active thumbnail in view
  useEffect(() => {
    const el = thumbRef.current
    if (!el) return
    const y = 10 + current * tPitch
    if (y < el.scrollTop || y + tH > el.scrollTop + el.clientHeight) el.scrollTop = y - el.clientHeight / 2 + tH / 2
  }, [current, tPitch, tH])

  const go = (i: number) => {
    const el = scrollRef.current
    if (el) el.scrollTop = PAD + Math.max(0, Math.min(meta.count - 1, i)) * pitch - Math.max(0, (size.h - slideH) / 2)
  }

  const onKeyDown = (e: React.KeyboardEvent) => {
    if (e.key === 'PageDown' || e.key === 'ArrowRight' || (e.key === ' ' && !e.shiftKey)) go(current + 1)
    else if (e.key === 'PageUp' || e.key === 'ArrowLeft') go(current - 1)
    else if (e.key === 'Home') go(0)
    else if (e.key === 'End') go(meta.count - 1)
    else return
    e.preventDefault()
  }

  const slots = []
  for (let i = first; i <= last; i++) {
    slots.push(
      <div key={i} className="de-slide-frame" style={{ top: PAD + i * pitch, left }} data-slide={i + 1}>
        <SlideSlot client={client} index={i} meta={meta} scale={scale} live />
      </div>,
    )
  }
  const thumbs = []
  if (showThumbs) {
    for (let i = tFirst; i <= tLast; i++) {
      thumbs.push(
        <div key={i} className="de-thumb" style={{ top: 10 + i * tPitch }} aria-current={i === current} onClick={() => go(i)}>
          <span className="de-thumb-n">{i + 1}</span>
          <SlideSlot client={client} index={i} meta={meta} scale={tScale} live />
        </div>,
      )
    }
  }

  return (
    <div ref={rootRef} style={{ display: 'flex', flex: 1, minWidth: 0 }}>
      {showThumbs && (
        <div ref={thumbRef} className="de-thumbs">
          <div style={{ height: 20 + meta.count * tPitch, position: 'relative' }}>{thumbs}</div>
        </div>
      )}
      <div style={{ position: 'relative', flex: 1, minWidth: 0, display: 'flex' }}>
        <div ref={scrollRef} className="de-scroll" tabIndex={0} onKeyDown={onKeyDown} style={{ outline: 'none' }}>
          <div style={{ position: 'relative', height: totalH, width: totalW }}>{slots}</div>
        </div>
        <ZoomControl zoom={zoom} effective={scale} onChange={setZoom} />
        <Status text={`Slide ${current + 1} of ${meta.count}`} onStatus={onStatus} />
      </div>
    </div>
  )
}
