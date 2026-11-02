import { useEffect, useRef, useState } from 'react'
import type { PresentationMeta } from '../../core/pptx/types'
import type { DocumentClient } from '../../core/worker/client'
import { SlideSlot } from './PptxView'

/** Full-screen slideshow: ←/→/Space/PageUp/PageDown/Home/End, click = next, Esc = exit. */
export function Presenter({ client, meta, start, onClose }: { client: DocumentClient; meta: PresentationMeta; start: number; onClose: () => void }) {
  const ref = useRef<HTMLDivElement>(null)
  const visible = meta.slides.filter((s) => !s.hidden).map((s) => s.index)
  const order = visible.length ? visible : meta.slides.map((s) => s.index)
  const [pos, setPos] = useState(Math.max(0, order.findIndex((i) => i >= start)))
  const [vw, setVw] = useState({ w: window.innerWidth, h: window.innerHeight })

  useEffect(() => {
    const el = ref.current
    el?.focus()
    el?.requestFullscreen?.().catch(() => {})
    const resize = () => setVw({ w: window.innerWidth, h: window.innerHeight })
    const exitFs = () => !document.fullscreenElement && onClose()
    window.addEventListener('resize', resize)
    document.addEventListener('fullscreenchange', exitFs)
    return () => {
      window.removeEventListener('resize', resize)
      document.removeEventListener('fullscreenchange', exitFs)
      if (document.fullscreenElement) document.exitFullscreen().catch(() => {})
    }
  }, [onClose])

  const go = (d: number) => setPos((p) => Math.max(0, Math.min(order.length - 1, p + d)))
  const scale = Math.min(vw.w / meta.w, vw.h / meta.h)
  return (
    <div
      ref={ref}
      className="de-present"
      tabIndex={-1}
      onClick={() => go(1)}
      onContextMenu={(e) => {
        e.preventDefault()
        go(-1)
      }}
      onKeyDown={(e) => {
        if (['ArrowRight', 'ArrowDown', 'PageDown', ' ', 'Enter', 'n'].includes(e.key)) go(1)
        else if (['ArrowLeft', 'ArrowUp', 'PageUp', 'Backspace', 'p'].includes(e.key)) go(-1)
        else if (e.key === 'Home') setPos(0)
        else if (e.key === 'End') setPos(order.length - 1)
        else if (e.key === 'Escape') onClose()
        else return
        e.preventDefault()
      }}
    >
      <div style={{ width: meta.w * scale, height: meta.h * scale }}>
        <SlideSlot key={order[pos]} client={client} index={order[pos]} meta={meta} scale={scale} live />
      </div>
      <div className="de-present-n">
        {pos + 1} / {order.length}
      </div>
    </div>
  )
}
