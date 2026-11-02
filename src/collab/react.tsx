/** React bindings: session lifecycle, presence avatars / status, collaborators' cursors. */
import { useEffect, useMemo, useRef, useState, useSyncExternalStore, type CSSProperties } from 'react'
import type { PresentationMeta } from '../core/pptx/types'
import type { DocumentClient } from '../core/worker/client'
import { getPeers, subscribePeers, type PeerView } from './hosts'
import { startCollab, type Collab, type CollabOptions } from './index'
import type { CollabStatus } from './session'

const transportIds = new WeakMap<object, number>()
let nextTransportId = 0

function optsKey(o: CollabOptions | undefined): string {
  if (!o) return ''
  let t: string
  if (typeof o.transport === 'string') t = o.transport
  else {
    let id = transportIds.get(o.transport)
    if (id === undefined) transportIds.set(o.transport, (id = ++nextTransportId))
    t = `#${id}`
  }
  return `${t}|${o.room}|${o.base ?? 0}|${o.user.id ?? ''}|${o.user.name}|${o.user.color ?? ''}`
}

export interface CollabState {
  collab: Collab | null
  status: CollabStatus | null
  error?: string
}

/** runs a collaboration session for `client` while `opts` is set */
export function useCollab(client: DocumentClient | null | undefined, opts: CollabOptions | undefined): CollabState {
  const [state, setState] = useState<CollabState>({ collab: null, status: null })
  const key = optsKey(opts)
  const optsRef = useRef(opts)
  optsRef.current = opts
  useEffect(() => {
    const o = optsRef.current
    if (!client || !o) {
      setState({ collab: null, status: null })
      return
    }
    let c: Collab
    try {
      c = startCollab(client, o)
    } catch (e) {
      setState({ collab: null, status: 'error', error: (e as Error).message })
      return
    }
    const sync = () => setState({ collab: c, status: c.status, error: c.session.error })
    sync()
    const offs = [c.on('status', sync), c.on('error', sync)]
    return () => {
      offs.forEach((f) => f())
      c.close()
    }
  }, [client, key])
  return state
}

const initials = (name: string) =>
  name
    .split(/\s+/)
    .filter(Boolean)
    .slice(0, 2)
    .map((w) => w[0]!.toUpperCase())
    .join('') || '?'

const LABEL: Record<CollabStatus, string> = {
  connecting: 'Connecting…',
  syncing: 'Syncing…',
  live: 'Live',
  offline: 'Offline · changes will sync',
  error: 'Not synced',
}

/** status chip + collaborators' avatars (top bar) */
export function CollabBar({ client, state }: { client: DocumentClient; state: CollabState }) {
  const peers = useSyncExternalStore(
    (fn) => subscribePeers(client, fn),
    () => getPeers(client),
  )
  const [pending, setPending] = useState(0)
  useEffect(() => {
    const s = state.collab?.session
    if (!s) return
    const t = setInterval(() => setPending(s.pending), 500)
    return () => clearInterval(t)
  }, [state.collab])
  if (!state.status) return null
  const title = state.status === 'error' ? state.error ?? 'Collaboration stopped' : state.status === 'live' && pending ? `Saving ${pending} change${pending > 1 ? 's' : ''}…` : LABEL[state.status]
  return (
    <span className="de-collab">
      <span className={`de-collab-status de-collab-${state.status}`} title={title}>
        <span className="de-collab-dot" />
        <span className="de-collab-label">{state.status === 'live' ? (peers.length ? `${peers.length + 1} editing` : 'Live') : LABEL[state.status]}</span>
      </span>
      {peers.length > 0 && (
        <span className="de-peers">
          {peers.slice(0, 5).map((p) => (
            <span key={p.site} className="de-peer" style={{ background: p.color }} title={p.name}>
              {initials(p.name)}
            </span>
          ))}
          {peers.length > 5 && (
            <span className="de-peer de-peer-more" title={peers.slice(5).map((p) => p.name).join(', ')}>
              +{peers.length - 5}
            </span>
          )}
        </span>
      )}
    </span>
  )
}

interface Mark {
  key: string
  color: string
  name: string
  caret?: { x: number; y: number; h: number }
  boxes: { x: number; y: number; w: number; h: number }[]
  outline?: { x: number; y: number; w: number; h: number }
}

/**
 * Collaborators' cursors drawn over the document (DOCX carets / selections, PPTX selected shapes).
 * Spreadsheets draw their own (in-grid) selection boxes. Render inside a positioned container.
 */
export function PeerOverlay({ client, collab }: { client: DocumentClient; collab: Collab | null }) {
  const ref = useRef<HTMLDivElement>(null)
  const peers = useSyncExternalStore(
    (fn) => subscribePeers(client, fn),
    () => getPeers(client),
  )
  const [marks, setMarks] = useState<Mark[]>([])
  const active = useMemo(() => peers.filter((p) => p.presence && (p.presence as { k?: string }).k !== 'xlsx'), [peers])

  useEffect(() => {
    const layer = ref.current
    const root = layer?.parentElement
    if (!layer || !root || !active.length) {
      setMarks([])
      return
    }
    let raf = 0
    const measure = () => {
      raf = 0
      const base = layer.getBoundingClientRect()
      const rel = (r: DOMRect) => ({ x: r.left - base.left, y: r.top - base.top, w: r.width, h: r.height })
      const out: Mark[] = []
      for (const p of active) {
        const m = place(p, root, collab, client)
        if (!m) continue
        out.push({
          key: String(p.site),
          color: p.color,
          name: p.name,
          caret: m.caret ? { x: m.caret.left - base.left, y: m.caret.top - base.top, h: m.caret.height || 16 } : undefined,
          boxes: m.sel.map(rel),
          outline: m.outline ? rel(m.outline) : undefined,
        })
      }
      setMarks(out)
    }
    const schedule = () => {
      if (!raf) raf = requestAnimationFrame(measure)
    }
    schedule()
    root.addEventListener('scroll', schedule, true)
    window.addEventListener('resize', schedule)
    const offChange = collab?.on('change', schedule)
    // views re-render asynchronously (paging, slide loading): keep up cheaply
    const t = setInterval(schedule, 400)
    return () => {
      cancelAnimationFrame(raf)
      root.removeEventListener('scroll', schedule, true)
      window.removeEventListener('resize', schedule)
      offChange?.()
      clearInterval(t)
    }
  }, [active, collab, client])

  return (
    <div ref={ref} className="de-peer-layer" aria-hidden>
      {marks.map((m) => (
        <div key={m.key} style={{ '--peer': m.color } as CSSProperties}>
          {m.boxes.map((b, i) => (
            <div key={i} className="de-peer-sel" style={{ left: b.x, top: b.y, width: b.w, height: b.h }} />
          ))}
          {m.outline && (
            <div className="de-peer-outline" style={{ left: m.outline.x - 3, top: m.outline.y - 3, width: m.outline.w + 6, height: m.outline.h + 6 }}>
              <span className="de-peer-flag">{m.name}</span>
            </div>
          )}
          {m.caret && (
            <div className="de-peer-caret" style={{ left: m.caret.x, top: m.caret.y, height: m.caret.h }}>
              <span className="de-peer-flag">{m.name}</span>
            </div>
          )}
        </div>
      ))}
    </div>
  )
}

function place(p: PeerView, root: HTMLElement, collab: Collab | null, client: DocumentClient): { caret?: DOMRect | null; sel: DOMRect[]; outline?: DOMRect } | null {
  const pr = p.presence as { k?: string; slide?: string; top?: number }
  if (pr.k === 'docx') {
    const r = collab?.rects(root, p)
    return r ? { caret: r.caret, sel: r.sel } : null
  }
  if (pr.k === 'pptx' && pr.slide) {
    const meta = client.info.pptx as PresentationMeta | undefined
    const i = meta?.slides.findIndex((s) => s.path === pr.slide) ?? -1
    const frame = i >= 0 ? root.querySelector<HTMLElement>(`[data-slide="${i + 1}"]`) : null
    if (!frame) return null
    const el = pr.top !== undefined ? frame.querySelector<HTMLElement>(`[data-top="${pr.top}"]`) : null
    return { sel: [], outline: (el ?? frame).getBoundingClientRect() }
  }
  return null
}
