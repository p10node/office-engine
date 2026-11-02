import { useEffect, useMemo, useRef, useState, type ReactNode } from 'react'
import type { Box, PresentationMeta, ShapeKind, ShapeOp, SlideModel, SlideOp } from '../../core/pptx/types'
import type { DocumentClient } from '../../core/worker/client'
import { useForceUpdate, useIsoLayoutEffect, useScroll, useSize } from '../common/hooks'
import { useFileOps, useViewer } from '../common/Slots'
import { slidesToText } from '../file/text'
import { Status } from '../common/Status'
import { ZoomControl, type Zoom } from '../common/ZoomControl'
import { getSel } from '../edit/input'
import { PptxEditor } from './PptxEditor'
import { PptxToolbar, type AlignTo, type Arrange, type Rotate, type SlideCmd } from './PptxToolbar'
import { ContextMenu, MOD, type MenuItem } from '../common/ContextMenu'
import { Presenter } from './Presenter'
import { pickFiles } from '../edit/Toolbar'
import { SlideView, type SlideEditing } from './Slide'
import { publishLocal, registerHost } from '../../collab/hosts'

interface SlideCache {
  models: Map<number, SlideModel>
  pending: Map<number, Promise<SlideModel>>
  listeners: Map<number, Set<(m: SlideModel) => void>>
}

const caches = new WeakMap<DocumentClient, SlideCache>()

function cacheFor(client: DocumentClient): SlideCache {
  let c = caches.get(client)
  if (!c) caches.set(client, (c = { models: new Map(), pending: new Map(), listeners: new Map() }))
  return c
}

/** replaces a cached slide (after an edit) and notifies mounted views */
function putSlide(client: DocumentClient, index: number, m: SlideModel) {
  const c = cacheFor(client)
  c.models.set(index, m)
  for (const fn of c.listeners.get(index) ?? []) fn(m)
}

/** forget every cached slide (after slides were added / removed / reordered) */
function resetSlides(client: DocumentClient) {
  const c = cacheFor(client)
  c.models.clear()
  c.pending.clear()
}

export function useSlide(client: DocumentClient, index: number): SlideModel | { error: string } | undefined {
  const cache = cacheFor(client)
  const [model, setModel] = useState<SlideModel | { error: string } | undefined>(() => cache.models.get(index))
  useEffect(() => {
    let set = cache.listeners.get(index)
    if (!set) cache.listeners.set(index, (set = new Set()))
    const fn = (m: SlideModel) => setModel(m)
    set.add(fn)
    return () => {
      set!.delete(fn)
    }
  }, [cache, index])
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

export function SlideSlot(props: { client: DocumentClient; index: number; meta: PresentationMeta; scale: number; live: boolean; editing?: SlideEditing; overlay?: ReactNode }) {
  const { client, index, meta, scale, live, editing, overlay } = props
  const model = useSlide(client, index)
  if (!model) return <div className="de-slide" style={{ width: meta.w * scale, height: meta.h * scale, background: '#fff' }} />
  if ('error' in model) {
    return (
      <div className="de-slide de-ph" style={{ width: meta.w * scale, height: meta.h * scale, display: 'flex' }}>
        {model.error}
      </div>
    )
  }
  return <SlideView model={model} client={client} scale={scale} live={live} editing={editing} overlay={overlay} />
}

type DragMode = 'move' | 'n' | 's' | 'e' | 'w' | 'ne' | 'nw' | 'se' | 'sw'
const HANDLES: DragMode[] = ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w']

function boxOf(model: SlideModel | undefined, top: number): Box | undefined {
  const els = model?.els.filter((e) => e.top === top) ?? []
  if (!els.length) return undefined
  const x0 = Math.min(...els.map((e) => e.box.x))
  const y0 = Math.min(...els.map((e) => e.box.y))
  const x1 = Math.max(...els.map((e) => e.box.x + e.box.w))
  const y1 = Math.max(...els.map((e) => e.box.y + e.box.h))
  return { x: x0, y: y0, w: x1 - x0, h: y1 - y0, rot: els.length === 1 ? els[0].box.rot : undefined }
}

function resized(b: Box, mode: DragMode, dx: number, dy: number, keep: boolean): Box {
  if (mode === 'move') return { ...b, x: b.x + dx, y: b.y + dy }
  let { x, y, w, h } = b
  if (mode.includes('e')) w += dx
  if (mode.includes('s')) h += dy
  if (mode.includes('w')) {
    x += dx
    w -= dx
  }
  if (mode.includes('n')) {
    y += dy
    h -= dy
  }
  if (keep && mode.length === 2 && b.w && b.h) {
    const k = Math.max(w / b.w, h / b.h)
    const nw = b.w * k
    const nh = b.h * k
    if (mode.includes('w')) x = b.x + b.w - nw
    if (mode.includes('n')) y = b.y + b.h - nh
    w = nw
    h = nh
  }
  return { ...b, x, y, w: Math.max(4, w), h: Math.max(4, h) }
}

function Selection({ box, scale, onHandle }: { box: Box; scale: number; onHandle: (m: DragMode, e: React.MouseEvent) => void }) {
  const hs = 9 / scale
  const pos: Record<DragMode, [number, number]> = { nw: [0, 0], n: [0.5, 0], ne: [1, 0], e: [1, 0.5], se: [1, 1], s: [0.5, 1], sw: [0, 1], w: [0, 0.5], move: [0, 0] }
  return (
    <div className="de-shape-sel" style={{ left: box.x, top: box.y, width: box.w, height: box.h, borderWidth: 1.5 / scale, transform: box.rot ? `rotate(${box.rot}deg)` : undefined }}>
      {HANDLES.map((m) => (
        <div
          key={m}
          className={`de-handle de-h-${m}`}
          style={{ width: hs, height: hs, left: pos[m][0] * box.w - hs / 2, top: pos[m][1] * box.h - hs / 2, borderWidth: 1 / scale }}
          onMouseDown={(e) => {
            e.stopPropagation()
            e.preventDefault()
            onHandle(m, e)
          }}
        />
      ))}
    </div>
  )
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
  const viewer = useViewer()
  const meta = client.info.pptx as PresentationMeta
  useFileOps({
    printable: async () => {
      const m = client.info.pptx as PresentationMeta
      // hidden flags are known once a slide is parsed: filter on the models
      const all = await Promise.all(m.slides.map((s) => client.call('pptx.slide', s.index)))
      const shown = all.filter((s) => !s.hidden)
      const slides = shown.length ? shown : all
      return {
        size: `${m.w}px ${m.h}px`,
        margin: '0',
        node: slides.map((model) => (
          <div key={model.index} className="de-print-page" style={{ width: m.w, height: m.h }}>
            <SlideView model={{ ...model, hidden: false }} client={client} scale={1} />
          </div>
        )),
      }
    },
    text: async () => {
      const m = client.info.pptx as PresentationMeta
      return slidesToText(await Promise.all(m.slides.map((s) => client.call('pptx.slide', s.index))))
    },
    stats: () => {
      const m = client.info.pptx as PresentationMeta
      return [
        ['Slides', m.count.toLocaleString()],
        ['Slide size', `${+(m.w / 96).toFixed(2)} × ${+(m.h / 96).toFixed(2)} in`],
      ]
    },
  })
  const [rootRef, rootSize] = useSize<HTMLDivElement>()
  const [scrollRef, size] = useSize<HTMLDivElement>()
  const scroll = useScroll(scrollRef)
  const thumbRef = useRef<HTMLDivElement>(null)
  const thumbScroll = useScroll(thumbRef)
  const [zoom, setZoom] = useState<Zoom>(initialZoom)
  const showThumbs = thumbnails && rootSize.w >= 720 && meta.count > 1
  const [, bump] = useForceUpdate()
  const editor = useMemo(() => new PptxEditor(client), [client])
  const caretAt = useRef<{ x: number; y: number } | null>(null)
  const textRoot = useRef<HTMLDivElement | null>(null)
  const [selected, setSelected] = useState<{ slide: number; top: number } | null>(null)
  const [drag, setDrag] = useState<{ slide: number; top: number; mode: DragMode; sx: number; sy: number; box: Box; cur: Box; moved: boolean; sid?: number; px: number; py: number } | null>(null)
  const [slidesVer, setSlidesVer] = useState(0)
  const [presenting, setPresenting] = useState<number | null>(null)
  const [ctxMenu, setCtxMenu] = useState<{ x: number; y: number; items: MenuItem[] } | null>(null)
  /** shape edits that can be undone */
  const [undoDepth, setUndoDepth] = useState(0)

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

  // ---- editing ----
  useEffect(() => {
    editor.onChange = bump
    editor.onDirty = () => viewer.setDirty(true)
    editor.onCommitted = (slide, m) => putSlide(client, slide, m)
    editor.shortcut = (e) => {
      if ((e.metaKey || e.ctrlKey) && e.key.toLowerCase() === 's') {
        void viewer.save()
        return true
      }
      if (e.key === 'Escape') {
        stopEditing()
        return true
      }
      return false
    }
  })
  useEffect(() => {
    if (!viewer.editable) stopEditing()
    return () => editor.end()
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [viewer.editable])

  const stopEditing = () => {
    if (!editor.active) return
    editor.end()
    if (textRoot.current) getSel(textRoot.current)?.removeAllRanges()
    // keep keyboard shortcuts (Del, ⌘D, arrows…) working on the shape that stays selected
    scrollRef.current?.focus({ preventScroll: true })
    bump()
  }

  // ---- collaboration (src/collab) ----
  const stopRef = useRef(stopEditing)
  stopRef.current = stopEditing
  useEffect(() => registerHost(client, 'pptx', { endEditing: () => stopRef.current() }), [client])
  useEffect(
    () =>
      client.on('collab:pptx', (r) => {
        if (r.meta) {
          resetSlides(client)
          setSelected(null)
          setSlidesVer((v) => v + 1)
        } else if (r.slide !== undefined && r.model) putSlide(client, r.slide, r.model)
      }),
    [client],
  )
  useEffect(() => publishLocal(client, { slide: selected?.slide ?? current, top: selected?.top }), [client, selected?.slide, selected?.top, current])

  const editing: SlideEditing | undefined =
    editor.active && editor.text()
      ? {
          sid: editor.sid,
          text: editor.text()!,
          edit: {
            root: (el) => {
              textRoot.current = el
              if (el) editor.input.attach(el)
            },
          },
        }
      : undefined

  useIsoLayoutEffect(() => {
    const root = textRoot.current
    if (!editor.active || !root) return
    if (caretAt.current) {
      const { x, y } = caretAt.current
      caretAt.current = null
      root.focus({ preventScroll: true })
      const d = root.ownerDocument as Document & { caretPositionFromPoint?: (x: number, y: number) => { offsetNode: Node; offset: number } | null }
      const pos = d.caretPositionFromPoint?.(x, y)
      const range = pos ? null : d.caretRangeFromPoint?.(x, y)
      const node = pos?.offsetNode ?? range?.startContainer
      const off = pos?.offset ?? range?.startOffset ?? 0
      const s = getSel(root)
      if (node && root.contains(node)) s?.collapse(node, off)
      else {
        s?.selectAllChildren(root)
        s?.collapseToEnd()
      }
      editor.input.lastSel = editor.input.readSel()
      return
    }
    const sel = editor.pendingSel
    if (sel) {
      editor.pendingSel = null
      editor.input.restore(sel)
    }
  })

  const modelOf = (i: number) => cacheFor(client).models.get(i)

  const beginText = (index: number, sid: number, x: number, y: number) => {
    const el = modelOf(index)?.els.find((e) => e.t === 'shape' && e.sid === sid)
    if (!el || el.t !== 'shape' || !el.text) return
    editor.begin(index, sid, el.text)
    caretAt.current = { x, y }
    bump()
  }

  const onSlideMouseDown = (index: number, e: React.MouseEvent) => {
    if (!viewer.editable || e.button !== 0) return
    const t = e.target as HTMLElement
    if (editor.active && textRoot.current?.contains(t)) return
    const topEl = t.closest('[data-top]') as HTMLElement | null
    const top = topEl ? Number(topEl.dataset.top) : NaN
    stopEditing()
    if (!topEl || Number.isNaN(top)) {
      setSelected(null)
      return
    }
    e.preventDefault()
    scrollRef.current?.focus({ preventScroll: true })
    const box = boxOf(modelOf(index), top)
    setSelected({ slide: index, top })
    if (!box) return
    const sidEl = t.closest('[data-sid]') as HTMLElement | null
    setDrag({ slide: index, top, mode: 'move', sx: e.clientX, sy: e.clientY, box, cur: box, moved: false, sid: sidEl ? Number(sidEl.dataset.sid) : undefined, px: e.clientX, py: e.clientY })
  }

  const runShapes = async (slide: number, ops: ShapeOp[], selectNew = false) => {
    try {
      const m = await client.call('pptx.shapes', slide, ops)
      putSlide(client, slide, m)
      viewer.setDirty(true)
      setUndoDepth((d) => d + 1)
      if (selectNew) {
        const tops = m.els.map((x) => x.top ?? -1)
        setSelected({ slide, top: Math.max(...tops) })
      }
      return m
    } catch (err) {
      console.error('[office-engine]', err)
    }
  }

  useEffect(() => {
    if (!drag) return
    const move = (e: MouseEvent) => {
      const dx = (e.clientX - drag.sx) / scale
      const dy = (e.clientY - drag.sy) / scale
      const moved = drag.moved || Math.abs(e.clientX - drag.sx) + Math.abs(e.clientY - drag.sy) > 3
      setDrag({ ...drag, moved, cur: moved ? resized(drag.box, drag.mode, dx, dy, e.shiftKey) : drag.box })
    }
    const up = () => {
      const d = drag
      setDrag(null)
      if (d.moved) void runShapes(d.slide, [{ op: 'xfrm', top: d.top, box: d.cur }])
      // a click on text edits it (Esc then selects the whole shape, like PowerPoint)
      else if (d.mode === 'move' && d.sid !== undefined) beginText(d.slide, d.sid, d.px, d.py)
    }
    window.addEventListener('mousemove', move)
    window.addEventListener('mouseup', up, { once: true })
    return () => {
      window.removeEventListener('mousemove', move)
      window.removeEventListener('mouseup', up)
    }
  })

  const insert = async (kind: ShapeKind, at?: number) => {
    const slide = at ?? selected?.slide ?? current
    const w = kind === 'text' ? meta.w * 0.4 : kind === 'line' ? meta.w * 0.25 : meta.w * 0.2
    const h = kind === 'text' ? meta.h * 0.1 : kind === 'line' ? 1 : meta.w * 0.2
    let box: Box = { x: (meta.w - w) / 2, y: (meta.h - h) / 2, w, h }
    let rid: string | undefined
    if (kind === 'image') {
      const [f] = await pickFiles('image/*', false)
      if (!f) return
      rid = await client.call('pptx.image', slide, await f.arrayBuffer(), (f.type.split('/')[1] ?? 'png').replace('svg+xml', 'svg'))
      const dims = await new Promise<{ w: number; h: number }>((res) => {
        const img = new Image()
        img.onload = () => res({ w: img.naturalWidth, h: img.naturalHeight })
        img.onerror = () => res({ w: 400, h: 300 })
        img.src = URL.createObjectURL(f)
      })
      const k = Math.min(1, (meta.w * 0.6) / dims.w, (meta.h * 0.6) / dims.h)
      box = { x: (meta.w - dims.w * k) / 2, y: (meta.h - dims.h * k) / 2, w: dims.w * k, h: dims.h * k }
    }
    const m = await runShapes(slide, [{ op: 'insert', kind, box, rid, text: kind === 'text' ? 'Text' : undefined }], true)
    // the new shape is selected: arrow keys, Delete, ⌘D act on it right away
    scrollRef.current?.focus({ preventScroll: true })
    if (m && kind === 'text') {
      const el = m.els[m.els.length - 1]
      if (el?.t === 'shape' && el.sid !== undefined && el.text) {
        editor.begin(slide, el.sid, el.text)
        bump()
      }
    }
  }

  const slideAction = async (op: SlideOp, goTo?: number) => {
    stopEditing()
    setSelected(null)
    const next = await client.call('pptx.slides', op)
    Object.assign(meta, next)
    resetSlides(client)
    viewer.setDirty(true)
    setSlidesVer((v) => v + 1)
    if (goTo !== undefined) requestAnimationFrame(() => go(goTo))
  }

  const undoShapes = async (redo: boolean) => {
    const r = await client.call('pptx.undo', redo)
    if (!r) return
    putSlide(client, r.index, r.model)
    setUndoDepth((d) => Math.max(0, d + (redo ? 1 : -1)))
    setSelected(null)
    go(r.index)
  }

  const slideCmd = (kind: SlideCmd, i = current) => {
    if (kind === 'add') void slideAction({ op: 'add', after: i }, i + 1)
    else if (kind === 'duplicate') void slideAction({ op: 'duplicate', index: i }, i + 1)
    else if (kind === 'delete') void slideAction({ op: 'delete', index: i }, Math.max(0, i - 1))
    else if (kind === 'up') i > 0 && void slideAction({ op: 'move', from: i, to: i - 1 }, i - 1)
    else if (kind === 'down') i < meta.count - 1 && void slideAction({ op: 'move', from: i, to: i + 1 }, i + 1)
    else if (kind === 'hide') void slideAction({ op: 'hide', index: i, hidden: !modelOf(i)?.hidden }, i)
  }

  const elOf = (sel: { slide: number; top: number } | null) => (sel ? modelOf(sel.slide)?.els.find((x) => x.top === sel.top) : undefined)
  // menus act on the shape that was right-clicked even before React re-renders
  const selRef = useRef(selected)
  selRef.current = selected
  const shapeCmd = {
    arrange: (to: Arrange) => {
      const selected = selRef.current
      if (selected) void runShapes(selected.slide, [{ op: 'order', top: selected.top, to }])
    },
    duplicate: () => {
      const selected = selRef.current
      if (selected) void runShapes(selected.slide, [{ op: 'duplicate', top: selected.top }], true)
    },
    remove: () => {
      const selected = selRef.current
      if (selected) void runShapes(selected.slide, [{ op: 'delete', top: selected.top }])
      setSelected(null)
    },
    rotate: (k: Rotate) => {
      const selected = selRef.current
      const el = elOf(selected)
      const box = selected && boxOf(modelOf(selected.slide), selected.top)
      if (!selected || !box) return
      const b: Box = { ...box, flipH: el?.box.flipH, flipV: el?.box.flipV }
      if (k === 'cw') b.rot = (b.rot ?? 0) + 90
      else if (k === 'ccw') b.rot = (b.rot ?? 0) - 90
      else if (k === 'flipH') b.flipH = !b.flipH
      else b.flipV = !b.flipV
      void runShapes(selected.slide, [{ op: 'xfrm', top: selected.top, box: b }])
    },
    align: (k: AlignTo | 'both') => {
      const selected = selRef.current
      const box = selected && boxOf(modelOf(selected.slide), selected.top)
      if (!selected || !box) return
      const b = { ...box }
      // align the visible (rotated) extent
      const r = ((box.rot ?? 0) * Math.PI) / 180
      const vw = Math.abs(box.w * Math.cos(r)) + Math.abs(box.h * Math.sin(r))
      const vh = Math.abs(box.w * Math.sin(r)) + Math.abs(box.h * Math.cos(r))
      const dx = (vw - box.w) / 2
      const dy = (vh - box.h) / 2
      if (k === 'left') b.x = dx
      else if (k === 'center' || k === 'both') b.x = (meta.w - b.w) / 2
      else if (k === 'right') b.x = meta.w - b.w - dx
      else if (k === 'top') b.y = dy
      else if (k === 'middle') b.y = (meta.h - b.h) / 2
      else if (k === 'bottom') b.y = meta.h - b.h - dy
      if (k === 'both') b.y = (meta.h - b.h) / 2
      void runShapes(selected.slide, [{ op: 'xfrm', top: selected.top, box: b }])
    },
    anchor: (a: 't' | 'ctr' | 'b') => {
      if (editor.active && editor.body) {
        editor.body = { ...editor.body, anchor: a }
        void runShapes(editor.slide, [{ op: 'anchor', sid: editor.sid, anchor: a }])
        bump()
        return
      }
      const selected = selRef.current
      const el = elOf(selected)
      if (selected && el?.t === 'shape' && el.sid !== undefined) void runShapes(selected.slide, [{ op: 'anchor', sid: el.sid, anchor: a }])
    },
  }

  const shapeMenu = (): MenuItem[] => [
    { label: 'Duplicate', icon: 'copy', hint: `${MOD}D`, run: shapeCmd.duplicate },
    { label: 'Delete', icon: 'trash', hint: 'Del', run: shapeCmd.remove },
    { label: 'Bring to front', icon: 'front', hint: `${MOD}⇧↑`, sep: true, run: () => shapeCmd.arrange('front') },
    { label: 'Bring forward', hint: `${MOD}↑`, run: () => shapeCmd.arrange('forward') },
    { label: 'Send backward', hint: `${MOD}↓`, run: () => shapeCmd.arrange('backward') },
    { label: 'Send to back', icon: 'back', hint: `${MOD}⇧↓`, run: () => shapeCmd.arrange('back') },
    { label: 'Rotate 90° clockwise', icon: 'rotate', sep: true, run: () => shapeCmd.rotate('cw') },
    { label: 'Rotate 90° counter-clockwise', icon: 'rotateCcw', run: () => shapeCmd.rotate('ccw') },
    { label: 'Flip horizontally', icon: 'flipH', run: () => shapeCmd.rotate('flipH') },
    { label: 'Flip vertically', icon: 'flipV', run: () => shapeCmd.rotate('flipV') },
    { label: 'Center on slide', icon: 'alignObj', sep: true, run: () => shapeCmd.align('both') },
  ]
  const slideMenu = (i: number): MenuItem[] => [
    { label: 'New slide', icon: 'slideAdd', hint: `${MOD}M`, run: () => slideCmd('add', i) },
    { label: 'Duplicate slide', icon: 'copy', run: () => slideCmd('duplicate', i) },
    { label: 'Delete slide', icon: 'trash', run: () => slideCmd('delete', i), disabled: meta.count <= 1 },
    { label: modelOf(i)?.hidden ? 'Show slide' : 'Hide slide', icon: 'hide', run: () => slideCmd('hide', i) },
    { label: 'Move slide up', icon: 'up', sep: true, run: () => slideCmd('up', i), disabled: i === 0 },
    { label: 'Move slide down', icon: 'down', run: () => slideCmd('down', i), disabled: i >= meta.count - 1 },
    { label: 'Insert text box', icon: 'textbox', sep: true, run: () => void insert('text', i) },
    { label: 'Insert image…', icon: 'image', run: () => void insert('image', i) },
    { label: 'Present from this slide', icon: 'play', sep: true, run: () => setPresenting(i) },
  ]

  const onShapeKey = (e: React.KeyboardEvent) => {
    if (!selected || editor.active || !viewer.editable) return false
    const box = boxOf(modelOf(selected.slide), selected.top)
    const step = e.shiftKey ? 10 : 1
    const nudge: Record<string, [number, number]> = { ArrowUp: [0, -step], ArrowDown: [0, step], ArrowLeft: [-step, 0], ArrowRight: [step, 0] }
    const mod = e.metaKey || e.ctrlKey
    if (mod && e.key.toLowerCase() === 'd') shapeCmd.duplicate()
    else if (mod && (e.key === 'ArrowUp' || e.key === 'ArrowDown')) shapeCmd.arrange(e.key === 'ArrowUp' ? (e.shiftKey ? 'front' : 'forward') : e.shiftKey ? 'back' : 'backward')
    else if (e.key === 'Delete' || e.key === 'Backspace') {
      void runShapes(selected.slide, [{ op: 'delete', top: selected.top }])
      setSelected(null)
    } else if (nudge[e.key] && box) void runShapes(selected.slide, [{ op: 'xfrm', top: selected.top, box: { ...box, x: box.x + nudge[e.key][0], y: box.y + nudge[e.key][1] } }])
    else if (e.key === 'Escape') setSelected(null)
    else return false
    e.preventDefault()
    return true
  }

  const go = (i: number) => {
    const el = scrollRef.current
    if (el) el.scrollTop = PAD + Math.max(0, Math.min(meta.count - 1, i)) * pitch - Math.max(0, (size.h - slideH) / 2)
  }

  const onKeyDown = (e: React.KeyboardEvent) => {
    // already handled by the text editor (e.g. Esc that just left text editing)
    if (editor.active || e.defaultPrevented) return
    if (onShapeKey(e)) return
    if ((e.metaKey || e.ctrlKey) && e.key.toLowerCase() === 's') {
      e.preventDefault()
      void viewer.save()
      return
    }
    if (e.key === 'F5') {
      e.preventDefault()
      setPresenting(e.shiftKey ? 0 : current)
      return
    }
    const mod = e.metaKey || e.ctrlKey
    if (viewer.editable && mod && (e.key.toLowerCase() === 'z' || e.key.toLowerCase() === 'y')) {
      e.preventDefault()
      void undoShapes(e.key.toLowerCase() === 'y' || e.shiftKey)
      return
    }
    if (viewer.editable && mod && e.key.toLowerCase() === 'm') {
      e.preventDefault()
      slideCmd('add')
      return
    }
    if (e.key === 'PageDown' || e.key === 'ArrowRight' || (e.key === ' ' && !e.shiftKey)) go(current + 1)
    else if (e.key === 'PageUp' || e.key === 'ArrowLeft') go(current - 1)
    else if (e.key === 'Home') go(0)
    else if (e.key === 'End') go(meta.count - 1)
    else return
    e.preventDefault()
  }

  const selOverlay = (i: number) => {
    const box = drag && drag.slide === i && drag.moved ? drag.cur : selected ? boxOf(modelOf(i), selected.top) : undefined
    if (!box || !selected) return undefined
    return (
      <Selection
        box={box}
        scale={scale}
        onHandle={(mode, e) => {
          const b = boxOf(modelOf(i), selected.top)
          if (b) setDrag({ slide: i, top: selected.top, mode, sx: e.clientX, sy: e.clientY, box: b, cur: b, moved: false, px: e.clientX, py: e.clientY })
        }}
      />
    )
  }

  const slots = []
  for (let i = first; i <= last; i++) {
    slots.push(
      <div key={`${slidesVer}:${i}`} className={viewer.editable ? 'de-slide-frame de-slide-editable' : 'de-slide-frame'} style={{ top: PAD + i * pitch, left }} data-slide={i + 1}
        onMouseDown={(e) => onSlideMouseDown(i, e)}
        onContextMenu={(e) => {
          if (!viewer.editable) return
          const t = e.target as HTMLElement
          if (editor.active && textRoot.current?.contains(t)) return
          e.preventDefault()
          const topEl = t.closest('[data-top]') as HTMLElement | null
          if (topEl) {
            stopEditing()
            const sel = { slide: i, top: Number(topEl.dataset.top) }
            selRef.current = sel
            setSelected(sel)
            setCtxMenu({ x: e.clientX, y: e.clientY, items: shapeMenu() })
          } else setCtxMenu({ x: e.clientX, y: e.clientY, items: slideMenu(i) })
        }}
      >
        <SlideSlot
          client={client}
          index={i}
          meta={meta}
          scale={scale}
          live
          editing={editor.slide === i ? editing : undefined}
          overlay={viewer.editable && selected?.slide === i ? selOverlay(i) : undefined}
        />
      </div>,
    )
  }
  if (editor.active && (editor.slide < first || editor.slide > last)) stopEditing()
  const thumbs = []
  if (showThumbs) {
    for (let i = tFirst; i <= tLast; i++) {
      thumbs.push(
        <div
          key={`${slidesVer}:${i}`}
          className="de-thumb"
          style={{ top: 10 + i * tPitch }}
          aria-current={i === current}
          onClick={() => go(i)}
          onContextMenu={(e) => {
            if (!viewer.editable) return
            e.preventDefault()
            go(i)
            setCtxMenu({ x: e.clientX, y: e.clientY, items: slideMenu(i) })
          }}
        >
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
        <div ref={scrollRef} className="de-scroll" tabIndex={0} onKeyDown={onKeyDown} style={{ outline: 'none' }} onMouseDown={(e) => e.target === e.currentTarget && stopEditing()}>
          <div style={{ position: 'relative', height: totalH, width: totalW }} onMouseDown={(e) => e.target === e.currentTarget && stopEditing()}>
            {slots}
          </div>
        </div>
        <ZoomControl zoom={zoom} effective={scale} onChange={setZoom} />
        {!viewer.editable && (
          <button type="button" className="de-mode de-present-btn" title="Present (F5)" onClick={() => setPresenting(current)}>
            ▶ Present
          </button>
        )}
        <Status text={`Slide ${current + 1} of ${meta.count}`} onStatus={onStatus} />
        {viewer.editable && (
          <PptxToolbar
            editor={editor}
            shape={
              selected && !editor.active
                ? (() => {
                    const el = elOf(selected)
                    const sh = el?.t === 'shape' ? el : undefined
                    return { canText: sh?.sid !== undefined, anchor: sh?.text?.anchor }
                  })()
                : null
            }
            textAnchor={editor.body?.anchor}
            canUndo={undoDepth > 0}
            onUndo={(redo) => void undoShapes(redo)}
            onInsert={(k) => void insert(k)}
            onFill={(color) => selected && void runShapes(selected.slide, [{ op: 'fill', top: selected.top, color }])}
            onLine={(p) => selected && void runShapes(selected.slide, [{ op: 'line', top: selected.top, ...p }])}
            onArrange={shapeCmd.arrange}
            onRotate={shapeCmd.rotate}
            onAlign={shapeCmd.align}
            onAnchor={shapeCmd.anchor}
            onDuplicate={shapeCmd.duplicate}
            onDelete={shapeCmd.remove}
            onBackground={(color) => void runShapes(current, [{ op: 'bg', color }])}
            onSlide={(k) => slideCmd(k)}
            onPresent={(fromStart) => setPresenting(fromStart ? 0 : current)}
          />
        )}
        {ctxMenu && <ContextMenu x={ctxMenu.x} y={ctxMenu.y} items={ctxMenu.items} onClose={() => setCtxMenu(null)} />}
        {presenting !== null && <Presenter client={client} meta={meta} start={presenting} onClose={() => setPresenting(null)} />}
      </div>
    </div>
  )
}
