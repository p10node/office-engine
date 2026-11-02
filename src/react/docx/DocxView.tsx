import { useEffect, useMemo, useRef, useState, type CSSProperties } from 'react'
import type { Block, DocxMeta, PageSetup, Paragraph } from '../../core/docx/types'
import { DEFAULT_PAGE } from '../../core/docx/types'
import { registerHost } from '../../collab/hosts'
import type { DocumentClient } from '../../core/worker/client'
import { lowerIndex, useForceUpdate, useIsoLayoutEffect, useScroll, useSize } from '../common/hooks'
import { useViewer } from '../common/Slots'
import { Status } from '../common/Status'
import { ZoomControl, type Zoom } from '../common/ZoomControl'
import { ContextMenu, MOD, type MenuItem } from '../common/ContextMenu'
import { PromptDialog, Toast, type PromptSpec } from '../common/Dialog'
import { domFromPos, getSel, opInsert, posFromDom } from '../edit/input'
import { download, pickFiles } from '../edit/Toolbar'
import { runsText, type RRun } from '../edit/rich'
import { DocxEditor, type DocState, type FormatState } from './DocxEditor'
import { DocxToolbar } from './DocxToolbar'
import { FindBar, type Match } from './FindBar'
import { BlockView, contextual, type RenderCtx } from './render'

const PX = 4 / 3
const GAP = 16
const PAD = 24

interface PageSlice {
  start: number
  end: number
  used: number
}

function paginate(blocks: Block[], heights: WeakMap<Block, number>, count: number, avail: number): PageSlice[] {
  const pages: PageSlice[] = []
  let cur: PageSlice = { start: 0, end: 0, used: 0 }
  for (let i = 0; i < count; i++) {
    const b = blocks[i]
    const h = heights.get(b) ?? 0
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

/** the rendered element of the inline object at (pa, o) — its <img> when it is a picture */
function findObj(root: HTMLElement, pa: string, o: number): HTMLElement | null {
  for (const p of root.querySelectorAll<HTMLElement>(`[data-pa="${pa}"]`)) {
    for (const sp of p.querySelectorAll<HTMLElement>(`[data-obj][data-o="${o}"]`)) {
      if (sp.closest('[data-pa]') === p) return sp.querySelector('img') ?? sp
    }
  }
  return null
}

interface Box {
  x: number
  y: number
  w: number
  h: number
}
const HANDLES = ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w'] as const

/** heading entries for the outline panel */
function outlineOf(blocks: Block[], titleId: string | undefined) {
  const out: { bi: number; lvl: number; text: string }[] = []
  blocks.forEach((b, bi) => {
    if (b.t !== 'p' || b.pid === undefined) return
    const title = !!titleId && b.sid === titleId
    const lvl = title ? 0 : b.ps.outline
    if (lvl === undefined || lvl > 5) return
    const text = runsText(b.runs as RRun[]).trim()
    if (text) out.push({ bi, lvl: title ? 0 : lvl + (titleId ? 1 : 0), text })
  })
  return out
}
const editable = (b: Block | undefined): b is Paragraph => !!b && b.t === 'p' && b.pid !== undefined && !b.ro

export interface DocxViewProps {
  client: DocumentClient
  onStatus?: (s: string) => void
  initialZoom?: Zoom
}

export function DocxView({ client, onStatus, initialZoom = 'fit' }: DocxViewProps) {
  const viewer = useViewer()
  const st = useRef<DocState>({ blocks: [], heights: new WeakMap(), dirty: new Set(), frontier: 0 }).current
  const [version, bump] = useForceUpdate()
  const [meta, setMeta] = useState<DocxMeta | null>(null)
  const [done, setDone] = useState<{ error?: string } | null>(null)
  const [zoom, setZoom] = useState<Zoom>(initialZoom)
  const [scrollRef, size] = useSize<HTMLDivElement>()
  const scroll = useScroll(scrollRef)
  const pagesRef = useRef<HTMLDivElement>(null)
  const editor = useMemo(() => new DocxEditor(client, st, bump), [client, st, bump])
  const isEditing = viewer.editable && !!done && !done.error
  const [find, setFind] = useState<{ open: boolean; q: string; idx: number }>({ open: false, q: '', idx: 0 })
  const [ctxMenu, setCtxMenu] = useState<{ x: number; y: number } | null>(null)
  const [prompt, setPrompt] = useState<PromptSpec | null>(null)
  const [toast, setToast] = useState<string | null>(null)
  const [outline, setOutline] = useState(false)
  const [painter, setPainter] = useState(false)
  const [img, setImg] = useState<FormatState['img'] | null>(null)
  const [imgBox, setImgBox] = useState<Box | null>(null)
  const resizing = useRef(false)
  /** pt per screen px of the picture being resized */
  const resizeRatio = useRef<number | null>(null)

  useEffect(() => {
    editor.onImage = (x) => setImg((p) => (p?.pa === x?.pa && p?.o === x?.o && p?.w === x?.w && p?.h === x?.h && p?.target === x?.target && p?.alt === x?.alt ? p : x))
    editor.onStyles = (styles) => setMeta((m) => (m ? { ...m, styles } : m))
    editor.onPainter = setPainter
    return () => {
      editor.onImage = editor.onStyles = editor.onPainter = undefined
    }
  }, [editor])

  // collaboration sessions edit through this editor once the document is fully loaded
  useEffect(() => (done && !done.error ? registerHost(client, 'docx', { editor, st }) : undefined), [client, editor, st, done])

  useEffect(() => {
    const offs = [
      client.on('docx:blocks', (b) => {
        for (const x of b) st.blocks.push(x)
        bump()
      }),
      client.on('docx:meta', (m) => {
        setMeta(m)
        if (m.styles) editor.styles = m.styles
      }),
      client.on('docx:done', setDone),
    ]
    if (!started.has(client)) {
      started.add(client)
      client.call('docx.start').catch(() => {})
    }
    return () => offs.forEach((f) => f())
  }, [client, bump, st, editor])

  const page: PageSetup = meta?.page ?? DEFAULT_PAGE
  const contentW = page.w - page.ml - page.mr
  const availPx = (page.h - page.mt - page.mb) * PX

  // ---- measurement: edited blocks first, then the not-yet-measured tail (time-sliced) ----
  const [batch, setBatch] = useState(40)
  const measureRef = useRef<HTMLDivElement>(null)
  const widthKey = `${contentW}|${meta?.font}|${meta?.sz}`
  const lastKey = useRef(widthKey)
  if (lastKey.current !== widthKey) {
    lastKey.current = widthKey
    st.heights = new WeakMap()
    st.dirty.clear()
    st.frontier = 0
  }
  const total = st.blocks.length
  const queue: { b: Block; i: number }[] = []
  for (const b of st.dirty) {
    if (queue.length >= 50) break
    const i = st.blocks.indexOf(b)
    if (i < 0) st.dirty.delete(b)
    else queue.push({ b, i })
  }
  const tailEnd = Math.min(total, st.frontier + batch)
  for (let i = st.frontier; i < tailEnd; i++) queue.push({ b: st.blocks[i], i })
  const renderStart = performance.now()
  const measureCtx = useMemo<RenderCtx>(() => ({ client, page: -1, pages: 0, measuring: true, editable: false }), [client])

  useIsoLayoutEffect(() => {
    const el = measureRef.current
    if (!el || !queue.length || !meta) return
    const hs = Array.from(el.children).map((c) => (c as HTMLElement).getBoundingClientRect().height)
    const dt = performance.now() - renderStart
    const id = setTimeout(() => {
      queue.forEach((q, k) => {
        st.heights.set(q.b, hs[k] ?? 0)
        st.dirty.delete(q.b)
      })
      if (tailEnd > st.frontier) st.frontier = tailEnd
      setBatch((b) => (dt < 10 ? Math.min(b * 2, 1500) : dt > 30 ? Math.max(10, b >> 1) : b))
      bump()
    }, 0)
    return () => clearTimeout(id)
  })

  const pages = useMemo(
    () => paginate(st.blocks, st.heights, Math.min(st.frontier, total), availPx),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [version, availPx, total],
  )

  // ---- layout ----
  const pageW = page.w * PX
  const pageH = page.h * PX
  const fit = size.w ? Math.max(0.2, Math.min(2, (size.w - PAD * 2) / pageW)) : 1
  const z = zoom === 'fit' ? Math.min(fit, 1.25) : zoom
  const complete = !!done && st.frontier >= total
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
  // keep the page holding the caret mounted so the DOM selection survives scrolling
  const caretBi = isEditing && editor.input.lastSel ? Number(editor.input.lastSel.f.pa.split('.')[0]) : -1
  const caretPage = caretBi >= 0 ? pages.findIndex((p) => caretBi >= p.start && caretBi < p.end) : -1

  editor.contentW = contentW
  editor.defaultFont = meta?.font
  editor.defaultSz = meta?.sz
  const cur = tops.length ? lowerIndex(tops, scroll.y + size.h / 3) + 1 : 0
  // word count (debounced, off the typing path)
  const [words, setWords] = useState(0)
  useEffect(() => {
    if (!done) return
    const id = setTimeout(() => {
      let n = 0
      const count = (b: Block) => {
        if (b.t === 'p') n += runsText(b.runs as RRun[]).split(/\s+/).filter((w) => /[\p{L}\p{N}]/u.test(w)).length
        else if (b.t === 'tbl') for (const r of b.rows) for (const c of r.cells) c.blocks.forEach(count)
      }
      st.blocks.forEach(count)
      setWords(n)
    }, 600)
    return () => clearTimeout(id)
  }, [version, done, st])
  const status = done?.error
    ? `Error: ${done.error}`
    : complete
      ? `Page ${cur} of ${pages.length} · ${words.toLocaleString()} words`
      : `Page ${cur} of ${pages.length}+ · loading…`

  // ---- editing wiring ----
  useEffect(() => {
    const el = pagesRef.current
    if (!isEditing || !el) return
    editor.input.attach(el)
    editor.onDirty = () => viewer.setDirty(true)
    return () => editor.input.detach()
  }, [isEditing, editor, viewer])

  useEffect(() => {
    editor.shortcut = (e: KeyboardEvent) => {
      const mod = e.metaKey || e.ctrlKey
      if (mod && e.key.toLowerCase() === 's') {
        void viewer.save()
        return true
      }
      if (mod && e.key.toLowerCase() === 'f') {
        setFind((f) => ({ ...f, open: true }))
        return true
      }
      if (mod && e.key === 'Enter') {
        editor.insertPageBreak()
        return true
      }
      if (mod && e.key.toLowerCase() === 'k') {
        askLink()
        return true
      }
      // Ctrl/⌘+Alt+0…6: Normal text, Heading 1…6
      if (mod && e.altKey && /^Digit[0-6]$/.test(e.code)) {
        const n = Number(e.code.slice(5))
        void editor.applyCanon(n ? (`h${n}` as 'h1') : 'normal')
        return true
      }
      if (mod && e.key === '\\') {
        editor.clearFormat()
        return true
      }
      if (e.key === 'Escape' && editor.painter) {
        editor.copyFormat()
        return true
      }
      return false
    }
  }, [editor, viewer]) // eslint-disable-line react-hooks/exhaustive-deps

  useIsoLayoutEffect(() => {
    const sel = editor.pendingSel
    if (!sel || !pagesRef.current) return
    if (!editor.input.restore(sel)) {
      // caret paragraph not rendered (scrolled away): bring its page into view, retry next render
      const bi = Number(sel.f.pa.split('.')[0])
      const pi = pages.findIndex((p) => bi >= p.start && bi < p.end)
      if (pi >= 0 && scrollRef.current) scrollRef.current.scrollTop = Math.max(0, tops[pi] - 16)
      if (pi < 0 || (pi >= first && pi <= last)) editor.pendingSel = null
      return
    }
    editor.pendingSel = null
    // keep the caret visible
    const ds = getSel(pagesRef.current)
    const range = ds?.rangeCount ? ds.getRangeAt(0) : null
    const sc = scrollRef.current
    if (range && sc) {
      const r = range.getBoundingClientRect()
      const box = sc.getBoundingClientRect()
      if (r.height && r.bottom > box.bottom - 24) sc.scrollTop += r.bottom - box.bottom + 48
      else if (r.height && r.top < box.top + 8) sc.scrollTop -= box.top - r.top + 48
    }
  })

  // ---- find ----
  const matches = useMemo<Match[]>(() => {
    const q = find.q.toLowerCase()
    if (!find.open || !q) return []
    const out: Match[] = []
    const scan = (p: Paragraph, pa: string, bi: number) => {
      const t = runsText(p.runs as RRun[]).toLowerCase()
      for (let i = t.indexOf(q); i >= 0 && out.length < 5000; i = t.indexOf(q, i + q.length)) out.push({ pa, a: i, b: i + q.length, bi })
    }
    st.blocks.forEach((b, bi) => {
      if (editable(b)) scan(b, String(bi), bi)
      else if (b.t === 'tbl')
        b.rows.forEach((r, ri) => r.cells.forEach((c, ci) => c.blocks.forEach((x, pi) => editable(x) && scan(x, `${bi}.${ri}.${ci}.${pi}`, bi))))
    })
    return out
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [find.q, find.open, version])

  const gotoMatch = (idx: number) => {
    const m = matches[idx]
    const sc = scrollRef.current
    if (!m || !sc) return
    const pi = pages.findIndex((p) => m.bi >= p.start && m.bi < p.end)
    if (pi >= 0) sc.scrollTop = Math.max(0, tops[pi] - 16)
    setFind((f) => ({ ...f, idx }))
    requestAnimationFrame(() =>
      requestAnimationFrame(() => {
        const root = pagesRef.current
        if (!root) return
        const a = domFromPos(root, { pa: m.pa, o: m.a })
        const b = domFromPos(root, { pa: m.pa, o: m.b })
        if (!a || !b) return
        const r = document.createRange()
        r.setStart(a[0], a[1])
        r.setEnd(b[0], b[1])
        const rect = r.getBoundingClientRect()
        const box = sc.getBoundingClientRect()
        sc.scrollTop += rect.top - box.top - box.height / 3
        if (isEditing) editor.input.restore({ a: { pa: m.pa, o: m.a }, f: { pa: m.pa, o: m.b } }, false)
      }),
    )
  }

  // highlight visible matches (CSS Custom Highlight API)
  useIsoLayoutEffect(() => {
    const H = (globalThis as unknown as { Highlight?: new (...r: Range[]) => unknown }).Highlight
    const reg = (CSS as unknown as { highlights?: Map<string, unknown> }).highlights
    if (!H || !reg) return
    const root = pagesRef.current
    if (!root || !matches.length) {
      reg.delete('de-find')
      reg.delete('de-find-cur')
      return
    }
    const all: Range[] = []
    let curR: Range | undefined
    matches.forEach((m, i) => {
      if (m.bi < (pages[first]?.start ?? 0) || m.bi >= (pages[last]?.end ?? 0)) return
      const a = domFromPos(root, { pa: m.pa, o: m.a })
      const b = domFromPos(root, { pa: m.pa, o: m.b })
      if (!a || !b) return
      const r = document.createRange()
      r.setStart(a[0], a[1])
      r.setEnd(b[0], b[1])
      if (i === find.idx) curR = r
      else all.push(r)
    })
    reg.set('de-find', new H(...all))
    if (curR) reg.set('de-find-cur', new H(curR))
    else reg.delete('de-find-cur')
  })

  const replaceMatch = (m: Match, text: string) => opInsert(editor, { a: { pa: m.pa, o: m.a }, f: { pa: m.pa, o: m.b } }, text)

  const docStyle: CSSProperties = { fontFamily: meta?.font, fontSize: meta ? `${meta.sz}pt` : undefined }

  // ---- dialogs / clipboard ----
  const askLink = () => {
    const cur = editor.state()?.link
    setPrompt({
      title: cur ? 'Edit link' : 'Insert link',
      placeholder: 'https://…',
      initial: cur,
      ok: 'Apply',
      onOk: (v) => void editor.setLink(v.trim() || null),
      extra: cur ? { label: 'Remove', run: () => void editor.setLink(null) } : undefined,
    })
  }
  const askAlt = () => {
    const cur = editor.state()?.img
    if (!cur) return
    setPrompt({
      title: 'Alt text',
      label: 'Describe the picture for screen readers',
      initial: cur.alt ?? '',
      area: true,
      ok: 'Save',
      onOk: (v) => editor.setAltText(v.trim()),
    })
  }
  const paste = async (plain: boolean) => {
    const sel = editor.input.sel()
    if (!sel) return
    try {
      if (!plain && navigator.clipboard?.read) {
        for (const it of await navigator.clipboard.read()) {
          const pic = it.types.find((t) => t.startsWith('image/'))
          if (pic) return void editor.insertImages([await it.getType(pic)], sel)
          if (it.types.includes('text/html')) {
            const html = await (await it.getType('text/html')).text()
            const dt = { files: [] as unknown as FileList, getData: (t: string) => (t === 'text/html' ? html : '') } as unknown as DataTransfer
            if (editor.pasteRich(sel, dt)) return
          }
        }
      }
      const text = await navigator.clipboard.readText()
      if (text) opInsert(editor, sel, text)
    } catch {
      setToast(`The browser blocked clipboard access — use ${MOD}V to paste`)
    }
  }
  const copyCut = (cmd: 'copy' | 'cut') => {
    if (!document.execCommand(cmd)) setToast(`Use ${MOD}${cmd === 'copy' ? 'C' : 'X'}`)
  }
  const pickImage = () => void pickFiles().then((f) => void (f.length && editor.insertImages(f)))
  const imgEl = () => (img && pagesRef.current ? (findObj(pagesRef.current, img.pa, img.o) as HTMLImageElement | null) : null)
  const imageActions = {
    original: () => {
      const el = imgEl()
      if (el?.naturalWidth) editor.imageOriginal(el.naturalWidth, el.naturalHeight)
    },
    replace: () => void pickFiles('image/*', false).then((f) => f[0] && editor.replaceImage(f[0])),
    alt: askAlt,
    download: () => {
      const el = imgEl()
      if (!el?.src || !img) return
      void fetch(el.src)
        .then((r) => r.blob())
        .then((b) => download(b, img.target.split('/').pop() ?? 'image'))
    },
  }

  // ---- image selection overlay (resize handles) ----
  useIsoLayoutEffect(() => {
    if (resizing.current) return
    const root = pagesRef.current
    const sc = scrollRef.current
    let box: Box | null = null
    if (isEditing && img && root && sc) {
      const el = findObj(root, img.pa, img.o)
      if (el) {
        const r = el.getBoundingClientRect()
        const b = sc.getBoundingClientRect()
        box = { x: r.left - b.left + sc.scrollLeft, y: r.top - b.top + sc.scrollTop, w: r.width, h: r.height }
      }
    }
    setImgBox((p) => (p === box || (p && box && p.x === box.x && p.y === box.y && p.w === box.w && p.h === box.h) ? p : box))
  })

  const startResize = (dir: (typeof HANDLES)[number], e: React.MouseEvent) => {
    e.preventDefault()
    e.stopPropagation()
    const el = imgEl()
    if (!img || !imgBox || !el) return
    const { w: w0, h: h0, x: x0, y: y0 } = imgBox
    const ptPerPx = img.w / Math.max(1, w0)
    const sx = e.clientX
    const sy = e.clientY
    let w = w0
    let h = h0
    resizing.current = true
    resizeRatio.current = ptPerPx
    const move = (ev: MouseEvent) => {
      const dx = ev.clientX - sx
      const dy = ev.clientY - sy
      let nw = w0 + (dir.includes('e') ? dx : dir.includes('w') ? -dx : 0)
      let nh = h0 + (dir.includes('s') ? dy : dir.includes('n') ? -dy : 0)
      if (dir.length === 2 && !ev.shiftKey) {
        const k = Math.max(nw / w0, nh / h0)
        nw = w0 * k
        nh = h0 * k
      }
      const maxW = (contentW / ptPerPx) * 1.0
      w = Math.max(8, Math.min(nw, maxW * 2))
      h = Math.max(8, nh)
      el.style.width = `${w * ptPerPx}pt`
      el.style.height = `${h * ptPerPx}pt`
      setImgBox({ x: x0, y: y0, w, h })
    }
    const up = () => {
      window.removeEventListener('mousemove', move)
      window.removeEventListener('mouseup', up)
      resizing.current = false
      resizeRatio.current = null
      if (w !== w0 || h !== h0) editor.setImageSize(w * ptPerPx, h * ptPerPx)
      else setImgBox(null)
    }
    window.addEventListener('mousemove', move)
    window.addEventListener('mouseup', up)
  }

  // ---- context menu ----
  const ctxItems = (): MenuItem[] => {
    const s = editor.state()
    const sel = editor.input.sel()
    const ranged = !!sel && (sel.a.pa !== sel.f.pa || sel.a.o !== sel.f.o)
    const items: MenuItem[] = []
    if (s?.img) {
      items.push(
        { head: 'Image', label: 'Original size', icon: 'imgOrig', run: imageActions.original },
        { label: 'Fit to page width', icon: 'fit', run: () => editor.resizeImage('fit') },
        { label: 'Replace image…', icon: 'imgReplace', run: imageActions.replace },
        { label: 'Alt text…', icon: 'alt', run: askAlt },
        { label: 'Download image', icon: 'save', run: imageActions.download },
        { label: 'Delete image', icon: 'trash', run: () => editor.deleteSelection() },
      )
    }
    items.push(
      { label: 'Cut', icon: 'cut', hint: `${MOD}X`, disabled: !ranged, run: () => copyCut('cut'), sep: items.length > 0 },
      { label: 'Copy', icon: 'copy', hint: `${MOD}C`, disabled: !ranged, run: () => copyCut('copy') },
      { label: 'Paste', icon: 'paste', hint: `${MOD}V`, run: () => void paste(false) },
      { label: 'Paste without formatting', icon: 'pastePlain', hint: `${MOD}⇧V`, run: () => void paste(true) },
      { label: 'Delete', icon: 'trash', disabled: !ranged || !!s?.img, run: () => editor.deleteSelection() },
    )
    if (s?.link) {
      items.push(
        { label: 'Open link', icon: 'external', sep: true, run: () => window.open(s.link, '_blank', 'noopener') },
        { label: 'Edit link…', icon: 'link', hint: `${MOD}K`, run: askLink },
        { label: 'Remove link', icon: 'unlink', run: () => void editor.setLink(null) },
      )
    } else items.push({ label: 'Insert link…', icon: 'link', hint: `${MOD}K`, sep: true, run: askLink, disabled: !s })
    if (s?.inTable) {
      const op = (k: Parameters<DocxEditor['tableOp']>[0]) => () => editor.tableOp(k)
      items.push(
        { head: 'Table', label: 'Insert row above', icon: 'rowAbove', sep: true, run: op('rowAbove') },
        { label: 'Insert row below', icon: 'rowBelow', run: op('rowBelow') },
        { label: 'Insert column left', icon: 'colLeft', run: op('colLeft') },
        { label: 'Insert column right', icon: 'colRight', run: op('colRight') },
        { label: 'Delete row', icon: 'delRow', run: op('delRow') },
        { label: 'Delete column', icon: 'delCol', run: op('delCol') },
        { label: 'Delete table', icon: 'trash', run: op('delTable') },
      )
    } else {
      items.push(
        { label: 'Insert image…', icon: 'image', sep: true, run: pickImage, disabled: !s },
        { label: 'Insert table', icon: 'table', run: () => editor.insertTable(3, 3), disabled: !s },
        { label: 'Page break', icon: 'pagebreak', hint: `${MOD}↵`, run: () => editor.insertPageBreak(), disabled: !s },
      )
    }
    items.push(
      { label: 'Clear formatting', icon: 'clear', hint: `${MOD}\\`, sep: true, run: () => editor.clearFormat(), disabled: !s },
      { label: 'Select all', icon: 'selectAll', hint: `${MOD}A`, run: () => document.execCommand('selectAll') },
    )
    return items
  }

  // ---- outline ----
  const titleId = meta?.styles?.find((x) => x.canon === 'title')?.id
  const headings = useMemo(
    () => (outline ? outlineOf(st.blocks, titleId) : []),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [outline, version, titleId],
  )
  const topBlock = pages[Math.max(0, cur - 1)]?.start ?? 0
  const activeHeading = headings.reduce((acc, h, i) => (h.bi <= Math.max(topBlock, caretBi) ? i : acc), -1)
  const gotoBlock = (bi: number) => {
    const sc = scrollRef.current
    const pi = pages.findIndex((p) => bi >= p.start && bi < p.end)
    if (!sc || pi < 0) return
    sc.scrollTop = Math.max(0, tops[pi] - 16)
    requestAnimationFrame(() =>
      requestAnimationFrame(() => {
        const el = pagesRef.current?.querySelector(`[data-pa="${bi}"]`)
        if (el) sc.scrollTop += el.getBoundingClientRect().top - sc.getBoundingClientRect().top - 24
        if (isEditing) editor.input.restore({ a: { pa: String(bi), o: 0 }, f: { pa: String(bi), o: 0 } })
      }),
    )
  }

  return (
    <div style={{ position: 'relative', flex: 1, minWidth: 0, display: 'flex' }}>
      <div
        ref={scrollRef}
        className="de-scroll"
        tabIndex={-1}
        onKeyDown={(e) => {
          if ((e.metaKey || e.ctrlKey) && e.key.toLowerCase() === 'f') {
            e.preventDefault()
            setFind((f) => ({ ...f, open: true }))
          }
        }}
      >
        {meta && queue.length > 0 && (
          <div ref={measureRef} className="de-measure de-doc" style={{ ...docStyle, width: `${contentW}pt` }}>
            {queue.map(({ b, i }) => (
              <BlockView key={i} block={b} ctx={measureCtx} sp={contextual(b, st.blocks[i - 1], st.blocks[i + 1])} />
            ))}
          </div>
        )}
        <div
          ref={pagesRef}
          className={isEditing ? 'de-pages de-editing' : 'de-pages'}
          style={{ height: totalH, width: totalW }}
          contentEditable={isEditing || undefined}
          suppressContentEditableWarning
          spellCheck={false}
          onMouseDown={(e) => {
            if (!isEditing) return
            // clicking a picture selects it (resize handles, image tools)
            const t = e.target as HTMLElement
            const obj = t.tagName === 'IMG' ? t.closest<HTMLElement>('[data-obj]') : null
            const pos = obj && pagesRef.current ? posFromDom(pagesRef.current, obj, 0) : null
            if (pos) {
              e.preventDefault()
              editor.selectObject(pos.pa, pos.o)
            }
          }}
          onMouseUp={() => {
            if (editor.painter) setTimeout(() => editor.applyPainter(), 0)
          }}
          onContextMenu={(e) => {
            if (!isEditing) return
            e.preventDefault()
            setCtxMenu({ x: e.clientX, y: e.clientY })
          }}
        >
          {pages.map((p, i) =>
            (i < first || i > last) && i !== caretPage ? null : (
              <PageView
                key={i}
                index={i}
                count={pageCount}
                slice={p}
                blocks={st.blocks}
                meta={meta}
                page={page}
                top={tops[i]}
                left={left}
                height={pageHeights[i]}
                zoom={z}
                client={client}
                docStyle={docStyle}
                editing={isEditing}
              />
            ),
          )}
          {skeleton > 0 && <div className="de-page-skeleton" contentEditable={false} style={{ top: acc, left, width: pageW * z, height: pageH * z }} />}
        </div>
        {imgBox && (
          <div className="de-img-sel" style={{ left: imgBox.x, top: imgBox.y, width: imgBox.w, height: imgBox.h }}>
            {HANDLES.map((d) => (
              <span key={d} className={`de-img-h de-img-${d}`} onMouseDown={(e) => startResize(d, e)} />
            ))}
            {resizeRatio.current && (
              <span className="de-img-size">
                {Math.round(imgBox.w * resizeRatio.current)} × {Math.round(imgBox.h * resizeRatio.current)} pt
              </span>
            )}
          </div>
        )}
      </div>
      {outline && (
        <nav className="de-outline" aria-label="Document outline">
          <div className="de-outline-head">
            <span>Outline</span>
            <button type="button" className="de-tb" title="Close outline" aria-label="Close outline" onClick={() => setOutline(false)}>
              ×
            </button>
          </div>
          {headings.length ? (
            headings.map((h, i) => (
              <button key={h.bi} type="button" className={i === activeHeading ? 'de-outline-item on' : 'de-outline-item'} style={{ paddingLeft: 10 + h.lvl * 14 }} onClick={() => gotoBlock(h.bi)}>
                {h.text}
              </button>
            ))
          ) : (
            <div className="de-outline-empty">Headings you add to the document will appear here.</div>
          )}
        </nav>
      )}
      {find.open && (
        <FindBar
          query={find.q}
          index={find.idx}
          count={matches.length}
          editable={isEditing}
          onQuery={(q) => setFind({ open: true, q, idx: 0 })}
          onNav={(d) => matches.length && gotoMatch((find.idx + d + matches.length) % matches.length)}
          onClose={() => setFind({ open: false, q: find.q, idx: 0 })}
          onReplace={(text, all) => {
            if (!matches.length) return
            if (all) [...matches].reverse().forEach((m) => replaceMatch(m, text))
            else replaceMatch(matches[Math.min(find.idx, matches.length - 1)], text)
          }}
        />
      )}
      {isEditing && (
        <DocxToolbar
          editor={editor}
          styles={meta?.styles ?? []}
          onFind={() => setFind((f) => ({ ...f, open: true }))}
          outline={outline}
          onOutline={() => setOutline((o) => !o)}
          painter={painter}
          onLink={askLink}
          image={imageActions}
        />
      )}
      {!isEditing && done && (
        <button type="button" className={outline ? 'de-outline-toggle on' : 'de-outline-toggle'} title="Document outline" onClick={() => setOutline((o) => !o)}>
          ☰
        </button>
      )}
      {ctxMenu && <ContextMenu x={ctxMenu.x} y={ctxMenu.y} items={ctxItems()} onClose={() => setCtxMenu(null)} />}
      {prompt && <PromptDialog spec={prompt} onClose={() => setPrompt(null)} />}
      {toast && <Toast text={toast} onDone={() => setToast(null)} />}
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
  editing: boolean
}

function PageView({ index, count, slice, blocks, meta, page, top, left, height, zoom, client, docStyle, editing }: PageViewProps) {
  const ctx = useMemo<RenderCtx>(() => ({ client, page: index, pages: count, measuring: false, editable: editing }), [client, index, count, editing])
  const hfCtx = useMemo<RenderCtx>(() => ({ ...ctx, editable: false }), [ctx])
  const first = index === 0 && meta?.titlePg
  const header = first ? meta?.firstHeader : meta?.header
  const footer = first ? meta?.firstFooter : meta?.footer
  const items = []
  for (let i = slice.start; i < slice.end; i++) {
    const b = blocks[i]
    items.push(
      <BlockView
        key={`${i}:${b.t === 'p' ? b.ver ?? 0 : 0}`}
        block={b}
        ctx={ctx}
        pa={editable(b) || b.t === 'tbl' ? String(i) : undefined}
        sp={contextual(b, blocks[i - 1], blocks[i + 1])}
      />,
    )
  }
  return (
    <div className="de-page" style={{ top, left, width: page.w * PX * zoom, height }} data-page={index + 1}>
      <div className="de-page-inner de-doc" style={{ ...docStyle, width: `${page.w}pt`, height: height / zoom, zoom }}>
        {header && (
          <div className="de-hf" contentEditable={editing ? false : undefined} style={{ top: `${page.header}pt`, padding: `0 ${page.mr}pt 0 ${page.ml}pt` }}>
            {header.map((b, i) => (
              <BlockView key={i} block={b} ctx={hfCtx} />
            ))}
          </div>
        )}
        <div style={{ position: 'absolute', top: `${page.mt}pt`, left: `${page.ml}pt`, width: `${page.w - page.ml - page.mr}pt` }}>{items}</div>
        {footer && (
          <div className="de-hf" contentEditable={editing ? false : undefined} style={{ bottom: `${page.footer}pt`, padding: `0 ${page.mr}pt 0 ${page.ml}pt` }}>
            {footer.map((b, i) => (
              <BlockView key={i} block={b} ctx={hfCtx} />
            ))}
          </div>
        )}
      </div>
    </div>
  )
}
