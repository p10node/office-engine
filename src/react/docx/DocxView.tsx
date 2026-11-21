import { useCallback, useEffect, useMemo, useRef, useState, type CSSProperties } from 'react'
import type { Block, DocxMeta, HFKind, PageSetup, Paragraph } from '../../core/docx/types'
import { DEFAULT_PAGE } from '../../core/docx/types'
import { registerHost } from '../../collab/hosts'
import { safeHref } from '../../core/url'
import type { DocumentClient } from '../../core/worker/client'
import { claimFocus, lowerIndex, useForceUpdate, useIsoLayoutEffect, useScroll, useSize } from '../common/hooks'
import { useFileOps, useSelectAll, useViewer } from '../common/Slots'
import { Status } from '../common/Status'
import { ZoomControl, type Zoom } from '../common/ZoomControl'
import { useViewCommands, useControllerObserved } from '../controller'
import { applyDocx, docxFormat } from '../edit/format'
import { ContextMenu, MOD, type MenuItem } from '../common/ContextMenu'
import { PromptDialog, Toast, type PromptSpec } from '../common/Dialog'
import { domFromPos, getSel, opInsert, paraRanges, posFromDom, selPoints } from '../edit/input'
import { download, Icon, pickFiles } from '../edit/Toolbar'
import { runsText, type RRun, type Sel } from '../edit/rich'
import { DocxEditor, type CommentRange, type DocState, type FormatState } from './DocxEditor'
import { markCss, useCommentsCtl, useCommentsState, type CommentHost, type PlacedThread } from '../comments/Comments'
import { COMMENT_KEYS } from '../comments/CommentsPanel'
import { DocxToolbar, styleChoices, type ImageActions } from './DocxToolbar'
import { FindBar, type Match } from './FindBar'
import { findAll } from './find'
import { HeaderFooterDialog } from './HeaderFooterDialog'
import { PageSetupDialog } from './PageSetupDialog'
import { mergeSource } from '../../core/plainmerge'
import { docxBodyHtml, docxToHtml, docxToMarkdown, docxToText, pictureData } from '../file/text'
import { BlockView, contextual, type RenderCtx } from './render'
import { columnsOf, headOf, pageAt, pageHF, pagesOf, paginate, pieceOf, type Lines, type PageSlice } from './paginate'
import { lineStarts, paraLines, tableLines } from './measure'
import { formatCount, formatNumber, useStrings } from '../strings'
import { docxFonts, docxPageFonts } from '../../core/fonts'
import { useDocFonts } from '../common/fonts'

const PX = 4 / 3
const GAP = 16
const PAD = 24
// Stay below browser element-height limits, as the spreadsheet view does.
const MAX_SCROLL_PX = 12_000_000

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

/** a paragraph's text without its tracked deletions (word count, outline) */
const shownText = (p: Paragraph) => runsText(p.runs.filter((r) => r.t !== 'text' || r.tc?.k !== 'del') as RRun[])

/** heading entries for the outline panel */
function outlineOf(blocks: Block[], titleId: string | undefined) {
  const out: { bi: number; lvl: number; text: string }[] = []
  blocks.forEach((b, bi) => {
    if (b.t !== 'p' || b.pid === undefined) return
    const title = !!titleId && b.sid === titleId
    const lvl = title ? 0 : b.ps.outline
    if (lvl === undefined || lvl > 5) return
    const text = shownText(b).trim()
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
  const str = useStrings()
  const st = useRef<DocState & { lines: WeakMap<Block, Lines> }>({ blocks: [], heights: new WeakMap(), lines: new WeakMap(), dirty: new Set(), frontier: 0 }).current
  const [version, bump] = useForceUpdate()
  const [meta, setMeta] = useState<DocxMeta | null>(null)
  const [pageDlg, setPageDlg] = useState(false)
  /** header / footer dialog, opened on that part */
  const [hfDlg, setHfDlg] = useState<HFKind | null>(null)
  const openHF = useCallback((k: HFKind) => setHfDlg(k), [])
  const [done, setDone] = useState<{ error?: string } | null>(null)
  const [zoom, setZoom] = useState<Zoom>(initialZoom)
  const [scrollRef, size] = useSize<HTMLDivElement>()
  const scroll = useScroll(scrollRef)
  const pagesRef = useRef<HTMLDivElement>(null)
  const [domSel, setDomSel] = useState<number[]>([])
  const editor = useMemo(() => new DocxEditor(client, st, bump), [client, st, bump])
  const colResize = useCallback((bi: number, cols: number[]) => editor.setColumnWidths(bi, cols), [editor])
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
    editor.onHeaders = (h) => setMeta((m) => (m ? { ...m, ...h } : m))
    editor.onPage = (pg) => setMeta((m) => (m ? { ...m, page: pg } : m))
    editor.onHF = (r) => setMeta((m) => (m ? { ...m, header: undefined, footer: undefined, firstHeader: undefined, firstFooter: undefined, ...r.headers, titlePg: r.titlePg } : m))
    editor.onPainter = setPainter
    return () => {
      editor.onImage = editor.onStyles = editor.onHeaders = editor.onPainter = editor.onPage = editor.onHF = undefined
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
        // a page setup made before the file finished loading (a collaborator's, replayed) wins over the file's
        let next = editor.pageNow ? { ...m, page: editor.pageNow } : m
        if (editor.hfNow) next = { ...next, header: undefined, footer: undefined, firstHeader: undefined, firstFooter: undefined, ...editor.hfNow.headers, titlePg: editor.hfNow.titlePg }
        setMeta(next)
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
  const cols = page.cols && page.cols > 1 ? page.cols : 1
  const colGap = cols > 1 ? (page.colGap ?? 36) : 0
  // text is laid out (and measured) one column wide
  const contentW = (page.w - page.ml - page.mr - colGap * (cols - 1)) / cols
  const availPx = (page.h - page.mt - page.mb) * PX

  // ---- measurement (see below, once the pages shown are known) ----
  const [batch, setBatch] = useState(40)
  /** printing waits until the line starts of every paragraph a page cuts are read */
  const printWait = useRef<(() => void) | null>(null)
  /** the pages are laid out again before this render is painted: what depends on where things are waits for that */
  const relayout = useRef(false)
  const measureRef = useRef<HTMLDivElement>(null)
  const widthKey = `${contentW}|${meta?.font}|${meta?.sz}`
  const lastKey = useRef(widthKey)
  if (lastKey.current !== widthKey) {
    lastKey.current = widthKey
    st.heights = new WeakMap()
    st.lines = new WeakMap()
    st.dirty.clear()
    st.frontier = 0
  }
  const total = st.blocks.length
  // measured first: edited blocks
  const queue: { b: Block; i: number; starts?: boolean }[] = []
  const queued = new Set<Block>()
  const enqueue = (b: Block, i: number, starts?: boolean) => {
    if (queued.has(b)) return
    queued.add(b)
    queue.push({ b, i, starts })
  }
  for (const b of st.dirty) {
    if (queue.length >= 50) break
    const i = st.blocks.indexOf(b)
    if (i < 0) st.dirty.delete(b)
    else enqueue(b, i, true)
  }
  const edited = queue.length
  // A few edited blocks (typing) are measured again before this render is painted, and the pages cut again then (see
  // the measure effect): until that, the pages stay as they were cut. Else they are cut again on every update.
  const soon = edited > 0 && edited <= 10
  const cut = useRef<{ version: number; key: string; total: number; pages: PageSlice[] } | null>(null)
  const cutKey = `${availPx}|${cols}|${meta?.final}|${meta?.sectKind}`
  const c0 = cut.current
  let laid: PageSlice[]
  if (c0 && c0.key === cutKey && c0.total === total && (c0.version === version || soon)) laid = c0.pages
  else {
    laid = paginate(st.blocks, st.heights, Math.min(st.frontier, total), availPx, cols, (meta?.final && meta.sectKind) || 'nextPage', st.lines)
    cut.current = { version, key: cutKey, total, pages: laid }
  }
  // while an input method composes, the pages stay as they are: a paragraph cut elsewhere would be rebuilt under it
  const frozen = useRef<{ pages: PageSlice[]; total: number } | null>(null)
  const composing = isEditing && editor.input.isComposing
  const pages = composing && frozen.current?.total === total ? frozen.current.pages : laid
  if (!composing) frozen.current = { pages, total }

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
  const spacerH = Math.min(totalH, MAX_SCROLL_PX)
  const maxLogicalY = Math.max(0, totalH - size.h)
  const maxScrollY = Math.max(0, spacerH - size.h)
  const scrollScale = maxScrollY > 0 ? Math.max(1, maxLogicalY / maxScrollY) : 1
  const scrollY = Math.min(maxLogicalY, scroll.y * scrollScale)
  const scrollOffset = scrollY - scroll.y
  const scrollToY = (y: number) => {
    if (scrollRef.current) scrollRef.current.scrollTop = Math.max(0, Math.min(maxLogicalY, y)) / scrollScale
  }
  /** the logical position (in `tops`) of a physical `scrollTop` */
  const logicalY = (top: number) => Math.min(maxLogicalY, top * scrollScale)
  const scrollByY = (delta: number) => {
    if (scrollRef.current) scrollRef.current.scrollTop += delta / scrollScale
  }
  const renderedScroll = useRef(scroll.y)
  useIsoLayoutEffect(() => { renderedScroll.current = scroll.y })
  // A native scroll changes the viewport before useScroll's frame-batched
  // update repositions the pages. Measure a target only after that commit.
  const afterScroll = (run: () => void) => {
    let frames = 0
    const tick = () => {
      const sc = scrollRef.current
      if (!sc || ++frames > 120) return
      if (Math.abs(sc.scrollTop - renderedScroll.current) > 1) requestAnimationFrame(tick)
      else run()
    }
    requestAnimationFrame(tick)
  }
  const totalW = Math.max(size.w, pageW * z + PAD * 2)
  const left = Math.max(PAD, (totalW - pageW * z) / 2)

  const first = tops.length ? lowerIndex(tops, scrollY - pageH * z) : 0
  const last = tops.length ? lowerIndex(tops, scrollY + size.h + 200) : -1
  // measured again for a font: the first block in view stays where it is on screen, until the measuring is done or the
  // reader scrolls. Found by the model (blocks stacked from the top margin, a block begun on the page before by the part
  // of it on this page): the text on screen is already in the new font while the pages are not. Put back by what is
  // drawn once it is: the model's heights are measured at zoom 1, and a page at another zoom rounds its lines a little
  // differently, which adds up over the blocks above
  const anchor = useRef<{ gen: number; bi: number; dy: number; y: number } | null>(null)
  useIsoLayoutEffect(() => {
    const sc = scrollRef.current
    const root = pagesRef.current
    const a = anchor.current
    if (!sc || !root || !tops.length) return
    /** the height block `bi` takes on page `pi`: one begun on the page before, from the cut (its space after kept) */
    const partH = (pi: number, bi: number) => {
      const b = st.blocks[bi]
      const h = st.heights.get(b) ?? 0
      const from = bi === pages[pi].start ? pages[pi].from : 0
      const l = from ? st.lines.get(b) : undefined
      return l ? h - l.ys[from] : h
    }
    const blockTop = (pi: number, bi: number) => {
      let y = tops[pi] + page.mt * PX * z
      for (let i = pages[pi].start; i < bi; i++) y += partH(pi, i) * z
      return y
    }
    if (again.current.next === 0 && a?.gen !== fonts.gen) {
      // (nothing measured again yet: these are the pages the reader sees)
      const y = logicalY(sc.scrollTop)
      const pi = lowerIndex(tops, y)
      let bi = pages[pi].start
      while (bi < pages[pi].end - 1 && blockTop(pi, bi + 1) <= y) bi++
      anchor.current = { gen: fonts.gen, bi, dy: blockTop(pi, bi) - y, y }
    } else if (a && Math.abs(sc.scrollTop - a.y) <= 1) {
      const pi = pages.findIndex((p) => a.bi >= p.start && a.bi < p.end)
      // the first part of the block as drawn; not there while where its lines start is read, or when its page is out of view
      const el = root.querySelector<HTMLElement>(`[data-pa="${a.bi}"]:not([data-from])`)
      if (el) scrollByY(el.getBoundingClientRect().top - sc.getBoundingClientRect().top - a.dy)
      else if (pi >= 0) scrollToY(blockTop(pi, a.bi) - a.dy)
      a.y = sc.scrollTop
      // (released once the block is placed by what is drawn)
      if (again.current.next < 0 && el) anchor.current = null
    } else anchor.current = null
  })
  const pageCount = complete ? pages.length : 0
  // Both ends of the selection stay mounted, so the DOM selection survives scrolling. A removed page moves its end of
  // the range to the pages container, and a page mounted again there lands after that point: outside a range that ends on it.
  const pinned = new Set<number>(domSel)
  const lastSel = isEditing ? editor.input.lastSel : null
  const caretBi = lastSel ? Number(lastSel.f.pa.split('.')[0]) : -1
  /** the page showing a place: of a block split across pages, the one holding that line / row */
  const pageOfPos = (pa: string, o = 0) => {
    const [bi, ri] = pa.split('.').map(Number)
    const [a, b] = pagesOf(pages, bi)
    if (a === b) return a
    if (ri === undefined) return pageAt(pages, bi, o, st.lines.get(st.blocks[bi]))
    for (let pi = b; pi > a; pi--) if (ri >= pieceOf(pages[pi], bi).from) return pi
    return a
  }
  // (an offset where a page cuts a paragraph is shown at the end of one page and the start of the next)
  for (const p of lastSel ? [lastSel.a, lastSel.f] : []) pinned.add(pageOfPos(p.pa, p.o)).add(pageOfPos(p.pa, Math.max(0, p.o - 1)))
  /** where the pages cut the paragraphs holding the selection: the caret is put back when that changes */
  const selCuts = lastSel
    ? [lastSel.a.pa, lastSel.f.pa]
        .map((pa) => {
          const bi = Number(pa.split('.')[0])
          const [a, b] = pagesOf(pages, bi)
          const at = st.lines.get(st.blocks[bi])?.at
          let sig = `${bi}`
          for (let pi = a; pi >= 0 && pi < b; pi++) sig += `:${pages[pi].to === undefined ? '' : at?.[pages[pi].to!]}`
          return sig
        })
        .join('|')
    : ''
  const lastCuts = useRef(selCuts)

  // then the paragraphs a page shown ends in (where their lines start is read), then the not-yet-measured tail
  // (time-sliced)
  /** the paragraphs page `pi` cuts whose line starts are not read yet */
  const unread = (pi: number, out: Map<Block, number>) => {
    for (const c of pages[pi] ? columnsOf(pages[pi]) : []) {
      for (const bi of c.end - 1 > c.start ? [c.start, c.end - 1] : [c.start]) {
        const b = st.blocks[bi]
        if (b?.t !== 'p') continue
        const { from, to } = pieceOf(c, bi)
        const at = st.lines.get(b)?.at
        if ((from && at?.[from] === undefined) || (to !== undefined && at?.[to] === undefined)) out.set(b, bi)
      }
    }
  }
  const shownUnread = new Map<Block, number>()
  for (let pi = first; pi <= last; pi++) unread(pi, shownUnread)
  for (const pi of pinned) unread(pi, shownUnread)
  for (const [b, i] of shownUnread) enqueue(b, i, true)
  // printing: every page
  const allUnread = new Map<Block, number>()
  if (printWait.current) for (let pi = 0; pi < pages.length && allUnread.size < 200; pi++) unread(pi, allUnread)
  for (const [b, i] of allUnread) enqueue(b, i, true)
  const tailEnd = Math.min(total, st.frontier + batch)
  for (let i = st.frontier; i < tailEnd; i++) enqueue(st.blocks[i], i)
  // ---- fonts: blocks are measured once the host's fonts they are set in are ready (loaded, failed or timed out). A
  // font that arrives later (one that timed out, the page's own web font) has the measured blocks measured again in
  // slices, their old heights and lines kept meanwhile: the pages stay, and so do the caret and the page in view ----
  const fonts = useDocFonts()
  const again = useRef({ gen: fonts.gen, next: -1 })
  if (again.current.gen !== fonts.gen) again.current = { gen: fonts.gen, next: st.frontier > 0 ? 0 : -1 }
  const againFrom = again.current.next
  const againTo = st.frontier
  const againEnd = againFrom < 0 ? -1 : Math.min(againTo, againFrom + batch)
  // (a paragraph a page cuts has where its lines start read again with it, so the cut does not wait for a second pass)
  for (let i = Math.max(0, againFrom); i < againEnd; i++) if (!st.dirty.has(st.blocks[i])) enqueue(st.blocks[i], i, st.lines.get(st.blocks[i])?.at !== undefined)
  const fontsReady = !queue.length || !meta || fonts.ready(docxFonts(queue.map((q) => q.b), docxPageFonts(meta)))
  const renderStart = performance.now()
  const measureCtx = useMemo<RenderCtx>(() => ({ client, page: -1, pages: 0, measuring: true, editable: false }), [client])

  useIsoLayoutEffect(() => {
    const el = measureRef.current
    if (!el || !queue.length || !meta || !fontsReady) return
    const res = queue.map((q, k) => {
      const c = el.children[k] as HTMLElement | undefined
      let l: Lines | undefined
      if (c && q.b.t === 'tbl') l = tableLines(c, q.b)
      else if (c && q.b.t === 'p') {
        l = paraLines(c)
        if (q.starts && l.ys.length > 2) {
          const at = lineStarts(c, l.ys)
          // lines that cannot be told apart: the paragraph is not cut
          l = at ? { ys: l.ys, at } : { ys: [l.ys[0], l.ys[l.ys.length - 1]] }
        }
      }
      return { h: c?.getBoundingClientRect().height ?? 0, l }
    })
    const dt = performance.now() - renderStart
    const save = () => {
      queue.forEach((q, k) => {
        st.heights.set(q.b, res[k].h)
        if (res[k].l) st.lines.set(q.b, res[k].l!)
        else st.lines.delete(q.b)
        st.dirty.delete(q.b)
      })
      if (tailEnd > st.frontier) st.frontier = tailEnd
      if (againEnd >= 0) again.current.next = againEnd >= againTo ? -1 : againEnd
      setBatch((b) => (dt < 10 ? Math.min(b * 2, 1500) : dt > 30 ? Math.max(10, b >> 1) : b))
      bump()
    }
    // a page shown waits for these, and a few edited blocks (typing) are laid out again before anything is painted: no
    // page shows them where they were (the caret is put back once, after that)
    if (shownUnread.size || soon) {
      relayout.current = true
      return void save()
    }
    const id = setTimeout(save, 0)
    return () => clearTimeout(id)
  })
  useEffect(() => {
    if (!printWait.current || allUnread.size) return
    const done = printWait.current
    printWait.current = null
    done()
  })

  editor.contentW = contentW
  editor.defaultFont = meta?.font
  editor.defaultSz = meta?.sz
  const cur = tops.length ? lowerIndex(tops, scrollY + size.h / 3) + 1 : 0
  // word count (debounced, off the typing path)
  const [words, setWords] = useState(0)
  useEffect(() => {
    if (!done) return
    const id = setTimeout(() => {
      let n = 0
      const count = (b: Block) => {
        if (b.t === 'p') n += shownText(b).split(/\s+/).filter((w) => /[\p{L}\p{N}]/u.test(w)).length
        else if (b.t === 'tbl') for (const r of b.rows) for (const c of r.cells) c.blocks.forEach(count)
      }
      st.blocks.forEach(count)
      setWords(n)
    }, 600)
    return () => clearTimeout(id)
  }, [version, done, st])
  const status = done?.error
    ? str.common.error(done.error)
    : complete
      ? str.docx.status(cur, pages.length, words, str.locale)
      : str.docx.statusLoading(cur, pages.length, str.locale)

  // the pages holding the ends of a ranged DOM selection (viewing has no model selection to tell them)
  useEffect(() => {
    const pageOf = (root: HTMLElement, n: Node | null) => {
      const el = n && root.contains(n) ? (n instanceof Element ? n : n.parentElement)?.closest<HTMLElement>('.de-page') : null
      return el ? Number(el.dataset.page) - 1 : -1
    }
    const onSel = () => {
      const root = pagesRef.current
      const s = root && selPoints(root)
      const ranged = !!s && (s.anchorNode !== s.focusNode || s.anchorOffset !== s.focusOffset)
      const next = root && ranged ? [pageOf(root, s.anchorNode), pageOf(root, s.focusNode)].filter((i) => i >= 0) : []
      setDomSel((cur) => (next.length === cur.length && next.every((x, i) => x === cur[i]) ? cur : next))
    }
    document.addEventListener('selectionchange', onSel)
    return () => document.removeEventListener('selectionchange', onSel)
  }, [])

  // ---- editing wiring ----
  useEffect(() => {
    const el = pagesRef.current
    if (!isEditing || !el) { editor.input.clearSelection(); return }
    editor.input.attach(el)
    return () => editor.input.detach()
  }, [isEditing, editor])

  // Host callback updates must not detach the input and cancel a queued selection update.
  const setDirty = viewer.setDirty
  useEffect(() => {
    if (isEditing) editor.onDirty = () => setDirty(true)
  }, [editor, isEditing, setDirty])

  // Ctrl/⌘+A: every paragraph when editing, the rendered pages' text when viewing (not the viewer around them)
  useSelectAll(() => {
    if (isEditing) editor.selectAll()
    else if (pagesRef.current) getSel(pagesRef.current)?.selectAllChildren(pagesRef.current)
  })

  // keyboard first: shortcuts and typing work without a click (a focused host page control keeps the focus)
  const rendered = pages.some((p) => p.end > p.start)
  useEffect(() => {
    if (!isEditing) {
      claimFocus(scrollRef.current)
      return
    }
    // typing needs a caret: at the start of the document once its first page has content
    if (!rendered || !claimFocus(pagesRef.current) || editor.input.readSel()) return
    const all = editor.all()
    if (all && !editor.input.restore({ a: all.a, f: all.a })) editor.pendingSel = { a: all.a, f: all.a }
  }, [isEditing, rendered]) // eslint-disable-line react-hooks/exhaustive-deps

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
    if (relayout.current) {
      relayout.current = false
      return
    }
    const cutsMoved = lastCuts.current !== selCuts
    lastCuts.current = selCuts
    const sel = editor.pendingSel
    if (!sel || !pagesRef.current) {
      // a paragraph that moved to another page took the caret with it: it goes back to where it was; so does a caret
      // in a paragraph now cut elsewhere by the end of a page (its text moved from one part to the other)
      if (isEditing && !editor.input.repair() && cutsMoved) editor.input.reflow()
      return
    }
    const pi = pageOfPos(sel.f.pa, sel.f.o)
    if (pi >= 0 && (pi < first || pi > last)) {
      scrollToY(tops[pi] - 16)
      return
    }
    const noRefocus = editor.pendingNoRefocus
    editor.pendingNoRefocus = false
    if (!editor.input.restore(sel, noRefocus ? 'preserve' : true)) {
      // caret paragraph not rendered (scrolled away): bring its page into view, retry next render
      if (pi >= 0) scrollToY(tops[pi] - 16)
      if (pi < 0 || (pi >= first && pi <= last)) editor.pendingSel = null
      return
    }
    editor.pendingSel = null
    if (noRefocus) return
    // keep the caret visible
    const ds = getSel(pagesRef.current)
    const range = ds?.rangeCount ? ds.getRangeAt(0) : null
    const sc = scrollRef.current
    if (range && sc) {
      const r = range.getBoundingClientRect()
      const box = sc.getBoundingClientRect()
      if (r.height && r.bottom > box.bottom - 24) scrollByY(r.bottom - box.bottom + 48)
      else if (r.height && r.top < box.top + 8) scrollByY(-(box.top - r.top + 48))
    }
  })

  // ---- find ----
  const matches = useMemo<Match[]>(() => {
    const q = find.q
    if (!find.open || !q) return []
    const out: Match[] = []
    const scan = (p: Paragraph, pa: string, bi: number) => {
      for (const [a, b] of findAll(runsText(p.runs as RRun[]), q, 5000 - out.length)) out.push({ pa, a, b, bi })
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
    const pi = pageOfPos(m.pa, m.a)
    if (pi >= 0) scrollToY(tops[pi] - 16)
    setFind((f) => ({ ...f, idx }))
    afterScroll(() => {
      const root = pagesRef.current
      const r = root && paraRanges(root, m.pa, m.a, m.b)[0]
      if (!r) return
      const rect = r.getBoundingClientRect()
      const box = sc.getBoundingClientRect()
      scrollByY(rect.top - box.top - box.height / 3)
      if (isEditing) editor.input.restore({ a: { pa: m.pa, o: m.a }, f: { pa: m.pa, o: m.b } }, false)
    })
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
    let curR: Range[] | undefined
    matches.forEach((m, i) => {
      if (m.bi < (pages[first]?.start ?? 0) || m.bi >= (pages[last]?.end ?? 0)) return
      // a match over the end of a page: one range on each page
      const rs = paraRanges(root, m.pa, m.a, m.b)
      if (i === find.idx) curR = rs
      else all.push(...rs)
    })
    reg.set('de-find', new H(...all))
    if (curR?.length) reg.set('de-find-cur', new H(...curR))
    else reg.delete('de-find-cur')
  })

  const replaceMatch = (m: Match, text: string) => opInsert(editor, { a: { pa: m.pa, o: m.a }, f: { pa: m.pa, o: m.b } }, text)

  const docStyle: CSSProperties = { fontFamily: meta?.font, fontSize: meta ? `${meta.sz}pt` : undefined }

  // ---- File → Download as / Print: the pages as laid out on screen ----
  useFileOps({
    printable: async () => {
      if (!complete) throw new Error(str.common.stillLoading)
      await new Promise<void>((resolve) => {
        printWait.current = resolve
        bump()
      })
      return {
        size: `${page.w}pt ${page.h}pt`,
        margin: '0',
        node: pages.map((p, i) => {
          // a page whose content is taller than the sheet (a table row taller than a page) is as tall as it is, and the
          // browser cuts it into sheets: nothing is clipped at the sheet's edge
          const h = Math.max(pageH, (page.mt + page.mb) * PX + p.used)
          const tall = h > pageH
          return (
            <div key={i} className={tall ? 'de-print-page de-print-tall' : 'de-print-page'} style={{ width: `${page.w}pt`, height: tall ? `${h}px` : `${page.h}pt` }}>
              <PageView index={i} count={pages.length} slice={p} blocks={st.blocks} meta={meta} page={page} top={0} left={0} height={h} zoom={1} client={client} docStyle={docStyle} editing={false} lines={st.lines} avail={availPx} />
            </div>
          )
        }),
      }
    },
    text: async (format, base) => {
      if (!done) throw new Error(str.common.stillLoading)
      if (format === 'txt') return docxToText(st.blocks)
      const images = await pictureData(st.blocks, (t) => client.call('part', t) as Promise<Blob>)
      // a Markdown / HTML file edited here: what was not edited stays as the file has it
      const kept = base === undefined || (format !== 'html' && format !== 'md') ? null : mergeSource(base, format === 'html' ? 'html' : 'markdown', st.blocks, (bs) => (format === 'html' ? docxBodyHtml(bs, images) : docxToMarkdown(bs, images)))
      if (kept !== null) return kept
      return format === 'html' ? docxToHtml(st.blocks, { title: (client.name ?? '').replace(/\.[^.]+$/, ''), images }) : docxToMarkdown(st.blocks, images)
    },
    stats: () => {
      let chars = 0
      let paras = 0
      const count = (b: Block) => {
        if (b.t === 'p') {
          const t = runsText(b.runs as RRun[])
          chars += t.length
          if (t.trim()) paras++
        } else if (b.t === 'tbl') for (const r of b.rows) for (const c of r.cells) c.blocks.forEach(count)
      }
      st.blocks.forEach(count)
      const more = complete ? '' : '+'
      return [
        [str.docx.pages, formatCount(str, pages.length) + more],
        [str.docx.words, formatCount(str, words) + more],
        [str.docx.characters, formatCount(str, chars) + more],
        [str.docx.paragraphs, formatCount(str, paras) + more],
        [str.docx.pageSize, str.format.inches(formatNumber(str, +(page.w / 72).toFixed(2)), formatNumber(str, +(page.h / 72).toFixed(2)))],
      ]
    },
  })

  // ---- dialogs / clipboard ----
  const askLink = () => {
    // the selection now: the dialog takes the focus
    const at = editor.input.sel()
    const cur = editor.state(at)?.link
    setPrompt({
      title: cur ? str.edit.editLink : str.edit.link,
      placeholder: str.edit.linkPlaceholder,
      initial: cur,
      ok: str.common.apply,
      onOk: (v) => void editor.setLink(v.trim() || null, at).catch((e: Error) => setToast(e.message)),
      extra: cur ? { label: str.common.remove, run: () => void editor.setLink(null, at) } : undefined,
    })
  }
  const askAlt = () => {
    const cur = editor.state()?.img
    if (!cur) return
    setPrompt({
      title: str.edit.altText,
      label: str.edit.altLabel,
      initial: cur.alt ?? '',
      area: true,
      ok: str.common.save,
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
      setToast(str.edit.clipboardBlocked(`${MOD}V`))
    }
  }
  const copyCut = (cmd: 'copy' | 'cut') => {
    if (!document.execCommand(cmd)) setToast(str.edit.useKeys(`${MOD}${cmd === 'copy' ? 'C' : 'X'}`))
  }
  const pickImage = () => void pickFiles().then((f) => void (f.length && editor.insertImages(f)))
  const imgEl = () => (img && pagesRef.current ? (findObj(pagesRef.current, img.pa, img.o) as HTMLImageElement | null) : null)
  const imageActions: ImageActions = {
    original: () => {
      const el = imgEl()
      if (el?.naturalWidth) editor.imageOriginal(el.naturalWidth, el.naturalHeight)
    },
    replace: () => void pickFiles('image/*', false).then((f) => f[0] && editor.replaceImage(f[0])),
    alt: askAlt,
    download: !viewer.canDownload
      ? undefined
      : () => {
          if (!img) return
          // from the file, not from the page: an SVG is shown from a data: URL, which a host's content policy may not let be fetched
          void (client.call('part', img.target) as Promise<Blob>).then((b) => download(b, img.target.split('/').pop() ?? 'image'))
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
    if (cm?.canComment) items.push({ label: str.edit.comment, icon: 'commentAdd', hint: COMMENT_KEYS, run: () => cm.start(), disabled: !s })
    if (s?.img) {
      items.push(
        { head: str.docx.imageHead, label: str.docx.originalSize, icon: 'imgOrig', run: imageActions.original },
        { label: str.docx.fitWidthMenu, icon: 'fit', run: () => editor.resizeImage('fit') },
        { label: str.docx.replaceImageMenu, icon: 'imgReplace', run: imageActions.replace },
        { label: str.edit.altTextMenu, icon: 'alt', run: askAlt },
        ...(imageActions.download ? [{ label: str.docx.downloadImage, icon: 'save', run: imageActions.download } as MenuItem] : []),
        { label: str.docx.deleteImage, icon: 'trash', run: () => editor.deleteSelection() },
      )
    }
    items.push(
      { label: str.edit.cut, icon: 'cut', hint: `${MOD}X`, disabled: !ranged, run: () => copyCut('cut'), sep: items.length > 0 },
      { label: str.edit.copy, icon: 'copy', hint: `${MOD}C`, disabled: !ranged, run: () => copyCut('copy') },
      { label: str.edit.paste, icon: 'paste', hint: `${MOD}V`, run: () => void paste(false) },
      { label: str.edit.pastePlain, icon: 'pastePlain', hint: `${MOD}⇧V`, run: () => void paste(true) },
      { label: str.common.delete, icon: 'trash', disabled: !ranged || !!s?.img, run: () => editor.deleteSelection() },
    )
    if (s?.link) {
      // the target comes from the file (or a peer): opened only when it is a web, mail or phone link, like a click on it
      const href = safeHref(s.link)
      items.push(
        { label: str.edit.openLink, icon: 'external', sep: true, disabled: !href, run: () => href && window.open(href, '_blank', 'noopener,noreferrer') },
        { label: str.edit.editLinkMenu, icon: 'link', hint: `${MOD}K`, run: askLink },
        { label: str.edit.removeLink, icon: 'unlink', run: () => void editor.setLink(null) },
      )
    } else items.push({ label: str.edit.insertLinkMenu, icon: 'link', hint: `${MOD}K`, sep: true, run: askLink, disabled: !s })
    if (s?.inTable) {
      const op = (k: Parameters<DocxEditor['tableOp']>[0]) => () => editor.tableOp(k)
      items.push(
        { head: str.table.menu, label: str.table.rowAbove, icon: 'rowAbove', sep: true, run: op('rowAbove') },
        { label: str.table.rowBelow, icon: 'rowBelow', run: op('rowBelow') },
        { label: str.table.colLeft, icon: 'colLeft', run: op('colLeft') },
        { label: str.table.colRight, icon: 'colRight', run: op('colRight') },
        { label: str.table.delRow, icon: 'delRow', run: op('delRow') },
        { label: str.table.delCol, icon: 'delCol', run: op('delCol') },
        { label: str.table.delTable, icon: 'trash', run: op('delTable') },
        ...(s.cells && s.cells.n > 1 ? [{ label: str.table.merge, icon: 'merge', run: () => editor.mergeCells() } as MenuItem] : []),
        ...(s.cells?.merged && s.cells.n < 2 ? [{ label: str.table.unmerge, icon: 'merge', run: () => editor.unmergeCell() } as MenuItem] : []),
        { label: str.table.distribute, icon: 'colAdd', run: () => editor.distributeColumns() },
      )
    } else {
      items.push(
        { label: str.edit.imageMenu, icon: 'image', sep: true, run: pickImage, disabled: !s },
        { label: str.table.insert, icon: 'table', run: () => editor.insertTable(3, 3), disabled: !s },
        { label: str.docx.pageBreak, icon: 'pagebreak', hint: `${MOD}↵`, run: () => editor.insertPageBreak(), disabled: !s },
      )
    }
    // the style takes the look of the selection; every paragraph of that style follows
    const match = (label: string, value: string) => () =>
      void editor
        .updateStyle(value)
        .then((changed) => changed !== undefined && setToast(changed ? str.docx.styleUpdated(label) : str.docx.styleSame(label)))
        .catch((e) => console.error('[office-engine] style update failed', e))
    items.push(
      {
        label: str.docx.updateStyle,
        icon: 'styleMatch',
        sep: true,
        disabled: !s || !!s.img,
        items: styleChoices(meta?.styles ?? [], str).map((o) => ({ label: o.label, hint: o.value === s?.style ? str.docx.currentStyle : undefined, sep: o.sep, run: match(o.label, o.value) })),
      },
      { label: str.edit.clearFormat, icon: 'clear', hint: `${MOD}\\`, run: () => editor.clearFormat(), disabled: !s },
      { label: str.edit.selectAll, icon: 'selectAll', hint: `${MOD}A`, run: () => editor.selectAll() },
    )
    return items
  }

  // ---- comments: threads mark their text; the panel asks the editor where they are ----
  const cm = useCommentsCtl()
  // tracked changes are signed like comments: the collaboration user, else "Me"
  editor.author = () => cm?.me.name ?? 'Me'
  const cmState = useCommentsState(cm)
  const revealRef = useRef<(id: string) => void>(() => {})
  revealRef.current = (id: string) => {
    const place = editor.commentPlaces().get(id)
    const sc = scrollRef.current
    if (!place || !sc) return
    const pi = pageOfPos(place.pa, place.o)
    if (pi < 0) return
    const show = () => {
      const el = pagesRef.current?.querySelector(`[data-cm~="${CSS.escape(id)}"]`)
      if (!el) return false
      const r = el.getBoundingClientRect()
      const box = sc.getBoundingClientRect()
      if (r.top < box.top + 24 || r.bottom > box.bottom - 24) scrollByY(r.top - box.top - box.height / 3)
      return true
    }
    if (show()) return
    scrollToY(tops[pi] - 16)
    afterScroll(() => void show())
  }
  useEffect(() => {
    if (!cm) return
    const host: CommentHost = {
      kind: 'docx',
      hint: () => cm.t.docx.commentHint,
      pending: () => {
        const r = editor.commentRange()
        return r ? { quote: r.quote, range: r } : null
      },
      attach: (id, p) => editor.markComment(id, p.range as CommentRange),
      placed: (threads) => {
        const places = editor.commentPlaces()
        const order = new Map([...places.keys()].map((id, i) => [id, i]))
        return threads
          .filter((t) => places.has(t.root.id))
          .map((t): PlacedThread => ({ ...t, quote: places.get(t.root.id)!.quote }))
          .sort((a, b) => order.get(a.root.id)! - order.get(b.root.id)!)
      },
      reveal: (id) => revealRef.current(id),
    }
    cm.setHost(host)
    return () => cm.setHost(null)
  }, [cm, editor])
  useEffect(() => cm?.touch(), [cm, version])
  const cmMarks = cm ? cm.threads().filter((t) => !t.root.done || t.root.id === cmState.active).map((t) => t.root.id) : []
  // the text a thread being written covers stays highlighted while the box has the focus
  useIsoLayoutEffect(() => {
    const H = (globalThis as unknown as { Highlight?: new (...r: Range[]) => unknown }).Highlight
    const reg = (CSS as unknown as { highlights?: Map<string, unknown> }).highlights
    if (!H || !reg) return
    const range = cmState.draft?.range as CommentRange | undefined
    const root = pagesRef.current
    const a = range && editor.posOf(range.a)
    const f = range && editor.posOf(range.f)
    const da = root && a && domFromPos(root, a)
    const df = root && f && domFromPos(root, f)
    if (!da || !df) return void reg.delete('de-cm-draft')
    // within one paragraph, a range on each page it is on
    if (a!.pa === f!.pa) return void reg.set('de-cm-draft', new H(...paraRanges(root!, a!.pa, Math.min(a!.o, f!.o), Math.max(a!.o, f!.o))))
    const r = document.createRange()
    r.setStart(da[0], da[1])
    r.setEnd(df[0], df[1])
    reg.set('de-cm-draft', new H(r))
  })
  useEffect(() => () => void (CSS as unknown as { highlights?: Map<string, unknown> }).highlights?.delete('de-cm-draft'), [])

  // ---- outline ----
  const titleId = meta?.styles?.find((x) => x.canon === 'title')?.id
  const headings = useMemo(
    () => (outline ? outlineOf(st.blocks, titleId) : []),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [outline, version, titleId],
  )
  // the heading jumped to counts as current while its page is the one in view, though it may sit
  // below headings further up that page
  const [jumpBi, setJumpBi] = useState(-1)
  // until the caret goes somewhere else on that page: then the heading above the caret is current again
  useEffect(() => {
    if (caretBi >= 0 && caretBi !== jumpBi) setJumpBi(-1)
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [caretBi])
  const curPage = pages[Math.max(0, cur - 1)]
  const topBlock = curPage?.start ?? 0
  const jumped = curPage && jumpBi >= curPage.start && jumpBi < curPage.end ? jumpBi : -1
  const atBlock = Math.max(topBlock, caretBi, jumped)
  const activeHeading = headings.reduce((acc, h, i) => (h.bi <= atBlock ? i : acc), -1)
  const gotoBlock = (bi: number) => {
    const sc = scrollRef.current
    const pi = pagesOf(pages, bi)[0]
    if (!sc || pi < 0) return
    setJumpBi(bi)
    scrollToY(tops[pi] - 16)
    afterScroll(() => {
      const el = pagesRef.current?.querySelector(`[data-pa="${bi}"]`)
      if (el) scrollByY(el.getBoundingClientRect().top - sc.getBoundingClientRect().top - 24)
      if (isEditing) editor.input.restore({ a: { pa: String(bi), o: 0 }, f: { pa: String(bi), o: 0 } })
    })
  }

  // ---- controller: zoom, page, headings, undo ----
  const { ctl, headless } = viewer
  // the headings for the controller: read again once the document stops changing, not on every keystroke
  const observed = useControllerObserved(ctl)
  const [ctlHeadings, setCtlHeadings] = useState<{ bi: number; lvl: number; text: string }[]>([])
  useEffect(() => {
    if (!observed) return
    const id = setTimeout(() => setCtlHeadings(outlineOf(st.blocks, titleId)), 150)
    return () => clearTimeout(id)
  }, [observed, version, titleId, st])
  const ctlActive = ctlHeadings.reduce((acc, h, i) => (h.bi <= atBlock ? i : acc), -1)
  useEffect(() => {
    ctl?.update({ zoom: { value: zoom, effective: z } })
  }, [ctl, zoom, z])
  useEffect(() => {
    ctl?.update({ page: { current: cur, count: pageCount } })
  }, [ctl, cur, pageCount])
  useEffect(() => {
    ctl?.update({ outline: { kind: 'headings', items: ctlHeadings.map((h, i) => ({ id: String(h.bi), title: h.text, level: h.lvl, current: i === ctlActive })) } })
  }, [ctl, ctlHeadings, ctlActive])
  // undo and redo run only while editing (below), so they are offered only then
  useEffect(() => {
    if (!ctl) return
    const publish = (canUndo: boolean, canRedo: boolean) => ctl.update({ canUndo: isEditing && canUndo, canRedo: isEditing && canRedo })
    publish(editor.canUndo, editor.canRedo)
    editor.onHistory = publish
    return () => {
      editor.onHistory = undefined
    }
  }, [ctl, editor, isEditing])
  const publishedSelection = useRef<Sel | null>(null)
  const publishFormat = useViewCommands(ctl, {
    setZoom,
    goto: (id) => gotoBlock(Number(id)),
    gotoPage: (n) => {
      const top = tops[Math.min(Math.max(1, n), tops.length) - 1]
      if (top !== undefined) scrollToY(top - 16)
    },
    undo: () => {
      if (isEditing) editor.undo()
    },
    redo: () => {
      if (isEditing) editor.redo()
    },
    format: (c) => {
      const sel = publishedSelection.current
      if (isEditing && sel) editor.input.withFormatSelection(sel, () => applyDocx(editor, c))
    },
  })
  useEffect(() => {
    if (!ctl) return
    if (!isEditing) return publishFormat(null)
    const off = editor.subscribeState((s) => {
      const sel = s && editor.input.formatSelection()
      publishedSelection.current = sel
      const current = sel && editor.state(sel, { pendingList: true })
      publishFormat(current ? docxFormat(current) : null)
    })
    return () => {
      off()
      publishedSelection.current = null
      publishFormat(null)
    }
  }, [ctl, editor, isEditing, publishFormat])

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
        {meta && queue.length > 0 && fontsReady && (
          <div ref={measureRef} className="de-measure de-doc" style={{ ...docStyle, width: `${contentW}pt` }}>
            {queue.map(({ b, i }) => (
              <BlockView key={i} block={b} ctx={measureCtx} sp={contextual(b, st.blocks[i - 1], st.blocks[i + 1])} />
            ))}
          </div>
        )}
        <div
          ref={pagesRef}
          className={isEditing ? 'de-pages de-editing' : 'de-pages'}
          style={{ height: spacerH, width: totalW, overflow: scrollScale > 1 ? 'clip' : undefined }}
          contentEditable={isEditing || undefined}
          suppressContentEditableWarning
          spellCheck={false}
          onMouseDown={(e) => {
            if (!isEditing) return
            const t = e.target as HTMLElement
            // a copy of a table's header rows on a following page shows them, the rows themselves are edited on the first
            if (t.closest('.de-head-copy')) return void e.preventDefault()
            // clicking a picture selects it (resize handles, image tools)
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
          onClick={(e) => {
            if (!cm) return
            // a click on commented text shows its thread
            const ids = (e.target as HTMLElement).closest?.<HTMLElement>('[data-cm]')?.dataset.cm?.split(' ') ?? []
            const id = ids.find((x) => cmMarks.includes(x))
            if (id) cm.select(id, false)
            else if (cmState.active) cm.select(null)
          }}
          onContextMenu={(e) => {
            if (!isEditing) return
            e.preventDefault()
            setCtxMenu({ x: e.clientX, y: e.clientY })
          }}
        >
          {pages.map((p, i) =>
            (i < first || i > last) && !pinned.has(i) ? null : (
              <PageView
                key={i}
                index={i}
                count={pageCount}
                slice={p}
                blocks={st.blocks}
                meta={meta}
                page={page}
                top={tops[i] - scrollOffset}
                left={left}
                height={pageHeights[i]}
                zoom={z}
                client={client}
                docStyle={docStyle}
                editing={isEditing}
                colResize={colResize}
                onHF={isEditing && done ? openHF : undefined}
                lines={st.lines}
                avail={availPx}
              />
            ),
          )}
          {skeleton > 0 && <div className="de-page-skeleton" contentEditable={false} style={{ top: acc - scrollOffset, left, width: pageW * z, height: pageH * z }} />}
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
      {outline && !headless && (
        <nav className="de-outline" aria-label={str.docx.outline}>
          <div className="de-outline-head">
            <span>{str.docx.outlineHead}</span>
            <button type="button" className="de-tb" title={str.docx.closeOutline} aria-label={str.docx.closeOutline} onClick={() => setOutline(false)}>
              <Icon name="close" size={16} />
            </button>
          </div>
          {headings.length ? (
            headings.map((h, i) => (
              <button key={h.bi} type="button" className={i === activeHeading ? 'de-outline-item on' : 'de-outline-item'} style={{ paddingLeft: 10 + h.lvl * 14 }} onClick={() => gotoBlock(h.bi)}>
                {h.text}
              </button>
            ))
          ) : (
            <div className="de-outline-empty">{str.docx.outlineEmpty}</div>
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
          onOutline={headless ? undefined : () => setOutline((o) => !o)}
          painter={painter}
          onLink={askLink}
          image={imageActions}
          onPageSetup={done ? () => setPageDlg(true) : undefined}
          onHeaderFooter={done ? () => setHfDlg('header') : undefined}
          onNotice={setToast}
        />
      )}
      {!isEditing && done && !headless && (
        <button type="button" className={outline ? 'de-outline-toggle on' : 'de-outline-toggle'} title={str.docx.outline} onClick={() => setOutline((o) => !o)}>
          <Icon name="outline" />
        </button>
      )}
      {pageDlg && <PageSetupDialog page={page} onClose={() => setPageDlg(false)} onOk={(set) => void editor.setPage(set).catch((e) => console.error('[office-engine] page setup failed', e))} />}
      {hfDlg && meta && (
        <HeaderFooterDialog meta={meta} focus={hfDlg} onClose={() => setHfDlg(null)} onOk={(set) => void editor.setHeaderFooter(set).catch((e) => console.error('[office-engine] header / footer change failed', e))} />
      )}
      {ctxMenu && <ContextMenu x={ctxMenu.x} y={ctxMenu.y} items={ctxItems()} onClose={() => setCtxMenu(null)} />}
      {prompt && <PromptDialog spec={prompt} onClose={() => setPrompt(null)} />}
      {toast && <Toast text={toast} onDone={() => setToast(null)} />}
      {cmMarks.length > 0 && <style>{markCss('.de-pages', cmMarks, cmState.active)}</style>}
      {!headless && <ZoomControl zoom={zoom} effective={z} onChange={setZoom} />}
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
  colResize?: RenderCtx['colResize']
  /** editing: a double click on the header / footer area edits it */
  onHF?: (k: HFKind) => void
  /** where the blocks' lines / rows are, and the height of a page's text (px) */
  lines: WeakMap<Block, Lines>
  avail: number
}

function PageView({ index, count, slice, blocks, meta, page, top, left, height, zoom, client, docStyle, editing, colResize, onHF, lines, avail }: PageViewProps) {
  const str = useStrings()
  const ctx = useMemo<RenderCtx>(
    () => ({ client, page: index, pages: count, measuring: false, editable: editing, colResize: editing ? colResize : undefined, t: str }),
    [client, index, count, editing, colResize, str],
  )
  const hfCtx = useMemo<RenderCtx>(() => ({ ...ctx, editable: false }), [ctx])
  // the first page's, the even pages' (when the file has different odd and even ones) or the default
  const { header, footer, hk, fk } = pageHF(meta ?? undefined, index)
  // each column's blocks; a block split across pages (or columns) shows the lines / rows it has here
  const columns = columnsOf(slice).map((c) => {
    const items = []
    for (let i = c.start; i < c.end; i++) {
      const b = blocks[i]
      let { from, to }: { from?: number; to?: number } = pieceOf(c, i)
      let head: number | undefined
      let height: number | undefined
      if (b.t === 'tbl' && (from || to !== undefined)) head = from ? headOf(lines.get(b), avail) : undefined
      else if (b.t === 'p' && (from || to !== undefined)) {
        // lines [from, to): the characters they hold, the height they take
        const l = lines.get(b)
        const a = from ? l?.at?.[from] : 0
        const z = to !== undefined ? l?.at?.[to] : undefined
        if (!l || a === undefined || (to !== undefined && z === undefined)) {
          // where its lines start is being read (before this is painted): the place is kept
          items.push(<div key={`${i}:wait:${from}:${to}`} style={{ height: l ? l.ys[to ?? l.ys.length - 1] - (from ? l.ys[from] : 0) : 0 }} />)
          continue
        }
        if (to !== undefined) height = l.ys[to] - (from ? l.ys[from] : 0)
        from = a
        to = z
      } else from = to = undefined
      items.push(
        <BlockView
          // a part that starts or ends elsewhere is a new element (the caret is put back in it)
          key={`${i}:${b.t === 'p' ? (b.ver ?? 0) : 0}:${from ?? ''}:${to ?? ''}`}
          block={b}
          ctx={ctx}
          pa={editable(b) || b.t === 'tbl' ? String(i) : undefined}
          sp={contextual(b, blocks[i - 1], blocks[i + 1])}
          from={from || undefined}
          to={to}
          head={head}
          height={height}
        />,
      )
    }
    return items
  })
  const nCols = page.cols && page.cols > 1 ? page.cols : 1
  const gap = nCols > 1 ? (page.colGap ?? 36) : 0
  const colW = (page.w - page.ml - page.mr - gap * (nCols - 1)) / nCols
  return (
    <div className="de-page" style={{ top, left, width: page.w * PX * zoom, height }} data-page={index + 1}>
      <div className="de-page-inner de-doc" style={{ ...docStyle, width: `${page.w}pt`, height: height / zoom, zoom }}>
        {onHF && (
          <>
            <div className="de-hf-zone" contentEditable={false} title={str.docx.editHeader} style={{ top: 0, height: `${page.mt}pt` }} onDoubleClick={() => onHF(hk)} />
            <div className="de-hf-zone" contentEditable={false} title={str.docx.editFooter} style={{ bottom: 0, height: `${page.mb}pt` }} onDoubleClick={() => onHF(fk)} />
          </>
        )}
        {header && (
          <div className="de-hf" contentEditable={editing ? false : undefined} style={{ top: `${page.header}pt`, padding: `0 ${page.mr}pt 0 ${page.ml}pt` }} onDoubleClick={onHF && (() => onHF(hk))}>
            {header.map((b, i) => (
              <BlockView key={i} block={b} ctx={hfCtx} />
            ))}
          </div>
        )}
        <div
          className={nCols > 1 ? 'de-page-cols' : undefined}
          style={{ position: 'absolute', top: `${page.mt}pt`, left: `${page.ml}pt`, width: `${page.w - page.ml - page.mr}pt`, ...(nCols > 1 ? { display: 'flex', columnGap: `${gap}pt` } : {}) }}
        >
          {nCols > 1
            ? columns.map((items, c) => (
                <div key={c} className="de-page-col" style={{ width: `${colW}pt`, flex: 'none' }}>
                  {items}
                </div>
              ))
            : columns[0]}
        </div>
        {footer && (
          <div className="de-hf" contentEditable={editing ? false : undefined} style={{ bottom: `${page.footer}pt`, padding: `0 ${page.mr}pt 0 ${page.ml}pt` }} onDoubleClick={onHF && (() => onHF(fk))}>
            {footer.map((b, i) => (
              <BlockView key={i} block={b} ctx={hfCtx} />
            ))}
          </div>
        )}
      </div>
    </div>
  )
}
