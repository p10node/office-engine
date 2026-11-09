import { useCallback, useEffect, useInsertionEffect, useMemo, useRef, useState, type CSSProperties, type ReactNode } from 'react'
import type { DocSource } from '../core/source'
import type { DocumentClient, OpenOptions } from '../core/worker/client'
import type { Zoom } from './common/ZoomControl'
import { DocxView } from './docx/DocxView'
import { ViewerContext, type FileOps, type ViewerApi } from './common/Slots'
import { StatusContext } from './common/Status'
import { ConfirmDialog, PromptDialog, Toast } from './common/Dialog'
import { download, Icon } from './edit/Toolbar'
import { DetailsDialog } from './file/DetailsDialog'
import { FileMenu, KIND_LABEL, useFileActions, type FileMenuItem } from './file/FileMenu'
import { importFile, type Imported } from './file/importFile'
import { blankPptx, parasToDocx, rowsToXlsx } from '../core/convert'
import type { DocKind } from '../core/opc'
import { extOf, HAS_SOURCE, PLAIN_EXT, PLAIN_FILE, type PlainFormat } from '../core/plain'
import { injectStyles } from './injectStyles'
import { isTextEntry, useDarkTheme, useSize, type Theme } from './common/hooks'
import { SourceEditor } from './source/SourceEditor'
import { PptxView } from './pptx/PptxView'
import { useDocument } from './useDocument'
import { XlsxView } from './xlsx/XlsxView'
import type { CollabOptions } from '../collab'
import { CollabBar, PeerOverlay, useCollab } from '../collab/react'
import { CommentsContext, CommentsCtl, useCommentsState, type CommentOptions } from './comments/Comments'
import { COMMENT_KEYS, CommentsPanel } from './comments/CommentsPanel'

export interface DocViewerProps {
  /**
   * File, Blob, ArrayBuffer, Uint8Array or URL (HTTP Range requests are used when the server supports them): an Office file,
   * an Apple Pages / Keynote / Numbers file (view only), or CSV / TSV / text / Markdown / HTML (opened as a workbook / document).
   */
  src: DocSource | null | undefined
  /** Display name (defaults to File.name or the URL's last segment); its extension also tells the format of a source without one. */
  name?: string
  className?: string
  style?: CSSProperties
  /** Show the top bar (name, status, progress). Default true. */
  toolbar?: boolean
  /** Extra toolbar content (rendered at the right). */
  toolbarExtra?: ReactNode
  /** Initial zoom for documents and presentations. Default 'fit'. */
  zoom?: Zoom
  /** Show slide thumbnails for presentations. Default true. */
  thumbnails?: boolean
  headers?: Record<string, string>
  credentials?: RequestCredentials
  /** Custom worker factory, or false to parse on the main thread. */
  worker?: OpenOptions['worker']
  /** Where to inject the stylesheet (document or a shadow root). */
  styleTarget?: Document | ShadowRoot
  /**
   * Color scheme of the viewer chrome (bars, menus, dialogs): 'light', 'dark' or 'auto' (follows the OS / browser setting).
   * The document itself keeps its own colors. Default 'light'.
   */
  theme?: Theme
  onLoad?: (client: DocumentClient) => void
  onError?: (error: Error) => void
  /** Start in edit mode. Default false (view). */
  editable?: boolean
  /** Show the Edit/View toggle. Default true. */
  allowEdit?: boolean
  /** Show the Save/Download button and enable Ctrl+S. Default true. */
  allowSave?: boolean
  /**
   * Allow downloading files to the user's device (the save fallback without `onSave`, "Download image").
   * With `onSave`, saving still works when this is false. Default true.
   */
  allowDownload?: boolean
  /** Show the File menu (new, open, save, rename, download as PDF / HTML / text…, details, print) and its shortcuts. Default true. */
  fileMenu?: boolean
  /** File → New / Open: replace the document with a blank one or a local file (Office, Pages, Keynote, Numbers, CSV, text, Markdown, HTML). Default true. */
  allowOpen?: boolean
  /** Extra File menu entries (Share, Version history…), shown before Close. */
  fileMenuItems?: FileMenuItem[]
  /** File → Rename: called with the new name (extension included). Rename is shown when `fileMenu` is on. */
  onRename?: (name: string) => void
  /** Adds File → Close. */
  onClose?: () => void
  /** Receives the saved file instead of downloading it. */
  onSave?: (file: Blob, name: string) => void | Promise<void>
  /** Called when the document gets / loses unsaved changes. */
  onDirtyChange?: (dirty: boolean) => void
  /**
   * Real-time collaboration: everyone opening the same base document (`src`) with the same `room`
   * edits it together through the hub at `transport` (see src/collab).
   */
  collab?: CollabOptions
  /**
   * Ctrl/⌘+A anywhere on the page selects the document (text fields outside the viewer keep their own).
   * Default true. With false the shortcut works from the viewer only (its content or toolbar).
   */
  captureSelectAll?: boolean
  /**
   * Comments (stored in the file, shared with collaborators): who writes them, who can be @mentioned, what the
   * host hears about. Shown by default, written while editing; `false` hides them.
   */
  comments?: CommentOptions | false
}

const EXT: Record<string, string> = { docx: '.docx', xlsx: '.xlsx', pptx: '.pptx' }
const FROM_LABEL: Record<'pages' | 'keynote' | 'numbers', string> = { pages: 'Apple Pages document', keynote: 'Apple Keynote presentation', numbers: 'Apple Numbers spreadsheet' }
const IMPORTED_BADGE: Record<PlainFormat, string> = { csv: 'csv', tsv: 'tsv', text: 'txt', markdown: 'md', html: 'html', mermaid: 'mmd', plantuml: 'puml', graphviz: 'dot' }
const IMPORTED_LABEL: Record<PlainFormat, string> = {
  csv: 'CSV file',
  tsv: 'Tab-separated values',
  text: 'Plain text',
  markdown: 'Markdown',
  html: 'HTML page',
  mermaid: 'Mermaid diagram',
  plantuml: 'PlantUML diagram',
  graphviz: 'Graphviz diagram',
}
const UNTITLED: Record<DocKind, string> = { docx: 'Untitled document', xlsx: 'Untitled spreadsheet', pptx: 'Untitled presentation' }

/** a key per opened client: a view is built again for each (the edited source's previews have one name) */
const clientIds = new WeakMap<object, number>()
let nextClientId = 0
const clientKey = (c: object) => {
  let id = clientIds.get(c)
  if (id === undefined) clientIds.set(c, (id = ++nextClientId))
  return id
}

function fmtBytes(n: number) {
  if (n < 1024) return `${n} B`
  if (n < 1024 * 1024) return `${(n / 1024).toFixed(0)} KB`
  return `${(n / 1024 / 1024).toFixed(1)} MB`
}

export function DocViewer(props: DocViewerProps) {
  const { src, name, className, style, toolbar = true, toolbarExtra, zoom, thumbnails, headers, credentials, worker, styleTarget, theme, onLoad, onError } = props
  const dark = useDarkTheme(theme)
  const { allowEdit = true, allowSave = true, allowDownload = true, fileMenu = true, allowOpen = true, captureSelectAll = true, onSave, onDirtyChange } = props
  const [editing, setEditing] = useState(!!props.editable)
  useEffect(() => {
    if (props.editable !== undefined) setEditing(!!props.editable)
  }, [props.editable])
  const [dirty, setDirtyState] = useState(false)
  const [saving, setSaving] = useState(false)
  const [barEl, setBarEl] = useState<HTMLDivElement | null>(null)
  useInsertionEffect(() => injectStyles(styleTarget), [styleTarget])
  // File → New / Open: a local file shown instead of `src` until the host passes another `src`
  const [opened, setOpened] = useState<Imported | null>(null)
  const [renamed, setRenamed] = useState<string | null>(null)
  useEffect(() => {
    setOpened(null)
    setRenamed(null)
  }, [src])
  const loadName = opened ? opened.name : name
  // the name is part of what is loaded: Rename only changes what is shown and saved
  // a preview of the edited source keeps the one shown until it is ready
  const { client, error, loading, progress } = useDocument(opened ? opened.src : src, { name: loadName, headers, credentials, worker }, !!opened?.live)
  const docName = renamed ?? loadName
  const collab = useCollab(client, props.collab)
  // converted from another format (Apple Pages, Keynote, Numbers): view only
  const readOnly = !!client?.info?.from
  // with collaboration, editing waits for the shared history (offline edits sync on reconnect)
  const canEdit = !readOnly && (!props.collab || collab.status === 'live' || collab.status === 'offline')
  const [statusEl, setStatusEl] = useState<HTMLSpanElement | null>(null)

  // ---- source (Markdown, HTML, diagrams): edited side by side with the document built from it ----
  const srcFormat = client?.info?.imported
  const hasSource = !!srcFormat && HAS_SOURCE.has(srcFormat)
  const [codeOpen, setCodeOpen] = useState(false)
  const codeShown = codeOpen && hasSource
  const [code, setCode] = useState<string | null>(null)
  const sourceEdited = useRef(false)
  const reloadTimer = useRef<ReturnType<typeof setTimeout> | undefined>(undefined)
  useEffect(() => () => clearTimeout(reloadTimer.current), [])
  // another document (not a preview of the edited source): its own source
  useEffect(() => {
    if (opened?.live) return
    clearTimeout(reloadTimer.current)
    setCode(null)
    sourceEdited.current = false
  }, [src, opened])
  useEffect(() => {
    if (!client || !codeShown || code !== null) return
    let live = true
    client.call('plain.source').then(
      (t) => live && setCode(t ?? ''),
      () => {},
    )
    return () => {
      live = false
    }
  }, [client, codeShown, code])
  const [bodyRef, bodySize] = useSize<HTMLDivElement>()
  const mainRef = useRef<HTMLDivElement>(null)
  const stacked = bodySize.w > 0 && bodySize.w < 720
  const [split, setSplit] = useState(0.45)
  // the preview's scroll position, kept across rebuilds
  const restoreScroll = useRef<number | null>(null)
  useEffect(() => {
    const ratio = restoreScroll.current
    if (ratio === null || !client) return
    restoreScroll.current = null
    let raf = 0
    let frames = 0
    let lastH = -1
    const tick = () => {
      const sc = mainRef.current?.querySelector<HTMLElement>('.de-scroll')
      if (sc && sc.scrollHeight > sc.clientHeight + 1) {
        sc.scrollTop = ratio * (sc.scrollHeight - sc.clientHeight)
        if (sc.scrollHeight === lastH) return
        lastH = sc.scrollHeight
      }
      if (++frames < 40) raf = requestAnimationFrame(tick)
    }
    raf = requestAnimationFrame(tick)
    return () => cancelAnimationFrame(raf)
  }, [client])

  // ---- comments ----
  const cmOn = props.comments !== false
  const cmCtl = useMemo(() => (client && cmOn ? new CommentsCtl(client) : null), [client, cmOn])
  useEffect(() => () => cmCtl?.dispose(), [cmCtl])
  const cmOpts = props.comments || {}
  if (cmCtl) {
    cmCtl.opts = cmOpts
    cmCtl.me = cmOpts.user ?? props.collab?.user ?? { name: 'Me' }
    cmCtl.canComment = !readOnly && (cmOpts.canComment ?? (editing && canEdit))
  }
  const cmState = useCommentsState(cmCtl)
  const cmOpenAtLoad = !!cmOpts.open
  useEffect(() => {
    if (cmCtl && cmOpenAtLoad) cmCtl.toggle(true)
  }, [cmCtl]) // eslint-disable-line react-hooks/exhaustive-deps
  const cmCount = cmCtl ? cmCtl.visible().length : 0
  // comments.focus: the thread shown once it is in the document (DOCX threads appear as the text loads)
  const cmFocus = cmOpts.focus
  const cmFocused = useRef<string | null>(null)
  useEffect(() => {
    if (!cmCtl || !cmFocus || cmFocused.current === cmFocus || !cmCtl.loaded) return
    const c = cmCtl.list.get(cmFocus)
    const root = c?.parent ?? cmFocus
    if (!cmCtl.threads().some((t) => t.root.id === root)) return
    cmFocused.current = cmFocus
    cmCtl.toggle(true)
    if (cmCtl.list.get(root)?.done) cmCtl.setShowResolved(true)
    cmCtl.select(root)
  }, [cmCtl, cmFocus, cmState.rev])

  useEffect(() => {
    if (client) onLoad?.(client)
  }, [client, onLoad])
  useEffect(() => {
    if (error) onError?.(error)
  }, [error, onError])

  // a converted file (CSV → xlsx…) starts unsaved
  const openedDirty = useRef(false)
  useEffect(() => {
    setDirtyState(!!client && openedDirty.current)
    if (client) openedDirty.current = false
  }, [client])
  const setDirty = useCallback(
    (d: boolean) =>
      setDirtyState((prev) => {
        if (prev !== d) onDirtyChange?.(d)
        return d
      }),
    [onDirtyChange],
  )
  const baseName = (renamed ?? (client?.name || docName || 'document')).replace(/\.[^.]+$/, '')
  const fileName = baseName + (client ? EXT[client.kind] : '')
  // saving without onSave means downloading
  const canSave = allowSave && (!!onSave || allowDownload)
  const save = useCallback(async () => {
    if (!client || saving || !canSave) return
    setSaving(true)
    try {
      const blob = await client.call('save')
      if (onSave) await onSave(blob, fileName)
      else download(blob, fileName)
      setDirty(false)
    } catch (e) {
      onError?.(e as Error)
      console.error('[office-engine] save failed', e)
    } finally {
      setSaving(false)
    }
  }, [client, saving, canSave, onSave, fileName, setDirty, onError])
  const fileOps = useRef<FileOps | null>(null)
  const selectAll = useRef<(() => void) | null>(null)
  // next to its source, the document is a preview: it is edited through the source
  const docEditable = editing && canEdit && !codeShown
  const api = useMemo<ViewerApi>(
    () => ({ bar: barEl, editable: docEditable, setDirty, save, canDownload: allowDownload, client, fileOps, selectAll }),
    [barEl, docEditable, setDirty, save, allowDownload, client],
  )

  const editSource = (text: string) => {
    setCode(text)
    if (!sourceEdited.current) setDirty(true)
    sourceEdited.current = true
    clearTimeout(reloadTimer.current)
    const fmt = srcFormat
    if (!fmt) return
    reloadTimer.current = setTimeout(() => {
      const sc = mainRef.current?.querySelector<HTMLElement>('.de-scroll')
      restoreScroll.current = sc && sc.scrollHeight > sc.clientHeight ? sc.scrollTop / (sc.scrollHeight - sc.clientHeight) : null
      // the same name: the same format, title and saved name
      let n = client?.name || loadName || `document.${PLAIN_FILE[fmt].ext}`
      if (PLAIN_EXT[extOf(n)] !== fmt) n += `.${PLAIN_FILE[fmt].ext}`
      openedDirty.current = true
      setOpened({ src: new Blob([text], { type: PLAIN_FILE[fmt].mime }), name: n, converted: true, live: true })
    }, 350)
  }
  const sourceName = () => `${(renamed ?? client?.name ?? docName ?? 'document').replace(/\.[^.]+$/, '')}.${PLAIN_FILE[srcFormat ?? 'text'].ext}`
  const dragSplit = (e: React.PointerEvent<HTMLDivElement>) => {
    const body = bodyRef.current
    if (!body) return
    e.preventDefault()
    const handle = e.currentTarget
    handle.setPointerCapture(e.pointerId)
    const r = body.getBoundingClientRect()
    const move = (ev: PointerEvent) => setSplit(Math.min(0.85, Math.max(0.15, stacked ? (ev.clientY - r.top) / r.height : (ev.clientX - r.left) / r.width)))
    const up = () => {
      handle.removeEventListener('pointermove', move)
      handle.removeEventListener('pointerup', up)
      handle.removeEventListener('pointercancel', up)
    }
    handle.addEventListener('pointermove', move)
    handle.addEventListener('pointerup', up)
    handle.addEventListener('pointercancel', up)
  }
  const rootRef = useRef<HTMLDivElement>(null)
  // captureSelectAll: Ctrl/⌘+A from the rest of the page too (the viewer's own keys are handled in onKeyDown below)
  useEffect(() => {
    const root = rootRef.current
    if (!captureSelectAll || !root) return
    const on = (e: KeyboardEvent) => {
      if (e.defaultPrevented || !(e.metaKey || e.ctrlKey) || e.altKey || e.shiftKey || e.key.toLowerCase() !== 'a') return
      const target = (e.composedPath()[0] ?? e.target) as Node | null
      if ((target && root.contains(target)) || isTextEntry(target) || !selectAll.current) return
      e.preventDefault()
      selectAll.current()
    }
    const doc = root.ownerDocument
    doc.addEventListener('keydown', on)
    return () => doc.removeEventListener('keydown', on)
  }, [captureSelectAll])

  // Ctrl/⌘+Alt+M: comment on the selection — from the viewer, or the page around it when nothing else has the focus
  // (a click on a slide shape leaves it on the page)
  const cmRef = useRef(cmCtl)
  cmRef.current = cmCtl
  useEffect(() => {
    const root = rootRef.current
    if (!root) return
    const on = (e: KeyboardEvent) => {
      const ctl = cmRef.current
      if (e.defaultPrevented || !(e.metaKey || e.ctrlKey) || !e.altKey || e.shiftKey || e.code !== 'KeyM' || !ctl?.canComment) return
      const target = (e.composedPath()[0] ?? e.target) as Node | null
      const doc = root.ownerDocument
      if (!(target && root.contains(target)) && target !== doc.body && target !== doc.documentElement) return
      e.preventDefault()
      ctl.start()
    }
    const doc = root.ownerDocument
    doc.addEventListener('keydown', on)
    return () => doc.removeEventListener('keydown', on)
  }, [])

  // ---- File menu ----
  const [toast, setToast] = useState<string | null>(null)
  const clearToast = useCallback(() => setToast(null), [])
  const [confirm, setConfirm] = useState<{ message: string; run: () => void } | null>(null)
  const [renaming, setRenaming] = useState(false)
  const [details, setDetails] = useState(false)
  /** replaces the document, after confirmation when there are unsaved changes */
  const replace = (message: string, run: () => void) => (dirty ? setConfirm({ message, run }) : run())
  const show = useCallback((doc: Imported, edit = false) => {
    openedDirty.current = doc.converted
    setRenamed(null)
    setOpened(doc)
    if (edit) setEditing(true)
  }, [])
  const canOpen = fileMenu && allowOpen && !props.collab
  const fileActions = useFileActions({
    kind: client?.kind,
    base: baseName,
    canOpen,
    canImport: fileMenu && allowEdit && canEdit,
    canSave: fileMenu && canSave,
    canDownload: fileMenu && allowDownload && !!client,
    ops: fileOps,
    create: (kind: DocKind) =>
      replace(`A new ${UNTITLED[kind].split(' ')[1]} replaces the current document. Your changes are not saved.`, () => {
        const blank = kind === 'docx' ? parasToDocx([]) : kind === 'xlsx' ? rowsToXlsx([], 'Sheet1') : blankPptx()
        blank.then((src) => show({ src, name: UNTITLED[kind] + EXT[kind], converted: false }, allowEdit), (e: Error) => setToast(e.message))
      }),
    open: (file) =>
      replace(`Opening “${file.name}” replaces the current document. Your changes are not saved.`, () => {
        importFile(file).then(show, (e: Error) => setToast(e.message))
      }),
    save,
    downloadOriginal: async () => {
      if (!client) return
      download(await client.call('save'), fileName)
      // without onSave, saving is downloading
      if (!onSave) setDirty(false)
    },
    rename: () => setRenaming(true),
    details: () => setDetails(true),
    close: props.onClose,
    extra: props.fileMenuItems ?? [],
    notify: setToast,
  })
  const onKeyDown = (e: React.KeyboardEvent) => {
    if (e.defaultPrevented || !(e.metaKey || e.ctrlKey) || e.altKey || e.shiftKey) return
    const k = e.key.toLowerCase()
    // Ctrl/⌘+S in the source: the viewer's Save
    if (k === 's' && (e.target as HTMLElement).closest?.('.de-src')) {
      e.preventDefault()
      void save()
      return
    }
    if (k === 'a') {
      // from the viewer's own chrome: the document, not the whole page (text fields keep their own select-all)
      if (isTextEntry(e.target) || !selectAll.current) return
      e.preventDefault()
      selectAll.current()
    } else if (!fileMenu) return
    else if (k === 'p') {
      // with downloads disabled, the browser must not print (save) the page either
      e.preventDefault()
      if (allowDownload) fileActions.run('print')
    } else if (k === 'o' && canOpen) {
      e.preventDefault()
      fileActions.run('open')
    }
  }

  const pct = progress && progress.total > 0 ? Math.min(100, (progress.loaded / progress.total) * 100) : 0
  const busy = loading || (!!progress && pct < 100)
  const title = renamed ?? client?.name ?? docName ?? (typeof src === 'string' ? decodeURIComponent(src.split(/[?#]/)[0].split('/').pop() ?? '') : (src as File | undefined)?.name) ?? ''

  let body: ReactNode
  if (error) {
    body = (
      <div className="de-center">
        <div className="de-error" role="alert">
          <strong>Cannot open document.</strong>
          <div>{error.message}</div>
        </div>
      </div>
    )
  } else if (!client) {
    body = src || opened ? (
      <div className="de-center">
        <div className="de-spinner" />
        <div>{progress?.phase === 'download' && progress.total ? `Downloading… ${fmtBytes(progress.loaded)} / ${fmtBytes(progress.total)}` : 'Opening…'}</div>
      </div>
    ) : (
      <div className="de-center">No document</div>
    )
  } else if (client.kind === 'docx') body = <DocxView key={clientKey(client)} client={client} initialZoom={zoom} />
  else if (client.kind === 'xlsx') body = <XlsxView client={client} />
  else body = <PptxView client={client} initialZoom={zoom} thumbnails={thumbnails} />

  return (
    <div ref={rootRef} className={'de-root' + (dark ? ' de-dark' : '') + (className ? ` ${className}` : '')} style={style} onKeyDown={onKeyDown}>
      {toolbar && (
        <div className="de-bar">
          {client && <span className={`de-badge de-badge-${client.kind}`}>{client.info?.from ?? (client.info?.imported ? IMPORTED_BADGE[client.info.imported] : client.kind)}</span>}
          {fileMenu && <FileMenu actions={fileActions} />}
          <span className="de-title" title={title}>
            {title}
          </span>
          <span className="de-status" ref={setStatusEl} hidden={!client} />
          {client && client.size > 0 && <span className="de-status">{fmtBytes(client.size)}</span>}
          {client && props.collab && <CollabBar client={client} state={collab} />}
          {toolbarExtra}
          {client && hasSource && (
            <button
              type="button"
              className={'de-mode' + (codeShown ? ' de-mode-on' : '')}
              onClick={() => setCodeOpen((v) => !v)}
              aria-pressed={codeShown}
              title={codeShown ? 'Hide the source' : 'Edit the source side by side with the document'}
              data-testid="de-code-toggle"
            >
              <Icon name="code" size={16} />
              Code
            </button>
          )}
          {cmCtl && (
            <button
              type="button"
              className={'de-mode de-cm-toggle' + (cmState.open ? ' de-mode-on' : '')}
              onClick={() => cmCtl.toggle()}
              title={cmCtl.canComment ? `Comments (${COMMENT_KEYS} to comment on the selection)` : 'Comments'}
              aria-pressed={cmState.open}
              data-testid="de-comments-toggle"
            >
              <Icon name="comment" size={16} />
              Comments
              {cmCount > 0 && <span className="de-cm-n">{cmCount}</span>}
            </button>
          )}
          {client && allowEdit && !readOnly && !codeShown && (
            <button type="button" className={editing ? 'de-mode de-mode-on' : 'de-mode'} onClick={() => setEditing((v) => !v)} title={editing ? 'Switch to viewing' : 'Edit document'}>
              <Icon name={editing ? 'view' : 'edit'} size={16} />
              {editing ? 'View' : 'Edit'}
            </button>
          )}
          {client && canSave && ((editing && !readOnly) || dirty) && (
            <button type="button" className={dirty ? 'de-mode de-save-dirty' : 'de-mode'} disabled={saving} onClick={() => void save()} title="Save a copy (Ctrl+S)">
              <Icon name="save" size={16} />
              {saving ? 'Saving…' : dirty || onSave ? 'Save' : 'Download'}
            </button>
          )}
          <div className="de-progress" style={{ width: `${busy ? Math.max(pct, 3) : 100}%`, opacity: busy ? 1 : 0 }} />
        </div>
      )}
      {toolbar && client && docEditable && <div className="de-ebar-host" ref={setBarEl} />}
      <div className={'de-body' + (codeShown && stacked ? ' de-body-stacked' : '')} ref={bodyRef}>
        {codeShown && (
          <div className="de-src-pane" style={stacked ? { height: `${split * 100}%` } : { width: `${split * 100}%` }}>
            {code === null ? (
              <div className="de-center">
                <div className="de-spinner" />
              </div>
            ) : (
              <SourceEditor
                text={code}
                format={srcFormat!}
                onChange={editSource}
                readOnly={!allowEdit || !!props.collab}
                notice={
                  dirty && !sourceEdited.current
                    ? 'Changes made in the document are not in this source: editing it builds the document again from the source.'
                    : undefined
                }
                onCopy={() =>
                  navigator.clipboard?.writeText(code).then(
                    () => setToast('Source copied'),
                    () => setToast('Cannot copy here'),
                  )
                }
                onDownload={allowDownload ? () => download(new Blob([code], { type: PLAIN_FILE[srcFormat!].mime }), sourceName()) : undefined}
              />
            )}
          </div>
        )}
        {codeShown && (
          <div
            className="de-split-handle"
            role="separator"
            aria-orientation={stacked ? 'horizontal' : 'vertical'}
            aria-label="Resize the source"
            aria-valuenow={Math.round(split * 100)}
            aria-valuemin={15}
            aria-valuemax={85}
            tabIndex={0}
            onPointerDown={dragSplit}
            onKeyDown={(e) => {
              const d = e.key === (stacked ? 'ArrowUp' : 'ArrowLeft') ? -0.05 : e.key === (stacked ? 'ArrowDown' : 'ArrowRight') ? 0.05 : 0
              if (!d) return
              e.preventDefault()
              setSplit((v) => Math.min(0.85, Math.max(0.15, v + d)))
            }}
          />
        )}
        <div className="de-main" ref={mainRef}>
          <ViewerContext.Provider value={api}>
            <CommentsContext.Provider value={cmCtl}>
              <StatusContext.Provider value={statusEl}>{body}</StatusContext.Provider>
            </CommentsContext.Provider>
          </ViewerContext.Provider>
        </div>
        {client && collab.collab && <PeerOverlay client={client} collab={collab.collab} />}
        {cmCtl && cmState.open && <CommentsPanel ctl={cmCtl} />}
      </div>
      {toast && <Toast text={toast} onDone={clearToast} />}
      {confirm && <ConfirmDialog title="Discard unsaved changes?" message={confirm.message} ok="Continue" onOk={confirm.run} onClose={() => setConfirm(null)} />}
      {renaming && client && (
        <PromptDialog
          spec={{
            title: 'Rename',
            label: 'File name',
            initial: baseName,
            ok: 'Rename',
            onOk: (v) => {
              const next = v.trim().replace(/[\\/:*?"<>|]+/g, ' ').replace(new RegExp(`\\${EXT[client.kind]}$`, 'i'), '')
              if (!next) return
              setRenamed(next + EXT[client.kind])
              props.onRename?.(next + EXT[client.kind])
            },
          }}
          onClose={() => setRenaming(false)}
        />
      )}
      {details && client && (
        <DetailsDialog
          client={client}
          file={[
            ['Name', fileName],
            [
              'Type',
              client.info?.from
                ? `${FROM_LABEL[client.info.from]} (view only)`
                : client.info?.imported
                  ? `${IMPORTED_LABEL[client.info.imported]}, opened as ${KIND_LABEL[client.kind]}`
                  : KIND_LABEL[client.kind],
            ],
            ...(client.size > 0 ? [['File size', fmtBytes(client.size)] as [string, string]] : []),
            ...(dirty ? [['Status', 'Unsaved changes'] as [string, string]] : []),
          ]}
          stats={fileOps.current?.stats() ?? []}
          onClose={() => setDetails(false)}
        />
      )}
    </div>
  )
}
