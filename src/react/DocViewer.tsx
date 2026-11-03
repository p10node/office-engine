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
import { injectStyles } from './injectStyles'
import { isTextEntry, useDarkTheme, type Theme } from './common/hooks'
import { PptxView } from './pptx/PptxView'
import { useDocument } from './useDocument'
import { XlsxView } from './xlsx/XlsxView'
import type { CollabOptions } from '../collab'
import { CollabBar, PeerOverlay, useCollab } from '../collab/react'

export interface DocViewerProps {
  /** File, Blob, ArrayBuffer, Uint8Array or URL (HTTP Range requests are used when the server supports them). */
  src: DocSource | null | undefined
  /** Display name (defaults to File.name or the URL's last segment). */
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
  /** File → New / Open: replace the document with a blank one or a local file (Office, CSV, text, Markdown, HTML). Default true. */
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
}

const EXT: Record<string, string> = { docx: '.docx', xlsx: '.xlsx', pptx: '.pptx' }
const UNTITLED: Record<DocKind, string> = { docx: 'Untitled document', xlsx: 'Untitled spreadsheet', pptx: 'Untitled presentation' }

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
  const { client, error, loading, progress } = useDocument(opened ? opened.src : src, { name: loadName, headers, credentials, worker })
  const docName = renamed ?? loadName
  const collab = useCollab(client, props.collab)
  // with collaboration, editing waits for the shared history (offline edits sync on reconnect)
  const canEdit = !props.collab || collab.status === 'live' || collab.status === 'offline'
  const [statusEl, setStatusEl] = useState<HTMLSpanElement | null>(null)

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
  const api = useMemo<ViewerApi>(
    () => ({ bar: barEl, editable: editing && canEdit, setDirty, save, canDownload: allowDownload, client, fileOps, selectAll }),
    [barEl, editing, canEdit, setDirty, save, allowDownload, client],
  )
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
  } else if (client.kind === 'docx') body = <DocxView key={String(client.name) + client.size} client={client} initialZoom={zoom} />
  else if (client.kind === 'xlsx') body = <XlsxView client={client} />
  else body = <PptxView client={client} initialZoom={zoom} thumbnails={thumbnails} />

  return (
    <div ref={rootRef} className={'de-root' + (dark ? ' de-dark' : '') + (className ? ` ${className}` : '')} style={style} onKeyDown={onKeyDown}>
      {toolbar && (
        <div className="de-bar">
          {client && <span className={`de-badge de-badge-${client.kind}`}>{client.kind}</span>}
          {fileMenu && <FileMenu actions={fileActions} />}
          <span className="de-title" title={title}>
            {title}
          </span>
          <span className="de-status" ref={setStatusEl} hidden={!client} />
          {client && client.size > 0 && <span className="de-status">{fmtBytes(client.size)}</span>}
          {client && props.collab && <CollabBar client={client} state={collab} />}
          {toolbarExtra}
          {client && allowEdit && (
            <button type="button" className={editing ? 'de-mode de-mode-on' : 'de-mode'} onClick={() => setEditing((v) => !v)} title={editing ? 'Switch to viewing' : 'Edit document'}>
              <Icon name={editing ? 'view' : 'edit'} size={16} />
              {editing ? 'View' : 'Edit'}
            </button>
          )}
          {client && canSave && (editing || dirty) && (
            <button type="button" className={dirty ? 'de-mode de-save-dirty' : 'de-mode'} disabled={saving} onClick={() => void save()} title="Save a copy (Ctrl+S)">
              <Icon name="save" size={16} />
              {saving ? 'Saving…' : dirty || onSave ? 'Save' : 'Download'}
            </button>
          )}
          <div className="de-progress" style={{ width: `${busy ? Math.max(pct, 3) : 100}%`, opacity: busy ? 1 : 0 }} />
        </div>
      )}
      {toolbar && client && editing && canEdit && <div className="de-ebar-host" ref={setBarEl} />}
      <div className="de-body">
        <ViewerContext.Provider value={api}>
          <StatusContext.Provider value={statusEl}>{body}</StatusContext.Provider>
        </ViewerContext.Provider>
        {client && collab.collab && <PeerOverlay client={client} collab={collab.collab} />}
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
            ['Type', KIND_LABEL[client.kind]],
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
