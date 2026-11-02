import { useCallback, useEffect, useInsertionEffect, useMemo, useState, type CSSProperties, type ReactNode } from 'react'
import type { DocSource } from '../core/source'
import type { DocumentClient, OpenOptions } from '../core/worker/client'
import type { Zoom } from './common/ZoomControl'
import { DocxView } from './docx/DocxView'
import { ViewerContext, type ViewerApi } from './common/Slots'
import { StatusContext } from './common/Status'
import { download, Icon } from './edit/Toolbar'
import { injectStyles } from './injectStyles'
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
  onLoad?: (client: DocumentClient) => void
  onError?: (error: Error) => void
  /** Start in edit mode. Default false (view). */
  editable?: boolean
  /** Show the Edit/View toggle. Default true. */
  allowEdit?: boolean
  /** Receives the saved file instead of downloading it. */
  onSave?: (file: Blob, name: string) => void | Promise<void>
  /** Called when the document gets / loses unsaved changes. */
  onDirtyChange?: (dirty: boolean) => void
  /**
   * Real-time collaboration: everyone opening the same base document (`src`) with the same `room`
   * edits it together through the hub at `transport` (see src/collab).
   */
  collab?: CollabOptions
}

const EXT: Record<string, string> = { docx: '.docx', xlsx: '.xlsx', pptx: '.pptx' }

function fmtBytes(n: number) {
  if (n < 1024) return `${n} B`
  if (n < 1024 * 1024) return `${(n / 1024).toFixed(0)} KB`
  return `${(n / 1024 / 1024).toFixed(1)} MB`
}

export function DocViewer(props: DocViewerProps) {
  const { src, name, className, style, toolbar = true, toolbarExtra, zoom, thumbnails, headers, credentials, worker, styleTarget, onLoad, onError } = props
  const { allowEdit = true, onSave, onDirtyChange } = props
  const [editing, setEditing] = useState(!!props.editable)
  useEffect(() => {
    if (props.editable !== undefined) setEditing(!!props.editable)
  }, [props.editable])
  const [dirty, setDirtyState] = useState(false)
  const [saving, setSaving] = useState(false)
  const [barEl, setBarEl] = useState<HTMLDivElement | null>(null)
  useInsertionEffect(() => injectStyles(styleTarget), [styleTarget])
  const { client, error, loading, progress } = useDocument(src, { name, headers, credentials, worker })
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

  useEffect(() => {
    setDirtyState(false)
  }, [client])
  const setDirty = useCallback(
    (d: boolean) =>
      setDirtyState((prev) => {
        if (prev !== d) onDirtyChange?.(d)
        return d
      }),
    [onDirtyChange],
  )
  const fileName = (client?.name || name || 'document').replace(/\.[^.]+$/, '') + (client ? EXT[client.kind] : '')
  const save = useCallback(async () => {
    if (!client || saving) return
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
  }, [client, saving, onSave, fileName, setDirty, onError])
  const api = useMemo<ViewerApi>(() => ({ bar: barEl, editable: editing && canEdit, setDirty, save, client }), [barEl, editing, canEdit, setDirty, save, client])

  const pct = progress && progress.total > 0 ? Math.min(100, (progress.loaded / progress.total) * 100) : 0
  const busy = loading || (!!progress && pct < 100)
  const title = client?.name ?? name ?? (typeof src === 'string' ? decodeURIComponent(src.split(/[?#]/)[0].split('/').pop() ?? '') : (src as File | undefined)?.name) ?? ''

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
    body = src ? (
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
    <div className={className ? `de-root ${className}` : 'de-root'} style={style}>
      {toolbar && (
        <div className="de-bar">
          {client && <span className={`de-badge de-badge-${client.kind}`}>{client.kind}</span>}
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
          {client && (editing || dirty) && (
            <button type="button" className={dirty ? 'de-mode de-save-dirty' : 'de-mode'} disabled={saving} onClick={() => void save()} title="Save a copy (Ctrl+S)">
              <Icon name="save" size={16} />
              {saving ? 'Saving…' : dirty ? 'Save' : 'Download'}
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
    </div>
  )
}
