import { useEffect, useInsertionEffect, useState, type CSSProperties, type ReactNode } from 'react'
import type { DocSource } from '../core/source'
import type { DocumentClient, OpenOptions } from '../core/worker/client'
import type { Zoom } from './common/ZoomControl'
import { DocxView } from './docx/DocxView'
import { StatusContext } from './common/Status'
import { injectStyles } from './injectStyles'
import { PptxView } from './pptx/PptxView'
import { useDocument } from './useDocument'
import { XlsxView } from './xlsx/XlsxView'

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
}

function fmtBytes(n: number) {
  if (n < 1024) return `${n} B`
  if (n < 1024 * 1024) return `${(n / 1024).toFixed(0)} KB`
  return `${(n / 1024 / 1024).toFixed(1)} MB`
}

export function DocViewer(props: DocViewerProps) {
  const { src, name, className, style, toolbar = true, toolbarExtra, zoom, thumbnails, headers, credentials, worker, styleTarget, onLoad, onError } = props
  useInsertionEffect(() => injectStyles(styleTarget), [styleTarget])
  const { client, error, loading, progress } = useDocument(src, { name, headers, credentials, worker })
  const [statusEl, setStatusEl] = useState<HTMLSpanElement | null>(null)

  useEffect(() => {
    if (client) onLoad?.(client)
  }, [client, onLoad])
  useEffect(() => {
    if (error) onError?.(error)
  }, [error, onError])

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
          {toolbarExtra}
          <div className="de-progress" style={{ width: `${busy ? Math.max(pct, 3) : 100}%`, opacity: busy ? 1 : 0 }} />
        </div>
      )}
      <div className="de-body">
        <StatusContext.Provider value={statusEl}>{body}</StatusContext.Provider>
      </div>
    </div>
  )
}
