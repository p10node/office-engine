import { useEffect, useRef, useState } from 'react'
import type { DocSource } from '../core/source'
import { DocumentClient, type OpenOptions } from '../core/worker/client'
import type { Events } from '../core/worker/protocol'

export interface DocumentState {
  client: DocumentClient | null
  error: Error | null
  loading: boolean
  progress: Events['progress'] | null
}

/**
 * Opens a document in a worker and closes it on unmount / source change.
 * `keep`: on a source change, the document shown stays (and usable) until the next one is open — the live preview
 * of an edited source.
 */
export function useDocument(src: DocSource | null | undefined, options: OpenOptions = {}, keep = false): DocumentState {
  const [state, setState] = useState<DocumentState>({ client: null, error: null, loading: !!src, progress: null })
  const { name, worker, credentials, language } = options
  const headersKey = options.headers ? JSON.stringify(options.headers) : ''
  const mailKey = options.mail ? JSON.stringify(options.mail) : ''
  const shown = useRef<DocumentClient | null>(null)
  const keepRef = useRef(keep)
  keepRef.current = keep
  /** the client shown from now on; the one before it is closed */
  const show = (c: DocumentClient | null) => {
    if (shown.current && shown.current !== c) shown.current.destroy()
    shown.current = c
  }
  useEffect(() => () => show(null), [])
  useEffect(() => {
    if (!src) {
      show(null)
      setState({ client: null, error: null, loading: false, progress: null })
      return
    }
    let live = true
    if (keepRef.current && shown.current) setState((s) => ({ ...s, error: null, loading: true }))
    else {
      show(null)
      setState({ client: null, error: null, loading: true, progress: null })
    }
    // a load given up (another source, the viewer gone) ends its worker: a large file is not downloaded and read to the end for nothing
    const giveUp = new AbortController()
    DocumentClient.open(src, {
      name,
      worker,
      credentials,
      language,
      headers: headersKey ? JSON.parse(headersKey) : undefined,
      mail: mailKey ? JSON.parse(mailKey) : undefined,
      onProgress: (p) => live && setState((s) => ({ ...s, progress: p })),
      signal: giveUp.signal,
    }).then(
      (c) => {
        if (!live) return c.destroy()
        show(c)
        c.on('progress', (p) => live && shown.current === c && setState((s) => ({ ...s, progress: p })))
        setState((s) => ({ ...s, client: c, loading: false }))
      },
      (e: Error) => {
        if (!live) return
        show(null)
        setState({ client: null, error: e, loading: false, progress: null })
      },
    )
    return () => {
      live = false
      giveUp.abort()
    }
  }, [src, name, worker, credentials, headersKey, language, mailKey])
  return state
}
