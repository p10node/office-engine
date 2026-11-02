import { useEffect, useState } from 'react'
import type { DocSource } from '../core/source'
import { DocumentClient, type OpenOptions } from '../core/worker/client'
import type { Events } from '../core/worker/protocol'

export interface DocumentState {
  client: DocumentClient | null
  error: Error | null
  loading: boolean
  progress: Events['progress'] | null
}

/** Opens a document in a worker and closes it on unmount / source change. */
export function useDocument(src: DocSource | null | undefined, options: OpenOptions = {}): DocumentState {
  const [state, setState] = useState<DocumentState>({ client: null, error: null, loading: !!src, progress: null })
  const { name, worker, credentials } = options
  const headersKey = options.headers ? JSON.stringify(options.headers) : ''
  useEffect(() => {
    if (!src) {
      setState({ client: null, error: null, loading: false, progress: null })
      return
    }
    let live = true
    let client: DocumentClient | null = null
    setState({ client: null, error: null, loading: true, progress: null })
    DocumentClient.open(src, {
      name,
      worker,
      credentials,
      headers: headersKey ? JSON.parse(headersKey) : undefined,
      onProgress: (p) => live && setState((s) => ({ ...s, progress: p })),
    }).then(
      (c) => {
        if (!live) return c.destroy()
        client = c
        c.on('progress', (p) => live && setState((s) => ({ ...s, progress: p })))
        setState((s) => ({ ...s, client: c, loading: false }))
      },
      (e: Error) => live && setState({ client: null, error: e, loading: false, progress: null }),
    )
    return () => {
      live = false
      client?.destroy()
    }
  }, [src, name, worker, credentials, headersKey])
  return state
}
