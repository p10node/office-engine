/**
 * Fixture for the browser tests: DocViewer configured from the query string, callbacks recorded on window.
 *   ?src=samples/sample.docx&editable=1&allowSave=0&onSave=1&items=1&close=1&collab=ws://…&room=r
 */
import { useState } from 'react'
import { createRoot } from 'react-dom/client'
import { DocViewer, type DocViewerProps, type FileMenuItem } from '../../src'

type Ev = { type: string; [k: string]: unknown }
declare global {
  interface Window {
    __events: Ev[]
    __setProps: (p: Partial<DocViewerProps>) => void
  }
}
window.__events = []
const log = (e: Ev) => window.__events.push(e)

const q = new URLSearchParams(location.search)
const flag = (k: string) => (q.has(k) ? q.get(k) !== '0' : undefined)

async function b64(blob: Blob) {
  const bytes = new Uint8Array(await blob.arrayBuffer())
  let s = ''
  for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode(...bytes.subarray(i, i + 0x8000))
  return btoa(s)
}

const items: FileMenuItem[] = [
  { label: 'Share…', icon: 'open', onClick: () => log({ type: 'item', label: 'Share' }) },
  { label: 'Version history', hint: 'beta', onClick: () => log({ type: 'item', label: 'Version history' }) },
]

const initial: DocViewerProps = {
  src: q.get('src'),
  name: q.get('name') ?? undefined,
  style: { height: '100vh' },
  toolbar: flag('toolbar'),
  editable: flag('editable'),
  allowEdit: flag('allowEdit'),
  allowSave: flag('allowSave'),
  allowDownload: flag('allowDownload'),
  fileMenu: flag('fileMenu'),
  allowOpen: flag('allowOpen'),
  thumbnails: flag('thumbnails'),
  onSave: q.has('onSave') ? async (blob, name) => void log({ type: 'save', name, b64: await b64(blob) }) : undefined,
  onDirtyChange: (dirty) => log({ type: 'dirty', dirty }),
  onRename: (name) => log({ type: 'rename', name }),
  onClose: q.has('close') ? () => log({ type: 'close' }) : undefined,
  fileMenuItems: q.has('items') ? items : undefined,
  onLoad: (client) => log({ type: 'load', kind: client.kind, name: client.name }),
  onError: (error) => log({ type: 'error', message: error.message }),
  collab: q.get('collab') ? { transport: q.get('collab')!, room: q.get('room') ?? 'e2e', base: Number(q.get('base')) || 0, user: { name: q.get('user') ?? 'Tester' } } : undefined,
}

function App() {
  const [props, setProps] = useState(initial)
  window.__setProps = (p) => setProps((cur) => ({ ...cur, ...p }))
  return <DocViewer {...props} />
}

createRoot(document.getElementById('root')!).render(<App />)
