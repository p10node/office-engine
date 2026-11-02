import { StrictMode, useCallback, useMemo, useState } from 'react'
import { createRoot } from 'react-dom/client'
import { DocViewer, type DocSource } from '../src'
import type { CollabOptions } from '../src/collab'

const SAMPLES = [
  { label: 'sample.docx', url: 'samples/sample.docx' },
  { label: 'sample.xlsx', url: 'samples/sample.xlsx' },
  { label: 'sample.pptx', url: 'samples/sample.pptx' },
  { label: 'large-doc.docx', url: 'samples/large-doc.docx' },
  { label: 'large-sheet.xlsx (300k rows)', url: 'samples/large-sheet.xlsx' },
  { label: 'large-deck.pptx (305 slides)', url: 'samples/large-deck.pptx' },
]

const bar: React.CSSProperties = {
  display: 'flex', flexWrap: 'wrap', gap: 8, alignItems: 'center', padding: '10px 16px',
  borderBottom: '1px solid #dadde1', background: '#fff',
}
const btn: React.CSSProperties = {
  border: '1px solid #d0d4d9', background: '#fff', borderRadius: 6, padding: '5px 10px', cursor: 'pointer', font: 'inherit',
}

// on GitHub Pages the demo lives at <site>/demo/, docs at <site>/
const DOCS_URL = import.meta.env.DEV ? 'https://p10node.github.io/office-engine/' : '../'

// collaboration server (bun run collab): ?collab=http://host:port overrides it
const params = new URLSearchParams(location.search)
const COLLAB_HTTP = params.get('collab') ?? `${location.protocol}//${location.hostname}:8787`
const COLLAB_WS = COLLAB_HTTP.replace(/^http/, 'ws') + '/collab'
const roomFile = (room: string) => `${COLLAB_HTTP}/rooms/${encodeURIComponent(room)}/file`

function userName(): string {
  try {
    const saved = localStorage.getItem('de-user')
    if (saved) return saved
    const n = `Guest ${Math.floor(100 + Math.random() * 900)}`
    localStorage.setItem('de-user', n)
    return n
  } catch {
    return 'Guest'
  }
}

async function sourceBytes(s: DocSource): Promise<Blob> {
  if (typeof s === 'string' || s instanceof URL) return (await fetch(String(s))).blob()
  if (s instanceof Blob) return s
  return new Blob([s as ArrayBuffer])
}

function App() {
  const initialRoom = params.get('room')
  const initial = initialRoom ? roomFile(initialRoom) : params.get('src')
  const [src, setSrc] = useState<DocSource | null>(initial)
  const [room, setRoom] = useState<string | null>(initialRoom)
  const [user, setUser] = useState(userName)
  const [sharing, setSharing] = useState<string | null>(null)
  const [url, setUrl] = useState('')
  const [drag, setDrag] = useState(false)
  const [loadMs, setLoadMs] = useState<number | null>(null)
  const [t0, setT0] = useState(() => performance.now())

  const open = (s: DocSource, inRoom: string | null = null) => {
    setT0(performance.now())
    setLoadMs(null)
    setSrc(s)
    setRoom(inRoom)
    const q = new URLSearchParams(location.search)
    if (inRoom) q.set('room', inRoom)
    else q.delete('room')
    history.replaceState(null, '', q.size ? `?${q}` : location.pathname)
  }
  /** uploads the opened file as the shared base and joins its room */
  const share = async () => {
    if (!src) return
    setSharing('Uploading…')
    try {
      const blob = await sourceBytes(src)
      const id = Math.random().toString(36).slice(2, 10)
      const name = (src as File).name ?? (typeof src === 'string' ? decodeURIComponent(src.split(/[?#]/)[0].split('/').pop() ?? '') : 'document')
      const r = await fetch(roomFile(id), { method: 'PUT', body: blob, headers: { 'X-File-Name': encodeURIComponent(name) } })
      if (!r.ok) throw new Error(`HTTP ${r.status}`)
      open(roomFile(id), id)
      setSharing(null)
    } catch (e) {
      setSharing(`Share failed: is the collab server running (bun run collab)? ${(e as Error).message}`)
    }
  }
  const collab = useMemo<CollabOptions | undefined>(() => (room ? { transport: COLLAB_WS, room, user: { name: user } } : undefined), [room, user])
  const link = room ? `${location.origin}${location.pathname}?room=${room}${params.get('collab') ? `&collab=${encodeURIComponent(COLLAB_HTTP)}` : ''}` : ''
  const onLoad = useCallback(() => setLoadMs(Math.round(performance.now() - t0)), [t0])

  return (
    <div
      style={{ display: 'flex', flexDirection: 'column', height: '100%' }}
      onDragOver={(e) => { e.preventDefault(); setDrag(true) }}
      onDragLeave={(e) => { if (e.currentTarget === e.target) setDrag(false) }}
      onDrop={(e) => {
        e.preventDefault()
        setDrag(false)
        const f = e.dataTransfer.files[0]
        if (f) open(f)
      }}
    >
      <div style={bar}>
        <strong>office-engine</strong>
        <a href={DOCS_URL} style={{ marginRight: 8, color: '#1a73e8', textDecoration: 'none' }}>Docs</a>
        {SAMPLES.map((s) => (
          <button key={s.url} style={btn} onClick={() => open(s.url)}>{s.label}</button>
        ))}
        <label style={{ ...btn, display: 'inline-block' }}>
          Open file…
          <input type="file" accept=".docx,.xlsx,.pptx,.docm,.xlsm,.pptm" hidden onChange={(e) => e.target.files?.[0] && open(e.target.files[0])} />
        </label>
        <form onSubmit={(e) => { e.preventDefault(); if (url) open(url) }} style={{ display: 'flex', gap: 6, flex: '1 1 260px' }}>
          <input value={url} onChange={(e) => setUrl(e.target.value)} placeholder="https://…/file.xlsx (Range requests)" style={{ ...btn, flex: 1, cursor: 'text' }} />
          <button style={btn}>Open URL</button>
        </form>
        {loadMs !== null && <span style={{ color: '#69707a' }}>opened in {loadMs} ms</span>}
        <span style={{ flex: '1 0 0' }} />
        <input
          value={user}
          onChange={(e) => setUser(e.target.value)}
          onBlur={() => {
            try {
              localStorage.setItem('de-user', user)
            } catch {
              /* private mode */
            }
          }}
          title="Your name for collaborators"
          style={{ ...btn, width: 110, cursor: 'text' }}
        />
        {room ? (
          <>
            <button style={btn} onClick={() => void navigator.clipboard?.writeText(link)} title={link}>
              Copy share link
            </button>
            <button style={btn} onClick={() => open(src!, null)}>Stop sharing</button>
          </>
        ) : (
          <button style={{ ...btn, borderColor: '#1a73e8', color: '#1a73e8' }} disabled={!src} onClick={() => void share()} title="Edit together in real time">
            Share live
          </button>
        )}
        {sharing && <span style={{ color: sharing.startsWith('Share failed') ? '#c5221f' : '#69707a' }}>{sharing}</span>}
      </div>
      <div style={{ flex: 1, minHeight: 0, position: 'relative' }}>
        <DocViewer src={src} onLoad={onLoad} editable collab={collab} />
        {drag && (
          <div style={{ position: 'absolute', inset: 12, border: '2px dashed #1a73e8', borderRadius: 12, background: 'rgba(26,115,232,.06)', display: 'grid', placeItems: 'center', fontSize: 16, color: '#1a73e8', pointerEvents: 'none' }}>
            Drop a .docx, .xlsx or .pptx file
          </div>
        )}
      </div>
    </div>
  )
}

createRoot(document.getElementById('root')!).render(
  <StrictMode>
    <App />
  </StrictMode>,
)
