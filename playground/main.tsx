import { StrictMode, useCallback, useState } from 'react'
import { createRoot } from 'react-dom/client'
import { DocViewer, type DocSource } from '../src'

const SAMPLES = [
  { label: 'sample.docx', url: '/samples/sample.docx' },
  { label: 'sample.xlsx', url: '/samples/sample.xlsx' },
  { label: 'sample.pptx', url: '/samples/sample.pptx' },
  { label: 'large-doc.docx', url: '/samples/large-doc.docx' },
  { label: 'large-sheet.xlsx (300k rows)', url: '/samples/large-sheet.xlsx' },
  { label: 'large-deck.pptx (305 slides)', url: '/samples/large-deck.pptx' },
]

const bar: React.CSSProperties = {
  display: 'flex', flexWrap: 'wrap', gap: 8, alignItems: 'center', padding: '10px 16px',
  borderBottom: '1px solid #dadde1', background: '#fff',
}
const btn: React.CSSProperties = {
  border: '1px solid #d0d4d9', background: '#fff', borderRadius: 6, padding: '5px 10px', cursor: 'pointer', font: 'inherit',
}

function App() {
  const initial = new URLSearchParams(location.search).get('src')
  const [src, setSrc] = useState<DocSource | null>(initial)
  const [url, setUrl] = useState('')
  const [drag, setDrag] = useState(false)
  const [loadMs, setLoadMs] = useState<number | null>(null)
  const [t0, setT0] = useState(() => performance.now())

  const open = (s: DocSource) => {
    setT0(performance.now())
    setLoadMs(null)
    setSrc(s)
  }
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
        <strong style={{ marginRight: 8 }}>docs-engine</strong>
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
      </div>
      <div style={{ flex: 1, minHeight: 0, position: 'relative' }}>
        <DocViewer src={src} onLoad={onLoad} />
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
