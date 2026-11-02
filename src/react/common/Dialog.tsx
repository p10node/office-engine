import { useEffect, useRef, useState, type ReactNode } from 'react'

export interface PromptSpec {
  title: string
  label?: string
  initial?: string
  placeholder?: string
  ok?: string
  /** multi-line input */
  area?: boolean
  onOk: (value: string) => void
  /** extra button (e.g. "Remove link") */
  extra?: { label: string; run: () => void }
  note?: ReactNode
}

/** Small modal prompt (no window.prompt: works in sandboxed / embedded pages). */
export function PromptDialog({ spec, onClose }: { spec: PromptSpec; onClose: () => void }) {
  const [v, setV] = useState(spec.initial ?? '')
  const ref = useRef<HTMLInputElement & HTMLTextAreaElement>(null)
  useEffect(() => {
    const el = ref.current
    el?.focus()
    el?.select()
  }, [])
  const done = () => {
    onClose()
    spec.onOk(v)
  }
  return (
    <div
      className="de-dialog-back"
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
      onKeyDown={(e) => {
        if (e.key === 'Escape') {
          e.stopPropagation()
          onClose()
        }
      }}
    >
      <form
        className="de-dialog"
        role="dialog"
        aria-label={spec.title}
        onSubmit={(e) => {
          e.preventDefault()
          done()
        }}
      >
        <div className="de-dialog-title">{spec.title}</div>
        {spec.label && <label className="de-dialog-label">{spec.label}</label>}
        {spec.area ? (
          <textarea ref={ref} rows={3} value={v} placeholder={spec.placeholder} onChange={(e) => setV(e.target.value)} />
        ) : (
          <input ref={ref} value={v} placeholder={spec.placeholder} onChange={(e) => setV(e.target.value)} />
        )}
        {spec.note && <div className="de-dialog-note">{spec.note}</div>}
        <div className="de-dialog-btns">
          {spec.extra && (
            <button
              type="button"
              className="de-btn de-dialog-extra"
              onClick={() => {
                onClose()
                spec.extra!.run()
              }}
            >
              {spec.extra.label}
            </button>
          )}
          <button type="button" className="de-btn" onClick={onClose}>
            Cancel
          </button>
          <button type="submit" className="de-btn de-btn-primary">
            {spec.ok ?? 'OK'}
          </button>
        </div>
      </form>
    </div>
  )
}

/** Short-lived message at the bottom of the viewer. */
export function Toast({ text, onDone }: { text: string; onDone: () => void }) {
  useEffect(() => {
    const id = setTimeout(onDone, 3500)
    return () => clearTimeout(id)
  }, [text, onDone])
  return (
    <div className="de-toast" role="status">
      {text}
    </div>
  )
}
