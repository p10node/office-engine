import { useEffect, useRef, useState, type ReactNode } from 'react'
import { Btn } from '../edit/Toolbar'

export interface Match {
  pa: string
  a: number
  b: number
  bi: number
}

export function FindBar(props: {
  query: string
  index: number
  count: number
  editable: boolean
  onQuery: (q: string) => void
  onNav: (d: 1 | -1) => void
  onClose: () => void
  onReplace: (text: string, all: boolean) => void
  placeholder?: string
  /** extra controls (match case…) under the search row */
  options?: ReactNode
  /** open with the replace row visible */
  replace?: boolean
}) {
  const ref = useRef<HTMLInputElement>(null)
  const [rep, setRep] = useState('')
  const [showRep, setShowRep] = useState(!!props.replace)
  useEffect(() => {
    ref.current?.focus()
    ref.current?.select()
  }, [])
  return (
    <div className="de-find" role="search" onKeyDown={(e) => e.key === 'Escape' && props.onClose()}>
      <div className="de-find-row">
        <input
          ref={ref}
          value={props.query}
          placeholder={props.placeholder ?? 'Find in document'}
          aria-label="Find"
          onChange={(e) => props.onQuery(e.target.value)}
          onKeyDown={(e) => {
            if (e.key === 'Enter') {
              e.preventDefault()
              props.onNav(e.shiftKey ? -1 : 1)
            }
          }}
        />
        <span className="de-find-count">{props.query ? (props.count ? `${Math.min(props.index + 1, props.count)} / ${props.count}` : '0 / 0') : ''}</span>
        <Btn icon="up" title="Previous (Shift+Enter)" onClick={() => props.onNav(-1)} disabled={!props.count} />
        <Btn icon="down" title="Next (Enter)" onClick={() => props.onNav(1)} disabled={!props.count} />
        {props.editable && <Btn label="⇄" title="Replace" active={showRep} onClick={() => setShowRep((v) => !v)} />}
        <Btn icon="close" title="Close (Esc)" onClick={props.onClose} />
      </div>
      {props.options && <div className="de-find-row de-find-opts">{props.options}</div>}
      {props.editable && showRep && (
        <div className="de-find-row">
          <input value={rep} placeholder="Replace with" aria-label="Replace with" onChange={(e) => setRep(e.target.value)} />
          <button type="button" className="de-btn" disabled={!props.count} onClick={() => props.onReplace(rep, false)}>
            Replace
          </button>
          <button type="button" className="de-btn" disabled={!props.count} onClick={() => props.onReplace(rep, true)}>
            All
          </button>
        </div>
      )}
    </div>
  )
}
