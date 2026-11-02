import { useEffect, useLayoutEffect, useRef, useState } from 'react'
import { Icon } from '../edit/Toolbar'

export interface MenuItem {
  label: string
  run: () => void
  disabled?: boolean
  /** separator above this item */
  sep?: boolean
  /** shortcut / secondary text */
  hint?: string
  icon?: string
  /** small section title above this item */
  head?: string
}

/** Right-click menu at viewport coordinates; never steals focus from the editor. */
export function ContextMenu({ x, y, items, onClose }: { x: number; y: number; items: MenuItem[]; onClose: () => void }) {
  const ref = useRef<HTMLDivElement>(null)
  const [pos, setPos] = useState({ left: x, top: y })
  const [hi, setHi] = useState(-1)
  const close = useRef(onClose)
  close.current = onClose

  // keep inside the viewport (flip up / left when needed)
  useLayoutEffect(() => {
    const el = ref.current
    if (!el) return
    const r = el.getBoundingClientRect()
    const vw = document.documentElement.clientWidth
    const vh = document.documentElement.clientHeight
    setPos({
      left: x + r.width > vw - 4 ? Math.max(4, x - r.width) : x,
      top: y + r.height > vh - 4 ? Math.max(4, Math.min(y - r.height, vh - r.height - 4)) : y,
    })
  }, [x, y])

  useEffect(() => {
    const down = (e: MouseEvent) => {
      if (ref.current && !e.composedPath().includes(ref.current)) close.current()
    }
    const key = (e: KeyboardEvent) => {
      const n = items.length
      const step = (from: number, d: number) => {
        let i = from
        for (let k = 0; k < n; k++) {
          i = (i + d + n) % n
          if (!items[i].disabled) return i
        }
        return from
      }
      if (e.key === 'Escape') close.current()
      else if (e.key === 'ArrowDown') setHi((h) => step(h, 1))
      else if (e.key === 'ArrowUp') setHi((h) => step(h < 0 ? 0 : h, -1))
      else if (e.key === 'Enter' && hi >= 0 && !items[hi].disabled) {
        close.current()
        items[hi].run()
      } else return
      e.preventDefault()
      e.stopPropagation()
    }
    const away = () => close.current()
    window.addEventListener('mousedown', down, true)
    window.addEventListener('keydown', key, true)
    window.addEventListener('blur', away)
    window.addEventListener('resize', away)
    return () => {
      window.removeEventListener('mousedown', down, true)
      window.removeEventListener('keydown', key, true)
      window.removeEventListener('blur', away)
      window.removeEventListener('resize', away)
    }
  }, [items, hi])

  return (
    <div ref={ref} className="de-ctx" style={pos} role="menu" onMouseDown={(e) => e.preventDefault()} onContextMenu={(e) => e.preventDefault()}>
      {items.map((it, i) => (
        <div key={i} className="de-ctx-row">
          {it.sep && i > 0 && <div className="de-ctx-sep" role="separator" />}
          {it.head && <div className="de-ctx-head">{it.head}</div>}
          <button
            type="button"
            role="menuitem"
            className={i === hi ? 'de-ctx-hi' : undefined}
            disabled={it.disabled}
            onMouseEnter={() => setHi(i)}
            onClick={() => {
              close.current()
              it.run()
            }}
          >
            <span className="de-ctx-icon">{it.icon && <Icon name={it.icon} size={16} />}</span>
            <span className="de-ctx-label">{it.label}</span>
            {it.hint && <span className="de-ctx-hint">{it.hint}</span>}
          </button>
        </div>
      ))}
    </div>
  )
}

/** platform modifier label for shortcut hints */
export const MOD = typeof navigator !== 'undefined' && /Mac|iPhone|iPad/.test(navigator.platform) ? '⌘' : 'Ctrl+'
