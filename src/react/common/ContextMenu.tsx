import { useEffect, useLayoutEffect, useRef, useState } from 'react'
import { Icon } from '../edit/Toolbar'

export interface MenuItem {
  label: string
  /** what picking the item does (an item with `items` opens them instead) */
  run?: () => void
  disabled?: boolean
  /** separator above this item */
  sep?: boolean
  /** shortcut / secondary text */
  hint?: string
  icon?: string
  /** small section title above this item */
  head?: string
  /** submenu: the item opens these beside it */
  items?: MenuItem[]
}

/** next enabled item from `from` in direction `d` (wraps); `from` itself when there is none */
function step(items: MenuItem[], from: number, d: number): number {
  const n = items.length
  let i = from
  for (let k = 0; k < n; k++) {
    i = (i + d + n) % n
    if (!items[i].disabled) return i
  }
  return from
}

function Rows(props: { items: MenuItem[]; hi: number; open?: number; onHover: (i: number) => void; onPick: (i: number) => void; rowRef?: (i: number, el: HTMLButtonElement | null) => void }) {
  const { items, hi, open } = props
  return (
    <>
      {items.map((it, i) => (
        <div key={i} className="de-ctx-row">
          {it.sep && i > 0 && <div className="de-ctx-sep" role="separator" />}
          {it.head && <div className="de-ctx-head">{it.head}</div>}
          <button
            ref={props.rowRef && ((el) => props.rowRef!(i, el))}
            type="button"
            role="menuitem"
            aria-haspopup={it.items ? 'menu' : undefined}
            aria-expanded={it.items ? open === i : undefined}
            className={i === hi || i === open ? 'de-ctx-hi' : undefined}
            disabled={it.disabled}
            onMouseEnter={() => props.onHover(i)}
            onClick={() => props.onPick(i)}
          >
            <span className="de-ctx-icon">{it.icon && <Icon name={it.icon} size={16} />}</span>
            <span className="de-ctx-label">{it.label}</span>
            {it.hint && <span className="de-ctx-hint">{it.hint}</span>}
            {it.items && (
              <span className="de-ctx-more">
                <Icon name="chevron" size={14} />
              </span>
            )}
          </button>
        </div>
      ))}
    </>
  )
}

/** submenu panel beside its row, kept inside the viewport (opens to the left when there is no room) */
function Sub(props: { anchor: HTMLElement; items: MenuItem[]; hi: number; onHover: (i: number) => void; onPick: (i: number) => void; onEnter: () => void }) {
  const ref = useRef<HTMLDivElement>(null)
  const [pos, setPos] = useState<{ left: number; top: number } | null>(null)
  useLayoutEffect(() => {
    const el = ref.current
    if (!el) return
    const a = props.anchor.getBoundingClientRect()
    const r = el.getBoundingClientRect()
    const vw = document.documentElement.clientWidth
    const vh = document.documentElement.clientHeight
    setPos({
      left: a.right + r.width > vw - 4 ? Math.max(4, a.left - r.width) : a.right,
      top: Math.max(4, Math.min(a.top - 6, vh - r.height - 4)),
    })
  }, [props.anchor, props.items.length])
  return (
    <div ref={ref} className="de-ctx de-ctx-sub" style={pos ?? { left: 0, top: 0, visibility: 'hidden' }} role="menu" onMouseEnter={props.onEnter}>
      <Rows items={props.items} hi={props.hi} onHover={props.onHover} onPick={props.onPick} />
    </div>
  )
}

/** Right-click menu at viewport coordinates; never steals focus from the editor. */
export function ContextMenu({ x, y, items, onClose }: { x: number; y: number; items: MenuItem[]; onClose: () => void }) {
  const ref = useRef<HTMLDivElement>(null)
  const rows = useRef<(HTMLButtonElement | null)[]>([])
  const [pos, setPos] = useState({ left: x, top: y })
  const [hi, setHi] = useState(-1)
  /** the open submenu: its row, and the item highlighted in it */
  const [sub, setSub] = useState<{ at: number; hi: number } | null>(null)
  const close = useRef(onClose)
  close.current = onClose
  /** closes the submenu a moment after the pointer left its row (time to reach the panel across other rows) */
  const leave = useRef<ReturnType<typeof setTimeout> | undefined>(undefined)
  const stay = () => clearTimeout(leave.current)
  useEffect(() => stay, [])

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

  const subItems = sub ? items[sub.at]?.items : undefined
  const openSub = (at: number, first = -1) => {
    stay()
    setSub((s) => (s?.at === at && first < 0 ? s : { at, hi: first }))
  }
  const pick = (list: MenuItem[], i: number, at?: number) => {
    const it = list[i]
    if (!it || it.disabled) return
    if (it.items && at === undefined) return openSub(i)
    close.current()
    it.run?.()
  }

  useEffect(() => {
    const down = (e: MouseEvent) => {
      if (ref.current && !e.composedPath().includes(ref.current)) close.current()
    }
    const key = (e: KeyboardEvent) => {
      if (e.key === 'Escape') {
        if (subItems) setSub(null)
        else close.current()
      } else if (subItems && sub) {
        if (e.key === 'ArrowDown') setSub({ at: sub.at, hi: step(subItems, sub.hi, 1) })
        else if (e.key === 'ArrowUp') setSub({ at: sub.at, hi: step(subItems, sub.hi < 0 ? 0 : sub.hi, -1) })
        else if (e.key === 'ArrowLeft') setSub(null)
        else if (e.key === 'Enter') pick(subItems, sub.hi, sub.at)
        else return
      } else if (e.key === 'ArrowDown') setHi((h) => step(items, h, 1))
      else if (e.key === 'ArrowUp') setHi((h) => step(items, h < 0 ? 0 : h, -1))
      else if ((e.key === 'ArrowRight' || e.key === 'Enter') && items[hi]?.items && !items[hi].disabled) openSub(hi, step(items[hi].items!, -1, 1))
      else if (e.key === 'Enter') pick(items, hi)
      else return
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
  })

  const anchor = sub ? rows.current[sub.at] : null
  return (
    <div ref={ref} className="de-ctx" style={pos} role="menu" onMouseDown={(e) => e.preventDefault()} onContextMenu={(e) => e.preventDefault()}>
      <Rows
        items={items}
        hi={hi}
        open={sub?.at}
        rowRef={(i, el) => (rows.current[i] = el)}
        onHover={(i) => {
          setHi(i)
          const it = items[i]
          if (it.items && !it.disabled) openSub(i)
          else if (sub) {
            stay()
            leave.current = setTimeout(() => setSub(null), 250)
          }
        }}
        onPick={(i) => pick(items, i)}
      />
      {sub && subItems && anchor && (
        <Sub anchor={anchor} items={subItems} hi={sub.hi} onEnter={stay} onHover={(i) => setSub({ at: sub.at, hi: i })} onPick={(i) => pick(subItems, i, sub.at)} />
      )}
    </div>
  )
}

/** platform modifier label for shortcut hints */
export const MOD = typeof navigator !== 'undefined' && /Mac|iPhone|iPad/.test(navigator.platform) ? '⌘' : 'Ctrl+'
