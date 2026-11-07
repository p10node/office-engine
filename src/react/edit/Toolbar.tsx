import {
  Children,
  Fragment,
  isValidElement,
  useEffect,
  useLayoutEffect,
  useRef,
  useState,
  useSyncExternalStore,
  type CSSProperties,
  type ReactElement,
  type ReactNode,
  type RefObject,
} from 'react'

const BORDER_GRID = <path d="M4 4h16v16H4zM12 4v16M4 12h16" strokeDasharray="1.2 2.6" opacity={0.5} strokeWidth={1.4} />

const P: Record<string, ReactNode> = {
  undo: <path d="M9 14 4 9l5-5M4 9h10.5a5.5 5.5 0 0 1 0 11H11" />,
  redo: <path d="m15 14 5-5-5-5M20 9H9.5a5.5 5.5 0 0 0 0 11H13" />,
  bold: <path d="M7 5h6a3.5 3.5 0 0 1 0 7H7zM7 12h7a3.5 3.5 0 0 1 0 7H7z" strokeWidth={2.4} />,
  italic: <path d="M14 5h-4m4 14h-4m3-14-2 14" strokeWidth={2.2} />,
  underline: <path d="M7 4v7a5 5 0 0 0 10 0V4M5 20h14" />,
  strike: <path d="M16 6.5C15.3 5 13.8 4 12 4c-2.5 0-4 1.3-4 3 0 4 8 2.5 8 7 0 2-1.8 3.5-4.2 3.5-2 0-3.6-1-4.3-2.5M4 12h16" />,
  color: <path d="m6 16 6-12 6 12M8.3 11.5h7.4" />,
  highlight: <path d="m9 11-5 5v3h3l5-5M14.5 5.5l4 4L12 16l-4-4zM15 4l5 5" />,
  fill: <path d="m19 11-8-8-8.5 8.5a2 2 0 0 0 0 2.8l5.7 5.7a2 2 0 0 0 2.8 0zM5 2l5 5M21 15s2 2.5 2 4a2 2 0 0 1-4 0c0-1.5 2-4 2-4" />,
  alignLeft: <path d="M4 6h16M4 10h10M4 14h16M4 18h10" />,
  alignCenter: <path d="M4 6h16M7 10h10M4 14h16M7 18h10" />,
  alignRight: <path d="M4 6h16M10 10h10M4 14h16M10 18h10" />,
  justify: <path d="M4 6h16M4 10h16M4 14h16M4 18h16" />,
  bullets: <path d="M9 6h11M9 12h11M9 18h11M4.5 6h.01M4.5 12h.01M4.5 18h.01" strokeWidth={2.2} />,
  numbers: <path d="M10 6h10M10 12h10M10 18h10M4 4.5h1.5v4M4 8.5h3M4 14.5c0-1 2.5-1.2 2.5.3 0 1.2-2.5 1.8-2.5 3.2h2.8" />,
  indent: <path d="M4 5h16M11 10h9M11 14h9M4 19h16M4 9l3 3-3 3" />,
  outdent: <path d="M4 5h16M11 10h9M11 14h9M4 19h16M7 9l-3 3 3 3" />,
  clear: <path d="M5 5h10M10 5l-3 14M14 14l5 5m0-5-5 5" />,
  search: <path d="m20 20-4.2-4.2M17 11a6 6 0 1 1-12 0 6 6 0 0 1 12 0" />,
  save: <path d="M12 4v11m-4.5-4.5L12 15l4.5-4.5M5 19h14" />,
  wrap: <path d="M4 6h16M4 12h13a3 3 0 0 1 0 6h-4m2-2-2 2 2 2M4 18h5" />,
  edit: <path d="m4 20 1-4L16 5l3 3L8 19zM14 7l3 3" />,
  view: <path d="M2 12s3.5-7 10-7 10 7 10 7-3.5 7-10 7S2 12 2 12m10 3a3 3 0 1 0 0-6 3 3 0 0 0 0 6" />,
  close: <path d="M6 6l12 12M18 6 6 18" />,
  image: <path d="M4 5h16v14H4zM4 16l5-5 4 4 3-3 4 4M15.5 9.5h.01" />,
  table: <path d="M4 5h16v14H4zM4 10h16M4 15h16M10 5v14M15 5v14" />,
  link: <path d="M10 14a4 4 0 0 0 5.7 0l3-3a4 4 0 0 0-5.7-5.7l-1 1M14 10a4 4 0 0 0-5.7 0l-3 3a4 4 0 0 0 5.7 5.7l1-1" />,
  pagebreak: <path d="M5 4v5h14V4M5 20v-5h14v5M3 12h2m3 0h2m3 0h2m3 0h3" />,
  sup: <path d="m4 8 8 10m0-10L4 18M16 4h4l-4 5h4" />,
  sub: <path d="m4 6 8 10m0-10L4 16M16 15h4l-4 5h4" />,
  lineSpacing: <path d="M10 6h10M10 12h10M10 18h10M5 4v16M3 6l2-2 2 2M3 18l2 2 2-2" />,
  plus: <path d="M12 5v14M5 12h14" />,
  minus: <path d="M5 12h14" />,
  fit: <path d="M4 9V4h5M20 9V4h-5M4 15v5h5M20 15v5h-5" />,
  shape: <path d="M4 4h9v9H4zM17 11a4 4 0 1 1 0 8 4 4 0 0 1 0-8" />,
  textbox: <path d="M4 5h16v14H4zM8 9h8M12 9v7" />,
  slideAdd: <path d="M3 6h14v12H3zM20 8v6M17 11h6" />,
  copy: <path d="M8 8h12v12H8zM4 16V4h12" />,
  trash: <path d="M4 7h16M9 7V4h6v3M6 7l1 13h10l1-13" />,
  play: <path d="M7 4l13 8-13 8z" />,
  front: <path d="M8 8h12v12H8zM4 4h10v3M4 4v10h3" />,
  back: <path d="M4 4h12v12H4zM20 8v12H8v-3" />,
  merge: <path d="M4 5h16v14H4zM9 12H4m16 0h-5m-6-3 3 3-3 3m6-6-3 3 3 3" />,
  borders: <path d="M4 4h16v16H4zM12 4v16M4 12h16" strokeDasharray="2 2" />,
  freeze: <path d="M4 4h16v16H4zM4 9h16M9 4v16" />,
  sum: <path d="M17 5H7l6 7-6 7h10" />,
  rowAdd: <path d="M4 4h16v6H4zM12 14v6M9 17h6" />,
  colAdd: <path d="M4 4h6v16H4zM14 12h6M17 9v6" />,
  sort: <path d="M7 4v16M4 17l3 3 3-3M14 6h6M14 12h4M14 18h2" />,
  up: <path d="m6 15 6-6 6 6" />,
  down: <path d="m6 9 6 6 6-6" />,
  chevron: <path d="m7 10 5 5 5-5" />,
  lineWeight: <><path d="M4 6h16" strokeWidth={1} /><path d="M4 11h16" strokeWidth={2.2} /><path d="M4 17.5h16" strokeWidth={3.6} /></>,
  rotate: <path d="M20 11a8 8 0 1 0-2.3 5.7M20 4v7h-7" />,
  rotateCcw: <path d="M4 11a8 8 0 1 1 2.3 5.7M4 4v7h7" />,
  flipH: <path d="M12 3v18M9 7 4 17h5zM15 7l5 10h-5z" />,
  flipV: <path d="M3 12h18M7 9 17 4v5zM7 15l10 5v-5z" />,
  alignObj: <path d="M4 3v18M8 7h12v4H8zM8 14h7v4H8z" />,
  currency: <path d="M12 3v18M16.5 7.5C16 5.9 14.3 5 12 5c-2.6 0-4.5 1.3-4.5 3.3 0 4.7 9 2.6 9 7.4 0 2-1.9 3.3-4.5 3.3-2.4 0-4.2-1-4.7-2.8" />,
  vTop: <path d="M5 4h14M12 20V8M8.5 11.5 12 8l3.5 3.5" />,
  vMiddle: <path d="M5 12h14M12 3v6M9.5 6.5 12 9l2.5-2.5M12 21v-6M9.5 17.5 12 15l2.5 2.5" />,
  vBottom: <path d="M5 20h14M12 4v12M8.5 12.5 12 16l3.5-3.5" />,
  rotNone: <path d="M5 17h14M7 14l3-8 3 8M8 11.5h4" />,
  rotUp: <path d="M5 19h14M6 15l9.5-9.5M9 6.5l4.5 7.5M8 10l4-.5" />,
  rotDown: <path d="M5 5h14M6 9l9.5 9.5M9 17.5l4.5-7.5M8 14l4 .5" />,
  rotStack: <path d="M9 4l2 5 2-5M9.6 6.5h2.8M9 11h4M11 11v4M9 15h4M9 17h3a1.5 1.5 0 0 1 0 3H9z" strokeWidth={1.5} />,
  rotUp90: <path d="M19 5v14M15 18l-8-3 8-3M12.5 13v4M7 7l2-2 2 2M9 5v6" />,
  rotDown90: <path d="M19 5v14M7 6l8 3-8 3M9.5 11V7M13 17l-2 2-2-2M11 19v-6" />,
  sortAZ: <path d="M7 4v16M4 17l3 3 3-3M13 11l2.5-7 2.5 7M13.7 9h3.6M13 14h5l-5 6h5" />,
  sortZA: <path d="M7 4v16M4 17l3 3 3-3M13 4h5l-5 6h5M13 20l2.5-7 2.5 7M13.7 18h3.6" />,
  gridlines: <path d="M4 4h16v16H4zM4 9.3h16M4 14.6h16M9.3 4v16M14.6 4v16" />,
  lineStyle: <path d="M4 6h16M4 12h16M4 18h3M10 18h4M17 18h3" strokeWidth={2} />,
  paint: <path d="M5 4h11v5H5zM16 6.5h3v5l-7 1.5v3M11 16h2v5h-2z" />,
  styleMatch: <path d="m3 18 4.5-12L12 18M4.7 14h5.6M21 9.5a4 4 0 0 0-7-1.8M14 5v3h3M14 14.5a4 4 0 0 0 7 1.8M21 19v-3h-3" />,
  omega: <path d="M5 19h4v-2.5A6.5 6.5 0 1 1 15 16.5V19h4" />,
  outline: <path d="M4 6h.01M4 12h.01M4 18h.01M8 6h12M11 12h9M11 18h9" strokeWidth={2} />,
  imgOrig: <path d="M4 5h16v14H4zM8 15V9l-1.5 1.5M12 10h.01M12 14h.01M16.5 15V9L15 10.5" />,
  imgReplace: <path d="M4 5h10v8H4zM4 11l3-3 3 3 2-2 2 2M14 19h6v-6M20 19l-5-5M7 16v3h4" />,
  alt: <path d="M4 5h16v14H4zM7.5 15l2.5-6 2.5 6M8.3 13h3.4M15 9v6h2.5" />,
  cut: <path d="M6 6a2.5 2.5 0 1 0 0 .01M6 18a2.5 2.5 0 1 0 0 .01M8 7.5 19 18M8 16.5 19 6" />,
  paste: <path d="M9 4h6v3H9zM15 5h3v15H6V5h3" />,
  pastePlain: <path d="M9 4h6v3H9zM15 5h3v15H6V5h3M9 11h6M12 11v6" />,
  external: <path d="M14 4h6v6M20 4l-9 9M18 14v6H4V6h6" />,
  unlink: <path d="M10 14a4 4 0 0 0 5.7 0l3-3a4 4 0 0 0-5.7-5.7l-1 1M14 10a4 4 0 0 0-5.7 0l-3 3a4 4 0 0 0 5.7 5.7l1-1M3 3l18 18" />,
  selectAll: <path d="M4 4h3M10 4h4M17 4h3v3M20 10v4M20 17v3h-3M14 20h-4M7 20H4v-3M4 14v-4M4 7V4M8 8h8v8H8z" />,
  hideShow: <path d="M3 3l18 18M10.6 6.1A10 10 0 0 1 12 6c6.5 0 10 6 10 6a17 17 0 0 1-3.1 3.8M6.6 7.6C3.8 9.3 2 12 2 12s3.5 6 10 6c1.6 0 3-.4 4.3-1" />,
  percent: <path d="M19 5 5 19M7.5 9.5a2 2 0 1 0 0-4 2 2 0 0 0 0 4M16.5 18.5a2 2 0 1 0 0-4 2 2 0 0 0 0 4" />,
  decLess: <path d="M3.5 16h.01M8 9.5a2.5 3.5 0 1 1 0 7 2.5 3.5 0 0 1 0-7M21 19h-7m0 0 2-2m-2 2 2 2" strokeWidth={1.7} />,
  decMore: <path d="M3.5 16h.01M8 9.5a2.5 3.5 0 1 1 0 7 2.5 3.5 0 0 1 0-7M14.5 9.5a2 3 0 1 1 0 6 2 3 0 0 1 0-6M13 19h8m0 0-2-2m2 2-2 2" strokeWidth={1.7} />,
  check: <path d="m5 12.5 4.5 4.5L19 7.5" strokeWidth={2.2} />,
  more: <path d="M12 5.5h.01M12 12h.01M12 18.5h.01" strokeWidth={3.2} />,
  bAll: <path d="M4 4h16v16H4zM12 4v16M4 12h16" />,
  bOuter: <>{BORDER_GRID}<path d="M4 4h16v16H4z" /></>,
  bInner: <>{BORDER_GRID}<path d="M12 4v16M4 12h16" /></>,
  bTop: <>{BORDER_GRID}<path d="M4 4h16" strokeWidth={2.2} /></>,
  bBottom: <>{BORDER_GRID}<path d="M4 20h16" strokeWidth={2.2} /></>,
  bLeft: <>{BORDER_GRID}<path d="M4 4v16" strokeWidth={2.2} /></>,
  bRight: <>{BORDER_GRID}<path d="M20 4v16" strokeWidth={2.2} /></>,
  bNone: BORDER_GRID,
  rect: <path d="M4 6h16v12H4z" />,
  roundRect: <path d="M8 6h8a4 4 0 0 1 4 4v4a4 4 0 0 1-4 4H8a4 4 0 0 1-4-4v-4a4 4 0 0 1 4-4" />,
  ellipse: <path d="M12 6c4.4 0 8 2.7 8 6s-3.6 6-8 6-8-2.7-8-6 3.6-6 8-6" />,
  triangle: <path d="M12 5l8 14H4z" />,
  arrow: <path d="M4 10h9V6l7 6-7 6v-4H4z" />,
  star: <path d="m12 4 2.4 5 5.6.7-4.1 3.8 1.1 5.5-5-2.8-5 2.8 1.1-5.5L4 9.7 9.6 9z" />,
  line: <path d="M5 19 19 5" />,
  rowAbove: <path d="M4 12h16v8H4zM12 3v6M9 6h6" />,
  rowBelow: <path d="M4 4h16v8H4zM12 15v6M9 18h6" />,
  colLeft: <path d="M12 4h8v16h-8zM3 12h6M6 9v6" />,
  colRight: <path d="M4 4h8v16H4zM15 12h6M18 9v6" />,
  delRow: <path d="M4 8h16v8H4zM9 10.5l6 3m0-3-6 3" />,
  delCol: <path d="M8 4h8v16H8zM10.5 9l3 6m0-6-3 6" />,
  hide: <path d="M3 3l18 18M10.6 6.1A10 10 0 0 1 12 6c6.5 0 10 6 10 6a17 17 0 0 1-3.1 3.8M6.6 7.6C3.8 9.3 2 12 2 12s3.5 6 10 6c1.6 0 3-.4 4.3-1M9.9 9.9a3 3 0 0 0 4.2 4.2" />,
  file: <path d="M14 3H7a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8zM14 3v5h5" />,
  open: <path d="M4 19V6a1 1 0 0 1 1-1h4l2 2h7a1 1 0 0 1 1 1v2M4 19l2.6-7.2a1 1 0 0 1 .9-.8H20a1 1 0 0 1 1 1.3L18.6 19z" />,
  print: <path d="M7 9V4h10v5M7 17H5a1 1 0 0 1-1-1v-6a1 1 0 0 1 1-1h14a1 1 0 0 1 1 1v6a1 1 0 0 1-1 1h-2M7 14h10v6H7z" />,
  pdf: <path d="M14 3H7a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8zM14 3v5h5M8.5 16.5v-4h1.2a1.2 1.2 0 0 1 0 2.4H8.5M13 12.5v4h.8a1.6 1.6 0 0 0 1.6-1.6v-.8a1.6 1.6 0 0 0-1.6-1.6z" strokeWidth={1.5} />,
  html: <path d="M14 3H7a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8zM14 3v5h5M10 12l-2 2.5 2 2.5M14 12l2 2.5-2 2.5" />,
  doc: <path d="M14 3H7a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8zM14 3v5h5M9 13h6M9 17h4" />,
  slides: <path d="M4 5h16v11H4zM12 16v4M8 20h8" />,
  info: <path d="M12 21a9 9 0 1 0 0-18 9 9 0 0 0 0 18zM12 11v5M12 8h.01" />,
}

export function Icon({ name, size = 18 }: { name: string; size?: number }) {
  return (
    <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth={1.8} strokeLinecap="round" strokeLinejoin="round" aria-hidden>
      {P[name]}
    </svg>
  )
}

const keep = (e: React.MouseEvent) => e.preventDefault()

/** popover panel kept inside the viewport horizontally */
function PopBox(props: { className?: string; children: ReactNode; onMouseDown?: (e: React.MouseEvent) => void }) {
  const ref = useRef<HTMLDivElement>(null)
  const [dx, setDx] = useState(0)
  useLayoutEffect(() => {
    const el = ref.current
    if (!el) return
    const r = el.getBoundingClientRect()
    const vw = document.documentElement.clientWidth
    if (r.right > vw - 8) setDx(-Math.min(r.left - 8, r.right - vw + 8))
    else if (r.left < 8) setDx(8 - r.left)
  }, [])
  return (
    <div ref={ref} className={props.className ?? 'de-tb-pop'} style={dx ? { transform: `translateX(${dx}px)` } : undefined} onMouseDown={props.onMouseDown}>
      {props.children}
    </div>
  )
}

/** closes a popup on outside mousedown (shadow-DOM safe: uses the composed path) */
export function useDismiss(ref: RefObject<HTMLElement | null>, open: boolean, close: () => void) {
  const cb = useRef(close)
  cb.current = close
  useEffect(() => {
    if (!open) return
    const on = (e: MouseEvent) => {
      const el = ref.current
      if (el && !e.composedPath().includes(el)) cb.current()
    }
    window.addEventListener('mousedown', on, true)
    return () => window.removeEventListener('mousedown', on, true)
  }, [open, ref])
}

export function Btn(props: { icon?: string; label?: ReactNode; title: string; active?: boolean; disabled?: boolean; onClick: () => void; wide?: boolean }) {
  return (
    <button
      type="button"
      className={props.active ? 'de-tb de-tb-on' : 'de-tb'}
      title={props.title}
      aria-label={props.title}
      aria-pressed={props.active}
      disabled={props.disabled}
      onMouseDown={keep}
      onClick={props.onClick}
      style={props.wide ? { width: 'auto', padding: '0 8px' } : undefined}
    >
      {props.icon ? <Icon name={props.icon} /> : props.label}
    </button>
  )
}

export const Sep = () => <span className="de-tb-sep" />

export interface DropOption<T> {
  value: T
  label: string
  /** secondary text on the right (example, shortcut) */
  hint?: string
  icon?: string
  /** label style (font / style previews) */
  style?: CSSProperties
  disabled?: boolean
  /** separator above this option */
  sep?: boolean
  /** group heading above this option */
  head?: string
}

/**
 * Styled dropdown (listbox). Never takes focus from the editor: mouse down is prevented and,
 * while open, keys are captured at the window (arrows, Enter, Escape, type-ahead).
 */
export function Dropdown<T extends string | number>(props: {
  title: string
  options: DropOption<T>[]
  onChange: (v: T) => void
  /** selected value (check mark); omit for command menus */
  value?: T | ''
  placeholder?: string
  /** fixed trigger content (menus); null = icon only */
  label?: ReactNode
  icon?: string
  width?: number
  menuWidth?: number
  disabled?: boolean
  /** popup aligned to the trigger's right edge */
  alignRight?: boolean
  /** extra class of the trigger button */
  className?: string
}) {
  const { options } = props
  const [open, setOpen] = useState(false)
  const [hi, setHi] = useState(-1)
  const [shift, setShift] = useState(0)
  const ref = useRef<HTMLSpanElement>(null)
  const listRef = useRef<HTMLDivElement>(null)
  const typed = useRef({ text: '', at: 0 })
  const check = props.value !== undefined
  const sel = check ? options.findIndex((o) => o.value === props.value) : -1
  const current = sel >= 0 ? options[sel] : undefined
  useDismiss(ref, open, () => setOpen(false))

  const pick = (o: DropOption<T> | undefined) => {
    if (!o || o.disabled) return
    setOpen(false)
    props.onChange(o.value)
  }
  const show = () => {
    setHi(sel >= 0 ? sel : options.findIndex((o) => !o.disabled))
    setShift(0)
    setOpen(true)
  }

  // keyboard while open (captured so the document editor does not see the keys)
  useEffect(() => {
    if (!open) return
    const on = (e: KeyboardEvent) => {
      const n = options.length
      const step = (from: number, d: number) => {
        let i = from
        for (let k = 0; k < n; k++) {
          i = (i + d + n) % n
          if (!options[i].disabled) return i
        }
        return from
      }
      if (e.key === 'ArrowDown') setHi((h) => step(h, 1))
      else if (e.key === 'ArrowUp') setHi((h) => step(h < 0 ? 0 : h, -1))
      else if (e.key === 'Home') setHi(step(-1, 1))
      else if (e.key === 'End') setHi(step(0, -1))
      else if (e.key === 'Enter' || (e.key === ' ' && !typed.current.text)) pick(options[hi])
      else if (e.key === 'Escape') setOpen(false)
      else if (e.key === 'Tab') return setOpen(false)
      else if (e.key.length === 1 && !e.metaKey && !e.ctrlKey && !e.altKey) {
        // type-ahead on labels
        const t = typed.current
        t.text = (Date.now() - t.at < 700 ? t.text : '') + e.key.toLowerCase()
        t.at = Date.now()
        const i = options.findIndex((o) => !o.disabled && o.label.toLowerCase().startsWith(t.text))
        if (i >= 0) setHi(i)
      } else return
      e.preventDefault()
      e.stopPropagation()
    }
    window.addEventListener('keydown', on, true)
    return () => window.removeEventListener('keydown', on, true)
  }, [open, hi, options]) // eslint-disable-line react-hooks/exhaustive-deps

  // keep the popup inside the viewport, highlighted option visible
  useLayoutEffect(() => {
    const el = listRef.current
    if (!open || !el) return
    const r = el.getBoundingClientRect()
    const vw = document.documentElement.clientWidth
    if (r.right > vw - 8) setShift((s) => s - (r.right - vw + 8))
    else if (r.left < 8) setShift((s) => s + (8 - r.left))
    // the viewer clips its content: keep the list above its bottom edge (scrolls when longer)
    const root = el.closest('.de-root')
    const bottom = Math.min(document.documentElement.clientHeight, root ? root.getBoundingClientRect().bottom : Infinity)
    const cssMax = parseFloat(getComputedStyle(el).maxHeight)
    const room = bottom - r.top - 8
    if (room > 0) el.style.maxHeight = `${Math.max(120, Math.min(Number.isFinite(cssMax) ? cssMax : room, room))}px`
  }, [open])
  useLayoutEffect(() => {
    if (open) listRef.current?.querySelector(`[data-i="${hi}"]`)?.scrollIntoView({ block: 'nearest' })
  }, [open, hi])

  const label = props.label !== undefined ? props.label : (current?.label ?? props.placeholder ?? '')
  return (
    <span ref={ref} className="de-dd-wrap">
      <button
        type="button"
        className={(open ? 'de-dd de-dd-open' : 'de-dd') + (props.className ? ' ' + props.className : '')}
        style={{ width: props.width }}
        title={props.title}
        aria-label={props.title}
        aria-haspopup="listbox"
        aria-expanded={open}
        disabled={props.disabled}
        onMouseDown={keep}
        onClick={() => (open ? setOpen(false) : show())}
        onKeyDown={(e) => {
          if (!open && (e.key === 'ArrowDown' || e.key === 'Enter' || e.key === ' ')) {
            e.preventDefault()
            show()
          }
        }}
      >
        {props.icon && <Icon name={props.icon} />}
        {label !== null && <span className={current || props.label !== undefined ? 'de-dd-label' : 'de-dd-label de-dd-ph'}>{label}</span>}
        <Icon name="chevron" size={14} />
      </button>
      {open && (
        <div
          ref={listRef}
          className="de-dd-pop"
          role="listbox"
          aria-label={props.title}
          style={{ minWidth: props.menuWidth, ...(props.alignRight ? { right: -shift } : { left: shift }) }}
          onMouseDown={keep}
        >
          {options.map((o, i) => (
            <Fragment key={String(o.value)}>
              {o.sep && i > 0 && <div className="de-dd-sep" role="separator" />}
              {o.head && (
                <div className="de-dd-head" role="presentation">
                  {o.head}
                </div>
              )}
              <div
                data-i={i}
                role="option"
                aria-selected={i === sel}
                aria-disabled={o.disabled || undefined}
                className={'de-dd-opt' + (i === hi ? ' de-dd-hi' : '') + (o.disabled ? ' de-dd-dis' : '')}
                onMouseMove={() => i !== hi && !o.disabled && setHi(i)}
                onClick={() => pick(o)}
              >
                {check && <span className="de-dd-check">{i === sel && <Icon name="check" size={16} />}</span>}
                {o.icon && <Icon name={o.icon} />}
                <span className="de-dd-text" style={o.style}>
                  {o.label}
                </span>
                {o.hint && <span className="de-dd-hint">{o.hint}</span>}
              </div>
            </Fragment>
          ))}
        </div>
      )}
    </span>
  )
}

/** Value picker (fonts, styles, formats…). */
export function Select<T extends string | number>(props: {
  value: T | ''
  options: DropOption<T>[]
  onChange: (v: T) => void
  title: string
  width?: number
  placeholder?: string
  menuWidth?: number
  disabled?: boolean
}) {
  return <Dropdown {...props} />
}

/** Font size box: − [size] + with a list of common sizes (Google Docs style). */
export function SizeBox(props: { value?: number; placeholder?: number; sizes?: number[]; onChange: (n: number) => void; title?: string; disabled?: boolean }) {
  const sizes = props.sizes ?? SIZES
  const [text, setText] = useState<string | null>(null)
  const [open, setOpen] = useState(false)
  const ref = useRef<HTMLSpanElement>(null)
  const inputRef = useRef<HTMLInputElement>(null)
  useDismiss(ref, open, () => setOpen(false))
  const base = props.value ?? props.placeholder ?? 11
  const set = (n: number) => {
    setText(null)
    setOpen(false)
    if (Number.isFinite(n) && n >= 1 && n <= 400) props.onChange(Math.round(n * 2) / 2)
  }
  const done = () => {
    if (text !== null) set(parseFloat(text))
    inputRef.current?.blur()
  }
  const title = props.title ?? 'Font size'
  return (
    <span ref={ref} className="de-size">
      <button type="button" className="de-tb de-size-step" title="Decrease font size" aria-label="Decrease font size" disabled={props.disabled} onMouseDown={keep} onClick={() => set(Math.max(1, Math.ceil(base) - 1))}>
        <Icon name="minus" size={16} />
      </button>
      <input
        ref={inputRef}
        className="de-size-input"
        title={title}
        aria-label={title}
        inputMode="decimal"
        disabled={props.disabled}
        value={text ?? (props.value !== undefined ? String(props.value) : '')}
        placeholder={props.placeholder !== undefined ? String(props.placeholder) : ''}
        onFocus={(e) => {
          e.currentTarget.select()
          setOpen(true)
        }}
        onChange={(e) => setText(e.target.value.replace(/[^\d.]/g, ''))}
        onKeyDown={(e) => {
          if (e.key === 'Enter') {
            e.preventDefault()
            done()
          } else if (e.key === 'Escape') {
            e.preventDefault()
            setText(null)
            setOpen(false)
            inputRef.current?.blur()
          } else if (e.key === 'ArrowUp' || e.key === 'ArrowDown') {
            e.preventDefault()
            const cur = parseFloat(text ?? String(base))
            const next = e.key === 'ArrowUp' ? sizes.find((n) => n > cur) : [...sizes].reverse().find((n) => n < cur)
            if (next !== undefined) setText(String(next))
          }
        }}
        onBlur={() => {
          if (text !== null) set(parseFloat(text))
          setOpen(false)
        }}
      />
      <button type="button" className="de-tb de-size-step" title="Increase font size" aria-label="Increase font size" disabled={props.disabled} onMouseDown={keep} onClick={() => set(Math.floor(base) + 1)}>
        <Icon name="plus" size={16} />
      </button>
      {open && (
        <div className="de-dd-pop de-size-pop" role="listbox" aria-label={title} onMouseDown={keep}>
          {sizes.map((n) => (
            <div key={n} role="option" aria-selected={n === props.value} className={'de-dd-opt' + (String(n) === (text ?? String(props.value)) ? ' de-dd-hi' : '')} onClick={() => set(n)}>
              {n}
            </div>
          ))}
        </div>
      )}
    </span>
  )
}

// ---------------------------------------------------------------- responsive toolbar row

function flatten(children: ReactNode): ReactElement[] {
  const out: ReactElement[] = []
  Children.forEach(children, (c) => {
    if (!isValidElement(c)) return
    if (c.type === Fragment) out.push(...flatten((c.props as { children?: ReactNode }).children))
    else out.push(c)
  })
  return out
}

const signature = (el: ReactElement) => {
  const p = el.props as { title?: string; icon?: string; label?: unknown }
  const t = typeof el.type === 'string' ? el.type : (el.type as { name?: string }).name
  return `${t}:${p.title ?? p.icon ?? (typeof p.label === 'string' ? p.label : '')}`
}

/**
 * One-line toolbar: items that do not fit move into a "more" popover (like Google Docs).
 * Item widths are measured once per set of items and cached for resizes.
 */
export function ToolbarRow({ children }: { children: ReactNode }) {
  const items = flatten(children)
  const sig = items.map(signature).join('|')
  const rowRef = useRef<HTMLDivElement>(null)
  const widths = useRef<number[]>([])
  const [fit, setFit] = useState<{ sig: string; cut: number; w: number }>({ sig: '', cut: items.length, w: 0 })
  const [more, setMore] = useState(false)
  const moreRef = useRef<HTMLSpanElement>(null)
  useDismiss(moreRef, more, () => setMore(false))
  const measured = fit.sig === sig
  const cut = measured ? fit.cut : items.length

  const layout = (w: number) => {
    const ws = widths.current
    const gap = 2
    const total = ws.reduce((a, b) => a + b + gap, 0)
    let k = items.length
    if (total > w) {
      const room = w - 34
      let used = 0
      k = 0
      while (k < items.length && used + ws[k] + gap <= room) used += ws[k++] + gap
    }
    setFit((f) => (f.sig === sig && f.cut === k && f.w === w ? f : { sig, cut: k, w }))
  }
  const avail = () => {
    const el = rowRef.current
    if (!el) return 0
    const cs = getComputedStyle(el)
    return el.clientWidth - parseFloat(cs.paddingLeft) - parseFloat(cs.paddingRight)
  }

  // new set of items: everything is rendered once, measured, then cut
  useLayoutEffect(() => {
    const el = rowRef.current
    if (!el || measured) return
    widths.current = [...el.querySelectorAll<HTMLElement>(':scope > [data-i]')].map((x) => x.getBoundingClientRect().width)
    layout(avail())
  })
  const relayout = useRef(() => {})
  relayout.current = () => {
    if (widths.current.length === items.length) layout(avail())
  }
  useEffect(() => {
    const el = rowRef.current
    if (!el) return
    const ro = new ResizeObserver(() => relayout.current())
    ro.observe(el)
    return () => ro.disconnect()
  }, [])

  const isSep = (el: ReactElement | undefined) => !!el && el.type === Sep
  let end = cut
  while (end > 0 && isSep(items[end - 1]) && end < items.length) end--
  const rest = items.slice(cut).filter((el, i, all) => !(isSep(el) && (i === 0 || i === all.length - 1)))
  return (
    <div ref={rowRef} className="de-ebar" role="toolbar">
      {items.slice(0, end).map((el, i) => (
        <span key={i} data-i={i} className="de-ebar-it">
          {el}
        </span>
      ))}
      {rest.length > 0 && (
        <span ref={moreRef} className="de-tb-color de-ebar-morewrap">
          <button type="button" className={more ? 'de-tb de-tb-on' : 'de-tb'} title="More" aria-label="More tools" aria-expanded={more} onMouseDown={keep} onClick={() => setMore((m) => !m)}>
            <Icon name="more" />
          </button>
          {more && (
            <PopBox className="de-tb-pop de-ebar-more">
              {rest.map((el, i) => (
                <span key={i} className="de-ebar-it">
                  {el}
                </span>
              ))}
            </PopBox>
          )}
        </span>
      )}
    </div>
  )
}

const PALETTE = [
  '#000000', '#434343', '#666666', '#999999', '#b7b7b7', '#cccccc', '#d9d9d9', '#ffffff',
  '#980000', '#ff0000', '#ff9900', '#ffff00', '#00ff00', '#00ffff', '#4a86e8', '#0000ff',
  '#9900ff', '#ff00ff', '#e6b8af', '#f4cccc', '#fce5cd', '#fff2cc', '#d9ead3', '#d0e0e3',
  '#c9daf8', '#cfe2f3', '#d9d2e9', '#ead1dc', '#cc4125', '#e06666', '#f6b26b', '#ffd966',
  '#93c47d', '#76a5af', '#6d9eeb', '#6fa8dc', '#8e7cc3', '#c27ba0', '#1155cc', '#38761d',
]

// custom colors picked by the user: shared by every color button, kept across sessions
const CUSTOM_KEY = 'office-engine:custom-colors'
const MAX_CUSTOM = 7
const HEX6 = /^#[0-9a-f]{6}$/
let customColors: string[] | null = null
const customListeners = new Set<() => void>()

function getCustomColors(): string[] {
  if (customColors) return customColors
  customColors = []
  try {
    const v: unknown = JSON.parse(localStorage.getItem(CUSTOM_KEY) ?? '[]')
    if (Array.isArray(v)) customColors = v.filter((c): c is string => typeof c === 'string' && HEX6.test(c)).slice(0, MAX_CUSTOM)
  } catch {
    // storage unavailable (private mode, sandbox)
  }
  return customColors
}

function addCustomColor(c: string) {
  customColors = [c, ...getCustomColors().filter((x) => x !== c)].slice(0, MAX_CUSTOM)
  try {
    localStorage.setItem(CUSTOM_KEY, JSON.stringify(customColors))
  } catch {
    // not persisted
  }
  customListeners.forEach((f) => f())
}

const subscribeCustom = (f: () => void) => {
  customListeners.add(f)
  return () => {
    customListeners.delete(f)
  }
}

export function ColorBtn(props: { icon: string; title: string; value?: string; onPick: (c: string | null) => void; palette?: string[]; noneLabel?: string }) {
  const [open, setOpen] = useState(false)
  const ref = useRef<HTMLSpanElement>(null)
  const pickerRef = useRef<HTMLInputElement>(null)
  const custom = useSyncExternalStore(subscribeCustom, getCustomColors, getCustomColors)
  useDismiss(ref, open, () => setOpen(false))
  const onPick = useRef(props.onPick)
  onPick.current = props.onPick
  const cur = props.value?.toLowerCase()

  // native 'change' fires once the OS picker is confirmed (React's onChange would fire while dragging)
  useEffect(() => {
    const el = pickerRef.current
    if (!el) return
    const on = () => {
      const c = el.value.toLowerCase()
      if (!HEX6.test(c)) return
      addCustomColor(c)
      onPick.current(c)
    }
    el.addEventListener('change', on)
    return () => el.removeEventListener('change', on)
  }, [])

  const openPicker = () => {
    const el = pickerRef.current
    if (!el) return
    el.value = cur && HEX6.test(cur) ? cur : (custom[0] ?? '#000000')
    setOpen(false)
    try {
      el.showPicker()
    } catch {
      el.click()
    }
  }
  const swatch = (c: string) => (
    <button key={c} type="button" title={c} className={cur === c ? 'on' : undefined} style={{ background: c }} onClick={() => (props.onPick(c), setOpen(false))} />
  )
  return (
    <span ref={ref} className="de-tb-color">
      <button type="button" className={open ? 'de-tb de-tb-on' : 'de-tb'} title={props.title} aria-label={props.title} aria-expanded={open} onMouseDown={keep} onClick={() => setOpen((o) => !o)}>
        <Icon name={props.icon} />
        <span className="de-tb-swatch" style={{ background: props.value ?? 'transparent' }} />
      </button>
      {/* kept mounted while the OS picker is open (the popover closes) */}
      <input ref={pickerRef} type="color" className="de-tb-native" tabIndex={-1} aria-hidden />
      {open && (
        <PopBox onMouseDown={keep}>
          <button type="button" className="de-tb-none" onClick={() => (props.onPick(null), setOpen(false))}>
            {props.noneLabel ?? 'None'}
          </button>
          <div className="de-tb-grid">{(props.palette ?? PALETTE).map(swatch)}</div>
          <div className="de-tb-custom">Custom</div>
          <div className="de-tb-grid">
            {custom.map(swatch)}
            <button type="button" className="de-tb-add" title="Custom color…" aria-label="Custom color" onClick={openPicker}>
              <Icon name="plus" size={14} />
            </button>
          </div>
        </PopBox>
      )}
    </span>
  )
}

export const FONTS = ['Arial', 'Calibri', 'Cambria', 'Courier New', 'Georgia', 'Helvetica', 'Segoe UI', 'Tahoma', 'Times New Roman', 'Verdana']
export const SIZES = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 40, 48, 60, 72]

/** Font picker options (each shown in its own face); keeps an unknown current font at the top. */
export function fontOptions(current?: string): DropOption<string>[] {
  const list = current && !FONTS.includes(current) ? [current, ...FONTS] : FONTS
  return list.map((f, i) => ({ value: f, label: f, style: { fontFamily: `"${f}", system-ui, sans-serif` }, sep: i === 1 && list !== FONTS }))
}

/** Picks the first name of a CSS font stack. */
export function fontName(css: string | undefined): string {
  if (!css) return ''
  const m = /^\s*"?([^",]+)"?/.exec(css)
  return m ? m[1] : ''
}

export function download(blob: Blob, name: string) {
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url
  a.download = name
  document.body.appendChild(a)
  a.click()
  a.remove()
  setTimeout(() => URL.revokeObjectURL(url), 30_000)
}

/** Command menu: fixed label, items with optional icons. */
export function Menu(props: { title: string; label: string; icon?: string; items: { value: string; label: string; icon?: string; disabled?: boolean; sep?: boolean; hint?: string }[]; onPick: (v: string) => void; width?: number }) {
  return <Dropdown title={props.title} label={props.label.replace(/\s*▾$/, '')} icon={props.icon} options={props.items} onChange={props.onPick} />
}

/** Small popover anchored under a toolbar button. */
export function Pop(props: { icon: string; title: string; children: (close: () => void) => ReactNode; active?: boolean; disabled?: boolean; chevron?: boolean }) {
  const [open, setOpen] = useState(false)
  const ref = useRef<HTMLSpanElement>(null)
  useDismiss(ref, open, () => setOpen(false))
  return (
    <span ref={ref} className="de-tb-color">
      <button
        type="button"
        className={(props.active || open ? 'de-tb de-tb-on' : 'de-tb') + (props.chevron ? ' de-tb-chev' : '')}
        title={props.title}
        aria-label={props.title}
        aria-expanded={open}
        disabled={props.disabled}
        onMouseDown={keep}
        onClick={() => setOpen((o) => !o)}
      >
        <Icon name={props.icon} />
        {props.chevron && <Icon name="chevron" size={14} />}
      </button>
      {open && (
        <PopBox onMouseDown={(e) => (e.target as HTMLElement).closest('input, textarea') || e.preventDefault()}>
          {props.children(() => setOpen(false))}
        </PopBox>
      )}
    </span>
  )
}

/** rows × cols picker for new tables */
export function GridPicker({ onPick }: { onPick: (r: number, c: number) => void }) {
  const [hover, setHover] = useState<[number, number]>([0, 0])
  return (
    <div>
      <div className="de-gridpick" onMouseLeave={() => setHover([0, 0])}>
        {Array.from({ length: 64 }, (_, i) => {
          const r = Math.floor(i / 8) + 1
          const c = (i % 8) + 1
          return <button key={i} type="button" className={r <= hover[0] && c <= hover[1] ? 'on' : ''} onMouseEnter={() => setHover([r, c])} onMouseDown={keep} onClick={() => onPick(r, c)} />
        })}
      </div>
      <div className="de-gridpick-label">{hover[0] ? `${hover[0]} × ${hover[1]}` : 'Table size'}</div>
    </div>
  )
}

/** Inline text prompt (URL, sheet name…) inside a popover. */
export function PromptBox(props: { placeholder: string; initial?: string; ok: string; onOk: (v: string) => void; extra?: ReactNode }) {
  const [v, setV] = useState(props.initial ?? '')
  return (
    <form
      className="de-prompt"
      onSubmit={(e) => {
        e.preventDefault()
        props.onOk(v)
      }}
    >
      <input autoFocus value={v} placeholder={props.placeholder} onChange={(e) => setV(e.target.value)} />
      <button type="submit" className="de-btn">
        {props.ok}
      </button>
      {props.extra}
    </form>
  )
}

/** Opens the OS file picker (images by default). */
export function pickFiles(accept = 'image/*', multiple = true): Promise<File[]> {
  return new Promise((resolve) => {
    const input = document.createElement('input')
    input.type = 'file'
    input.accept = accept
    input.multiple = multiple
    input.onchange = () => resolve(Array.from(input.files ?? []))
    input.click()
  })
}
