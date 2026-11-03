import { useCallback, useEffect, useLayoutEffect, useReducer, useRef, useState } from 'react'
import type { DocumentClient } from '../../core/worker/client'
import { isWebImage } from '../../core/opc'

export const useIsoLayoutEffect = typeof window === 'undefined' ? useEffect : useLayoutEffect

/** Color scheme of the viewer chrome: 'auto' follows the browser / OS setting (prefers-color-scheme). */
export type Theme = 'light' | 'dark' | 'auto'

const DARK_MQ = '(prefers-color-scheme: dark)'
const prefersDark = () => typeof matchMedia === 'function' && matchMedia(DARK_MQ).matches

/** Whether the dark theme applies for `theme`; 'auto' tracks prefers-color-scheme live. */
export function useDarkTheme(theme: Theme = 'light'): boolean {
  const [dark, setDark] = useState(() => theme === 'auto' && prefersDark())
  useEffect(() => {
    if (theme !== 'auto' || typeof matchMedia !== 'function') return
    const mq = matchMedia(DARK_MQ)
    const on = () => setDark(mq.matches)
    on()
    mq.addEventListener('change', on)
    return () => mq.removeEventListener('change', on)
  }, [theme])
  return theme === 'dark' || (theme === 'auto' && dark)
}

export function useForceUpdate(): [number, () => void] {
  const [n, bump] = useReducer((x: number) => x + 1, 0)
  return [n, bump as () => void]
}

/** Element size via ResizeObserver. */
export function useSize<T extends HTMLElement>(): [React.RefObject<T | null>, { w: number; h: number }] {
  const ref = useRef<T>(null)
  const [size, setSize] = useState({ w: 0, h: 0 })
  useIsoLayoutEffect(() => {
    const el = ref.current
    if (!el) return
    const update = () => {
      const w = el.clientWidth
      const h = el.clientHeight
      setSize((s) => (s.w === w && s.h === h ? s : { w, h }))
    }
    update()
    const ro = new ResizeObserver(update)
    ro.observe(el)
    return () => ro.disconnect()
  }, [])
  return [ref, size]
}

/** Scroll position of an element, batched to animation frames. Follows the element when it mounts after the first render. */
export function useScroll(ref: React.RefObject<HTMLElement | null>): { x: number; y: number } {
  const [pos, setPos] = useState({ x: 0, y: 0 })
  const [el, setEl] = useState<HTMLElement | null>(null)
  // a conditionally rendered panel (slide thumbnails) gets its element later than the first effect run
  useIsoLayoutEffect(() => {
    if (ref.current !== el) setEl(ref.current)
  })
  useEffect(() => {
    if (!el) return
    let raf = 0
    const read = () => setPos((p) => (p.x === el.scrollLeft && p.y === el.scrollTop ? p : { x: el.scrollLeft, y: el.scrollTop }))
    const on = () => {
      if (raf) return
      raf = requestAnimationFrame(() => {
        raf = 0
        read()
      })
    }
    read()
    el.addEventListener('scroll', on, { passive: true })
    return () => {
      el.removeEventListener('scroll', on)
      cancelAnimationFrame(raf)
    }
  }, [el])
  return pos
}

/** Lazily resolves a package part to an object URL (null while loading / unsupported). */
export function useMediaUrl(client: DocumentClient, path: string | undefined, enabled = true): string | null | 'unsupported' {
  const [url, setUrl] = useState<string | null | 'unsupported'>(null)
  useEffect(() => {
    if (!path || !enabled) return
    if (!isWebImage(path)) {
      setUrl('unsupported')
      return
    }
    let live = true
    client.mediaUrl(path).then(
      (u) => live && setUrl(u),
      () => live && setUrl('unsupported'),
    )
    return () => {
      live = false
    }
  }, [client, path, enabled])
  return url
}

/** Binary search: last index i with arr[i] <= v. */
export function lowerIndex(arr: ArrayLike<number>, v: number): number {
  let lo = 0
  let hi = arr.length - 1
  let ans = 0
  while (lo <= hi) {
    const mid = (lo + hi) >> 1
    if (arr[mid] <= v) {
      ans = mid
      lo = mid + 1
    } else hi = mid - 1
  }
  return ans
}

export function useLatest<T>(v: T) {
  const r = useRef(v)
  r.current = v
  return r
}

export function useStableCallback<A extends unknown[], R>(fn: (...a: A) => R) {
  const r = useLatest(fn)
  return useCallback((...a: A) => r.current(...a), [r])
}

/** A control that takes typed text: Ctrl/⌘+A selects its own content. */
export function isTextEntry(target: EventTarget | null): boolean {
  const el = target as HTMLElement | null
  if (!el || el.nodeType !== 1) return false
  const tag = el.tagName
  if (tag === 'TEXTAREA' || tag === 'SELECT') return true
  if (tag === 'INPUT') return !/^(button|checkbox|radio|submit|reset|range|color|file|image)$/i.test((el as HTMLInputElement).type)
  return !!el.isContentEditable
}

/**
 * Focuses `el` so keyboard shortcuts work without a click. The focus stays where it is when a control
 * outside the viewer has it (a host page form); the viewer's own chrome (or the page body) yields.
 */
export function claimFocus(el: HTMLElement | null): boolean {
  if (!el) return false
  const doc = el.ownerDocument
  const active = doc.activeElement
  const rn = el.getRootNode()
  const host = typeof ShadowRoot !== 'undefined' && rn instanceof ShadowRoot ? rn.host : null
  const within = el.closest('.de-root') ?? el
  if (active && active !== doc.body && active !== host && !within.contains(active)) return false
  el.focus({ preventScroll: true })
  return true
}
