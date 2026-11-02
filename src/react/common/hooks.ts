import { useCallback, useEffect, useLayoutEffect, useReducer, useRef, useState } from 'react'
import type { DocumentClient } from '../../core/worker/client'
import { isWebImage } from '../../core/opc'

export const useIsoLayoutEffect = typeof window === 'undefined' ? useEffect : useLayoutEffect

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

/** Scroll position of an element, batched to animation frames. */
export function useScroll(ref: React.RefObject<HTMLElement | null>): { x: number; y: number } {
  const [pos, setPos] = useState({ x: 0, y: 0 })
  useEffect(() => {
    const el = ref.current
    if (!el) return
    let raf = 0
    const on = () => {
      if (raf) return
      raf = requestAnimationFrame(() => {
        raf = 0
        setPos({ x: el.scrollLeft, y: el.scrollTop })
      })
    }
    el.addEventListener('scroll', on, { passive: true })
    return () => {
      el.removeEventListener('scroll', on)
      cancelAnimationFrame(raf)
    }
  }, [ref])
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
