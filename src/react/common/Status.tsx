import { createContext, useContext, useEffect } from 'react'
import { createPortal } from 'react-dom'

/** Toolbar slot provided by DocViewer; views portal their status text into it (no parent re-render). */
export const StatusContext = createContext<HTMLElement | null>(null)

export function Status({ text, onStatus }: { text: string; onStatus?: (s: string) => void }) {
  const el = useContext(StatusContext)
  useEffect(() => {
    onStatus?.(text)
  }, [text, onStatus])
  return el ? createPortal(text, el) : null
}
