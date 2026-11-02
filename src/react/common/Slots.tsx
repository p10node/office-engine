import { createContext, useContext, type ReactNode } from 'react'
import { createPortal } from 'react-dom'
import { ToolbarRow } from '../edit/Toolbar'
import type { DocumentClient } from '../../core/worker/client'

export interface ViewerApi {
  /** editor bar mount point (null when the viewer toolbar is hidden) */
  bar: HTMLElement | null
  editable: boolean
  /** report unsaved changes */
  setDirty: (dirty: boolean) => void
  /** build + hand off the edited file */
  save: () => Promise<void>
  client: DocumentClient | null
}

export const ViewerContext = createContext<ViewerApi>({
  bar: null,
  editable: false,
  setDirty: () => {},
  save: async () => {},
  client: null,
})

export const useViewer = () => useContext(ViewerContext)

/** Renders the editor bar into the viewer's toolbar row. */
export function EditorBar({ children }: { children: ReactNode }) {
  const { bar } = useViewer()
  return bar ? createPortal(<ToolbarRow>{children}</ToolbarRow>, bar) : null
}
