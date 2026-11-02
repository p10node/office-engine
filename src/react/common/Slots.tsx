import { createContext, useContext, useEffect, type ReactNode } from 'react'
import { createPortal } from 'react-dom'
import { ToolbarRow } from '../edit/Toolbar'
import type { DocumentClient } from '../../core/worker/client'
import type { Printable } from '../file/print'

/** What the active view offers to the File menu. */
export interface FileOps {
  /** all pages, for PDF / print and HTML */
  printable(): Promise<Printable>
  /** text formats: 'txt' / 'md' (documents, presentations), 'csv' / 'tsv' (current sheet) */
  text(format: string): Promise<string>
  /** File → Document details: [label, value] rows */
  stats(): [string, string][]
  /** File → Import (spreadsheets): rows into a new sheet */
  importTable?(rows: string[][], name: string): Promise<void>
}

export interface ViewerApi {
  /** editor bar mount point (null when the viewer toolbar is hidden) */
  bar: HTMLElement | null
  editable: boolean
  /** report unsaved changes */
  setDirty: (dirty: boolean) => void
  /** build + hand off the edited file */
  save: () => Promise<void>
  /** whether files may be downloaded to the user's device */
  canDownload: boolean
  client: DocumentClient | null
  /** set by the mounted view (see useFileOps) */
  fileOps: { current: FileOps | null }
}

export const ViewerContext = createContext<ViewerApi>({
  bar: null,
  editable: false,
  setDirty: () => {},
  save: async () => {},
  canDownload: true,
  client: null,
  fileOps: { current: null },
})

export const useViewer = () => useContext(ViewerContext)

/** Publishes the view's file operations (latest closure each render) to the File menu. */
export function useFileOps(ops: FileOps) {
  const slot = useViewer().fileOps
  useEffect(() => {
    slot.current = ops
  })
  useEffect(
    () => () => {
      slot.current = null
    },
    [slot],
  )
}

/** Renders the editor bar into the viewer's toolbar row. */
export function EditorBar({ children }: { children: ReactNode }) {
  const { bar } = useViewer()
  return bar ? createPortal(<ToolbarRow>{children}</ToolbarRow>, bar) : null
}
