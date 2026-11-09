import { createContext, useContext, useEffect, type ReactNode } from 'react'
import { createPortal } from 'react-dom'
import { ToolbarRow } from '../edit/Toolbar'
import type { DocumentClient } from '../../core/worker/client'
import type { Printable } from '../file/print'

/** What the active view offers to the File menu. */
export interface FileOps {
  /** all pages, for PDF / print and HTML */
  printable(): Promise<Printable>
  /** text formats: 'txt' / 'md' (documents, presentations), 'csv' / 'tsv' (current sheet; a Blob, as a sheet can be large) */
  text(format: string): Promise<string | Blob>
  /** File → Document details: [label, value] rows */
  stats(): [string, string][]
  /** File → Import (spreadsheets): the rows of a delimited text file (CSV / TSV) into a new sheet */
  importText?(file: Blob, name: string): Promise<void>
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
  /** set by the mounted view (see useSelectAll): Ctrl/⌘+A */
  selectAll: { current: (() => void) | null }
}

export const ViewerContext = createContext<ViewerApi>({
  bar: null,
  editable: false,
  setDirty: () => {},
  save: async () => {},
  canDownload: true,
  client: null,
  fileOps: { current: null },
  selectAll: { current: null },
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

/** Publishes the view's select-all: Ctrl/⌘+A from the viewer chrome or, with `captureSelectAll`, from anywhere on the page. */
export function useSelectAll(fn: () => void) {
  const slot = useViewer().selectAll
  useEffect(() => {
    slot.current = fn
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
