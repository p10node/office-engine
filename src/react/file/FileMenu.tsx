import { useRef, useState } from 'react'
import type { DocKind } from '../../core/opc'
import { MOD } from '../common/ContextMenu'
import type { FileOps } from '../common/Slots'
import { download, Dropdown, pickFiles, type DropOption } from '../edit/Toolbar'
import { OPEN_ACCEPT } from './importFile'
import { exportHtml, printDocument } from './print'

export const KIND_LABEL: Record<DocKind, string> = {
  docx: 'Word document',
  xlsx: 'Excel workbook',
  pptx: 'PowerPoint presentation',
}
const EXT: Record<DocKind, string> = { docx: '.docx', xlsx: '.xlsx', pptx: '.pptx' }

const TEXT_FORMATS: Record<DocKind, { id: string; label: string; ext: string; mime: string }[]> = {
  docx: [
    { id: 'txt', label: 'Plain text', ext: '.txt', mime: 'text/plain' },
    { id: 'md', label: 'Markdown', ext: '.md', mime: 'text/markdown' },
  ],
  xlsx: [
    { id: 'csv', label: 'CSV (current sheet)', ext: '.csv', mime: 'text/csv' },
    { id: 'tsv', label: 'TSV (current sheet)', ext: '.tsv', mime: 'text/tab-separated-values' },
  ],
  pptx: [{ id: 'txt', label: 'Plain text', ext: '.txt', mime: 'text/plain' }],
}

/** Host-provided File menu entry (Share, Version history…). */
export interface FileMenuItem {
  label: string
  /** built-in icon name */
  icon?: string
  /** shortcut / note on the right */
  hint?: string
  disabled?: boolean
  onClick: () => void
}

export interface FileActionsOptions {
  kind: DocKind | undefined
  /** file name without extension */
  base: string
  /** New / Open (replace the document) */
  canOpen: boolean
  /** Import into the document (spreadsheets) */
  canImport: boolean
  canSave: boolean
  canDownload: boolean
  ops: { current: FileOps | null }
  create: (kind: DocKind) => void
  open: (file: File) => void
  save: () => Promise<void>
  /** the document in its own format */
  downloadOriginal: () => Promise<void>
  /** a picture file (SVG): downloads are the Word document, PDF and the picture itself */
  picture?: { label: string; ext: string; download: () => Promise<void> }
  rename?: () => void
  details?: () => void
  close?: () => void
  extra: FileMenuItem[]
  notify: (message: string) => void
}

export interface FileActions {
  options: DropOption<string>[]
  busy: boolean
  run: (action: string) => void
}

/** actions that open a picker / dialog right away (no busy state) */
const QUICK = /^(new:|open$|import$|rename$|details$|close$|extra:)/

/** File menu entries and their actions (also used by the Ctrl+O / Ctrl+P shortcuts). */
export function useFileActions(o: FileActionsOptions): FileActions {
  const [busy, setBusy] = useState(false)
  const busyRef = useRef(false)
  const opts = useRef(o)
  opts.current = o
  const kind = o.kind

  const options: DropOption<string>[] = []
  const section = (items: DropOption<string>[]) => {
    if (!items.length) return
    items[0] = { ...items[0], sep: options.length > 0 }
    options.push(...items)
  }
  section(
    o.canOpen
      ? [
          { value: 'new:docx', label: 'Document', icon: 'doc', head: 'New' },
          { value: 'new:xlsx', label: 'Spreadsheet', icon: 'table' },
          { value: 'new:pptx', label: 'Presentation', icon: 'slides' },
          { value: 'open', label: 'Open…', icon: 'open', hint: `${MOD}O`, sep: true },
        ]
      : [],
  )
  if (o.canImport && kind === 'xlsx') options.push({ value: 'import', label: 'Import CSV into new sheet…', icon: 'table', sep: !o.canOpen && options.length > 0 })
  const fileItems: DropOption<string>[] = []
  if (kind && o.canSave) fileItems.push({ value: 'save', label: 'Save', icon: 'save', hint: `${MOD}S` })
  if (kind && o.rename) fileItems.push({ value: 'rename', label: 'Rename…', icon: 'edit' })
  section(fileItems)
  if (o.canDownload && kind) {
    section([
      { value: 'original', label: KIND_LABEL[kind], icon: 'save', hint: EXT[kind], head: 'Download as' },
      { value: 'pdf', label: 'PDF document', icon: 'pdf', hint: '.pdf' },
      ...(o.picture
        ? [{ value: 'picture', label: o.picture.label, icon: 'image', hint: o.picture.ext }]
        : [{ value: 'html', label: 'Web page', icon: 'html', hint: '.html' }, ...TEXT_FORMATS[kind].map((f) => ({ value: 'text:' + f.id, label: f.label, icon: 'doc', hint: f.ext }))]),
    ])
  }
  const tail: DropOption<string>[] = []
  if (kind && o.details) tail.push({ value: 'details', label: 'Document details', icon: 'info' })
  if (kind && o.canDownload) tail.push({ value: 'print', label: 'Print', icon: 'print', hint: `${MOD}P` })
  section(tail)
  section(o.extra.map((x, i) => ({ value: `extra:${i}`, label: x.label, icon: x.icon, hint: x.hint, disabled: x.disabled })))
  if (o.close) section([{ value: 'close', label: 'Close', icon: 'close' }])

  const exec = async (action: string) => {
    const c = opts.current
    if (action.startsWith('new:')) return c.create(action.slice(4) as DocKind)
    if (action.startsWith('extra:')) return c.extra[Number(action.slice(6))]?.onClick()
    switch (action) {
      case 'open': {
        const [file] = await pickFiles(OPEN_ACCEPT, false)
        if (file) c.open(file)
        return
      }
      case 'rename':
        return c.rename?.()
      case 'details':
        return c.details?.()
      case 'close':
        return c.close?.()
      case 'save':
        return c.save()
    }
    const ops = c.ops.current
    if (action === 'import') {
      const [file] = await pickFiles('.csv,.tsv,.tab,.txt', false)
      if (!file) return
      if (!ops?.importText) throw new Error('The document is still loading — try again in a moment')
      await ops.importText(file, file.name.replace(/\.[^.]+$/, '').slice(0, 31) || 'Imported')
      return
    }
    if (!c.kind || !c.canDownload) return
    if (action === 'original') return c.downloadOriginal()
    if (action === 'picture') return c.picture?.download()
    if (!ops) throw new Error('The document is still loading — try again in a moment')
    if (action.startsWith('text:')) {
      const f = TEXT_FORMATS[c.kind].find((x) => 'text:' + x.id === action)
      if (!f) return
      download(new Blob([await ops.text(f.id)], { type: `${f.mime};charset=utf-8` }), c.base + f.ext)
      return
    }
    if (action === 'pdf' || action === 'print' || action === 'html') {
      c.notify(action === 'pdf' ? 'Preparing PDF… choose “Save as PDF” in the print dialog' : 'Preparing…')
      const p = await ops.printable()
      if (p.notice) c.notify(p.notice)
      if (action === 'html') download(await exportHtml(p, c.base), c.base + '.html')
      else await printDocument(p, c.base)
    }
  }

  const run = (action: string) => {
    const quick = QUICK.test(action)
    if (busyRef.current && !quick) return
    if (!quick) {
      busyRef.current = true
      setBusy(true)
    }
    Promise.resolve()
      .then(() => exec(action))
      .catch((e: Error) => opts.current.notify(e.message || String(e)))
      .finally(() => {
        if (quick) return
        busyRef.current = false
        setBusy(false)
      })
  }

  return { options, busy, run }
}

export function FileMenu({ actions }: { actions: FileActions }) {
  if (!actions.options.length) return null
  return <Dropdown title="File" label="File" className="de-file" options={actions.options} onChange={actions.run} menuWidth={280} disabled={actions.busy} />
}
