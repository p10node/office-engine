import { useRef, useState } from 'react'
import type { DocKind } from '../../core/opc'
import { MOD } from '../common/ContextMenu'
import type { FileOps } from '../common/Slots'
import { download, Dropdown, pickFiles, type DropOption, type IconName } from '../edit/Toolbar'
import { OPEN_ACCEPT } from './importFile'
import { exportHtml, printDocument } from './print'
import { useStrings } from '../strings'

const EXT: Record<DocKind, string> = { docx: '.docx', xlsx: '.xlsx', pptx: '.pptx' }

const TEXT_FORMATS: Record<DocKind, { id: 'txt' | 'md' | 'csv' | 'tsv'; ext: string; mime: string }[]> = {
  docx: [
    { id: 'txt', ext: '.txt', mime: 'text/plain' },
    { id: 'md', ext: '.md', mime: 'text/markdown' },
  ],
  xlsx: [
    { id: 'csv', ext: '.csv', mime: 'text/csv' },
    { id: 'tsv', ext: '.tsv', mime: 'text/tab-separated-values' },
  ],
  pptx: [{ id: 'txt', ext: '.txt', mime: 'text/plain' }],
}

/** Host-provided File menu entry (Share, Version history…). */
export interface FileMenuItem {
  label: string
  /** built-in icon name */
  icon?: IconName
  /** shortcut / note on the right */
  hint?: string
  disabled?: boolean
  onClick: () => void
}

export interface FileActionsOptions {
  kind: DocKind | undefined
  /** extension of the Office file when it is not the plain one of its kind (".xlsm", ".dotx"…) */
  ext?: string
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
  /** a mail message's attachments: each is saved from the menu */
  attachments?: { name: string; size: string; download: () => Promise<void> }[]
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
  const t = useStrings()
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
          { value: 'new:docx', label: t.file.newDocx, icon: 'doc', head: t.file.new },
          { value: 'new:xlsx', label: t.file.newXlsx, icon: 'table' },
          { value: 'new:pptx', label: t.file.newPptx, icon: 'slides' },
          { value: 'open', label: t.file.open, icon: 'open', hint: `${MOD}O`, sep: true },
        ]
      : [],
  )
  if (o.canImport && kind === 'xlsx') options.push({ value: 'import', label: t.file.importCsv, icon: 'table', sep: !o.canOpen && options.length > 0 })
  const fileItems: DropOption<string>[] = []
  if (kind && o.canSave) fileItems.push({ value: 'save', label: t.file.save, icon: 'save', hint: `${MOD}S` })
  if (kind && o.rename) fileItems.push({ value: 'rename', label: t.file.rename, icon: 'edit' })
  section(fileItems)
  if (o.canDownload && kind) {
    section([
      { value: 'original', label: t.kind[kind], icon: 'save', hint: o.ext || EXT[kind], head: t.file.downloadAs },
      { value: 'pdf', label: t.file.pdf, icon: 'pdf', hint: '.pdf' },
      ...(o.picture
        ? [{ value: 'picture', label: o.picture.label, icon: 'image', hint: o.picture.ext } as const]
        : [{ value: 'html', label: t.file.html, icon: 'html', hint: '.html' } as const, ...TEXT_FORMATS[kind].map((f) => ({ value: 'text:' + f.id, label: t.file[f.id], icon: 'doc', hint: f.ext }) as const)]),
    ])
  }
  if (o.canDownload && kind && o.attachments?.length) {
    section(o.attachments.map((a, i) => ({ value: `attachment:${i}`, label: a.name, icon: 'save', hint: a.size, ...(i ? {} : { head: t.file.attachments }) }) as DropOption<string>))
  }
  const tail: DropOption<string>[] = []
  if (kind && o.details) tail.push({ value: 'details', label: t.file.details, icon: 'info' })
  if (kind && o.canDownload) tail.push({ value: 'print', label: t.file.print, icon: 'print', hint: `${MOD}P` })
  section(tail)
  section(o.extra.map((x, i) => ({ value: `extra:${i}`, label: x.label, icon: x.icon, hint: x.hint, disabled: x.disabled })))
  if (o.close) section([{ value: 'close', label: t.file.close, icon: 'close' }])

  const exec = async (action: string) => {
    const c = opts.current
    if (action.startsWith('new:')) {
      const kind = action.slice(4)
      if (c.canOpen && (kind === 'docx' || kind === 'xlsx' || kind === 'pptx')) c.create(kind)
      return
    }
    if (action.startsWith('extra:')) {
      const item = c.extra[Number(action.slice(6))]
      if (item && !item.disabled) item.onClick()
      return
    }
    switch (action) {
      case 'open': {
        if (!c.canOpen) return
        const [file] = await pickFiles(OPEN_ACCEPT, false)
        if (file && opts.current.canOpen) opts.current.open(file)
        return
      }
      case 'rename':
        return c.kind ? c.rename?.() : undefined
      case 'details':
        return c.kind ? c.details?.() : undefined
      case 'close':
        return c.close?.()
      case 'save':
        return c.kind && c.canSave ? c.save() : undefined
    }
    const ops = c.ops.current
    if (action === 'import') {
      if (!c.canImport || c.kind !== 'xlsx') return
      const [file] = await pickFiles('.csv,.tsv,.tab,.txt', false)
      if (!file || !opts.current.canImport) return
      if (!ops?.importText) throw new Error(t.common.stillLoading)
      await ops.importText(file, file.name.replace(/\.[^.]+$/, '').slice(0, 31) || 'Imported')
      return
    }
    if (!c.kind || !c.canDownload) return
    if (action === 'original') return c.downloadOriginal()
    if (action === 'picture') return c.picture?.download()
    if (action.startsWith('attachment:')) return c.attachments?.[Number(action.slice(11))]?.download()
    if (!ops) throw new Error(t.common.stillLoading)
    await ops.flush?.()
    if (action.startsWith('text:')) {
      const f = TEXT_FORMATS[c.kind].find((x) => 'text:' + x.id === action)
      if (!f) return
      download(new Blob([await ops.text(f.id)], { type: `${f.mime};charset=utf-8` }), c.base + f.ext)
      return
    }
    if (action === 'pdf' || action === 'print' || action === 'html') {
      c.notify(action === 'pdf' ? t.file.preparingPdf : t.file.preparing)
      const p = await ops.printable()
      if (p.notice) c.notify(p.notice)
      if (action === 'html') download(await exportHtml(p, c.base, t), c.base + '.html')
      else await printDocument(p, c.base, t)
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
  const t = useStrings()
  if (!actions.options.length) return null
  return <Dropdown title={t.file.menu} label={t.file.menu} className="de-file" options={actions.options} onChange={actions.run} menuWidth={280} disabled={actions.busy} />
}
