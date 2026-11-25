import { useCallback, useEffect, useInsertionEffect, useMemo, useRef, useState, type CSSProperties, type ReactNode } from 'react'
import type { DocSource } from '../core/source'
import { DocumentClient, type OpenOptions } from '../core/worker/client'
import type { Zoom } from './common/ZoomControl'
import { DocxView } from './docx/DocxView'
import { ViewerContext, type FileOps, type ViewerApi } from './common/Slots'
import { StatusContext } from './common/Status'
import { ConfirmDialog, PromptDialog, Toast } from './common/Dialog'
import { download, Icon, IconsProvider, type Icons } from './edit/Toolbar'
import { UIComponentsProvider, type UIComponents } from './ui'
import { createViewerController, internalController, type ViewerController } from './controller'
import { DetailsDialog } from './file/DetailsDialog'
import { FileMenu, useFileActions, type FileMenuItem } from './file/FileMenu'
import { importFile, type Imported } from './file/importFile'
import { blankPptx, parasToDocx, rowsToXlsx } from '../core/convert'
import type { DocKind } from '../core/opc'
import { extOf, HAS_SOURCE, PLAIN_EXT, PLAIN_FILE, unpackedName, type PlainFormat } from '../core/plain'
import { injectStyles } from './injectStyles'
import { isTextEntry, useIsoLayoutEffect, useSize } from './common/hooks'
import { useThemeRoot, type Theme } from './theme'
import { SourceEditor } from './source/SourceEditor'
import { CodeView } from './source/CodeView'
import { LANGUAGE_EXT, LANGUAGE_NAME, languageOf } from '../core/syntax'
import { PictureView } from './picture/PictureView'
import { PptxView } from './pptx/PptxView'
import { useDocument } from './useDocument'
import { XlsxView } from './xlsx/XlsxView'
import type { CollabOptions } from '../collab'
import { CollabBar, PeerOverlay, useCollab } from '../collab/react'
import { CommentsContext, CommentsCtl, useCommentsState, type CommentOptions } from './comments/Comments'
import { COMMENT_KEYS, CommentsPanel } from './comments/CommentsPanel'
import { formatBytes, StringsProvider, useStrings, type Locale, type StringsOverride } from './strings'
import type { FontSource } from '../core/fonts'
import { FontsContext, useHostFonts } from './common/fonts'

export interface DocViewerProps {
  /**
   * File, Blob, ArrayBuffer, Uint8Array or URL (HTTP Range requests are used when the server supports them): an Office file,
   * an Apple Pages / Keynote / Numbers file, a Microsoft Visio drawing or an e-mail message (.eml, .msg; view only), CSV / TSV / text / Markdown / HTML (opened as a
   * workbook / document),
   * an SVG picture, a Mermaid / PlantUML / Graphviz diagram, or a source file (main.ts, script.py, config.json…: shown as code).
   */
  src: DocSource | null | undefined
  /** Display name (defaults to File.name or the URL's last segment); its extension also tells the format of a source without one. */
  name?: string
  /**
   * The file is source code in this language ("typescript", "ts", "python", "c++"…): a text file opens in the code
   * view, colored in it, whatever its name and MIME type say (a language the viewer does not color is shown plain).
   * Without it, a source file is told by its name's extension or its MIME type.
   */
  language?: string
  className?: string
  style?: CSSProperties
  /** Show the top bar (name, status, progress). Default true. */
  toolbar?: boolean
  /** Extra toolbar content (rendered at the right). */
  toolbarExtra?: ReactNode
  /** Initial zoom for documents and presentations. Default 'fit'. */
  zoom?: Zoom
  /** Show slide thumbnails for presentations. Default true. */
  thumbnails?: boolean
  headers?: Record<string, string>
  credentials?: RequestCredentials
  /** Custom worker factory, or false to parse on the main thread. */
  worker?: OpenOptions['worker']
  /** Where to inject the stylesheet (document or a shadow root). */
  styleTarget?: Document | ShadowRoot
  /**
   * Inject the stylesheet. Default true. Set false when the page loads `@p10node/office-engine/styles.css` itself
   * (a CSP without inline styles, a cascade layer of its own, a bundler that extracts CSS).
   */
  injectStyles?: boolean
  /**
   * Color scheme of the viewer chrome (bars, menus, dialogs): 'light', 'dark' or 'auto' (follows the OS / browser setting),
   * or a theme object with that mode and design token overrides (see docs/theming.md). The document itself keeps its own
   * colors. Default 'light'.
   */
  theme?: Theme
  onLoad?: (client: DocumentClient) => void
  onError?: (error: Error) => void
  /** Start in edit mode. Default false (view). */
  editable?: boolean
  /** Show the Edit/View toggle. Default true. */
  allowEdit?: boolean
  /** Show the Save/Download button and enable Ctrl+S. Default true. */
  allowSave?: boolean
  /**
   * Allow downloading files to the user's device (the save fallback without `onSave`, "Download image").
   * With `onSave`, saving still works when this is false. Default true.
   */
  allowDownload?: boolean
  /** Show the File menu (new, open, save, rename, download as PDF / HTML / text…, details, print) and its shortcuts. Default true. */
  fileMenu?: boolean
  /** File → New / Open: replace the document with a blank one or a local file (Office, Pages, Keynote, Numbers, Visio, e-mail, CSV, text, Markdown, HTML). Default true. */
  allowOpen?: boolean
  /** Extra File menu entries (Share, Version history…), shown before Close. */
  fileMenuItems?: FileMenuItem[]
  /** File → Rename: called with the new name (extension included). Rename is shown when `fileMenu` is on. */
  onRename?: (name: string) => void
  /** Adds File → Close. */
  onClose?: () => void
  /** Receives the saved file instead of downloading it. */
  onSave?: (file: Blob, name: string) => void | Promise<void>
  /** Called when the document gets / loses unsaved changes. */
  onDirtyChange?: (dirty: boolean) => void
  /**
   * Real-time collaboration: everyone opening the same base document (`src`) with the same `room`
   * edits it together through the hub at `transport` (see src/collab).
   */
  collab?: CollabOptions
  /**
   * Ctrl/⌘+A anywhere on the page selects the document (text fields outside the viewer keep their own).
   * Default true. With false the shortcut works from the viewer only (its content or toolbar).
   */
  captureSelectAll?: boolean
  /**
   * Comments (stored in the file, shared with collaborators): who writes them, who can be @mentioned, what the
   * host hears about. Shown by default, written while editing; `false` hides them.
   */
  comments?: CommentOptions | false
  /**
   * The host's icons, shown instead of the built-in ones: a map by name (see `ICON_NAMES`) or a resolver
   * `(name, { size }) => node`; a missing entry, or `undefined` / `null` from the resolver, keeps the built-in icon.
   * Each is rendered in `span.de-icon`, sized by font-size (1em).
   */
  icons?: Icons
  /**
   * The host's own UI components (`Button`, `Select`, `Separator`, `Dialog`, `Toast`), rendered in the toolbars and
   * dialogs instead of the built-in ones; each one left out keeps the built-in one. See docs/theming.md.
   */
  components?: Partial<UIComponents>
  /**
   * BCP 47 tag for number/date display and typed numbers ("de-DE" reads 1,5). An explicitly supplied tag
   * takes precedence over the table or enclosing StringsProvider. No tag uses a decimal point for input.
   * This does not select UI texts or change formulas or saved values.
   */
  locale?: Locale
  /** Texts replacing the enclosing provider's or English texts, key by key (see `Strings`): a group may be given in part, e.g. `{ file: { menu: 'Datei' } }`. */
  strings?: StringsOverride
  /**
   * The viewer's state and commands for the host (`createViewerController()`): zoom, page, outline, editing, unsaved
   * changes, save, undo… Subscribe to it to draw your own controls; see `headless`.
   */
  controller?: ViewerController
  /**
   * The document alone: no top bar, zoom control, outline panel, sheet tabs or slide list (the host draws its own,
   * driven by `controller`). While editing, the formatting toolbar keeps its row. The top bar shows when `toolbar` is
   * set explicitly. Default false.
   */
  headless?: boolean
  /**
   * Font files the host supplies (the library ships none): `{ family, src, weight?, style?, unicodeRange? }`, `src` a URL
   * or the bytes. A file is loaded when a document's text names its family, directly (`Calibri`) or as the fallback
   * the font list puts after it (`Carlito` for Calibri, `Caladea` for Cambria, `Liberation Sans / Serif / Mono` for
   * Arial, Times New Roman, Courier New); documents are laid out once it is there. One that fails is reported to
   * `onError`, one slower than 3 s is waited for no longer: the text is shown in the next font of its list, and laid
   * out again when it arrives.
   */
  fonts?: readonly FontSource[]
}

const EXT: Record<string, string> = { docx: '.docx', xlsx: '.xlsx', pptx: '.pptx' }
const IMPORTED_BADGE: Record<PlainFormat, string> = { csv: 'csv', tsv: 'tsv', text: 'txt', markdown: 'md', html: 'html', svg: 'svg', mermaid: 'mmd', plantuml: 'puml', graphviz: 'dot', code: 'code' }
/** plain formats saved back as themselves: Markdown and HTML built again from the document when it was edited, SVG and source files from their text */
const SAVES_SOURCE = new Set<PlainFormat>(['markdown', 'html', 'svg', 'code'])
/** the language of each plain format shown in the source pane, and what the pane calls it */
const SOURCE_LANG: Partial<Record<PlainFormat, [string, string]>> = {
  markdown: ['markdown', 'Markdown'],
  html: ['html', 'HTML'],
  svg: ['xml', 'SVG'],
  mermaid: ['mermaid', 'Mermaid'],
  plantuml: ['plantuml', 'PlantUML'],
  graphviz: ['graphviz', 'Graphviz'],
}
/** a Word document of source code: one code paragraph in its language (the colors are drawn, not stored) */
function codeDocx(text: string, language: string): Promise<Blob> {
  const code = text.replace(/^\uFEFF/, '').replace(/\r\n?/g, '\n')
  return parasToDocx([{ runs: code ? [{ text: code }] : [], code: language }])
}

/** a key per opened client: a view is built again for each (the edited source's previews have one name) */
const clientIds = new WeakMap<object, number>()
let nextClientId = 0
const clientKey = (c: object) => {
  let id = clientIds.get(c)
  if (id === undefined) clientIds.set(c, (id = ++nextClientId))
  return id
}

export function DocViewer(props: DocViewerProps) {
  return (
    <StringsProvider locale={props.locale} strings={props.strings}>
      <Viewer {...props} />
    </StringsProvider>
  )
}

function Viewer(props: DocViewerProps) {
  const t = useStrings()
  const fmtBytes = (n: number) => formatBytes(t, n)
  const { src, name, className, style, toolbarExtra, zoom, thumbnails, headers, credentials, worker, styleTarget, theme, onLoad, onError } = props
  const headless = !!props.headless
  const toolbar = props.toolbar ?? !headless
  // without a controller from the host the viewer keeps its own: the view publishes to one either way
  const ownCtl = useMemo(() => createViewerController(), [])
  const controller = internalController(props.controller ?? ownCtl)
  const ctl = useMemo(() => controller.createBinding(), [controller])
  const root = useThemeRoot(theme, style)
  const fonts = useHostFonts(props.fonts, onError)
  const { allowEdit = true, allowSave = true, allowDownload = true, fileMenu = true, allowOpen = true, captureSelectAll = true, onSave, onDirtyChange } = props
  const [editing, setEditing] = useState(!!props.editable)
  useEffect(() => {
    if (props.editable !== undefined) setEditing(!!props.editable)
  }, [props.editable])
  const [dirty, setDirtyState] = useState(false)
  const [saving, setSaving] = useState(false)
  const [barEl, setBarEl] = useState<HTMLDivElement | null>(null)
  const inject = props.injectStyles ?? true
  useInsertionEffect(() => {
    if (inject) injectStyles(styleTarget)
  }, [styleTarget, inject])
  // File → New / Open: a local file shown instead of `src` until the host passes another `src`
  const [opened, setOpened] = useState<Imported | null>(null)
  const [renamed, setRenamed] = useState<string | null>(null)
  useEffect(() => {
    setOpened(null)
    setRenamed(null)
  }, [src])
  const loadName = opened ? opened.name : name
  // the name is part of what is loaded: Rename only changes what is shown and saved
  // a preview of the edited source keeps the one shown until it is ready
  // (the host's language hint is for its own file, not one opened from the File menu)
  // (a mail message's header is written in the viewer's language)
  const mailTexts = useMemo(() => ({ texts: t.mail, locale: t.locale }), [t])
  const { client, error, loading, progress } = useDocument(opened ? opened.src : src, { name: loadName, headers, credentials, worker, language: opened ? undefined : props.language, mail: mailTexts }, !!opened?.live)
  // Save ownership follows the document, including a source change before its client is ready.
  const saveSession = useMemo(() => ({ live: true, saving: false, revision: 0 }), [client, src, opened])
  useEffect(() => {
    saveSession.live = true
    setSaving(false)
    return () => {
      saveSession.live = false
    }
  }, [saveSession])
  const viewCtl = useMemo(() => ctl.createView(), [ctl, client])
  const docName = renamed ?? loadName
  const collab = useCollab(client, props.collab)
  const srcFormat = client?.info?.imported
  // an SVG file: the picture alone, read only (its text is edited in the source pane)
  const isPicture = srcFormat === 'svg'
  // a source file: its code, colored, edited as text (with collaboration: read only, the session shares the document)
  const isCode = srcFormat === 'code'
  const codeLang = client?.info?.language ?? ''
  // converted from another format (Apple Pages, Keynote, Numbers, Microsoft Visio, a mail message): view only
  const readOnly = !!client?.info?.from || isPicture || (isCode && !!props.collab)
  // with collaboration, editing waits for the shared history (offline edits sync on reconnect)
  const canEdit = !readOnly && (!props.collab || collab.status === 'live' || collab.status === 'offline')
  const [statusEl, setStatusEl] = useState<HTMLSpanElement | null>(null)

  // ---- source (Markdown, HTML, SVG, diagrams): edited side by side with the document built from it ----
  // (a source file's code is the view itself)
  const hasSource = !!srcFormat && HAS_SOURCE.has(srcFormat) && !isCode
  const [codeOpen, setCodeOpen] = useState(false)
  const codeShown = codeOpen && hasSource
  // Markdown / HTML / SVG files are saved in their own format; Markdown and HTML are built again from an edited document
  const saveFormat = srcFormat && SAVES_SOURCE.has(srcFormat) ? srcFormat : null
  const rebuilds = saveFormat === 'markdown' || saveFormat === 'html'
  /** what was edited last: the document (its source is stale) or the source (the document is built from it) */
  const lastEdit = useRef<'doc' | 'source' | null>(null)
  const fileOps = useRef<FileOps | null>(null)
  const [code, setCode] = useState<string | null>(null)
  const sourceEdited = useRef(false)
  const crlf = useRef(false)
  const reloadTimer = useRef<ReturnType<typeof setTimeout> | undefined>(undefined)
  useEffect(() => () => clearTimeout(reloadTimer.current), [])
  // another document (not a preview of the edited source): its own source
  useEffect(() => {
    if (opened?.live) return
    clearTimeout(reloadTimer.current)
    setCode(null)
    sourceEdited.current = false
    lastEdit.current = null
  }, [src, opened])
  // the source shown: the file's text, or (Markdown / HTML edited as a document) the text built from the document
  useEffect(() => {
    if (!client || (!codeShown && !isCode)) return
    const stale = rebuilds && lastEdit.current === 'doc'
    if (code !== null && !stale) return
    let live = true
    const ops = fileOps.current
    // the source the document was built from: what was not edited is kept from it
    const built =
      stale && ops
        ? Promise.resolve(code ?? (client.call('plain.source') as Promise<string | null>))
            .then((base) => ops.text(saveFormat === 'html' ? 'html' : 'md', base ?? undefined))
            .catch(() => null)
        : Promise.resolve(null)
    built
      .then(async (t) => (typeof t === 'string' ? t : ((await client.call('plain.source')) as string | null)))
      .then(
        (t) => {
          if (!live) return
          if (stale) lastEdit.current = null
          // a source file in CRLF: the text edited (LF, as a textarea gives it) is saved in CRLF too
          if (isCode && code === null) crlf.current = !!t && /\r\n/.test(t) && !/(^|[^\r])\n/.test(t)
          setCode(t ?? '')
        },
        () => {},
      )
    return () => {
      live = false
    }
  }, [client, codeShown, isCode, code, rebuilds, saveFormat])
  const [bodyRef, bodySize] = useSize<HTMLDivElement>()
  const mainRef = useRef<HTMLDivElement>(null)
  const stacked = bodySize.w > 0 && bodySize.w < 720
  const [split, setSplit] = useState(0.45)
  // the preview's scroll position, kept across rebuilds
  const restoreScroll = useRef<number | null>(null)
  useEffect(() => {
    const ratio = restoreScroll.current
    if (ratio === null || !client) return
    restoreScroll.current = null
    let raf = 0
    let frames = 0
    let lastH = -1
    const tick = () => {
      const sc = mainRef.current?.querySelector<HTMLElement>('.de-scroll')
      if (sc && sc.scrollHeight > sc.clientHeight + 1) {
        sc.scrollTop = ratio * (sc.scrollHeight - sc.clientHeight)
        if (sc.scrollHeight === lastH) return
        lastH = sc.scrollHeight
      }
      if (++frames < 40) raf = requestAnimationFrame(tick)
    }
    raf = requestAnimationFrame(tick)
    return () => cancelAnimationFrame(raf)
  }, [client])

  // ---- comments ----
  const cmOn = props.comments !== false && !isPicture && !isCode
  const cmCtl = useMemo(() => (client && cmOn ? new CommentsCtl(client) : null), [client, cmOn])
  useEffect(() => () => cmCtl?.dispose(), [cmCtl])
  const cmOpts = props.comments || {}
  if (cmCtl) {
    cmCtl.opts = cmOpts
    cmCtl.t = t
    cmCtl.me = cmOpts.user ?? props.collab?.user ?? { name: 'Me' }
    cmCtl.canComment = !readOnly && (cmOpts.canComment ?? (editing && canEdit))
  }
  const cmState = useCommentsState(cmCtl)
  const cmOpenAtLoad = !!cmOpts.open
  useEffect(() => {
    if (cmCtl && cmOpenAtLoad) cmCtl.toggle(true)
  }, [cmCtl]) // eslint-disable-line react-hooks/exhaustive-deps
  const cmCount = cmCtl ? cmCtl.visible().length : 0
  // comments.focus: the thread shown once it is in the document (DOCX threads appear as the text loads)
  const cmFocus = cmOpts.focus
  const cmFocused = useRef<string | null>(null)
  useEffect(() => {
    if (!cmCtl || !cmFocus || cmFocused.current === cmFocus || !cmCtl.loaded) return
    const c = cmCtl.list.get(cmFocus)
    const root = c?.parent ?? cmFocus
    if (!cmCtl.threads().some((t) => t.root.id === root)) return
    cmFocused.current = cmFocus
    cmCtl.toggle(true)
    if (cmCtl.list.get(root)?.done) cmCtl.setShowResolved(true)
    cmCtl.select(root)
  }, [cmCtl, cmFocus, cmState.rev])

  // once per document: a new `onLoad` function (a host's inline callback, with every render) is not another load
  const onLoadRef = useRef(onLoad)
  onLoadRef.current = onLoad
  useEffect(() => {
    if (client) onLoadRef.current?.(client)
  }, [client])
  useEffect(() => {
    if (error) onError?.(error)
  }, [error, onError])

  // a converted file (CSV → xlsx…) starts unsaved
  const openedDirty = useRef(false)
  // The host is told of a change once, from where it happened (an effect, an event): not from inside a state updater,
  // which React runs while rendering (twice in StrictMode), where a host that sets its own state is in the wrong.
  const dirtyNow = useRef(false)
  const onDirtyRef = useRef(onDirtyChange)
  onDirtyRef.current = onDirtyChange
  const changeDirty = useCallback((d: boolean) => {
    if (dirtyNow.current === d) return
    dirtyNow.current = d
    setDirtyState(d)
    onDirtyRef.current?.(d)
  }, [])
  useEffect(() => {
    // (the host is told too: a new document is not unsaved)
    changeDirty(!!client && openedDirty.current)
    if (client) openedDirty.current = false
  }, [client, changeDirty])
  const setDirty = useCallback(
    (d: boolean) => {
      if (!saveSession.live) return
      if (d) saveSession.revision++
      changeDirty(d)
    },
    [changeDirty, saveSession],
  )
  const baseName = (renamed ?? (client?.name || docName || 'document')).replace(/\.[^.]+$/, '')
  // the Office file as it is saved: a workbook with macros stays an .xlsm, a template an .xltx (named otherwise, Excel refuses it)
  const officeExt = client ? (client.info?.ext ? `.${client.info.ext}` : EXT[client.kind]) : ''
  const fileName = baseName + officeExt
  // what Save writes: the Office file, or a Markdown / HTML / SVG / source file in its own format (a source file keeps its extension)
  const codeLanguage = LANGUAGE_NAME[languageOf(codeLang) ?? ''] ?? codeLang
  const codeExt = isCode ? extOf(unpackedName(client?.name ?? '') ?? client?.name) || LANGUAGE_EXT[languageOf(codeLang) ?? ''] || 'txt' : ''
  const saveExt = isCode ? `.${codeExt}` : saveFormat ? `.${PLAIN_FILE[saveFormat].ext}` : officeExt
  const saveName = baseName + saveExt
  /** the text of a file saved in its own format: the source as edited, or built from the edited document */
  const sourceText = useCallback(async (): Promise<string> => {
    if (!client) return ''
    if (rebuilds && lastEdit.current === 'doc' && fileOps.current) {
      const text = (await fileOps.current.text(saveFormat === 'html' ? 'html' : 'md', code ?? ((await client.call('plain.source')) as string | null) ?? undefined)) as string
      lastEdit.current = null
      setCode(text)
      return text
    }
    const text = code ?? ((await client.call('plain.source')) as string | null) ?? ''
    return isCode && crlf.current ? text.replace(/\r?\n/g, '\r\n') : text
  }, [client, rebuilds, saveFormat, code, isCode])
  // saving without onSave means downloading
  const canSave = allowSave && (!!onSave || allowDownload)
  const saveFile = useCallback(async (original = false) => {
    if (!client || !saveSession.live || saveSession.saving || !(original ? allowDownload : canSave)) return
    saveSession.saving = true
    setSaving(true)
    try {
      await fileOps.current?.flush?.()
      if (!saveSession.live) return
      const revision = saveSession.revision
      // a source file not edited: its bytes as they came (its encoding, byte order mark and line breaks);
      // as a Word document: its code as it is now, one code paragraph
      const kept = !isCode
        ? null
        : original
          ? await codeDocx(await sourceText(), codeLang)
          : !sourceEdited.current
            ? ((await client.call('plain.original')) as Blob | null)
            : null
      const blob = kept ?? (!original && saveFormat ? new Blob([await sourceText()], { type: `${PLAIN_FILE[saveFormat].mime};charset=utf-8` }) : await client.call('save'))
      if (!saveSession.live) return
      if (!original && onSave) await onSave(blob, saveName)
      else download(blob, original ? fileName : saveName)
      if (saveSession.live && saveSession.revision === revision && (!original || !onSave)) setDirty(false)
    } catch (e) {
      if (!saveSession.live) return
      onError?.(e as Error)
      console.error('[office-engine] save failed', e)
      setToast((e as Error).message)
    } finally {
      saveSession.saving = false
      if (saveSession.live) setSaving(false)
    }
  }, [client, saveSession, canSave, allowDownload, onSave, fileName, saveName, saveFormat, sourceText, setDirty, onError, isCode, codeLang])
  const save = useCallback(() => saveFile(), [saveFile])
  // an edit in the document: its source (Markdown / HTML) is built again when shown or saved
  const setDocDirty = useCallback(
    (d: boolean) => {
      if (d) lastEdit.current = 'doc'
      setDirty(d)
    },
    [setDirty],
  )
  const selectAll = useRef<(() => void) | null>(null)
  // next to its source, the document is a preview: it is edited through the source (a source file's code is edited as text)
  const docEditable = editing && canEdit && !codeShown && !isCode
  const api = useMemo<ViewerApi>(
    () => ({ bar: barEl, editable: docEditable, setDirty: setDocDirty, save, canDownload: allowDownload, client, fileOps, selectAll, ctl: viewCtl, headless }),
    [barEl, docEditable, setDocDirty, save, allowDownload, client, viewCtl, headless],
  )

  const editSource = (text: string) => {
    setCode(text)
    if (!sourceEdited.current) setDirty(true)
    sourceEdited.current = true
    lastEdit.current = 'source'
    clearTimeout(reloadTimer.current)
    const fmt = srcFormat
    if (!fmt) return
    reloadTimer.current = setTimeout(() => {
      const sc = mainRef.current?.querySelector<HTMLElement>('.de-scroll')
      restoreScroll.current = sc && sc.scrollHeight > sc.clientHeight ? sc.scrollTop / (sc.scrollHeight - sc.clientHeight) : null
      // the same name: the same format, title and saved name
      let n = client?.name || loadName || `document.${PLAIN_FILE[fmt].ext}`
      if (PLAIN_EXT[extOf(unpackedName(n) ?? n)] !== fmt) n += `.${PLAIN_FILE[fmt].ext}`
      openedDirty.current = true
      setOpened({ src: new Blob([text], { type: PLAIN_FILE[fmt].mime }), name: n, converted: true, live: true })
    }, 350)
  }
  // a source file's code as shown and edited: in LF, as a textarea holds it (saved in the file's own line breaks)
  const shownCode = useMemo(() => (isCode && code !== null ? code.replace(/\r\n?/g, '\n') : null), [isCode, code])
  /** an edit of a source file's code: what is saved (nothing is built from it) */
  const editCode = (text: string) => {
    setCode(text)
    sourceEdited.current = true
    setDirty(true)
  }
  const sourceName = () => `${(renamed ?? client?.name ?? docName ?? 'document').replace(/\.[^.]+$/, '')}.${PLAIN_FILE[srcFormat ?? 'text'].ext}`
  const dragSplit = (e: React.PointerEvent<HTMLDivElement>) => {
    const body = bodyRef.current
    if (!body) return
    e.preventDefault()
    const handle = e.currentTarget
    handle.setPointerCapture(e.pointerId)
    const r = body.getBoundingClientRect()
    const move = (ev: PointerEvent) => setSplit(Math.min(0.85, Math.max(0.15, stacked ? (ev.clientY - r.top) / r.height : (ev.clientX - r.left) / r.width)))
    const up = () => {
      handle.removeEventListener('pointermove', move)
      handle.removeEventListener('pointerup', up)
      handle.removeEventListener('pointercancel', up)
    }
    handle.addEventListener('pointermove', move)
    handle.addEventListener('pointerup', up)
    handle.addEventListener('pointercancel', up)
  }
  const rootRef = useRef<HTMLDivElement>(null)
  // captureSelectAll: Ctrl/⌘+A from the rest of the page too (the viewer's own keys are handled in onKeyDown below)
  useEffect(() => {
    const root = rootRef.current
    if (!captureSelectAll || !root) return
    const on = (e: KeyboardEvent) => {
      if (e.defaultPrevented || !(e.metaKey || e.ctrlKey) || e.altKey || e.shiftKey || e.key.toLowerCase() !== 'a') return
      const target = (e.composedPath()[0] ?? e.target) as Node | null
      if ((target && root.contains(target)) || isTextEntry(target) || !selectAll.current) return
      e.preventDefault()
      selectAll.current()
    }
    const doc = root.ownerDocument
    doc.addEventListener('keydown', on)
    return () => doc.removeEventListener('keydown', on)
  }, [captureSelectAll])

  // Ctrl/⌘+Alt+M: comment on the selection — from the viewer, or the page around it when nothing else has the focus
  // (a click on a slide shape leaves it on the page)
  const cmRef = useRef(cmCtl)
  cmRef.current = cmCtl
  useEffect(() => {
    const root = rootRef.current
    if (!root) return
    const on = (e: KeyboardEvent) => {
      const ctl = cmRef.current
      if (e.defaultPrevented || !(e.metaKey || e.ctrlKey) || !e.altKey || e.shiftKey || e.code !== 'KeyM' || !ctl?.canComment) return
      const target = (e.composedPath()[0] ?? e.target) as Node | null
      const doc = root.ownerDocument
      if (!(target && root.contains(target)) && target !== doc.body && target !== doc.documentElement) return
      e.preventDefault()
      ctl.start()
    }
    const doc = root.ownerDocument
    doc.addEventListener('keydown', on)
    return () => doc.removeEventListener('keydown', on)
  }, [])

  // ---- File menu ----
  const [toast, setToast] = useState<string | null>(null)
  const clearToast = useCallback(() => setToast(null), [])
  // the file opened but not whole (a CSV with more columns than a sheet holds): what is missing is said, not left to be found
  useEffect(() => client?.on('warning', (w) => setToast(w.text)), [client])
  const [confirm, setConfirm] = useState<{ message: string; run: () => void } | null>(null)
  const [renaming, setRenaming] = useState(false)
  const [details, setDetails] = useState(false)
  /** replaces the document, after confirmation when there are unsaved changes */
  const replace = (message: string, run: () => void) => (dirty ? setConfirm({ message, run }) : run())
  const show = useCallback((doc: Imported, edit = false) => {
    // a file shown as it is, converted or not, has no changes: only the preview of a source being edited is unsaved
    openedDirty.current = !!doc.live
    setRenamed(null)
    setOpened(doc)
    if (edit) setEditing(true)
  }, [])
  const canOpen = allowOpen && !props.collab
  const fileActions = useFileActions({
    kind: client?.kind,
    ext: officeExt,
    base: baseName,
    canOpen,
    canImport: allowEdit && canEdit,
    canSave,
    canDownload: allowDownload && !!client,
    ops: fileOps,
    create: (kind: DocKind) =>
      replace(t.bar.replaceNew(t.newWhat[kind]), () => {
        const blank = kind === 'docx' ? parasToDocx([]) : kind === 'xlsx' ? rowsToXlsx([], 'Sheet1') : blankPptx()
        blank.then((src) => show({ src, name: t.untitled[kind] + EXT[kind], converted: false }, allowEdit), (e: Error) => setToast(e.message))
      }),
    open: (file) =>
      replace(t.bar.replaceOpen(file.name), () => {
        importFile(file, t.file.unsupported)
          .then(async (doc) => {
            // Unsaved changes go with the document shown: not for a file that does not open (a damaged one, a
            // protected one), of which the error would be all that is left. It is opened once to see, then shown.
            if (dirty) (await DocumentClient.open(doc.src, { name: doc.name, worker })).destroy()
            show(doc)
          })
          .catch((e: Error) => setToast(e.message))
      }),
    save,
    downloadOriginal: () => saveFile(true),
    picture: isPicture
      ? {
          label: t.file.picture,
          ext: '.svg',
          download: async () => download(new Blob([await sourceText()], { type: 'image/svg+xml' }), saveName),
        }
      : undefined,
    attachments: client?.info?.mail?.attachments.map((a, i) => ({
      name: a.name,
      size: fmtBytes(a.size),
      download: async () => download((await client.call('mail.attachment', i)) as Blob, a.name),
    })),
    rename: () => setRenaming(true),
    details: () => setDetails(true),
    close: props.onClose,
    extra: props.fileMenuItems ?? [],
    notify: setToast,
  })
  const onKeyDown = (e: React.KeyboardEvent) => {
    if (e.defaultPrevented || !(e.metaKey || e.ctrlKey) || e.altKey || e.shiftKey) return
    const k = e.key.toLowerCase()
    // Ctrl/⌘+S in the source: the viewer's Save. And anywhere else in the viewer while a document is edited, where the
    // view did not take the key itself: on a sheet tab, a thumbnail, a button of the bar (never the browser's "save page").
    if (k === 's' && ((e.target as HTMLElement).closest?.('.de-src') || (editing && canSave && !!client))) {
      e.preventDefault()
      void save()
      return
    }
    if (k === 'a') {
      // from the viewer's own chrome: the document, not the whole page (text fields keep their own select-all)
      if (isTextEntry(e.target) || !selectAll.current) return
      e.preventDefault()
      selectAll.current()
    } else if (k === 'p' && !allowDownload) {
      // with downloads disabled, the browser must not print (save) the page either: with or without the File menu
      e.preventDefault()
    } else if (!fileMenu) return
    else if (k === 'p') {
      e.preventDefault()
      fileActions.run('print')
    } else if (k === 'o' && canOpen) {
      e.preventDefault()
      fileActions.run('open')
    }
  }

  const pct = progress && progress.total > 0 ? Math.min(100, (progress.loaded / progress.total) * 100) : 0
  const busy = loading || (!!progress && pct < 100)
  const title = renamed ?? client?.name ?? docName ?? (typeof src === 'string' ? decodeURIComponent(src.split(/[?#]/)[0].split('/').pop() ?? '') : (src as File | undefined)?.name) ?? ''

  // ---- controller: the viewer's part of the state, and its commands ----
  const editAllowed = !readOnly && !codeShown
  useEffect(() => {
    ctl.update({ kind: client?.kind ?? null, loading: busy, editing: docEditable, canEdit: !!client && editAllowed && canEdit, dirty, saving })
  }, [ctl, client, busy, docEditable, editAllowed, canEdit, dirty, saving])
  const commands = useRef({ editAllowed, save, fileActions, cmCtl })
  commands.current = { editAllowed, save, fileActions, cmCtl }
  useIsoLayoutEffect(() => {
    return ctl.attachViewer({
      setEditing: (on) => {
        if (!on || commands.current.editAllowed) setEditing(on)
      },
      save: () => commands.current.save(),
      fileAction: (id) => commands.current.fileActions.run(id),
      toggleComments: (open) => commands.current.cmCtl?.toggle(open),
    })
  }, [ctl])
  useIsoLayoutEffect(() => viewCtl.activate(), [viewCtl])

  let body: ReactNode
  if (error) {
    body = (
      <div className="de-center">
        <div className="de-error" role="alert">
          <strong>{t.bar.cannotOpen}</strong>
          <div>{error.message}</div>
        </div>
      </div>
    )
  } else if (!client) {
    body = src || opened ? (
      <div className="de-center">
        <div className="de-spinner" />
        <div>{progress?.phase === 'download' && progress.total ? t.bar.downloading(fmtBytes(progress.loaded), fmtBytes(progress.total)) : t.bar.opening}</div>
      </div>
    ) : (
      <div className="de-center">{t.bar.noDocument}</div>
    )
  } else if (isPicture) body = <PictureView key={clientKey(client)} client={client} />
  else if (isCode)
    body =
      code === null || shownCode === null ? (
        <div className="de-center">
          <div className="de-spinner" />
        </div>
      ) : (
        <CodeView
          key={clientKey(client)}
          text={shownCode}
          language={codeLang}
          readOnly={!(editing && canEdit && allowEdit)}
          onChange={editCode}
          onCopy={() =>
            navigator.clipboard?.writeText(code).then(
              () => setToast(t.bar.codeCopied),
              () => setToast(t.bar.cannotCopy),
            )
          }
        />
      )
  else if (client.kind === 'docx') body = <DocxView key={clientKey(client)} client={client} initialZoom={zoom} />
  else if (client.kind === 'xlsx') body = <XlsxView client={client} />
  else body = <PptxView client={client} initialZoom={zoom} thumbnails={thumbnails} />

  return (
    <IconsProvider icons={props.icons}>
    <UIComponentsProvider components={props.components}>
    <div ref={rootRef} className={root.className + (headless ? ' de-headless' : '') + (className ? ` ${className}` : '')} style={root.style} onKeyDown={onKeyDown}>
      {toolbar && (
        <div className="de-bar">
          {client && (
            <span className={`de-badge de-badge-${client.kind}`}>
              {client.info?.from ?? client.info?.converted ?? (isCode ? codeExt : client.info?.imported ? IMPORTED_BADGE[client.info.imported] : client.kind)}
            </span>
          )}
          {fileMenu && <FileMenu actions={fileActions} />}
          <span className="de-title" title={title}>
            {title}
          </span>
          <span className="de-status" ref={setStatusEl} hidden={!client} />
          {client && client.size > 0 && <span className="de-status">{fmtBytes(client.size)}</span>}
          {client && props.collab && <CollabBar client={client} state={collab} />}
          {toolbarExtra}
          {client && hasSource && (
            <button
              type="button"
              className={'de-mode' + (codeShown ? ' de-mode-on' : '')}
              onClick={() => setCodeOpen((v) => !v)}
              aria-pressed={codeShown}
              title={codeShown ? t.bar.codeHide : t.bar.codeShow}
              data-testid="de-code-toggle"
            >
              <Icon name="code" size={16} />
              {t.bar.code}
            </button>
          )}
          {cmCtl && (
            <button
              type="button"
              className={'de-mode de-cm-toggle' + (cmState.open ? ' de-mode-on' : '')}
              onClick={() => cmCtl.toggle()}
              title={cmCtl.canComment ? t.bar.commentsTitle(COMMENT_KEYS) : t.bar.comments}
              aria-pressed={cmState.open}
              data-testid="de-comments-toggle"
            >
              <Icon name="comment" size={16} />
              {t.bar.comments}
              {cmCount > 0 && <span className="de-cm-n">{cmCount}</span>}
            </button>
          )}
          {client && allowEdit && !readOnly && !codeShown && (
            <button type="button" className={editing ? 'de-mode de-mode-on' : 'de-mode'} onClick={() => setEditing((v) => !v)} title={editing ? t.bar.toView : t.bar.toEdit}>
              <Icon name={editing ? 'view' : 'edit'} size={16} />
              {editing ? t.bar.view : t.bar.edit}
            </button>
          )}
          {client && canSave && ((editing && !readOnly) || dirty) && (
            <button type="button" className={dirty ? 'de-mode de-save-dirty' : 'de-mode'} disabled={saving} onClick={() => void save()} title={t.bar.saveTitle}>
              <Icon name="save" size={16} />
              {saving ? t.bar.saving : dirty || onSave ? t.bar.save : t.bar.download}
            </button>
          )}
          <div className="de-progress" style={{ width: `${busy ? Math.max(pct, 3) : 100}%`, opacity: busy ? 1 : 0 }} />
        </div>
      )}
      {(toolbar || headless) && client && docEditable && <div className="de-ebar-host" ref={setBarEl} />}
      <div className={'de-body' + (codeShown && stacked ? ' de-body-stacked' : '')} ref={bodyRef}>
        {codeShown && (
          <div className="de-src-pane" style={stacked ? { height: `${split * 100}%` } : { width: `${split * 100}%` }}>
            {code === null ? (
              <div className="de-center">
                <div className="de-spinner" />
              </div>
            ) : (
              <SourceEditor
                text={code}
                language={SOURCE_LANG[srcFormat!]?.[0] ?? 'text'}
                label={SOURCE_LANG[srcFormat!]?.[1]}
                onChange={editSource}
                readOnly={!allowEdit || !!props.collab}
                notice={
                  dirty && !sourceEdited.current && !rebuilds
                    ? t.bar.sourceStale
                    : undefined
                }
                onCopy={() =>
                  navigator.clipboard?.writeText(code).then(
                    () => setToast(t.bar.sourceCopied),
                    () => setToast(t.bar.cannotCopy),
                  )
                }
                onDownload={allowDownload ? () => download(new Blob([code], { type: PLAIN_FILE[srcFormat!].mime }), sourceName()) : undefined}
              />
            )}
          </div>
        )}
        {codeShown && (
          <div
            className="de-split-handle"
            role="separator"
            aria-orientation={stacked ? 'horizontal' : 'vertical'}
            aria-label={t.bar.resizeSource}
            aria-valuenow={Math.round(split * 100)}
            aria-valuemin={15}
            aria-valuemax={85}
            tabIndex={0}
            onPointerDown={dragSplit}
            onKeyDown={(e) => {
              const d = e.key === (stacked ? 'ArrowUp' : 'ArrowLeft') ? -0.05 : e.key === (stacked ? 'ArrowDown' : 'ArrowRight') ? 0.05 : 0
              if (!d) return
              e.preventDefault()
              setSplit((v) => Math.min(0.85, Math.max(0.15, v + d)))
            }}
          />
        )}
        <div className="de-main" ref={mainRef}>
          <ViewerContext.Provider value={api}>
            <CommentsContext.Provider value={cmCtl}>
              <StatusContext.Provider value={statusEl}>
                <FontsContext.Provider value={fonts}>{body}</FontsContext.Provider>
              </StatusContext.Provider>
            </CommentsContext.Provider>
          </ViewerContext.Provider>
        </div>
        {client && collab.collab && <PeerOverlay client={client} collab={collab.collab} />}
        {cmCtl && cmState.open && <CommentsPanel ctl={cmCtl} />}
      </div>
      {toast && <Toast text={toast} onDone={clearToast} />}
      {confirm && <ConfirmDialog title={t.bar.discardTitle} message={confirm.message} ok={t.bar.discardOk} onOk={confirm.run} onClose={() => setConfirm(null)} />}
      {renaming && client && (
        <PromptDialog
          spec={{
            title: t.bar.renameTitle,
            label: t.bar.renameLabel,
            initial: baseName,
            ok: t.bar.renameOk,
            onOk: (v) => {
              const next = v.trim().replace(/[\\/:*?"<>|]+/g, ' ').replace(new RegExp(`(\\${officeExt}|\\${saveExt})$`, 'i'), '')
              if (!next) return
              setRenamed(next + officeExt)
              props.onRename?.(next + saveExt)
            },
          }}
          onClose={() => setRenaming(false)}
        />
      )}
      {details && client && (
        <DetailsDialog
          client={client}
          file={[
            [t.details.name, saveName],
            [
              t.details.type,
              client.info?.from
                ? t.details.viewOnly(t.from[client.info.from])
                : client.info?.converted
                  ? t.details.openedAs(t.converted[client.info.converted], t.kind[client.kind])
                  : isCode
                    ? codeLanguage ? `${t.imported.code} (${codeLanguage})` : t.imported.code
                    : client.info?.imported
                      ? t.details.openedAs(t.imported[client.info.imported], t.kind[client.kind])
                      : t.kind[client.kind],
            ],
            ...(client.size > 0 ? [[t.details.fileSize, fmtBytes(client.size)] as [string, string]] : []),
            ...(dirty ? [[t.details.status, t.details.unsaved] as [string, string]] : []),
          ]}
          stats={fileOps.current?.stats() ?? []}
          onClose={() => setDetails(false)}
        />
      )}
    </div>
    </UIComponentsProvider>
    </IconsProvider>
  )
}
