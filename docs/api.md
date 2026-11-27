# API

## React

```tsx
import { DocViewer } from '@p10node/office-engine'

const fileTexts = { file: { menu: 'Files' } }

<DocViewer src={file} style={{ height: 600 }} />                 // File | Blob | ArrayBuffer | Uint8Array
<DocViewer src="https://cdn.example.com/report.xlsx" />          // URL → Range requests
<DocViewer src={url} headers={{ Authorization: `Bearer ${t}` }} credentials="include" />
<DocViewer src={file} editable onSave={(blob, name) => upload(blob, name)} />  // editing
<DocViewer src={file} allowSave={false} allowDownload={false} />  // no Save/Download, no image download
<DocViewer src={file} fileMenu={false} />                         // no File menu
<DocViewer src={file} theme="dark" />                             // dark chrome; "auto" follows the OS setting
<DocViewer src={file} editable locale="de-DE" />                  // numbers typed as 1,5 and 1.234,5
<DocViewer src={file} fonts={[{ family: 'Carlito', src: '/fonts/Carlito-Regular.ttf' }]} />  // font files, see Fonts
<DocViewer src={file} strings={fileTexts} />                       // override individual UI texts
```

### File menu

| Item                      | DOCX                       | XLSX                           | PPTX                       |
|---------------------------|----------------------------|--------------------------------|----------------------------|
| New                       | blank document             | blank spreadsheet              | blank presentation         |
| Open… (`Ctrl+O`)          | any supported file         | any supported file             | any supported file         |
| Import CSV into new sheet | –                          | `.csv` / `.tsv`                | –                          |
| Save (`Ctrl+S`)           | ✓                          | ✓                              | ✓                          |
| Rename…                   | ✓                          | ✓                              | ✓                          |
| Download as (native)      | `.docx`                    | `.xlsx`                        | `.pptx`                    |
| PDF                       | pages as shown             | current sheet, fit to width    | one slide per page         |
| Web page                  | `.html`, images inlined    | `.html`                        | `.html`, images inlined    |
| Text                      | `.txt`, `.md`              | `.csv`, `.tsv` (current sheet) | `.txt` (incl. notes)       |
| Document details          | pages, words, characters   | sheets, used range             | slides, slide size         |
| Print (`Ctrl+P`)          | ✓                          | ✓                              | ✓                          |
| Host items, Close         | `fileMenuItems`, `onClose` | `fileMenuItems`, `onClose`     | `fileMenuItems`, `onClose` |

Open converts other formats: `.csv` / `.tsv` → workbook, `.txt` / `.md` / `.html` → document (headings, lists, bold / italic, links), `.mmd` / `.puml` / `.dot` / `.gv` → document with the drawing. Converted files start unsaved; Save writes the Office file.
New / Open replace the shown document until the host passes another `src` (unsaved changes are confirmed first; both are hidden during collaboration).
Rename changes the shown name and the name used by Save / Download (`onRename` receives it); it does not reload the document.
Document details also lists the file's core properties (title, author, created, modified…).
PDF goes through the browser print dialog ("Save as PDF"), so text stays selectable. `allowDownload={false}` also hides Download as / Print and blocks `Ctrl+P`.

```tsx
<DocViewer
  src={file}
  fileMenuItems={[{ label: 'Share…', icon: 'open', onClick: share }, { label: 'Version history', onClick: history }]}
  onRename={(name) => api.rename(id, name)}
  onClose={() => router.back()}
/>
```

CSS is auto-injected, no import needed (or import `@p10node/office-engine/styles.css` with `injectStyles={false}`). Container must have a height. Colors, type and shape: [Theming](theming.md).

| Prop               | Type                                                         | Default   | Notes                                                                                                                                                        |
|--------------------|--------------------------------------------------------------|-----------|--------------------------------------------------------------------------------------------------------------------------------------------------------------|
| `src`              | `File \| Blob \| ArrayBuffer \| Uint8Array \| string \| URL` | –         | Office (also 97-2003); OpenDocument; RTF; Pages/Keynote/Numbers; Visio; CSV/TSV, text, code, Markdown, HTML, diagrams                                        |
| `name`             | `string`                                                     | auto      | Display name; its extension also tells the format of a `Blob` / `ArrayBuffer`                                                                                |
| `language`         | `string`                                                     | –         | Source code in this language (`'typescript'`, `'ts'`, `'python'`, `'c++'`…): the code view, whatever the name says                                           |
| `toolbar`          | `boolean`                                                    | `true`    | Top bar (name, status, progress)                                                                                                                             |
| `toolbarExtra`     | `ReactNode`                                                  | –         | Extra buttons on the right of the toolbar                                                                                                                    |
| `zoom`             | `number \| 'fit'`                                            | `'fit'`   | Initial zoom (docx/pptx)                                                                                                                                     |
| `thumbnails`       | `boolean`                                                    | `true`    | Thumbnail column for pptx                                                                                                                                    |
| `headers`          | `Record<string,string>`                                      | –         | Headers for URL requests                                                                                                                                     |
| `credentials`      | `RequestCredentials`                                         | –         | Cross-origin cookies                                                                                                                                         |
| `worker`           | `() => Worker \| false`                                      | inline    | Custom worker (CSP), `false` = main thread                                                                                                                   |
| `styleTarget`      | `Document \| ShadowRoot`                                     | document  | Where CSS is injected                                                                                                                                        |
| `injectStyles`     | `boolean`                                                    | `true`    | Inject the stylesheet; false when the page loads `@p10node/office-engine/styles.css` itself                                                                  |
| `theme`            | `'light' \| 'dark' \| 'auto' \| ThemeOptions`                | `'light'` | Color scheme of the chrome; `'auto'` follows `prefers-color-scheme`; an object adds token overrides ([Theming](theming.md)). Documents keep their own colors |
| `locale`           | `string`                                                     | inherited | BCP 47 tag for number/date display and numeric input. See [UI texts](#ui-texts) and [input](#numbers-typed-in-a-locale)                                      |
| `strings`          | `StringsOverride`                                            | –         | Texts merged over the enclosing provider or English, key by key; may include a formatting `locale`                                                           |
| `editable`         | `boolean`                                                    | `false`   | Open in edit mode (Pages, Keynote, Numbers and Visio files stay view only)                                                                                   |
| `allowEdit`        | `boolean`                                                    | `true`    | Show Edit/View button                                                                                                                                        |
| `allowSave`        | `boolean`                                                    | `true`    | Show Save/Download button, enable Ctrl+S                                                                                                                     |
| `allowDownload`    | `boolean`                                                    | `true`    | Allow downloads to device (save without `onSave`, images)                                                                                                    |
| `fileMenu`         | `boolean`                                                    | `true`    | File menu: open, download as PDF / HTML / text, print                                                                                                        |
| `allowOpen`        | `boolean`                                                    | `true`    | File → New / Open (off during collaboration)                                                                                                                 |
| `captureSelectAll` | `boolean`                                                    | `true`    | Ctrl/⌘+A from anywhere on the page selects the document; `false`: from the viewer only                                                                       |
| `fileMenuItems`    | `FileMenuItem[]`                                             | –         | Extra File menu entries (`label`, `onClick`, `icon`, `hint`)                                                                                                 |
| `icons`            | `Icons`                                                      | built-in  | Your own icons: a map by `IconName` or a resolver `(name, { size }) => node`; see [Icons](#icons)                                                            |
| `components`       | `Partial<UIComponents>`                                      | built-in  | Replacements for toolbar and dialog controls. See [Theming](theming.md#your-own-components).                                                                 |
| `fonts`            | `FontSource[]`                                               | –         | Font files for the documents' text (the library ships none), loaded when a document uses them; see [Fonts](#fonts)                                           |
| `onRename`         | `(name) => void`                                             | –         | File → Rename result (extension included)                                                                                                                    |
| `onClose`          | `() => void`                                                 | –         | Adds File → Close                                                                                                                                            |
| `onSave`           | `(blob, name) => void`                                       | download  | Receives saved file (downloads by default): the Office file, or a Markdown / HTML / SVG / source file as itself                                              |
| `onDirtyChange`    | `(dirty) => void`                                            | –         | Whether there are unsaved changes                                                                                                                            |
| `collab`           | `CollabOptions`                                              | –         | Real-time collaboration, see [collaboration](collaboration.md)                                                                                               |
| `comments`         | `CommentOptions \| false`                                    | `{}`      | Comments (see below); `false` hides them                                                                                                                     |
| `onLoad`           | `(client) => void`                                           | –         | Once per opened document                                                                                                                                     |
| `onError`          | `(error) => void`                                            | –         |                                                                                                                                                              |
| `controller`       | `ViewerController`                                           | its own   | The viewer's state and commands for the host, see [headless mode](#headless-mode)                                                                            |
| `headless`         | `boolean`                                                    | `false`   | The document alone: no top bar, zoom control, outline panel, sheet tabs or slide list                                                                        |

### Headless mode

For a host that draws its own controls around the document: `headless` leaves the viewer with the document alone, and a
`controller` (`createViewerController()`) carries its state and the commands that drive it. While editing, the formatting
toolbar keeps its row above the document; dialogs, the find bar and the comments panel are the viewer's own. `toolbar`
set explicitly brings back the top bar (File menu, comments, Save) and leaves the rest headless.

```tsx
import { createViewerController, DocViewer, useViewerState } from '@p10node/office-engine'
import { useMemo } from 'react'

function Viewer({ file }: { file: File }) {
  const ctl = useMemo(() => createViewerController(), [])
  const s = useViewerState(ctl)
  return (
    <>
      <nav>
        {s.outline?.items.map((i) => (
          <button key={i.id} aria-current={i.current} style={{ paddingLeft: i.level * 12 }} onMouseDown={(e) => e.preventDefault()} onClick={() => ctl.goto(i.id)}>
            {i.title}
          </button>
        ))}
      </nav>
      {s.zoom && <button onClick={() => ctl.zoomIn()}>{s.zoom.value === 'fit' ? 'Fit' : `${Math.round(s.zoom.effective * 100)}%`}</button>}
      {s.canEdit && <button onClick={() => ctl.setEditing(!s.editing)}>{s.editing ? 'View' : 'Edit'}</button>}
      {s.dirty && <button onClick={() => ctl.save()}>Save</button>}
      <DocViewer src={file} controller={ctl} headless style={{ height: 600 }} />
    </>
  )
}
```

| `ViewerState`        | Meaning                                                                                                                                                                                                           |
|----------------------|-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| `rev`                | counts the changes (the state object is replaced on each)                                                                                                                                                         |
| `kind`               | `'docx'`, `'xlsx'`, `'pptx'`, or `null` before a document is open                                                                                                                                                 |
| `loading`            | loading or downloading                                                                                                                                                                                            |
| `editing`, `canEdit` | in edit mode; whether editing is possible (allowed, not view only)                                                                                                                                                |
| `dirty`, `saving`    | unsaved changes; a save running                                                                                                                                                                                   |
| `canUndo`, `canRedo` | whether undo / redo can run (spreadsheets and presentations count the edits made here; in a text of a slide both are offered while it is edited)                                                                  |
| `zoom`               | `{ value: 'fit' \| number, effective }`, `null` for spreadsheets                                                                                                                                                  |
| `page`               | `{ current, count }` from 1: the page (documents; `count` 0 while laid out) or slide in view; `null` for spreadsheets                                                                                             |
| `outline`            | `{ kind, items: [{ id, title, level, current }] }`: a document's headings, a workbook's visible sheets, a presentation's slides (titled by their title placeholder, else "Slide n")                               |
| `status`             | the status line ("Page 2 of 7 · 1,204 words", "Slide 3 of 12", "200 rows × 8 cols")                                                                                                                               |
| `format`             | while editing, the formatting at the selection (`TextFormat`, below), retained while using focused host controls; `null` without a valid target, while viewing, and in a presentation unless text is being edited |

Commands: `setEditing(on)`, `save()`, `undo()`, `redo()`, `zoomIn()`, `zoomOut()`, `setZoom('fit' | number)`,
`goto(outlineId)`, `gotoPage(n)`, `present()` (a slideshow from the slide in view), `print()`, `fileAction(id)` (a File menu
action: `pdf`, `html`, `text:md`, `details`, `rename`…), `toggleComments(open?)`, `format(command)` (below). A command the
open document has no use for does nothing. `subscribe(fn)` returns its unsubscribe; `snapshot()` the current state;
`useViewerState(ctl)` is the state in React. Without `react`, `mount()` returns the controller (`mounted.controller`), and
`<office-viewer headless>` has it as `element.controller`. Keep a subscription active to receive the complete DOCX
headings and PPTX slide titles.

#### Formatting

One formatting API for documents, cells and slide text, for a toolbar of your own:

```tsx
const f = useViewerState(ctl).format
<button aria-pressed={!!f?.bold} disabled={!f} onMouseDown={(e) => e.preventDefault()}
  onClick={() => ctl.format({ op: 'toggle', mark: 'bold' })}>B</button>
```

`TextFormat`: `bold`, `italic`, `underline`, `strike`, `font`, `size` (points), `color`, `highlight` (a cell's fill in a
spreadsheet), `align` (`'left' | 'center' | 'right' | 'justify'`), `list` (`'bullet' | 'number' | null`; undefined for
cells). `FormatCommand`: `{ op: 'toggle', mark }`, `{ op: 'font', name }`, `{ op: 'size', pt }`, `{ op: 'color', color }`,
`{ op: 'highlight', color }` (`null` removes either), `{ op: 'align', align }`, `{ op: 'list', list }`, `{ op: 'clear' }`.
Sizes must be finite and in **1–400 points**, rounded to the nearest half point, just like the built-in size box.
Invalid sizes do nothing before mutation, dirty state or history. Font names are trimmed and must be nonempty without control characters.

Color commands accept opaque **`#rgb`, `#rrggbb`, or `#rrggbbff`** (normalized to six hex digits); `null` removes the color.
Other CSS syntax, including names (`red`), functions (`rgb(...)`), and translucent colors, is rejected before mutation.
The snapshot may still contain other CSS colors read from the original file.

`list` commands **set every selected paragraph** to `'bullet'` or `'number'`, including mixed selections; `list: null`
explicitly removes lists. Repeating a setter is idempotent. The built-in toolbar buttons still toggle.
The snapshot reports the first selected paragraph's list, without aggregating mixed selections: `list: null` means no list;
`list: undefined` means lists are unsupported (cells). Other optional `TextFormat` fields may be undefined when unspecified
or unsupported; `format: null` means there is no formatting target. Unknown operations and invalid arguments do nothing.

Each mutating format command is a separate undo step (caret-only pending marks follow normal typing behavior). For a live
font/size field, keep a local draft and commit on blur or Enter when one undo step per finished value is desired; do not
publish every prefix of a font name. The controller also supports live commands while preserving native input focus.

Cells support marks, font, size, text color, fill, left/center/right alignment and clear; lists and justified alignment do
nothing. DOCX and PPTX support every command above. Tables, pictures, tracked changes and cell borders stay with the built-in
toolbars.

Focused host buttons, menu items, inputs, textareas and selects keep the logical editor target. Formatting preserves
the host control's focus, including keyboard activation and live `onChange` commands. Prevent mouse down on nonfocusable
menu options to avoid moving focus to the page background and losing the target; buttons may also prevent mouse down
to keep focus in the document. Native text fields retain their own text selection. A true selection loss, leaving edit mode, ending slide text editing, or replacing the source clears
the target. For example, a host size input can use `disabled={!f}` and send a size on change:

```tsx
<input aria-label="Font size" type="number" disabled={!f} defaultValue={f?.size}
  onChange={(e) => ctl.format({ op: 'size', pt: e.currentTarget.valueAsNumber })} />
```

Controller actions obey `allowSave`, `allowDownload` and `allowOpen` even when the File menu is hidden.
`fileMenu={false}` hides the menu and disables its Open and Print shortcuts. Ctrl/⌘+S and host controller commands remain available, subject to their own permissions.
`allowEdit` controls the Edit/View toggle and CSV import; `editable` and `setEditing()` control edit mode, including when the toggle is hidden.
Prevent the default mouse-down action on host outline buttons (as above) to preserve the editor focus and an open XLSX formula while switching sheets.
New/Open replace the document and remain available for a read-only file when `allowOpen` permits them.
Import requires `allowEdit` and an editable workbook; it remains available with `allowOpen={false}`. Controllers are created with
`createViewerController()`; publishing and attachment methods are internal to the engine.

Use one controller per mounted viewer. Give a modal or fullscreen viewer its own controller, even when it shows the same file.
During a temporary overlap, the latest mounted viewer owns a shared controller. Releasing it clears the controller;
an older viewer that remains mounted does not automatically regain ownership.

#### Ready-made controls

For a host that wants the controls but its own layout, each of these is driven by the controller alone and placed
anywhere around a headless viewer:

```tsx
import { DocViewer, FormatBar, OfficeUIProvider, OutlineList, SheetTabs, StatusText, useThemeRoot, ZoomBar } from '@p10node/office-engine'
import { de } from '@p10node/office-engine/locales/de'

const root = useThemeRoot(theme, undefined, { host: true })
<OfficeUIProvider icons={icons} components={components} strings={de}>
  <header className={root.className} style={root.style}>
    <ZoomBar controller={ctl} /> <StatusText controller={ctl} />
    <FormatBar controller={ctl} />
  </header>
  <aside className={root.className} style={root.style}><OutlineList controller={ctl} /><SheetTabs controller={ctl} /></aside>
  <DocViewer src={file} controller={ctl} headless theme={theme} />
</OfficeUIProvider>
```

| Control | Shows |
|---|---|
| `ZoomBar` | zoom out, the zoom (a click switches between fitting the width and 100 %), zoom in; nothing for spreadsheets |
| `StatusText` | the status line |
| `OutlineList` | headings (indented), sheets or slides; a click shows it, the one in view is marked |
| `SheetTabs` | a workbook's sheets as tabs; nothing for other documents |
| `FormatBar` | while editing: undo, redo, size, bold / italic / underline / strike, alignment, lists (where there are lists), clear formatting |

They render through the same components, icons, texts and design tokens as the built-in toolbars: outside `DocViewer`,
`OfficeUIProvider` supplies the components, icons and strings to both the controls and the viewer. An element with `useThemeRoot()`'s class and style supplies the tokens. Give a
container of controls `{ host: true }`: it is sized by them, and their popups (the size list, the overflow menu) show
beyond it, above the viewer.

### Comments

```tsx
<DocViewer
  src={file}
  editable
  comments={{
    user: { id: me.id, name: me.name },                              // author of the comments written here
    mentions: (q) => api.members(nodeId, q),                         // or a list: [{ id, name }, …]
    onEvent: (e) => e.mentioned.length && api.notify(e.mentioned, e), // notify people mentioned
    open: true,                                                      // panel open on load
  }}
/>
```

| `CommentOptions` | Type                                                      | Default                      |
|------------------|-----------------------------------------------------------|------------------------------|
| `user`           | `{ name, id?, initials? }`                                | the `collab` user, else "Me" |
| `mentions`       | `CommentUser[] \| (query) => CommentUser[] \| Promise<…>` | none (no `@` suggestions)    |
| `canComment`     | `boolean`                                                 | while editing                |
| `onEvent`        | `(e: CommentEvent) => void`                               | –                            |
| `open`           | `boolean`                                                 | `false`                      |
| `focus`          | `string` (a comment id)                                   | –                            |

`focus` shows a thread when the document opens (a link from a notification): the panel opens on it and the document scrolls to its place.

`onEvent` is called for the changes made in this viewer (not for collaborators' — each of them reports their own):

| `CommentEvent` | Meaning                                                                                     |
|----------------|---------------------------------------------------------------------------------------------|
| `type`         | `'add'`, `'reply'`, `'edit'`, `'delete'`, `'resolve'`, `'reopen'`                           |
| `comment`      | the comment (`DocComment`: `id`, `parent?`, `author`, `time`, `text`, `mentions?`, `done?`) |
| `thread`       | the thread's first comment                                                                  |
| `mentioned`    | people mentioned by this change: all of a new comment's, those an edit added                |
| `participants` | the people who wrote the thread's other comments (not the one who made the change)          |
| `quote`        | what the thread points at: the commented text, `Sheet1!B3`, `Slide 2`                       |

A comment's author and mentions carry the `id` the host gave (stored in the file's people / persons / authors list), so a host can tell its users apart after the file was saved and reopened. Ownership (edit / delete) is by `user.id` when there is one, else by name.

Use individual views / build your own UI:

```tsx
import { useDocument, useThemeRoot, DocxView, XlsxView, PptxView } from '@p10node/office-engine'
import '@p10node/office-engine/styles.css' // or call injectStyles() once: only DocViewer loads the stylesheet by itself

const { client, error, progress } = useDocument(src)
const root = useThemeRoot()
if (client?.kind === 'xlsx') return <div className={root.className} style={{ ...root.style, height: 600 }}><XlsxView client={client} /></div>
```

The views style themselves from design tokens on their root: `useThemeRoot(theme)` gives its class name and style (see [Theming](theming.md)). Without the stylesheet and that root they render unstyled, and like `DocViewer` they need a container with a height.

### Icons

Every icon the viewer draws, toolbars, menus, zoom, outline, Present and the add-sheet tab included, has a name (`IconName`, all of them in `ICON_NAMES`). `icons` replaces any of them, either as a map or as a resolver that is called for every icon:

```tsx
import { DocViewer, type IconName, type Icons } from '@p10node/office-engine'

// Define outside the host component so the map keeps its identity across renders.
const icons: Icons = { undo: <MyUndo />, redo: <MyRedo /> }

<DocViewer src={file} icons={icons} />
```

A name the map leaves out, or a resolver returning `undefined` / `null`, keeps the built-in icon, so a set can cover just part of the names. Your icon is rendered inside `<span class="de-icon">` whose font-size is the icon's size in pixels: draw it at `1em` (a bare `<svg>` child is sized to `1em` for you) and with `currentColor`, so it follows the theme. The resolver also gets `{ size }` in pixels.

The names exported in `ICON_NAMES` are stable public keys, including names originally used only inside the viewer.
New names may be added; existing keys must not be renamed as an internal refactor.

With [Lucide](https://lucide.dev) (`lucide-react`):

```tsx
import { Bold, Italic, ListTree, Plus, Presentation, Redo2, Replace, Search, Undo2, X, ZoomIn, ZoomOut, type LucideIcon } from 'lucide-react'

const LUCIDE: Partial<Record<IconName, LucideIcon>> = {
  undo: Undo2, redo: Redo2, bold: Bold, italic: Italic, search: Search, replace: Replace,
  zoomIn: ZoomIn, zoomOut: ZoomOut, outline: ListTree, close: X, present: Presentation, sheetAdd: Plus,
}
const lucide = (name: IconName) => {
  const C = LUCIDE[name]
  return C && <C size="1em" strokeWidth={1.8} />
}

<DocViewer src={file} icons={lucide} />
```

With [Phosphor](https://phosphoricons.com) (`@phosphor-icons/react`):

```tsx
import { ArrowArcLeft, ArrowArcRight, ArrowsLeftRight, ListBullets, MagnifyingGlass, MagnifyingGlassMinus, MagnifyingGlassPlus, Plus, PresentationChart, TextB, TextItalic, X, type Icon } from '@phosphor-icons/react'

const PHOSPHOR: Partial<Record<IconName, Icon>> = {
  undo: ArrowArcLeft, redo: ArrowArcRight, bold: TextB, italic: TextItalic, search: MagnifyingGlass, replace: ArrowsLeftRight,
  zoomIn: MagnifyingGlassPlus, zoomOut: MagnifyingGlassMinus, outline: ListBullets, close: X, present: PresentationChart, sheetAdd: Plus,
}
const phosphor = (name: IconName) => {
  const C = PHOSPHOR[name]
  return C && <C size="1em" />
}
```

The standalone views get the same icons from `IconsProvider`, which `DocViewer` uses internally; a nested provider falls back to the enclosing one for names it leaves out:

```tsx
import { IconsProvider, XlsxView } from '@p10node/office-engine'

<IconsProvider icons={lucide}>
  <div className="de-root"><XlsxView client={client} /></div>
</IconsProvider>
```

### UI texts

Viewer UI texts (titles, labels, menu entries, placeholders, dialogs, toasts, status lines such as "Page 1 of 5") come from one typed table, `Strings`, with the exceptions listed below. Only English UI texts with US spelling are included in the default viewer and standalone embed bundle. German is a separate optional import with standard German spelling and `de-DE` number/date formatting (for example `1.234.567,89`):

```tsx
import { DocViewer } from '@p10node/office-engine'
import { de } from '@p10node/office-engine/locales/de'

<DocViewer src={file} strings={de} />
```

The locale module is included only when your app imports it. The German table uses standard German spelling and `de-DE` formatting and numeric input. `locale="de-DE"` alone changes number/date formatting and numeric input while keeping English UI texts; it does not load German. With no table or formatting override, dates and grouped counts retain the browser's locale. Page/slide numbers, measurements, file sizes and zoom labels retain ASCII digits and a decimal point, as before. An explicit formatting tag applies to all of these displays. The English labels now use singular forms for counts of one. Sheet counts and paste limits use the active formatting locale or the browser default. Zoom percentages use Intl rounding, which can differ at exact halves.

`strings` merges over the enclosing `StringsProvider` (or English) key by key, so a group can be given in part. Nested providers and `DocViewer` inherit the outer texts and formatting. Undefined values do not replace inherited values: `strings={en}` changes the texts while retaining an enclosing formatting tag; also pass `locale="en-US"` to select US formatting explicitly. A value is a string, or a function where the text takes numbers or names. Count callbacks receive raw numbers for pluralization; callbacks that format grouped counts also receive the active formatting tag as their final optional argument. Hoist override objects outside render, or memoize them when they depend on props, to keep the provider value stable. To customize an imported language, spread that table and any groups you change:

```tsx
import { DocViewer } from '@p10node/office-engine'
import { de } from '@p10node/office-engine/locales/de'

const germanTexts = {
  ...de,
  file: { ...de.file, menu: 'Ablage' },
  docx: {
    ...de.docx,
    status: (page: number, pages: number, words: number, locale?: string) =>
      `${page.toLocaleString(locale)}/${pages.toLocaleString(locale)} · ${words.toLocaleString(locale)} ${words === 1 ? 'Wort' : 'Wörter'}`,
  },
}

<DocViewer src={file} strings={germanTexts} />
```

Views used on their own get their texts from a `StringsProvider` (English without one); `useStrings()` reads them in your own components:

```tsx
import { StringsProvider, useStrings, XlsxView } from '@p10node/office-engine'
import { de } from '@p10node/office-engine/locales/de'

const sheetTexts = { ...de, xlsx: { ...de.xlsx, addSheet: 'Neues Blatt' } }

<StringsProvider strings={sheetTexts}>
  <div className="de-root"><XlsxView client={client} /></div>
</StringsProvider>
```

For a regional variant, override only the texts that differ. The `locale` prop is a BCP 47 formatting tag and takes precedence over `strings.locale`; otherwise the local table’s locale takes precedence over the enclosing provider. With neither, displays use the defaults above. Malformed tags (for example `de_DE`) are ignored: a valid table or inherited tag still applies, otherwise the default formatting above is used. For display, well-formed but unsupported tags follow `Intl` fallback. Numeric input instead uses the engine’s default syntax when the resolved tag is missing or unsupported. For example, a host can use US English as the base and supply British formatting and selected British spellings:

```tsx
import { DocViewer, StringsProvider, en, type StringsOverride } from '@p10node/office-engine'

const usEnglish = { ...en, locale: 'en-US' }
const britishOverrides: StringsOverride = {
  edit: { textColor: 'Text colour' },
  format: { borderColor: 'Border colour' },
}

<StringsProvider strings={usEnglish}>
  <DocViewer src={file} locale="en-GB" strings={britishOverrides} />
</StringsProvider>
```

The same overrides work with `StringsProvider`, `mount()` and the web component's `strings` property. The example changes only those two labels; other text keeps the US English defaults. Override any other regional differences you need.

For a complete language of your own, pass a `Strings` table as `strings`. Starting from `en` (or `de`) keeps untranslated entries as fallbacks; the following example only translates Cancel. Replace the remaining entries to complete the translation:

```ts
import { en, type Strings } from '@p10node/office-engine'

export const fr: Strings = { ...en, locale: 'fr-FR', common: { ...en.common, cancel: 'Annuler' } }
// <DocViewer src={file} strings={fr} />
```

A host can keep its own language tables in its app and pass them through `strings`; no engine registration is needed. A new optional package locale follows the German entry: add its typed table, a `src/locales/<code>.ts` entry and the corresponding build/package export. Locale parity tests check complete keys and nonempty texts.

For vanilla embeds, pass the same optional table to `mount(el, { src, strings: de })` or assign it to the web component:

```html
<script src="office-engine.embed.js"></script>
<office-viewer id="viewer" src="/file.xlsx"></office-viewer>
<script type="module">
  import { de } from './locales/de.js' // optional module next to the distributed embed bundle
  document.getElementById('viewer').strings = de
</script>
```

Document content stays in its original language. Default new sheet names, the author "Me" and new text box content also stay unchanged; the “Page 1 of 5” header/footer preset inserts text in the UI language. Font names, formula names and help, number format previews, host-provided labels and engine/transport errors (including main-thread fetch errors such as `HTTP 404 while fetching …`, collaboration errors such as `Document is still syncing`, and parsing-worker messages) are not translated by the table. Print, PDF and HTML export inherit the active UI table for viewer-generated labels; document content remains unchanged.

### Numbers typed in a locale

`locale` (a BCP 47 tag) says how the numbers a person types are read: into a cell of a sheet, the font size box, the margins and column spacing of Page setup, the size in Header & footer. Its decimal and grouping separators, minus and percent signs and digits come from `Intl.NumberFormat`:

```tsx
<DocViewer src={file} editable locale="de-DE" />   // 1,5 → 1.5; 1.234,5 → 1234.5; 50 % → 0.5; 1.5 stays text
<DocViewer src={file} editable locale="fr-FR" />   // 1 234,5 (any space) → 1234.5
<DocViewer src={file} editable />                  // as before: 1.5 and 1,234.5; 1,5 stays text
```

- With no supported tag from the prop, local table or enclosing provider, numeric input uses a decimal point and commas grouping by thousands. The browser’s language does not affect numeric input.
- The whole text is read as one number or not at all: `12,5` without a locale is not 12, `1.5` with `de-DE` is not 15. A cell keeps what is not a number as text; a field keeps its previous value (the font size box shows it again, Page setup disables Apply). The rules, grouping included: [Features](features.md#numbers-typed-in-a-locale).
- What is saved is the number (`<v>1.5</v>`), the same in every locale; changing `locale` changes no value of a workbook. A number edited again shows the locale's decimal separator (`1,5` in the formula bar and the cell), which reads back as the same number.
- Formulas are not localized: `=1.5+A1`, `=SUM(1,5)` and the English function names, whatever the locale.
- With `locale`, the Page setup and Header & footer fields are text fields read by the same rules; without it they stay the browser's number fields.
- Collaboration: the tag goes with each cell edit (`locale` of the shared `set` operation), so every participant, and a server replaying the log with `compact`, reads `1,5` the same. Participants should run a version that knows it: an older one reads such an edit with a decimal point.

For standalone views or controls, wrap them in `<StringsProvider locale="de-DE">` or supply a table with that locale. The worker also accepts the tag directly in `client.call('xlsx.set', sheet, cells, { locale: 'de-DE' })`; `mount(el, { src, locale: 'de-DE' })`; `<office-viewer locale="de-DE">`.

### Fonts

The library ships no font files. Text is shown in the fonts the browser has, and where a line or a page breaks depends on them: a document in Calibri on a machine without Calibri is set in another font and breaks elsewhere. Documents name metric-compatible fallbacks after their fonts (the same widths, under the SIL Open Font License), so supplying those files is enough for most Office documents to break as they do in Office:

| Document font          | Fallback                 |
|------------------------|--------------------------|
| Calibri, Calibri Light | Carlito                  |
| Cambria                | Caladea                  |
| Arial, Helvetica       | Liberation Sans, Arimo   |
| Times New Roman        | Liberation Serif, Tinos  |
| Courier New            | Liberation Mono, Cousine |

```tsx
import { DocViewer, type FontSource } from '@p10node/office-engine'

// Served by your app (Carlito and Caladea: Google Fonts / Debian fonts-crosextra-*; Liberation: github.com/liberationfonts)
const fonts: FontSource[] = [
  { family: 'Carlito', src: '/fonts/Carlito-Regular.ttf' },
  { family: 'Carlito', src: '/fonts/Carlito-Bold.ttf', weight: 700 },
  { family: 'Carlito', src: '/fonts/Carlito-Italic.ttf', style: 'italic' },
  { family: 'Carlito', src: '/fonts/Carlito-BoldItalic.ttf', weight: 700, style: 'italic' },
  { family: 'Caladea', src: '/fonts/Caladea-Regular.ttf' },
  { family: 'Liberation Sans', src: '/fonts/LiberationSans-Regular.ttf' },
  { family: 'Liberation Serif', src: '/fonts/LiberationSerif-Regular.ttf' },
  { family: 'Liberation Mono', src: '/fonts/LiberationMono-Regular.ttf' },
]

<DocViewer src={file} fonts={fonts} />
```

| `FontSource`   | Type                                               | Notes                                                                                 |
|----------------|----------------------------------------------------|---------------------------------------------------------------------------------------|
| `family`       | `string`                                           | The name text asks for: a fallback (`Carlito`) or the document's own name (`Calibri`) |
| `src`          | `string \| ArrayBuffer \| ArrayBufferView \| Blob` | A URL (resolved against the page) or the file's bytes: TTF, OTF, WOFF, WOFF2          |
| `weight`       | `number \| string`                                 | CSS `font-weight` of the file: `700`, `'bold'`, `'100 900'` for a variable font       |
| `style`        | `string`                                           | `'normal'`, `'italic'`, `'oblique'`                                                   |
| `unicodeRange` | `string`                                           | CSS `unicode-range`, when a family comes in several files                             |

- **Which name**: under a fallback's name (`Carlito`), the file is used where the original is not installed, and a machine that has Calibri keeps it. Under the document's own name (`Calibri`), it is used everywhere, before an installed font of that name.
- **When**: a file is fetched when a document whose text is set in its family is shown (through the fallback list too: a declared `Carlito` is fetched for a document in Calibri), not before. Nothing is fetched when no font is declared. The faces are added to the page's `document.fonts` (the `FontFace` API, no stylesheet) and shared by the viewers on the page that declare the same file; they are removed when no viewer declares them any more.
- **Layout**: a document is measured and paginated once the fonts it uses are loaded. A file that fails to load is reported to `onError` and in the console, and the text is shown in the next font of its list; one that takes more than 3 s is not waited for any longer: the document is laid out without it, and again when it arrives (the page in view and the caret stay where they are). A font the page provides itself (an `@font-face` of the host page) that arrives late also has the document laid out again. Spreadsheets fit numbers to their columns again; slides are laid out by the browser and follow by themselves. Printing and *Download as PDF* use the same files.
- **Only from the host**: fonts are taken from this list only. A document's font names pick among them, and fonts embedded in a document (`word/fontTable.xml`) are not read.

## Without a React build

```html
<script src="office-engine.embed.js"></script>   <!-- dist/, bundles React + worker -->

<office-viewer src="/files/report.docx" style="height:600px"></office-viewer>

<input id="pick" type="file">
<div id="v" style="height:600px"></div>
<script>
  const viewer = OfficeEngine.mount(document.getElementById('v'), { src: '/files/deck.pptx' })
  document.getElementById('pick').onchange = (e) => viewer.update({ src: e.target.files[0] })
  // viewer.destroy()
</script>
```

`<office-viewer>`: attributes `headless`, `src`, `name`, `language` (a source file's, as the prop), `editable`, `toolbar="false"`, `thumbnails="false"`, `save="false"`, `download="false"`, `file-menu="false"`, `open="false"`, `capture-select-all="false"`, `theme="dark"` (or `"auto"`), `locale="de-DE"` (number/date formatting and numeric input), `collab` / `room` / `user` / `base` (real-time collaboration; `user` also writes the comments), `comments="false"`, `fonts` (a JSON list of `FontSource` with URLs: `fonts='[{"family":"Carlito","src":"/fonts/Carlito-Regular.ttf"}]'`); properties `.file = File`, `.icons` (map or resolver, as `icons` above) `.strings = { … }` (see [UI texts](#ui-texts)) and `.fonts` (as `fonts` above, bytes too; it wins over the attribute), and `.controller` (see [headless mode](#headless-mode)); events `load`, `error`, `comment` (`detail`: a `CommentEvent`).

## Core API (no React)

```ts
import { DocumentClient } from '@p10node/office-engine/core'

const doc = await DocumentClient.open(fileOrUrl, { onProgress: console.log })
doc.kind                                   // 'docx' | 'xlsx' | 'pptx'
doc.info.from                              // 'pages' / 'keynote' / 'numbers' / 'visio': an Apple or Visio file, converted to docx / pptx / xlsx / pptx (view only)
doc.info.imported                          // 'csv' / 'tsv' / 'text' / 'markdown' / 'html' / 'svg' / 'mermaid' / 'plantuml' / 'graphviz' / 'code': a plain file, built as xlsx / docx (editable)
doc.info.language                          // 'code' (a source file): its language, 'typescript' / 'python' / 'json'… (the hint's own word when it is not one the highlighter knows)
doc.info.converted                         // 'doc' / 'xls' / 'ppt' / 'rtf' / 'odt' / 'ods' / 'odp': converted to docx / xlsx / pptx (editable, saved as such)
doc.info.ext                               // 'xlsm' / 'xltx' / 'docm' / 'dotx' / 'pptm' / 'potx' / 'ppsx' …: the file is that variant, and is saved as it (undefined: plain docx / xlsx / pptx)
doc.on('docx:blocks', (blocks) => {})      // docx: streamed blocks
await doc.call('docx.start')
doc.call('xlsx.load', 0)                   // xlsx: stream sheet
doc.on('xlsx:progress', (p) => {})
await doc.call('xlsx.block', 0, 1000, 1100) // formatted cells, rows [1000, 1100)
await doc.call('pptx.slide', 12)           // slide model with layout/master/theme resolved
await doc.call('xlsx.set', 0, [{ r: 4, c: 1, input: '=SUM(A1:A4)' }])
await doc.call('xlsx.set', 0, [{ r: 5, c: 1, input: '1.234,5' }], { locale: 'de-DE' }) // numbers typed in a locale: 1234.5
const clip = await doc.call('xlsx.copy', 0, { r0: 0, c0: 0, r1: 9, c1: 2 })   // cells as they are (values, formats, formulas) + `text` for the clipboard
await doc.call('xlsx.sheetOp', 0, { op: 'paste', range: { r0: 20, c0: 0, r1: 29, c1: 2 }, cells: clip.cells, from: { sheet: clip.sheet, r0: 0, c0: 0 } })
await doc.call('xlsx.sheetOp', 0, { op: 'clear', range: { r0: 0, c0: 3, r1: 1048575, c1: 3 } })   // a whole column, every row of it
await doc.call('xlsx.import', 0, csvFile, { name: 'cities' }) // a new sheet after sheet 0 with the rows of a CSV / TSV, streamed in when loaded
await doc.call('xlsx.text', 0, ',') // the sheet as CSV (a Blob; '\t' for TSV), the values as shown
// open(src, { residentCells: 12_000_000 }): cells of a CSV / TSV kept in memory; the other rows are read from the file as they are shown
// open(src, { encoding: 'windows-1258' }): the encoding of a text file (CSV, TSV, text, Markdown, HTML…), a TextDecoder label; without it the
// byte order mark, the declared charset, UTF-8, then (Vietnamese text) Windows-1258, then Windows-1252 decide. An unknown label is an error
// open(src, { language: 'python' }): a text file is source code in this language (as the `language` prop)
// the worker also emits `warning` ({ text }) when a file opened but not whole: a CSV with more than 16,384 columns (the viewer shows it as a message)
const comments = await doc.call('comments.list')  // DocComment[]: threads are roots + replies (threadsOf)
doc.on('comment', ({ id, c }) => {})       // a comment set / deleted (c null), here or by a collaborator
await doc.call('comments.set', c.id, { ...c, done: true }) // XLSX / PPTX threads carry their place (`at`)
const blob = await doc.call('save')        // edited file
await doc.call('plain.source')             // text of a Markdown / HTML / diagram / source file (null for other files)
await doc.call('plain.original')           // a source file's bytes as opened (null for other files)
doc.destroy()

// syntax colors, DOM-free (the tokenizer of the code view, the Markdown code blocks and @p10node/markdown)
import { highlight, highlightLines, languageOf } from '@p10node/office-engine/core'
highlight('ts', 'const a = 1') // [['k', 'const'], ['', ' a = '], ['n', '1']]: c comment, k keyword, s string, n number / constant,
                               // f function, y type, a attribute / property, t tag, m meta, o operator ('' plain)
languageOf('c++')              // 'cpp' (undefined: a language it does not color)
```

### Reading font files

The same entry reads font files without a browser: glyph widths, kerning, ligatures and line metrics come from the file's own tables, so a text is as wide in every browser, in a worker and on a server. Nothing is drawn and no DOM is used. This is what a layout with the same line breaks everywhere, and PDF output, are built on; the viewer itself still leaves text layout to the browser.

```ts
import { DocumentClient, FontBook, Font, layoutText, subsetFont, toUnicodeCMap, cssFont } from '@p10node/office-engine/core'

// the list DocViewer's `fonts` takes; files are fetched and read when text first asks for them
const fonts = [
  { family: 'Carlito', src: '/fonts/Carlito-Regular.ttf' },
  { family: 'Carlito', src: '/fonts/Carlito-Bold.ttf', weight: 700 },
]

// in the worker of a document (or of none: `new Engine(emit)` on a server)
await doc.call('fonts.set', fonts)
const [m] = await doc.call('fonts.measure', [{ text: 'Total: 1,234.50', font: cssFont('Calibri')!, size: 11, weight: 700 }])
m.width                                  // in the unit of `size`
m.ascent; m.descent; m.lineGap           // of the first font of the list that has a file
m.exact                                  // true: what a browser sets for this text in these files
m.unsure                                 // why not: see below
doc.on('font:error', ({ family, message }) => {})   // a file that could not be fetched or read

// the same without a worker; `pieces` then carry the glyphs
const book = new FontBook(fonts, { fetch: (url) => readFile(url) })
const measured = await book.measure({ text: 'office', font: 'Carlito', size: 11 })
measured.pieces[0].run                   // { glyphs, advances, dx, dy, clusters, width, exact… } in font units

// one file
const font = await Font.open(bytes)      // TTF, OTF, WOFF, WOFF2; Font.open(bytes, 1): the second font of a TTC
font.unitsPerEm; font.names.family; font.weight; font.italic; font.metrics
font.lineMetrics()                       // { ascent, descent, lineGap } as browsers take `line-height: normal`
font.lineMetrics('word')                 // as Word spaces single lines (the OS/2 Windows values)
const run = layoutText(font, 'AV office', { kerning: true, ligatures: true })

// a subset to embed, and the text of its glyphs for a PDF
const sub = subsetFont(font, run.glyphs)  // { bytes (a TTF / OTF of its own), glyphs: Map(old → new), cff? }
toUnicodeCMap(run.glyphs.map((g, i) => [sub.glyphs.get(g)!, 'AV office'.slice(run.clusters[i], run.clusters[i + 1])]))
```

- **Which file**: as a browser chooses among `@font-face` rules. The families of the font list in order (`cssFont('Calibri')` names Carlito after Calibri), in a family the nearest style and then the nearest weight, `unicodeRange` respected, and each character in the first file that has a glyph for it.
- **What is applied**: the character map, glyph widths, pair kerning (GPOS pair adjustments, else the `kern` table), single, multiple and ligature substitutions of the features browsers apply to plain text (`liga`, `clig`, `ccmp`, `locl`, `rlig`, `calt`, `rclt`). Text is cut where its script changes, as browsers do before shaping. `kerning: false` and `ligatures: false` turn the two off, as Word has them unless a document asks; `letterSpacing` turns ligatures off, as in CSS.
- **Never exact by accident**: `exact` is false, with the reason in `unsure` (`inexact` on a run), wherever the result is not what a browser's shaping gives:

| Reason      | Meaning                                                                                               |
| ----------- | ----------------------------------------------------------------------------------------------------- |
| `font`      | no declared file serves the text: its width is 0, to be measured some other way                       |
| `missing`   | a character no file of the list has a glyph for                                                       |
| `shaping`   | text that needs a shaping engine (Arabic, Hebrew, Indic scripts, Thai, combining marks, emoji sequences) and no shaper |
| `features`  | the font has a rule for these glyphs that is not applied (contextual forms, mark placement, Apple's `morx` / `kerx` / `trak` tables) |
| `kerning`   | the font's kerning is in a table of a kind that is not read                                           |
| `synthetic` | bold or italic text served by a regular file: the browser thickens or slants it itself               |
| `variable`  | a variable font asked for at another weight than its tables give widths for                           |

- **Shaping**: the library ships no shaping engine. A host can pass one, for example HarfBuzz compiled to WebAssembly: `new FontBook(fonts, { shaper })`, `layoutText(font, text, { shaper })`, or `engine.shaper = shaper` before `fonts.set` on an engine that runs where its host does (a function does not cross to a worker). It is called for the text that needs it, with the font as a bare TTF / OTF (`font.sfnt()`):

```ts
const shaper: Shaper = {
  // text of one script and direction → glyphs in the order they are drawn, lengths in font units
  shape: (fontBytes, text, { script, direction, features }) => [{ glyph, advance, dx, dy, cluster }],
}
```

- **Subsets**: `subsetFont` keeps the glyphs given and what they are made of, renumbered from 0 (glyph 0 stays), with a character map for the `chars` passed as a third argument. TrueType outlines are copied; PostScript ones (CFF, also CID-keyed) keep their charstrings and subroutines, and come as an OTF and as the bare CFF table (`cff`) a PDF embeds. `font.embeddable` is false for a font whose license forbids embedding.
- **Limits**: `FONT_LIMITS` (a file of 64 MB, 128 MB unpacked, 2,000,000 layout records). A file that is not a font, is damaged or is past a limit throws a `FontError` that says which. Not read: fonts embedded in a document, CFF2 outlines, the other instances of a variable font, bitmap and color glyphs. WOFF2 needs a runtime whose `DecompressionStream` unpacks Brotli: Node 26, Bun 1.4 and Firefox 157 do, Chromium 156 does not. There a WOFF2 file is refused with a `FontError` that says so (`font:error`, `onError`) and its text goes to the next font of the list, so declare TTF, OTF or WOFF files for text that has to be measured in Chrome.

### PDF output

A presentation is written as a PDF file by the library itself: no browser, no print dialog, no dependency, in the worker or on a server (`new Engine(emit)`). Its text is set in the font files of `fonts.set` and embedded as subsets, with the text of their glyphs, so it can be selected, copied and searched. Spreadsheets (#47) and documents (#48) follow.

```ts
await doc.call('fonts.set', fonts)                     // the files text is set in (see "Reading font files")
const pdf = await doc.call('pdf')                      // a Blob, application/pdf: every slide but the hidden ones
await doc.call('pdf', { slides: [0, 2], title: 'Q3' }) // these slides; hidden ones only with `hidden: true`, even when named
doc.on('warning', ({ text }) => {})                    // what could not be drawn as it is, once each
const svg = await doc.call('pptx.svg', 0)              // a slide as an SVG picture, drawn as the PDF draws it

// without a document: pages of drawing commands, and the backends that paint them
import { slidePage, writePdf, pageToSvg, svgItems, FontBook } from '@p10node/office-engine/core'
const page = await slidePage(slideModel, { fonts: new FontBook(fonts), part: async (name) => bytesOf(name) })
const bytes = await writePdf([page], { title: 'Deck', onWarning: console.warn })
```

- **Pages**: a slide is a page at the slide's size; shapes, preset and custom outlines, solid and gradient fills, picture fills, outlines with their dashes and arrowheads, flips and rotation, pictures with their crop and shape, tables (rows grow to their text), charts (the viewer's chart, as shapes and text) and text boxes are placed as the viewer places them. One page of the outline per slide title.
- **Text**: lines broken, aligned and spaced as the viewer's CSS does (`pre-wrap`, a unitless `line-height`), bullets in their hanging indent, underline, strikethrough, highlight, superscript and subscript, links, vertical text. Text no declared file serves is set in Helvetica (the PDF's own font, Windows-1252 characters only) and a warning names the font, so a server should declare its font files: Carlito, Caladea and Liberation stand in for Calibri, Cambria, Arial and Times New Roman.
- **Fonts' licenses**: a font that may not be embedded, or only as bitmaps, has its text set in Helvetica with a warning; one that may not be subset is embedded whole (but a CJK font keyed by CID, whose glyphs are not its characters, is set in Helvetica too). A damaged font file falls back the same way and the rest of the file is written.
- **Pictures**: JPEG copied as it is, PNG and GIF with their transparency, SVG as shapes and text. EMF, WMF, TIFF, BMP and WebP are left out with a warning. A picture is unpacked to at most 64 million pixels, an SVG picture draws at most 100,000 elements and 1,000,000 path segments (counted as drawn, references included); one past them, or one that cannot be read, is left out with a warning.
- **Not yet**: shadows and other effects, pattern fills (drawn in their foreground color), tiled pictures (stretched), text columns, autofit that shrinks text (the file's own scale is used) or grows the shape to its text, notes pages, SmartArt and OLE objects.

## Standalone packages

Diagrams and Markdown without the viewer (no DOM needed, no dependency): [`@p10node/diagram`](../packages/diagram), [`@p10node/mermaid`](../packages/mermaid), [`@p10node/plantuml`](../packages/plantuml), [`@p10node/graphviz`](../packages/graphviz), [`@p10node/markdown`](../packages/markdown).

```ts
import { render, renderPage } from '@p10node/diagram'
const { svg, width, height } = render('sequenceDiagram\n  Ann->>Ben: hi') // language detected
renderPage() // <pre class="mermaid">, <code class="language-dot">… of the page become drawings

import { render as markdown } from '@p10node/markdown'
el.innerHTML = markdown(text) // safe HTML; ```mermaid / ```plantuml / ```dot drawn, $…$ / $$…$$ math as MathML
el.innerHTML = markdown(text, { highlight: true }) // and code blocks colored: <span class="hl-keyword">…
```

## CSP

The default worker is an inline blob → requires `worker-src blob:`. If that's not allowed, host `dist/office-engine.worker.js` (exported as `@p10node/office-engine/worker`) on the same origin:

```tsx
<DocViewer src={src} worker={() => new Worker('/static/office-engine.worker.js', { type: 'module' })} />
```

## Links in documents

Links from a document (DOCX text and pictures, PPTX text) are clickable in view mode when they point to `http(s)`, `mailto`, `tel`, `sms`, `ftp(s)`, a relative path or a place in the document (`#bookmark`). Other targets (`javascript:`, `data:`, `vbscript:`, `file:`…) show as plain text, also in the web page and Markdown exports, so a crafted file cannot run script in the host page. The file itself keeps its links unchanged. External links open in a new tab with `rel="noreferrer noopener"`.
