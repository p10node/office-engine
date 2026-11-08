# Features

## Loading and rendering

- **Lazy ZIP**: reads only the central directory, inflates each part on demand (native `DecompressionStream`, 0 dependencies)
- **HTTP Range**: open a URL → fetch only the bytes needed (256 KB block cache); server without Range support → automatic full-download fallback
- **Streaming XML parser** runs in a **Web Worker** → UI never blocks
- **Progressive render**:
  - DOCX: blocks (paragraph/table) stream in → offscreen height measurement in time slices → pagination → only visible pages render
  - XLSX: rows stream in, 2D virtualized grid (only cells in viewport render), freeze panes, merge, number format, sheet tabs
  - PPTX: slides parsed on demand when scrolled into view, virtualized thumbnails

## Format support

| Format | Supported                                                                                                                                                                         | Not yet                                                                    |
|--------|-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|----------------------------------------------------------------------------|
| DOCX   | inherited styles, run formatting, lists/numbering, tables (merge, border, shading, table style), images, hyperlinks, header/footer, PAGE field, text boxes, page/section breaks, text columns, tracked changes | footnote body, charts, exact float wrapping, paragraphs split across pages |
| XLSX   | shared strings, number/date formats, font/fill/border/align, merge, freeze, hidden columns/rows, wrap, formula calculation (134 functions, recalculated on edit), multiple sheets | charts, conditional formatting, images                                     |
| PPTX   | master/layout/theme, placeholder inheritance, ~50 preset shapes, custGeom, images (crop), tables + table style, groups, bullet/autonum, autofit                                   | charts (placeholder), SmartArt, animations, video playback                 |

`.doc/.xls/.ppt` (legacy binary) and encrypted files are not supported.

### Apple Pages, Keynote and Numbers (view only)

Files from Pages 5, Keynote 6 and Numbers 3 (2013) and later open read-only: the worker converts a `.pages` document to DOCX, a `.key` presentation to PPTX and a `.numbers` spreadsheet to XLSX, in memory, and the matching view shows it. Find, print, word count, slideshow, sheet tabs and *Download as* work as usual; *Download as* **Word** / **PowerPoint** / **Excel** saves the converted file. The bar shows a `PAGES` / `KEYNOTE` / `NUMBERS` badge and no Edit button, even with `editable`. Single files and zipped package folders (`Report.pages/…` inside a ZIP) both open.

**Pages**

- **Converted**: paragraph and character styles (named styles become Word styles), fonts, sizes, bold/italic/underline/strikethrough, colors, highlights, super/subscript, caps, alignment, indents, spacing, page breaks, bullet and numbered lists (levels), links, inline images, text boxes and shapes with a fill, tables (merged cells, header rows, cell fills, banded rows, lines, text/number/date cells, rich text cells), header and footer (left, center and right parts, page number), footnotes (listed after the text), page size and margins
- **Not yet**: charts, equations, comments, tracked changes, exact positions of floating objects (placed in the text flow: those on the first page before the body, the others after it; on page layout documents, page by page), image cropping, several sections with different headers

**Keynote**

- **Converted**: slide size and order, skipped slides (hidden), master slide objects (drawn under the slide's), slide backgrounds (color, gradient, image), title / body / slide number placeholders, text boxes (text styles, bullet and numbered lists, links), shapes (rectangles, rounded rectangles, polygons, arrows, stars, custom paths and curves, fill and line), images cropped by their mask, movie poster frames, groups, rotation and flips, tables (merged cells, header rows, cell fills, banded rows, lines)
- **Not yet**: charts (a placeholder box), builds and transitions, speaker notes, shadows and reflections, text that shrinks to fit

**Numbers**

- **Converted**: every table becomes a worksheet (named after its sheet, plus the table's name when a sheet holds several), hidden sheets stay hidden; values keep their type: numbers, currencies, percentages, scientific and dates with the matching Excel number format (Numbers date patterns translated), booleans; durations, fractions, numbers in other bases, checkboxes, star ratings and pop-up menu choices as the text Numbers shows; rich text cells as their text; cell look (font, size, bold/italic/underline/strikethrough, color, alignment, vertical alignment, wrap, fill, banded rows, lines), column widths, row heights (grown to fit wrapped text), merged cells, hidden rows / columns, frozen header rows / columns
- **Not yet**: formulas (cells hold the values Numbers computed), charts, images and text boxes placed on sheets, custom formats (shown with the General format), accounting-style currency, the document locale's decimal / thousands separators, conditional highlighting, comments

Cell values in Pages and Keynote tables use the same formats.

**Not supported**: password-protected files and the iWork '09 and older XML format: each gives an error saying so.

### CSV, text, Markdown and HTML

A `src` (file, Blob, buffer or URL) that is not an Office package opens too, converted in the worker and **editable**:

| File                      | Opens as                                                                                                |
|---------------------------|---------------------------------------------------------------------------------------------------------|
| `.csv`, `.tsv` / `.tab`   | workbook, one sheet named after the file; numbers and booleans typed, codes like `007` kept as text     |
| `.txt`, `.text`, `.log`   | document, one paragraph per line                                                                        |
| `.md`, `.markdown`        | document: headings, lists, bold / italic / strike, links, code, quotes, tables (as tab-separated lines) |
| `.html`, `.htm`, `.xhtml` | document: headings, paragraphs, lists, inline formatting and links; scripts and styles are dropped      |

- **Format**: the extension (of `name`, the File or the URL), else the MIME type (Blob type, `Content-Type`), else the content (HTML, delimited rows, or text). A file named `.docx` / `.xlsx` / … is always read as a package, so an HTML error page served for a missing file still fails
- **CSV**: comma, semicolon (then `12,5` is a number) or tab, detected from the first line; quoted fields, newlines inside quotes
- **Encoding**: byte order mark (UTF-8, UTF-16), the declared `charset`, UTF-8, else Windows-1252 (Excel's CSV on Windows)
- The bar shows a `CSV` / `TSV` / `TXT` / `MD` / `HTML` badge; saving writes the workbook / document (`.xlsx` / `.docx`); *Download as* CSV / text / Markdown gives the plain format back

## Editing

### DOCX

- **Typing**: type/delete/Enter, paste (keeps formatting when pasting from Word/Google Docs/web), IME input (e.g. Vietnamese)
- **Text formatting**: B/I/U/S, superscript/subscript, font, size, color, highlight, paint format
- **Styles**: Normal, Title, Subtitle, Heading 1–6 always available (created in the file when missing; Ctrl/⌘+Alt+0…6), update a style to match the selection (right-click → Update style to match: font, size, color, B/I/U/S, alignment and spacing go into the style, every paragraph of it follows; written to that file's own styles only), document outline panel
- **Paragraphs**: alignment, line spacing, bullet/number list, indent, page break
- **Images**: insert (file picker/paste/drag & drop), click to select with 8 resize handles, original size, fit width, replace, alt text, download
- **Tables**: insert, add/delete rows & columns, delete table, Tab between cells; select cells by dragging across them, merge / unmerge (across and down), cell color, cell borders (all / outer / inner / none / one side, color, line style), vertical alignment, distribute columns, drag a column edge to resize
- **Page setup**: paper size (A4, Letter, Legal, A3, A5, B5), portrait / landscape, margins (cm, presets), 1–3 text columns with spacing; applies to the whole document
- **Headers & footers**: header, footer and first-page parts (double-click the page's header / footer area, or the toolbar), alignment, bold, size, page number / page count fields with presets ("Page 1 of 5", "Trang 1/5"), different first page
- **Tracked changes**: insertions and deletions from Word shown in their author's color and kept on save; next / previous change, accept / reject (the change at the caret, the selection, or all); *Track changes* records your own edits as suggestions (typing, deleting, pasting, replacing)
- **Other**: select all (Ctrl/⌘+A), link (dialog), special characters, right-click context menu (clipboard, link, table, image, insert, styles), find & replace, word count, undo/redo

### XLSX

- **Cells**: edit, range selection, select all (Ctrl/⌘+A), copy/cut/paste, fill down/right, sort, find & replace (match case, entire cell, in formulas)
- **Formulas**: auto-calculated (134 functions), function name & argument hints while typing, click/drag cells (incl. other sheets, arrow keys) to insert references, colored references, auto-close parentheses, functions menu (SUM/AVERAGE/COUNT/MAX/MIN)
- **Formatting**: paint format, currency/percent/decimal buttons + 20 number formats, font, size, B/I/U/S, text/fill color, borders (8 kinds, color, line style), merge all/horizontally/vertically, horizontal & vertical alignment, wrap, text rotation, clear formatting
- **Rows & columns**: insert/delete/hide/unhide, drag to resize column width/row height, freeze, gridlines
- **Sheets**: add/rename/delete/reorder
- **Other**: context menu, undo/redo

### PPTX

- **Text**: font, size, B/I/U/S, color, highlight, superscript/subscript, alignment, vertical alignment, line spacing, bullet/numbered list, indent level, clear formatting
- **Shapes**: select/drag/resize/delete (arrow keys to nudge), fill, border color/weight/dash, arrange (front/forward/backward/back), rotate 90°/flip, align on slide, duplicate (⌘D), undo/redo
- **Insert**: 30 shapes, text box, image
- **Slides**: background, add/duplicate/delete/move/hide, slideshow from current slide or beginning
- **Other**: context menus on shapes/slides/thumbnails

### Toolbar

Single row, buttons that don't fit collapse into `⋮`; custom-drawn dropdowns (font/style preview, number formats with examples, ↑↓/Enter/Esc/type-ahead keys, never steals focus from the document); font size box `− 11 +` accepts direct input; cell borders picked from an icon grid; color pickers have a palette plus a **Custom** row (`+` opens the system color picker, recent custom colors are remembered across sessions).

### Dark mode

`theme="dark"`, or `"auto"` to follow the OS / browser setting (`prefers-color-scheme`), darkens the chrome: bars, toolbar, menus, dialogs, sheet headers and tabs, thumbnails, popovers, scrollbars. Pages, slides and cells keep the colors the document gives them, like paper in a dark PDF viewer. Everything is driven by CSS variables on `.de-root`, so a host can also tune the palette.

## Comments

- **Where**: DOCX — the selected text (or the word at the caret); XLSX — a cell (one thread per cell, as in Excel: commenting on a cell that has one answers it); PPTX — next to the selected shape, or where the slide is right-clicked
- **Start one**: `Ctrl/⌘+Alt+M`, right-click → Comment, or Comments → New comment
- **Threads**: replies, resolve / reopen (answering a resolved thread reopens it), edit and delete your own comments, deleting a thread's first comment deletes the thread
- **@mentions**: type `@` to pick someone from the host's list (`comments.mentions`); the host hears who was mentioned (`comments.onEvent`) to notify them
- **Panel**: threads in document order with what they point at (the commented text, `Sheet1!B3`, `Slide 2`), resolved ones on demand; picking a thread shows its place, clicking commented text / a marked cell / a pin on a slide shows its thread
- **In the document**: commented text is highlighted (stronger for the current thread), cells get a mark in their top right corner, slides a pin with the author's initials; resolved threads are not shown
- **DOCX ranges follow the text**: typing inside a commented range extends it, typing right after it does not; splitting a paragraph keeps both parts covered; undo takes a new comment away (redo brings it back); a thread whose text is deleted leaves the file
- **XLSX**: threads move with inserted / deleted rows and columns and go with a deleted row, column or sheet (undo brings them back)
- **PPTX**: a slide's threads stay with it (not copied by Duplicate slide, deleted with the slide)
- **Shared rooms**: comments are shared operations like edits — every editor sees them live, newcomers get them from the file or the history
- **Read-only** while viewing (default: comments can be written while editing, `comments.canComment` overrides)

Stored the Office way, so Word, Excel and PowerPoint show them:

| Format | Parts                                                                                                                 | Notes                                                                                                                                                  |
|--------|-----------------------------------------------------------------------------------------------------------------------|--------------------------------------------------------------------------------------------------------------------------------------------------------|
| DOCX   | `comments.xml`, `commentsExtended.xml` (threads, resolved), `commentsIds.xml`, `commentsExtensible.xml`, `people.xml` | ranges as `commentRangeStart` / `commentRangeEnd` + reference; untouched comments keep their XML (rich text); comments in headers / footnotes are kept |
| XLSX   | `threadedComments/*.xml`, `persons/person.xml`, plus the legacy `comments*.xml` + VML Excel keeps next to them        | notes (legacy comments) are read as one-comment threads; a note nobody answered or changed stays a note                                                |
| PPTX   | modern comments `comments/modernComment_*.xml`, `authors.xml`                                                         | legacy comments are read as one-comment threads and written back as modern comments                                                                    |

Comment text is plain (lines kept); formatting inside comments read from a file is kept while the comment is not edited. Ids are durable: a comment keeps its id after saving and reopening (DOCX durable ids, XLSX / PPTX GUIDs).

## File menu

- **New**: blank document, spreadsheet or presentation (16:9, Office theme, title / content / blank layouts)
- **Open** (`Ctrl/⌘+O`): `.docx` `.xlsx` `.pptx` open as they are; `.csv` / `.tsv` become a workbook (numbers and booleans typed, codes like `007` kept as text); `.txt` / `.md` / `.html` become a document (headings, lists, bold / italic / strike, links, code)
- **Download as**: the native format, PDF, web page (`.html`, images embedded), plain text / Markdown (DOCX), CSV / TSV of the current sheet (XLSX), slide text with notes (PPTX)
- **Import CSV into new sheet** (XLSX): adds a sheet after the current one with the file's rows
- **Save** (`Ctrl/⌘+S`), **Rename** (shown name and the name used for saving), **Document details** (file, core properties, page / word / slide counts)
- **Host entries** (`fileMenuItems`: Share, Version history…) and **Close** (`onClose`)
- **PDF / Print** (`Ctrl/⌘+P`): rendered off-screen and sent to the browser print dialog (choose "Save as PDF"): DOCX pages as paginated on screen, PPTX one slide per page, XLSX current sheet fitted to the paper width

## Saving

- DOCX: edited paragraphs are regenerated from the original XML (keeps `pPr`, `rPr`, images, bookmarks, comments); untouched paragraphs keep original bytes
- XLSX: only `<row>`s containing edited cells are rewritten (streamed, never loads the whole sheet into a DOM); new strings written as inline strings; `fullCalcOnLoad` set so Excel recalculates formulas
- PPTX: only `a:p` of edited shapes are regenerated
- Unchanged ZIP entries are copied as-is (no recompression)

| Operation                  | File                   | Time   |
|----------------------------|------------------------|--------|
| Type 1 char (browser)      | DOCX 2073 pages        | ~7 ms  |
| Save after edit            | DOCX 20k paragraphs    | ~0.5 s |
| Save after editing 2 cells | XLSX 300k rows (25 MB) | ~3 s   |

## Limitations

- Formulas: no defined names, structured refs (`Table1[Col]`), 3-D refs
- DOCX: text boxes and nested tables are read-only; paragraphs with fields spanning multiple paragraphs (TOC) are read-only; headers / footers are edited as text lines (pictures, tables and shapes in them are replaced when edited, with a warning); a header taller than the top margin overlaps the text; tracked changes don't cover formatting changes, paragraph marks or pictures; page setup applies to the whole document (not per section)
- XLSX: no chart/pivot/conditional format/data validation/filter editing; inserting/deleting rows & columns doesn't shift conditional formats/data validation
- PPTX: can't rotate shapes with the mouse yet, can't edit text inside tables yet, no table insert, no animations/transitions; slide list changes (add/delete/move) are not undoable
- Comments: plain text only; DOCX comments can't cover text in more than one table cell or across a table; comments in headers, footers and footnotes are kept but not shown; XLSX comments don't move with sorted rows; PPTX comments anchored to a shape are shown at the shape's position when the file is opened and keep that anchor until moved
