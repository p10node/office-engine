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
| DOCX   | inherited styles, run formatting, lists/numbering, tables (merge, border, shading, table style), images, hyperlinks, header/footer, PAGE field, text boxes, page/section breaks   | footnote body, charts, exact float wrapping, paragraphs split across pages |
| XLSX   | shared strings, number/date formats, font/fill/border/align, merge, freeze, hidden columns/rows, wrap, formula calculation (134 functions, recalculated on edit), multiple sheets | charts, conditional formatting, images                                     |
| PPTX   | master/layout/theme, placeholder inheritance, ~50 preset shapes, custGeom, images (crop), tables + table style, groups, bullet/autonum, autofit                                   | charts (placeholder), SmartArt, animations, video playback                 |

`.doc/.xls/.ppt` (legacy binary) and encrypted files are not supported.

## Editing

### DOCX

- **Typing**: type/delete/Enter, paste (keeps formatting when pasting from Word/Google Docs/web), IME input (e.g. Vietnamese)
- **Text formatting**: B/I/U/S, superscript/subscript, font, size, color, highlight, paint format
- **Styles**: Normal, Title, Subtitle, Heading 1–6 always available (created in the file when missing; Ctrl/⌘+Alt+0…6), document outline panel
- **Paragraphs**: alignment, line spacing, bullet/number list, indent, page break
- **Images**: insert (file picker/paste/drag & drop), click to select with 8 resize handles, original size, fit width, replace, alt text, download
- **Tables**: insert, add/delete rows & columns, delete table, Tab between cells
- **Other**: link (dialog), special characters, right-click context menu (clipboard, link, table, image, insert), find & replace, word count, undo/redo

### XLSX

- **Cells**: edit, range selection, copy/cut/paste, fill down/right, sort, find & replace (match case, entire cell, in formulas)
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

Single row, buttons that don't fit collapse into `⋮`; custom-drawn dropdowns (font/style preview, number formats with examples, ↑↓/Enter/Esc/type-ahead keys, never steals focus from the document); font size box `− 11 +` accepts direct input; cell borders picked from an icon grid.

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
- DOCX: header/footer, text boxes and nested tables are read-only; paragraphs with fields spanning multiple paragraphs (TOC) are read-only; "update style to match selection" is not available
- XLSX: no chart/pivot/conditional format/data validation/filter editing; inserting/deleting rows & columns doesn't shift conditional formats/data validation
- PPTX: can't rotate shapes with the mouse yet, can't edit text inside tables yet, no table insert, no animations/transitions; slide list changes (add/delete/move) are not undoable
