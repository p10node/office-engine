# Word 97-2003 fixtures

The `.doc` files of `test/fixtures/legacy/`, used by `test/legacy-doc.test.ts`.

- `sample.doc`: the playground sample, modeled on `playground/public/samples/sample.docx`. Title, headings, bold /
  italic / underlined red / large text, a hyperlink field, a footnote, a bullet list and a numbered list (list tables),
  an inline picture (PNG in the Data stream), a caption, a table with a shaded header row, a merged footer cell and
  borders, a tab with a dot leader, a justified paragraph, small caps, and a header and a footer with PAGE and
  NUMPAGES fields. Written by `test/doc-builder.ts` (a Word 97 writer for tests: compound file, FIB, piece table,
  FKPs, stylesheet, fonts, lists, sections, headers, notes, fields, pictures); the output is the same on every run:

  ```sh
  bun scripts/legacy-samples/doc.ts
  ```

- `textutil.doc`: written by macOS `textutil` (Cocoa's Word writer): 8-bit text pieces, direct formatting (bold,
  italic, underline, strike, color, size, font, superscript / subscript), alignment and a table, from this HTML:

  ```html
  <html><head><meta charset="utf-8"><title>Sample</title></head><body>
  <h1>Quarterly Report</h1>
  <p>This is <b>bold</b>, <i>italic</i>, <u>underlined</u>, <s>struck</s>, <span style="color:#c00000">red</span>, <span style="font-size:18pt">large</span>, <span style="font-family:Courier New">monospace</span>, x<sup>2</sup> and H<sub>2</sub>O.</p>
  <h2>Highlights</h2>
  <ul><li>First item</li><li>Second item</li></ul>
  <ol><li>Step one</li><li>Step two</li></ol>
  <p style="text-align:center">Centered paragraph</p>
  <p style="text-align:right">Right aligned</p>
  <p>A link to <a href="https://example.com/">example.com</a>.</p>
  <table border="1"><tr><th>Region</th><th>Sales</th></tr><tr><td>North</td><td>120</td></tr><tr><td>South</td><td>95</td></tr></table>
  <p>End of document.</p>
  </body></html>
  ```

  ```sh
  textutil -convert doc rich.html -output test/fixtures/legacy/textutil.doc
  ```

  textutil writes no styles, list tables, fields or pictures; the tests build those with `test/doc-builder.ts`.

Real Word files (Apache POI `test-data/document`, LibreOffice `sw/qa/extras/ww8import` and `ww8export`, Apache Tika
test documents) are used for validation and are not committed.
