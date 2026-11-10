# RTF fixtures

The `.rtf` files of `test/fixtures/legacy/`, used by `test/legacy-rtf.test.ts`. Most RTF features are covered by
handwritten RTF strings in the test itself.

## sample.rtf

A report laid out the way Word writes RTF: a stylesheet with Word's built-in styles (Normal, Title, Heading 1 / 2,
List Paragraph, header, footer, footnote text, Hyperlink), list tables (bullets with a nested level, numbers with a
lettered level), a header and a footer with PAGE / NUMPAGES fields, a footnote, an external and an internal link (to a
bookmark), character formatting, a picture as `\shppict` PNG with a `\nonshppict` metafile fallback, a table (shaded
header row repeated on each page, borders, a merged cell, right-aligned numbers), a right tab with a dot leader, a
paragraph with borders and shading, and a page break. The picture is the gradient of
`playground/public/samples/sample.docx`.

```sh
bun scripts/legacy-samples/rtf.ts
```

## textutil.rtf

macOS `textutil` (Cocoa text system, macOS 27) from this HTML:

```html
<html><body>
<h1>Project Notes</h1>
<p>Plain text with <b>bold</b>, <i>italic</i>, <u>underline</u>, <s>strike</s>, <span style="color:#c00000">red</span> and <a href="https://example.com/docs">a link</a>.</p>
<h2>Tasks</h2>
<ul><li>Write the parser</li><li>Test real files<ul><li>Word and pandoc</li></ul></li></ul>
<ol><li>First step</li><li>Second step</li></ol>
<p style="text-align:center">Centered line</p>
<table border="1"><tr><th>Name</th><th>Count</th></tr><tr><td>Alpha</td><td>3</td></tr><tr><td colspan="2">Merged footer</td></tr></table>
<p>Caf&eacute; costs 5 &euro; &mdash; done.</p>
</body></html>
```

```sh
textutil -convert rtf textutil.html -output test/fixtures/legacy/textutil.rtf
```

## pandoc.rtf

pandoc 3.13's RTF writer, from the playground's Word sample (headings as formatted text, lists, a picture, a table,
a link):

```sh
pandoc playground/public/samples/sample.docx -s -t rtf -o test/fixtures/legacy/pandoc.rtf
```
