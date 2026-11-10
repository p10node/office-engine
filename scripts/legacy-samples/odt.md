# OpenDocument text fixtures

The `.odt` / `.fodt` / `.ott` files of `test/fixtures/legacy/`, used by `test/legacy-odt.test.ts`.

## sample.odt, sample.fodt, template.ott

Built with [odfpy](https://pypi.org/project/odfpy/) by `odt.py`, laid out the way LibreOffice writes documents (its
style names, automatic styles for direct formatting, a master page with a header and footer). The picture is the
gradient of `playground/public/samples/sample.docx`.

```sh
python3 -m venv /tmp/odf-venv && /tmp/odf-venv/bin/pip install odfpy
unzip -o -j playground/public/samples/sample.docx word/media/image1.png -d /tmp/odf-out
/tmp/odf-venv/bin/python scripts/legacy-samples/odt.py test/fixtures/legacy /tmp/odf-out/image1.png
```

- `sample.odt`: title, table of contents, headings with outline numbering, character formatting, a footnote, links and
  bookmarks, bullet and numbered lists (nested, one continued), a table with a header row and merged cells, a picture,
  a text frame, tab stops with leaders, a page break, a two-column section, header and footer with page fields
- `sample.fodt`: the same document as flat XML (the picture inline as base64)
- `template.ott`: a short template (`application/vnd.oasis.opendocument.text-template`)

## pandoc.odt

pandoc 3.13's ODT writer, from the playground's Word sample:

```sh
pandoc playground/public/samples/sample.docx -o test/fixtures/legacy/pandoc.odt
```

## textutil.odt

macOS `textutil`'s ODT writer (OpenDocument 1.0), from the playground's HTML sample:

```sh
textutil -convert odt -output test/fixtures/legacy/textutil.odt playground/public/samples/sample.html
```
