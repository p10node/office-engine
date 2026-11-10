# OpenDocument presentation fixtures

The fixtures of `test/fixtures/legacy/` for the ODP → PPTX converter are generated with [odfpy](https://github.com/eea/odfpy) by `scripts/legacy-samples/odp.py`.

| File          | Content                                                                    |
| ------------- | -------------------------------------------------------------------------- |
| `sample.odp`  | the demo deck (6 slides, about 6 KB)                                       |
| `sample.otp`  | the same deck as a template (`presentation-template` mimetype)             |
| `sample.fodp` | the same deck as flat XML, the picture inline as `office:binary-data`      |

The deck follows `playground/public/samples/sample.pptx`:

1. Title slide: title and subtitle placeholders, speaker notes.
2. Outline with two levels of bullets, bold, italic, underlined, highlighted and superscript runs, a hyperlink, notes.
3. Shapes: a rounded rectangle, an ellipse, a right arrow with its enhanced path and equations, a star turned by `draw:transform`, a text box with colored bold text, a PNG picture with a border, a line with an arrow head.
4. A table with column widths, a header row and body cell fills.
5. A numbered list with a lettered sub-list on a tinted slide background.
6. A hidden slide.

The master page has a gradient band, title and outline placeholder areas, a footer and a slide number; the outline styles (`Default-outline1..3`) carry the bullets. The picture is drawn by the script (a small landscape), so nothing is copied from elsewhere.

## Regenerate

```sh
python3 -m venv /tmp/odp-venv
/tmp/odp-venv/bin/pip install odfpy
/tmp/odp-venv/bin/python -I scripts/legacy-samples/odp.py test/fixtures/legacy
```

The output is deterministic except for the creation date odfpy writes into `meta.xml`.

Other test documents (transforms, custom shape paths, fields, list headers, master pages and footers, crops, merged table cells, damaged input) are built as flat XML strings in `test/legacy-odp.test.ts`.
