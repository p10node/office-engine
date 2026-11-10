# OpenDocument spreadsheet fixtures

The `.ods` / `.fods` files of `test/fixtures/legacy/`, used by `test/legacy-ods.test.ts`. Built with
[odfpy](https://pypi.org/project/odfpy/) 1.4 by `ods.py`, laid out the way LibreOffice writes spreadsheets (a
"Default" cell style, automatic cell / column / row / table styles, data styles with `style:map` conditions, a master
page with a header and footer, view settings in `settings.xml`). The values are random with a fixed seed.

```sh
python3 -m venv /tmp/ods-venv && /tmp/ods-venv/bin/pip install odfpy
/tmp/ods-venv/bin/python scripts/legacy-samples/ods.py test/fixtures/legacy
```

- `sample.ods`: the playground sample, modeled on `playground/public/samples/sample.xlsx`
  - "Sales": a merged title, a header row with a fill and a bottom border, 60 banded rows of dates, integers,
    currency, a `[.C3]*[.D3]` formula with its cached value, percentages and a `VLOOKUP` into the hidden "Lookup"
    sheet, a total row of `SUM` formulas; a comment on a header cell; extra values (a negative number in red
    parentheses, a boolean, a long date and time, a duration over 24 hours); frozen at B3; landscape A4 with the sheet
    name as header and "Page n of m" as footer
  - "Notes": a tall row of wrapped text in two paragraphs, italic underlined text, a hyperlink, indented items, struck
    colored text in another size, rotated text with borders; a red tab and no gridlines
  - "Data": 400 rows under an autofilter
  - "Lookup": hidden
  - a named range ("Revenue")
- `sample.fods`: the same document as flat XML, with 12 sales rows and without the "Data" sheet to stay small

Huge repeated rows and cells, value types, the null date, cell styles, sheet names Excel does not take, view settings,
links and comments, number formats and formulas are tested on packages the tests build from XML strings.
