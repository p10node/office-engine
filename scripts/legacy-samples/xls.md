# Excel 97-2003 fixtures

The `.xls` files of `test/fixtures/legacy/`, used by `test/legacy-xls.test.ts`. Built with
[xlwt](https://pypi.org/project/xlwt/) 1.3.0 (BIFF8) by `xls.py`; the output is the same on every run.

```sh
python3 -m venv /tmp/xls-venv && /tmp/xls-venv/bin/pip install xlwt==1.3.0
/tmp/xls-venv/bin/python scripts/legacy-samples/xls.py test/fixtures/legacy
```

- `sample.xls`: the playground sample, modeled on `playground/public/samples/sample.xlsx`. A "Sales" sheet (merged
  title, a header row with a fill and borders, dates, currency, percentages, formulas with a total row, references to
  other sheets, frozen panes, column widths, landscape page setup with header and footer, fit to one page wide), a
  "Notes" sheet (wrapped text, italic / bold / underlined / struck / colored text, fills, rotation, indent, rich text,
  a hidden row and column, another font) and a 200-row "Data" sheet
- `strings.xls`: 300 strings, 8-bit and 16-bit, and two long ones: a shared string table over many CONTINUE records,
  with strings split across them
- `dates1904.xls`: dates in the 1904 date system

xlwt writes no cached formula results, comments, hyperlinks, shared or array formulas, defined names or BIFF5. The
tests build those record by record with `test/xls-builder.ts`.
