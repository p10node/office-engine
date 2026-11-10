# PowerPoint 97-2003 fixtures

The fixture is synthesized record by record, so building it needs neither PowerPoint nor LibreOffice, by
`test/ppt-builder.ts` (records, OfficeArt drawings, text runs, master styles, the persist directory, the edit and the
Current User stream, in a compound file written by `test/cfb-writer.ts`). The output is the same on every run.

```sh
bun scripts/legacy-samples/ppt.ts
```

- `sample.ppt`: the playground sample, modeled on `playground/public/samples/sample.pptx`. A main master (title, body
  and footer placeholders, a colored band, five bullet levels) and a title master with a gradient background; a title
  slide; a bullet slide (levels, bold / italic / colored / underlined runs, a Wingdings bullet, a link, a numbered
  list stored as PowerPoint 2000 extras, speaker notes, a PowerPoint 2002 comment); a shapes slide (rounded
  rectangle, oval, arrow with a gradient, text box, picture with a crop, dashed line with an arrowhead, freeform, arc,
  rotated text); a table stored as a group of cells; a hidden slide with its own background and color scheme; footers
  with slide numbers

The tests build their other cases with the same builder (incremental saves, a missing Current User stream, encrypted
and PowerPoint 95 files, JPEG / DIB / EMF pictures, groups in groups, comments, damaged files).

macOS Quick Look renders the fixture like the converted PPTX (`qlmanage -t -s 600 test/fixtures/legacy/sample.ppt`
shows the first slide), which checks the builder against a reader that is not ours.
