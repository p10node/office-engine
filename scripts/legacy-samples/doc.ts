/**
 * Writes test/fixtures/legacy/sample.doc, a Word 97 document with the content of the DOCX sample: title, headings,
 * character formatting, a link, bullet and numbered lists, a picture, a table with merged and shaded cells,
 * a footnote, and a header and footer with a page number. Usage: bun scripts/legacy-samples/doc.ts
 */
import { png } from '../../test/iwork-builder'
import { C, DocBuilder, P, S, STYLES } from '../../test/doc-builder'

const d = new DocBuilder()
const bullets = d.list([{ nfc: 23, text: '\uf0b7', left: 720, hang: 360, font: 1 }])
const numbers = d.list([{ nfc: 0, text: '\u0000.', nums: [1], left: 720, hang: 360 }])

d.p('Quarterly Report', STYLES.title)
d.p([
  { text: 'This document shows ' },
  { text: 'bold', c: C.bold() },
  { text: ', ' },
  { text: 'italic', c: C.italic() },
  { text: ', ' },
  { text: 'underlined red', c: [...C.underline(1), ...C.color('C00000')] },
  { text: ' and ' },
  { text: 'large', c: C.size(16) },
  { text: ' text. Visit ' },
  ...d.hyperlink('https://example.com/', 'example.com'),
  { text: '.' },
  d.footnote('Figures are unaudited and rounded to one decimal place.'),
])
d.p('Highlights', STYLES.heading1)
for (const item of ['Revenue grew 24% year over year', 'Two new regions launched', 'Customer churn fell to 3.1%']) d.p(item, STYLES.listBullet, P.list(bullets))
d.p('Next steps', STYLES.heading1)
for (const item of ['Hire the platform team', 'Ship the viewer engine', 'Expand to APAC']) d.p(item, STYLES.listNumber, P.list(numbers))
d.p('Figures', STYLES.heading2)
d.p([d.picture(png(240, 90), 4320, 1620)], 0, P.jc(1))
d.p('Figure 1 — generated gradient image', STYLES.caption)
d.p('Table', STYLES.heading2)
const head = (t: string) => ({ paras: [[{ text: t, c: C.bold() }]], shd: 'D9E2F3' })
d.table({
  widths: [2400, 1800, 1800, 1800],
  header: true,
  rows: [
    [head('Region'), head('Q1'), head('Q2'), head('Total')],
    [{ paras: ['North'] }, { paras: ['120'] }, { paras: ['140'] }, { paras: ['260'] }],
    [{ paras: ['South'] }, { paras: ['98'] }, { paras: ['110'] }, { paras: ['208'] }],
    [{ paras: ['East'] }, { paras: ['132'] }, { paras: ['150'] }, { paras: ['282'] }],
    [{ paras: [[{ text: 'Merged footer cell (3 columns)', c: C.italic() }]], hm: 1, shd: 'F2F2F2' }, { paras: [''], hm: 2 }, { paras: [''], hm: 2 }, { paras: [[{ text: '750', c: C.bold() }]], shd: 'F2F2F2' }],
  ],
})
d.p('Appendix', STYLES.heading1)
d.p([{ text: 'Notes', c: C.style(STYLES.strong) }, { text: '\tFigures in thousands.' }], 0, P.tabs([[1440, 0, 1]]))
d.p('The quarter closed ahead of plan. Growth came from the two new regions and from renewals in the existing ones; costs stayed flat.', 0, [...P.jc(3), ...P.first(360)])
d.p([{ text: 'Prepared by the finance team. ' }, { text: 'Confidential', c: [...C.smallCaps(), ...C.color('7F7F7F')] }])
d.headers([undefined, [{ text: 'Quarterly Report\tFinance' }], undefined, [{ text: 'Page ' }, ...d.field('PAGE', '1'), { text: ' of ' }, ...d.field('NUMPAGES', '1')], undefined, undefined])

const bytes = d.build([...S.page(12240, 15840), ...S.margins(1440, 1440, 1440, 1440), ...S.headerDist(720, 720)])
const out = `${import.meta.dir}/../../test/fixtures/legacy/sample.doc`
await Bun.write(out, bytes)
console.log('wrote', out, bytes.length, 'bytes')
