/**
 * Writes test/fixtures/legacy/sample.rtf: a report laid out the way Word writes RTF (a stylesheet with built-in
 * styles, list tables, \shppict pictures with a \nonshppict fallback, header and footer with page fields, a footnote).
 * The picture is the gradient of playground/public/samples/sample.docx.
 * Usage: bun scripts/legacy-samples/rtf.ts
 */
import { BlobSource } from '../../src/core/source'
import { ZipArchive } from '../../src/core/zip'

const root = `${import.meta.dir}/../..`
const docx = await ZipArchive.open(new BlobSource(Bun.file(`${root}/playground/public/samples/sample.docx`)))
const png = await docx.bytes('word/media/image1.png')
const hex = Buffer.from(png).toString('hex').replace(/(.{128})/g, '$1\n')
const pw = (png[16] << 24) | (png[17] << 16) | (png[18] << 8) | png[19]
const ph = (png[20] << 24) | (png[21] << 16) | (png[22] << 8) | png[23]

const NORMAL = '\\ql \\li0\\ri0\\sa160\\sl259\\slmult1\\widctlpar\\wrapdefault\\aspalpha\\aspnum\\faauto\\adjustright\\rin0\\lin0\\itap0'
const CHP = '\\rtlch\\fcs1 \\af0\\afs22\\alang1025 \\ltrch\\fcs0 \\f1\\fs22\\lang1033\\langfe1033\\cgrid\\langnp1033\\langfenp1033'
const H1 = '\\s1\\ql \\li0\\ri0\\sb240\\keepn\\widctlpar\\wrapdefault\\aspalpha\\aspnum\\faauto\\outlinelevel0\\adjustright\\rin0\\lin0\\itap0\\sl259\\slmult1'
const H1C = '\\rtlch\\fcs1 \\af0\\afs32\\alang1025 \\ltrch\\fcs0 \\f2\\fs32\\cf3\\lang1033\\langfe1033\\cgrid\\langnp1033\\langfenp1033'
const H2 = '\\s2\\ql \\li0\\ri0\\sb40\\keepn\\widctlpar\\wrapdefault\\aspalpha\\aspnum\\faauto\\outlinelevel1\\adjustright\\rin0\\lin0\\itap0\\sl259\\slmult1'
const H2C = '\\rtlch\\fcs1 \\af0\\afs26\\alang1025 \\ltrch\\fcs0 \\f2\\fs26\\cf3\\lang1033\\langfe1033\\cgrid\\langnp1033\\langfenp1033'
const TITLE = '\\s15\\ql \\li0\\ri0\\widctlpar\\wrapdefault\\aspalpha\\aspnum\\faauto\\adjustright\\rin0\\lin0\\itap0\\contextualspace'
const TITLEC = '\\rtlch\\fcs1 \\af0\\afs56\\alang1025 \\ltrch\\fcs0 \\f2\\fs56\\expnd-2\\expndtw-10\\kerning28\\lang1033\\langfe1033\\cgrid\\langnp1033\\langfenp1033'
const LIST = '\\s17\\ql \\fi-360\\li720\\ri0\\sa160\\sl259\\slmult1\\widctlpar\\wrapdefault\\aspalpha\\aspnum\\faauto\\adjustright\\rin0\\lin720\\itap0\\contextualspace'
const r = (text: string, fmt = '') => `{\\rtlch\\fcs1 \\af0 \\ltrch\\fcs0 ${fmt}${fmt ? ' ' : ''}${text}}`
const para = (props: string, chp: string, body: string) => `\\pard\\plain \\ltrpar${props} ${chp} ${body}\\par\n`

const bullet = (text: string, lvl = 0) =>
  `{\\listtext\\pard\\plain\\ltrpar \\rtlch\\fcs1 \\af0\\afs22 \\ltrch\\fcs0 ${lvl ? '\\f4\\fs22 o' : "\\f3\\fs22 \\'b7"}\\tab}` +
  para(LIST.replace('\\fi-360\\li720', `\\fi-360\\li${720 * (lvl + 1)}`).replace('\\lin720', `\\lin${720 * (lvl + 1)}`) + `\\ls1\\ilvl${lvl}`, CHP, r(text))
const numbered = (n: string, text: string, lvl = 0) =>
  `{\\listtext\\pard\\plain\\ltrpar \\rtlch\\fcs1 \\af0\\afs22 \\ltrch\\fcs0 \\fs22 ${n}\\tab}` +
  para(LIST.replace('\\fi-360\\li720', `\\fi-360\\li${720 * (lvl + 1)}`).replace('\\lin720', `\\lin${720 * (lvl + 1)}`) + `\\ls2\\ilvl${lvl}`, CHP, r(text))

// the table: a shaded header row repeated on each page, three data rows, a total row whose first cell spans three columns
const cellBorders = '\\clbrdrt\\brdrs\\brdrw4\\brdrcf8 \\clbrdrl\\brdrs\\brdrw4\\brdrcf8 \\clbrdrb\\brdrs\\brdrw4\\brdrcf8 \\clbrdrr\\brdrs\\brdrw4\\brdrcf8 '
const rowDef = (header: boolean, merged = false) =>
  `\\trowd \\irow0\\irowband0\\ltrrow\\ts11\\trgaph108\\trleft-108${header ? '\\trhdr' : ''}\\trbrdrt\\brdrs\\brdrw4\\brdrcf8 \\trbrdrl\\brdrs\\brdrw4\\brdrcf8 \\trbrdrb\\brdrs\\brdrw4\\brdrcf8 \\trbrdrr\\brdrs\\brdrw4\\brdrcf8 \\trbrdrh\\brdrs\\brdrw4\\brdrcf8 \\trbrdrv\\brdrs\\brdrw4\\brdrcf8 ` +
  '\\trftsWidth1\\trftsWidthB3\\trftsWidthA3\\trautofit1\\trpaddl108\\trpaddr108\\trpaddfl3\\trpaddft3\\trpaddfb3\\trpaddfr3\\tblrsid1\\tblind0\\tblindtype3 ' +
  [2340, 4680, 7020, 9360]
    .map((x, k) => `\\clvertalt${cellBorders}${header ? '\\clcbpat5' : ''}${merged ? (k === 0 ? '\\clmgf' : k < 3 ? '\\clmrg' : '') : ''}\\cltxlrtb\\clftsWidth3\\clwWidth2340${header ? '' : '\\clshdrawnil'} \\cellx${x}`)
    .join('')
const CELL = '\\ql \\li0\\ri0\\widctlpar\\intbl\\wrapdefault\\aspalpha\\aspnum\\faauto\\adjustright\\rin0\\lin0'
const row = (cells: string[], header = false, merged = false) =>
  `\\pard\\plain \\ltrpar${CELL} ${CHP} ` +
  cells.map((c, k) => (k ? `\\pard\\plain \\ltrpar${k === 3 ? CELL.replace('\\ql', '\\qr') : CELL} ${CHP} ` : '') + r(c, header ? '\\b' : '') + '\\cell ').join('') +
  `\\pard\\plain \\ltrpar${CELL} ${CHP} {\\rtlch\\fcs1 \\af0 \\ltrch\\fcs0 ${rowDef(header, merged)}\\row }\n`

const picture =
  `{\\*\\shppict{\\pict{\\*\\picprop\\shplid1025{\\sp{\\sn shapeType}{\\sv 75}}{\\sp{\\sn fFlipH}{\\sv 0}}{\\sp{\\sn fFlipV}{\\sv 0}}{\\sp{\\sn pibName}{\\sv image1.png}}}` +
  `\\picscalex100\\picscaley100\\piccropl0\\piccropr0\\piccropt0\\piccropb0\\picw${pw}\\pich${ph}\\picwgoal6480\\pichgoal${Math.round((6480 * ph) / pw)}\\pngblip\\bliptag-1\n${hex}}}` +
  // readers without \shppict get a metafile; a converter that shows the PNG skips it
  '{\\nonshppict{\\pict\\picscalex100\\picscaley100\\picw100\\pich100\\picwgoal1440\\pichgoal1440\\wmetafile8\n010009000003220000000000}}'

const rtf =
  '{\\rtf1\\adeflang1025\\ansi\\ansicpg1252\\uc1\\adeff0\\deff0\\stshfdbch0\\stshfloch1\\stshfhich1\\stshfbi0\\deflang1033\\deflangfe1033\\themelang1033\\themelangfe0\\themelangcs0\n' +
  '{\\fonttbl{\\f0\\fbidi \\froman\\fcharset0\\fprq2{\\*\\panose 02020603050405020304}Times New Roman;}{\\f1\\fswiss\\fcharset0\\fprq2{\\*\\panose 020f0502020204030204}Calibri;}\n' +
  '{\\f2\\fswiss\\fcharset0\\fprq2{\\*\\panose 020f0302020204030204}Calibri Light;}{\\f3\\fbidi \\froman\\fcharset2\\fprq2{\\*\\panose 05050102010706020507}Symbol;}{\\f4\\fbidi \\fmodern\\fcharset0\\fprq1{\\*\\panose 02070309020205020404}Courier New;}}\n' +
  '{\\colortbl;\\red0\\green0\\blue0;\\red0\\green0\\blue255;\\red47\\green84\\blue150;\\red192\\green0\\blue0;\\red217\\green226\\blue243;\\red255\\green255\\blue0;\\red5\\green99\\blue193;\\red191\\green191\\blue191;\\red242\\green242\\blue242;}\n' +
  '{\\*\\defchp \\f1\\fs22\\lang1033\\langfe1033\\langfenp1033 }{\\*\\defpap \\ql \\li0\\ri0\\sa160\\sl259\\slmult1\\widctlpar\\wrapdefault\\aspalpha\\aspnum\\faauto\\adjustright\\rin0\\lin0\\itap0 }\n' +
  '{\\stylesheet' +
  `{${NORMAL} ${CHP} \\snext0 \\sqformat \\spriority0 Normal;}\n` +
  `{${H1} ${H1C} \\sbasedon0 \\snext0 \\slink16 \\sqformat \\spriority9 heading 1;}\n` +
  `{${H2} ${H2C} \\sbasedon0 \\snext0 \\sqformat \\spriority9 \\sunhideused heading 2;}\n` +
  '{\\*\\cs10 \\additive \\ssemihidden \\sunhideused \\spriority1 Default Paragraph Font;}\n' +
  '{\\*\\ts11\\tsrowd\\trftsWidthB3\\trpaddl108\\trpaddr108\\trpaddfl3\\trpaddft3\\trpaddfb3\\trpaddfr3\\tblind0\\tblindtype3\\tsvertalt\\tsbrdrt\\tsbrdrl\\tsbrdrb\\tsbrdrr\\tsbrdrdgl\\tsbrdrdgr\\tsbrdrh\\tsbrdrv \\ql \\li0\\ri0\\sa160\\sl259\\slmult1\\widctlpar\\wrapdefault\\aspalpha\\aspnum\\faauto\\adjustright\\rin0\\lin0\\itap0 \\rtlch\\fcs1 \\af0\\afs22\\alang1025 \\ltrch\\fcs0 \\f1\\fs22\\lang1033\\langfe1033\\cgrid\\langnp1033\\langfenp1033 \\snext11 \\ssemihidden \\sunhideused Normal Table;}\n' +
  `{${TITLE} ${TITLEC} \\sbasedon0 \\snext0 \\sqformat \\spriority10 Title;}\n` +
  '{\\*\\cs16 \\additive \\rtlch\\fcs1 \\af0\\afs32 \\ltrch\\fcs0 \\f2\\fs32\\cf3 \\sbasedon10 \\slink1 \\spriority9 Heading 1 Char;}\n' +
  `{${LIST} ${CHP} \\sbasedon0 \\sqformat \\spriority34 List Paragraph;}\n` +
  '{\\*\\cs18 \\additive \\rtlch\\fcs1 \\af0 \\ltrch\\fcs0 \\ul\\cf7 \\sbasedon10 \\sunhideused Hyperlink;}\n' +
  '{\\s19\\ql \\li0\\ri0\\widctlpar\\tqc\\tx4680\\tqr\\tx9360\\wrapdefault\\aspalpha\\aspnum\\faauto\\adjustright\\rin0\\lin0\\itap0 ' + CHP + ' \\sbasedon0 \\sunhideused header;}\n' +
  '{\\s20\\qc \\li0\\ri0\\widctlpar\\wrapdefault\\aspalpha\\aspnum\\faauto\\adjustright\\rin0\\lin0\\itap0 ' + CHP + ' \\sbasedon0 \\sunhideused footer;}\n' +
  '{\\s21\\ql \\li0\\ri0\\widctlpar\\wrapdefault\\aspalpha\\aspnum\\faauto\\adjustright\\rin0\\lin0\\itap0 \\rtlch\\fcs1 \\af0\\afs20\\alang1025 \\ltrch\\fcs0 \\f1\\fs20\\lang1033\\langfe1033\\cgrid\\langnp1033\\langfenp1033 \\sbasedon0 \\ssemihidden \\sunhideused footnote text;}\n' +
  '{\\*\\cs22 \\additive \\rtlch\\fcs1 \\af0 \\ltrch\\fcs0 \\super \\sbasedon10 \\ssemihidden \\sunhideused footnote reference;}}\n' +
  '{\\*\\listtable' +
  '{\\list\\listtemplateid1\\listhybrid' +
  "{\\listlevel\\levelnfc23\\levelnfcn23\\leveljc0\\leveljcn0\\levelfollow0\\levelstartat1\\levelspace0\\levelindent0{\\leveltext\\leveltemplateid67698689\\'01\\u-3913 ?;}{\\levelnumbers;}\\f3\\fbias0 \\fi-360\\li720\\lin720 }" +
  "{\\listlevel\\levelnfc23\\levelnfcn23\\leveljc0\\leveljcn0\\levelfollow0\\levelstartat1\\lvltentative\\levelspace0\\levelindent0{\\leveltext\\leveltemplateid67698691\\'01o;}{\\levelnumbers;}\\f4\\fbias0 \\fi-360\\li1440\\lin1440 }" +
  '{\\listname ;}\\listid100}\n' +
  '{\\list\\listtemplateid2\\listhybrid' +
  "{\\listlevel\\levelnfc0\\levelnfcn0\\leveljc0\\leveljcn0\\levelfollow0\\levelstartat1\\levelspace0\\levelindent0{\\leveltext\\leveltemplateid67698703\\'02\\'00.;}{\\levelnumbers\\'01;}\\fi-360\\li720\\lin720 }" +
  "{\\listlevel\\levelnfc4\\levelnfcn4\\leveljc0\\leveljcn0\\levelfollow0\\levelstartat1\\lvltentative\\levelspace0\\levelindent0{\\leveltext\\leveltemplateid67698713\\'02\\'01.;}{\\levelnumbers\\'01;}\\fi-360\\li1440\\lin1440 }" +
  '{\\listname ;}\\listid200}}\n' +
  '{\\*\\listoverridetable{\\listoverride\\listid100\\listoverridecount0\\ls1}{\\listoverride\\listid200\\listoverridecount0\\ls2}}\n' +
  '{\\*\\rsidtbl \\rsid1\\rsid2}{\\mmathPr\\mmathFont34\\mbrkBin0\\mbrkBinSub0\\msmallFrac0\\mdispDef1\\mlMargin0\\mrMargin0\\mdefJc1\\mwrapIndent1440\\mintLim0\\mnaryLim1}\n' +
  '{\\info{\\title Quarterly Report}{\\author Office Engine}{\\operator Office Engine}{\\creatim\\yr2026\\mo11\\dy10\\hr9\\min0}{\\version1}{\\nofpages2}{\\nofwords240}}\n' +
  '{\\*\\xmlnstbl {\\xmlns1 http://schemas.microsoft.com/office/word/2003/wordml}}\n' +
  '\\paperw12240\\paperh15840\\margl1440\\margr1440\\margt1440\\margb1440\\gutter0\\ltrsect\n' +
  '\\deftab720\\widowctrl\\ftnbj\\aenddoc\\trackmoves0\\trackformatting1\\donotembedsysfont1\\relyonvml0\\donotembedlingdata0\\grfdocevents0\\validatexml1\\showplaceholdtext0\\ignoremixedcontent0\\saveinvalidxml0\\showxmlerrors1\\noxlattoyen\n' +
  '\\expshrtn\\noultrlspc\\dntblnsbdb\\nospaceforul\\formshade\\horzdoc\\dgmargin\\dghspace180\\dgvspace180\\dghorigin1440\\dgvorigin1440\\dghshow1\\dgvshow1\n' +
  '\\jexpand\\viewkind1\\viewscale100\\pgbrdrhead\\pgbrdrfoot\\splytwnine\\ftnlytwnine\\htmautsp\\nolnhtadjtbl\\useltbaln\\alntblind\\lytcalctblwd\\lyttblrtgr\\lnbrkrule\\nobrkwrptbl\\snaptogridincell\\allowfieldendsel\\wrppunct\n' +
  '\\asianbrkrule\\rsidroot1\\newtblstyruls\\nogrowautofit\\usenormstyforlist\\noindnmbrts\\felnbrelev\\nocxsptable\\indrlsweleven\\noafcnsttbl\\afelev\\utinl\\hwelev\\spltpgpar\\notcvasp\\notbrkcnstfrctbl\\notvatxbx\\krnprsnet\\cachedcolbal \\nouicompat \\fet0\n' +
  '{\\*\\wgrffmtfilter 2450}\\nofeaturethrottle1\\ilfomacatclnup0\n' +
  '{\\*\\ftnsep \\ltrpar \\pard\\plain \\ltrpar\\ql \\li0\\ri0\\widctlpar\\wrapdefault\\aspalpha\\aspnum\\faauto\\adjustright\\rin0\\lin0\\itap0 ' + CHP + ' {\\rtlch\\fcs1 \\af0 \\ltrch\\fcs0 \\chftnsep \\par }}\n' +
  '\\ltrpar \\sectd \\ltrsect\\linex0\\headery708\\footery708\\colsx708\\endnhere\\sectlinegrid360\\sectdefaultcl\\sftnbj\n' +
  '{\\header \\ltrpar ' + para('\\s19\\ql \\li0\\ri0\\widctlpar\\tqc\\tx4680\\tqr\\tx9360\\wrapdefault\\aspalpha\\aspnum\\faauto\\adjustright\\rin0\\lin0\\itap0', CHP, r('Office Engine\\tab Quarterly Report\\tab 2026', '\\cf3')) + '}\n' +
  '{\\footer \\ltrpar ' +
  para(
    '\\s20\\qc \\li0\\ri0\\widctlpar\\wrapdefault\\aspalpha\\aspnum\\faauto\\adjustright\\rin0\\lin0\\itap0',
    CHP,
    r('Page ') + '{\\field{\\*\\fldinst {\\rtlch\\fcs1 \\af0 \\ltrch\\fcs0  PAGE   \\\\* MERGEFORMAT }}{\\fldrslt {\\rtlch\\fcs1 \\af0 \\ltrch\\fcs0 \\noproof 1}}}' +
      r(' of ') + '{\\field{\\*\\fldinst {\\rtlch\\fcs1 \\af0 \\ltrch\\fcs0  NUMPAGES   \\\\* MERGEFORMAT }}{\\fldrslt {\\rtlch\\fcs1 \\af0 \\ltrch\\fcs0 \\noproof 2}}}',
  ) +
  '}\n' +
  para(TITLE, TITLEC, r('Quarterly Report')) +
  para(
    NORMAL,
    CHP,
    r('This document shows ') + r('bold', '\\b') + r(', ') + r('italic', '\\i') + r(', ') + r('underlined red', '\\ul\\cf4') + r(', ') +
      r('struck', '\\strike') + r(', ') + r('highlighted', '\\highlight6') + r(' and ') + r('large', '\\fs36') + r(' text, H') + r('2', '\\sub') +
      r('O, x') + r('2', '\\super') + r(', ') + r('Small Caps', '\\scaps') + r(", caf\\'e9 \\u8364\\'80 and \\ldblquote quotes\\rdblquote . Visit ") +
      '{\\field{\\*\\fldinst {\\rtlch\\fcs1 \\af0 \\ltrch\\fcs0 HYPERLINK "https://example.com" }}{\\fldrslt {\\rtlch\\fcs1 \\af0 \\ltrch\\fcs0 \\cs18\\ul\\cf7 example.com}}}' +
      r('.') + '{\\rtlch\\fcs1 \\af0 \\ltrch\\fcs0 \\cs22\\super \\chftn {\\footnote \\ltrpar \\pard\\plain \\ltrpar\\s21\\ql \\li0\\ri0\\widctlpar\\wrapdefault\\aspalpha\\aspnum\\faauto\\adjustright\\rin0\\lin0\\itap0 \\rtlch\\fcs1 \\af0\\afs20\\alang1025 \\ltrch\\fcs0 \\f1\\fs20\\lang1033\\langfe1033\\cgrid\\langnp1033\\langfenp1033 {\\rtlch\\fcs1 \\af0 \\ltrch\\fcs0 \\cs22\\super \\chftn }{\\rtlch\\fcs1 \\af0 \\ltrch\\fcs0  Figures are unaudited.}}}',
  ) +
  para(H1, H1C, '{\\*\\bkmkstart highlights}' + r('Highlights') + '{\\*\\bkmkend highlights}') +
  bullet('Revenue grew 24% year over year') +
  bullet('Two new regions launched') +
  bullet('North and South', 1) +
  bullet('Customer churn fell to 3.1%') +
  para(H1, H1C, r('Next steps')) +
  numbered('1.', 'Hire the platform team') +
  numbered('2.', 'Ship the viewer engine') +
  numbered('a.', 'Documents, workbooks and presentations', 1) +
  numbered('3.', 'Expand to APAC') +
  para(H2, H2C, r('Figures')) +
  para(NORMAL.replace('\\ql', '\\qc'), CHP, picture) +
  para(NORMAL.replace('\\ql', '\\qc'), CHP, r('Figure 1 \\emdash  generated gradient image', '\\i')) +
  para(H2, H2C, r('Table')) +
  row(['Region', 'Q1', 'Q2', 'Total'], true) +
  row(['North', '120', '140', '260']) +
  row(['South', '98', '110', '208']) +
  row(['East', '132', '150', '282']) +
  row(['Merged total (3 columns)', '', '', '750'], false, true) +
  para(NORMAL + '\\tqr\\tldot\\tx9000', CHP, r('Grand total\\tab 750')) +
  para(
    NORMAL + '\\brdrt\\brdrs\\brdrw10\\brsp20 \\brdrl\\brdrs\\brdrw10\\brsp80 \\brdrb\\brdrs\\brdrw10\\brsp20 \\brdrr\\brdrs\\brdrw10\\brsp80 \\cbpat9',
    CHP,
    r('Note: ', '\\b') + r('the totals include all regions. See ') +
      '{\\field{\\*\\fldinst {\\rtlch\\fcs1 \\af0 \\ltrch\\fcs0 HYPERLINK \\\\l "highlights" }}{\\fldrslt {\\rtlch\\fcs1 \\af0 \\ltrch\\fcs0 \\cs18\\ul\\cf7 Highlights}}}' +
      r('.'),
  ) +
  para(NORMAL + '\\pagebb', CHP, r('Appendix', '\\b\\fs28')) +
  para(
    NORMAL.replace('\\ql', '\\qj').replace('\\li0', '\\li864').replace('\\lin0', '\\lin864'),
    CHP,
    r('Enim do nostrud reprehenderit sit amet, commodo adipiscing veniam, aute sit ex ut dolor consectetur ullamco exercitation amet, et consectetur consequat.'),
  ) +
  '}\n'

const out = `${root}/test/fixtures/legacy/sample.rtf`
await Bun.write(out, rtf)
console.log('wrote', out, rtf.length, 'bytes')
