/**
 * Crafted files for the security tests (test/security.test.ts, e2e/security.spec.ts), built from the samples:
 *
 *  - `probe*`: every color of the file says "fff url(https://probe.invalid/<id>.png)". Given to CSS as it is, that
 *    is a valid `background` and the browser requests the address when the file is opened.
 *  - `inherited*`: names a plain object answers with what it inherits ("constructor", "__proto__") where the file
 *    names a preset color, a border style, an alignment…
 */
import { BlobSource } from '../src/core/source'
import { ZipArchive } from '../src/core/zip'
import { rebuildZip } from '../src/core/zipwrite'

const SAMPLES = `${import.meta.dir}/../playground/public/samples`

/** never resolves: a request to it is a leak, and reaches nobody */
export const PROBE_HOST = 'probe.invalid'

let seq = 0
/** a color value that is also a url() to fetch; `id` tells the probes apart in a failure */
export const probe = (id: string) => `fff url(https://${PROBE_HOST}/${id}-${++seq}.png)`

type Edit = (xml: string) => string

async function rebuilt(sample: string, edits: Record<string, Edit>): Promise<Blob> {
  const zip = await ZipArchive.open(new BlobSource(new Blob([await Bun.file(`${SAMPLES}/${sample}`).arrayBuffer()])))
  const changes = new Map<string, string>()
  for (const [path, edit] of Object.entries(edits)) {
    const before = await zip.text(path)
    const after = edit(before)
    if (after === before) throw new Error(`probe: nothing changed in ${sample} ${path}`)
    changes.set(path, after)
  }
  return rebuildZip(zip, changes)
}

/** `xml` with `add` put before the first `anchor` */
function before(xml: string, anchor: string, add: string): string {
  const i = xml.indexOf(anchor)
  if (i < 0) throw new Error(`probe: no ${anchor}`)
  return xml.slice(0, i) + add + xml.slice(i)
}

/** `xml` with `add` put after the first `anchor` */
function after(xml: string, anchor: string, add: string): string {
  const i = xml.indexOf(anchor)
  if (i < 0) throw new Error(`probe: no ${anchor}`)
  return xml.slice(0, i + anchor.length) + add + xml.slice(i + anchor.length)
}

const drawingColors = (id: string): Edit => (xml) =>
  xml.replace(/(<a:srgbClr val=")[0-9A-Fa-f]{6}(")/g, (_, a, b) => a + probe(id) + b).replace(/(lastClr=")[0-9A-Fa-f]{6}(")/g, (_, a, b) => a + probe(id + '-sys') + b)

const solid = (id: string) => `<a:solidFill><a:srgbClr val="${probe(id)}"/></a:solidFill>`

function shape(n: number, y: number, geom: string, fill: string, extra = ''): string {
  return (
    `<p:sp><p:nvSpPr><p:cNvPr id="${900 + n}" name="probe ${n}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
    `<p:spPr><a:xfrm><a:off x="200000" y="${y}"/><a:ext cx="2400000" cy="600000"/></a:xfrm><a:prstGeom prst="${geom}"><a:avLst/></a:prstGeom>${fill}${extra}</p:spPr>` +
    `<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:buClr><a:srgbClr val="${probe('pptx-bullet')}"/></a:buClr><a:buChar char="•"/></a:pPr>` +
    `<a:r><a:rPr lang="en-US">${solid('pptx-text')}<a:highlight><a:srgbClr val="${probe('pptx-highlight')}"/></a:highlight></a:rPr><a:t>probe ${n}</a:t></a:r></a:p></p:txBody></p:sp>`
  )
}

/** slide background, shapes (solid, gradient, outlined, with arrow heads), text, bullets, a table, the theme */
export function probePptx(): Promise<Blob> {
  const line = (id: string, ends = '') => `<a:ln w="25400">${solid(id)}${ends}</a:ln>`
  const grad = `<a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="${probe('pptx-grad')}"/></a:gs><a:gs pos="100000"><a:srgbClr val="${probe('pptx-grad')}"/></a:gs></a:gsLst><a:lin ang="5400000"/></a:gradFill>`
  const table =
    `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="950" name="probe table"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>` +
    `<p:xfrm><a:off x="3000000" y="200000"/><a:ext cx="3000000" cy="400000"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">` +
    `<a:tbl><a:tblPr/><a:tblGrid><a:gridCol w="3000000"/></a:tblGrid><a:tr h="400000"><a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>cell</a:t></a:r></a:p></a:txBody>` +
    `<a:tcPr><a:lnL w="12700">${solid('pptx-cell-line')}</a:lnL><a:lnB w="12700">${solid('pptx-cell-line')}</a:lnB>${solid('pptx-cell')}</a:tcPr></a:tc></a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame>`
  const shapes =
    shape(1, 200000, 'rect', solid('pptx-rect')) +
    shape(2, 1000000, 'ellipse', solid('pptx-ellipse'), line('pptx-line')) +
    shape(3, 1800000, 'roundRect', grad) +
    shape(4, 2600000, 'line', '', line('pptx-arrow', '<a:headEnd type="triangle"/><a:tailEnd type="triangle"/>')) +
    table
  return rebuilt('sample.pptx', {
    'ppt/slides/slide1.xml': (xml) => before(after(xml, '<p:cSld>', `<p:bg><p:bgPr>${solid('pptx-bg')}<a:effectLst/></p:bgPr></p:bg>`), '</p:spTree>', shapes),
    'ppt/slides/slide3.xml': drawingColors('pptx-slide3'),
    'ppt/theme/theme1.xml': drawingColors('pptx-theme'),
  })
}

/** fonts, fills (solid and gradient), borders, the workbook's own palette, the theme */
export function probeXlsx(): Promise<Blob> {
  return rebuilt('sample.xlsx', {
    'xl/styles.xml': (xml) => {
      const gradient = `<fill><gradientFill degree="90"><stop position="0"><color rgb="${probe('xlsx-grad')}"/></stop><stop position="1"><color rgb="${probe('xlsx-grad')}"/></stop></gradientFill></fill>`
      const out = xml.replace('<fill><patternFill patternType="solid"><fgColor rgb="00D9E1F2"/></patternFill></fill>', gradient)
      if (out === xml) throw new Error('probe: no fill to make a gradient of')
      return out.replace(/(\brgb=")[0-9A-Fa-f]{8}(")/g, (_, a, b) => a + probe('xlsx-style') + b)
    },
    'xl/theme/theme1.xml': drawingColors('xlsx-theme'),
  })
}

const WPS = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape'

/** a text box (fill and outline), run colors, shading, borders, the theme */
export function probeDocx(): Promise<Blob> {
  const box =
    `<w:p><w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><wp:extent cx="1905000" cy="952500"/><wp:docPr id="991" name="probe box"/>` +
    `<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="${WPS}"><wps:wsp xmlns:wps="${WPS}">` +
    `<wps:spPr>${solid('docx-box')}<a:ln>${solid('docx-box-line')}</a:ln></wps:spPr>` +
    `<wps:txbx><w:txbxContent><w:p><w:r><w:t>probe box</w:t></w:r></w:p></w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`
  const wordColors: Edit = (xml) =>
    xml
      .replace(/(w:fill=")[0-9A-Fa-f]{6}(")/g, (_, a, b) => a + probe('docx-shd') + b)
      .replace(/(<w:color w:val=")[0-9A-Fa-f]{6}(")/g, (_, a, b) => a + probe('docx-color') + b)
      .replace(/(w:color=")[0-9A-Fa-f]{6}(")/g, (_, a, b) => a + probe('docx-border') + b)
  return rebuilt('sample.docx', {
    'word/document.xml': (xml) => wordColors(after(xml, '<w:body>', box)),
    'word/styles.xml': wordColors,
    'word/theme/theme1.xml': drawingColors('docx-theme'),
  })
}

/** names every object inherits, where a file names one of a known few */
export const INHERITED = ['constructor', '__proto__', 'toString', 'hasOwnProperty', 'valueOf']

export function inheritedPptx(name: string): Promise<Blob> {
  const sp = (n: number, fill: string, ln = '', pPr = '') =>
    `<p:sp><p:nvSpPr><p:cNvPr id="${960 + n}" name="k ${n}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="200000" y="${n * 700000}"/><a:ext cx="2000000" cy="500000"/></a:xfrm>` +
    `<a:prstGeom prst="${name}"><a:avLst><a:gd name="${name}" fmla="val 5"/></a:avLst></a:prstGeom>${fill}${ln}</p:spPr>` +
    `<p:txBody><a:bodyPr anchor="${name}" vert="${name}"/><a:lstStyle/><a:p>${pPr}<a:r><a:rPr><a:latin typeface="${name}"/><a:solidFill><a:schemeClr val="${name}"/></a:solidFill></a:rPr><a:t>k ${n}</a:t></a:r></a:p></p:txBody></p:sp>`
  const shapes =
    sp(1, `<a:solidFill><a:prstClr val="${name}"/></a:solidFill>`, `<a:ln w="12700"><a:solidFill><a:schemeClr val="${name}"/></a:solidFill><a:prstDash val="${name}"/><a:headEnd type="${name}"/></a:ln>`, `<a:pPr algn="${name}"/>`) +
    sp(2, `<a:solidFill><a:schemeClr val="${name}"/></a:solidFill>`) +
    sp(3, `<a:solidFill><a:sysClr val="${name}" lastClr="${name}"/></a:solidFill>`)
  return rebuilt('sample.pptx', {
    'ppt/slides/slide1.xml': (xml) =>
      before(
        after(xml, '<p:cSld>', `<p:bg><p:bgPr><a:solidFill><a:prstClr val="${name}"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>`),
        '</p:spTree>',
        shapes,
      ),
    // the slide's own color map sends every name to an inherited one
    'ppt/slides/slide2.xml': (xml) =>
      before(
        xml.replace(/<p:clrMapOvr>[\s\S]*?<\/p:clrMapOvr>/, ''),
        '</p:sld>',
        `<p:clrMapOvr><a:overrideClrMapping bg1="${name}" tx1="${name}" bg2="${name}" tx2="${name}" accent1="${name}" accent2="${name}" accent3="${name}" accent4="${name}" accent5="${name}" accent6="${name}" hlink="${name}" folHlink="${name}"/></p:clrMapOvr>`,
      ),
  })
}

export function inheritedXlsx(name: string): Promise<Blob> {
  return rebuilt('sample.xlsx', {
    'xl/styles.xml': (xml) =>
      xml
        .replace(/(<bottom style=")[a-zA-Z]+(")/g, `$1${name}$2`)
        .replace(/(<name val=")[^"]*(")/g, `$1${name}$2`)
        .replace(/formatCode="yyyy-mm-dd"/, `formatCode="[${name}]0.00"`)
        .replace(/<color theme="1"\/>/, `<color theme="${name}"/>`)
        .replace(/horizontal="center"/g, `horizontal="${name}" vertical="${name}"`),
  })
}

export function inheritedDocx(name: string): Promise<Blob> {
  const para =
    `<w:p><w:pPr><w:pBdr><w:top w:val="${name}" w:sz="4" w:space="1" w:color="${name}"/></w:pBdr><w:jc w:val="${name}"/></w:pPr>` +
    `<w:r><w:rPr><w:rFonts w:ascii="${name}" w:hAnsi="${name}"/><w:highlight w:val="${name}"/><w:u w:val="${name}"/><w:vertAlign w:val="${name}"/><w:color w:val="${name}"/></w:rPr><w:t>inherited ${name}</w:t></w:r></w:p>`
  return rebuilt('sample.docx', {
    'word/document.xml': (xml) => after(xml, '<w:body>', para).replace(/<w:tblStyle w:val="[^"]*"\/>/, `<w:tblStyle w:val="${name}"/>`),
    'word/styles.xml': (xml) => before(xml, '</w:styles>', `<w:style w:type="table" w:styleId="${name}"><w:name w:val="${name}"/><w:tblStylePr w:type="${name}"><w:tcPr><w:shd w:val="clear" w:fill="FF0000"/></w:tcPr></w:tblStylePr></w:style>`),
  })
}

/** the sample document, its link pointing at a script instead of a web page */
export function scriptLinkDocx(): Promise<Blob> {
  return rebuilt('sample.docx', {
    'word/_rels/document.xml.rels': (xml) => xml.replace('Target="https://example.com"', 'Target="javascript:window.__ran=1"'),
  })
}
