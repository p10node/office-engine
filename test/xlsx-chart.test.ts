import { describe, expect, test } from 'bun:test'
import { Engine } from '../src/core/worker/engine'
import { rowsToXlsx } from '../src/core/convert'
import { ZipArchive } from '../src/core/zip'
import { BlobSource } from '../src/core/source'
import { rebuildZip } from '../src/core/zipwrite'
import type { SheetDrawing } from '../src/core/xlsx/drawings'
import type { ChartModel } from '../src/core/xlsx/chart'
import { niceScale, renderChart } from '../src/core/xlsx/chart-svg'

const RELS = 'http://schemas.openxmlformats.org/package/2006/relationships'
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const NS = 'xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"'
const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'

const ROWS = [
  ['Quarter', 'Sales', 'Cost'],
  ['Q1', '10', '4'],
  ['Q2', '20', '8'],
  ['Q3', '15', '9'],
]

/** a numeric data source: a reference and a cache whose values differ from the cells (so the test tells them apart) */
const numRef = (f: string, cache: number[], fmt = 'General') =>
  `<c:numRef><c:f>${f}</c:f><c:numCache><c:formatCode>${fmt}</c:formatCode><c:ptCount val="${cache.length}"/>${cache.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join('')}</c:numCache></c:numRef>`
const strRef = (f: string, cache: string[]) =>
  `<c:strRef><c:f>${f}</c:f><c:strCache><c:ptCount val="${cache.length}"/>${cache.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join('')}</c:strCache></c:strRef>`
const ser = (i: number, name: string, val: string, extra = '') =>
  `<c:ser><c:idx val="${i}"/><c:order val="${i}"/><c:tx>${strRef(`Data!$${name}$1`, ['cached name'])}</c:tx>${extra}<c:cat>${strRef('Data!$A$2:$A$4', ['x', 'y', 'z'])}</c:cat><c:val>${val}</c:val></c:ser>`
const AXES = '<c:catAx><c:axId val="1"/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx><c:valAx><c:axId val="2"/><c:axPos val="l"/><c:majorGridlines/><c:crossAx val="1"/></c:valAx>'

/** a workbook whose sheet "Data" holds ROWS and one chart (its `plotArea` body given); `hideRow` hides a row (1-based) */
async function withChart(plot: string, opts: { hideRow?: number; title?: string; titleXml?: string; style?: number; anchors?: number } = {}): Promise<Blob> {
  const base = await rowsToXlsx(ROWS, 'Data')
  const zip = await ZipArchive.open(new BlobSource(base))
  const files = new Map<string, string>()
  let sheet = (await zip.text('xl/worksheets/sheet1.xml')).replace('</worksheet>', '<drawing r:id="rId1"/></worksheet>')
  if (!sheet.includes('xmlns:r=')) sheet = sheet.replace('<worksheet ', `<worksheet xmlns:r="${R}" `)
  if (opts.hideRow) sheet = sheet.replace(`<row r="${opts.hideRow}">`, `<row r="${opts.hideRow}" hidden="1">`)
  files.set('xl/worksheets/sheet1.xml', sheet)
  files.set('xl/worksheets/_rels/sheet1.xml.rels', `${HEAD}<Relationships xmlns="${RELS}"><Relationship Id="rId1" Type="${R}/drawing" Target="../drawings/drawing1.xml"/></Relationships>`)
  files.set('xl/drawings/_rels/drawing1.xml.rels', `${HEAD}<Relationships xmlns="${RELS}"><Relationship Id="rId1" Type="${R}/chart" Target="../charts/chart1.xml"/></Relationships>`)
  files.set(
    'xl/drawings/drawing1.xml',
    `${HEAD}<xdr:wsDr ${NS}>${Array.from({ length: opts.anchors ?? 1 }, (_, k) => `<xdr:twoCellAnchor><xdr:from><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${1 + k * 20}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>10</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${16 + k * 20}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id="${2 + k}" name="Chart ${1 + k}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId1"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>`).join('')}</xdr:wsDr>`,
  )
  const title = opts.titleXml ?? (opts.title ? `<c:title><c:tx><c:rich><a:p><a:r><a:t>${opts.title}</a:t></a:r></a:p></c:rich></c:tx></c:title>` : '')
  files.set('xl/charts/chart1.xml', `${HEAD}<c:chartSpace ${NS}>${opts.style ? `<c:style val="${opts.style}"/>` : ''}<c:chart>${title}<c:plotArea><c:layout/>${plot}</c:plotArea><c:legend><c:legendPos val="b"/></c:legend><c:plotVisOnly val="1"/></c:chart></c:chartSpace>`)
  const ct = (await zip.text('[Content_Types].xml')).replace(
    '</Types>',
    '<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/><Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/></Types>',
  )
  files.set('[Content_Types].xml', ct)
  return rebuildZip(zip, files, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
}

async function modelOf(book: Blob, engine = new Engine(() => {})): Promise<{ model: ChartModel | undefined; engine: Engine; drawing: SheetDrawing }> {
  if (!(engine as unknown as { xlsx?: unknown }).xlsx) await engine.open(book, {})
  const [drawing] = (await engine.call('xlsx.drawings', [0])) as SheetDrawing[]
  return { model: drawing?.model, engine, drawing }
}

const CLUSTERED = `<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>${ser(0, 'B', numRef('Data!$B$2:$B$4', [1, 1, 1]))}${ser(1, 'C', numRef('Data!$C$2:$C$4', [2, 2, 2]))}<c:gapWidth val="150"/><c:axId val="1"/><c:axId val="2"/></c:barChart>${AXES}`

describe('xlsx charts: model', () => {
  test('series read from the cells, not the saved cache; names, categories, theme colors', async () => {
    const { model } = await modelOf(await withChart(CLUSTERED, { title: 'Revenue' }))
    expect(model).toBeDefined()
    const g = model!.groups[0]
    expect(g).toMatchObject({ type: 'bar', dir: 'col', grouping: 'clustered', gapWidth: 150, axIds: ['1', '2'] })
    expect(g.series.map((s) => s.name)).toEqual(['Sales', 'Cost'])
    expect(g.series[0].cats).toEqual(['Q1', 'Q2', 'Q3'])
    expect(g.series[0].values).toEqual([10, 20, 15])
    expect(g.series[1].values).toEqual([4, 8, 9])
    // no spPr: the theme's first two accents
    expect(g.series[0].fill).toMatch(/^#[0-9a-f]{6}$/i)
    expect(g.series[0].fill).not.toBe(g.series[1].fill)
    expect(model!.title?.text).toBe('Revenue')
    expect(model!.legend?.pos).toBe('b')
    expect(model!.axes.map((a) => [a.id, a.kind, a.gridlines])).toEqual([['1', 'cat', false], ['2', 'val', true]])
  })

  test('a defined name or another workbook: the saved cache stands in', async () => {
    const plot = `<c:lineChart><c:grouping val="standard"/>${ser(0, 'B', numRef('SalesRange', [7, 8, 9]))}${ser(1, 'C', numRef('[1]Data!$C$2:$C$4', [3, 2, 1]))}<c:axId val="1"/><c:axId val="2"/></c:lineChart>${AXES}`
    const { model } = await modelOf(await withChart(plot))
    expect(model!.groups[0].series.map((s) => s.values)).toEqual([[7, 8, 9], [3, 2, 1]])
    expect(model!.groups[0].series[0].marker).not.toBeNull()
  })

  test('hidden rows are left out of a chart of visible cells', async () => {
    const { model } = await modelOf(await withChart(CLUSTERED, { hideRow: 3 }))
    expect(model!.groups[0].series[0].values).toEqual([10, 15])
    expect(model!.groups[0].series[0].cats).toEqual(['Q1', 'Q3'])
  })

  test('an edited cell is in the chart read after the edit', async () => {
    const book = await withChart(CLUSTERED)
    const events: { event: string; data: { sheet?: number; done?: boolean } }[] = []
    const engine = new Engine((event, data) => events.push({ event, data: data as { sheet?: number; done?: boolean } }))
    await engine.open(book, {})
    await engine.call('xlsx.load', [0])
    while (!events.some((e) => e.event === 'xlsx:progress' && e.data.sheet === 0 && e.data.done)) await new Promise((r) => setTimeout(r, 5))
    await engine.call('xlsx.set', [0, [{ r: 2, c: 1, input: '42' }]])
    const { model } = await modelOf(book, engine)
    expect(model!.groups[0].series[0].values).toEqual([10, 42, 15])
  })

  test('pie: one series, a color per slice; a single series titles the chart', async () => {
    const plot = `<c:pieChart><c:varyColors val="1"/>${ser(0, 'B', numRef('Data!$B$2:$B$4', [1, 1, 1]), '<c:dPt><c:idx val="1"/><c:explosion val="20"/></c:dPt>')}<c:firstSliceAng val="90"/></c:pieChart>`
    const { model } = await modelOf(await withChart(plot))
    const s = model!.groups[0].series[0]
    expect(model!.groups[0].type).toBe('pie')
    expect(model!.groups[0].firstSliceAng).toBe(90)
    expect(Object.keys(s.points)).toEqual(['0', '1', '2'])
    expect(new Set(Object.values(s.points).map((p) => p.color)).size).toBe(3)
    expect(s.points[1].explosion).toBe(20)
    expect(model!.title?.text).toBe('Sales')
  })

  test('radar: read and drawn, a spoke per category, markers for style "marker"', async () => {
    const plot = `<c:radarChart><c:radarStyle val="marker"/>${ser(0, 'B', numRef('Data!$B$2:$B$4', [1, 1, 1]))}<c:axId val="1"/><c:axId val="2"/></c:radarChart>${AXES}`
    const { model, drawing } = await modelOf(await withChart(plot))
    expect(drawing.chart).toBe('radar')
    expect(model!.groups[0]).toMatchObject({ type: 'radar', radarStyle: 'marker' })
    expect(model!.groups[0].series[0].marker).not.toBeNull()
    const svg = renderChart(model!, 400, 300)
    for (const t of ['Q1', 'Q2', 'Q3']) expect(svg).toContain(`>${t}<`)
    expect(svg).toContain('<polygon')
    expect(svg).not.toContain('NaN')
  })

  test('bubble: sizes read from the cells, larger values drawn as larger bubbles', async () => {
    const bub = `<c:ser><c:idx val="0"/><c:order val="0"/><c:tx>${strRef('Data!$B$1', ['x'])}</c:tx><c:xVal>${numRef('Data!$B$2:$B$4', [0, 0, 0])}</c:xVal><c:yVal>${numRef('Data!$C$2:$C$4', [0, 0, 0])}</c:yVal><c:bubbleSize>${numRef('Data!$B$2:$B$4', [1, 1, 1])}</c:bubbleSize></c:ser>`
    const plot = `<c:bubbleChart><c:varyColors val="0"/>${bub}<c:bubbleScale val="100"/><c:axId val="1"/><c:axId val="2"/></c:bubbleChart><c:valAx><c:axId val="1"/><c:axPos val="b"/><c:crossAx val="2"/></c:valAx><c:valAx><c:axId val="2"/><c:axPos val="l"/><c:majorGridlines/><c:crossAx val="1"/></c:valAx>`
    const { model } = await modelOf(await withChart(plot))
    const s = model!.groups[0].series[0]
    expect(s.sizes).toEqual([10, 20, 15])
    expect(s.xs).toEqual([10, 20, 15])
    const svg = renderChart(model!, 400, 300)
    const radii = [...svg.matchAll(/<circle [^>]*r="([\d.]+)"[^>]*fill="(?!none)/g)].map((m) => Number(m[1]))
    expect(radii.length).toBe(3)
    // drawn largest first: 20, 15, 10
    expect(radii[0]).toBeGreaterThan(radii[1])
    expect(radii[1]).toBeGreaterThan(radii[2])
  })

  test('stock: high-low lines and up-down bars between the first and last series', async () => {
    const plot = `<c:stockChart>${ser(0, 'B', numRef('Data!$B$2:$B$4', [1, 1, 1]))}${ser(1, 'C', numRef('Data!$C$2:$C$4', [1, 1, 1]))}<c:hiLowLines><c:spPr><a:ln w="12700"><a:solidFill><a:srgbClr val="112233"/></a:solidFill></a:ln></c:spPr></c:hiLowLines><c:upDownBars><c:gapWidth val="100"/><c:upBars><c:spPr><a:solidFill><a:srgbClr val="00AA00"/></a:solidFill></c:spPr></c:upBars><c:downBars><c:spPr><a:solidFill><a:srgbClr val="AA0000"/></a:solidFill></c:spPr></c:downBars></c:upDownBars><c:axId val="1"/><c:axId val="2"/></c:stockChart>${AXES}`
    const { model } = await modelOf(await withChart(plot))
    expect(model!.groups[0].hiLow).toMatchObject({ color: '#112233' })
    const svg = renderChart(model!, 400, 300)
    // every quarter falls (Sales to Cost): three falling bars, three high-low lines
    expect(svg.match(/fill="#aa0000"/gi)?.length).toBe(3)
    expect(svg.match(/stroke="#112233"/gi)?.length).toBe(3)
  })

  test('surface: value bands in the legend and the plot filled by band', async () => {
    const plot = `<c:surfaceChart><c:wireframe val="0"/>${ser(0, 'B', numRef('Data!$B$2:$B$4', [1, 1, 1]))}${ser(1, 'C', numRef('Data!$C$2:$C$4', [1, 1, 1]))}<c:axId val="1"/><c:axId val="2"/><c:axId val="3"/></c:surfaceChart>${AXES}<c:serAx><c:axId val="3"/><c:axPos val="b"/><c:crossAx val="2"/></c:serAx>`
    const { model } = await modelOf(await withChart(plot))
    expect(model!.groups[0].type).toBe('surface')
    expect(model!.axes.some((a) => a.kind === 'ser')).toBe(true)
    const svg = renderChart(model!, 400, 300)
    expect(svg).toContain('>0-5<')
    expect(svg).toContain('>20-25<')
    const colors = new Set([...svg.matchAll(/<rect [^>]*fill="(#[0-9a-f]{6})"/gi)].map((m) => m[1].toLowerCase()))
    expect(colors.size).toBeGreaterThanOrEqual(4)
    expect(svg).not.toContain('NaN')
  })
})

describe('xlsx charts: looks', () => {
  const colored = (spPr: string) => CLUSTERED.replace('<c:tx>', `<c:spPr>${spPr}</c:spPr><c:tx>`).replace(`${ser(0, 'B', numRef('Data!$B$2:$B$4', [1, 1, 1]))}`, ser(0, 'B', numRef('Data!$B$2:$B$4', [1, 1, 1]), `<c:spPr>${spPr}</c:spPr>`))

  test('a gradient fill is drawn as a gradient, a pattern fill as a pattern', async () => {
    const grad = '<a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="1F4E79"/></a:gs><a:gs pos="100000"><a:srgbClr val="9DC3E6"/></a:gs></a:gsLst><a:lin ang="5400000"/></a:gradFill>'
    const { model } = await modelOf(await withChart(colored(grad)))
    const s = model!.groups[0].series[0]
    expect(s.fill?.toLowerCase()).toBe('#1f4e79')
    expect(s.look).toMatchObject({ kind: 'gradient', angle: 90, radial: false })
    const svg = renderChart(model!, 400, 300)
    expect(svg).toMatch(/<linearGradient id="de-fill\d+"[^>]*>.*stop-color="#9DC3E6"/i)
    expect(svg).toMatch(/<rect [^>]*fill="url\(#de-fill\d+\)"/)

    const patt = '<a:pattFill prst="dkDnDiag"><a:fgClr><a:srgbClr val="C55A11"/></a:fgClr><a:bgClr><a:srgbClr val="FFFFFF"/></a:bgClr></a:pattFill>'
    const { model: pm } = await modelOf(await withChart(colored(patt)))
    expect(pm!.groups[0].series[0].look).toMatchObject({ kind: 'pattern', preset: 'dkDnDiag', fg: '#c55a11', bg: '#ffffff' })
    expect(renderChart(pm!, 400, 300)).toMatch(/<pattern id="de-fill\d+"[^>]*>.*stroke="#C55A11"/i)
  })

  test('axis and legend text, gridlines and the plot area in the colors the file gives', async () => {
    const tx = (c: string) => `<c:txPr><a:bodyPr/><a:p><a:pPr><a:defRPr sz="900"><a:solidFill><a:srgbClr val="${c}"/></a:solidFill></a:defRPr></a:pPr></a:p></c:txPr>`
    const axes = `<c:catAx><c:axId val="1"/><c:axPos val="b"/>${tx('AA0001')}<c:crossAx val="2"/></c:catAx><c:valAx><c:axId val="2"/><c:axPos val="l"/><c:majorGridlines><c:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="4B5563"/></a:solidFill><a:prstDash val="dash"/></a:ln></c:spPr></c:majorGridlines>${tx('AA0002')}<c:crossAx val="1"/></c:valAx><c:spPr><a:solidFill><a:srgbClr val="374151"/></a:solidFill></c:spPr>`
    const plot = CLUSTERED.replace(AXES, axes)
    const book = await withChart(plot)
    const { model } = await modelOf(book)
    model!.legend!.text = { color: '#aa0003' }
    const svg = renderChart(model!, 400, 300)
    expect(svg).toMatch(/fill="#AA0001"[^>]*>Q1</i)
    expect(svg).toMatch(/fill="#AA0002"[^>]*>25</i)
    expect(svg).toMatch(/fill="#aa0003"[^>]*>Sales</)
    expect(svg).toMatch(/<line [^>]*stroke="#4B5563"[^>]*stroke-dasharray/i)
    expect(model!.plotFill?.toLowerCase()).toBe('#374151')
    expect(svg).toContain('fill="#374151"')
  })

  test('the theme gives the default grays; a dark built-in style a dark chart with light text', async () => {
    const { model } = await modelOf(await withChart(CLUSTERED))
    expect(model!.gridColor.toLowerCase()).toBe('#d9d9d9')
    expect(model!.textColor.toLowerCase()).toBe('#595959')
    const dark = await withChart(CLUSTERED, { style: 42 })
    const { model: dm } = await modelOf(dark)
    expect(dm!.bg).toBeTruthy()
    expect(dm!.textColor.toLowerCase()).toBe('#ffffff')
    // style 42 is the second of the dark row: the theme's accents in turn
    expect(dm!.groups[0].series.map((s) => s.fill)).toEqual(model!.groups[0].series.map((s) => s.fill))
    // style 3, one accent's shades: two series, two shades of accent 1
    const { model: mono } = await modelOf(await withChart(CLUSTERED, { style: 3 }))
    const [a, b] = mono!.groups[0].series.map((s) => s.fill)
    expect(a).not.toBe(b)
  })
})

describe('xlsx charts: drawing', () => {
  test('axis bounds as Excel picks them', () => {
    expect(niceScale(4, 20, { ticks: 6 })).toMatchObject({ min: 0, max: 25, ticks: [0, 5, 10, 15, 20, 25] })
    expect(niceScale(95, 100, { ticks: 5 }).min).toBeGreaterThan(80)
    expect(niceScale(-3, 8, { ticks: 6 }).ticks[0]).toBeLessThan(-2)
    expect(niceScale(0, 1, { min: 0, max: 1, major: 0.25 }).ticks).toEqual([0, 0.25, 0.5, 0.75, 1])
  })

  test('a clustered column chart: a bar per value, the categories, gridlines and the legend', async () => {
    const { model } = await modelOf(await withChart(CLUSTERED, { title: 'Revenue' }))
    const svg = renderChart(model!, 480, 300)
    expect(svg.startsWith('<svg')).toBe(true)
    // nothing that runs or loads: url() only names the picture's own clip path, gradients and patterns
    expect(svg).not.toMatch(/<script|href=|url\((?!#de-)/i)
    const fills = new Set(model!.groups[0].series.map((s) => s.fill))
    const bars = [...svg.matchAll(/<rect [^>]*fill="([^"]+)"/g)].filter((m) => fills.has(m[1]))
    // 3 categories × 2 series, plus a legend swatch each
    expect(bars.length).toBe(8)
    for (const t of ['Revenue', 'Q1', 'Q2', 'Q3', 'Sales', 'Cost', '25']) expect(svg).toContain(`>${t}<`)
  })

  test('stacked bars, lines and pies draw without errors and with their values', async () => {
    const stacked = CLUSTERED.replace('<c:grouping val="clustered"/>', '<c:grouping val="stacked"/><c:overlap val="100"/>').replace('<c:barDir val="col"/>', '<c:barDir val="bar"/>')
    const line = `<c:lineChart><c:grouping val="standard"/>${ser(0, 'B', numRef('Data!$B$2:$B$4', [1, 1, 1]), '<c:dLbls><c:showVal val="1"/></c:dLbls>')}<c:marker val="1"/><c:axId val="1"/><c:axId val="2"/></c:lineChart>${AXES}`
    const doughnut = `<c:doughnutChart><c:varyColors val="1"/>${ser(0, 'B', numRef('Data!$B$2:$B$4', [1, 1, 1]), '<c:dLbls><c:showPercent val="1"/></c:dLbls>')}<c:holeSize val="60"/></c:doughnutChart>`
    for (const plot of [stacked, line, doughnut]) {
      const { model } = await modelOf(await withChart(plot))
      const svg = renderChart(model!, 400, 260)
      expect(svg).toMatch(/^<svg[\s\S]*<\/svg>$/)
      expect(svg).not.toContain('NaN')
    }
    const { model: lm } = await modelOf(await withChart(line))
    expect(renderChart(lm!, 400, 260)).toContain('>20<')
    const { model: dm } = await modelOf(await withChart(doughnut))
    expect(renderChart(dm!, 400, 260)).toContain('>44%<')
  })

  test('100% stacked: the axis ends at 100%, with no step above', async () => {
    const pct = CLUSTERED.replace('<c:grouping val="clustered"/>', '<c:grouping val="percentStacked"/><c:overlap val="100"/>')
    const { model } = await modelOf(await withChart(pct))
    const svg = renderChart(model!, 400, 300)
    expect(svg).toContain('>100%<')
    expect(svg).not.toContain('>120%<')
    expect(svg).not.toContain('>110%<')
  })

  test('a title from a cell: the cell text, not the automatic one', async () => {
    const titleXml = `<c:title><c:tx>${strRef('Data!$C$1', ['cached title'])}</c:tx></c:title>`
    const { model } = await modelOf(await withChart(CLUSTERED, { titleXml }))
    expect(model!.title?.text).toBe('Cost')
  })

  test('pie: a negative value is a slice of its magnitude, its label keeps the sign', async () => {
    const plot = `<c:pieChart><c:varyColors val="1"/>${ser(0, 'B', numRef('Data!$B$2:$B$4', [1, 1, 1]))}</c:pieChart>`
    const { model } = await modelOf(await withChart(plot))
    const s = model!.groups[0].series[0]
    s.values = [-10, 20, 30]
    s.labels = { value: true, percent: false, category: false, series: false }
    const svg = renderChart(model!, 400, 260)
    expect(svg.match(/<path d="M[^"]* A/g)?.length).toBe(3)
    expect(svg).toContain('>-10<')
  })

  const SCATTER = `<c:scatterChart><c:scatterStyle val="lineMarker"/><c:ser><c:idx val="0"/><c:order val="0"/><c:xVal>${numRef('Data!$B$2:$B$4', [1, 2, 3])}</c:xVal><c:yVal>${numRef('Data!$C$2:$C$4', [1, 2, 3])}</c:yVal></c:ser><c:axId val="1"/><c:axId val="2"/></c:scatterChart><c:valAx><c:axId val="1"/><c:axPos val="b"/><c:crossAx val="2"/></c:valAx><c:valAx><c:axId val="2"/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>`
  /** the x position of each value label in a rendered chart, by its text */
  const labelX = (svg: string) => Object.fromEntries([...svg.matchAll(/<text x="([\d.]+)"[^>]*>(\d+)</g)].map((m) => [m[2], Number(m[1])]))

  test('scatter: a point missing its X is not drawn; labels stay with their points', async () => {
    const { model } = await modelOf(await withChart(SCATTER))
    const s = model!.groups[0].series[0]
    s.labels = { value: true, percent: false, category: false, series: false }
    // values no axis shows as a tick label
    s.xs = [1, null, 3]
    s.values = [517, 733, 951]
    let at = labelX(renderChart(model!, 400, 260))
    expect('733' in at).toBe(false)
    expect(at['517'] < at['951']).toBe(true)
    // a missing Y: the next points keep their own labels
    s.xs = [1, 2, 3]
    s.values = [null, 733, 951]
    at = labelX(renderChart(model!, 400, 260))
    expect(at['733'] < at['951']).toBe(true)
  })

  test('series stop at the plot area when the axis range leaves values outside it', async () => {
    const { model } = await modelOf(await withChart(CLUSTERED))
    const val = model!.axes.find((a) => a.kind === 'val')!
    val.max = 5
    const svg = renderChart(model!, 400, 260)
    expect(svg).toContain('<clipPath id="de-plot">')
    expect(svg).toMatch(/<g clip-path="url\(#de-plot\)">[\s\S]*<\/g>/)
  })
})

describe('xlsx charts: files that ask for more than they hold', () => {
  test('a saved point count is capped before anything is allocated for it', async () => {
    const huge = `<c:numRef><c:f>SalesRange</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="1000000000"/><c:pt idx="0"><c:v>7</c:v></c:pt><c:pt idx="999999999"><c:v>8</c:v></c:pt></c:numCache></c:numRef>`
    const plot = `<c:lineChart><c:grouping val="standard"/>${ser(0, 'B', huge)}<c:axId val="1"/><c:axId val="2"/></c:lineChart>${AXES}`
    const { model } = await modelOf(await withChart(plot))
    const values = model!.groups[0].series[0].values
    expect(values.length).toBeLessThanOrEqual(100_000)
    expect(values[0]).toBe(7)
  })

  test('a scatter series without X values numbers its points', async () => {
    const yOnly = `<c:ser><c:idx val="0"/><c:order val="0"/><c:tx>${strRef('Data!$B$1', ['y'])}</c:tx><c:spPr><a:ln><a:noFill/></a:ln></c:spPr><c:marker><c:symbol val="circle"/></c:marker><c:yVal>${numRef('Data!$B$2:$B$4', [0, 0, 0])}</c:yVal></c:ser>`
    const plot = `<c:scatterChart><c:scatterStyle val="lineMarker"/>${yOnly}<c:axId val="1"/><c:axId val="2"/></c:scatterChart><c:valAx><c:axId val="1"/><c:axPos val="b"/><c:crossAx val="2"/></c:valAx><c:valAx><c:axId val="2"/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>`
    const { model } = await modelOf(await withChart(plot))
    expect(model!.groups[0].series[0].xs).toBeUndefined()
    // three markers, plus the legend's
    expect(renderChart(model!, 400, 260).match(/<circle /g)?.length).toBe(4)
  })

  test('a doughnut of two series draws two rings', async () => {
    const plot = `<c:doughnutChart><c:varyColors val="1"/>${ser(0, 'B', numRef('Data!$B$2:$B$4', [1, 1, 1]))}${ser(1, 'C', numRef('Data!$C$2:$C$4', [1, 1, 1]))}<c:holeSize val="50"/></c:doughnutChart>`
    const { model } = await modelOf(await withChart(plot))
    // three slices a ring
    expect(renderChart(model!, 400, 300).match(/<path d="M[^"]* A[^"]* L[^"]* A[^"]*Z"/g)?.length).toBe(6)
  })

  test('bars on an axis of negative values only start at its top, the bound nearest zero', async () => {
    const { model } = await modelOf(await withChart(CLUSTERED))
    const s = model!.groups[0].series
    s[0].values = [-95, -92, -95]
    s[1].values = [-95, -95, -95]
    const val = model!.axes.find((a) => a.kind === 'val')!
    val.min = -100
    val.max = -90
    const fills = new Set(s.map((x) => x.fill))
    // the bars, inside the plot area (the legend's swatches come before it): per category the first series, then the second
    const plotted = renderChart(model!, 400, 260).split('<g clip-path')[1]
    const heights = [...plotted.matchAll(/<rect [^>]*height="([\d.]+)"[^>]*fill="([^"]+)"/g)].filter((m) => fills.has(m[2])).map((m) => Number(m[1]))
    // -95 reaches further from -90 than -92 does
    expect(heights[0]).toBeGreaterThan(heights[2])
  })
})

describe('xlsx charts: large series, log axes, hidden points, markers', () => {
  const LINE = (extra = '') => `<c:lineChart><c:grouping val="standard"/>${ser(0, 'B', numRef('Data!$B$2:$B$4', [1, 1, 1]))}${extra}<c:axId val="1"/><c:axId val="2"/></c:lineChart>${AXES}`

  test('three series of 50,000 points each are drawn (no spread past the argument limit)', async () => {
    const { model } = await modelOf(await withChart(CLUSTERED))
    const g = model!.groups[0]
    const big = Array.from({ length: 50_000 }, (_, k) => k % 97)
    g.series = [0, 1, 2].map((i) => ({ ...g.series[i % g.series.length], values: big, cats: big.map(String) }))
    expect(() => renderChart(model!, 400, 260)).not.toThrow()
  })

  test('a log X axis places 10 halfway between 1 and 100', async () => {
    const xy = `<c:ser><c:idx val="0"/><c:order val="0"/><c:tx>${strRef('Data!$B$1', ['y'])}</c:tx><c:marker><c:symbol val="circle"/></c:marker><c:xVal>${numRef('Data!$B$2:$B$4', [0, 0, 0])}</c:xVal><c:yVal>${numRef('Data!$C$2:$C$4', [0, 0, 0])}</c:yVal></c:ser>`
    const plot = `<c:scatterChart><c:scatterStyle val="marker"/>${xy}<c:axId val="1"/><c:axId val="2"/></c:scatterChart><c:valAx><c:axId val="1"/><c:scaling><c:logBase val="10"/></c:scaling><c:axPos val="b"/><c:crossAx val="2"/></c:valAx><c:valAx><c:axId val="2"/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>`
    const { model } = await modelOf(await withChart(plot))
    model!.groups[0].series[0].xs = [1, 10, 100]
    model!.legend = undefined
    const xs = [...renderChart(model!, 400, 260).matchAll(/<circle cx="([\d.]+)"/g)].map((m) => Number(m[1]))
    expect(xs.length).toBe(3)
    expect(Math.abs(xs[1] - (xs[0] + xs[2]) / 2)).toBeLessThan(1)
  })

  test('a log value axis takes its bounds from the positive values; zero is left out', async () => {
    const { model } = await modelOf(await withChart(LINE()))
    model!.groups[0].series[0].values = [0, 10, 100]
    model!.axes.find((a) => a.kind === 'val')!.logBase = 10
    const svg = renderChart(model!, 400, 260)
    expect(svg).not.toMatch(/E-1[0-2]</)
    expect(svg).toContain('>10<')
  })

  test('a point\'s own color follows it when a hidden row before it is left out', async () => {
    const pt = `<c:dPt><c:idx val="2"/><c:spPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></c:spPr></c:dPt>`
    const plot = `<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>${ser(0, 'B', numRef('Data!$B$2:$B$4', [1, 1, 1]), pt)}<c:axId val="1"/><c:axId val="2"/></c:barChart>${AXES}`
    // row 3 holds Q2, the second point: Q3 (point 2 in the file) becomes the second point shown
    const { model } = await modelOf(await withChart(plot, { hideRow: 3 }))
    const s = model!.groups[0].series[0]
    expect(s.values).toEqual([10, 15])
    expect(s.points[1]?.color?.toLowerCase()).toBe('#ff0000')
    expect(s.points[2]).toBeUndefined()
  })

  test('a line chart that turns its markers off has none on series without their own', async () => {
    const { model } = await modelOf(await withChart(LINE('<c:marker val="0"/>')))
    expect(model!.groups[0].series[0].marker).toBeNull()
    const { model: on } = await modelOf(await withChart(LINE()))
    expect(on!.groups[0].series[0].marker).not.toBeNull()
  })
})

describe('xlsx charts: axis formats, labels at the edge, empty cells, date axes, axis crossings', () => {
  const LINE = (extra = '') => `<c:lineChart><c:grouping val="standard"/>${ser(0, 'B', numRef('Data!$B$2:$B$4', [1, 1, 1]), extra)}<c:axId val="1"/><c:axId val="2"/></c:lineChart>${AXES}`
  /** the points of the first line path drawn in the plot */
  const linePoints = (svg: string) => {
    const d = /<path d="(M[^"]*)"[^>]*stroke=/.exec(svg.split('<g clip-path')[1])?.[1] ?? ''
    return [...d.matchAll(/[ML]([\d.]+) ([\d.]+)/g)].map((m) => ({ x: Number(m[1]), y: Number(m[2]) }))
  }

  test('a scatter X axis linked to its source is written in the X values\' format, not the Y values\'', async () => {
    const xy = `<c:ser><c:idx val="0"/><c:order val="0"/><c:tx>${strRef('Data!$B$1', ['y'])}</c:tx><c:xVal>${numRef('Data!$B$2:$B$4', [0, 0, 0])}</c:xVal><c:yVal>${numRef('Data!$C$2:$C$4', [0, 0, 0], '0%')}</c:yVal></c:ser>`
    const plot = `<c:scatterChart><c:scatterStyle val="marker"/>${xy}<c:axId val="1"/><c:axId val="2"/></c:scatterChart><c:valAx><c:axId val="1"/><c:axPos val="b"/><c:crossAx val="2"/></c:valAx><c:valAx><c:axId val="2"/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>`
    const { model } = await modelOf(await withChart(plot))
    const s = model!.groups[0].series[0]
    s.xs = [1, 2, 3]
    s.values = [0.1, 0.2, 0.3]
    s.format = '0%'
    const svg = renderChart(model!, 400, 260)
    expect(svg).toContain('>20%<')
    expect(svg).not.toMatch(/>[123]00%</)
  })

  test('a data label beside a point on the plot\'s edge is drawn outside the plot\'s clipping', async () => {
    const { model } = await modelOf(await withChart(LINE('<c:dLbls><c:showVal val="1"/></c:dLbls>')))
    model!.axes.find((a) => a.kind === 'val')!.max = 20
    const svg = renderChart(model!, 400, 260)
    const clip = svg.indexOf('<g clip-path')
    const closed = svg.indexOf('</g>', clip)
    // Q2 is 20, the axis maximum: its label sits above the plot
    expect(svg.lastIndexOf('>20<')).toBeGreaterThan(closed)
  })

  test('empty cells in a line: a gap, zero or spanned, as the chart says', async () => {
    const { model } = await modelOf(await withChart(LINE()))
    model!.groups[0].series[0].values = [10, null, 15]
    model!.blanks = 'gap'
    expect(linePoints(renderChart(model!, 400, 260)).length).toBe(0)
    model!.blanks = 'span'
    expect(linePoints(renderChart(model!, 400, 260)).length).toBe(2)
    model!.blanks = 'zero'
    expect(linePoints(renderChart(model!, 400, 260)).length).toBe(3)
  })

  test('a date axis places its points by date: a gap of eight days is eight times one of a day', async () => {
    const { model } = await modelOf(await withChart(LINE()))
    model!.axes.find((a) => a.kind === 'cat')!.kind = 'date'
    model!.groups[0].series[0].catValues = [46000, 46001, 46009]
    const [a, b, c] = linePoints(renderChart(model!, 400, 260))
    expect(Math.abs((c.x - b.x) / (b.x - a.x) - 8)).toBeLessThan(0.05)
  })

  test('columns grow from where the category axis crosses: at the top when it crosses at the maximum', async () => {
    const { model } = await modelOf(await withChart(CLUSTERED))
    const fills = new Set(model!.groups[0].series.map((s) => s.fill))
    const tops = (svg: string) => [...svg.split('<g clip-path')[1].matchAll(/<rect [^>]*y="([\d.]+)"[^>]*fill="([^"]+)"/g)].filter((m) => fills.has(m[2])).map((m) => Number(m[1]))
    const atZero = tops(renderChart(model!, 400, 260))
    model!.axes.find((a) => a.kind === 'cat')!.crosses = 'max'
    const atMax = tops(renderChart(model!, 400, 260))
    // from the top every column starts at the same place, the axis maximum
    expect(new Set(atMax).size).toBe(1)
    expect(atMax[0]).toBeLessThan(Math.min(...atZero))
  })

  test('a range given in full keeps its empty cells at the end as points', async () => {
    // B2:B6 reaches two rows past the filled ones
    const plot = `<c:lineChart><c:grouping val="standard"/>${ser(0, 'B', numRef('Data!$B$2:$B$6', [1, 1, 1, 1, 1]))}<c:axId val="1"/><c:axId val="2"/></c:lineChart>${AXES}`
    const { model } = await modelOf(await withChart(plot))
    expect(model!.groups[0].series[0].values).toEqual([10, 20, 15, null, null])
  })

  test('empty cells shown as zero are inside the axis range', async () => {
    const { model } = await modelOf(await withChart(LINE()))
    model!.groups[0].series[0].values = [90, null, 100]
    model!.blanks = 'zero'
    const svg = renderChart(model!, 400, 260)
    // the axis reaches down to the zero, and the line's three points are all inside the plot
    expect(svg).toContain('>0<')
    const plot = /<clipPath id="de-plot"><rect x="[\d.]+" y="([\d.]+)" width="[\d.]+" height="([\d.]+)"/.exec(svg)!
    const bottom = Number(plot[1]) + Number(plot[2])
    expect(linePoints(svg).every((p) => p.y <= bottom)).toBe(true)
  })

  test('values all alike keep zero in view; a bound fixed past the values still gives the axis a span', () => {
    expect(niceScale(10, 10)).toMatchObject({ min: 0 })
    const above = niceScale(4, 20, { min: 30 })
    expect(above.min).toBe(30)
    expect(above.max).toBeGreaterThan(30)
    expect(above.ticks.length).toBeGreaterThan(1)
    const below = niceScale(4, 20, { max: 2 })
    expect(below.max).toBe(2)
    expect(below.min).toBeLessThan(2)
  })

  test('columns of equal values have a height', async () => {
    const { model } = await modelOf(await withChart(CLUSTERED))
    for (const s of model!.groups[0].series) s.values = [10, 10, 10]
    const fills = new Set(model!.groups[0].series.map((s) => s.fill))
    const heights = [...renderChart(model!, 400, 260).split('<g clip-path')[1].matchAll(/<rect [^>]*height="([\d.]+)"[^>]*fill="([^"]+)"/g)].filter((m) => fills.has(m[2])).map((m) => Number(m[1]))
    expect(heights.length).toBe(6)
    expect(heights.every((h) => h > 10)).toBe(true)
  })

  test('an Excel 2010 style number (142) the XML reader already took out of its AlternateContent is style 42', async () => {
    const { model } = await modelOf(await withChart(CLUSTERED, { style: 142 }))
    expect(model!.textColor.toLowerCase()).toBe('#ffffff')
  })
})

describe('xlsx charts: hostile input', () => {
  const cached = (n: number, ptCount: number | string = n) =>
    `<c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="${ptCount}"/>${Array.from({ length: n }, (_, i) => `<c:pt idx="${i}"><c:v>${i}</c:v></c:pt>`).join('')}</c:numLit>`

  test('a dash named after an inherited key ("constructor") is drawn solid', async () => {
    const dashed = `<c:spPr><a:ln w="19050"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:prstDash val="constructor"/></a:ln></c:spPr>`
    const plot = `<c:lineChart><c:grouping val="standard"/>${ser(0, 'B', numRef('Data!$B$2:$B$4', [1, 2, 3]), dashed)}<c:axId val="1"/><c:axId val="2"/></c:lineChart>${AXES}`
    const { model } = await modelOf(await withChart(plot))
    const svg = renderChart(model!, 400, 260)
    expect(svg).toContain('<path')
    expect(svg).not.toContain('stroke-dasharray')
  })

  test('a group named after an inherited key (<c:constructor>) is skipped, and the chart beside it kept', async () => {
    const { model } = await modelOf(await withChart(`<c:constructor>${ser(0, 'B', numRef('Data!$B$2:$B$4', [1, 1, 1]))}</c:constructor>${CLUSTERED}`))
    expect(model!.groups.map((g) => g.type)).toEqual(['bar'])
    // a model of plain data: what the worker posts
    expect(() => structuredClone(model)).not.toThrow()
  })

  test('an empty pie group beside a bar group draws the bars', async () => {
    const { model } = await modelOf(await withChart(`<c:pieChart><c:varyColors val="1"/></c:pieChart>${CLUSTERED}`))
    expect(model!.groups.map((g) => g.type)).toEqual(['bar'])
    expect(renderChart(model!, 400, 260)).toContain('<rect')
    // a model made elsewhere with the empty group still draws
    model!.groups.unshift({ ...model!.groups[0], type: 'pie', series: [] })
    expect(() => renderChart(model!, 400, 260)).not.toThrow()
  })

  test('a point count that is not a whole number is read as the whole part', async () => {
    const plot = `<c:lineChart><c:grouping val="standard"/>${ser(0, 'B', `<c:numRef><c:f>SalesRange</c:f>${cached(3, '3.5').replace(/c:numLit/g, 'c:numCache')}</c:numRef>`)}<c:axId val="1"/><c:axId val="2"/></c:lineChart>${AXES}`
    const { model } = await modelOf(await withChart(plot))
    expect(model!.groups[0].series[0].values).toEqual([0, 1, 2])
  })

  test('more series than Excel allows: the placeholder', async () => {
    const many = Array.from({ length: 256 }, (_, i) => `<c:ser><c:idx val="${i}"/><c:order val="${i}"/><c:val>${cached(1)}</c:val></c:ser>`).join('')
    const { model, drawing } = await modelOf(await withChart(`<c:lineChart><c:grouping val="standard"/>${many}<c:axId val="1"/><c:axId val="2"/></c:lineChart>${AXES}`))
    expect(model).toBeUndefined()
    expect(drawing.chart).toBe('line')
  })

  test('saved values past the chart-wide budget: the placeholder, not a chart of millions of points', async () => {
    // 30 series of a saved count of 100 000 each: within the cap of one cache, past the chart's
    const many = Array.from({ length: 30 }, (_, i) => `<c:ser><c:idx val="${i}"/><c:order val="${i}"/><c:val>${cached(1, 100_000)}</c:val></c:ser>`).join('')
    const { model } = await modelOf(await withChart(`<c:lineChart><c:grouping val="standard"/>${many}<c:axId val="1"/><c:axId val="2"/></c:lineChart>${AXES}`))
    expect(model).toBeUndefined()
  })

  test('a reference of many areas counts its cells together: its saved values stand in', async () => {
    // 200 areas of 300 cells: each is small, together they are past the cap of one reference
    const areas = Array.from({ length: 200 }, () => 'Data!$B$1:$B$300').join(',')
    const plot = `<c:lineChart><c:grouping val="standard"/>${ser(0, 'B', numRef(`(${areas})`, [5, 6, 7]))}<c:axId val="1"/><c:axId val="2"/></c:lineChart>${AXES}`
    const { model } = await modelOf(await withChart(plot))
    expect(model!.groups[0].series[0].values).toEqual([5, 6, 7])
  })

  test('a chart shown by several anchors is read once, each anchor with a model of its own', async () => {
    const engine = new Engine(() => {})
    await engine.open(await withChart(CLUSTERED, { anchors: 3 }), {})
    const pkg = (engine as unknown as { xlsx: { pkg: { xml: (part: string) => Promise<unknown> } } }).xlsx.pkg
    const xml = pkg.xml.bind(pkg)
    let reads = 0
    pkg.xml = (part) => ((reads += part.endsWith('chart1.xml') ? 1 : 0), xml(part))
    const drawings = (await engine.call('xlsx.drawings', [0])) as SheetDrawing[]
    expect(drawings.length).toBe(3)
    expect(reads).toBe(1)
    expect(drawings.every((d) => d.model?.groups[0].series.length === 2)).toBe(true)
    expect(new Set(drawings.map((d) => d.model)).size).toBe(3)
  })

  test('a gap width, overlap and hole size outside what Excel takes are kept to it', async () => {
    const plot = CLUSTERED.replace('<c:gapWidth val="150"/>', '<c:gapWidth val="-100"/><c:overlap val="-400"/>')
    const { model } = await modelOf(await withChart(plot))
    expect(model!.groups[0].gapWidth).toBe(0)
    expect(model!.groups[0].overlap).toBe(-100)
    expect(renderChart(model!, 400, 260)).not.toContain('NaN')
    const ring = `<c:doughnutChart><c:varyColors val="1"/>${ser(0, 'B', numRef('Data!$B$2:$B$4', [1, 1, 1]))}<c:holeSize val="5000"/></c:doughnutChart>`
    expect((await modelOf(await withChart(ring))).model!.groups[0].holeSize).toBe(90)
  })

  test('a control character in a title is left out, so the picture stays well-formed XML', async () => {
    const { model } = await modelOf(await withChart(CLUSTERED, { title: 'Sales&#1;\u{1F4C8} 2026' }))
    const svg = renderChart(model!, 400, 260)
    expect(svg).toContain('Sales\u{1F4C8} 2026')
    expect(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/.test(svg)).toBe(false)
    expect(renderChart({ ...model!, title: { ...model!.title!, text: 'a\uD800b' } }, 400, 260)).toContain('>ab<')
  })

  test('fixed bounds the wrong way round or a step of zero still give a scale', () => {
    for (const opts of [{ min: 10, max: 2 }, { major: 0 }, { min: 0, max: 0 }, { major: Infinity }, { min: NaN, max: Infinity }]) {
      const s = niceScale(1, 9, opts)
      expect(Number.isFinite(s.min) && Number.isFinite(s.max) && s.max > s.min).toBe(true)
      expect(s.ticks.length).toBeGreaterThan(1)
      expect(s.ticks.every(Number.isFinite)).toBe(true)
    }
  })
})
