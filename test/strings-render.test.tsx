import { describe, expect, test } from 'bun:test'
import { renderToStaticMarkup } from 'react-dom/server'
import { DocViewer } from '../src/react/DocViewer'
import { PromptDialog } from '../src/react/common/Dialog'
import { FindBar } from '../src/react/docx/FindBar'
import { resolveStrings, StringsProvider, useStrings } from '../src/react/strings'
import { de } from '../src/locales/de'
import { drawingEl } from '../src/react/xlsx/drawings'
import type { ChartModel } from '../src/core/xlsx/chart'

const noop = () => {}
const find = <FindBar query="Office" index={1} count={3} editable replace onQuery={noop} onNav={noop} onClose={noop} onReplace={noop} />

describe('rendered UI strings', () => {
  test('German find/replace labels and a partial override reach the rendered controls', () => {
    const html = renderToStaticMarkup(
      <StringsProvider strings={{ ...de, find: { ...de.find, replaceAll: 'Alle Treffer ersetzen' } }}>
        {find}
        <PromptDialog spec={{ title: 'Test', onOk: noop }} onClose={noop} />
      </StringsProvider>,
    )
    expect(html).toContain('placeholder="Im Dokument suchen"')
    expect(html).toContain('aria-label="Suchen"')
    expect(html).toContain('aria-label="Vorheriges (Umschalt+Eingabe)"')
    expect(html).toContain('aria-label="Ersetzen durch"')
    expect(html).toContain('>Ersetzen</button>')
    expect(html).toContain('>Alle Treffer ersetzen</button>')
    expect(html).toContain('>Abbrechen</button>')
    expect(html).not.toContain('Find in document')
    expect(html).not.toContain('>Cancel</button>')
  })

  test('standalone controls default to English and the German render leaves that default intact', () => {
    const html = renderToStaticMarkup(find)
    expect(html).toContain('placeholder="Find in document"')
    expect(html).toContain('aria-label="Replace with"')
    expect(html).toContain('>All</button>')
  })

  test('drawing placeholders translate labels while keeping the document title', () => {
    const chart = { kind: 'chart' as const, chart: 'column', title: 'Sales 2026' }
    expect(renderToStaticMarkup(drawingEl('c', chart, { width: 200, height: 100 }, resolveStrings('en')))).toContain('<strong>Column chart</strong>')
    const german = resolveStrings(undefined, de)
    const html = renderToStaticMarkup(drawingEl('c', chart, { width: 200, height: 100 }, german))
    expect(html).toContain('title="Säulendiagramm: Sales 2026"')
    expect(html).toContain('<strong>Säulendiagramm</strong>')
    expect(renderToStaticMarkup(drawingEl('p', { kind: 'picture', title: '' }, { width: 200, height: 100 }, german))).toContain('<strong>Bild</strong>')
    for (const kind of ['', 'unknown', 'constructor', '__proto__']) {
      expect(renderToStaticMarkup(drawingEl('c', { ...chart, chart: kind }, { width: 200, height: 100 }, german))).toContain('<strong>Diagramm</strong>')
    }
  })

  test('a chart that cannot be drawn shows its placeholder box instead of taking the view down', () => {
    const broken = { groups: null } as unknown as ChartModel
    const warn = console.warn
    console.warn = () => {}
    try {
      const html = renderToStaticMarkup(drawingEl('c', { kind: 'chart', chart: 'column', title: 'Sales', model: broken }, { width: 200, height: 100 }, resolveStrings('en')))
      expect(html).toContain('de-xl-drawing')
      expect(html).toContain('<strong>Column chart</strong>')
      // asked again (the next render), it is not drawn again
      expect(renderToStaticMarkup(drawingEl('c', { kind: 'chart', chart: 'column', title: 'Sales', model: broken }, { width: 200, height: 100 }, resolveStrings('en')))).toContain('de-xl-drawing')
    } finally {
      console.warn = warn
    }
  })

  test('drawing labels accept partial overrides and pictures retain their original alt text', () => {
    const strings = resolveStrings(undefined, { ...de, xlsx: { ...de.xlsx, drawings: { ...de.xlsx.drawings, picture: 'Grafik', charts: { ...de.xlsx.drawings.charts, column: 'Umsatzdiagramm' } } } })
    expect(renderToStaticMarkup(drawingEl('c', { kind: 'chart', chart: 'column', title: '' }, { width: 200, height: 100 }, strings))).toContain('<strong>Umsatzdiagramm</strong>')
    expect(renderToStaticMarkup(drawingEl('p', { kind: 'picture', title: '' }, { width: 200, height: 100 }, strings))).toContain('<strong>Grafik</strong>')
    expect(renderToStaticMarkup(drawingEl('c', { kind: 'chart', chart: 'line', title: '' }, { width: 200, height: 100 }, strings))).toContain('<strong>Liniendiagramm</strong>')
    const picture = renderToStaticMarkup(drawingEl('p', { kind: 'picture', title: 'Company logo', url: 'blob:picture' }, { width: 200, height: 100 }, strings))
    expect(picture).toContain('alt="Company logo"')
    expect(picture).not.toContain('de-xl-drawing')
  })
})

function TextProbe() {
  const s = useStrings()
  return <span>{s.file.menu}|{s.common.cancel}|{s.locale ?? 'browser'}</span>
}

test('nested providers inherit texts and formatting; explicit props override locally', () => {
  const html = renderToStaticMarkup(<StringsProvider strings={de}>
    <StringsProvider><TextProbe /></StringsProvider>
    <StringsProvider locale="en-GB" strings={{ file: { menu: 'Ablage' } }}><TextProbe /></StringsProvider>
  </StringsProvider>)
  expect(html).toContain('Datei|Abbrechen|de-DE')
  expect(html).toContain('Ablage|Abbrechen|en-GB')
})

test('DocViewer inherits an enclosing German provider and applies partial overrides', () => {
  const html = renderToStaticMarkup(<StringsProvider strings={de}><DocViewer src={null} strings={{ bar: { noDocument: 'Datei wählen' } }} /></StringsProvider>)
  expect(html).toContain('Datei wählen')
  expect(html).toContain('Datei')
  expect(html).not.toContain('Open a file')
})
