/**
 * What a crafted file (or a collaboration peer) must not be able to do: have the browser request an address,
 * name an inherited key where a known one is expected, announce sizes the reader then allocates, nest XML past the
 * stack, or have other clients run any worker method.
 */
import { afterEach, describe, expect, test } from 'bun:test'
import { colorOf, cssColor, cssPaint, customIndexed, drawingColor, fileHex, officeHex, sheetColor } from '../src/core/color'
import { cssFont } from '../src/core/fonts'
import { parseRaw } from '../src/core/raw'
import { BlobSource, BufferSource } from '../src/core/source'
import { parseTheme } from '../src/core/theme'
import { Engine } from '../src/core/worker/engine'
import { MAX_XML_DEPTH, parseXml, type XEl } from '../src/core/xml'
import { ZIP_LIMITS, ZipArchive } from '../src/core/zip'
import { rebuildZip } from '../src/core/zipwrite'
import { CollabHub } from '../src/collab/hub'
import { isSharedCall } from '../src/collab/docx'
import { CollabSession, type Adapter } from '../src/collab/session'
import { MemoryTransport } from '../src/collab/transport'
import { PEER_COLORS } from '../src/collab/protocol'
import { INHERITED, inheritedDocx, inheritedPptx, inheritedXlsx, PROBE_HOST, probeDocx, probePptx, probeXlsx } from './probes'

const S = (n: string) => `${import.meta.dir}/../playground/public/samples/${n}`
const URL_COLOR = 'fff url(https://probe.invalid/x.png)'
const el = (n: string, a: Record<string, string> = {}, c: XEl[] = []): XEl => ({ n, a, c, t: '' })

async function open(file: Blob) {
  const events: { event: string; data: any }[] = []
  const engine = new Engine((event, data) => events.push({ event, data }))
  const info = await engine.open(file, {})
  return { engine, info, events }
}

/** everything the views are given for a document: what the worker would send to the page */
async function models(file: Blob): Promise<unknown[]> {
  const { engine, info, events } = await open(file)
  const out: unknown[] = [info]
  if (info.kind === 'docx') {
    await engine.call('docx.start', [])
    expect(events.find((e) => e.event === 'docx:done')?.data).toEqual({})
    out.push(...events.filter((e) => e.event === 'docx:blocks' || e.event === 'docx:meta').map((e) => e.data))
  } else if (info.kind === 'xlsx') {
    for (let s = 0; s < info.xlsx!.sheets.length; s++) {
      await engine.call('xlsx.load', [s])
      for (let i = 0; i < 300 && !events.some((e) => e.event === 'xlsx:progress' && e.data.sheet === s && e.data.done); i++) await new Promise((r) => setTimeout(r, 10))
      out.push(await engine.call('xlsx.block', [s, 0, 200]))
    }
    out.push(...events.filter((e) => e.event === 'xlsx:progress').map((e) => e.data))
  } else {
    for (let i = 0; i < info.pptx!.count; i++) out.push(await engine.call('pptx.slide', [i]))
  }
  return out
}

describe('colors given to CSS', () => {
  test('cssColor: a color and nothing more', () => {
    for (const ok of ['#fff', '#FFFF', '#a1b2c3', '#a1b2c3d4', 'rgb(1, 2, 3)', 'rgba(1,2,3,0.5)', 'hsl(120 50% 50% / 0.5)', 'red', 'transparent', 'currentColor']) expect(cssColor(ok)).toBe(ok)
    expect(cssColor('  #fff ')).toBe('#fff')
    for (const bad of [
      '#' + URL_COLOR,
      'url(https://probe.invalid/x.png)',
      '#fff url(x)',
      'red url(x)',
      'u\\72l(x)',
      'image-set("x" 1x)',
      'rgb(1,2,3) url(x)',
      'rgb(url(x))',
      'var(--x)',
      'linear-gradient(#fff, #000)',
      '#fff;background:url(x)',
      '#ff',
      '#fffff',
      '',
      'a',
      'x'.repeat(200),
      undefined,
      null,
      7,
      {},
      () => 'red',
    ])
      expect(cssColor(bad)).toBeUndefined()
  })

  test('cssPaint: a color, or a gradient of colors the way the engine writes it', () => {
    for (const ok of ['#fff', 'linear-gradient(180deg, #ff0000 0%, rgba(1,2,3,0.5) 100%)', 'radial-gradient(circle, #ff0000 0%, #00ff00 50%, #0000ff 100%)', 'linear-gradient(135.5deg, #fff, #000)'])
      expect(cssPaint(ok)).toBe(ok)
    for (const bad of [
      'linear-gradient(180deg, #fff 0%, url(https://probe.invalid/x.png) 100%)',
      'linear-gradient(180deg, #fff 0%, #000 100%), url(https://probe.invalid/x.png)',
      'linear-gradient(180deg, #fff url(x) 0%)',
      'radial-gradient(circle, #fff) url(x)',
      'conic-gradient(#fff, #000)',
      'image-set(url(x) 1x)',
      'element(#x)',
      '#fff url(x)',
      undefined,
      {},
    ])
      expect(cssPaint(bad)).toBeUndefined()
  })

  test('fileHex / officeHex: hex digits only', () => {
    expect(fileHex('FF0000')).toBe('#ff0000')
    expect(fileHex('80FF0000')).toBe('#ff0000')
    expect(fileHex('F00')).toBe('#ff0000')
    for (const bad of [URL_COLOR, 'FF0000 ', 'GG0000', 'FF00', '', undefined, '#FF0000']) expect(fileHex(bad)).toBeUndefined()
    expect(officeHex('#ff0000')).toBe('FF0000')
    expect(officeHex('ff0000')).toBe('FF0000')
    expect(officeHex('#f00')).toBe('FF0000')
    for (const bad of ['#' + URL_COLOR, '"/><x a="', '#80ff0000', 'red', 'auto', '', null, undefined]) expect(officeHex(bad)).toBeUndefined()
  })

  test('DrawingML colors: srgbClr, sysClr, theme and placeholder colors', () => {
    expect(drawingColor(el('srgbClr', { val: 'FF0000' }), {})).toBe('#ff0000')
    expect(drawingColor(el('srgbClr', { val: URL_COLOR }), {})).toBeUndefined()
    expect(drawingColor(el('srgbClr', { val: URL_COLOR }, [el('lumMod', { val: '50000' })]), {})).toBeUndefined()
    expect(drawingColor(el('sysClr', { val: 'window', lastClr: URL_COLOR }), {})).toBe('#ffffff')
    expect(drawingColor(el('sysClr', { val: 'windowText', lastClr: '112233' }), {})).toBe('#112233')
    // a color that reaches the resolver from elsewhere (the theme, a style's placeholder color) is checked too
    const theme = { ...parseTheme(undefined), colors: { accent1: '#' + URL_COLOR } }
    expect(drawingColor(el('schemeClr', { val: 'accent1' }), { theme })).toBeUndefined()
    expect(drawingColor(el('schemeClr', { val: 'phClr' }), { phClr: '#' + URL_COLOR })).toBeUndefined()
    expect(colorOf(el('solidFill', {}, [el('srgbClr', { val: URL_COLOR })]), {})).toBeUndefined()
  })

  test('theme colors: one that is not hex digits keeps the default', () => {
    const scheme = el('clrScheme', {}, [el('accent1', {}, [el('srgbClr', { val: URL_COLOR })]), el('accent2', {}, [el('srgbClr', { val: '00FF00' })]), el('dk1', {}, [el('sysClr', { val: 'windowText', lastClr: URL_COLOR })])])
    const theme = parseTheme(el('theme', {}, [el('themeElements', {}, [scheme])]))
    expect(theme.colors.accent1).toBe('#4472c4')
    expect(theme.colors.accent2).toBe('#00ff00')
    expect(theme.colors.dk1).toBe('#000000')
  })

  test('SpreadsheetML colors: rgb, the workbook palette', () => {
    expect(sheetColor(el('color', { rgb: 'FF4472C4' }), undefined)).toBe('#4472c4')
    expect(sheetColor(el('color', { rgb: URL_COLOR }), undefined)).toBeUndefined()
    const styles = el('styleSheet', {}, [el('colors', {}, [el('indexedColors', {}, [el('rgbColor', { rgb: URL_COLOR }), el('rgbColor', { rgb: '00112233' })])])])
    const palette = customIndexed(styles)
    expect(palette).toEqual(['#000000', '#112233'])
    expect(sheetColor(el('color', { indexed: '0' }), undefined, palette)).toBe('#000000')
  })

  test('a font name is quoted for CSS without what would end the quote', () => {
    expect(cssFont('Evil"\\ \n; background: url(x)')).toBe('"Evil ; background: url(x)", sans-serif')
    expect(cssFont('constructor')).toBe('"constructor", sans-serif')
  })

  for (const [kind, build] of [['pptx', probePptx], ['xlsx', probeXlsx], ['docx', probeDocx]] as const) {
    test(`${kind}: no color of a crafted file reaches the views with a url() in it`, async () => {
      const json = JSON.stringify(await models(await build()))
      expect(json.length).toBeGreaterThan(1000)
      expect(json).not.toContain('url(')
      expect(json).not.toContain(PROBE_HOST)
    })
  }

  test('an edit writes a color into the file only when it is one', async () => {
    const bad = '#fff"/><a:evil xmlns:a="x'
    const pptx = await open(Bun.file(S('sample.pptx')) as unknown as Blob)
    await pptx.engine.call('pptx.shapes', [0, [{ op: 'bg', color: bad }]])
    const slide = (await pptx.engine.call('pptx.slide', [0])) as { bg?: { t: string; c?: string } }
    expect(JSON.stringify(slide)).not.toContain('evil')
    await pptx.engine.call('pptx.shapes', [0, [{ op: 'bg', color: '#336699' }]])
    expect(((await pptx.engine.call('pptx.slide', [0])) as { bg?: unknown }).bg).toEqual({ t: 'solid', c: '#336699' })

    const xlsx = await open(Bun.file(S('sample.xlsx')) as unknown as Blob)
    await xlsx.engine.call('xlsx.load', [0])
    for (let i = 0; i < 300 && !xlsx.events.some((e) => e.event === 'xlsx:progress' && e.data.done); i++) await new Promise((r) => setTimeout(r, 10))
    const range = { r0: 0, c0: 0, r1: 0, c1: 0 }
    const res = (await xlsx.engine.call('xlsx.format', [0, range, { fill: '#' + URL_COLOR, color: bad, border: 'all', borderColor: bad }])) as { styles?: { css: Record<string, unknown> }[] }
    const css = JSON.stringify(res.styles ?? [])
    expect(css).not.toContain('url(')
    expect(css).not.toContain('evil')
    const saved = await ZipArchive.open(new BlobSource((await xlsx.engine.call('save', [])) as Blob))
    expect(await saved.text('xl/styles.xml')).not.toContain('probe.invalid')
  })
})

describe('keys every object inherits', () => {
  test('lookups by a name the file gives', () => {
    for (const name of INHERITED) {
      expect(drawingColor(el('prstClr', { val: name }), {})).toBe('#000000')
      expect(drawingColor(el('schemeClr', { val: name }), { theme: parseTheme(undefined), clrMap: { bg1: 'lt1' } })).toBeUndefined()
      expect(sheetColor(el('color', { theme: name }), parseTheme(undefined))).toBeUndefined()
      expect(typeof cssFont(name)).toBe('string')
      expect(cssFont(name)).not.toContain('function')
    }
  })

  for (const [kind, build] of [['pptx', inheritedPptx], ['xlsx', inheritedXlsx], ['docx', inheritedDocx]] as const) {
    test(`${kind}: a file naming them opens, and what the views get can leave the worker`, async () => {
      for (const name of INHERITED) {
        const all = await models(await build(name))
        // a function (or anything else that can't be cloned) in a model fails the worker's answer: a blank document
        expect(() => structuredClone(all)).not.toThrow()
        expect(JSON.stringify(all)).not.toContain('[native code]')
      }
    })
  }
})

describe('ZIP reader limits', () => {
  const limits = { ...ZIP_LIMITS }
  afterEach(() => void Object.assign(ZIP_LIMITS, limits))

  const u16 = (n: number) => [n & 255, (n >> 8) & 255]
  const u32 = (n: number) => [n & 255, (n >>> 8) & 255, (n >>> 16) & 255, (n >>> 24) & 255]

  /** a ZIP of stored (or as-given) entries; `size`: the unpacked size the directory announces */
  function zip(files: { name: string; data?: Uint8Array; method?: number; size?: number }[], patch?: (eocd: number[]) => void): Uint8Array {
    const enc = new TextEncoder()
    const local: number[] = []
    const central: number[] = []
    for (const f of files) {
      const name = [...enc.encode(f.name)]
      const data = [...(f.data ?? new Uint8Array())]
      const size = f.size ?? data.length
      const head = [...u16(20), ...u16(0), ...u16(f.method ?? 0), ...u16(0), ...u16(0), ...u32(0), ...u32(data.length), ...u32(size), ...u16(name.length), ...u16(0)]
      central.push(...u32(0x02014b50), ...u16(20), ...head, ...u16(0), ...u16(0), ...u16(0), ...u32(0), ...u32(local.length), ...name)
      local.push(...u32(0x04034b50), ...head, ...name, ...data)
    }
    const eocd = [...u32(0x06054b50), ...u16(0), ...u16(0), ...u16(files.length), ...u16(files.length), ...u32(central.length), ...u32(local.length), ...u16(0)]
    patch?.(eocd)
    return new Uint8Array([...local, ...central, ...eocd])
  }
  const archive = (bytes: Uint8Array) => ZipArchive.open(new BufferSource(bytes))
  const deflated = async (bytes: Uint8Array) =>
    new Uint8Array(await new Response(new Blob([bytes as BlobPart]).stream().pipeThrough(new CompressionStream('deflate-raw') as unknown as TransformStream<Uint8Array, Uint8Array>)).arrayBuffer())

  test('a well-formed archive reads as before', async () => {
    const z = await archive(zip([{ name: 'a.txt', data: new TextEncoder().encode('hello') }, { name: 'dir/b.txt', data: new TextEncoder().encode('world') }]))
    expect(await z.text('a.txt')).toBe('hello')
    expect(await z.text('DIR/B.TXT')).toBe('world')
    const packed = await archive(zip([{ name: 'a.txt', data: await deflated(new TextEncoder().encode('x'.repeat(5000))), method: 8, size: 5000 }]))
    expect((await packed.text('a.txt')).length).toBe(5000)
  })

  test('two entries of one name are refused, by case too', async () => {
    const one = new TextEncoder().encode('1')
    await expect(archive(zip([{ name: 'word/document.xml', data: one }, { name: 'word/document.xml', data: one }]))).rejects.toThrow(/two files named/)
    await expect(archive(zip([{ name: 'word/document.xml', data: one }, { name: 'Word/Document.xml', data: one }]))).rejects.toThrow(/two files named/)
    await expect(archive(zip([{ name: 'a.xml', data: one }, { name: '/a.xml', data: one }]))).rejects.toThrow(/two files named/)
  })

  test('an announced size past the limit is refused before anything is allocated', async () => {
    const z = await archive(zip([{ name: 'big.xml', data: new Uint8Array(4), method: 8, size: 0xfffffff0 }]))
    await expect(z.bytes('big.xml')).rejects.toThrow(/too large/)
    await expect(z.text('big.xml')).rejects.toThrow(/too large/)
    ZIP_LIMITS.stream = 1000
    const s = await archive(zip([{ name: 'big.xml', data: new Uint8Array(4), size: 5000 }]))
    await expect(s.stream('big.xml')).rejects.toThrow(/too large/)
  })

  test('an entry that unpacks to more than it announced stops at the limit', async () => {
    const bomb = await deflated(new Uint8Array(4 * 1024 * 1024))
    expect(bomb.length).toBeLessThan(8192)
    const z = await archive(zip([{ name: 'bomb.xml', data: bomb, method: 8, size: 10 }]))
    // within the limit: the bytes are all there, whatever the directory said
    expect((await z.bytes('bomb.xml')).length).toBe(4 * 1024 * 1024)
    ZIP_LIMITS.whole = 1024 * 1024
    await expect(z.bytes('bomb.xml')).rejects.toThrow(/too large/)
    ZIP_LIMITS.stream = 1024 * 1024
    const reader = (await z.stream('bomb.xml')).getReader()
    const drain = async () => {
      for (;;) if ((await reader.read()).done) return
    }
    await expect(drain()).rejects.toThrow(/too large/)
  })

  test('a directory or an entry outside the file, and too many entries, are refused', async () => {
    const one = new TextEncoder().encode('1')
    // central directory size and offset pointing past the end
    await expect(archive(zip([{ name: 'a', data: one }], (e) => e.splice(12, 4, ...u32(0x7fffffff))))).rejects.toThrow(/outside the file/)
    await expect(archive(zip([{ name: 'a', data: one }], (e) => e.splice(16, 4, ...u32(0x7fffffff))))).rejects.toThrow(/outside the file/)
    ZIP_LIMITS.entries = 2
    await expect(archive(zip([{ name: 'a', data: one }, { name: 'b', data: one }, { name: 'c', data: one }]))).rejects.toThrow(/Too many files/)
    Object.assign(ZIP_LIMITS, limits)
    // an entry whose data would run past the end of the file
    const bytes = zip([{ name: 'a', data: one }])
    const dv = new DataView(bytes.buffer)
    const cd = bytes.length - 22 - (46 + 1)
    dv.setUint32(cd + 20, 5000, true)
    const z = await archive(bytes)
    await expect(z.bytes('a')).rejects.toThrow(/past the end/)
  })

  test('a document with a part named twice does not open', async () => {
    const src = await ZipArchive.open(new BlobSource(new Blob([await Bun.file(S('sample.docx')).arrayBuffer()])))
    const doc = await src.bytes('word/document.xml')
    const twice = await archive(zip([{ name: '[Content_Types].xml', data: await src.bytes('[Content_Types].xml') }, { name: 'word/document.xml', data: doc }, { name: 'word/Document.xml', data: doc }]))
      .then(() => 'opened')
      .catch((e: Error) => e.message)
    expect(twice).toMatch(/two files named/)
  })
})

describe('XML nesting', () => {
  const nested = (depth: number) => '<a>'.repeat(depth) + 'x' + '</a>'.repeat(depth)

  test('as deep as the limit parses; deeper is refused by every parser', () => {
    expect(parseXml(nested(MAX_XML_DEPTH))?.n).toBe('a')
    expect(() => parseXml(nested(MAX_XML_DEPTH + 1))).toThrow(/nested deeper/)
    expect(() => parseXml(nested(200_000))).toThrow(/nested deeper/)
    expect(parseRaw(nested(MAX_XML_DEPTH)).length).toBe(1)
    expect(() => parseRaw(nested(MAX_XML_DEPTH + 1))).toThrow(/nested deeper/)
    // elements one after the other are not nesting
    expect(parseXml('<r>' + '<a><b/></a>'.repeat(50_000) + '</r>')?.c.length).toBe(50_000)
  })

  test('a document nested past the limit fails to open with that error, not a stack overflow', async () => {
    const zip = await ZipArchive.open(new BlobSource(new Blob([await Bun.file(S('sample.docx')).arrayBuffer()])))
    const xml = await zip.text('word/document.xml')
    const deep = '<w:sdt><w:sdtContent>'.repeat(60_000) + '<w:p><w:r><w:t>deep</w:t></w:r></w:p>' + '</w:sdtContent></w:sdt>'.repeat(60_000)
    const file = await rebuildZip(zip, new Map([['word/document.xml', xml.replace('<w:body>', '<w:body>' + deep)]]))
    const { engine, events } = await open(file)
    await engine.call('docx.start', [])
    expect(events.find((e) => e.event === 'docx:done')?.data.error).toMatch(/nested deeper/)
  })
})

describe('collaboration: what a peer may ask for', () => {
  test('a `call` op names a shared document edit, with arguments of its shape', () => {
    expect(isSharedCall('docx.style', ['h1'])).toBe(true)
    expect(isSharedCall('docx.styleSet', ['Heading1', { b: true, color: '#ff0000', sz: 14, font: null, line: 1.5, lineRule: 'exact' }])).toBe(true)
    expect(isSharedCall('docx.page', [{ w: 612, h: 792, cols: 2 }])).toBe(true)
    expect(isSharedCall('docx.hf', [{ parts: { header: [{ text: 'Page {PAGE}', align: 'center', b: true, sz: 10 }] }, titlePg: true }])).toBe(true)
    expect(isSharedCall('docx.hf', [{ raw: { footer: '<w:ftr/>' } }])).toBe(true)

    for (const method of ['open', 'save', 'part', 'docx.start', 'docx.edit', 'docx.table', 'docx.image', 'docx.link', 'docx.convert', 'xlsx.import', 'xlsx.set', 'pptx.shapes', 'collab.xlsx.apply', 'comments.set', 'constructor', '__proto__', 'toString', ''])
      expect(isSharedCall(method, ['https://probe.invalid/x'])).toBe(false)
    expect(isSharedCall(undefined, [])).toBe(false)
    expect(isSharedCall({ toString: () => 'docx.page' }, [{}])).toBe(false)
    expect(isSharedCall('docx.style', 'h1')).toBe(false)
    expect(isSharedCall('docx.style', ['constructor'])).toBe(false)
    expect(isSharedCall('docx.style', ['h1', 'h2'])).toBe(false)
    expect(isSharedCall('docx.styleSet', ['Heading1', { b: true, evil: 1 }])).toBe(false)
    expect(isSharedCall('docx.styleSet', ['Heading1', { color: { toString: () => '#fff' } }])).toBe(false)
    expect(isSharedCall('docx.styleSet', [{ id: 'x' }, {}])).toBe(false)
    expect(isSharedCall('docx.page', [{ w: '612' }])).toBe(false)
    expect(isSharedCall('docx.page', [{ w: 612, __proto__: null, src: 'x' }])).toBe(false)
    expect(isSharedCall('docx.page', [[612]])).toBe(false)
    expect(isSharedCall('docx.hf', [{ parts: { body: [] } }])).toBe(false)
    expect(isSharedCall('docx.hf', [{ parts: { header: [{ text: 5 }] } }])).toBe(false)
    expect(isSharedCall('docx.hf', [{ raw: { header: { xml: '<x/>' } } }])).toBe(false)
  })

  test('a `call` from a peer naming another method is not run by the clients that receive it', async () => {
    const { DocumentClient } = await import('../src/core/worker/client')
    const { startCollab } = await import('../src/collab')
    const { registerHost } = await import('../src/collab/hosts')
    const { DocxEditor } = await import('../src/react/docx/DocxEditor')
    const hub = new CollabHub()
    const join = async (name: string) => {
      const client = await DocumentClient.open(Bun.file(S('sample.docx')) as unknown as Blob, { worker: false })
      const st = { blocks: [] as any[], heights: new WeakMap(), dirty: new Set<number>(), frontier: 0 }
      client.on('docx:blocks', (b) => st.blocks.push(...b))
      const done = new Promise<void>((r) => client.on('docx:done', () => r()))
      void client.call('docx.start')
      await done
      registerHost(client, 'docx', { editor: new DocxEditor(client, st as any, () => {}), st: st as any })
      const calls: string[] = []
      const direct = client.direct.bind(client) as (m: string, ...a: unknown[]) => Promise<unknown>
      client.direct = ((m: string, ...a: unknown[]) => (calls.push(m), direct(m, ...a))) as typeof client.direct
      const transport = new MemoryTransport(hub, () => 1)
      const collab = startCollab(client, { transport, room: 'r', user: { name } })
      return { client, collab, transport, calls }
    }
    const a = await join('Ann')
    const b = await join('Bob')
    const settle = async () => {
      for (let i = 0; i < 300; i++) {
        await new Promise((r) => setTimeout(r, 10))
        await Promise.all([a, b].map((c) => c.collab.session.idle()))
        if ([a, b].every((c) => c.collab.pending === 0 && c.collab.status === 'live') && a.collab.session.rev === b.collab.session.rev) return
      }
      throw new Error('did not settle')
    }
    await settle()
    const errors: string[] = []
    b.collab.session.on('error', (e) => errors.push(e.message))
    const quiet = console.error
    console.error = () => {}
    try {
      // what a modified client would send: any worker method, with its arguments
      a.transport.send({ t: 'ops', id: 'evil-1', base: a.collab.session.rev, ops: [{ t: 'call', method: 'open', args: ['https://probe.invalid/doc.docx', {}] }] })
      a.transport.send({ t: 'ops', id: 'evil-2', base: a.collab.session.rev, ops: [{ t: 'call', method: 'save', args: [] }] })
      a.transport.send({ t: 'ops', id: 'evil-3', base: a.collab.session.rev, ops: [{ t: 'call', method: 'docx.page', args: [{ w: 400, evil: 1 }] }] })
      for (let i = 0; i < 100 && b.collab.session.rev < 3; i++) await new Promise((r) => setTimeout(r, 10))
      await b.collab.session.idle()
    } finally {
      console.error = quiet
    }
    expect(b.collab.session.rev).toBe(3)
    expect(b.calls).not.toContain('open')
    expect(b.calls).not.toContain('save')
    expect(b.calls).not.toContain('docx.page')
    expect(errors.length).toBe(3)
    expect(errors[0]).toMatch(/refused/)
    // the room goes on: a page setup one client asks for reaches the other
    b.calls.length = 0
    await b.client.call('docx.page', { w: 500 })
    for (let i = 0; i < 100 && !a.calls.includes('docx.page'); i++) await new Promise((r) => setTimeout(r, 10))
    expect(a.calls).toContain('docx.page')
    expect(b.calls).toContain('docx.page')
    a.collab.close()
    b.collab.close()
  })

  test('the color of a peer is a color for everyone in the room', async () => {
    const hub = new CollabHub()
    const adapter: Adapter<unknown> = { transform: (op) => op, apply: () => ({}), revert: () => {} }
    const bad = '#fff url(https://probe.invalid/peer.png)'
    const evil = new CollabSession(new MemoryTransport(hub, () => 1), adapter, { room: 'c', user: { name: 'Eve', color: bad } })
    const ann = new CollabSession(new MemoryTransport(hub, () => 1), adapter, { room: 'c', user: { name: 'Ann', color: 'teal' } })
    for (let i = 0; i < 200 && !(ann.peers.size === 1 && evil.peers.size === 1); i++) await new Promise((r) => setTimeout(r, 5))
    const seen = [...ann.peers.values()][0]
    expect(seen.name).toBe('Eve')
    expect(PEER_COLORS).toContain(seen.color)
    expect([...evil.peers.values()][0].color).toBe('teal')

    // a hub that relays what it is given (another implementation, an older one): the client checks as well
    const relay = new CollabSession(
      {
        connect() {
          this.onStatus?.('open')
          queueMicrotask(() => this.onMessage?.({ t: 'welcome', site: 1, seq: 0, from: 0, log: [], peers: [{ site: 2, name: 'Eve', color: bad }] }))
          queueMicrotask(() => this.onMessage?.({ t: 'join', peer: { site: 3, name: 'Mal', color: 'red;background:url(x)' } }))
        },
        send() {},
        close() {},
      } as ConstructorParameters<typeof CollabSession>[0],
      adapter,
      { room: 'c', user: { name: 'Ann' } },
    )
    for (let i = 0; i < 200 && relay.peers.size < 2; i++) await new Promise((r) => setTimeout(r, 5))
    for (const p of relay.peers.values()) expect(PEER_COLORS).toContain(p.color)
    evil.close()
    ann.close()
    relay.close()
  })
})
