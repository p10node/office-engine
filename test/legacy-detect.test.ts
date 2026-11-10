import { describe, expect, test } from 'bun:test'
import { Cfb, isCfb } from '../src/core/legacy/cfb'
import { convertCfb, convertOdfZip, isFlatOdf, isRtf, LEGACY_EXT, LEGACY_KIND, legacyFormatOf } from '../src/core/legacy'
import { odfFormatOf } from '../src/core/legacy/odf'
import { OdfPackage, parseOdfXml, textContent, toPt } from '../src/core/legacy/odf/common'
import { mayBePlain, plainFormat } from '../src/core/plain'
import { BlobSource } from '../src/core/source'
import { ZipArchive } from '../src/core/zip'
import { writeZip } from '../src/core/zipwrite'
import { Engine } from '../src/core/worker/engine'
import { writeCfb } from './cfb-writer'

const enc = new TextEncoder()
const bytes = (n: number, seed = 1) => Uint8Array.from({ length: n }, (_, i) => (i * 31 + seed) & 0xff)

describe('compound file reader', () => {
  test('small (mini stream) and large streams, nested storages, case-insensitive paths', () => {
    const big = bytes(10_000)
    const small = bytes(100, 7)
    const file = writeCfb({ WordDocument: big, '1Table': small, 'ObjectPool/_123/\u0001Ole': bytes(20, 3), Empty: new Uint8Array(0) })
    expect(isCfb(file)).toBe(true)
    const cfb = Cfb.open(file)
    expect(cfb.stream('WordDocument')).toEqual(big)
    expect(cfb.stream('1table')).toEqual(small)
    expect(cfb.stream('ObjectPool/_123/\u0001Ole')).toEqual(bytes(20, 3))
    expect(cfb.stream('Empty')).toEqual(new Uint8Array(0))
    expect(cfb.stream('Missing')).toBeUndefined()
    // a storage is not a stream
    expect(cfb.stream('ObjectPool')).toBeUndefined()
    expect(cfb.has('objectpool/_123')).toBe(true)
    expect(cfb.children('').map((e) => e.name).sort()).toEqual(['1Table', 'Empty', 'ObjectPool', 'WordDocument'])
    expect(cfb.children('ObjectPool/_123').map((e) => e.path)).toEqual(['ObjectPool/_123/\u0001Ole'])
  })

  test('many streams spread over several directory and mini FAT sectors', () => {
    const streams: Record<string, Uint8Array> = {}
    for (let i = 0; i < 40; i++) streams[`S${i}`] = bytes(50 + i * 90, i)
    const cfb = Cfb.open(writeCfb(streams))
    for (const [k, v] of Object.entries(streams)) expect(cfb.stream(k)).toEqual(v)
  })

  test('damaged files fail instead of looping or reading past the end', () => {
    expect(() => Cfb.open(enc.encode('not a compound file at all'))).toThrow()
    const file = writeCfb({ WordDocument: bytes(10_000) })
    // a FAT that points the big stream's first sector back to itself
    const loop = file.slice()
    const dv = new DataView(loop.buffer)
    const fatAt = 512
    const first = [...Array(128).keys()].find((i) => i > 1 && dv.getUint32(fatAt + i * 4, true) === i + 1)!
    dv.setUint32(fatAt + first * 4, first, true)
    expect(() => Cfb.open(loop).stream('WordDocument')).toThrow(/loops/)
    // truncated: the reader keeps to the bytes it has
    const cut = file.slice(0, 1536)
    expect(() => Cfb.open(cut).stream('WordDocument')).not.toThrow()
  })

  test('FAT sectors the header counts but does not list (textutil past ~7 MB) follow the listed ones', () => {
    const big = bytes(200_000)
    const file = writeCfb({ WordDocument: big })
    const dv = new DataView(file.buffer)
    expect(dv.getUint32(0x2c, true)).toBe(4)
    // only the first two of four FAT sectors listed, no DIFAT sector
    dv.setUint32(0x4c + 2 * 4, 0xffffffff, true)
    dv.setUint32(0x4c + 3 * 4, 0xffffffff, true)
    expect(Cfb.open(file).stream('WordDocument')).toEqual(big)
  })

  test('a compound file without a document stream, and an encrypted package, say so', async () => {
    await expect(convertCfb(writeCfb({ Other: bytes(10) }))).rejects.toThrow(/Not a Word, Excel or PowerPoint file/)
    await expect(convertCfb(writeCfb({ EncryptionInfo: bytes(10), EncryptedPackage: bytes(5000) }))).rejects.toThrow(/Password-protected/)
    await expect(convertCfb(writeCfb({ PP40: bytes(100) }))).rejects.toThrow(/PowerPoint 95 and older/)
  })
})

describe('format detection', () => {
  test('first bytes: compound file, RTF, flat OpenDocument; anything else is not a legacy format', () => {
    expect(legacyFormatOf(writeCfb({ A: bytes(1) }))).toBe('cfb')
    expect(legacyFormatOf(enc.encode('{\\rtf1\\ansi Hello}'))).toBe('rtf')
    expect(isRtf(enc.encode('{\\rtf1'))).toBe(true)
    // a byte order mark, blank space before it, the Unicode variant
    expect(isRtf(enc.encode('\ufeff{\\rtf1'))).toBe(true)
    expect(isRtf(enc.encode('\r\n {\\rtf1'))).toBe(true)
    expect(isRtf(enc.encode('{\\urtf1'))).toBe(true)
    expect(isRtf(enc.encode('{\\rtx'))).toBe(false)
    const fodt = '<?xml version="1.0" encoding="UTF-8"?>\n<office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" office:mimetype="application/vnd.oasis.opendocument.text"></office:document>'
    expect(legacyFormatOf(enc.encode(fodt))).toBe('flat-odf')
    expect(isFlatOdf(enc.encode('﻿' + fodt))).toBe(true)
    expect(legacyFormatOf(enc.encode('<html><body>hi</body></html>'))).toBeNull()
    expect(legacyFormatOf(enc.encode('a,b\n1,2'))).toBeNull()
    // a bare Excel stream: BOF record 0x0809, length 16, BIFF8 workbook globals
    expect(legacyFormatOf(Uint8Array.of(0x09, 0x08, 0x10, 0x00, 0x00, 0x06, 0x05, 0x00, 0, 0, 0, 0, 0, 0, 0, 0))).toBe('biff')
    // a tab-separated file starting with a tab is text
    expect(legacyFormatOf(enc.encode('\tA\tB\n1\t2'))).toBeNull()
    expect(legacyFormatOf(Uint8Array.of(0x50, 0x4b, 3, 4))).toBeNull()
  })

  test('extensions: every format opens as the right kind and is never read as text', () => {
    for (const [ext, f] of Object.entries(LEGACY_EXT)) {
      expect(LEGACY_KIND[f]).toBeDefined()
      // .dot is also Graphviz: the bytes decide
      if (ext !== 'dot') expect(mayBePlain(`x.${ext}`)).toBe(false)
    }
    // an RTF file named .txt is still RTF to the engine (checked before plain text); plainFormat leaves RTF names alone
    expect(plainFormat('x.rtf', undefined, enc.encode('{\\rtf1 hi}'), true)).toBeNull()
    expect(plainFormat('x.odt', undefined, Uint8Array.of(0x50, 0x4b, 3, 4), false)).toBeNull()
    // a text .dot is a Graphviz graph, a compound .dot a Word template
    expect(plainFormat('g.dot', undefined, enc.encode('digraph { a -> b }'), true)).toBe('graphviz')
    expect(plainFormat('t.dot', undefined, writeCfb({ WordDocument: bytes(10) }), false)).toBeNull()
  })

  test('OpenDocument packages by mimetype or manifest, templates too; flat XML by office:mimetype', async () => {
    const zip = async (parts: Record<string, string>) => ZipArchive.open(new BlobSource(await writeZip(new Map(Object.entries(parts)), 'application/zip')))
    expect(await odfFormatOf(await zip({ mimetype: 'application/vnd.oasis.opendocument.text', 'content.xml': '<x/>' }))).toBe('odt')
    expect(await odfFormatOf(await zip({ mimetype: 'application/vnd.oasis.opendocument.spreadsheet-template', 'content.xml': '<x/>' }))).toBe('ods')
    const manifest =
      '<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.presentation"/></manifest:manifest>'
    expect(await odfFormatOf(await zip({ 'META-INF/manifest.xml': manifest, 'content.xml': '<x/>' }))).toBe('odp')
    // a drawing, a formula, or a ZIP that is not OpenDocument
    expect(await odfFormatOf(await zip({ mimetype: 'application/vnd.oasis.opendocument.graphics', 'content.xml': '<x/>' }))).toBeNull()
    expect(await odfFormatOf(await zip({ 'word/document.xml': '<x/>' }))).toBeNull()
    expect(await odfFormatOf('<office:document office:mimetype="application/vnd.oasis.opendocument.spreadsheet">')).toBe('ods')
    await expect(convertOdfZip(await zip({ mimetype: 'application/vnd.sun.xml.writer', 'content.xml': '<x/>' }))).rejects.toThrow(/OpenOffice.org 1.x/)
    expect(await convertOdfZip(await zip({ 'word/document.xml': '<x/>' }))).toBeNull()
  })
})

describe('OpenDocument plumbing', () => {
  test('mixed content in order, canonical prefixes whatever the file declares, entities', () => {
    const root = parseOdfXml(
      '<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0">' +
        '<t:p t:style-name="P1">Fish &amp; chips <t:span>hot</t:span><t:s t:c="2"/>now<t:tab/>!<t:line-break/>end</t:p></o:document-content>',
    )
    expect(root.name).toBe('office:document-content')
    const p = root.kids[0] as { name: string; attrs: Record<string, string> }
    expect(p.name).toBe('text:p')
    expect(p.attrs['text:style-name']).toBe('P1')
    expect(textContent(root)).toBe('Fish & chips hot  now\t!\nend')
  })

  test('comments, CDATA, processing instructions; a namespace declared on an inner element stays in its subtree', () => {
    const root = parseOdfXml(
      '<?xml version="1.0"?><!-- head --><root xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0">' +
        '<a xmlns:x="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><x:p x:style-name="S">one<![CDATA[ <two> ]]><!-- c --></x:p></a>' +
        '<x:p/><b xmlns:t="urn:oasis:names:tc:opendocument:xmlns:table:1.0"/><t:table/><office:x/></root>',
    )
    const kids = root.kids.filter((k): k is Exclude<typeof k, string> => typeof k !== 'string').map((k) => k.name)
    expect(kids).toEqual(['a', 'x:p', 'b', 't:table', 'office:x'])
    const p = (root.kids[0] as typeof root).kids[0] as typeof root
    expect(p.name).toBe('text:p')
    expect(p.attrs).toEqual({ 'text:style-name': 'S' })
    expect(textContent(p)).toBe('one <two> ')
  })

  test('a password-protected package says so', async () => {
    const manifest =
      '<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/>' +
      '<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"><manifest:encryption-data manifest:checksum-type="SHA1/1K"/></manifest:file-entry></manifest:manifest>'
    const zip = await ZipArchive.open(new BlobSource(await writeZip(new Map(Object.entries({ mimetype: 'application/vnd.oasis.opendocument.text', 'content.xml': 'garbage', 'META-INF/manifest.xml': manifest })), 'application/zip')))
    await expect(OdfPackage.open(zip)).rejects.toThrow(/Password-protected/)
    // a password on a macro library only: the document opens
    const macro = manifest.replace('content.xml', 'Basic/Standard/Module1.xml')
    const zip2 = await ZipArchive.open(new BlobSource(await writeZip(new Map(Object.entries({ mimetype: 'application/vnd.oasis.opendocument.text', 'content.xml': '<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"/>', 'META-INF/manifest.xml': macro })), 'application/zip')))
    expect((await OdfPackage.open(zip2)).content.name).toBe('office:document-content')
  })

  test('lengths in points', () => {
    expect(toPt('2.54cm')).toBeCloseTo(72)
    expect(toPt('10mm')).toBeCloseTo(28.346, 2)
    expect(toPt('0.5in')).toBe(36)
    expect(toPt('12pt')).toBe(12)
    expect(toPt('1pc')).toBe(12)
    expect(toPt('50%')).toBeUndefined()
  })

  test('styles inherit from their parents and the family default; automatic content styles win', async () => {
    const NS = 'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"'
    const flat =
      `<office:document ${NS} office:mimetype="application/vnd.oasis.opendocument.text">` +
      '<office:styles><style:default-style style:family="paragraph"><style:text-properties fo:font-size="12pt" fo:color="#000000"/></style:default-style>' +
      '<style:style style:name="Base" style:family="paragraph"><style:text-properties fo:font-weight="bold"/></style:style>' +
      '<style:style style:name="Child" style:family="paragraph" style:parent-style-name="Base"><style:text-properties fo:color="#ff0000"/></style:style></office:styles>' +
      '<office:automatic-styles><style:style style:name="P1" style:family="paragraph" style:parent-style-name="Child"><style:paragraph-properties fo:text-align="center"/></style:style></office:automatic-styles>' +
      '<office:body><office:text><text:p text:style-name="P1">x</text:p></office:text></office:body></office:document>'
    const pkg = await OdfPackage.open(flat)
    const props = pkg.styles.props('paragraph', 'P1')
    expect(props['style:text-properties']).toEqual({ 'fo:font-size': '12pt', 'fo:color': '#ff0000', 'fo:font-weight': 'bold' })
    expect(props['style:paragraph-properties']).toEqual({ 'fo:text-align': 'center' })
    expect(pkg.styles.chain('paragraph', 'P1').map((s) => s.attrs['style:name'] ?? 'default')).toEqual(['P1', 'Child', 'Base', 'default'])
    expect(textContent(pkg.body('text')!)).toBe('x')
  })
})

describe('engine', () => {
  test('a bare Excel stream (no compound file) opens as a workbook', async () => {
    const xls = new Uint8Array(await Bun.file(`${import.meta.dir}/fixtures/legacy/sample.xls`).arrayBuffer())
    const stream = Cfb.open(xls).stream('Workbook')!
    const engine = new Engine(() => {})
    const info = await engine.open(new Blob([stream as Uint8Array<ArrayBuffer>]), { name: 'export.xls' })
    expect(info.kind).toBe('xlsx')
    expect(info.converted).toBe('xls')
    expect(info.xlsx?.sheets.length).toBeGreaterThan(0)
  })
})
