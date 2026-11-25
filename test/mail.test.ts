import { describe, expect, test } from 'bun:test'
import type { Block, Paragraph, Table } from '../src/core/docx/types'
import { isMsg, mailTextFormat, openMail, readEml, readEmlx, readMsg } from '../src/core/mail'
import { decodeWords, parseAddresses, parseParams, unflow } from '../src/core/mail/mime'
import { pictureSize } from '../src/core/mail/picture'
import { decompressRtf, unwrapRtf } from '../src/core/mail/rtf'
import { BlobSource } from '../src/core/source'
import { Engine } from '../src/core/worker/engine'
import { ZipArchive } from '../src/core/zip'
import { writeCfb } from './cfb-writer'
import { b64, compressRtf, latin1, msgFile, png, sampleEml, sampleMsg, utf8 } from './mail-builder'

async function open(bytes: Uint8Array, name: string, params: Record<string, unknown> = {}) {
  const events: { event: string; data: any }[] = []
  const engine = new Engine((event, data) => events.push({ event, data }))
  const info = await engine.open(new Blob([bytes as Uint8Array<ArrayBuffer>]), { name, ...params })
  await engine.call('docx.start', [])
  const done = events.find((e) => e.event === 'docx:done')?.data
  if (done?.error) throw new Error(done.error)
  const blocks: Block[] = events.filter((e) => e.event === 'docx:blocks').flatMap((e) => e.data)
  return { engine, info, blocks }
}

const textOf = (b: Block): string =>
  b.t === 'p' ? b.runs.map((r) => (r.t === 'text' ? r.text : r.t === 'img' ? '<img>' : '')).join('') : b.t === 'tbl' ? b.rows.map((r) => r.cells.map((c) => c.blocks.map(textOf).join('/')).join(' | ')).join('\n') : ''
const allText = (blocks: Block[]) => blocks.map(textOf).join('\n')
const tables = (blocks: Block[]) => blocks.filter((b): b is Table => b.t === 'tbl')
const paras = (blocks: Block[]) => blocks.filter((b): b is Paragraph => b.t === 'p')

/** an Internet message of header lines and a body */
const eml = (headers: string[], body: string) => latin1([...headers, '', body].join('\r\n'))

describe('mail: Internet message headers', () => {
  test('encoded words: B and Q, charsets, a character split between two words, spaces between words dropped', () => {
    expect(decodeWords('=?UTF-8?B?w6l0w6k=?= report')).toBe('été report')
    expect(decodeWords('=?iso-8859-1?Q?Caf=E9_cr=E8me?=')).toBe('Café crème')
    // "€" (E2 82 AC) split over two words
    expect(decodeWords('=?utf-8?Q?Price_=E2=82?= =?utf-8?Q?=AC5?=')).toBe('Price €5')
    expect(decodeWords('=?utf-8?q?a?= b =?utf-8?q?c?=')).toBe('a b c')
    // 8-bit UTF-8 in a header (RFC 6532)
    expect(decodeWords(String.fromCharCode(...utf8('Grüße')))).toBe('Grüße')
    // an unknown charset: read as UTF-8
    expect(decodeWords('=?x-unknown?Q?plain?=')).toBe('plain')
  })

  test('addresses: names quoted, with commas, encoded, in comments; groups; mailto', () => {
    expect(parseAddresses('"Lee, Mei" <mei@x.org>, Bob <bob@y.com>, ann@z.org (Ann Z), plain@w.net')).toEqual([
      { name: 'Lee, Mei', addr: 'mei@x.org' },
      { name: 'Bob', addr: 'bob@y.com' },
      { name: 'Ann Z', addr: 'ann@z.org' },
      { name: '', addr: 'plain@w.net' },
    ])
    expect(parseAddresses('=?utf-8?Q?Doe=2C_J=C3=BCrgen?= <j@x.de>')).toEqual([{ name: 'Doe, Jürgen', addr: 'j@x.de' }])
    expect(parseAddresses('Team: a@x.org, b@x.org;, undisclosed-recipients:;')).toEqual([
      { name: '', addr: 'a@x.org' },
      { name: '', addr: 'b@x.org' },
    ])
    expect(parseAddresses('<mailto:m@x.org>')).toEqual([{ name: '', addr: 'm@x.org' }])
  })

  test('parameters: quoted, RFC 2231 encoded and continued, encoded words in a quoted name', () => {
    expect(parseParams('text/HTML; charset="utf-8"; format=flowed')).toEqual({ value: 'text/html', params: { charset: 'utf-8', format: 'flowed' } } as any)
    expect(parseParams("attachment; filename*=UTF-8''r%C3%A9sum%C3%A9.pdf").params.filename).toBe('résumé.pdf')
    expect(parseParams("attachment; filename*0*=UTF-8''long%20na; filename*1=me.txt").params.filename).toBe('long name.txt')
    expect(parseParams('attachment; filename="=?UTF-8?B?w7xiZXIucGRm?="').params.filename).toBe('über.pdf')
  })

  test('format=flowed: soft breaks joined, quotes kept apart, space-stuffing undone', () => {
    expect(unflow('A line that \r\ngoes on.\r\n> quoted \r\n> more\r\n plain', false)).toBe('A line that goes on.\n> quoted more\nplain')
    expect(unflow('one \r\ntwo', true)).toBe('onetwo')
  })
})

describe('mail: Internet message structure', () => {
  test('the sample: subject, people, date, the HTML of the alternatives, the related picture shown, the CSV attached', () => {
    const m = readEml(sampleEml())
    expect(m.subject).toBe('Quarterly review – agenda')
    expect(m.from).toEqual([{ name: 'Ana Martínez', addr: 'ana@example.org' }])
    expect(m.to.map((a) => a.name)).toEqual(['Jonas Weber', 'Lee, Mei'])
    expect(m.cc).toEqual([{ name: '', addr: 'team@example.com' }])
    expect(m.date?.toISOString()).toBe('2026-11-19T08:14:00.000Z')
    expect(m.body.length).toBe(1)
    expect('html' in m.body[0] && m.body[0].html).toContain('<h2>Quarterly review: agenda</h2>')
    expect(m.attachments.map((a) => [a.name, a.type, a.cid])).toEqual([
      ['logo.png', 'image/png', 'logo@example.org'],
      ['q3-figures.csv', 'text/csv', undefined],
    ])
  })

  test('a plain message: no MIME headers, LF line ends, quoted-printable in another charset', () => {
    const m = readEml(latin1('From: a@x.org\nTo: b@y.org\nSubject: Hi\nContent-Type: text/plain; charset=iso-8859-1\nContent-Transfer-Encoding: quoted-printable\n\nCaf=E9 =\nat noon'))
    expect(m.body).toEqual([{ text: 'Café at noon' }])
  })

  test('mixed parts: text after an attachment goes on in the body; signatures dropped; a forwarded message read', () => {
    const inner = ['From: Old <old@x.org>', 'Subject: Original', 'Date: Mon, 2 Nov 2026 10:00:00 +0000', '', 'The first message.'].join('\r\n')
    const m = readEml(
      eml(
        ['From: a@x.org', 'Subject: Fwd', 'Content-Type: multipart/signed; boundary=s; protocol="application/pkcs7-signature"'],
        [
          '--s',
          'Content-Type: multipart/mixed; boundary=m',
          '',
          '--m',
          'Content-Type: text/plain',
          '',
          'Before.',
          '--m',
          'Content-Type: application/octet-stream; name=data.bin',
          'Content-Transfer-Encoding: base64',
          '',
          b64(Uint8Array.of(1, 2, 3)),
          '--m',
          'Content-Type: text/plain',
          'Content-Disposition: inline',
          '',
          'After.',
          '--m',
          'Content-Type: message/rfc822',
          '',
          inner,
          '--m--',
          '--s',
          'Content-Type: application/pkcs7-signature; name=smime.p7s',
          '',
          'MIAG',
          '--s--',
        ].join('\r\n'),
      ),
    )
    expect(m.body).toEqual([{ text: 'Before.' }, { text: 'After.' }])
    expect(m.attachments.map((a) => a.name)).toEqual(['data.bin', 'Original.eml'])
    expect([...m.attachments[0].data!]).toEqual([1, 2, 3])
    expect(m.attachments[1].message).toMatchObject({ subject: 'Original', from: [{ name: 'Old', addr: 'old@x.org' }], body: [{ text: 'The first message.' }] })
  })

  test('a message cut short still shows what is there; an encrypted one says so; no date is no date', () => {
    const cut = readEml(eml(['Subject: Cut', 'Content-Type: multipart/alternative; boundary=b'], '--b\r\nContent-Type: text/plain\r\n\r\nPartial te'))
    expect(cut.body).toEqual([{ text: 'Partial te' }])
    const enc = readEml(eml(['Subject: Secret', 'Content-Type: application/pkcs7-mime; smime-type=enveloped-data; name=smime.p7m', 'Content-Transfer-Encoding: base64'], 'MIAGCSqG'))
    expect(enc.encrypted).toBe(true)
    const odd = readEml(eml(['Subject: x', 'Date: sometime soon'], 'x'))
    expect(odd.date).toBeUndefined()
    expect(odd.dateText).toBe('sometime soon')
  })

  test('emlx: the length line and the property list after the message are left out', () => {
    const msg = 'From: a@x.org\nSubject: Apple\n\nBody.\n'
    const m = readEmlx(latin1(`${msg.length}\n${msg}<?xml version="1.0"?><plist><dict/></plist>`))
    expect(m.subject).toBe('Apple')
    expect(m.body).toEqual([{ text: 'Body.\n' }])
  })
})

describe('mail: Outlook items', () => {
  test('the sample: sender, recipients by type, sent time, importance, HTML body, attachments in order, the attached item', () => {
    const m = readMsg(sampleMsg())
    expect(m.subject).toBe('Contract renewal – signed copy')
    expect(m.from).toEqual([{ name: 'Jonas Weber', addr: 'jonas@example.com' }])
    expect(m.to).toEqual([{ name: 'Ana Martínez', addr: 'ana@example.org' }])
    expect(m.cc).toEqual([{ name: 'Legal', addr: 'legal@example.org' }])
    expect(m.date?.toISOString()).toBe('2026-11-20T15:42:00.000Z')
    expect(m.importance).toBe('high')
    expect('html' in m.body[0]).toBe(true)
    expect(m.attachments.map((a) => [a.name, a.hidden ?? false, !!a.message])).toEqual([
      ['contract-signed.pdf', false, false],
      ['signature.png', true, false],
      ['Original request.msg', false, true],
    ])
    expect(m.attachments[2].message).toMatchObject({ subject: 'Contract renewal', body: [{ text: expect.stringContaining('signed renewal') }] })
  })

  test('an Exchange sender has no address to show; recipients from the display strings when there are no recipient storages', () => {
    const m = readMsg(msgFile({ subject: 'x', fromName: 'Pat', fromExchange: '/O=EXCHANGELABS/OU=GROUP/CN=RECIPIENTS/CN=PAT', extra: { 0x0e04: { str: 'Ann Lee; Bob' } } }))
    expect(m.from).toEqual([{ name: 'Pat', addr: '' }])
    expect(m.to).toEqual([
      { name: 'Ann Lee', addr: '' },
      { name: 'Bob', addr: '' },
    ])
  })

  test('message classes other than a note are said; a compound file that is no Outlook item is not one', () => {
    expect(readMsg(msgFile({ subject: 'Sync', messageClass: 'IPM.Schedule.Meeting.Request' })).itemClass).toBe('IPM.Schedule.Meeting.Request')
    expect(readMsg(msgFile({ subject: 'x' })).itemClass).toBeUndefined()
    expect(isMsg(sampleMsg())).toBe(true)
    expect(isMsg(writeCfb({ WordDocument: new Uint8Array(512) }))).toBe(false)
    expect(isMsg(latin1('not a compound file'))).toBe(false)
  })
})

describe('mail: Outlook RTF bodies', () => {
  const htmlRtf = String.raw`{\rtf1\ansi\ansicpg1252\fromhtml1 \deff0{\fonttbl{\f0\fswiss\fcharset0 Arial;}}` +
    String.raw`{\*\htmltag19 <html>}{\*\htmltag34 <head>}{\*\htmltag41 </head>}{\*\htmltag50 <body>}` +
    String.raw`\htmlrtf {\htmlrtf0{\*\htmltag64 <p>}Caf\'e9 荤?5 {\*\htmltag84 <b>}bold{\*\htmltag92 </b>}` +
    String.raw`\htmlrtf\par\htmlrtf0{\*\htmltag72 </p>}\htmlrtf }\htmlrtf0{\*\htmltag58 </body>}{\*\htmltag27 </html>}}`

  test('compressed RTF: LZFu (literal runs, the end reference) and uncompressed (MELA)', () => {
    const rtf = latin1(htmlRtf)
    expect(new TextDecoder().decode(decompressRtf(compressRtf(rtf)))).toBe(htmlRtf)
    expect(new TextDecoder().decode(decompressRtf(compressRtf(rtf, 'MELA')))).toBe(htmlRtf)
    // references into the initial dictionary: "{\rtf1\ansi\mac" is its first 15 bytes (offset 0, length 15)
    const head = new Uint8Array(16)
    new DataView(head.buffer).setUint32(4, 15, true)
    new DataView(head.buffer).setUint32(8, 0x75465a4c, true)
    const ref = Uint8Array.from([0x03, 0x00, 0x0d, 0x0c, 0xf0])
    new DataView(head.buffer).setUint32(0, ref.length + 12, true)
    const out = decompressRtf(new Uint8Array([...head, ...ref]))
    expect(new TextDecoder().decode(out)).toBe('{\\rtf1\\ansi\\mac')
    expect(() => decompressRtf(new Uint8Array(4))).toThrow()
  })

  test('the HTML an RTF body wraps: its tags, its text in the code page and \\u, the RTF-only text left out', () => {
    const r = unwrapRtf(latin1(htmlRtf))
    expect('html' in r && r.html.replace(/\s+/g, ' ')).toBe('<html><head></head><body><p>Café €5 <b>bold</b></p></body></html>')
  })

  test('plain text wrapped (\\fromtext) and RTF of its own', () => {
    expect(unwrapRtf(latin1(String.raw`{\rtf1\ansi\fromtext{\fonttbl{\f0 Courier;}}Line one\par Line two}`))).toEqual({ text: 'Line one\nLine two' })
    expect('rtf' in unwrapRtf(latin1(String.raw`{\rtf1\ansi Plain {\b rich} text}`))).toBe(true)
  })

  test('an item whose body is only compressed RTF: the wrapped HTML is shown; RTF of its own is converted with the header around it', async () => {
    const wrapped = await open(msgFile({ subject: 'From RTF', rtf: compressRtf(latin1(htmlRtf)) }), 'r.msg')
    expect(allText(wrapped.blocks)).toContain('Café €5 bold')
    const own = await open(msgFile({ subject: 'Rich', fromName: 'Ann', body: 'fallback', rtf: compressRtf(latin1(String.raw`{\rtf1\ansi\deff0{\fonttbl{\f0 Arial;}}Some {\b bold} words\par}`)) }), 'own.msg')
    const ps = paras(own.blocks)
    expect(textOf(ps[0])).toBe('Rich')
    expect(tables(own.blocks)[0].rows[0].cells.map((c) => c.blocks.map(textOf).join(''))).toEqual(['From', 'Ann'])
    const body = ps.find((p) => textOf(p) === 'Some bold words')!
    expect(body.runs.find((r) => r.t === 'text' && r.text === 'bold')).toMatchObject({ s: { b: true } })
  })
})

describe('mail: the document', () => {
  test('opened by the engine: read-only, the attachments listed and saved; the pictures the body shows are not', async () => {
    const { engine, info, blocks } = await open(sampleEml(), 'sample.eml')
    expect(info).toMatchObject({ kind: 'docx', from: 'eml', mail: { subject: 'Quarterly review – agenda', attachments: [{ name: 'q3-figures.csv', type: 'text/csv', size: 66 }] } })
    const header = tables(blocks)[0]
    expect(header.rows.map((r) => textOf(r.cells[0].blocks[0] as Paragraph))).toEqual(['From', 'Sent', 'To', 'Cc', 'Attachments'])
    // the logo is a picture of the body, at the size the HTML gives
    const img = paras(blocks).flatMap((p) => p.runs).find((r) => r.t === 'img')
    expect(img).toMatchObject({ t: 'img', w: 90, h: 30 })
    // a mailto: link on each address
    expect(paras(header.rows[0].cells[1].blocks).flatMap((p) => p.runs).some((r) => r.t === 'text' && r.href === 'mailto:ana@example.org')).toBe(true)
    // the agenda table keeps its lines and the shading of its header row
    const agenda = tables(blocks)[1]
    expect(agenda.rows[0].cells.map((c) => c.shd)).toEqual(['#e8eef9', '#e8eef9', '#e8eef9'])
    const csv = (await engine.call('mail.attachment', [0])) as Blob
    expect(await csv.text()).toStartWith('Quarter,Revenue')
    expect(csv.type).toBe('text/csv')
    await expect(engine.call('mail.attachment', [5])).rejects.toThrow('No such attachment')
  })

  test('the header in the host’s words and locale', async () => {
    const { blocks } = await open(sampleMsg(), 'sample.msg', { mail: { texts: { from: 'Von', sent: 'Gesendet', item: { report: 'Bericht' } }, locale: 'de-DE' } })
    const text = allText(blocks)
    expect(text).toContain('Von | Jonas Weber')
    expect(text).toContain('Gesendet | Freitag, 20. November 2026')
    // the words not given stay English
    expect(text).toContain('To | Ana Martínez')
  })

  test('remote pictures are left out and said; links that run script are not links', async () => {
    const m = eml(['From: a@x.org', 'Subject: Web', 'Content-Type: text/html'], '<p><img src="https://t.example/p.gif"><img src=//cdn.example/x.png>Hi <a href="javascript:alert(1)">there</a></p>')
    const { blocks } = await open(m, 'web.eml')
    const text = allText(blocks)
    expect(text).toContain('Pictures from the web are not shown')
    expect(paras(blocks).flatMap((p) => p.runs).some((r) => r.t === 'img')).toBe(false)
    expect(paras(blocks).flatMap((p) => p.runs).some((r) => r.t === 'text' && r.href)).toBe(false)
  })

  test('a plain-text body: a paragraph per line, web addresses linked, quoted lines grey', async () => {
    const { blocks } = await open(eml(['From: a@x.org', 'Subject: Text'], 'See https://example.org/a.\r\n> earlier\r\nEnd'), 'text.eml')
    const body = paras(blocks).slice(-3)
    expect(body.map(textOf)).toEqual(['See https://example.org/a.', '> earlier', 'End'])
    expect(body[0].runs.find((r) => r.t === 'text' && r.href)).toMatchObject({ text: 'https://example.org/a' })
    expect(body[1].runs[0]).toMatchObject({ s: { color: '#595959' } })
  })

  test('found by its content (no name, a MIME type), not mistaken: web pages saved as MHTML, other text', () => {
    const head = latin1('Received: from x\r\nFrom: a@x.org\r\nTo: b@y.org\r\nSubject: s\r\n\r\nbody')
    expect(mailTextFormat(undefined, '', head)).toBe('eml')
    expect(mailTextFormat('download', 'message/rfc822', latin1('junk'))).toBe('eml')
    expect(mailTextFormat('page.mht', '', head)).toBe(null)
    expect(mailTextFormat(undefined, '', latin1('From: <Saved by Blink>\r\nSnapshot-Content-Location: https://x\r\nSubject: p\r\nMIME-Version: 1.0\r\n\r\n'))).toBe(null)
    expect(mailTextFormat(undefined, '', latin1('Name: value\r\nOther: thing\r\n'))).toBe(null)
  })

  test('damaged or odd files: no crash, a message that says so', async () => {
    await expect(open(latin1('garbage'), 'x.msg')).rejects.toThrow('not an Outlook message')
    // (some programs save Internet messages as .msg)
    const mime = await open(eml(['From: a@x.org', 'To: b@y.org', 'Subject: Saved as msg'], 'Body'), 'saved.msg')
    expect(mime.info.from).toBe('eml')
    // a damaged compound file
    const msg = sampleMsg()
    await expect(open(msg.subarray(0, 600), 'cut.msg')).rejects.toThrow()
    // an empty .eml is an empty message
    const empty = await open(new Uint8Array(0), 'empty.eml')
    expect(allText(empty.blocks)).toContain('(no subject)')
  })

  test('pictures: the size of PNG, GIF, JPEG and BMP read from their bytes', async () => {
    expect(pictureSize(png(7, 3))).toEqual({ w: 7, h: 3 })
    expect(pictureSize(latin1('GIF89a\x05\x00\x04\x00' + '\0'.repeat(30)))).toEqual({ w: 5, h: 4 })
    const jpeg = Uint8Array.from([0xff, 0xd8, 0xff, 0xe0, 0, 4, 0, 0, 0xff, 0xc0, 0, 11, 8, 0, 20, 0, 30, 3, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0])
    expect(pictureSize(jpeg)).toEqual({ w: 30, h: 20 })
    const bmp = new Uint8Array(40)
    bmp.set([0x42, 0x4d])
    new DataView(bmp.buffer).setInt32(18, 9, true)
    new DataView(bmp.buffer).setInt32(22, -6, true)
    expect(pictureSize(bmp)).toEqual({ w: 9, h: 6 })
    // the built document has the PNG as a part of its own
    const { docx } = await openMail(sampleEml(), 'eml')
    const zip = await ZipArchive.open(new BlobSource(docx))
    expect(zip.list().map((e) => e.name)).toContain('word/media/image1.png')
  })
})
