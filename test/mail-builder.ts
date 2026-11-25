/**
 * Test helper: mail messages. Internet messages (.eml) are written as text; Outlook items (.msg) as a compound file
 * of MAPI properties (MS-OXMSG), with an RTF body compressed as Outlook does (MS-OXRTFCP). Also the sample messages
 * of the playground (scripts/mail-samples.ts).
 */
import { deflateSync } from 'node:zlib'
import { crc32 } from '../src/core/zipwrite'
import { writeCfb } from './cfb-writer'

/** each character one byte */
export const latin1 = (s: string) => Uint8Array.from(s, (c) => c.charCodeAt(0) & 0xff)
export const utf8 = (s: string) => new TextEncoder().encode(s)
export const b64 = (b: Uint8Array) => Buffer.from(b).toString('base64').replace(/(.{76})/g, '$1\r\n')

/** a PNG of w × h pixels of one color */
export function png(w: number, h: number, rgb: [number, number, number] = [0x25, 0x63, 0xeb]): Uint8Array {
  const chunk = (type: string, data: Uint8Array) => {
    const body = new Uint8Array(4 + data.length)
    body.set(latin1(type))
    body.set(data, 4)
    const out = new Uint8Array(12 + data.length)
    new DataView(out.buffer).setUint32(0, data.length)
    out.set(body, 4)
    new DataView(out.buffer).setUint32(8 + data.length, crc32(body))
    return out
  }
  const ihdr = new Uint8Array(13)
  new DataView(ihdr.buffer).setUint32(0, w)
  new DataView(ihdr.buffer).setUint32(4, h)
  ihdr.set([8, 2, 0, 0, 0], 8)
  const raw = new Uint8Array(h * (1 + w * 3))
  for (let y = 0; y < h; y++) for (let x = 0; x < w; x++) raw.set(rgb, y * (1 + w * 3) + 1 + x * 3)
  return concat([Uint8Array.of(0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a), chunk('IHDR', ihdr), chunk('IDAT', new Uint8Array(deflateSync(raw))), chunk('IEND', new Uint8Array())])
}

export function concat(parts: Uint8Array[]): Uint8Array<ArrayBuffer> {
  const out = new Uint8Array(parts.reduce((n, p) => n + p.length, 0))
  let k = 0
  for (const p of parts) {
    out.set(p, k)
    k += p.length
  }
  return out
}

// ---------------------------------------------------------------- compressed RTF

const PREBUF_LEN = 207

/** PR_RTF_COMPRESSED of `rtf`: LZFu of literal bytes only (valid, if not small), ended by a reference to the write position */
export function compressRtf(rtf: Uint8Array, kind: 'LZFu' | 'MELA' = 'LZFu'): Uint8Array {
  const head = (comp: number, type: number) => {
    const h = new Uint8Array(16)
    const dv = new DataView(h.buffer)
    dv.setUint32(0, comp, true)
    dv.setUint32(4, rtf.length, true)
    dv.setUint32(8, type, true)
    return h
  }
  if (kind === 'MELA') return concat([head(rtf.length + 12, 0x414c454d), rtf])
  const body: number[] = []
  let w = PREBUF_LEN
  let k = 0
  for (;;) {
    let control = 0
    const run: number[] = []
    let bit = 0
    for (; bit < 8 && k < rtf.length; bit++) {
      run.push(rtf[k++])
      w = (w + 1) & 4095
    }
    let ended = false
    if (bit < 8) {
      control |= 1 << bit
      run.push((w >> 4) & 0xff, (w << 4) & 0xf0)
      ended = true
    }
    body.push(control, ...run)
    if (ended) break
  }
  return concat([head(body.length + 12, 0x75465a4c), Uint8Array.from(body)])
}

// ---------------------------------------------------------------- .msg

type Value = { str: string } | { bin: Uint8Array } | { long: number } | { bool: boolean } | { time: Date }

const hex4 = (n: number) => n.toString(16).toUpperCase().padStart(4, '0')

/** a storage's streams: its property stream (with a header of `headerSize` bytes) and the variable-size values */
function props(dir: string, values: Record<number, Value | undefined>, headerSize: number, header?: (dv: DataView) => void): Record<string, Uint8Array> {
  const out: Record<string, Uint8Array> = {}
  const entries = Object.entries(values).filter((e): e is [string, Value] => e[1] !== undefined)
  const stream = new Uint8Array(headerSize + entries.length * 16)
  const dv = new DataView(stream.buffer)
  header?.(dv)
  entries.forEach(([idText, v], k) => {
    const id = Number(idText)
    const at = headerSize + k * 16
    let type: number
    if ('str' in v) {
      type = 0x001f
      const s = new Uint8Array((v.str.length + 1) * 2)
      for (let i = 0; i < v.str.length; i++) new DataView(s.buffer).setUint16(i * 2, v.str.charCodeAt(i), true)
      out[`${dir}__substg1.0_${hex4(id)}${hex4(type)}`] = s
      dv.setUint32(at + 8, s.length, true)
    } else if ('bin' in v) {
      type = 0x0102
      out[`${dir}__substg1.0_${hex4(id)}${hex4(type)}`] = v.bin
      dv.setUint32(at + 8, v.bin.length, true)
    } else if ('long' in v) {
      type = 0x0003
      dv.setInt32(at + 8, v.long, true)
    } else if ('bool' in v) {
      type = 0x000b
      dv.setUint8(at + 8, v.bool ? 1 : 0)
    } else {
      type = 0x0040
      const ticks = (v.time.getTime() + 11644473600000) * 10000
      dv.setUint32(at + 8, ticks % 4294967296, true)
      dv.setUint32(at + 12, Math.floor(ticks / 4294967296), true)
    }
    dv.setUint32(at, ((id << 16) | type) >>> 0, true)
    dv.setUint32(at + 4, 6, true)
  })
  out[`${dir}__properties_version1.0`] = stream
  return out
}

export interface MsgRecipient {
  name: string
  smtp?: string
  /** 1 To, 2 Cc, 3 Bcc */
  type?: number
}

export interface MsgAttachment {
  name: string
  data?: Uint8Array
  mime?: string
  cid?: string
  hidden?: boolean
  /** an attached Outlook item */
  message?: MsgSpec
}

export interface MsgSpec {
  subject?: string
  messageClass?: string
  fromName?: string
  fromSmtp?: string
  /** an Exchange sender: an X.500 address and no SMTP one */
  fromExchange?: string
  sent?: Date
  importance?: number
  body?: string
  html?: string | Uint8Array
  rtf?: Uint8Array
  recipients?: MsgRecipient[]
  attachments?: MsgAttachment[]
  /** extra properties of the message */
  extra?: Record<number, Value>
}

/** the streams of an item in `dir`; `top`: the file's own item (a 32-byte property header, a name-to-id map) */
function msgStreams(spec: MsgSpec, dir: string, top: boolean): Record<string, Uint8Array> {
  const v: Record<number, Value | undefined> = {
    0x001a: { str: spec.messageClass ?? 'IPM.Note' },
    ...(spec.subject !== undefined ? { 0x0037: { str: spec.subject } } : {}),
    ...(spec.fromName ? { 0x0042: { str: spec.fromName }, 0x0c1a: { str: spec.fromName } } : {}),
    ...(spec.fromSmtp ? { 0x5d02: { str: spec.fromSmtp }, 0x5d01: { str: spec.fromSmtp } } : {}),
    ...(spec.fromExchange ? { 0x0064: { str: 'EX' }, 0x0065: { str: spec.fromExchange } } : {}),
    ...(spec.sent ? { 0x0039: { time: spec.sent } } : {}),
    ...(spec.importance !== undefined ? { 0x0017: { long: spec.importance } } : {}),
    ...(spec.body !== undefined ? { 0x1000: { str: spec.body } } : {}),
    ...(spec.html !== undefined ? { 0x1013: typeof spec.html === 'string' ? { bin: utf8(spec.html) } : { bin: spec.html } } : {}),
    ...(spec.rtf ? { 0x1009: { bin: spec.rtf } } : {}),
    ...(typeof spec.html === 'string' ? { 0x3fde: { long: 65001 } } : {}),
    ...spec.extra,
  }
  const recips = spec.recipients ?? []
  const atts = spec.attachments ?? []
  const out: Record<string, Uint8Array> = props(dir, v, top ? 32 : 24, (dv) => {
    dv.setUint32(8, recips.length, true)
    dv.setUint32(12, atts.length, true)
    dv.setUint32(16, recips.length, true)
    dv.setUint32(20, atts.length, true)
  })
  if (top) out['__nameid_version1.0/__substg1.0_00020102'] = new Uint8Array(0)
  recips.forEach((r, k) => {
    const rd = `${dir}__recip_version1.0_#${k.toString(16).toUpperCase().padStart(8, '0')}/`
    Object.assign(out, props(rd, { 0x0c15: { long: r.type ?? 1 }, 0x3001: { str: r.name }, ...(r.smtp ? { 0x39fe: { str: r.smtp }, 0x3002: { str: 'SMTP' }, 0x3003: { str: r.smtp } } : {}) }, 8))
  })
  atts.forEach((a, k) => {
    const ad = `${dir}__attach_version1.0_#${k.toString(16).toUpperCase().padStart(8, '0')}/`
    const av: Record<number, Value | undefined> = {
      0x3705: { long: a.message ? 5 : 1 },
      0x3707: { str: a.name },
      ...(a.mime ? { 0x370e: { str: a.mime } } : {}),
      ...(a.cid ? { 0x3712: { str: a.cid } } : {}),
      ...(a.hidden ? { 0x7ffe: { bool: true } } : {}),
      ...(a.data ? { 0x3701: { bin: a.data } } : {}),
    }
    Object.assign(out, props(ad, av, 8))
    if (a.message) Object.assign(out, msgStreams(a.message, `${ad}__substg1.0_3701000D/`, false))
  })
  return out
}

/** an Outlook item (.msg) */
export function msgFile(spec: MsgSpec): Uint8Array<ArrayBuffer> {
  return writeCfb(msgStreams(spec, '', true))
}

// ---------------------------------------------------------------- samples

/** the playground's sample .eml: an HTML newsletter-style message with a picture of its own, a table, an attachment and a picture from the web */
export function sampleEml(): Uint8Array<ArrayBuffer> {
  const logo = png(120, 40)
  const html = [
    '<html><head><style>p { margin: 0 }</style></head><body style="font-family: Arial">',
    '<p><img src="cid:logo@example.org" width="120" height="40" alt="Logo"></p>',
    '<h2>Quarterly review: agenda</h2>',
    '<p>Hi team,</p><p>here is the agenda for <b>Thursday</b>. The figures are in the attached sheet; the dashboard is at <a href="https://example.org/q3">example.org/q3</a>.</p>',
    '<table border="1" style="border-collapse: collapse"><tr><th bgcolor="#e8eef9">Time</th><th bgcolor="#e8eef9">Topic</th><th bgcolor="#e8eef9">Owner</th></tr>',
    '<tr><td>10:00</td><td>Revenue and margins</td><td>Ana</td></tr><tr><td>10:30</td><td>Hiring plan</td><td>Jonas</td></tr><tr><td>11:00</td><td>Roadmap</td><td>Mei</td></tr></table>',
    '<ul><li>Bring the updated forecast</li><li>Read the notes from last time</li></ul>',
    '<p>Thanks,<br>Ana</p><p><img src="https://tracker.example.com/open.gif" width="1" height="1"></p>',
    '</body></html>',
  ].join('\r\n')
  const sheet = utf8('Quarter,Revenue,Margin\r\nQ1,1200,0.31\r\nQ2,1350,0.33\r\nQ3,1420,0.35\r\n')
  return latin1(
    [
      'Return-Path: <ana@example.org>',
      'Received: from mail.example.org by mx.example.com; Thu, 19 Nov 2026 09:14:02 +0100',
      'From: =?UTF-8?Q?Ana_Mart=C3=ADnez?= <ana@example.org>',
      'To: Jonas Weber <jonas@example.com>, "Lee, Mei" <mei@example.com>',
      'Cc: team@example.com',
      'Subject: =?UTF-8?B?UXVhcnRlcmx5IHJldmlldyDigJMgYWdlbmRh?=',
      'Date: Thu, 19 Nov 2026 09:14:00 +0100',
      'Message-ID: <q3-review@example.org>',
      'MIME-Version: 1.0',
      'Content-Type: multipart/mixed; boundary="mixed"',
      '',
      '--mixed',
      'Content-Type: multipart/related; boundary="related"',
      '',
      '--related',
      'Content-Type: multipart/alternative; boundary="alt"',
      '',
      '--alt',
      'Content-Type: text/plain; charset=utf-8',
      'Content-Transfer-Encoding: quoted-printable',
      '',
      'Hi team, here is the agenda for Thursday.',
      '--alt',
      'Content-Type: text/html; charset=utf-8',
      'Content-Transfer-Encoding: base64',
      '',
      b64(utf8(html)),
      '--alt--',
      '--related',
      'Content-Type: image/png',
      'Content-Transfer-Encoding: base64',
      'Content-ID: <logo@example.org>',
      'Content-Disposition: inline; filename="logo.png"',
      '',
      b64(logo),
      '--related--',
      '--mixed',
      'Content-Type: text/csv; name="q3-figures.csv"',
      'Content-Transfer-Encoding: base64',
      'Content-Disposition: attachment; filename="q3-figures.csv"',
      '',
      b64(sheet),
      '--mixed--',
      '',
    ].join('\r\n'),
  )
}

/** the playground's sample .msg: an Outlook item with an HTML body, recipients, a hidden picture and an attached message */
export function sampleMsg(): Uint8Array<ArrayBuffer> {
  return msgFile({
    subject: 'Contract renewal – signed copy',
    fromName: 'Jonas Weber',
    fromSmtp: 'jonas@example.com',
    sent: new Date(Date.UTC(2026, 10, 20, 15, 42)),
    importance: 2,
    recipients: [
      { name: 'Ana Martínez', smtp: 'ana@example.org' },
      { name: 'Legal', smtp: 'legal@example.org', type: 2 },
    ],
    body: 'Hello Ana,\r\n\r\nattached is the signed contract. Please file it.\r\n\r\nJonas',
    html: '<html><body><p>Hello Ana,</p><p>attached is the <b>signed contract</b>. Please file it.</p><p><img src="cid:sig"></p><p>Jonas</p></body></html>',
    attachments: [
      { name: 'contract-signed.pdf', mime: 'application/pdf', data: latin1('%PDF-1.4\n% sample\n') },
      { name: 'signature.png', mime: 'image/png', cid: 'sig', hidden: true, data: png(160, 24, [0x1f, 0x29, 0x37]) },
      {
        name: 'Original request.msg',
        message: {
          subject: 'Contract renewal',
          fromName: 'Ana Martínez',
          fromSmtp: 'ana@example.org',
          sent: new Date(Date.UTC(2026, 10, 18, 8, 5)),
          recipients: [{ name: 'Jonas Weber', smtp: 'jonas@example.com' }],
          body: 'Jonas,\r\n\r\ncould you send the signed renewal by Friday?\r\n\r\nThanks, Ana',
        },
      },
    ],
  })
}
