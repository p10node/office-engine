/**
 * Outlook items (.msg; .oft templates): a compound file of MAPI properties (MS-OXMSG). Fixed-size values are in the
 * `__properties_version1.0` stream, the others in `__substg1.0_<id><type>` streams; recipients, attachments and
 * attached items are storages of their own.
 */
import { Cfb } from '../legacy/cfb'
import { decode } from '../legacy/rtf/codepage'
import { parseAddresses, parseDate, parseEntity, binary } from './mime'
import { decompressRtf, unwrapRtf } from './rtf'
import type { MailAddress, MailAttachment, MailBody, MailMessage } from './types'

const PT_SHORT = 0x0002
const PT_LONG = 0x0003
const PT_BOOLEAN = 0x000b
const PT_STRING8 = 0x001e
const PT_UNICODE = 0x001f
const PT_SYSTIME = 0x0040
const PT_BINARY = 0x0102

// property ids (MS-OXPROPS)
const SUBJECT = 0x0037
const MESSAGE_CLASS = 0x001a
const IMPORTANCE = 0x0017
const TRANSPORT_HEADERS = 0x007d
const SENT_REPRESENTING_NAME = 0x0042
const SENT_REPRESENTING_ADDRTYPE = 0x0064
const SENT_REPRESENTING_EMAIL = 0x0065
const SENT_REPRESENTING_SMTP = 0x5d02
const SENDER_NAME = 0x0c1a
const SENDER_ADDRTYPE = 0x0c1e
const SENDER_EMAIL = 0x0c1f
const SENDER_SMTP = 0x5d01
const DISPLAY_BCC = 0x0e02
const DISPLAY_CC = 0x0e03
const DISPLAY_TO = 0x0e04
const CLIENT_SUBMIT_TIME = 0x0039
const DELIVERY_TIME = 0x0e06
const CREATION_TIME = 0x3007
const LAST_MODIFICATION_TIME = 0x3008
const BODY = 0x1000
const RTF_COMPRESSED = 0x1009
const HTML = 0x1013
const INTERNET_CPID = 0x3fde
const MESSAGE_CODEPAGE = 0x3ffd
const RECIPIENT_TYPE = 0x0c15
const DISPLAY_NAME = 0x3001
const ADDRTYPE = 0x3002
const EMAIL_ADDRESS = 0x3003
const SMTP_ADDRESS = 0x39fe
const ATTACH_DATA = 0x3701
const ATTACH_FILENAME = 0x3704
const ATTACH_METHOD = 0x3705
const ATTACH_LONG_FILENAME = 0x3707
const ATTACH_MIME_TAG = 0x370e
const ATTACH_CONTENT_ID = 0x3712
const ATTACH_CONTENT_LOCATION = 0x3713
const ATTACHMENT_HIDDEN = 0x7ffe

const ATTACH_EMBEDDED_MSG = 5

/** attached items inside one another, at most */
const MAX_DEPTH = 16

const hex4 = (n: number) => n.toString(16).toUpperCase().padStart(4, '0')

/** the MAPI properties of one object (the message, a recipient, an attachment): a storage of the compound file */
class Props {
  private fixed = new Map<number, { type: number; at: number }>()
  private dv: DataView
  private codepage = 1252

  /** `dir`: the storage's path with a trailing "/" ('' for the root); `skip`: the size of the property stream's header */
  constructor(private cfb: Cfb, private dir: string, skip: number) {
    const s = cfb.stream(dir + '__properties_version1.0') ?? new Uint8Array(0)
    this.dv = new DataView(s.buffer, s.byteOffset, s.byteLength)
    for (let p = skip; p + 16 <= s.length; p += 16) {
      const tag = this.dv.getUint32(p, true)
      this.fixed.set(tag >>> 16, { type: tag & 0xffff, at: p + 8 })
    }
    this.codepage = this.long(INTERNET_CPID) ?? this.long(MESSAGE_CODEPAGE) ?? 1252
  }

  /** text in the code page of the message (also of an attachment or recipient, which have none of their own) */
  setCodepage(cp: number) {
    if (!this.fixed.has(INTERNET_CPID) && !this.fixed.has(MESSAGE_CODEPAGE)) this.codepage = cp
  }

  get cp(): number {
    return this.codepage
  }

  long(id: number): number | undefined {
    const f = this.fixed.get(id)
    if (!f) return undefined
    if (f.type === PT_LONG) return this.dv.getInt32(f.at, true)
    if (f.type === PT_SHORT) return this.dv.getInt16(f.at, true)
    return undefined
  }

  bool(id: number): boolean {
    const f = this.fixed.get(id)
    return !!f && f.type === PT_BOOLEAN && this.dv.getUint8(f.at) !== 0
  }

  time(id: number): Date | undefined {
    const f = this.fixed.get(id)
    if (!f || f.type !== PT_SYSTIME) return undefined
    const ticks = this.dv.getUint32(f.at + 4, true) * 4294967296 + this.dv.getUint32(f.at, true)
    if (!ticks) return undefined
    // 100 ns intervals since 1601-01-01
    const d = new Date(ticks / 10000 - 11644473600000)
    return Number.isNaN(d.getTime()) ? undefined : d
  }

  bytes(id: number): Uint8Array | undefined {
    return this.cfb.stream(`${this.dir}__substg1.0_${hex4(id)}${hex4(PT_BINARY)}`)
  }

  str(id: number): string {
    const uni = this.cfb.stream(`${this.dir}__substg1.0_${hex4(id)}${hex4(PT_UNICODE)}`)
    if (uni) return new TextDecoder('utf-16le').decode(uni).replace(/\0+$/, '')
    const ansi = this.cfb.stream(`${this.dir}__substg1.0_${hex4(id)}${hex4(PT_STRING8)}`)
    if (ansi) return decode(ansi, this.codepage).replace(/\0+$/, '')
    return ''
  }
}

/** an Exchange address ("/O=EXCHANGELABS/OU=…/CN=RECIPIENTS/CN=…"): not one to show */
const exchangeDn = (addr: string) => /^\/o=/i.test(addr)

function address(name: string, addr: string): MailAddress {
  const a = exchangeDn(addr) ? '' : addr.trim()
  const n = name.trim()
  return { name: n === a ? '' : n, addr: a }
}

/** Outlook's display strings ("Ann Lee; Bob"): names only */
const displayNames = (s: string): MailAddress[] =>
  s
    .split(';')
    .map((x) => x.trim())
    .filter(Boolean)
    .map((name) => ({ name, addr: '' }))

class MsgReader {
  constructor(private cfb: Cfb) {}

  /** the storages of recipients or attachments in `dir`, in their order (the hex number their names end with) */
  private storages(dir: string, prefix: RegExp) {
    return this.cfb
      .children(dir.replace(/\/$/, ''))
      .filter((e) => e.type === 'storage' && prefix.test(e.name))
      .sort((a, b) => parseInt(a.name.slice(a.name.indexOf('#') + 1), 16) - parseInt(b.name.slice(b.name.indexOf('#') + 1), 16))
  }

  /** the item in storage `dir` (its property stream's header: 32 bytes for the file's item, 24 for an attached one) */
  read(dir: string, depth: number): MailMessage {
    const p = new Props(this.cfb, dir, depth ? 24 : 32)
    const headers = p.str(TRANSPORT_HEADERS)
    const mime = headers ? parseEntity(headers.replace(/\r?\n?$/, '\r\n\r\n')) : null
    const mimeHeader = (name: string) => mime?.headers.find((h) => h[0] === name)?.[1] ?? ''

    const recipients = { to: [] as MailAddress[], cc: [] as MailAddress[], bcc: [] as MailAddress[] }
    for (const e of this.storages(dir, /^__recip_version1\.0_#/i)) {
      const r = new Props(this.cfb, e.path + '/', 8)
      r.setCodepage(p.cp)
      const type = r.long(RECIPIENT_TYPE) ?? 1
      const smtp = r.str(SMTP_ADDRESS)
      const email = r.str(EMAIL_ADDRESS)
      const a = address(r.str(DISPLAY_NAME), smtp || (r.str(ADDRTYPE).toUpperCase() !== 'EX' ? email : ''))
      if (!a.name && !a.addr) continue
      ;(type === 3 ? recipients.bcc : type === 2 ? recipients.cc : recipients.to).push(a)
    }
    // no recipient storages (some writers): the transport headers, else Outlook's display strings
    const fallback = (list: MailAddress[], name: string, display: number) =>
      list.length ? list : mimeHeader(name) ? parseAddresses(mimeHeader(name)) : displayNames(p.str(display))

    const smtpOf = (smtp: number, type: number, email: number) => p.str(smtp) || (p.str(type).toUpperCase() === 'SMTP' ? p.str(email) : '') || (exchangeDn(p.str(email)) ? '' : p.str(email))
    let from = [address(p.str(SENT_REPRESENTING_NAME), smtpOf(SENT_REPRESENTING_SMTP, SENT_REPRESENTING_ADDRTYPE, SENT_REPRESENTING_EMAIL))]
    if (!from[0].name && !from[0].addr) from = [address(p.str(SENDER_NAME), smtpOf(SENDER_SMTP, SENDER_ADDRTYPE, SENDER_EMAIL))]
    if (!from[0].addr && mimeHeader('from')) {
      const h = parseAddresses(mimeHeader('from'))
      if (h.length) from = [{ name: from[0].name || h[0].name, addr: h[0].addr }]
    }
    if (!from[0].name && !from[0].addr) from = []

    const date = p.time(CLIENT_SUBMIT_TIME) ?? p.time(DELIVERY_TIME) ?? parseDate(mimeHeader('date')) ?? p.time(CREATION_TIME) ?? p.time(LAST_MODIFICATION_TIME)
    const importance = p.long(IMPORTANCE)
    const itemClass = p.str(MESSAGE_CLASS)
    const msg: MailMessage = {
      subject: p.str(SUBJECT),
      from,
      to: fallback(recipients.to, 'to', DISPLAY_TO),
      cc: fallback(recipients.cc, 'cc', DISPLAY_CC),
      bcc: fallback(recipients.bcc, 'bcc', DISPLAY_BCC),
      ...(date ? { date } : {}),
      body: this.body(p),
      attachments: [],
      ...(importance === 2 ? { importance: 'high' as const } : importance === 0 ? { importance: 'low' as const } : {}),
      ...(itemClass && !/^IPM\.Note$/i.test(itemClass) ? { itemClass } : {}),
    }
    if (/^IPM\.Note\.(?:SMIME|Secure)(?!\.MultipartSigned)/i.test(itemClass)) msg.encrypted = true
    for (const e of this.storages(dir, /^__attach_version1\.0_#/i)) {
      const a = this.attachment(e.path + '/', p.cp, depth)
      if (a) msg.attachments.push(a)
    }
    return msg
  }

  /** the body: HTML (PR_HTML, else the HTML the RTF body wraps), else the RTF itself, else the plain text */
  private body(p: Props): MailBody[] {
    const html = p.bytes(HTML)
    if (html?.length) return [{ html: htmlText(html, p.cp) }]
    const htmlStr = p.str(HTML)
    if (htmlStr) return [{ html: htmlStr }]
    const text = p.str(BODY)
    const compressed = p.bytes(RTF_COMPRESSED)
    if (compressed?.length) {
      try {
        const rtf = unwrapRtf(decompressRtf(compressed))
        // (a plain-text body wrapped in RTF: PR_BODY is the same text, without the RTF's guesses)
        if ('html' in rtf) return [rtf]
        if ('rtf' in rtf) return [text ? { rtf: rtf.rtf, text } : rtf]
        return [{ text: text || rtf.text }]
      } catch {
        // damaged: the plain text
      }
    }
    return text ? [{ text }] : []
  }

  private attachment(dir: string, cp: number, depth: number): MailAttachment | null {
    const a = new Props(this.cfb, dir, 8)
    a.setCodepage(cp)
    const method = a.long(ATTACH_METHOD)
    const name = (a.str(ATTACH_LONG_FILENAME) || a.str(ATTACH_FILENAME) || a.str(DISPLAY_NAME)).replace(/[\\/]+/g, '_')
    const cid = a.str(ATTACH_CONTENT_ID).replace(/^<|>$/g, '')
    const location = a.str(ATTACH_CONTENT_LOCATION)
    const type = a.str(ATTACH_MIME_TAG).toLowerCase()
    const common = { ...(cid ? { cid } : {}), ...(location ? { location } : {}), ...(a.bool(ATTACHMENT_HIDDEN) ? { hidden: true } : {}) }
    const inner = `${dir}__substg1.0_${hex4(ATTACH_DATA)}000D`
    if (this.cfb.has(inner)) {
      if (method === ATTACH_EMBEDDED_MSG && depth < MAX_DEPTH) {
        const message = this.read(inner + '/', depth + 1)
        return { name: name || `${message.subject || 'message'}.msg`, type: 'application/vnd.ms-outlook', message, ...common }
      }
      // an OLE object (a document Outlook shows as an object): not read
      return null
    }
    const data = a.bytes(ATTACH_DATA)
    if (!data) return null
    return { name: name || `attachment.${type.split('/')[1] ?? 'bin'}`, type, data, ...common }
  }
}

/** the text of an HTML body stored as bytes: in its <meta> charset, else the item's code page */
function htmlText(b: Uint8Array, cp: number): string {
  const meta = /<meta[^>]+charset\s*=\s*["']?([\w-]+)/i.exec(binary(b.subarray(0, 2048)))?.[1]
  if (meta) {
    try {
      return new TextDecoder(meta).decode(b)
    } catch {
      // unknown label: the code page
    }
  }
  return decode(b, cp)
}

/** the compound file is an Outlook item: a message's property stream and a message class or subject */
export function isMsgCfb(cfb: Cfb): boolean {
  return cfb.has('__properties_version1.0') && (cfb.has('__substg1.0_001A001F') || cfb.has('__substg1.0_001A001E') || cfb.has('__nameid_version1.0'))
}

/** Reads an Outlook item (the bytes of a .msg / .oft file). */
export function readMsg(bytes: Uint8Array): MailMessage {
  const cfb = Cfb.open(bytes)
  if (!isMsgCfb(cfb)) throw new Error('Not an Outlook item (.msg)')
  return new MsgReader(cfb).read('', 0)
}
