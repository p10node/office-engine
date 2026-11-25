/**
 * Internet messages (.eml; Apple Mail's .emlx): RFC 5322 headers, MIME parts (RFC 2045-2049), encoded words
 * (RFC 2047) and parameters (RFC 2231), read into the message shown. Structure is read on a "binary string" (one
 * character per byte); text is decoded in its part's charset.
 */
import { decode } from '../legacy/rtf/codepage'
import type { MailAddress, MailAttachment, MailBody, MailMessage } from './types'

/** parts inside one another, at most (a message attached to a message attached to…) */
const MAX_DEPTH = 32

/** bytes as a string of one character per byte */
export function binary(b: Uint8Array): string {
  let s = ''
  for (let k = 0; k < b.length; k += 0x8000) s += String.fromCharCode(...b.subarray(k, k + 0x8000))
  return s
}

/** a string of one character per byte as its bytes */
export function bytesOf(s: string): Uint8Array {
  const b = new Uint8Array(s.length)
  for (let k = 0; k < s.length; k++) b[k] = s.charCodeAt(k)
  return b
}

/** text in a charset label ("utf-8", "iso-8859-1", "ks_c_5601-1987"…); UTF-8, else Windows-1252, when the label is unknown */
export function decodeCharset(b: Uint8Array, charset: string | undefined): string {
  const label = (charset ?? '').trim().replace(/^["']|["']$/g, '').replace(/\*.*$/, '')
  if (label) {
    // (TextDecoder's "latin1" / "us-ascii" are Windows-1252, as mail clients read them)
    try {
      return new TextDecoder(label).decode(b)
    } catch {
      const cp = /^(?:cp|windows-?|ms)(\d{3,5})$/i.exec(label)
      if (cp) return decode(b, Number(cp[1]))
    }
  }
  try {
    return new TextDecoder('utf-8', { fatal: true }).decode(b)
  } catch {
    return new TextDecoder('windows-1252').decode(b)
  }
}

export function base64(s: string): Uint8Array {
  let t = s.replace(/[^A-Za-z0-9+/]/g, '')
  if (t.length % 4 === 1) t = t.slice(0, -1)
  t += '='.repeat((4 - (t.length % 4)) % 4)
  try {
    return bytesOf(atob(t))
  } catch {
    return new Uint8Array(0)
  }
}

export function quotedPrintable(s: string): string {
  return s.replace(/=(?:\r?\n|([0-9A-Fa-f]{2}))/g, (_, h: string | undefined) => (h ? String.fromCharCode(parseInt(h, 16)) : ''))
}

// ---------------------------------------------------------------- headers

export interface Entity {
  /** header names lower-cased → raw values (unfolded, one character per byte), in order */
  headers: [string, string][]
  /** the body, one character per byte */
  body: string
}

/** Splits an entity (a message or a part) into its headers and body. */
export function parseEntity(raw: string): Entity {
  // the blank line: CRLF CRLF, or LF LF in files saved on Unix
  const m = /\r?\n\r?\n/.exec(raw)
  // a part with no header ends its first line at the blank line
  const blank = /^\r?\n/.exec(raw)
  const cut = blank ? 0 : m ? m.index : raw.length
  const after = blank ? blank[0].length : m ? m.index + m[0].length : raw.length
  const headers: [string, string][] = []
  for (const line of raw.slice(0, cut).split(/\r?\n(?![ \t])/)) {
    const colon = line.indexOf(':')
    if (colon <= 0) continue
    const name = line.slice(0, colon).trim().toLowerCase()
    if (!/^[!-9;-~]+$/.test(name)) continue
    headers.push([name, line.slice(colon + 1).replace(/\r?\n[ \t]/g, ' ').trim()])
  }
  return { headers, body: raw.slice(after) }
}

const header = (e: Entity, name: string) => e.headers.find((h) => h[0] === name)?.[1]

/** 8-bit header text: UTF-8 (RFC 6532), else Windows-1252 */
function eightBit(raw: string): string {
  // eslint-disable-next-line no-control-regex
  return /[^\x00-\x7f]/.test(raw) ? decodeCharset(bytesOf(raw), undefined) : raw
}

const WORD = /=\?([^?\s]+)\?([BbQq])\?([^?\s]*)\?=/g

/** RFC 2047 encoded words decoded; the space between two of them is dropped, and words of one charset are decoded together (a character may be split between them) */
export function decodeWords(raw: string): string {
  const text = eightBit(raw)
  if (!text.includes('=?')) return text
  let out = ''
  let at = 0
  let pending: { charset: string; bytes: string } | null = null
  const flush = () => {
    if (pending) out += decodeCharset(bytesOf(pending.bytes), pending.charset)
    pending = null
  }
  for (const m of text.matchAll(WORD)) {
    const gap = text.slice(at, m.index)
    if (gap && (!pending || /\S/.test(gap))) {
      flush()
      out += gap
    }
    const charset = m[1].toLowerCase()
    const bytes = m[2].toUpperCase() === 'B' ? binary(base64(m[3])) : quotedPrintable(m[3].replace(/_/g, ' '))
    if (pending && pending.charset === charset) pending.bytes += bytes
    else {
      flush()
      pending = { charset, bytes }
    }
    at = m.index! + m[0].length
  }
  flush()
  return out + text.slice(at)
}

/** a header's value decoded (encoded words), '' when there is none */
export const headerText = (e: Entity, name: string) => decodeWords(header(e, name) ?? '')

/** splits on `sep` outside quotes, comments and angle brackets */
function splitOutside(s: string, sep: string): string[] {
  const out: string[] = []
  let cur = ''
  let quoted = false
  let depth = 0
  let angle = false
  for (let k = 0; k < s.length; k++) {
    const c = s[k]
    if (quoted) {
      if (c === '\\' && k + 1 < s.length) cur += c + s[++k]
      else {
        if (c === '"') quoted = false
        cur += c
      }
      continue
    }
    if (c === '"') quoted = true
    else if (c === '(') depth++
    else if (c === ')' && depth) depth--
    else if (c === '<') angle = true
    else if (c === '>') angle = false
    else if (c === sep && !depth && !angle) {
      out.push(cur)
      cur = ''
      continue
    }
    cur += c
  }
  out.push(cur)
  return out
}

const unquote = (s: string) => (/^".*"$/s.test(s) ? s.slice(1, -1).replace(/\\(.)/g, '$1') : s)

/** An address list ("Ann <ann@x.org>, "Doe, Bob" <bob@y.com>, team: a@z.org, b@z.org;"): its names decoded. */
export function parseAddresses(raw: string): MailAddress[] {
  const out: MailAddress[] = []
  for (let item of splitOutside(eightBit(raw), ',')) {
    // a group: "Team: a@x, b@y;" (its name is left out; "undisclosed-recipients:;" has no address)
    const group = /^[^"<(]*?:/.exec(item)
    if (group && !/^\s*[^\s@]+@/.test(item)) item = item.slice(group[0].length)
    item = item.replace(/;\s*$/, '').trim()
    if (!item) continue
    const angle = /<([^>]*)>/.exec(item)
    let name: string
    let addr: string
    if (angle) {
      addr = angle[1].trim()
      name = (item.slice(0, angle.index) + item.slice(angle.index + angle[0].length)).trim()
    } else {
      // "ann@x.org (Ann)": the comment is the name
      const comment = /\(([^)]*)\)/.exec(item)
      addr = item.replace(/\([^)]*\)/g, '').trim()
      name = comment ? comment[1].trim() : ''
    }
    name = decodeWords(unquote(name.replace(/\s+/g, ' ').trim())).trim()
    if (/^mailto:/i.test(addr)) addr = addr.slice(7)
    if (name === addr) name = ''
    if (name || addr) out.push({ name, addr })
  }
  return out
}

/** A structured header ("text/html; charset=utf-8", "attachment; filename*=UTF-8''r%C3%A9sum%C3%A9.pdf"): its value lower-cased and its parameters, RFC 2231 joined and decoded. */
export function parseParams(raw: string | undefined): { value: string; params: Record<string, string> } {
  const [first, ...rest] = splitOutside(raw ?? '', ';')
  const params: Record<string, string> = Object.create(null)
  const pieces: Record<string, { k: number; v: string; enc: boolean }[]> = Object.create(null)
  for (const p of rest) {
    const eq = p.indexOf('=')
    if (eq < 0) continue
    const key = p.slice(0, eq).trim().toLowerCase()
    const v = unquote(p.slice(eq + 1).trim())
    const m = /^([^*]+)(?:\*(\d+))?(\*)?$/.exec(key)
    if (!m) continue
    if (m[2] === undefined && !m[3]) {
      // (non-standard but common: an encoded word in a quoted file name)
      if (!(m[1] in params)) params[m[1]] = decodeWords(v)
      continue
    }
    ;(pieces[m[1]] ??= []).push({ k: Number(m[2] ?? 0), v, enc: !!m[3] })
  }
  for (const [name, list] of Object.entries(pieces)) {
    list.sort((a, b) => a.k - b.k)
    let charset = ''
    let bytes = ''
    list.forEach((p, i) => {
      let v = p.v
      if (p.enc && i === 0) {
        const m = /^([^']*)'[^']*'(.*)$/s.exec(v)
        if (m) {
          charset = m[1]
          v = m[2]
        }
      }
      bytes += p.enc ? v.replace(/%([0-9A-Fa-f]{2})/g, (_, h) => String.fromCharCode(parseInt(h, 16))) : v
    })
    // the RFC 2231 form wins over the plain one
    params[name] = decodeCharset(bytesOf(bytes), charset || 'utf-8')
  }
  return { value: first.trim().toLowerCase(), params }
}

/** a Date header as a date (RFC 5322 and the obsolete forms browsers read), or undefined */
export function parseDate(raw: string): Date | undefined {
  if (!raw) return undefined
  // "Tue, 3 Nov 2026 09:15:00 +0100 (CET)": the comment and the day name are not needed
  const t = raw.replace(/\([^)]*\)/g, '').replace(/^\s*[A-Za-z]{3},?\s*/, '').trim()
  const d = new Date(t)
  return Number.isNaN(d.getTime()) ? undefined : d
}

// ---------------------------------------------------------------- parts

/** the parts of a multipart body: between the boundary's delimiter lines, the preamble and epilogue left out */
function splitMultipart(body: string, boundary: string): string[] {
  const parts: string[] = []
  const delim = '--' + boundary
  let at = -1
  let from = 0
  for (;;) {
    let k = body.indexOf(delim, from)
    // a delimiter starts a line
    while (k > 0 && body[k - 1] !== '\n') k = body.indexOf(delim, k + 1)
    if (k < 0) {
      // a message cut short: what came after the last delimiter is its last part
      if (at >= 0) parts.push(body.slice(at))
      break
    }
    if (at >= 0) parts.push(body.slice(at, k).replace(/\r?\n$/, ''))
    const close = body.startsWith('--', k + delim.length)
    const eol = body.indexOf('\n', k)
    if (close) break
    if (eol < 0) break
    at = eol + 1
    from = at
  }
  return parts
}

/** a part's body bytes, its transfer encoding undone */
function partBytes(e: Entity): Uint8Array {
  const how = (header(e, 'content-transfer-encoding') ?? '').trim().toLowerCase()
  if (how === 'base64') return base64(e.body)
  if (how === 'quoted-printable') return bytesOf(quotedPrintable(e.body))
  return bytesOf(e.body)
}

const SIGNATURE = /^application\/(?:x-)?(?:pkcs7-signature|pgp-signature)$/
const ENCRYPTED = /^application\/(?:x-)?pkcs7-mime$|^multipart\/encrypted$/
/** text parts shown as text (a delivery report's details) */
const SHOWN_TEXT = /^(?:text\/plain|text\/rfc822-headers|message\/delivery-status|message\/disposition-notification|message\/feedback-report)$/

/** RFC 3676 format=flowed: lines ending in a space go on in the next (quoted lines with as many '>'), space-stuffing undone */
export function unflow(text: string, delsp: boolean): string {
  const out: string[] = []
  let soft = false
  let quote = -1
  for (const raw of text.split(/\r?\n/)) {
    const q = /^>*/.exec(raw)![0].length
    let line = raw.slice(q)
    if (line.startsWith(' ')) line = line.slice(1)
    const flowed = line.endsWith(' ') && line !== '-- '
    if (flowed && delsp) line = line.slice(0, -1)
    if (soft && q === quote) out[out.length - 1] += line
    else out.push('>'.repeat(q) + (q && line ? ' ' : '') + line)
    soft = flowed
    quote = q
  }
  return out.join('\n')
}

function textOf(e: Entity, ct: { value: string; params: Record<string, string> }): string {
  let text = decodeCharset(partBytes(e), ct.params.charset)
  if (ct.value === 'text/plain' && ct.params.format?.toLowerCase() === 'flowed') text = unflow(text, ct.params.delsp?.toLowerCase() === 'yes')
  return text
}

class Reader {
  readonly msg: MailMessage

  constructor(top: Entity, private depth: number) {
    const date = header(top, 'date') ?? ''
    const parsed = parseDate(date)
    const importance = (header(top, 'importance') ?? header(top, 'x-priority') ?? '').trim().toLowerCase()
    this.msg = {
      subject: headerText(top, 'subject'),
      from: parseAddresses(header(top, 'from') ?? header(top, 'sender') ?? ''),
      to: parseAddresses(top.headers.filter((h) => h[0] === 'to').map((h) => h[1]).join(', ')),
      cc: parseAddresses(top.headers.filter((h) => h[0] === 'cc').map((h) => h[1]).join(', ')),
      bcc: parseAddresses(top.headers.filter((h) => h[0] === 'bcc').map((h) => h[1]).join(', ')),
      ...(parsed ? { date: parsed } : date ? { dateText: decodeWords(date) } : {}),
      ...(/^(?:high|urgent|[12]\b)/.test(importance) ? { importance: 'high' as const } : /^(?:low|non-urgent|[45]\b)/.test(importance) ? { importance: 'low' as const } : {}),
      body: [],
      attachments: [],
    }
    this.entity(top, 0)
  }

  private get body(): MailBody[] {
    return this.msg.body
  }

  /** a part: shown (body), or attached */
  private entity(e: Entity, depth: number): void {
    const ct = parseParams(header(e, 'content-type') ?? 'text/plain')
    const type = ct.value || 'text/plain'
    if (depth > MAX_DEPTH) return
    if (ENCRYPTED.test(type)) this.msg.encrypted = true
    if (type.startsWith('multipart/')) return this.multipart(e, type, ct.params, depth)
    const disp = parseParams(header(e, 'content-disposition'))
    const name = disp.params.filename ?? ct.params.name ?? ''
    const attached = disp.value === 'attachment' || (!!name && disp.value !== 'inline')
    if (!attached && (type === 'text/html' || type === 'text/enriched' || SHOWN_TEXT.test(type) || (type === 'text/plain' && !name))) {
      if (type === 'text/html') this.body.push({ html: textOf(e, ct) })
      else this.body.push({ text: textOf(e, ct) })
      return
    }
    if (SIGNATURE.test(type)) return
    this.attach(e, type, name, depth)
  }

  private attach(e: Entity, type: string, name: string, depth: number) {
    const data = partBytes(e)
    const cid = header(e, 'content-id')?.trim().replace(/^<|>$/g, '')
    const location = header(e, 'content-location')?.trim()
    const a: MailAttachment = { name: name.replace(/[\\/]+/g, '_'), type, data, ...(cid ? { cid } : {}), ...(location ? { location } : {}) }
    if (type === 'message/rfc822' || type === 'message/global') {
      const inner = parseEntity(type === 'message/global' ? decodeCharset(data, 'utf-8') : binary(data))
      if (depth < MAX_DEPTH && this.depth < MAX_DEPTH) a.message = new Reader(inner, this.depth + 1).msg
      if (!a.name) a.name = (a.message?.subject || 'message').replace(/[\\/:*?"<>|]+/g, '_') + '.eml'
    }
    if (!a.name) a.name = defaultName(type, this.msg.attachments.length + 1)
    this.msg.attachments.push(a)
  }

  private multipart(e: Entity, type: string, params: Record<string, string>, depth: number): void {
    const parts = params.boundary ? splitMultipart(e.body, params.boundary).map(parseEntity) : []
    if (!parts.length) return
    if (type === 'multipart/alternative') {
      // the richest form that can be shown: the last one (RFC 2046) among HTML, related HTML, plain text
      const score = (p: Entity) => {
        const t = parseParams(header(p, 'content-type') ?? 'text/plain').value || 'text/plain'
        return t === 'text/html' ? 3 : t === 'multipart/related' || t === 'multipart/mixed' ? 2 : t === 'text/plain' ? 1 : 0
      }
      let best = parts[0]
      for (const p of parts) if (score(p) >= score(best)) best = p
      return this.entity(best, depth + 1)
    }
    if (type === 'multipart/related') {
      // the root (named by `start`, else the first part) is shown, the others are what it shows
      const start = params.start?.replace(/^<|>$/g, '')
      const root = (start && parts.find((p) => header(p, 'content-id')?.trim().replace(/^<|>$/g, '') === start)) || parts[0]
      this.entity(root, depth + 1)
      for (const p of parts) if (p !== root) this.related(p, depth + 1)
      return
    }
    if (type === 'multipart/signed') return this.entity(parts[0], depth + 1)
    // mixed, digest, report…: the parts in order; the first shown one is the body, more text goes on after it
    for (const p of parts) this.entity(p, depth + 1)
  }

  /** a resource of an HTML body (a picture): attached, the body shows it by its Content-ID or location */
  private related(e: Entity, depth: number): void {
    const ct = parseParams(header(e, 'content-type') ?? 'application/octet-stream')
    if (ct.value.startsWith('multipart/')) return this.multipart(e, ct.value, ct.params, depth)
    const disp = parseParams(header(e, 'content-disposition'))
    this.attach(e, ct.value || 'application/octet-stream', disp.params.filename ?? ct.params.name ?? '', depth)
  }
}

const EXT: Record<string, string> = { 'image/png': 'png', 'image/jpeg': 'jpg', 'image/gif': 'gif', 'image/bmp': 'bmp', 'image/webp': 'webp', 'text/calendar': 'ics', 'text/vcard': 'vcf', 'text/x-vcard': 'vcf', 'application/pdf': 'pdf' }

/** a file name for an attachment that has none */
function defaultName(type: string, n: number): string {
  return `attachment-${n}.${EXT[type] ?? (type.startsWith('text/') ? 'txt' : 'bin')}`
}

/** Reads an Internet message (the bytes of a .eml file). */
export function readEml(bytes: Uint8Array): MailMessage {
  return new Reader(parseEntity(binary(bytes)), 0).msg
}

/** An Apple Mail .emlx file: its first line is the message's length in bytes, an XML property list follows the message. */
export function readEmlx(bytes: Uint8Array): MailMessage {
  const head = binary(bytes.subarray(0, 32))
  const m = /^\s*(\d+)[ \t]*\r?\n/.exec(head)
  if (!m) return readEml(bytes)
  const start = m[0].length
  return readEml(bytes.subarray(start, Math.min(bytes.length, start + Number(m[1]))))
}

/** the start of an Internet message: header lines, among them those every message has */
export function looksLikeEml(head: Uint8Array): boolean {
  const text = binary(head.subarray(0, 4096))
  const lines = /^(?:[!-9;-~]+:[^\r\n]*\r?\n(?:[ \t][^\r\n]*\r?\n)*)+/.exec(text.replace(/^(?:From [^\r\n]*\r?\n)/, ''))
  if (!lines) return false
  const names = new Set([...lines[0].matchAll(/^([!-9;-~]+):/gm)].map((x) => x[1].toLowerCase()))
  // a web page saved by a browser (MHTML) has From / Subject / Date too, but neither recipients nor relays
  if (names.has('snapshot-content-location') || /^From: <Saved by /im.test(lines[0])) return false
  return (names.has('from') || names.has('received') || names.has('return-path')) && (names.has('to') || names.has('received') || names.has('message-id') || names.has('subject'))
}
