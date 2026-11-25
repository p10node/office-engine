/**
 * A mail message as a document: its subject, a table of who sent it to whom and when, its attachments, then its
 * body (HTML with its tables and the pictures it carries, plain text, or Outlook's RTF), then the messages attached
 * to it. Pictures the HTML loads from the web are never fetched.
 */
import { parasToDocx, parasXml, type ImportCell, type ImportPara, type ImportPicture, type ImportRun } from '../convert'
import { htmlToParas, type PastedPara } from '../html'
import { rtfToDocx } from '../legacy/rtf'
import { base64 } from './mime'
import { pictureExt, pictureSize } from './picture'
import type { MailAddress, MailAttachment, MailBody, MailMessage, MailTexts } from './types'

const GREY = '#595959'
const LABEL_SHARE = 0.16

/** what a message's document is built with */
interface Ctx {
  texts: MailTexts
  locale?: string
  /** attachments a body shows (its pictures): not listed */
  shown: Set<MailAttachment>
}

/** the label of an Outlook item that is not a plain message */
function itemLabel(itemClass: string | undefined, t: MailTexts): string | undefined {
  if (!itemClass) return undefined
  const c = itemClass.toLowerCase()
  const i = t.item
  if (c.startsWith('ipm.schedule.meeting.request')) return i.meetingRequest
  if (c.startsWith('ipm.schedule.meeting.canceled')) return i.meetingCanceled
  if (c.startsWith('ipm.schedule.meeting.resp.pos')) return i.meetingAccepted
  if (c.startsWith('ipm.schedule.meeting.resp.neg')) return i.meetingDeclined
  if (c.startsWith('ipm.schedule.meeting.resp.tent')) return i.meetingTentative
  if (c.startsWith('ipm.appointment')) return i.appointment
  if (c.startsWith('ipm.contact')) return i.contact
  if (c.startsWith('ipm.task')) return i.task
  if (c.startsWith('ipm.stickynote')) return i.note
  if (c.startsWith('report.')) return i.report
  return undefined
}

function formatDate(d: Date, locale: string | undefined): string {
  try {
    return new Intl.DateTimeFormat(locale, { dateStyle: 'full', timeStyle: 'short' }).format(d)
  } catch {
    return d.toUTCString()
  }
}

export function formatSize(n: number, locale?: string): string {
  const f = (v: number, digits: number) => new Intl.NumberFormat(locale, { maximumFractionDigits: digits }).format(v)
  if (n < 1024) return `${f(n, 0)} B`
  if (n < 1024 * 1024) return `${f(n / 1024, 0)} KB`
  return `${f(n / 1024 / 1024, 1)} MB`
}

/** an address as runs: the name, and the address in angle brackets (a mailto: link) */
function addressRuns(list: MailAddress[]): ImportRun[] {
  const runs: ImportRun[] = []
  list.forEach((a, k) => {
    if (k) runs.push({ text: '; ' })
    if (a.name) runs.push({ text: a.name })
    if (a.addr) {
      const href = /^[^\s@<>"]+@[^\s@<>"]+$/.test(a.addr) ? `mailto:${a.addr}` : undefined
      if (a.name) runs.push({ text: ' <', set: { color: GREY } })
      runs.push({ text: a.addr, set: a.name ? { color: GREY } : undefined, ...(href ? { href } : {}) })
      if (a.name) runs.push({ text: '>', set: { color: GREY } })
    }
  })
  return runs
}

const tight = (runs: ImportRun[]): ImportPara => ({ runs, after: 2 })

/** a row of the header table: its label and value */
function row(label: string, value: ImportPara[]): ImportCell[] {
  return [{ paras: [tight([{ text: label, set: { color: GREY } }])] }, { paras: value }]
}

// ---------------------------------------------------------------- bodies

/** a picture of a data: URL (the HTML's own, or a cid: one made from an attachment), at the size the HTML gives or its own */
function dataPicture(url: string, wPt: number, hPt: number, alt?: string): ImportPicture | null {
  const m = /^data:[^,]*?;base64,(.*)$/is.exec(url)
  if (!m) return null
  const data = base64(m[1])
  const ext = pictureExt(data)
  if (!ext) return null
  const own = pictureSize(data)
  let w = wPt * (4 / 3)
  let h = hPt * (4 / 3)
  if (own) {
    if (!w && !h) ({ w, h } = own)
    else if (!w) w = (h * own.w) / own.h
    else if (!h) h = (w * own.h) / own.w
  }
  if (!(w > 0 && h > 0)) return null
  return { data, ext, w, h, ...(alt ? { alt } : {}) }
}

/** clipboard-style paragraphs (the HTML reader's) as document paragraphs */
function fromPasted(paras: PastedPara[]): ImportPara[] {
  return paras.map((p): ImportPara => {
    if (p.table) {
      return {
        runs: [],
        table: {
          rows: p.table.map((r) => r.map((c) => ({ paras: fromPasted(c.paras), colSpan: c.colSpan, rowSpan: c.rowSpan, ...(c.fill ? { fill: c.fill } : {}) }))),
          borders: !!p.border,
        },
      }
    }
    const runs = p.runs.flatMap((r): ImportRun[] => {
      if (r.t === 'text') return [{ text: r.text, set: r.d, href: r.href }]
      if (r.t === 'img') {
        const picture = dataPicture(r.target, r.w, r.h, r.alt)
        return picture ? [{ text: '', picture }] : []
      }
      if (r.t === 'ref' && r.math) return [{ text: '', math: r.math }]
      return []
    })
    return { runs, heading: p.heading, list: p.list, ...(p.level ? { level: p.level } : {}) }
  })
}

const REMOTE_IMG = /<img\b[^>]*?\bsrc\s*=\s*["']?\s*(?:https?:)?\/\//gi

/** an HTML body: its cid: pictures made data: URLs from the attachments; how many pictures it loads from the web */
function htmlBody(html: string, attachments: MailAttachment[], ctx: Ctx): { paras: ImportPara[]; remote: number } {
  const find = (ref: string): MailAttachment | undefined => {
    const r = decodeURIComponent(ref.replace(/^cid:/i, '')).trim().toLowerCase()
    return attachments.find((a) => a.data && ((a.cid && a.cid.toLowerCase() === r) || (a.location && a.location.toLowerCase() === r) || a.name.toLowerCase() === r))
  }
  const resolved = html.replace(/(<img\b[^>]*?\bsrc\s*=\s*)(["']?)([^"'\s>]+)\2/gi, (all, pre: string, _q: string, src: string) => {
    if (/^(?:data:|https?:|\/\/)/i.test(src)) return all
    let a: MailAttachment | undefined
    try {
      a = find(src)
    } catch {
      return all
    }
    if (!a?.data || !pictureExt(a.data)) return all
    ctx.shown.add(a)
    let bin = ''
    for (let k = 0; k < a.data.length; k += 0x8000) bin += String.fromCharCode(...a.data.subarray(k, k + 0x8000))
    return `${pre}"data:${a.type || 'image/png'};base64,${btoa(bin)}"`
  })
  const remote = [...resolved.matchAll(REMOTE_IMG)].length
  return { paras: fromPasted(htmlToParas(resolved, false, { tables: true })), remote }
}

const URL_RE = /\bhttps?:\/\/[^\s<>"]+[^\s<>".,;:!?)\]'}]/g

/** a plain-text body: a paragraph per line, web addresses as links, quoted lines (">") grey */
function textBody(text: string): ImportPara[] {
  return text
    .replace(/^﻿/, '')
    .replace(/\s+$/, '')
    .split(/\r\n|\r|\n/)
    .map((line): ImportPara => {
      const set = /^\s*>/.test(line) ? { color: GREY } : undefined
      const runs: ImportRun[] = []
      let at = 0
      for (const m of line.matchAll(URL_RE)) {
        if (m.index! > at) runs.push({ text: line.slice(at, m.index), set })
        runs.push({ text: m[0], href: m[0], set: { ...set, color: '#0563c1', u: 'single' } })
        at = m.index! + m[0].length
      }
      if (at < line.length) runs.push({ text: line.slice(at), set })
      return { runs, after: 0 }
    })
}

function bodyParas(body: MailBody[], attachments: MailAttachment[], ctx: Ctx): { paras: ImportPara[]; remote: number } {
  const paras: ImportPara[] = []
  let remote = 0
  for (const b of body) {
    if ('html' in b) {
      const h = htmlBody(b.html, attachments, ctx)
      paras.push(...h.paras)
      remote += h.remote
    } else if ('rtf' in b) {
      // (an RTF body other than the document's own: its plain text)
      if (b.text) paras.push(...textBody(b.text))
    } else paras.push(...textBody(b.text))
  }
  return { paras, remote }
}

// ---------------------------------------------------------------- the document

/** the subject, the item's kind, the header table and notices; `depth` > 0: an attached message */
function headParas(msg: MailMessage, ctx: Ctx, depth: number, remote: number): ImportPara[] {
  const t = ctx.texts
  const out: ImportPara[] = []
  if (depth) out.push({ runs: [{ text: t.attached, set: { color: GREY, sz: 9 } }], after: 0 })
  out.push({ runs: [{ text: msg.subject || t.noSubject, set: { b: true, sz: depth ? 13 : 16 } }], heading: depth ? 2 : 1, after: 6 })
  const notes = [itemLabel(msg.itemClass, t), msg.importance === 'high' ? t.highImportance : msg.importance === 'low' ? t.lowImportance : undefined].filter(Boolean) as string[]
  if (notes.length) out.push({ runs: [{ text: notes.join(' · '), set: { color: msg.importance === 'high' ? '#c00000' : GREY } }], after: 6 })
  const rows: ImportCell[][] = []
  if (msg.from.length) rows.push(row(t.from, [tight(addressRuns(msg.from))]))
  if (msg.date || msg.dateText) rows.push(row(t.sent, [tight([{ text: msg.date ? formatDate(msg.date, ctx.locale) : msg.dateText! }])]))
  if (msg.to.length) rows.push(row(t.to, [tight(addressRuns(msg.to))]))
  if (msg.cc.length) rows.push(row(t.cc, [tight(addressRuns(msg.cc))]))
  if (msg.bcc.length) rows.push(row(t.bcc, [tight(addressRuns(msg.bcc))]))
  const listed = msg.attachments.filter((a) => !a.hidden && !ctx.shown.has(a))
  if (listed.length) {
    rows.push(row(t.attachments, listed.map((a) => tight([{ text: a.name }, ...(a.data ? [{ text: `  ${formatSize(a.data.length, ctx.locale)}`, set: { color: GREY } }] : [])]))))
  }
  if (rows.length) out.push({ runs: [], table: { rows, widths: [LABEL_SHARE, 1 - LABEL_SHARE] } })
  if (msg.encrypted) out.push({ runs: [{ text: t.encrypted, set: { i: true, color: GREY } }], after: 4 })
  if (remote) out.push({ runs: [{ text: t.remotePictures, set: { i: true, color: GREY, sz: 9 } }], after: 4 })
  out.push({ runs: [], rule: true, after: 12 })
  return out
}

/** a message and those attached to it, in order */
function messageParas(msg: MailMessage, ctx: Ctx, depth: number): ImportPara[] {
  // the body first: the pictures it shows are not listed as attachments
  const body = bodyParas(msg.body, msg.attachments, ctx)
  const out = [...headParas(msg, ctx, depth, body.remote), ...body.paras]
  for (const a of msg.attachments) if (a.message) out.push({ runs: [], after: 12 }, ...messageParas(a.message, ctx, depth + 1))
  return out
}

/** the attachments that can be saved: those listed in the document, with their bytes */
export function savedAttachments(msg: MailMessage, shown: Set<MailAttachment>): MailAttachment[] {
  return msg.attachments.filter((a) => a.data && !a.hidden && !shown.has(a))
}

/** Builds the document a message is shown as; also the attachments listed in it that can be saved. */
export async function mailToDocx(msg: MailMessage, texts: MailTexts, locale?: string): Promise<{ docx: Blob; attachments: MailAttachment[] }> {
  const ctx: Ctx = { texts, locale, shown: new Set() }
  const only = msg.body.length === 1 ? msg.body[0] : null
  if (only && 'rtf' in only) {
    // Outlook rich text: the RTF converted, the header and the attached messages around it (direct formatting only)
    const rest = msg.attachments.flatMap((a) => (a.message ? [{ runs: [], after: 12 } as ImportPara, ...messageParas(a.message, ctx, 1)] : []))
    const before = parasXml(headParas(msg, ctx, 0, 0))
    try {
      const docx = await rtfToDocx(only.rtf, { before, after: parasXml(rest) })
      return { docx, attachments: savedAttachments(msg, ctx.shown) }
    } catch {
      // damaged RTF: the plain text
      msg = { ...msg, body: only.text ? [{ text: only.text }] : [] }
      ctx.shown.clear()
    }
  }
  const docx = await parasToDocx(messageParas(msg, ctx, 0))
  return { docx, attachments: savedAttachments(msg, ctx.shown) }
}
