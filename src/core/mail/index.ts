/**
 * Mail messages, opened read-only as a Word document: Internet messages (.eml, Apple Mail's .emlx) and Outlook items
 * (.msg, .oft templates). The message's attachments can be saved from the viewer.
 */
import { Cfb } from '../legacy/cfb'
import { extOf } from '../plain'
import { mailToDocx } from './docx'
import { looksLikeEml, readEml, readEmlx } from './mime'
import { isMsgCfb, readMsg } from './msg'
import { MAIL_EXT, MAIL_TEXTS, type MailAttachment, type MailInfo, type MailTexts } from './types'

export { MAIL_EXT, MAIL_TEXTS, type MailInfo, type MailTexts } from './types'
export { readEml, readEmlx } from './mime'
export { readMsg } from './msg'
export { mailToDocx } from './docx'

export type MailFormat = 'eml' | 'msg'

/** a compound file that is an Outlook item (its directory says so; the file is read whole) */
export function isMsg(bytes: Uint8Array): boolean {
  try {
    return isMsgCfb(Cfb.open(bytes))
  } catch {
    return false
  }
}

/**
 * Whether a file that is not a compound file is an Internet message: its name (.eml, .emlx), its type
 * (message/rfc822), or (no name that says otherwise) its first bytes, which are mail headers.
 */
export function mailTextFormat(name: string | undefined, type: string | undefined, head: Uint8Array): 'eml' | 'emlx' | null {
  const ext = extOf(name)
  if (MAIL_EXT[ext] === 'eml' || MAIL_EXT[ext] === 'emlx') return MAIL_EXT[ext] as 'eml' | 'emlx'
  // (some programs save Internet messages as .msg)
  if (MAIL_EXT[ext] === 'msg') {
    if (looksLikeEml(head)) return 'eml'
    throw new Error(`${name ? `“${name}”` : 'The file'} is not an Outlook message (.msg): it is not a compound file, or it is damaged`)
  }
  if (/^message\/rfc822\b/i.test(type ?? '')) return 'eml'
  // a name with another extension is that format (a web page saved as .mht is MIME too)
  if (ext) return null
  return looksLikeEml(head) ? 'eml' : null
}

export interface OpenedMail {
  /** the document the message is shown as */
  docx: Blob
  from: MailFormat
  info: MailInfo
  /** the attachments that can be saved, in the order of `info.attachments` */
  attachments: MailAttachment[]
}

/** Reads a message and builds its document. */
export async function openMail(bytes: Uint8Array, format: 'eml' | 'emlx' | 'msg', texts: Partial<MailTexts> = {}, locale?: string): Promise<OpenedMail> {
  let msg
  try {
    msg = format === 'msg' ? readMsg(bytes) : format === 'emlx' ? readEmlx(bytes) : readEml(bytes)
  } catch (e) {
    throw new Error(`The message cannot be read: ${(e as Error).message.replace(/^./, (c) => c.toLowerCase())} (an incomplete download or copy?)`)
  }
  const t: MailTexts = { ...MAIL_TEXTS, ...texts, item: { ...MAIL_TEXTS.item, ...texts.item } }
  const { docx, attachments } = await mailToDocx(msg, t, locale)
  return {
    docx,
    from: format === 'msg' ? 'msg' : 'eml',
    info: { subject: msg.subject, attachments: attachments.map((a) => ({ name: a.name, type: a.type, size: a.data?.length ?? 0 })) },
    attachments,
  }
}
