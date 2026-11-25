/** A mail message as read from an Internet message (.eml, .emlx) or an Outlook item (.msg, .oft), for the document it is shown as. */

/** extensions of mail messages: Internet messages, Apple Mail's, Outlook items and templates */
export const MAIL_EXT: Record<string, 'eml' | 'emlx' | 'msg'> = { eml: 'eml', emlx: 'emlx', msg: 'msg', oft: 'msg' }

export interface MailAddress {
  /** the display name ('' when the address has none) */
  name: string
  /** the e-mail address ('' when only a name is known: an Exchange user, Outlook's display strings) */
  addr: string
}

/** a piece of the message's body, in order: most messages have one. An RTF body has its plain text with it (Outlook's) */
export type MailBody = { html: string } | { text: string } | { rtf: Uint8Array; text?: string }

export interface MailAttachment {
  name: string
  /** MIME type ('' when unknown) */
  type: string
  /** the file's bytes; undefined for an attached Outlook item (its message is `message`) */
  data?: Uint8Array
  /** Content-ID, without the angle brackets: a picture the HTML body shows as `cid:…` */
  cid?: string
  /** Content-Location: a picture the HTML body names by its URL */
  location?: string
  /** an attached message (forwarded as an attachment), read */
  message?: MailMessage
  /** Outlook: hidden from the attachment list (pictures of the body, signatures) */
  hidden?: boolean
}

export interface MailMessage {
  subject: string
  from: MailAddress[]
  to: MailAddress[]
  cc: MailAddress[]
  bcc: MailAddress[]
  /** when it was sent (else received, else written); undefined when the file says none, or nothing that reads as a date */
  date?: Date
  /** the Date header as written, when it does not read as a date */
  dateText?: string
  body: MailBody[]
  attachments: MailAttachment[]
  importance?: 'high' | 'low'
  /** Outlook's message class when it is not a plain message: "IPM.Schedule.Meeting.Request", "IPM.Contact"… */
  itemClass?: string
  /** S/MIME or PGP encrypted: the body cannot be read */
  encrypted?: boolean
}

/** the texts written into the document a message opens as (the viewer gives them in its language) */
export interface MailTexts {
  from: string
  sent: string
  to: string
  cc: string
  bcc: string
  attachments: string
  /** the heading of a message attached to the one shown */
  attached: string
  noSubject: string
  highImportance: string
  lowImportance: string
  /** said in place of the pictures an HTML body loads from the web, which are never fetched */
  remotePictures: string
  encrypted: string
  /** Outlook items other than messages, by the start of their message class */
  item: Record<'meetingRequest' | 'meetingCanceled' | 'meetingAccepted' | 'meetingDeclined' | 'meetingTentative' | 'appointment' | 'contact' | 'task' | 'note' | 'report', string>
}

export const MAIL_TEXTS: MailTexts = {
  from: 'From',
  sent: 'Sent',
  to: 'To',
  cc: 'Cc',
  bcc: 'Bcc',
  attachments: 'Attachments',
  attached: 'Attached message',
  noSubject: '(no subject)',
  highImportance: 'High importance',
  lowImportance: 'Low importance',
  remotePictures: 'Pictures from the web are not shown: nothing is downloaded when a message is opened.',
  encrypted: 'This message is encrypted: its content cannot be shown.',
  item: {
    meetingRequest: 'Meeting request',
    meetingCanceled: 'Meeting canceled',
    meetingAccepted: 'Meeting accepted',
    meetingDeclined: 'Meeting declined',
    meetingTentative: 'Meeting tentatively accepted',
    appointment: 'Appointment',
    contact: 'Contact',
    task: 'Task',
    note: 'Note',
    report: 'Delivery report',
  },
}

/** what the viewer is told of an opened message: its attachments, which `mail.attachment` gives */
export interface MailInfo {
  subject: string
  attachments: { name: string; type: string; size: number }[]
}
