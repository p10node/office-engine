/**
 * The RTF body of an Outlook item: compressed (MS-OXRTFCP), and often only a wrapper of the HTML or plain text the
 * message was written in (MS-OXRTFEX), which is taken out of it.
 */
import { charsetCp, decode } from '../legacy/rtf/codepage'
import { CLOSE, END, HEX, Lexer, OPEN, SYMBOL, TEXT, WORD } from '../legacy/rtf/lexer'

/** the dictionary's initial content (MS-OXRTFCP 3.1.1.1.1): 207 bytes */
const PREBUF =
  '{\\rtf1\\ansi\\mac\\deff0\\deftab720{\\fonttbl;}{\\f0\\fnil \\froman \\fswiss \\fmodern \\fscript \\fdecor MS Sans SerifSymbolArialTimes New RomanCourier' +
  '{\\colortbl\\red0\\green0\\blue0\r\n\\par \\pard\\plain\\f0\\fs20\\b\\i\\u\\tab\\tx'

const LZFU = 0x75465a4c
const MELA = 0x414c454d

/** Decompresses PR_RTF_COMPRESSED (its header, then LZ77 runs over a 4 KB ring). */
export function decompressRtf(b: Uint8Array): Uint8Array {
  if (b.length < 16) throw new Error('Compressed RTF: too short')
  const dv = new DataView(b.buffer, b.byteOffset, b.byteLength)
  const compSize = dv.getUint32(0, true)
  const rawSize = dv.getUint32(4, true)
  const type = dv.getUint32(8, true)
  const end = Math.min(b.length, compSize + 4)
  if (type === MELA) return b.slice(16, Math.min(b.length, 16 + rawSize))
  if (type !== LZFU) throw new Error('Compressed RTF: unknown type')
  const dict = new Uint8Array(4096)
  for (let k = 0; k < PREBUF.length; k++) dict[k] = PREBUF.charCodeAt(k)
  let w = PREBUF.length
  // (rawSize is what the writer says, the output may be shorter when the data is damaged; a control byte and 8
  // references of 2 bytes give at most 8 × 17 bytes, so a header cannot ask for more than the data can make)
  const out = new Uint8Array(Math.min(rawSize, Math.ceil(((end - 16) * 136) / 17)) + 18)
  let n = 0
  let p = 16
  while (p < end) {
    const control = b[p++]
    for (let bit = 0; bit < 8 && p < end; bit++) {
      if (!(control & (1 << bit))) {
        const c = b[p++]
        if (n < out.length) out[n++] = c
        dict[w] = c
        w = (w + 1) & 4095
        continue
      }
      if (p + 1 >= end) return out.slice(0, n)
      const ref = (b[p] << 8) | b[p + 1]
      p += 2
      const offset = ref >> 4
      const len = (ref & 15) + 2
      // a reference to where the next byte goes: the end
      if (offset === w) return out.slice(0, Math.min(n, rawSize))
      for (let k = 0; k < len; k++) {
        const c = dict[(offset + k) & 4095]
        if (n < out.length) out[n++] = c
        dict[w] = c
        w = (w + 1) & 4095
      }
    }
  }
  return out.slice(0, Math.min(n, rawSize))
}

/** What an Outlook RTF body is: a wrapped HTML or plain-text body, or RTF of its own (shown as RTF) */
export type RtfBody = { html: string } | { text: string } | { rtf: Uint8Array }

/** groups never shown: font, color and style tables, document information… */
const SKIPPED = new Set(['fonttbl', 'colortbl', 'stylesheet', 'info', 'listtable', 'listoverridetable', 'revtbl', 'rsidtbl', 'generator', 'pict', 'object', 'themedata', 'colorschememapping', 'latentstyles', 'datastore', 'xmlnstbl', 'mmathPr', 'header', 'footer', 'headerl', 'headerr', 'footerl', 'footerr'])

const SYMBOL_TEXT: Record<string, string> = { '\\': '\\', '{': '{', '}': '}', '~': ' ', '_': '‑', '-': '­', tab: '\t', emdash: '—', endash: '–', lquote: '‘', rquote: '’', ldblquote: '“', rdblquote: '”', bullet: '•', emspace: ' ', enspace: ' ', qmspace: ' ' }

/**
 * The body an Outlook RTF body stands for: the HTML (`\fromhtml1`: the HTML's tags are kept in `\*\htmltag` groups,
 * text Outlook added for RTF readers is between `\htmlrtf` and `\htmlrtf0`) or plain text (`\fromtext`) it wraps,
 * else the RTF itself.
 */
export function unwrapRtf(rtf: Uint8Array): RtfBody {
  const head = String.fromCharCode(...rtf.subarray(0, Math.min(rtf.length, 1024)))
  const html = /\\fromhtml1/.test(head)
  if (!html && !/\\fromtext/.test(head)) return { rtf }
  const lx = new Lexer(rtf)
  let ansi = 1252
  const fonts = new Map<number, number>()
  interface State {
    /** inside \htmlrtf: text for RTF readers only */
    rtfOnly: boolean
    /** inside an \*\htmltag group: the HTML's own text */
    tag: boolean
    /** inside \*\mhtmltag: the tag with its URL made a cid: */
    mtag: boolean
    cp: number
    /** \ucN: characters that follow \u to skip */
    uc: number
  }
  let st: State = { rtfOnly: false, tag: false, mtag: false, cp: 1252, uc: 1 }
  const stack: State[] = []
  let out = ''
  let bytes: number[] = []
  let skipChars = 0
  let fontTable = -1
  let curFont = -1
  const flush = () => {
    if (bytes.length) out += decode(Uint8Array.from(bytes), st.cp)
    bytes = []
  }
  const shown = () => (st.tag || st.mtag || !st.rtfOnly) && fontTable < 0
  for (;;) {
    const t = lx.next()
    if (t === END) break
    if (t === OPEN) {
      flush()
      stack.push(st)
      st = { ...st }
      continue
    }
    if (t === CLOSE) {
      flush()
      if (fontTable === stack.length) fontTable = -1
      st = stack.pop() ?? st
      continue
    }
    if (t === WORD) {
      const w = lx.word
      if (skipChars && w !== 'u') skipChars = 0
      switch (w) {
        case 'ansicpg':
          ansi = lx.param || 1252
          st.cp = ansi
          continue
        case 'fonttbl':
          fontTable = stack.length
          continue
        case 'f':
          if (fontTable >= 0) curFont = lx.param
          else {
            flush()
            st.cp = fonts.get(lx.param) ?? ansi
          }
          continue
        case 'fcharset':
          if (fontTable >= 0 && curFont >= 0) fonts.set(curFont, charsetCp(lx.param, ansi))
          continue
        case 'htmlrtf':
          flush()
          st.rtfOnly = !lx.hasParam || lx.param !== 0
          continue
        // (an \*\mhtmltag holds a tag whose URL was made a cid:, next to the \*\htmltag with the original URL: both are kept,
        // and the picture that is not a cid: is left out when the HTML is read)
        case 'htmltag':
          flush()
          st.tag = true
          continue
        case 'mhtmltag':
          flush()
          st.mtag = true
          continue
        case 'uc':
          st.uc = lx.param
          continue
        case 'u':
          if (!shown()) continue
          flush()
          out += String.fromCharCode(lx.param < 0 ? lx.param + 65536 : lx.param)
          skipChars = st.uc
          continue
        case 'par':
        case 'line':
          if (shown()) {
            flush()
            out += html && !st.tag && !st.mtag ? '\r\n' : '\n'
          }
          continue
        case 'htmlbase':
          continue
      }
      if (SKIPPED.has(w) && fontTable < 0) {
        flush()
        lx.skipGroup()
        st = stack.pop() ?? st
        continue
      }
      const sym = SYMBOL_TEXT[w]
      if (sym !== undefined && shown()) {
        flush()
        out += sym
      }
      continue
    }
    if (t === SYMBOL) {
      if (lx.word === '*') {
        // an ignorable destination other than the HTML's tags: left out
        const save = lx.pos
        const next = lx.next()
        const word: string = lx.word
        lx.pos = save
        if (next === WORD && (word === 'htmltag' || word === 'mhtmltag')) continue
        flush()
        lx.skipGroup()
        st = stack.pop() ?? st
        continue
      }
      if (skipChars) {
        skipChars--
        continue
      }
      const sym = SYMBOL_TEXT[lx.word]
      if (sym !== undefined && shown()) {
        flush()
        out += sym
      }
      continue
    }
    if (t === HEX) {
      if (skipChars) {
        skipChars--
        continue
      }
      if (shown()) bytes.push(lx.param)
      continue
    }
    if (t === TEXT) {
      let s = lx.start
      if (skipChars) {
        const k = Math.min(skipChars, lx.end - s)
        s += k
        skipChars -= k
      }
      if (fontTable >= 0 || !shown()) continue
      for (let k = s; k < lx.end; k++) bytes.push(rtf[k])
      continue
    }
  }
  flush()
  return html ? { html: out } : { text: out.replace(/\r\n/g, '\n') }
}
