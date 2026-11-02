import type { Inline, RunSet, RunStyle } from '../../core/docx/types'
import { cssFont } from '../../core/fonts'

export interface PastedPara {
  runs: Inline[]
  heading?: number
  list?: 'bullet' | 'number'
}

type Ctx = RunStyle & { href?: string; fontName?: string }

const BLOCK = new Set(['P', 'DIV', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'LI', 'TR', 'BLOCKQUOTE', 'PRE', 'SECTION', 'ARTICLE', 'HEADER', 'FOOTER', 'UL', 'OL', 'TABLE', 'TBODY', 'THEAD', 'DL', 'DT', 'DD', 'FIGURE'])
const SKIP = new Set(['SCRIPT', 'STYLE', 'HEAD', 'TITLE', 'META', 'LINK', 'NOSCRIPT', 'TEMPLATE'])

function hex(color: string | null | undefined): string | undefined {
  if (!color) return undefined
  const c = color.trim().toLowerCase()
  if (/^#[0-9a-f]{6}$/.test(c)) return c
  if (/^#[0-9a-f]{3}$/.test(c)) return '#' + [...c.slice(1)].map((x) => x + x).join('')
  const m = /^rgba?\((\d+),\s*(\d+),\s*(\d+)(?:,\s*([\d.]+))?\)/.exec(c)
  if (m) {
    if (m[4] !== undefined && Number(m[4]) === 0) return undefined
    return '#' + [m[1], m[2], m[3]].map((v) => Number(v).toString(16).padStart(2, '0')).join('')
  }
  return undefined
}

function sizePt(v: string): number | undefined {
  const m = /^([\d.]+)(pt|px|em|rem)?$/.exec(v.trim())
  if (!m) return undefined
  const n = Number(m[1])
  return m[2] === 'px' ? n * 0.75 : m[2] === 'em' || m[2] === 'rem' ? n * 11 : n
}

function styleOf(el: HTMLElement, base: Ctx): Ctx {
  const s: Ctx = { ...base }
  switch (el.tagName) {
    case 'B':
    case 'STRONG':
      s.b = true
      break
    case 'I':
    case 'EM':
      s.i = true
      break
    case 'U':
    case 'INS':
      s.u = 'single'
      break
    case 'S':
    case 'STRIKE':
    case 'DEL':
      s.strike = true
      break
    case 'SUP':
      s.va = 'sup'
      break
    case 'SUB':
      s.va = 'sub'
      break
    case 'A':
      if ((el as HTMLAnchorElement).href && !(el as HTMLAnchorElement).href.startsWith('javascript:')) {
        s.href = el.getAttribute('href') ?? undefined
        s.u = 'single'
        s.color = '#0563c1'
      }
      break
    case 'FONT': {
      const c = hex(el.getAttribute('color'))
      if (c) s.color = c
      const f = el.getAttribute('face')
      if (f) s.fontName = f.split(',')[0].replace(/['"]/g, '').trim()
      break
    }
    case 'CODE':
    case 'KBD':
      s.fontName = 'Courier New'
      break
  }
  const st = el.style
  if (st.fontWeight) s.b = st.fontWeight === 'bold' || st.fontWeight === 'bolder' || Number(st.fontWeight) >= 600
  if (st.fontStyle) s.i = st.fontStyle === 'italic' || st.fontStyle === 'oblique'
  const deco = `${st.textDecoration} ${st.textDecorationLine}`
  if (deco.includes('underline')) s.u = 'single'
  if (deco.includes('line-through')) s.strike = true
  if (st.textDecoration === 'none' && el.tagName !== 'A') s.u = undefined
  const c = hex(st.color)
  if (c) s.color = c
  const bg = hex(st.backgroundColor)
  if (bg && bg !== '#ffffff') s.hl = bg
  if (st.fontSize) {
    const pt = sizePt(st.fontSize)
    if (pt && pt >= 4 && pt <= 200) s.sz = Math.round(pt * 2) / 2
  }
  if (st.fontFamily) s.fontName = st.fontFamily.split(',')[0].replace(/['"]/g, '').trim()
  if (st.verticalAlign === 'super') s.va = 'sup'
  else if (st.verticalAlign === 'sub') s.va = 'sub'
  return s
}

function toRun(text: string, c: Ctx): Inline {
  const { href, fontName, ...rest } = c
  const s: RunStyle = { ...rest }
  if (fontName) s.font = cssFont(fontName)
  const d: RunSet = {}
  if (c.b !== undefined) d.b = c.b
  if (c.i !== undefined) d.i = c.i
  if (c.u) d.u = c.u
  if (c.strike) d.strike = true
  if (c.color) d.color = c.color
  if (c.sz) d.sz = c.sz
  if (fontName) d.font = fontName
  if (c.hl) d.hl = c.hl
  if (c.va) d.va = c.va
  return { t: 'text', text, s, d: Object.keys(d).length ? d : undefined, href } as Inline
}

/** Converts clipboard HTML (Word, Google Docs, web pages) into paragraphs of formatted runs. */
export function htmlToParas(html: string): PastedPara[] {
  const doc = new DOMParser().parseFromString(html, 'text/html')
  const out: PastedPara[] = []
  let cur: PastedPara | null = null
  const para = (props: Partial<PastedPara> = {}) => {
    cur = { runs: [], ...props }
    out.push(cur)
    return cur
  }
  const walk = (node: Node, c: Ctx, list: 'bullet' | 'number' | undefined, pre: boolean) => {
    for (const child of Array.from(node.childNodes)) {
      if (child.nodeType === 3) {
        let text = child.textContent ?? ''
        if (!pre) text = text.replace(/[\s ]+/g, (m) => (m.includes(' ') && m.length === 1 ? ' ' : ' '))
        if (!text || (!pre && !cur && !text.trim())) continue
        ;(cur ?? para()).runs.push(toRun(text, c))
        continue
      }
      if (child.nodeType !== 1) continue
      const el = child as HTMLElement
      if (SKIP.has(el.tagName)) continue
      if (el.tagName === 'BR') {
        ;(cur ?? para()).runs.push(toRun('\n', c))
        continue
      }
      if (el.tagName === 'TD' || el.tagName === 'TH') {
        if (cur && (cur as PastedPara).runs.length) (cur as PastedPara).runs.push(toRun('\t', c))
        walk(el, styleOf(el, c), list, pre)
        continue
      }
      const st = styleOf(el, c)
      if (BLOCK.has(el.tagName)) {
        const nextList = el.tagName === 'UL' ? 'bullet' : el.tagName === 'OL' ? 'number' : list
        if (el.tagName !== 'UL' && el.tagName !== 'OL' && el.tagName !== 'TABLE' && el.tagName !== 'TBODY' && el.tagName !== 'THEAD') {
          const h = /^H([1-6])$/.exec(el.tagName)
          para({ heading: h ? Number(h[1]) : undefined, list: el.tagName === 'LI' ? nextList ?? 'bullet' : undefined })
        }
        walk(el, st, nextList, pre || el.tagName === 'PRE')
        cur = null
        continue
      }
      walk(el, st, list, pre)
    }
  }
  walk(doc.body, {}, undefined, false)
  // trim whitespace-only edges
  for (const p of out) {
    const first = p.runs[0]
    if (first?.t === 'text') first.text = first.text.replace(/^ +/, '')
    const last = p.runs[p.runs.length - 1]
    if (last?.t === 'text') last.text = last.text.replace(/ +$/, '')
    p.runs = p.runs.filter((r) => r.t !== 'text' || r.text)
  }
  while (out.length && !out[0].runs.length) out.shift()
  while (out.length && !out[out.length - 1].runs.length) out.pop()
  return out
}
