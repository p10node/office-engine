/**
 * DOCX write-back: edited paragraphs are regenerated from their original XML
 * (pPr / rPr / inline objects copied verbatim, only changed properties patched);
 * everything else is written back byte-for-byte.
 */
import { relOfType } from '../opc'
import {
  clone, el, isNode, kidOf, kidsOf, local, parseRaw, prefixOf, removeKids, rootEl, serialize, setAttr, setOrdered, walkOoxml,
  escText, type RChild, type RNode,
} from '../raw'
import { parseXml } from '../xml'
import { Numbering } from './numbering'
import type { DocxDocument, GenRun } from './parser'
import { HIGHLIGHT } from './styles'
import type { Block, CanonStyle, CellSpec, EditOp, EditResult, ParaDirty, ParaSpec, Paragraph, RunSet, SpecRun, SrcRef, StyleInfo, TableOp, TableSpec } from './types'

export const DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
const NUMBERING_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering'
const NUMBERING_CT = 'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml'

const PPR_ORDER = [
  'pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr', 'widowControl', 'numPr', 'suppressLineNumbers', 'pBdr', 'shd',
  'tabs', 'suppressAutoHyphens', 'kinsoku', 'wordWrap', 'overflowPunct', 'topLinePunct', 'autoSpaceDE', 'autoSpaceDN', 'bidi',
  'adjustRightInd', 'snapToGrid', 'spacing', 'ind', 'contextualSpacing', 'mirrorIndents', 'suppressOverlap', 'jc', 'textDirection',
  'textAlignment', 'textboxTightWrap', 'outlineLvl', 'divId', 'cnfStyle', 'rPr', 'sectPr', 'pPrChange',
]
const RPR_ORDER = [
  'rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps', 'strike', 'dstrike', 'outline', 'shadow', 'emboss', 'imprint',
  'noProof', 'snapToGrid', 'vanish', 'webHidden', 'color', 'spacing', 'w', 'kern', 'position', 'sz', 'szCs', 'highlight', 'u',
  'effect', 'bdr', 'shd', 'fitText', 'vertAlign', 'rtl', 'cs', 'em', 'lang', 'eastAsianLayout', 'specVanish', 'oMath',
]
const START_MARKS = new Set(['bookmarkStart', 'commentRangeStart', 'permStart', 'moveFromRangeStart', 'moveToRangeStart'])
const END_MARKS = new Set(['bookmarkEnd', 'commentRangeEnd', 'permEnd', 'moveFromRangeEnd', 'moveToRangeEnd'])
const RUN_STOP = new Set(['p', 'del', 'moveFrom', 'oMath', 'oMathPara', 'pPr', 'rPr', 'r'])
const HL_NAME: Record<string, string> = Object.fromEntries(Object.entries(HIGHLIGHT).map(([k, v]) => [v, k]))

interface ParaLoc {
  node: RNode
  list: RChild[]
  index: number
}

/** Paragraphs of a body item in the same preorder the converter uses. */
export function collectParasRaw(nodes: RChild[]): ParaLoc[] {
  const out: ParaLoc[] = []
  nodes.forEach((n, i) => {
    if (!isNode(n)) return
    if (local(n.q) === 'p') out.push({ node: n, list: nodes, index: i })
    walkOoxml(n, (c, parent, idx) => {
      if (local(c.q) === 'p') out.push({ node: c, list: parent.c, index: idx })
    })
  })
  return out
}

export function collectRunsRaw(p: RNode): RNode[] {
  const out: RNode[] = []
  walkOoxml(p, (c) => {
    const l = local(c.q)
    if (l === 'r') out.push(c)
    return RUN_STOP.has(l)
  })
  return out
}

// ---------------------------------------------------------------- source lookup

const itemCache = new WeakMap<DocxDocument, Map<number, ParaLoc[]>>()

function itemParas(doc: DocxDocument, item: number): ParaLoc[] {
  let m = itemCache.get(doc)
  if (!m) itemCache.set(doc, (m = new Map()))
  let paras = m.get(item)
  if (!paras) {
    paras = collectParasRaw(parseRaw(doc.items[item] ?? ''))
    if (m.size > 64) m.delete(m.keys().next().value as number)
    m.set(item, paras)
  }
  return paras
}

export function origPara(doc: DocxDocument, pid: number): RNode | undefined {
  const loc = doc.origLoc.get(pid)
  if (!loc) return undefined
  return itemParas(doc, loc.item)[loc.ord]?.node
}

const runCache = new WeakMap<RNode, RNode[]>()
function origRun(doc: DocxDocument, src: SrcRef | undefined): RNode | undefined {
  if (!src) return undefined
  const p = origPara(doc, src[0])
  if (!p) return undefined
  let runs = runCache.get(p)
  if (!runs) runCache.set(p, (runs = collectRunsRaw(p)))
  return runs[src[1]]
}

// ---------------------------------------------------------------- generation

function applyRunSet(rPr: RNode, set: RunSet, W: string) {
  const val = W + 'val'
  const put = (l: string, node: RNode | undefined) => setOrdered(rPr, node, l, RPR_ORDER)
  const toggle = (l: string, on: boolean) => put(l, on ? el(W + l) : el(W + l, { [val]: '0' }))
  if (set.b !== undefined) {
    toggle('b', set.b)
    toggle('bCs', set.b)
  }
  if (set.i !== undefined) {
    toggle('i', set.i)
    toggle('iCs', set.i)
  }
  if (set.strike !== undefined) {
    toggle('strike', set.strike)
    if (set.strike) put('dstrike', undefined)
  }
  if (set.u !== undefined) put('u', el(W + 'u', { [val]: set.u && set.u !== 'none' ? set.u : 'none' }))
  if ('color' in set) put('color', set.color ? el(W + 'color', { [val]: set.color.replace('#', '').toUpperCase() }) : undefined)
  if (set.sz !== undefined) {
    const v = String(Math.round(set.sz * 2))
    put('sz', el(W + 'sz', { [val]: v }))
    put('szCs', el(W + 'szCs', { [val]: v }))
  }
  if (set.font !== undefined) {
    const f = set.font
    put('rFonts', el(W + 'rFonts', { [W + 'ascii']: f, [W + 'hAnsi']: f, [W + 'eastAsia']: f, [W + 'cs']: f }))
  }
  if ('hl' in set) {
    const name = set.hl ? HL_NAME[set.hl.toLowerCase()] : undefined
    if (!set.hl) {
      put('highlight', undefined)
      put('shd', undefined)
    } else if (name) put('highlight', el(W + 'highlight', { [val]: name }))
    else put('shd', el(W + 'shd', { [val]: 'clear', [W + 'color']: 'auto', [W + 'fill']: set.hl.replace('#', '').toUpperCase() }))
  }
  if ('va' in set) put('vertAlign', set.va ? el(W + 'vertAlign', { [val]: set.va === 'sup' ? 'superscript' : 'subscript' }) : undefined)
}

function textContentXml(text: string, W: string): string {
  let out = ''
  for (const piece of text.split(/(\t|\n)/)) {
    if (!piece) continue
    if (piece === '\t') out += `<${W}tab/>`
    else if (piece === '\n') out += `<${W}br/>`
    else out += `<${W}t xml:space="preserve">${escText(piece)}</${W}t>`
  }
  return out
}

function abstractXml(W: string, id: number, kind: 'bullet' | 'number'): string {
  const bullets = ['•', '◦', '▪']
  const fmts = ['decimal', 'lowerLetter', 'lowerRoman']
  let lv = ''
  for (let l = 0; l < 9; l++) {
    const fmt = kind === 'bullet' ? 'bullet' : fmts[l % 3]
    const text = kind === 'bullet' ? bullets[l % 3] : `%${l + 1}.`
    lv +=
      `<${W}lvl ${W}ilvl="${l}"><${W}start ${W}val="1"/><${W}numFmt ${W}val="${fmt}"/><${W}lvlText ${W}val="${text}"/>` +
      `<${W}lvlJc ${W}val="left"/><${W}pPr><${W}ind ${W}left="${720 * (l + 1)}" ${W}hanging="360"/></${W}pPr></${W}lvl>`
  }
  return `<${W}abstractNum ${W}abstractNumId="${id}"><${W}multiLevelType ${W}val="hybridMultilevel"/>${lv}</${W}abstractNum>`
}

const STYLES_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles'
const STYLES_CT = 'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml'

/** Word's built-in look for the canonical styles (used when a document lacks them) */
function canonDef(kind: Exclude<CanonStyle, 'normal'>, W: string): { id: string; name: string; ppr: string; rpr: string } {
  const sz = (v: number) => `<${W}sz ${W}val="${v}"/><${W}szCs ${W}val="${v}"/>`
  const color = (c: string) => `<${W}color ${W}val="${c}"/>`
  if (kind === 'title')
    return { id: 'Title', name: 'Title', ppr: `<${W}spacing ${W}after="80" ${W}line="240" ${W}lineRule="auto"/><${W}contextualSpacing/>`, rpr: `<${W}spacing ${W}val="-10"/><${W}kern ${W}val="28"/>${sz(56)}` }
  if (kind === 'subtitle') return { id: 'Subtitle', name: 'Subtitle', ppr: `<${W}spacing ${W}after="160"/>`, rpr: `${color('5A5A5A')}<${W}spacing ${W}val="15"/>${sz(28)}` }
  const n = Number(kind.slice(1))
  const look: Record<number, [number, number, string, boolean]> = {
    1: [240, 32, '2F5496', false],
    2: [40, 26, '2F5496', false],
    3: [40, 24, '1F3763', false],
    4: [40, 22, '2F5496', true],
    5: [40, 22, '2F5496', false],
    6: [40, 22, '1F3763', true],
  }
  const [before, size, c, italic] = look[n]
  return {
    id: `Heading${n}`,
    name: `heading ${n}`,
    ppr: `<${W}keepNext/><${W}keepLines/><${W}spacing ${W}before="${before}" ${W}after="0"/><${W}outlineLvl ${W}val="${n - 1}"/>`,
    rpr: `${italic ? `<${W}i/><${W}iCs/>` : ''}${color(c)}${sz(size)}`,
  }
}

/** Id of a built-in paragraph style (Title, Heading 1…), added to styles.xml when the document lacks it. */
export async function ensureStyle(doc: DocxDocument, kind: CanonStyle): Promise<{ id: string; styles: StyleInfo[] }> {
  const have = doc.styles.canonical()[kind]
  if (have || kind === 'normal') return { id: have ?? '', styles: doc.styles.quickStyles() }
  const rel = relOfType(doc.rels, 'styles')
  const path = rel?.target ?? 'word/styles.xml'
  const xml = doc.pkg.has(path)
    ? await doc.pkg.text(path)
    : '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:styles>'
  const nodes = parseRaw(xml)
  const root = rootEl(nodes)!
  const W = prefixOf(root.q)
  const def = canonDef(kind, W)
  let id = def.id
  for (let i = 2; doc.styles.has(id); i++) id = def.id + i
  const base = doc.styles.defaultPara
  const link = base ? `<${W}basedOn ${W}val="${base}"/><${W}next ${W}val="${base}"/>` : ''
  const styleXml =
    `<${W}style ${W}type="paragraph" ${W}styleId="${id}"><${W}name ${W}val="${def.name}"/>${link}<${W}uiPriority ${W}val="9"/><${W}qFormat/>` +
    `<${W}pPr>${def.ppr}</${W}pPr><${W}rPr>${def.rpr}</${W}rPr></${W}style>`
  root.c.push(parseRaw(styleXml)[0])
  root.self = false
  doc.pkg.setPart(path, serialize(nodes))
  if (!rel) {
    await doc.pkg.addRel(doc.mainPath, STYLES_REL, path)
    await doc.pkg.setContentType(path, STYLES_CT)
    doc.rels = await doc.pkg.rels(doc.mainPath)
  }
  const parsed = parseXml(W ? styleXml.replace(`<${W}style `, `<${W}style xmlns:${W.slice(0, -1)}="http://schemas.openxmlformats.org/wordprocessingml/2006/main" `) : styleXml)
  if (parsed) doc.styles.add(parsed)
  doc.markEdited()
  return { id, styles: doc.styles.quickStyles() }
}

/** Finds (or creates in numbering.xml) a list definition of the given kind. */
async function ensureList(doc: DocxDocument, kind: 'bullet' | 'number'): Promise<string> {
  const found = doc.numbering.findList(kind)
  if (found) return found
  const rel = relOfType(doc.rels, 'numbering')
  const path = rel?.target ?? 'word/numbering.xml'
  const xml = doc.pkg.has(path)
    ? await doc.pkg.text(path)
    : '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:numbering>'
  const nodes = parseRaw(xml)
  const root = rootEl(nodes)!
  const W = prefixOf(root.q)
  const { abstract, num } = doc.numbering.ids()
  const absNode = parseRaw(abstractXml(W, abstract, kind))[0] as RNode
  const numNode = parseRaw(`<${W}num ${W}numId="${num}"><${W}abstractNumId ${W}val="${abstract}"/></${W}num>`)[0] as RNode
  const idxOf = (l: string, last: boolean) => {
    let at = -1
    root.c.forEach((c, i) => {
      if (isNode(c) && local(c.q) === l && (last || at < 0)) at = i
    })
    return at
  }
  const lastAbs = idxOf('abstractNum', true)
  const firstNum = idxOf('num', false)
  root.c.splice(lastAbs >= 0 ? lastAbs + 1 : firstNum >= 0 ? firstNum : root.c.length, 0, absNode)
  const lastNum = idxOf('num', true)
  const cleanup = idxOf('numIdMacAtCleanup', false)
  root.c.splice(lastNum >= 0 ? lastNum + 1 : cleanup >= 0 ? cleanup : root.c.length, 0, numNode)
  root.self = false
  const out = serialize(nodes)
  doc.pkg.setPart(path, out)
  if (!rel) {
    await doc.pkg.addRel(doc.mainPath, NUMBERING_REL, path)
    await doc.pkg.setContentType(path, NUMBERING_CT)
    doc.rels = await doc.pkg.rels(doc.mainPath)
  }
  doc.numbering = new Numbering(parseXml(out), doc.theme)
  return String(num)
}

async function applyParaDirty(doc: DocxDocument, pPr: RNode, pd: ParaDirty, W: string) {
  const val = W + 'val'
  const put = (l: string, node: RNode | undefined) => setOrdered(pPr, node, l, PPR_ORDER)
  if (pd.style !== undefined) put('pStyle', pd.style && pd.style !== doc.styles.defaultPara ? el(W + 'pStyle', { [val]: pd.style }) : undefined)
  if (pd.align) put('jc', el(W + 'jc', { [val]: pd.align === 'justify' ? 'both' : pd.align }))
  if (pd.line !== undefined || pd.before !== undefined || pd.after !== undefined) {
    const sp = clone(kidOf(pPr, 'spacing') ?? el(W + 'spacing'))
    if (pd.before !== undefined) {
      setAttr(sp, W + 'before', String(Math.round(pd.before * 20)))
      setAttr(sp, W + 'beforeAutospacing', undefined)
    }
    if (pd.after !== undefined) {
      setAttr(sp, W + 'after', String(Math.round(pd.after * 20)))
      setAttr(sp, W + 'afterAutospacing', undefined)
    }
    if (pd.line !== undefined) {
      setAttr(sp, W + 'line', String(Math.round(pd.line * 240)))
      setAttr(sp, W + 'lineRule', 'auto')
    }
    put('spacing', sp)
  }
  if (pd.pageBreak !== undefined) put('pageBreakBefore', pd.pageBreak ? el(W + 'pageBreakBefore') : undefined)
  if (pd.indL !== undefined) {
    const ind = clone(kidOf(pPr, 'ind') ?? el(W + 'ind'))
    setAttr(ind, W + 'start', undefined)
    setAttr(ind, W + 'left', String(Math.max(0, Math.round(pd.indL * 20))))
    put('ind', ind)
  }
  const styleNum = doc.styles.para(readStyle(pPr)).numPr
  const numPr = (numId: string, ilvl: number) =>
    el(W + 'numPr', {}, [el(W + 'ilvl', { [val]: String(ilvl) }), el(W + 'numId', { [val]: numId })])
  if (pd.list === 'none') put('numPr', styleNum ? numPr('0', 0) : undefined)
  else if (pd.list) put('numPr', numPr(await ensureList(doc, pd.list), pd.ilvl ?? 0))
  else if (pd.ilvl !== undefined) {
    const cur = kidOf(pPr, 'numPr')
    const numId = (cur && attrOf(kidOf(cur, 'numId'), val)) ?? styleNum?.numId
    if (numId && numId !== '0') put('numPr', numPr(numId, pd.ilvl))
  }
}

function attrOf(n: RNode | undefined, q: string): string | undefined {
  if (!n) return undefined
  const m = new RegExp(`(?:^|\\s)${q.replace(':', '\\:')}="([^"]*)"`).exec(n.a)
  return m?.[1]
}

function readStyle(pPr: RNode): string | undefined {
  const ps = kidOf(pPr, 'pStyle')
  return ps ? /="([^"]*)"/.exec(ps.a)?.[1] : undefined
}

export async function genParagraph(doc: DocxDocument, spec: ParaSpec): Promise<{ xml: string; map: GenRun[] }> {
  const tpl = origPara(doc, spec.from)
  const pq = tpl?.q ?? 'w:p'
  const W = prefixOf(pq) || 'w:'
  const pPr = clone(kidOf(tpl, 'pPr') ?? el(W + 'pPr'))
  if (spec.pd) await applyParaDirty(doc, pPr, spec.pd, W)
  const keepSect = spec.keepSect ?? spec.pid === spec.from
  if (!keepSect) removeKids(pPr, 'sectPr')
  const markRPr = kidOf(pPr, 'rPr')

  const map: GenRun[] = []
  let body = ''
  let openLink: string | undefined
  const closeLink = () => {
    if (openLink !== undefined) body += `</${W}hyperlink>`
    openLink = undefined
  }
  const runXml = (rPr: RNode | undefined, content: string) => `<${W}r>${rPr && rPr.c.length ? serialize(rPr) : ''}${content}</${W}r>`
  const baseRPr = (src: SrcRef | undefined) => {
    const sr = origRun(doc, src)
    const r = kidOf(sr, 'rPr') ?? markRPr
    if (!r) return undefined
    const c = clone(r)
    c.q = W + 'rPr'
    removeKids(c, (l) => l === 'rPrChange' || l === 'ins' || l === 'del' || l === 'moveFrom' || l === 'moveTo')
    return c
  }

  spec.parts.forEach((part, pi) => {
    if (pi > 0) {
      closeLink()
      body += `<${W}r><${W}br ${W}type="page"/></${W}r>`
      map.push({})
    }
    for (const run of part as SpecRun[]) {
      if ('obj' in run) {
        closeLink()
        if (run.obj === 'field') {
          let rPr = baseRPr(undefined)
          if (run.set) applyRunSet((rPr ??= el(W + 'rPr')), run.set, W)
          body += `<${W}fldSimple ${W}instr=" ${run.kind} ">${runXml(rPr, `<${W}t>1</${W}t>`)}</${W}fldSimple>`
          map.push({ d: run.set })
        } else if (run.obj === 'image') {
          body += imageRunXml(W, run.rid, run.w, run.h, run.name, run.alt)
          map.push({})
        } else {
          const src = origRun(doc, run.src)
          if (src) {
            if (run.size || run.alt !== undefined || run.rid) {
              const r = clone(src)
              const cx = run.size && String(Math.round(run.size[0] * 12700))
              const cy = run.size && String(Math.round(run.size[1] * 12700))
              walkOoxml(r, (n) => {
                const l = local(n.q)
                if (cx && cy && (l === 'extent' || (l === 'ext' && /\scx=/.test(n.a)))) {
                  setAttr(n, 'cx', cx)
                  setAttr(n, 'cy', cy)
                }
                if (run.alt !== undefined && (l === 'docPr' || (l === 'cNvPr' && /\sid=/.test(n.a)))) setAttr(n, 'descr', run.alt || undefined)
                if (run.rid && l === 'blip') {
                  const m = /([\w.-]+):embed="/.exec(n.a)
                  setAttr(n, m ? `${m[1]}:embed` : 'r:embed', run.rid)
                }
              })
              body += serialize(r)
            } else body += serialize(src)
            map.push({ src: run.src })
          }
        }
        continue
      }
      if (run.lk !== openLink) {
        closeLink()
        if (run.lk) {
          const attr = run.lk.startsWith('r:') ? `r:id="${run.lk.slice(2)}"` : `${W}anchor="${escText(run.lk.slice(2))}"`
          body += `<${W}hyperlink ${attr}>`
          openLink = run.lk
        }
      }
      let rPr = baseRPr(run.src)
      if (run.set && Object.keys(run.set).length) applyRunSet((rPr ??= el(W + 'rPr')), run.set, W)
      if (!run.text) continue
      body += runXml(rPr, textContentXml(run.text, W))
      map.push({ src: run.src, d: run.set })
    }
  })
  closeLink()

  // bookmarks / comment ranges / comment references of the source paragraphs
  let starts = ''
  let ends = ''
  const orig = spec.orig ?? (spec.pid === spec.from ? [spec.from] : [])
  for (const pid of orig) {
    const op = origPara(doc, pid)
    if (!op) continue
    for (const c of kidsOf(op)) {
      const l = local(c.q)
      if (START_MARKS.has(l)) starts += serialize(c)
      else if (END_MARKS.has(l)) ends += serialize(c)
    }
    for (const r of collectRunsRaw(op)) {
      if (kidOf(r, 'commentReference')) {
        ends += serialize(r)
        map.push({})
      }
    }
  }

  let attrs = tpl?.a ?? ''
  if (spec.pid !== spec.from) attrs = attrs.replace(/\s+w14:(paraId|textId)="[^"]*"/g, '')
  const ppr = pPr.c.length ? serialize(pPr) : ''
  // comment references were appended after `body`; keep `map` order aligned with emission order
  return { xml: `<${pq}${attrs}>${ppr}${starts}${body}${ends}</${pq}>`, map }
}

// ---------------------------------------------------------------- edit operations

function groupFor(doc: DocxDocument, key: string): number[] {
  let g = doc.groups.get(key)
  if (!g) {
    const p = doc.ordPid.get(key)
    g = p !== undefined ? [p] : []
    doc.groups.set(key, g)
  }
  return g
}

function detach(doc: DocxDocument, pid: number) {
  const inf = doc.info.get(pid)
  if (!inf) return
  const g = groupFor(doc, `${inf.item}:${inf.ord}`)
  const i = g.indexOf(pid)
  if (i >= 0) g.splice(i, 1)
}

export async function applyEdit(doc: DocxDocument, op: EditOp): Promise<EditResult> {
  if (!doc.ready) throw new Error('Document is still loading')
  const anchor = doc.info.get(op.anchor)
  if (!anchor) throw new Error(`Unknown paragraph ${op.anchor}`)
  doc.markEdited()
  // list numbering only needs recomputing when list paragraphs or paragraph properties are involved
  const involved = [op.anchor, ...op.list.map((x) => x.pid), ...(op.remove ?? [])]
  const listy = op.list.some((x) => x.pd) || involved.some((pid) => doc.info.get(pid)?.num)
  const item = anchor.item
  const ord = anchor.ord
  const listPids = op.list.map((s) => s.pid)
  for (const pid of [...listPids, ...(op.remove ?? [])]) if (pid !== op.anchor) detach(doc, pid)
  const group = groupFor(doc, `${item}:${ord}`)
  const idx = group.indexOf(op.anchor)
  if (idx < 0) throw new Error(`Paragraph ${op.anchor} is not in the document`)
  group.splice(idx, 1, ...listPids)
  for (const pid of [op.anchor, ...(op.remove ?? [])]) if (!listPids.includes(pid)) doc.genXml.delete(pid)

  const blocks: Record<number, Block[]> = {}
  for (const spec of op.list) {
    const tplInfo = doc.info.get(spec.from) ?? anchor
    const prev = doc.info.get(spec.pid)
    let xml: string
    let map: GenRun[]
    const orig = spec.pristine && spec.pid === spec.from ? origPara(doc, spec.pid) : undefined
    if (orig) {
      xml = serialize(orig)
      map = collectRunsRaw(orig).map((_, k) => ({ src: [spec.pid, k] as SrcRef }))
    } else ({ xml, map } = await genParagraph(doc, spec))
    doc.genXml.set(spec.pid, xml)
    const { blocks: bl, rec } = doc.reconvert(spec.pid, xml, map, tplInfo.layer, prev?.marker)
    doc.info.set(spec.pid, { ...rec, item, ord, marker: prev?.marker, lst: prev?.lst })
    if (op.want) blocks[spec.pid] = bl
  }
  const listy2 = listy || op.list.some((x) => doc.info.get(x.pid)?.num)
  const changed = listy2 ? doc.recomputeMarkers() : new Map()
  const markers: EditResult['markers'] = {}
  for (const [pid, inf] of changed) markers[pid] = inf.marker ? { ...inf.marker, lst: inf.lst } : null
  // returned blocks carry the final markers
  for (const [pid, bl] of Object.entries(blocks)) {
    const inf = doc.info.get(Number(pid))
    bl.filter((b): b is Paragraph => b.t === 'p').forEach((p, i) => {
      delete p.marker
      delete p.lst
      if (i === 0 && inf?.marker) {
        p.marker = inf.marker
        p.lst = inf.lst
      }
    })
  }
  return { blocks, markers }
}

let drawingId = 7000

function imageRunXml(W: string, rid: string, w: number, h: number, name = 'Picture', alt?: string): string {
  const cx = Math.round(w * 12700)
  const cy = Math.round(h * 12700)
  const id = ++drawingId
  const A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
  const PIC = 'http://schemas.openxmlformats.org/drawingml/2006/picture'
  const WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
  const n = escText(name)
  const descr = alt ? ` descr="${escText(alt).replace(/"/g, '&quot;')}"` : ''
  return (
    `<${W}r><${W}drawing><wp:inline xmlns:wp="${WP}" distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/>` +
    `<wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="${id}" name="${n} ${id}"${descr}/><wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="${A}" noChangeAspect="1"/></wp:cNvGraphicFramePr>` +
    `<a:graphic xmlns:a="${A}"><a:graphicData uri="${PIC}"><pic:pic xmlns:pic="${PIC}"><pic:nvPicPr><pic:cNvPr id="${id}" name="${n}"${descr}/><pic:cNvPicPr/></pic:nvPicPr>` +
    `<pic:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></${W}drawing></${W}r>`
  )
}

const IMAGE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
const HYPERLINK_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'
const IMG_TYPES: Record<string, string> = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', svg: 'image/svg+xml', webp: 'image/webp' }

/** Adds an image part to the document; returns its relationship id and part path. */
export async function addImage(doc: DocxDocument, bytes: Uint8Array, ext: string): Promise<{ rid: string; path: string }> {
  ext = ext.toLowerCase().replace('jpeg', 'jpg')
  if (!IMG_TYPES[ext]) throw new Error(`Unsupported image type .${ext}`)
  const path = doc.pkg.freeName('word/media', 'image', ext)
  doc.pkg.setPart(path, bytes)
  await doc.pkg.ensureDefaultType(ext, IMG_TYPES[ext])
  const rid = await doc.pkg.addRel(doc.mainPath, IMAGE_REL, path)
  doc.rels = await doc.pkg.rels(doc.mainPath)
  doc.markEdited()
  return { rid, path }
}

/** External hyperlink relationship; returns the run link key ('r:rIdN'). */
export async function addLink(doc: DocxDocument, url: string): Promise<string> {
  const rid = await doc.pkg.addRel(doc.mainPath, HYPERLINK_REL, url, true)
  doc.rels = await doc.pkg.rels(doc.mainPath)
  return 'r:' + rid
}

// ---------------------------------------------------------------- tables

const TCPR_ORDER = ['cnfStyle', 'tcW', 'gridSpan', 'hMerge', 'vMerge', 'tcBorders', 'shd', 'noWrap', 'tcMar', 'textDirection', 'tcFitText', 'vAlign', 'hideMark']

function rawRows(tbl: RNode): RNode[] {
  const out: RNode[] = []
  const rec = (n: RNode) => {
    for (const c of kidsOf(n)) {
      const l = local(c.q)
      if (l === 'tr') out.push(c)
      else if (l === 'sdt' || l === 'sdtContent' || l === 'customXml') rec(c)
    }
  }
  rec(tbl)
  return out
}

function rawCells(tr: RNode): RNode[] {
  const out: RNode[] = []
  const rec = (n: RNode) => {
    for (const c of kidsOf(n)) {
      const l = local(c.q)
      if (l === 'tc') out.push(c)
      else if (l === 'sdt' || l === 'sdtContent' || l === 'customXml') rec(c)
    }
  }
  rec(tr)
  return out
}

/** Regenerates a whole table from the editor model, reusing original row/cell properties. */
export async function genTable(doc: DocxDocument, spec: TableSpec): Promise<string> {
  let orig: RNode | undefined
  if (spec.item !== undefined) {
    const nodes = parseRaw(doc.items[spec.item] ?? '')
    const top = rootEl(nodes)
    if (top) {
      if (local(top.q) === 'tbl') orig = top
      else walkOoxml(top, (n) => {
        if (!orig && local(n.q) === 'tbl') orig = n
        return !!orig
      })
    }
  }
  const W = orig ? prefixOf(orig.q) || 'w:' : 'w:'
  const val = W + 'val'
  const rows = orig ? rawRows(orig) : []
  let tblPr = orig ? kidOf(orig, 'tblPr') : undefined
  if (!tblPr) {
    const style = doc.styles.has('TableGrid') ? `<${W}tblStyle ${val}="TableGrid"/>` : ''
    const borders = style
      ? ''
      : `<${W}tblBorders>${['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map((b) => `<${W}${b} ${val}="single" ${W}sz="4" ${W}space="0" ${W}color="auto"/>`).join('')}</${W}tblBorders>`
    tblPr = parseRaw(`<${W}tblPr>${style}<${W}tblW ${W}w="0" ${W}type="auto"/>${borders}<${W}tblLook ${val}="04A0" ${W}firstRow="1" ${W}lastRow="0" ${W}firstColumn="1" ${W}lastColumn="0" ${W}noHBand="0" ${W}noVBand="1"/></${W}tblPr>`)[0] as RNode
  }
  const grid = `<${W}tblGrid>${spec.cols.map((w) => `<${W}gridCol ${W}w="${Math.round(w * 20)}"/>`).join('')}</${W}tblGrid>`
  const firstTcPr = rows.length ? kidOf(rawCells(rows[0])[0], 'tcPr') : undefined
  const nCols = spec.cols.length
  const cont: ({ left: number; span: number; tcPr?: RNode } | undefined)[] = []
  let out = `<${W}tbl${orig?.a ?? ''}>${serialize(tblPr)}${grid}`
  const tcPrFor = (src: [number, number] | undefined, g: number, span: number, vm?: 'restart' | 'continue', shd?: string) => {
    const tpl = src ? kidOf(rawCells(rows[src[0]] ?? ({ c: [] } as unknown as RNode))[src[1]], 'tcPr') : undefined
    const pr = clone(tpl ?? firstTcPr ?? el(W + 'tcPr'))
    pr.q = W + 'tcPr'
    removeKids(pr, (l) => l === 'gridSpan' || l === 'vMerge' || l === 'hMerge')
    const w = spec.cols.slice(g, g + span).reduce((a, b) => a + b, 0)
    setOrdered(pr, el(W + 'tcW', { [W + 'w']: String(Math.round(w * 20)), [W + 'type']: 'dxa' }), 'tcW', TCPR_ORDER)
    if (span > 1) setOrdered(pr, el(W + 'gridSpan', { [val]: String(span) }), 'gridSpan', TCPR_ORDER)
    if (vm) setOrdered(pr, el(W + 'vMerge', vm === 'restart' ? { [val]: 'restart' } : {}), 'vMerge', TCPR_ORDER)
    if (shd !== undefined) setOrdered(pr, shd ? el(W + 'shd', { [val]: 'clear', [W + 'color']: 'auto', [W + 'fill']: shd.replace('#', '').toUpperCase() }) : undefined, 'shd', TCPR_ORDER)
    return serialize(pr)
  }
  for (const row of spec.rows) {
    const trSrc = row.src !== undefined ? rows[row.src] : undefined
    const trPr = kidOf(trSrc, 'trPr')
    out += `<${W}tr${trSrc?.a ?? ''}>${trPr ? serialize(trPr) : ''}`
    let g = 0
    let k = 0
    const cells: CellSpec[] = row.cells
    while (g < nCols && (k < cells.length || cont[g])) {
      const c = cont[g]
      if (c && c.left > 0) {
        out += `<${W}tc>${c.tcPr ? '' : ''}${tcPrFor(undefined, g, c.span, 'continue')}<${W}p/></${W}tc>`
        c.left--
        if (!c.left) cont[g] = undefined
        g += c.span
        continue
      }
      const cell = cells[k++]
      if (!cell) break
      const span = Math.max(1, Math.min(cell.colSpan, nCols - g))
      let paras = ''
      for (const p of cell.paras) {
        const orig = p.pristine && p.pid === p.from ? origPara(doc, p.pid) : undefined
        paras += orig ? serialize(orig) : (await genParagraph(doc, p)).xml
      }
      out += `<${W}tc>${tcPrFor(cell.src, g, span, cell.rowSpan > 1 ? 'restart' : undefined, cell.shd)}${paras || `<${W}p/>`}</${W}tc>`
      if (cell.rowSpan > 1) cont[g] = { left: cell.rowSpan - 1, span }
      g += span
    }
    out += `</${W}tr>`
  }
  return out + `</${W}tbl>`
}

export async function applyTable(doc: DocxDocument, op: TableOp): Promise<void> {
  if (!doc.ready) throw new Error('Document is still loading')
  doc.markEdited()
  const xml = op.spec ? await genTable(doc, op.spec) : ''
  if (op.item !== undefined) {
    doc.itemXml.set(op.item, xml)
    return
  }
  // new table living inside a paragraph group
  let key: string | undefined
  for (const [k, g] of doc.groups) if (g.includes(op.tid)) key = k
  if (!op.spec) {
    if (key) doc.groups.get(key)!.splice(doc.groups.get(key)!.indexOf(op.tid), 1)
    doc.genXml.delete(op.tid)
    return
  }
  if (!key) {
    const anchor = op.anchor !== undefined ? doc.info.get(op.anchor) : undefined
    if (!anchor || op.anchor === undefined) throw new Error('Table anchor paragraph not found')
    const g = groupFor(doc, `${anchor.item}:${anchor.ord}`)
    g.splice(g.indexOf(op.anchor) + 1, 0, op.tid)
  }
  doc.genXml.set(op.tid, xml)
}

export async function saveDocx(doc: DocxDocument): Promise<Blob> {
  if (!doc.ready) throw new Error('Document is still loading')
  if (doc.groups.size || doc.itemXml.size) {
    const touched = new Set<number>()
    for (const key of doc.groups.keys()) touched.add(Number(key.split(':')[0]))
    const out: string[] = [doc.prefix]
    for (let i = 0; i < doc.items.length; i++) {
      const whole = doc.itemXml.get(i)
      if (whole !== undefined) {
        // keep whitespace that followed the item
        out.push(whole + (/\s*$/.exec(doc.items[i])?.[0] ?? ''))
        continue
      }
      if (!touched.has(i)) {
        out.push(doc.items[i])
        continue
      }
      const nodes = parseRaw(doc.items[i])
      const paras = collectParasRaw(nodes)
      for (let o = 0; o < paras.length; o++) {
        const g = doc.groups.get(`${i}:${o}`)
        if (!g) continue
        const loc = paras[o]
        loc.list[loc.index] = g.map((pid) => doc.genXml.get(pid) ?? (origPara(doc, pid) ? serialize(origPara(doc, pid)!) : '')).join('')
      }
      out.push(serialize(nodes))
    }
    out.push(doc.suffix)
    doc.pkg.setPart(doc.mainPath, out.join(''))
  }
  return doc.pkg.save(DOCX_MIME)
}
