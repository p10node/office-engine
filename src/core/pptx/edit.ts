/**
 * PPTX structural editing: shape geometry/fill/order, inserts, and slide list operations.
 * Works on the raw slide XML so everything not touched is preserved.
 */
import { relOfType, relsPathFor, relsOfType } from '../opc'
import {
  clone, el, getAttr, relIdAttrs, isNode, kidOf, kidsOf, local, parseRaw, prefixOf, removeKids, rootEl, serialize, setAttr, setOrdered, walkOoxml,
  escText, type RChild, type RNode,
} from '../raw'
import { EMU_PER_PX } from '../units'
import type { PptxPresentation } from './parser'
import type { Box, PresentationMeta, ShapeOp, SlideModel, SlideOp } from './types'

const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const NS_P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
const SLIDE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide'
const LAYOUT_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout'
const IMAGE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
const SLIDE_CT = 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml'
const IMG_TYPES: Record<string, string> = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', svg: 'image/svg+xml', webp: 'image/webp' }
const SPPR_ORDER = ['xfrm', 'custGeom', 'prstGeom', 'noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill', 'ln', 'effectLst', 'effectDag', 'scene3d', 'sp3d', 'extLst']
const FILLS = new Set(['noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill'])

const emu = (px: number) => String(Math.round(px * EMU_PER_PX))

function spTreeOf(nodes: RChild[]): RNode {
  const root = rootEl(nodes)!
  const tree = kidOf(kidOf(root, 'cSld'), 'spTree')
  if (!tree) throw new Error('Slide has no shape tree')
  return tree
}

function idOf(n: RNode): number | undefined {
  const nv = kidsOf(n).find((c) => /^nv.*Pr$/.test(local(c.q)))
  const id = nv ? getAttr(kidOf(nv, 'cNvPr') ?? nv, 'id') : undefined
  return id !== undefined ? Number(id) : undefined
}

function maxId(nodes: RChild[]): number {
  let max = 1
  for (const n of nodes) {
    if (!isNode(n)) continue
    walkOoxml(n, (c) => {
      if (local(c.q) === 'cNvPr') max = Math.max(max, Number(getAttr(c, 'id') ?? 0))
    })
  }
  return max
}

function aPrefix(nodes: RChild[]): string {
  let a = ''
  for (const n of nodes) {
    if (!isNode(n)) continue
    walkOoxml(n, (c) => {
      if (!a && /:(off|ext|t|p|bodyPr)$/.test(c.q)) a = prefixOf(c.q)
      return !!a
    })
  }
  return a || 'a:'
}

function setXfrm(node: RNode, box: Box, A: string, P: string) {
  const l = local(node.q)
  let host: RNode
  if (l === 'graphicFrame') host = node
  else {
    const prName = l === 'grpSp' ? 'grpSpPr' : 'spPr'
    let pr = kidOf(node, prName)
    if (!pr) {
      pr = el(P + prName)
      node.c.splice(1, 0, pr)
    }
    host = pr
  }
  let x = kidOf(host, 'xfrm')
  if (!x) {
    x = el((l === 'graphicFrame' ? P : A) + 'xfrm')
    if (l === 'graphicFrame') host.c.splice(1, 0, x)
    else setOrdered(host, x, 'xfrm', SPPR_ORDER)
  }
  const put = (name: string, attrs: Record<string, string>) => {
    let n = kidOf(x!, name)
    if (!n) {
      n = el(A + name)
      x!.c.splice(name === 'off' ? 0 : name === 'ext' ? 1 : x!.c.length, 0, n)
    }
    for (const [k, v] of Object.entries(attrs)) setAttr(n, k, v)
  }
  put('off', { x: emu(box.x), y: emu(box.y) })
  put('ext', { cx: emu(Math.max(1, box.w)), cy: emu(Math.max(1, box.h)) })
  setAttr(x, 'rot', box.rot ? String(Math.round((((box.rot % 360) + 360) % 360) * 60000)) : undefined)
  if (box.flipH !== undefined) setAttr(x, 'flipH', box.flipH ? '1' : undefined)
  if (box.flipV !== undefined) setAttr(x, 'flipV', box.flipV ? '1' : undefined)
  x.self = false
  host.self = false
}

function fillXml(A: string, color: string | null) {
  return color ? el(A + 'solidFill', {}, [el(A + 'srgbClr', { val: color.replace('#', '').toUpperCase() })]) : el(A + 'noFill')
}

function shapeXml(kind: string, id: number, box: Box, A: string, P: string, text?: string, rid?: string): string {
  const off = `<${A}off x="${emu(box.x)}" y="${emu(box.y)}"/><${A}ext cx="${emu(box.w)}" cy="${emu(box.h)}"/>`
  if (kind === 'image') {
    return (
      `<${P}pic><${P}nvPicPr><${P}cNvPr id="${id}" name="Picture ${id}"/><${P}cNvPicPr><${A}picLocks noChangeAspect="1"/></${P}cNvPicPr><${P}nvPr/></${P}nvPicPr>` +
      `<${P}blipFill><${A}blip r:embed="${rid}"/><${A}stretch><${A}fillRect/></${A}stretch></${P}blipFill>` +
      `<${P}spPr><${A}xfrm>${off}</${A}xfrm><${A}prstGeom prst="rect"><${A}avLst/></${A}prstGeom></${P}spPr></${P}pic>`
    )
  }
  const para = (t: string) => `<${A}p><${A}r><${A}rPr lang="en-US" dirty="0"/><${A}t>${escText(t)}</${A}t></${A}r></${A}p>`
  if (kind === 'text') {
    return (
      `<${P}sp><${P}nvSpPr><${P}cNvPr id="${id}" name="TextBox ${id}"/><${P}cNvSpPr txBox="1"/><${P}nvPr/></${P}nvSpPr>` +
      `<${P}spPr><${A}xfrm>${off}</${A}xfrm><${A}prstGeom prst="rect"><${A}avLst/></${A}prstGeom><${A}noFill/></${P}spPr>` +
      `<${P}txBody><${A}bodyPr wrap="square" rtlCol="0"><${A}spAutoFit/></${A}bodyPr><${A}lstStyle/>${para(text ?? 'Text')}</${P}txBody></${P}sp>`
    )
  }
  if (kind === 'line') {
    return (
      `<${P}cxnSp><${P}nvCxnSpPr><${P}cNvPr id="${id}" name="Connector ${id}"/><${P}cNvCxnSpPr/><${P}nvPr/></${P}nvCxnSpPr>` +
      `<${P}spPr><${A}xfrm>${off}</${A}xfrm><${A}prstGeom prst="line"><${A}avLst/></${A}prstGeom><${A}ln w="19050"><${A}solidFill><${A}schemeClr val="tx1"/></${A}solidFill></${A}ln></${P}spPr></${P}cxnSp>`
    )
  }
  return (
    `<${P}sp><${P}nvSpPr><${P}cNvPr id="${id}" name="Shape ${id}"/><${P}cNvSpPr/><${P}nvPr/></${P}nvSpPr>` +
    `<${P}spPr><${A}xfrm>${off}</${A}xfrm><${A}prstGeom prst="${kind}"><${A}avLst/></${A}prstGeom></${P}spPr>` +
    `<${P}style><${A}lnRef idx="2"><${A}schemeClr val="accent1"><${A}shade val="50000"/></${A}schemeClr></${A}lnRef><${A}fillRef idx="1"><${A}schemeClr val="accent1"/></${A}fillRef>` +
    `<${A}effectRef idx="0"><${A}schemeClr val="accent1"/></${A}effectRef><${A}fontRef idx="minor"><${A}schemeClr val="lt1"/></${A}fontRef></${P}style>` +
    `<${P}txBody><${A}bodyPr rtlCol="0" anchor="ctr"/><${A}lstStyle/><${A}p><${A}pPr algn="ctr"/>${text ? `<${A}r><${A}rPr lang="en-US" dirty="0"/><${A}t>${escText(text)}</${A}t></${A}r>` : ''}<${A}endParaRPr lang="en-US" dirty="0"/></${A}p></${P}txBody></${P}sp>`
  )
}

/** solid slide background (or back to the layout's when null) */
function setBackground(nodes: RChild[], A: string, P: string, color: string | null) {
  const cSld = kidOf(rootEl(nodes), 'cSld')
  if (!cSld) return
  cSld.c = cSld.c.filter((c) => !(isNode(c) && local(c.q) === 'bg'))
  if (!color) return
  const bg = parseRaw(`<${P}bg><${P}bgPr><${A}solidFill><${A}srgbClr val="${color.replace('#', '').toUpperCase()}"/></${A}solidFill><${A}effectLst/></${P}bgPr></${P}bg>`)[0] as RNode
  cSld.c.unshift(bg)
  cSld.self = false
}

async function editSlide(pres: PptxPresentation, index: number, fn: (nodes: RChild[], tree: RNode, A: string, P: string) => void | Promise<void>): Promise<SlideModel> {
  const info = pres.meta.slides[index]
  if (!info) throw new Error(`No slide #${index}`)
  const before = await pres.pkg.text(info.path)
  const nodes = parseRaw(before)
  const tree = spTreeOf(nodes)
  await fn(nodes, tree, aPrefix(nodes), prefixOf(tree.q))
  const after = serialize(nodes)
  pres.pkg.setPart(info.path, after)
  pres.touched = true
  if (after !== before) {
    pres.undoStack.push({ path: info.path, before, after })
    if (pres.undoStack.length > 200) pres.undoStack.shift()
    pres.redoStack = []
  }
  return pres.slide(index)
}

/** undoes / redoes the last shape edit; returns the slide index and its new model */
export async function undoShapes(pres: PptxPresentation, redo = false): Promise<{ index: number; model: SlideModel } | null> {
  const e = (redo ? pres.redoStack : pres.undoStack).pop()
  if (!e) return null
  ;(redo ? pres.undoStack : pres.redoStack).push(e)
  pres.pkg.setPart(e.path, redo ? e.after : e.before)
  const index = pres.meta.slides.findIndex((s) => s.path === e.path)
  if (index < 0) return null
  return { index, model: await pres.slide(index) }
}

export async function addSlideImage(pres: PptxPresentation, index: number, bytes: Uint8Array, ext: string): Promise<string> {
  ext = ext.toLowerCase().replace('jpeg', 'jpg')
  if (!IMG_TYPES[ext]) throw new Error(`Unsupported image type .${ext}`)
  const info = pres.meta.slides[index]
  const path = pres.pkg.freeName('ppt/media', 'image', ext)
  pres.pkg.setPart(path, bytes)
  await pres.pkg.ensureDefaultType(ext, IMG_TYPES[ext])
  return pres.pkg.addRel(info.path, IMAGE_REL, path)
}

export function shapeOps(pres: PptxPresentation, index: number, ops: ShapeOp[]): Promise<SlideModel> {
  return editSlide(pres, index, (nodes, tree, A, P) => {
    const find = (top: number) => {
      const i = tree.c.findIndex((c) => isNode(c) && idOf(c) === top)
      return i >= 0 ? { node: tree.c[i] as RNode, i } : undefined
    }
    for (const op of ops) {
      if (op.op === 'insert') {
        const id = maxId(nodes) + 1
        const xml = shapeXml(op.kind, id, op.box, A, P, op.text, op.rid)
        tree.c.push(parseRaw(xml)[0])
        continue
      }
      if (op.op === 'bg') {
        setBackground(nodes, A, P, op.color)
        continue
      }
      if (op.op === 'anchor') {
        let body: RNode | undefined
        walkOoxml(tree, (c) => {
          if (body) return true
          if (local(c.q) === 'sp' && Number(getAttr(kidOf(kidOf(c, 'nvSpPr'), 'cNvPr') ?? c, 'id')) === op.sid) body = kidOf(kidOf(c, 'txBody'), 'bodyPr')
        })
        if (body) setAttr(body, 'anchor', op.anchor)
        continue
      }
      const hit = find(op.top)
      if (!hit) continue
      const { node, i } = hit
      switch (op.op) {
        case 'xfrm':
          setXfrm(node, op.box, A, P)
          break
        case 'delete':
          tree.c.splice(i, 1)
          break
        case 'order': {
          const shapes = () => tree.c.map((c, k) => (isNode(c) && !/^(nvGrpSpPr|grpSpPr|extLst)$/.test(local(c.q)) ? k : -1)).filter((k) => k >= 0)
          const before = shapes()
          const pos = before.indexOf(i)
          tree.c.splice(i, 1)
          const rest = shapes()
          if (op.to === 'front') tree.c.splice(rest.length ? rest[rest.length - 1] + 1 : tree.c.length, 0, node)
          else if (op.to === 'back') tree.c.splice(rest.length ? rest[0] : tree.c.length, 0, node)
          else {
            // one step: swap with the neighbour shape
            const to = op.to === 'forward' ? Math.min(rest.length, pos + 1) : Math.max(0, pos - 1)
            tree.c.splice(to < rest.length ? rest[to] : rest.length ? rest[rest.length - 1] + 1 : tree.c.length, 0, node)
          }
          break
        }
        case 'duplicate': {
          const copy = clone(node)
          let id = maxId(nodes)
          walkOoxml(copy, (c) => {
            if (local(c.q) === 'cNvPr') setAttr(c, 'id', String(++id))
          })
          const off = (() => {
            let o: RNode | undefined
            walkOoxml(copy, (c) => {
              if (!o && local(c.q) === 'off') o = c
              return !!o
            })
            return o
          })()
          if (off) {
            setAttr(off, 'x', String(Number(getAttr(off, 'x') ?? 0) + Math.round((op.dx ?? 16) * EMU_PER_PX)))
            setAttr(off, 'y', String(Number(getAttr(off, 'y') ?? 0) + Math.round((op.dy ?? 16) * EMU_PER_PX)))
          }
          tree.c.splice(i + 1, 0, copy)
          break
        }
        case 'fill':
        case 'line': {
          const l = local(node.q)
          if (l !== 'sp' && l !== 'cxnSp' && l !== 'pic') break
          let pr = kidOf(node, 'spPr')
          if (!pr) {
            pr = el(P + 'spPr')
            node.c.splice(1, 0, pr)
          }
          if (op.op === 'fill') {
            removeKids(pr, (k) => FILLS.has(k))
            setOrdered(pr, fillXml(A, op.color), op.color ? 'solidFill' : 'noFill', SPPR_ORDER)
          } else {
            const ln = clone(kidOf(pr, 'ln') ?? el(A + 'ln'))
            if ('color' in op) {
              removeKids(ln, (k) => FILLS.has(k))
              ln.c.unshift(fillXml(A, op.color ?? null))
            } else if (!ln.c.some((c) => isNode(c) && FILLS.has(local(c.q)))) {
              // weight / dash on a shape whose outline comes from its style: make it visible
              ln.c.unshift(el(A + 'solidFill', {}, [el(A + 'schemeClr', { val: 'tx1' })]))
            }
            if (op.dash !== undefined) {
              removeKids(ln, (k) => k === 'prstDash' || k === 'custDash')
              const at = ln.c.findIndex((c) => isNode(c) && !FILLS.has(local(c.q)))
              ln.c.splice(at < 0 ? ln.c.length : at, 0, el(A + 'prstDash', { val: op.dash }))
            }
            ln.self = false
            if (op.w !== undefined) setAttr(ln, 'w', String(Math.round(op.w * 12700)))
            setOrdered(pr, ln, 'ln', SPPR_ORDER)
          }
          pr.self = false
          break
        }
      }
    }
  })
}

// ---------------------------------------------------------------- slides

async function presNodes(pres: PptxPresentation) {
  const nodes = parseRaw(await pres.pkg.text(pres.mainPath))
  const root = rootEl(nodes)!
  const P = prefixOf(root.q)
  let lst = kidOf(root, 'sldIdLst')
  if (!lst) {
    lst = el(P + 'sldIdLst')
    const at = root.c.findIndex((c) => isNode(c) && local(c.q) === 'sldMasterIdLst')
    root.c.splice(at + 1, 0, lst)
  }
  return { nodes, root, P, lst }
}

/** Re-reads the slide list from presentation.xml into `pres.meta`. */
async function refreshMeta(pres: PptxPresentation): Promise<PresentationMeta> {
  const rels = await pres.pkg.rels(pres.mainPath)
  const { lst } = await presNodes(pres)
  const slides: PresentationMeta['slides'] = []
  for (const s of kidsOf(lst, 'sldId')) {
    const rel = rels.get(getAttr(s, 'r:id') ?? getAttr(s, 'id', true) ?? '')
    if (rel && pres.pkg.has(rel.target)) slides.push({ index: slides.length, path: rel.target, hidden: false })
  }
  for (const s of slides) {
    const head = (await pres.pkg.text(s.path)).slice(0, 600)
    s.hidden = /<[\w:]*sld\b[^>]*\sshow="0"/.test(head)
  }
  pres.meta.slides = slides
  pres.meta.count = slides.length
  return pres.meta
}

async function pickLayout(pres: PptxPresentation, refSlide: string | undefined, want: string): Promise<string | undefined> {
  const ref = refSlide ? relOfType(await pres.pkg.rels(refSlide), 'slideLayout')?.target : undefined
  let masterPath = ref ? relOfType(await pres.pkg.rels(ref), 'slideMaster')?.target : undefined
  if (!masterPath) masterPath = relOfType(await pres.pkg.rels(pres.mainPath), 'slideMaster')?.target
  if (!masterPath) return ref
  const layouts = relsOfType(await pres.pkg.rels(masterPath), 'slideLayout').map((r) => r.target)
  const types = await Promise.all(layouts.map(async (l) => /<[\w:]*sldLayout\b[^>]*\stype="([^"]+)"/.exec((await pres.pkg.text(l)).slice(0, 800))?.[1]))
  const want2 = want === 'title' ? ['title'] : want === 'blank' ? ['blank'] : ['obj', 'txAndObj', 'tx']
  const i = types.findIndex((t) => t && want2.includes(t))
  return i >= 0 ? layouts[i] : ref ?? layouts[0]
}

async function newSlideXml(pres: PptxPresentation, layoutPath: string | undefined): Promise<string> {
  let shapes = ''
  if (layoutPath) {
    const nodes = parseRaw(await pres.pkg.text(layoutPath))
    const tree = kidOf(kidOf(rootEl(nodes), 'cSld'), 'spTree')
    let id = 2
    for (const sp of kidsOf(tree, 'sp')) {
      const ph = kidOf(kidOf(kidOf(sp, 'nvSpPr'), 'nvPr'), 'ph')
      if (!ph) continue
      const type = getAttr(ph, 'type')
      if (type && ['dt', 'ftr', 'sldNum'].includes(type)) continue
      const name = getAttr(kidOf(kidOf(sp, 'nvSpPr'), 'cNvPr') ?? sp, 'name') ?? `Placeholder ${id}`
      const phXml = `<p:ph${type ? ` type="${type}"` : ''}${getAttr(ph, 'idx') ? ` idx="${getAttr(ph, 'idx')}"` : ''}/>`
      shapes +=
        `<p:sp><p:nvSpPr><p:cNvPr id="${id++}" name="${escText(name)}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr>${phXml}</p:nvPr></p:nvSpPr>` +
        `<p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US" dirty="0"/></a:p></p:txBody></p:sp>`
    }
  }
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<p:sld xmlns:a="${NS_A}" xmlns:r="${NS_R}" xmlns:p="${NS_P}"><p:cSld><p:spTree>` +
    `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>` +
    `${shapes}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`
  )
}

async function insertSlidePart(pres: PptxPresentation, after: number, xml: string, rels: string): Promise<void> {
  const path = pres.pkg.freeName('ppt/slides', 'slide', 'xml')
  pres.pkg.setPart(path, xml)
  pres.pkg.setPart(relsPathFor(path), rels)
  await pres.pkg.setContentType(path, SLIDE_CT)
  const rid = await pres.pkg.addRel(pres.mainPath, SLIDE_REL, path)
  const { nodes, root, P, lst } = await presNodes(pres)
  const ids = kidsOf(lst, 'sldId').map((s) => Number(getAttr(s, 'id') ?? 0))
  const id = Math.max(255, ...ids) + 1
  const entry = el(P + 'sldId', { id: String(id), ...relIdAttrs(root, 'id', rid) })
  const items = kidsOf(lst, 'sldId')
  const at = after < 0 ? 0 : lst.c.indexOf(items[Math.min(after, items.length - 1)]) + 1
  lst.c.splice(at, 0, entry)
  lst.self = false
  pres.pkg.setPart(pres.mainPath, serialize(nodes))
}

export async function slideOp(pres: PptxPresentation, op: SlideOp): Promise<PresentationMeta> {
  const slides = pres.meta.slides
  pres.touched = true
  switch (op.op) {
    case 'add': {
      const ref = slides[Math.max(0, Math.min(op.after, slides.length - 1))]?.path
      const layout = await pickLayout(pres, ref, op.layout ?? 'content')
      const xml = await newSlideXml(pres, layout)
      const rels = layout
        ? `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${LAYOUT_REL}" Target="../slideLayouts/${layout.split('/').pop()}"/></Relationships>`
        : '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>'
      await insertSlidePart(pres, op.after, xml, rels)
      break
    }
    case 'duplicate': {
      const src = slides[op.index]
      if (!src) break
      const xml = await pres.pkg.text(src.path)
      const rp = relsPathFor(src.path)
      let rels = pres.pkg.has(rp) ? await pres.pkg.text(rp) : '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>'
      // notes belong to the original slide only
      const rn = parseRaw(rels)
      const rroot = rootEl(rn)!
      rroot.c = rroot.c.filter((c) => !(isNode(c) && /notesSlide$/.test(getAttr(c, 'Type') ?? '')))
      rels = serialize(rn)
      await insertSlidePart(pres, op.index, xml, rels)
      break
    }
    case 'delete': {
      const s = slides[op.index]
      if (!s || slides.length <= 1) break
      const { nodes, lst } = await presNodes(pres)
      const rels = await pres.pkg.rels(pres.mainPath)
      const entry = kidsOf(lst, 'sldId').find((x) => rels.get(getAttr(x, 'r:id') ?? '')?.target === s.path)
      if (entry) lst.c.splice(lst.c.indexOf(entry), 1)
      pres.pkg.setPart(pres.mainPath, serialize(nodes))
      const notes = relOfType(await pres.pkg.rels(s.path), 'notesSlide')?.target
      await pres.pkg.removeRelsTo(pres.mainPath, s.path)
      for (const part of [s.path, notes].filter(Boolean) as string[]) {
        pres.pkg.setPart(part, null)
        if (pres.pkg.has(relsPathFor(part))) pres.pkg.setPart(relsPathFor(part), null)
        await pres.pkg.setContentType(part, null)
      }
      break
    }
    case 'move': {
      const { nodes, lst } = await presNodes(pres)
      const items = kidsOf(lst, 'sldId')
      const item = items[op.from]
      if (!item || op.to < 0 || op.to >= items.length) break
      lst.c.splice(lst.c.indexOf(item), 1)
      const rest = kidsOf(lst, 'sldId')
      const at = op.to >= rest.length ? lst.c.length : lst.c.indexOf(rest[op.to])
      lst.c.splice(at, 0, item)
      pres.pkg.setPart(pres.mainPath, serialize(nodes))
      break
    }
    case 'hide': {
      const s = slides[op.index]
      if (!s) break
      const nodes = parseRaw(await pres.pkg.text(s.path))
      setAttr(rootEl(nodes)!, 'show', op.hidden ? '0' : undefined)
      pres.pkg.setPart(s.path, serialize(nodes))
      break
    }
  }
  return refreshMeta(pres)
}
