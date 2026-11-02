import { ZipArchive } from './zip'
import { kids, parseXml, type XEl } from './xml'

export type DocKind = 'docx' | 'xlsx' | 'pptx'

export interface Rel {
  id: string
  type: string
  target: string
  external: boolean
}

export type Rels = Map<string, Rel>

const MIME: Record<string, string> = {
  png: 'image/png',
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  jpe: 'image/jpeg',
  gif: 'image/gif',
  bmp: 'image/bmp',
  svg: 'image/svg+xml',
  webp: 'image/webp',
  tif: 'image/tiff',
  tiff: 'image/tiff',
  ico: 'image/x-icon',
  emf: 'image/emf',
  wmf: 'image/wmf',
  mp4: 'video/mp4',
  m4v: 'video/mp4',
  mov: 'video/quicktime',
  mp3: 'audio/mpeg',
  wav: 'audio/wav',
  m4a: 'audio/mp4',
}

export function mimeOf(path: string): string {
  const ext = path.split('.').pop()?.toLowerCase() ?? ''
  return MIME[ext] ?? 'application/octet-stream'
}

/** Browser-renderable image? (EMF/WMF/TIFF are not.) */
export function isWebImage(path: string): boolean {
  return /\.(png|jpe?g|jpe|gif|bmp|svg|webp|ico)$/i.test(path)
}

export function dirOf(part: string): string {
  const i = part.lastIndexOf('/')
  return i < 0 ? '' : part.slice(0, i)
}

export function resolvePath(baseDir: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1)
  const out = baseDir ? baseDir.split('/') : []
  for (const seg of target.split('/')) {
    if (seg === '..') out.pop()
    else if (seg && seg !== '.') out.push(seg)
  }
  return out.join('/')
}

export function relsPathFor(part: string): string {
  const d = dirOf(part)
  const f = part.slice(d.length ? d.length + 1 : 0)
  return (d ? d + '/' : '') + '_rels/' + f + '.rels'
}

export class OpcPackage {
  private relsCache = new Map<string, Promise<Rels>>()
  private xmlCache = new Map<string, Promise<XEl | undefined>>()
  private contentTypes?: Promise<Map<string, string>>

  constructor(readonly zip: ZipArchive) {}

  has(path: string) {
    return this.zip.has(path)
  }

  /** Parsed XML part (cached). Missing parts resolve to undefined. */
  xml(path: string, cache = true): Promise<XEl | undefined> {
    const load = async () => (this.zip.has(path) ? parseXml(await this.zip.text(path), SKIP_EXT) : undefined)
    if (!cache) return load()
    let p = this.xmlCache.get(path)
    if (!p) this.xmlCache.set(path, (p = load()))
    return p
  }

  rels(part: string): Promise<Rels> {
    let p = this.relsCache.get(part)
    if (!p) {
      p = (async () => {
        const map: Rels = new Map()
        const rp = relsPathFor(part)
        if (!this.zip.has(rp)) return map
        const root = parseXml(await this.zip.text(rp))
        const base = dirOf(part)
        for (const r of kids(root, 'Relationship')) {
          const external = r.a.TargetMode === 'External'
          const target = r.a.Target ?? ''
          map.set(r.a.Id, {
            id: r.a.Id,
            type: r.a.Type ?? '',
            external,
            target: external ? target : resolvePath(base, target),
          })
        }
        return map
      })()
      this.relsCache.set(part, p)
    }
    return p
  }

  async contentType(path: string): Promise<string | undefined> {
    if (!this.contentTypes) {
      this.contentTypes = (async () => {
        const m = new Map<string, string>()
        if (!this.zip.has('[Content_Types].xml')) return m
        const root = parseXml(await this.zip.text('[Content_Types].xml'))
        for (const d of kids(root, 'Default')) m.set('ext:' + (d.a.Extension ?? '').toLowerCase(), d.a.ContentType)
        for (const o of kids(root, 'Override')) m.set((o.a.PartName ?? '').replace(/^\//, '').toLowerCase(), o.a.ContentType)
        return m
      })()
    }
    const m = await this.contentTypes
    const key = path.toLowerCase()
    return m.get(key) ?? m.get('ext:' + (key.split('.').pop() ?? ''))
  }

  async blob(path: string): Promise<Blob> {
    const bytes = await this.zip.bytes(path)
    return new Blob([bytes as BlobPart], { type: mimeOf(path) })
  }

  async mainPart(): Promise<{ path: string; kind: DocKind }> {
    const rels = await this.rels('')
    let path: string | undefined
    for (const r of rels.values()) if (r.type.endsWith('/officeDocument')) path = r.target
    const candidates: [string, DocKind][] = [
      ['word/document.xml', 'docx'],
      ['xl/workbook.xml', 'xlsx'],
      ['ppt/presentation.xml', 'pptx'],
    ]
    if (!path) path = candidates.find(([p]) => this.zip.has(p))?.[0]
    if (!path) throw new Error('Not an Office Open XML document (no main part found)')
    const ct = (await this.contentType(path)) ?? ''
    let kind: DocKind | undefined
    if (ct.includes('wordprocessingml')) kind = 'docx'
    else if (ct.includes('spreadsheetml')) kind = 'xlsx'
    else if (ct.includes('presentationml')) kind = 'pptx'
    else kind = candidates.find(([p]) => path!.toLowerCase().startsWith(p.split('/')[0] + '/'))?.[1]
    if (!kind) throw new Error(`Unsupported document type (${ct || path})`)
    return { path, kind }
  }
}

export function relOfType(rels: Rels, suffix: string): Rel | undefined {
  for (const r of rels.values()) if (r.type.endsWith('/' + suffix)) return r
  return undefined
}

export function relsOfType(rels: Rels, suffix: string): Rel[] {
  return [...rels.values()].filter((r) => r.type.endsWith('/' + suffix))
}

const SKIP_EXT = new Set(['extLst'])
