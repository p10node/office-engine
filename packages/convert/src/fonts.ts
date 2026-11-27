/**
 * The font files of a server: found in the system's font folders (and any the host names), read once for their
 * names, styles and weights, and declared by their paths. Text is then set in them when a PDF is written; a file is
 * read from the disk when text first needs it.
 */
import { constants } from 'node:fs'
import { open, opendir, stat } from 'node:fs/promises'
import { homedir } from 'node:os'
import { join, resolve } from 'node:path'
import { Font, FONT_LIMITS } from '../../../src/core/font'

/** a font file PDFs set text in: the engine's `FontSource`, the list `DocViewer`'s `fonts` takes */
export interface FontFile {
  /** the family a document names (Calibri), or the one its font list falls back to (Carlito, sans-serif) */
  family: string
  /** a path, or the file's bytes */
  src: string | ArrayBuffer | ArrayBufferView | Blob
  /** CSS weight: 400, 700, '100 900' */
  weight?: number | string
  /** 'normal', 'italic' */
  style?: string
  unicodeRange?: string
}

/** where Linux, macOS and Windows keep their fonts */
export function systemFontDirs(): string[] {
  const home = homedir()
  const win = process.env.WINDIR ?? 'C:\\Windows'
  return [
    '/usr/share/fonts',
    '/usr/local/share/fonts',
    join(home, '.fonts'),
    join(home, '.local/share/fonts'),
    '/System/Library/Fonts',
    '/Library/Fonts',
    join(home, 'Library/Fonts'),
    join(win, 'Fonts'),
  ]
}

const FILE = /\.(ttf|otf|ttc|woff2?)$/i
/** files looked at, at most: a font folder is not a disk */
const MAX_FILES = 5000
const MAX_DEPTH = 6

/**
 * The generic families of a CSS font list (the last name of every list the engine writes), each served by the first
 * of these that is installed, then by fonts that cover other scripts
 */
const GENERIC: Record<string, string[]> = {
  'sans-serif': ['Liberation Sans', 'Arimo', 'Arial', 'Helvetica Neue', 'Helvetica', 'DejaVu Sans', 'Noto Sans', 'Verdana'],
  serif: ['Liberation Serif', 'Tinos', 'Times New Roman', 'Times', 'DejaVu Serif', 'Noto Serif', 'Georgia'],
  monospace: ['Liberation Mono', 'Cousine', 'Courier New', 'Courier', 'DejaVu Sans Mono', 'Menlo', 'Consolas', 'Noto Sans Mono'],
}
/** fonts of the scripts the ones above do not cover, after them in every generic family */
const COVERAGE = /^(Noto Sans (CJK|JP|SC|TC|KR|Arabic|Hebrew|Devanagari|Thai|Symbols)|Noto Serif CJK|Arial Unicode|PingFang|Hiragino Sans|Microsoft YaHei|Malgun Gothic|Apple SD Gothic|Segoe UI Symbol|Apple Symbols)/i

/**
 * Every font file in `dirs` (searched into their subfolders), declared under its family (and its typographic family,
 * "Avenir Next" for "Avenir Next Demi Bold"), weight and style, by its path; then the generic families. A collection
 * (.ttc) is declared by its first font. Files that cannot be read are left out.
 */
export async function findFonts(dirs: string[]): Promise<FontFile[]> {
  const files: string[] = []
  const seen = new Set<string>()
  // (entries read one by one, not a folder's whole list at once; links to folders are not followed: no loops)
  const walk = async (dir: string, depth: number) => {
    if (depth > MAX_DEPTH || files.length >= MAX_FILES) return
    let entries
    try {
      entries = await opendir(dir)
    } catch {
      return
    }
    try {
      for await (const e of entries) {
        if (files.length >= MAX_FILES) break
        const path = join(dir, e.name)
        if (e.isDirectory()) await walk(path, depth + 1)
        else if ((e.isFile() || e.isSymbolicLink()) && FILE.test(e.name) && !seen.has(path)) {
          seen.add(path)
          files.push(path)
        }
      }
    } catch {
      // (a folder that went away while it was read)
    }
  }
  for (const d of dirs) {
    const dir = resolve(d)
    const st = await stat(dir).catch(() => null)
    if (st?.isDirectory()) await walk(dir, 0)
    else if (st?.isFile() && FILE.test(dir) && !seen.has(dir)) (seen.add(dir), files.push(dir))
  }
  const out: FontFile[] = []
  for (const path of files) {
    try {
      const font = await Font.open(await readFontBytes(path))
      const face = { src: path, weight: font.weight, style: font.italic ? 'italic' : 'normal' }
      for (const family of new Set([font.names.family, font.names.typoFamily].filter(Boolean) as string[])) out.push({ family, ...face })
    } catch {
      // (not a font, damaged, or of a kind that is not read)
    }
  }
  return withGenerics(out)
}

/** `fonts` and, after them, the generic families served by the best of them */
export function withGenerics(fonts: FontFile[]): FontFile[] {
  const byFamily = new Map<string, FontFile[]>()
  for (const f of fonts) {
    const key = f.family.toLowerCase()
    const list = byFamily.get(key)
    if (list) list.push(f)
    else byFamily.set(key, [f])
  }
  // fonts of other scripts: the face of each nearest to regular
  // (all the files of that face: a family may come in several files, each with its own characters)
  const coverage = [...byFamily.keys()]
    .filter((k) => COVERAGE.test(k))
    .flatMap((k) => {
      const faces = byFamily.get(k)!
      const best = Math.min(...faces.map(rank))
      return faces.filter((f) => rank(f) === best)
    })
  const out = [...fonts]
  for (const [generic, names] of Object.entries(GENERIC)) {
    if (byFamily.has(generic)) continue
    const first = names.find((n) => byFamily.has(n.toLowerCase()))
    const main = first ? byFamily.get(first.toLowerCase())! : []
    for (const f of main) out.push({ ...f, family: generic })
    // declared at each weight and style the family before them has: text of any weight and style is set in the face
    // of the family it is nearest to and, where that has no glyph, in these (declared at others, the nearest would
    // be chosen among them and the family's own face left out)
    const looks = main.length ? [...new Map(main.map((f) => [`${f.weight}|${f.style}`, { weight: f.weight, style: f.style }])).values()] : [{ weight: 400, style: 'normal' }]
    for (const f of coverage) for (const look of looks) out.push({ ...f, ...look, family: generic })
  }
  return out
}

/** how far a face is from regular: its weight from 400, italic after upright */
const rank = (f: FontFile) => Math.abs(Number(f.weight ?? 400) - 400) + (/italic|oblique/i.test(f.style ?? '') ? 1000 : 0)

/**
 * A font file's bytes, refused past the size a font may have, also when it grows while it is read. Opened without
 * waiting (a pipe named like a font does not hold the scan up) and read only when it is a file.
 */
async function readFontBytes(path: string): Promise<Uint8Array> {
  const file = await open(path, constants.O_RDONLY | (constants.O_NONBLOCK ?? 0))
  try {
    const st = await file.stat()
    if (!st.isFile()) throw new Error(`${path} is not a file`)
    const max = FONT_LIMITS.file
    if (st.size > max) throw new Error(`${path} is larger than a font file may be`)
    // (room for one byte more than it says: a file that grew is seen to)
    let out = new Uint8Array(st.size + 1)
    let size = 0
    for (;;) {
      if (size === out.length) {
        if (size > max) throw new Error(`${path} is larger than a font file may be`)
        const more = new Uint8Array(Math.min(max + 1, out.length * 2))
        more.set(out)
        out = more
      }
      const { bytesRead } = await file.read(out, size, out.length - size, null)
      if (!bytesRead) return out.subarray(0, size)
      size += bytesRead
    }
  } finally {
    await file.close()
  }
}

let system: Promise<FontFile[]> | undefined

/** the system's fonts, found once */
export function systemFonts(): Promise<FontFile[]> {
  return (system ??= findFonts(systemFontDirs()))
}

/** bytes of font files read lately, by path, up to this many bytes: a server reads the same few for every file */
const CACHE_BYTES = 256 * 1024 * 1024
/** files read and kept, the most recently used last */
const cache = new Map<string, Uint8Array>()
let cached = 0
/** files being read: one read for every conversion that asks at the same time */
const reading = new Map<string, Promise<Uint8Array>>()

/** a font file by its path, from a cache of the ones read lately */
export function readFontFile(path: string): Promise<Uint8Array> {
  const hit = cache.get(path)
  if (hit) {
    cache.delete(path)
    cache.set(path, hit)
    return Promise.resolve(hit)
  }
  let p = reading.get(path)
  if (p) return p
  p = readFontBytes(path).then(
    (bytes) => {
      reading.delete(path)
      // (a file larger than the whole cache is not kept)
      if (bytes.byteLength <= CACHE_BYTES) {
        cache.set(path, bytes)
        cached += bytes.byteLength
        for (const [k, v] of cache) {
          if (cached <= CACHE_BYTES) break
          cache.delete(k)
          cached -= v.byteLength
        }
      }
      return bytes
    },
    (e) => {
      reading.delete(path)
      throw e
    },
  )
  reading.set(path, p)
  return p
}
