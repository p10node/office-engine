/**
 * Test helper: writes a compound file (MS-CFB version 3, 512-byte sectors) holding the given streams, for building
 * Word / Excel / PowerPoint 97-2003 test files. Paths use "/" between storages: "ObjectPool/_1/\u0001Ole".
 * Up to 109 FAT sectors (about 6.8 MB): enough for tests.
 */

const SECTOR = 512
const MINI = 64
const CUTOFF = 4096
const END = 0xfffffffe
const FREE = 0xffffffff
const FATSECT = 0xfffffffd
const NONE = 0xffffffff

interface Node {
  name: string
  type: 1 | 2 | 5
  data?: Uint8Array
  kids: Node[]
  id: number
  start: number
  size: number
  left: number
  right: number
  child: number
}

/** CFB name order: shorter names first, then by upper-cased code units */
const cmp = (a: string, b: string) => a.length - b.length || (a.toUpperCase() < b.toUpperCase() ? -1 : a.toUpperCase() > b.toUpperCase() ? 1 : 0)

export function writeCfb(streams: Record<string, Uint8Array>): Uint8Array<ArrayBuffer> {
  const root: Node = { name: 'Root Entry', type: 5, kids: [], id: 0, start: END, size: 0, left: NONE, right: NONE, child: NONE }
  for (const [path, data] of Object.entries(streams)) {
    const parts = path.split('/')
    let at = root
    parts.forEach((name, i) => {
      const leaf = i === parts.length - 1
      let n = at.kids.find((k) => k.name.toUpperCase() === name.toUpperCase())
      if (!n) {
        n = { name, type: leaf ? 2 : 1, kids: [], id: 0, start: END, size: 0, left: NONE, right: NONE, child: NONE }
        at.kids.push(n)
      }
      if (leaf) n.data = data
      at = n
    })
  }
  // directory order: breadth-first; each storage's children as a balanced tree (all black is accepted by readers)
  const all: Node[] = []
  const queue = [root]
  while (queue.length) {
    const n = queue.shift()!
    n.id = all.length
    all.push(n)
    n.kids.sort((a, b) => cmp(a.name, b.name))
    queue.push(...n.kids)
  }
  const tree = (list: Node[]): number => {
    if (!list.length) return NONE
    const mid = list.length >> 1
    const n = list[mid]
    n.left = tree(list.slice(0, mid))
    n.right = tree(list.slice(mid + 1))
    return n.id
  }
  for (const n of all) n.child = tree(n.kids)

  // mini stream: streams under the cutoff, in 64-byte mini sectors
  const mini: number[] = []
  const miniParts: Uint8Array[] = []
  for (const n of all) {
    if (n.type !== 2 || !n.data) continue
    n.size = n.data.length
    if (n.size >= CUTOFF || n.size === 0) continue
    const count = Math.ceil(n.size / MINI)
    n.start = mini.length
    for (let i = 0; i < count; i++) mini.push(i === count - 1 ? END : mini.length + 1)
    const padded = new Uint8Array(count * MINI)
    padded.set(n.data)
    miniParts.push(padded)
  }
  const miniStream = concat(miniParts)
  root.size = miniStream.length

  const secs = (bytes: number) => Math.ceil(bytes / SECTOR)
  const dirSecs = secs(all.length * 128)
  const miniFatSecs = secs(mini.length * 4)
  const miniStreamSecs = secs(miniStream.length)
  const big = all.filter((n) => n.type === 2 && n.data && n.size >= CUTOFF)
  const bigSecs = big.reduce((s, n) => s + secs(n.size), 0)
  let fatSecs = 1
  while (fatSecs * (SECTOR / 4) < fatSecs + dirSecs + miniFatSecs + miniStreamSecs + bigSecs) fatSecs++
  if (fatSecs > 109) throw new Error('cfb-writer: file too large for a test')

  const total = fatSecs + dirSecs + miniFatSecs + miniStreamSecs + bigSecs
  const fat = new Uint32Array(fatSecs * (SECTOR / 4)).fill(FREE)
  let next = 0
  const run = (count: number): number => {
    if (!count) return END
    const start = next
    for (let i = 0; i < count; i++) fat[next + i] = i === count - 1 ? END : next + i + 1
    next += count
    return start
  }
  for (let i = 0; i < fatSecs; i++) fat[next++] = FATSECT
  const dirStart = run(dirSecs)
  const miniFatStart = run(miniFatSecs)
  root.start = miniStream.length ? run(miniStreamSecs) : END
  for (const n of big) n.start = run(secs(n.size))

  const out = new Uint8Array(SECTOR * (1 + total))
  const dv = new DataView(out.buffer)
  out.set([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1])
  dv.setUint16(0x18, 0x3e, true)
  dv.setUint16(0x1a, 3, true)
  dv.setUint16(0x1c, 0xfffe, true)
  dv.setUint16(0x1e, 9, true)
  dv.setUint16(0x20, 6, true)
  dv.setUint32(0x2c, fatSecs, true)
  dv.setUint32(0x30, dirStart, true)
  dv.setUint32(0x38, CUTOFF, true)
  dv.setUint32(0x3c, miniFatSecs ? miniFatStart : END, true)
  dv.setUint32(0x40, miniFatSecs, true)
  dv.setUint32(0x44, END, true)
  for (let i = 0; i < 109; i++) dv.setUint32(0x4c + i * 4, i < fatSecs ? i : FREE, true)
  const at = (sector: number) => (sector + 1) * SECTOR
  for (let i = 0; i < fat.length; i++) dv.setUint32(at(0) + i * 4, fat[i], true)
  all.forEach((n, k) => {
    const off = at(dirStart) + k * 128
    for (let i = 0; i < n.name.length && i < 31; i++) dv.setUint16(off + i * 2, n.name.charCodeAt(i), true)
    dv.setUint16(off + 0x40, (Math.min(n.name.length, 31) + 1) * 2, true)
    out[off + 0x42] = n.type
    out[off + 0x43] = 1
    dv.setUint32(off + 0x44, n.left, true)
    dv.setUint32(off + 0x48, n.right, true)
    dv.setUint32(off + 0x4c, n.child, true)
    dv.setUint32(off + 0x74, n.type === 1 ? 0 : n.start, true)
    dv.setUint32(off + 0x78, n.size, true)
  })
  // unused directory slots: empty entries
  for (let k = all.length; k < dirSecs * 4; k++) {
    const off = at(dirStart) + k * 128
    dv.setUint32(off + 0x44, NONE, true)
    dv.setUint32(off + 0x48, NONE, true)
    dv.setUint32(off + 0x4c, NONE, true)
  }
  for (let i = 0; i < mini.length; i++) dv.setUint32(at(miniFatStart) + i * 4, mini[i], true)
  for (let i = mini.length; i < miniFatSecs * (SECTOR / 4); i++) dv.setUint32(at(miniFatStart) + i * 4, FREE, true)
  if (miniStream.length) out.set(miniStream, at(root.start))
  for (const n of big) out.set(n.data!, at(n.start))
  return out
}

function concat(parts: Uint8Array[]): Uint8Array {
  const out = new Uint8Array(parts.reduce((s, p) => s + p.length, 0))
  let o = 0
  for (const p of parts) {
    out.set(p, o)
    o += p.length
  }
  return out
}
