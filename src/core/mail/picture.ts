/** Pictures of a mail body: their format and size in pixels, read from their first bytes. */

/** the file extension a picture is written with ('png', 'jpeg', 'gif', 'bmp', 'webp'), or null for another format */
export function pictureExt(b: Uint8Array): string | null {
  if (b[0] === 0x89 && b[1] === 0x50 && b[2] === 0x4e && b[3] === 0x47) return 'png'
  if (b[0] === 0xff && b[1] === 0xd8) return 'jpeg'
  if (b[0] === 0x47 && b[1] === 0x49 && b[2] === 0x46) return 'gif'
  if (b[0] === 0x42 && b[1] === 0x4d) return 'bmp'
  if (b[0] === 0x52 && b[1] === 0x49 && b[2] === 0x46 && b[3] === 0x46 && b[8] === 0x57 && b[9] === 0x45 && b[10] === 0x42 && b[11] === 0x50) return 'webp'
  return null
}

/** width and height in pixels, or null when they cannot be read */
export function pictureSize(b: Uint8Array): { w: number; h: number } | null {
  const ext = pictureExt(b)
  if (!ext || b.length < 30) return null
  const dv = new DataView(b.buffer, b.byteOffset, b.byteLength)
  const ok = (w: number, h: number) => (w > 0 && h > 0 ? { w, h } : null)
  switch (ext) {
    case 'png':
      return ok(dv.getUint32(16), dv.getUint32(20))
    case 'gif':
      return ok(dv.getUint16(6, true), dv.getUint16(8, true))
    case 'bmp':
      return ok(Math.abs(dv.getInt32(18, true)), Math.abs(dv.getInt32(22, true)))
    case 'webp': {
      const kind = String.fromCharCode(b[12], b[13], b[14], b[15])
      if (kind === 'VP8 ') return ok(dv.getUint16(26, true) & 0x3fff, dv.getUint16(28, true) & 0x3fff)
      if (kind === 'VP8L') {
        const v = dv.getUint32(21, true)
        return ok((v & 0x3fff) + 1, ((v >> 14) & 0x3fff) + 1)
      }
      if (kind === 'VP8X') return ok((b[24] | (b[25] << 8) | (b[26] << 16)) + 1, (b[27] | (b[28] << 8) | (b[29] << 16)) + 1)
      return null
    }
    case 'jpeg': {
      // the frame header (SOF0-15, not DHT / JPG / DAC) says the size
      let p = 2
      while (p + 9 < b.length) {
        if (b[p] !== 0xff) return null
        const m = b[p + 1]
        if (m === 0xff) {
          p++
          continue
        }
        if (m >= 0xc0 && m <= 0xcf && m !== 0xc4 && m !== 0xc8 && m !== 0xcc) return ok(dv.getUint16(p + 7), dv.getUint16(p + 5))
        if (m === 0xd8 || (m >= 0xd0 && m <= 0xd7) || m === 0x01) {
          p += 2
          continue
        }
        p += 2 + dv.getUint16(p + 2)
      }
      return null
    }
  }
  return null
}
