/** Browser helpers shared by the diagram packages (not part of their typed API). */

/** A PNG of an SVG picture, `scale` times its size (2: sharp on high-density screens). Browser only. */
export async function svgToPng(svg: string, scale = 2): Promise<Blob> {
  const size = (k: string) => Number(new RegExp(`<svg[^>]*\\s${k}="([\\d.]+)"`).exec(svg)?.[1] ?? 0)
  const w = size('width')
  const h = size('height')
  if (!w || !h) throw new Error('Not a picture with a size')
  const img = new Image()
  img.src = `data:image/svg+xml;charset=utf-8,${encodeURIComponent(svg)}`
  await img.decode()
  const canvas = document.createElement('canvas')
  canvas.width = Math.max(1, Math.round(w * scale))
  canvas.height = Math.max(1, Math.round(h * scale))
  const g = canvas.getContext('2d')
  if (!g) throw new Error('No 2D canvas')
  g.scale(scale, scale)
  g.drawImage(img, 0, 0, w, h)
  return new Promise((resolve, reject) => canvas.toBlob((b) => (b ? resolve(b) : reject(new Error('PNG export failed'))), 'image/png'))
}

/** Puts pictures in an element (the SVG is built by the renderer: text escaped, no scripts or links). */
export function show(target: Element, svgs: string[]) {
  target.innerHTML = svgs.join('')
}
