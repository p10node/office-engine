import type { Inline, Paragraph, ParaSpec, SpecRun } from './types'

export function inlineSpec(x: Inline): SpecRun {
  switch (x.t) {
    case 'text':
      return { text: x.text, src: x.src, set: x.d && Object.keys(x.d).length ? x.d : undefined, lk: x.lk, cm: x.cm?.length ? x.cm : undefined }
    case 'field':
      return { obj: 'field', kind: x.kind }
    case 'img':
      if (x.src) return { obj: 'copy', src: x.src, size: x.resized ? [x.w, x.h] : undefined, alt: x.altSet ? (x.alt ?? '') : undefined, rid: x.repl ? x.rid : undefined, cm: x.cm?.length ? x.cm : undefined }
      return x.rid ? { obj: 'image', rid: x.rid, w: x.w, h: x.h, alt: x.alt, cm: x.cm?.length ? x.cm : undefined } : { text: '' }
    default:
      return x.src ? { obj: 'copy', src: x.src } : { text: '' }
  }
}

/** Write-back spec for a paragraph (all page-break parts sharing one pid). */
export function paraSpec(parts: Paragraph[]): ParaSpec {
  const p0 = parts[0]
  return {
    pid: p0.pid!,
    from: p0.from ?? p0.pid!,
    orig: p0.orig,
    pd: p0.pd,
    keepSect: p0.keepSect,
    pristine: !parts.some((p) => p.edited) && (p0.from === undefined || p0.from === p0.pid),
    parts: parts.map((p) => p.runs.map(inlineSpec)),
  }
}
