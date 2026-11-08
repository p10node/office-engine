import type { Inline, Paragraph, ParaSpec, SpecRun, Table, TableSpec } from './types'

export function inlineSpec(x: Inline): SpecRun {
  switch (x.t) {
    case 'text':
      return { text: x.text, src: x.src, set: x.d && Object.keys(x.d).length ? x.d : undefined, lk: x.lk, cm: x.cm?.length ? x.cm : undefined, ...(x.tc ? { tc: x.tc } : {}) }
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

/** Write-back spec for a table regenerated from the editor model (`t.tid` set). */
export function tableSpec(t: Table): TableSpec {
  return {
    tid: t.tid!,
    item: t.item,
    cols: t.cols,
    rows: t.rows.map((r) => ({
      src: r.src,
      cells: r.cells.map((c) => ({
        src: c.src,
        colSpan: c.colSpan,
        rowSpan: c.rowSpan,
        shd: c.set?.shd ? (c.shd ?? 'none') : undefined,
        borders: c.set?.borders ? c.borders : undefined,
        vAlign: c.set?.vAlign ? (c.vAlign ?? 'top') : undefined,
        paras: c.blocks.filter((b): b is Paragraph => b.t === 'p' && b.pid !== undefined).map((p) => paraSpec([p])),
      })),
    })),
  }
}
