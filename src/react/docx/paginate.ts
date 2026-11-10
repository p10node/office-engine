import type { Block } from '../../core/docx/types'

/** the blocks [start, end) of a page and the height they take */
export interface PageSlice {
  start: number
  end: number
  used: number
}

/**
 * Cuts the blocks into pages. With text columns (`cols` > 1) a page holds `cols` columns of `avail`: a block that does
 * not fit what is left of a column starts the next one (blocks are never split, as for pages), and `used` is the
 * tallest column.
 */
export function paginate(blocks: Block[], heights: WeakMap<Block, number>, count: number, avail: number, cols = 1, lastKind = 'nextPage'): PageSlice[] {
  // a section's type says how it starts: the break after a section is the next section's ('continuous': no new page);
  // the last section's is the body's (known once loaded)
  const breakAfter = new Map<number, string>()
  let nextKind = lastKind
  for (let i = blocks.length - 1; i >= 0; i--) {
    const b = blocks[i]
    if (b.t !== 'sect') continue
    breakAfter.set(i, nextKind)
    nextKind = b.kind
  }
  const pages: PageSlice[] = []
  let cur: PageSlice = { start: 0, end: 0, used: 0 }
  let col = 0
  let inCol = 0
  const next = (start: number) => {
    pages.push(cur)
    cur = { start, end: start, used: 0 }
    col = 0
    inCol = 0
  }
  for (let i = 0; i < count; i++) {
    const b = blocks[i]
    const h = heights.get(b) ?? 0
    if (b.t === 'sect') {
      cur.end = i + 1
      if (breakAfter.get(i) !== 'continuous') next(i + 1)
      continue
    }
    const forced = b.t === 'p' && b.pb && cur.end > cur.start
    if (forced) next(i)
    else if (inCol > 0 && inCol + h > avail) {
      if (col + 1 < cols) {
        col++
        inCol = 0
      } else next(i)
    }
    cur.end = i + 1
    inCol += h
    cur.used = Math.max(cur.used, cols > 1 ? Math.min(inCol, avail) : inCol)
  }
  if (cur.end > cur.start || !pages.length) pages.push(cur)
  return pages
}
