/** What each glyph of a text command stands for: copy, search and the SVG backend's characters. */
import type { DrawItem } from './types'

type Text = Extract<DrawItem, { t: 'text' }>

/**
 * For each glyph, the text it stands for: its cluster's characters (to the next cluster's start) for the first glyph
 * of a cluster, nothing for the others. (Worked out each time it is asked for: a command can be changed between two
 * exports.)
 */
export function glyphTexts(it: Text): string[] {
  const starts = [...new Set(it.clusters)].sort((a, b) => a - b)
  const end = new Map(starts.map((c, k) => [c, starts[k + 1] ?? it.text.length]))
  const seen = new Set<number>()
  return it.clusters.map((c) => {
    if (seen.has(c)) return ''
    seen.add(c)
    return it.text.slice(c, end.get(c))
  })
}
