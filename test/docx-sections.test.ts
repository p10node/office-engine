import { describe, expect, test } from 'bun:test'
import type { Block, PageSetup } from '../src/core/docx/types'
import { DEFAULT_PAGE } from '../src/core/docx/types'
import { paginate } from '../src/react/docx/paginate'

const page: PageSetup = { ...DEFAULT_PAGE }
const para = (): Block => ({ t: 'p', ps: {}, runs: [] }) as unknown as Block
const sect = (kind: string): Block => ({ t: 'sect', page, kind })

/** page starts of a list of blocks, each paragraph 10 high, pages 100 high */
function starts(blocks: Block[], lastKind?: string): number[] {
  const heights = new WeakMap<Block, number>()
  for (const b of blocks) heights.set(b, b.t === 'sect' ? 0 : 10)
  return paginate(blocks, heights, blocks.length, 100, 1, lastKind).map((p) => p.start)
}

describe('section breaks', () => {
  test("the break after a section is the next section's type", () => {
    // section 1 (nextPage) ends at 1, section 2 is continuous: no new page before it
    expect(starts([para(), sect('nextPage'), para(), sect('continuous'), para()], 'nextPage')).toEqual([0, 4])
    // section 1 is continuous (its own type says how it starts), section 2 starts on a new page
    expect(starts([para(), sect('continuous'), para(), sect('nextPage'), para()], 'nextPage')).toEqual([0, 2, 4])
  })

  test("the last section's type comes from the body", () => {
    expect(starts([para(), sect('nextPage'), para()], 'continuous')).toEqual([0])
    expect(starts([para(), sect('continuous'), para()])).toEqual([0, 2])
  })
})
