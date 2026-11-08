import { describe, expect, test } from 'bun:test'
import type { Inline, Paragraph, TrackMark } from '../src/core/docx/types'
import { deleteRange, insertText, type RRun } from '../src/react/edit/rich'
import { changeSpans, resolveRuns, trackEdit } from '../src/react/docx/track'

const ME: TrackMark = { k: 'ins', author: 'Vy', date: '2026-11-08T10:00:00Z' }
const para = (runs: Inline[], pid = 1): Paragraph => ({ t: 'p', ps: {}, runs, pid })
const t = (text: string, tc?: TrackMark['k']): Inline => ({ t: 'text', text, s: {}, ...(tc ? { tc: { k: tc, author: 'Lan' } } : {}) })
const caret = (o: number, pa = '0') => ({ a: { pa, o }, f: { pa, o } })
/** text with changes shown as [+ins] and [-del] */
const show = (p: Paragraph) => p.runs.map((r) => (r.t !== 'text' ? '￼' : r.tc ? `[${r.tc.k === 'ins' ? '+' : '-'}${r.text}]` : r.text)).join('')

describe('tracking an edit', () => {
  test('typing is marked inserted, by me', () => {
    const b = para([t('Hello world')])
    const n = para(insertText(b.runs as RRun[], 5, ' big') as Inline[])
    const r = trackEdit([b], [n], '0', caret(9), caret(5), ME)
    expect(show(r.next[0])).toBe('Hello[+ big] world')
    const ins = r.next[0].runs[1] as Inline & { t: 'text' }
    expect(ins.tc).toEqual({ k: 'ins', author: 'Vy', date: '2026-11-08T10:00:00Z' })
    expect(r.sel).toEqual(caret(9))
  })

  test('Backspace keeps the letter, deleted, and the caret moves before it', () => {
    const b = para([t('Hello')])
    const n = para(deleteRange(b.runs as RRun[], 4, 5) as Inline[])
    const r = trackEdit([b], [n], '0', caret(4), caret(5), ME)
    expect(show(r.next[0])).toBe('Hell[-o]')
    expect(r.sel).toEqual(caret(4))
    // again: the deleted letter is skipped, the one before goes
    const n2 = para(deleteRange(r.next[0].runs as RRun[], 3, 4) as Inline[])
    const r2 = trackEdit([r.next[0]], [n2], '0', caret(3), caret(4), ME)
    expect(show(r2.next[0])).toBe('Hel[-lo]')
  })

  test('Delete keeps the letter, deleted, and the caret moves past it', () => {
    const b = para([t('Hello')])
    const n = para(deleteRange(b.runs as RRun[], 0, 1) as Inline[])
    const r = trackEdit([b], [n], '0', caret(0), caret(0), ME)
    expect(show(r.next[0])).toBe('[-H]ello')
    expect(r.sel).toEqual(caret(1))
  })

  test('typing over a selection: deletion then insertion, caret after the new text', () => {
    const b = para([t('Revenue grew 24% this year')])
    const n = para(insertText(deleteRange(b.runs as RRun[], 13, 16), 13, '31%') as Inline[])
    const r = trackEdit([b], [n], '0', caret(16), { a: { pa: '0', o: 13 }, f: { pa: '0', o: 16 } }, ME)
    expect(show(r.next[0])).toBe('Revenue grew [-24%][+31%] this year')
    expect(r.sel).toEqual(caret(19))
  })

  test('removing text I inserted removes it for real; others’ deletions stay', () => {
    const b = para([t('ab'), { ...t('XY'), tc: { k: 'ins', author: 'Vy' } } as Inline, t('cd', 'del'), t('e')])
    const n = para(deleteRange(b.runs as RRun[], 1, 6) as Inline[])
    const r = trackEdit([b], [n], '0', caret(1), { a: { pa: '0', o: 1 }, f: { pa: '0', o: 6 } }, ME)
    // "XY" (mine) is gone; "b" is deleted by me, "cd" stays deleted by Lan
    expect(show(r.next[0])).toBe('a[-b][-cd]e')
    const marks = r.next[0].runs.filter((x) => x.t === 'text' && x.tc).map((x) => (x as { tc: TrackMark }).tc.author)
    expect(marks).toEqual(['Vy', 'Lan'])
  })

  test('formatting and paragraph splits are not tracked', () => {
    const b = para([t('Hello')])
    const n = para([{ ...t('Hello'), s: { b: true } } as Inline])
    expect(trackEdit([b], [n], '0', caret(0), caret(0), ME).next[0]).toBe(n)
    const split = [para([t('Hel')]), para([t('lo')], 2)]
    expect(trackEdit([b], split, '0', caret(0, '1'), caret(3), ME).next).toBe(split)
  })

  test('deleting across paragraphs keeps the text, deleted, in the joined paragraph', () => {
    const b = [para([t('one')]), para([t('two')], 2)]
    const n = [para([t('oo')])]
    const r = trackEdit(b, n, '0', caret(1), { a: { pa: '0', o: 1 }, f: { pa: '1', o: 2 } }, ME)
    expect(show(r.next[0])).toBe('o[-netw]o')
  })

  test('pasting two paragraphs marks both inserted', () => {
    const b = [para([t('ab')])]
    const n = [para([t('aX')]), para([t('Yb')], 2)]
    const r = trackEdit(b, n, '0', caret(1, '1'), caret(1), ME)
    expect(r.next.map(show)).toEqual(['a[+X]', '[+Y]b'])
  })
})

describe('accept / reject', () => {
  const p = para([t('a'), t('BC', 'ins'), t('de', 'del'), t('f')])
  test('accept keeps insertions, drops deletions', () => {
    expect(show(para(resolveRuns(p.runs, 0, 6, true)))).toBe('aBCf')
  })
  test('reject drops insertions, keeps deletions', () => {
    expect(show(para(resolveRuns(p.runs, 0, 6, false)))).toBe('adef')
  })
  test('only inside the range', () => {
    expect(show(para(resolveRuns(p.runs, 0, 3, true)))).toBe('aBC[-de]f')
  })
  test('change spans', () => {
    expect(changeSpans(p).map((x) => [x.a, x.b, x.tc.k])).toEqual([
      [1, 3, 'ins'],
      [3, 5, 'del'],
    ])
  })
})
