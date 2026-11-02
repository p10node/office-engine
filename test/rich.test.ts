import { describe, expect, test } from 'bun:test'
import { deleteRange, formatRange, insertText, normalize, prevBoundary, runsText, splitRuns, type RRun } from '../src/react/edit/rich'

const t = (text: string, s: object = {}): RRun => ({ t: 'text', text, s })
const img: RRun = { t: 'img' }

describe('rich text ops', () => {
  test('insert inherits left style, objects count as one char', () => {
    const runs = [t('Hello ', { b: true }), img, t('world')]
    const r = insertText(runs, 6, 'big ')
    expect(runsText(r)).toBe('Hello big ￼world')
    expect(r[0]).toEqual(t('Hello big ', { b: true }))
    expect(runsText(insertText(runs, 7, '!'))).toBe('Hello ￼!world')
  })
  test('delete across runs merges equal neighbours', () => {
    const runs = [t('ab', { i: true }), t('cd'), t('ef', { i: true })]
    const r = deleteRange(runs, 1, 5)
    expect(r).toEqual([t('af', { i: true })])
  })
  test('format splits runs and records direct formatting', () => {
    const r = formatRange([t('abcdef')], 2, 4, { b: true }, { b: true })
    expect(r.map((x) => [x.text, (x.s as any).b ?? false])).toEqual([['ab', false], ['cd', true], ['ef', false]])
    expect(r[1].d).toEqual({ b: true })
  })
  test('split + normalize + graphemes', () => {
    const [l, r] = splitRuns([t('ab'), img, t('cd')], 3)
    expect([runsText(l), runsText(r)]).toEqual(['ab￼', 'cd'])
    expect(normalize([t(''), t('a'), t('b')])).toEqual([t('ab')])
    const s = 'café👍🏽'
    expect(prevBoundary(s, s.length)).toBe(4)
  })
})
