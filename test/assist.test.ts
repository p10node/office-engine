import { describe, expect, test } from 'bun:test'
import { closeFormula, formulaContext, rangeText } from '../src/core/xlsx/formula/assist'
import { argSlot, functionCatalog, functionInfo, suggestFunctions } from '../src/core/xlsx/formula/catalog'
import { supportedFunctions } from '../src/core/xlsx/formula'

const ctx = (t: string, caret = t.length) => formulaContext(t, caret)

describe('formula catalog', () => {
  test('covers every implemented function', () => {
    const names = new Set(functionCatalog().map((f) => f.name))
    const missing = supportedFunctions().filter((n) => n !== 'SINGLE' && !names.has(n))
    expect(missing).toEqual([])
  })
  test('prefix suggestions', () => {
    expect(suggestFunctions('su').map((f) => f.name)).toEqual(['SUBSTITUTE', 'SUM', 'SUMIF', 'SUMIFS', 'SUMPRODUCT'])
    expect(suggestFunctions('')).toEqual([])
    expect(suggestFunctions('zz')).toEqual([])
  })
  test('argument slots', () => {
    const sum = functionInfo('sum')!
    expect([0, 1, 2, 5].map((i) => argSlot(sum, i))).toEqual([0, 1, 1, 1])
    const sumifs = functionInfo('SUMIFS')!
    expect([0, 1, 2, 3, 4].map((i) => argSlot(sumifs, i))).toEqual([0, 1, 2, 1, 2])
    const vl = functionInfo('VLOOKUP')!
    expect(argSlot(vl, 3)).toBe(3)
    expect(argSlot(vl, 4)).toBe(-1)
  })
})

describe('formula context', () => {
  test('word to autocomplete', () => {
    expect(ctx('=su').word).toEqual({ s: 1, e: 3, text: 'su', call: false })
    expect(ctx('=A1+VLO').word?.text).toBe('VLO')
    expect(ctx('=SUM(').word).toBeUndefined()
    expect(ctx('="su').word).toBeUndefined()
    expect(ctx('=Sheet1!A').word).toBeUndefined()
    // caret inside an existing call name
    expect(ctx('=SUM(A1)', 3).word).toEqual({ s: 1, e: 4, text: 'SU', call: true })
    expect(ctx('abc').word).toBeUndefined()
  })
  test('enclosing call and argument index', () => {
    expect(ctx('=SUM(').call).toEqual({ name: 'SUM', arg: 0 })
    expect(ctx('=IF(A1>0,"x",').call).toEqual({ name: 'IF', arg: 2 })
    expect(ctx('=IF(SUM(A1,B2),').call).toEqual({ name: 'IF', arg: 1 })
    expect(ctx('=IF(SUM(A1,B2').call).toEqual({ name: 'SUM', arg: 1 })
    expect(ctx('=SUM((A1+').call).toEqual({ name: 'SUM', arg: 0 })
    expect(ctx('=SUM(A1)').call).toBeUndefined()
    // inside an unterminated string the call is still known
    expect(ctx('=CONCAT(A1,"ab').call).toEqual({ name: 'CONCAT', arg: 1 })
  })
  test('pointable positions', () => {
    expect(ctx('=').pointable).toBe(true)
    expect(ctx('=SUM(').pointable).toBe(true)
    expect(ctx('=A1+').pointable).toBe(true)
    expect(ctx('=SUM(A1,').pointable).toBe(true)
    expect(ctx('=A1:').pointable).toBe(true)
    expect(ctx('=A1').pointable).toBe(false)
    expect(ctx('=SUM(A1)').pointable).toBe(false)
    expect(ctx('=SUM()', 5).pointable).toBe(true)
    expect(ctx('=SUM(B2)', 5).pointable).toBe(false)
    expect(ctx('="ab').pointable).toBe(false)
    expect(ctx('=5%').pointable).toBe(false)
    expect(ctx('abc').pointable).toBe(false)
  })
  test('reference spans', () => {
    const r = ctx("=SUM(A1:B3,'My Sheet'!C2)+$D$4").refs
    expect(r.map((x) => x.text)).toEqual(['A1:B3', "'My Sheet'!C2", '$D$4'])
    expect(r[1].ref.sheet).toBe('My Sheet')
    expect(r[0]).toMatchObject({ s: 5, e: 10 })
    expect(r[0].ref).toMatchObject({ r0: 0, c0: 0, r1: 2, c1: 1 })
    // refs before a syntax error are still found
    expect(ctx('=A1+B2+"x').refs.map((x) => x.text)).toEqual(['A1', 'B2'])
  })
})

describe('helpers', () => {
  test('rangeText', () => {
    expect(rangeText(0, 0, 0, 0)).toBe('A1')
    expect(rangeText(4, 2, 1, 0)).toBe('A2:C5')
    expect(rangeText(0, 3, 1048575, 3)).toBe('D:D')
    expect(rangeText(2, 0, 4, 16383)).toBe('3:5')
    expect(rangeText(0, 0, 0, 0, 'Sheet2')).toBe('Sheet2!A1')
    expect(rangeText(0, 0, 1, 1, 'My Sheet')).toBe("'My Sheet'!A1:B2")
  })
  test('closeFormula', () => {
    expect(closeFormula('=SUM(A1:A3')).toBe('=SUM(A1:A3)')
    expect(closeFormula('=IF(A1>0,SUM(B1')).toBe('=IF(A1>0,SUM(B1))')
    expect(closeFormula('=CONCAT("ab')).toBe('=CONCAT("ab")')
    expect(closeFormula('=SUM(A1)')).toBe('=SUM(A1)')
    expect(closeFormula('abc(')).toBe('abc(')
    expect(closeFormula("=round(c3*1.1,1)+sheet2!a1+'my s'!$b$2:c3+true")).toBe("=ROUND(C3*1.1,1)+sheet2!A1+'my s'!$B$2:C3+TRUE")
    expect(closeFormula('=if(a1="abc",x')).toBe('=IF(A1="abc",x)')
  })
})
