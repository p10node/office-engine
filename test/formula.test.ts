import { describe, expect, test } from 'bun:test'
import {
  colIndex,
  evaluate,
  evaluateFormula,
  FormulaError,
  FormulaSyntaxError,
  MAX_COL,
  MAX_ROW,
  parseFormula,
  references,
  renameSheetInFormula,
  shiftFormula,
  translateFormula,
  type CellProvider,
  type ErrorCode,
  type Value,
} from '../src/core/xlsx/formula'

// ---------------------------------------------------------------- in-memory provider

type Sheets = Record<string, Record<string, Value>>

const addr = (a: string) => {
  const m = /^([A-Z]+)(\d+)$/.exec(a)!
  return { r: Number(m[2]) - 1, c: colIndex(m[1]) }
}

function makeProvider(sheets: Sheets, extra: Partial<CellProvider> = {}): CellProvider {
  const cells = new Map<string, Value>()
  const ext = new Map<string, { rows: number; cols: number }>()
  const key = (s: string, r: number, c: number) => `${s.toLowerCase()}!${r},${c}`
  for (const [name, data] of Object.entries(sheets)) {
    let rows = 0
    let cols = 0
    for (const [a, v] of Object.entries(data)) {
      const { r, c } = addr(a)
      cells.set(key(name, r, c), v)
      rows = Math.max(rows, r + 1)
      cols = Math.max(cols, c + 1)
    }
    ext.set(name.toLowerCase(), { rows, cols })
  }
  const provider: CellProvider = {
    value(sheet, r, c) {
      const s = sheet ?? 'Sheet1'
      const v = cells.get(key(s, r, c)) ?? null
      // formula cells: strings starting with '=' are evaluated on demand
      if (typeof v === 'string' && v.startsWith('=')) return evaluateFormula(v.slice(1), { provider, sheet: s, r, c })
      return v
    },
    bounds: (sheet) => ext.get((sheet ?? 'Sheet1').toLowerCase()) ?? { rows: 0, cols: 0 },
    hasSheet: (name) => ext.has(name.toLowerCase()),
    now: () => new Date(2024, 0, 15, 13, 30, 0),
    ...extra,
  }
  return provider
}

const E = (code: ErrorCode) => new FormulaError(code)

const DATA: Sheets = {
  Sheet1: {
    A1: 10, A2: 20, A3: 30, A4: 40, A5: 50,
    B1: 'apple', B2: 'Banana', B3: 'cherry', B4: 'apple pie', B5: 'date',
    C1: 1, C2: 'x', C3: true, /* C4 empty */ C5: '3',
    E1: 1, E2: 3, E3: 5, E4: 7, E5: 9,
    F1: 'one', F2: 'three', F3: 'five', F4: 'seven', F5: 'nine',
    G1: '=A1*2', G2: '=G1+1',
    K1: 'a*b', K2: 'axb',
    M1: E('#DIV/0!'),
  },
  Sheet2: { A1: 100, B1: 200, A2: 5, B4: 'hello' },
  'My Sheet': { A1: 1, A2: 2, A3: 3 },
}
const P = makeProvider(DATA)

/** evaluate on Sheet1 at `at` (default J20) */
function ev(f: string, at = 'J20', provider: CellProvider = P, sheet = 'Sheet1'): Value {
  const { r, c } = addr(at)
  return evaluateFormula(f, { provider, sheet, r, c })
}

function checks(cases: [string, Value][]) {
  for (const [f, want] of cases) {
    const got = ev(f)
    if (typeof want === 'number' && typeof got === 'number' && !Number.isInteger(want)) expect(got, f).toBeCloseTo(want, 9)
    else expect(got, f).toEqual(want)
  }
}

// ---------------------------------------------------------------- parsing & operators

describe('parser / precedence', () => {
  test('arithmetic precedence', () =>
    checks([
      ['1+2*3', 7],
      ['(1+2)*3', 9],
      ['-2^2', 4],
      ['2^3^2', 64],
      ['2*3^2', 18],
      ['10-2-3', 5],
      ['12/2/3', 2],
      ['1+2&"x"', '3x'],
      ['"a"&1+1', 'a2'],
      ['1+1=2', true],
      ['1=1=TRUE', true],
      ['5%', 0.05],
      ['200%*2', 4],
      ['2^50%', Math.SQRT2],
      ['-5%', -0.05],
      ['--"3"', 3],
      ['+"abc"', 'abc'],
      ['1.5E+3', 1500],
      ['.5+1', 1.5],
      [' 1 + 2 ', 3],
      ['=1+1', 2],
    ]))

  test('literals', () =>
    checks([
      ['"say ""hi"""', 'say "hi"'],
      ['TRUE', true],
      ['false', false],
      ['#N/A', E('#N/A')],
      ['IF(ISNA(#N/A),1,2)', 1],
      ['SUM({1,2;3,4})', 10],
      ['SUM({1,-2,"x",TRUE})', -1],
    ]))

  test('syntax errors', () => {
    for (const bad of ['1+', 'SUM(1,2', '"open', '(1', '{1,2;3}', '1 2', 'A1:', '#BOGUS!']) {
      expect(() => parseFormula(bad), bad).toThrow(FormulaSyntaxError)
      expect(ev(bad), bad).toEqual(E('#NAME?'))
    }
  })

  test('ast shape', () => {
    expect(parseFormula('-2^2')).toEqual({
      type: 'binary',
      op: '^',
      left: { type: 'unary', op: '-', arg: { type: 'num', value: 2 } },
      right: { type: 'num', value: 2 },
    })
    const ast = parseFormula("_xlfn.IFERROR('My Sheet'!$A$1:B2,)")
    expect(ast.type === 'call' && ast.name).toBe('IFERROR')
    expect(ast.type === 'call' && ast.args[1]).toEqual({ type: 'missing' })
  })
})

describe('coercion & comparison', () => {
  test('arithmetic coercion', () =>
    checks([
      ['"3"+1', 4],
      ['"abc"+1', E('#VALUE!')],
      ['TRUE+TRUE', 2],
      ['C4+1', 1],
      ['C4&"x"', 'x'],
      ['A1&""', '10'],
      ['1/3&""', '0.333333333333333'],
      ['1E+20&""', '1E+20'],
      ['0.1+0.2=0.3', true],
      ['IF(C4,1,2)', 2],
      ['"2024-01-15"+0', 45306],
      ['"50%"*2', 1],
      ['"1,000"+1', 1001],
      ['C4', 0],
    ]))

  test('comparisons', () =>
    checks([
      ['1<"a"', true],
      ['"a"<TRUE', true],
      ['"abc"="ABC"', true],
      ['"a"<"B"', true],
      ['C4=0', true],
      ['C4=""', true],
      ['C4=FALSE', true],
      ['2>=2', true],
      ['"b"<>"B"', false],
    ]))

  test('errors', () =>
    checks([
      ['1/0', E('#DIV/0!')],
      ['M1+1', E('#DIV/0!')],
      ['SUM(M1,1)', E('#DIV/0!')],
      ['NOSUCHFN(1)', E('#NAME?')],
      ['foo+1', E('#NAME?')],
      ['SQRT(-1)', E('#NUM!')],
      ['Nope!A1', E('#REF!')],
      ['#REF!+1', E('#REF!')],
      ['0^0', E('#NUM!')],
      ['ABS()', E('#VALUE!')],
      ['(-8)^0.5', E('#NUM!')],
    ]))
})

describe('references', () => {
  test('ranges, implicit intersection, whole rows/cols', () => {
    expect(ev('A1:A5*2', 'J3')).toBe(60)
    expect(ev('A1:A5', 'J10')).toEqual(E('#VALUE!'))
    expect(ev('E1:F1', 'F9')).toBe('one')
    expect(ev('A1:B2', 'J1')).toEqual(E('#VALUE!'))
    checks([
      ['SUM(A1:A5)', 150],
      ['SUM(A:A)', 150],
      ['SUM($A$1:$A$3)', 60],
      ['SUM(A5:A1)', 150],
      ['SUM(A1:A2:A4)', 100],
      ['SUM(A1:INDEX(A1:A5,3))', 60],
      ['SUM((A1,A3))', 40],
      ['COUNT((A1:A2,E1:E2))', 4],
      ['SUM(Sheet2!1:1)', 300],
      ['SUM(Sheet2!A:A)', 105],
    ])
  })

  test('cross-sheet', () =>
    checks([
      ['Sheet2!A1+1', 101],
      ["SUM('My Sheet'!A1:A3)", 6],
      ["'My Sheet'!A2*10", 20],
      ['sheet2!B4', 'hello'],
      ['UPPER(Sheet2!B4)', 'HELLO'],
    ]))

  test('formula cells', () =>
    checks([
      ['G1', 20],
      ['G2', 21],
      ['SUM(G1:G2)', 41],
    ]))

  test('evaluate() with pre-parsed ast and other sheet context', () => {
    const ast = parseFormula('A1+B1')
    expect(evaluate(ast, { provider: P, sheet: 'Sheet2', r: 5, c: 5 })).toBe(300)
    expect(ev('A2*2', 'A1', P, 'Sheet2')).toBe(10)
  })
})

// ---------------------------------------------------------------- functions

describe('aggregates & statistics', () => {
  test('basic aggregates', () =>
    checks([
      ['SUM(C1:C5)', 1],
      ['SUM(1,"2",TRUE)', 4],
      ['SUM("x")', E('#VALUE!')],
      ['SUM()', 0],
      ['COUNT(C1:C5)', 1],
      ['COUNT(1,"2","x",TRUE)', 3],
      ['COUNTA(C1:C5)', 4],
      ['COUNTBLANK(C1:C5)', 1],
      ['COUNTBLANK(A:A)', MAX_ROW + 1 - 5],
      ['AVERAGE(A1:A5)', 30],
      ['AVERAGE(B1:B5)', E('#DIV/0!')],
      ['MIN(A1:A5)', 10],
      ['MAX(A1:A5,99)', 99],
      ['MIN(B1:B2)', 0],
      ['PRODUCT(A1:A3)', 6000],
      ['MEDIAN(A1:A4)', 25],
      ['MEDIAN(A1:A5)', 30],
      ['MODE({1,2,2,3,3})', 2],
      ['MODE({1,2,3})', E('#N/A')],
      ['STDEV(A1:A5)', 15.811388300841896],
      ['STDEV.S(A1:A5)', 15.811388300841896],
      ['STDEV.P(A1:A5)', 14.142135623730951],
      ['VAR(A1:A5)', 250],
      ['VAR.S(A1:A5)', 250],
      ['VAR.P(A1:A5)', 200],
      ['STDEV(1)', E('#DIV/0!')],
      ['LARGE(A1:A5,2)', 40],
      ['SMALL(A1:A5,2)', 20],
      ['SMALL(A1:A5,9)', E('#NUM!')],
      ['RANK(30,A1:A5)', 3],
      ['RANK(20,A1:A5,1)', 2],
      ['RANK.EQ(50,A1:A5)', 1],
      ['RANK(31,A1:A5)', E('#N/A')],
    ]))

  test('SUMPRODUCT & array formulas', () =>
    checks([
      ['SUMPRODUCT({1,2},{3,4})', 11],
      ['SUMPRODUCT((A1:A5>20)*A1:A5)', 120],
      ['SUMPRODUCT(--(B1:B5="apple"))', 1],
      ['SUMPRODUCT(A1:A5>20)', 0],
      ['SUMPRODUCT(LEN(B1:B5))', 30],
      ['SUMPRODUCT(A1:A2,E1:E3)', E('#VALUE!')],
      ['SUM(IF(A1:A5>20,1,0))', 3],
      ['SUM(A1:A5*{1;0;0;0;1})', 60],
      ['MAX(IFERROR(1/(A1:A5-30),0))', 0.1],
      ['SUMPRODUCT(ROW(A1:A3))', 6],
    ]))

  test('criteria functions', () =>
    checks([
      ['SUMIF(A1:A5,">25")', 120],
      ['SUMIF(B1:B5,"apple*",A1:A5)', 50],
      ['SUMIF(B1:B5,"APPLE",A1:A5)', 10],
      ['SUMIF(B1:B5,"apple",A1)', 10],
      ['SUMIF(A:A,">=40",E:E)', 16],
      ['SUMIFS(A1:A5,B1:B5,"a*",A1:A5,"<30")', 10],
      ['SUMIFS(A1:A5,B1:B5,"a*",A1:A3,"<30")', E('#VALUE!')],
      ['AVERAGEIF(A1:A5,">=30")', 40],
      ['AVERAGEIF(A1:A5,">100")', E('#DIV/0!')],
      ['AVERAGEIFS(A1:A5,A1:A5,">10",A1:A5,"<50")', 30],
      ['COUNTIF(B1:B5,"*an*")', 1],
      ['COUNTIF(A1:A5,30)', 1],
      ['COUNTIF(A1:A5,"30")', 1],
      ['COUNTIF(A1:A5,"<>30")', 4],
      ['COUNTIF(A1:A5,"<=20")', 2],
      ['COUNTIF(C1:C5,"")', 1],
      ['COUNTIF(C1:C5,"<>")', 4],
      ['COUNTIF(C1:C5,"=")', 1],
      ['COUNTIF(B1:B5,"<>apple")', 4],
      ['COUNTIF(B1:B5,"=apple")', 1],
      ['COUNTIF(B1:B5,">c")', 2],
      ['COUNTIF(C1:C5,TRUE)', 1],
      ['COUNTIF(C1:C5,"true")', 1],
      ['COUNTIF(C1:C5,3)', 1],
      ['COUNTIF(K1:K2,"a~*b")', 1],
      ['COUNTIF(K1:K2,"a*b")', 2],
      ['COUNTIF(K1:K2,"a?b")', 2],
      ['COUNTIF(M1:M2,"#DIV/0!")', 1],
      ['COUNTIF(B1:B5,B1)', 1],
      ['COUNTIFS(A1:A5,">10",B1:B5,"*e*")', 3],
      ['MINIFS(A1:A5,B1:B5,"apple*")', 10],
      ['MAXIFS(A1:A5,B1:B5,"apple*")', 40],
      ['MAXIFS(A1:A5,B1:B5,"zzz")', 0],
    ]))
})

describe('logical & info', () => {
  test('logical', () =>
    checks([
      ['IF(A1>5,"yes","no")', 'yes'],
      ['IF(FALSE,1)', false],
      ['IF(TRUE,)', 0],
      ['IF("x",1,2)', E('#VALUE!')],
      ['IF(1/0>1,1,2)', E('#DIV/0!')],
      ['IF(TRUE,1,1/0)', 1],
      ['IFS(A1>100,"a",A1>5,"b")', 'b'],
      ['IFS(FALSE,1)', E('#N/A')],
      ['IFERROR(1/0,"err")', 'err'],
      ['IFERROR(5,"x")', 5],
      ['IFNA(#N/A,1)', 1],
      ['IFNA(1/0,1)', E('#DIV/0!')],
      ['AND(TRUE,1)', true],
      ['AND(A1:A5)', true],
      ['AND(TRUE,0)', false],
      ['AND("x")', E('#VALUE!')],
      ['AND(B1:B2)', E('#VALUE!')],
      ['OR(FALSE,0)', false],
      ['OR(A1:A5>45)', true],
      ['XOR(TRUE,TRUE,TRUE)', true],
      ['XOR(TRUE,TRUE)', false],
      ['NOT(0)', true],
      ['TRUE()', true],
      ['FALSE()', false],
      ['SWITCH(2,1,"a",2,"b","z")', 'b'],
      ['SWITCH(9,1,"a","z")', 'z'],
      ['SWITCH(9,1,"a")', E('#N/A')],
      ['SWITCH("B","a",1,"b",2)', 2],
      ['CHOOSE(2,"a","b","c")', 'b'],
      ['CHOOSE(4,"a","b","c")', E('#VALUE!')],
      ['SUM(CHOOSE(1,A1:A2,A3:A4))', 30],
    ]))

  test('info', () =>
    checks([
      ['ISBLANK(C4)', true],
      ['ISBLANK(C1)', false],
      ['ISNUMBER(A1)', true],
      ['ISNUMBER("1")', false],
      ['ISTEXT(B1)', true],
      ['ISLOGICAL(C3)', true],
      ['ISERROR(M1)', true],
      ['ISERROR(A1)', false],
      ['ISERR(NA())', false],
      ['ISERR(1/0)', true],
      ['ISNA(NA())', true],
      ['NA()', E('#N/A')],
      ['ERROR.TYPE(1/0)', 2],
      ['ERROR.TYPE(NA())', 7],
      ['ERROR.TYPE(1)', E('#N/A')],
      ['N(TRUE)', 1],
      ['N("x")', 0],
      ['T("x")', 'x'],
      ['T(1)', ''],
      ['SUMPRODUCT(--ISNUMBER(C1:C5))', 1],
    ]))
})

describe('math', () => {
  test('rounding & arithmetic', () =>
    checks([
      ['ROUND(2.675,2)', 2.68],
      ['ROUND(-2.5,0)', -3],
      ['ROUND(1234.5,-2)', 1200],
      ['ROUND(1.005,2)', 1.01],
      ['ROUNDUP(3.2,0)', 4],
      ['ROUNDUP(-3.2,0)', -4],
      ['ROUNDUP(3.14159,3)', 3.142],
      ['ROUNDDOWN(3.9,0)', 3],
      ['ROUNDDOWN(-3.9,0)', -3],
      ['INT(-3.5)', -4],
      ['TRUNC(-3.5)', -3],
      ['TRUNC(3.14159,2)', 3.14],
      ['ABS(-3)', 3],
      ['MOD(-1,3)', 2],
      ['MOD(5,-3)', -1],
      ['MOD(5,0)', E('#DIV/0!')],
      ['POWER(2,10)', 1024],
      ['SQRT(16)', 4],
      ['EXP(1)', Math.E],
      ['LN(EXP(2))', 2],
      ['LOG(100)', 2],
      ['LOG(8,2)', 3],
      ['LOG10(1000)', 3],
      ['LN(0)', E('#NUM!')],
      ['PI()', Math.PI],
      ['SIGN(-5)', -1],
      ['CEILING(2.5,1)', 3],
      ['CEILING(-2.5,-2)', -4],
      ['CEILING(-2.5,2)', -2],
      ['CEILING(2.5,-1)', E('#NUM!')],
      ['FLOOR(2.5,1)', 2],
      ['FLOOR(0.3,0.1)', 0.3],
      ['FLOOR(-2.5,1)', -3],
      ['CEILING.MATH(-2.5)', -2],
      ['CEILING.MATH(-2.5,1,1)', -3],
      ['CEILING.MATH(6.3,5)', 10],
      ['FLOOR.MATH(-2.5)', -3],
      ['FLOOR.MATH(-2.5,1,1)', -2],
      ['FLOOR.MATH(26.75,5)', 25],
      ['FACT(5)', 120],
      ['FACT(-1)', E('#NUM!')],
      ['EVEN(1.5)', 2],
      ['EVEN(-1)', -2],
      ['ODD(2)', 3],
      ['ODD(-2)', -3],
      ['ODD(0)', 1],
      ['GCD(12,18)', 6],
      ['GCD(A1:A5)', 10],
      ['LCM(4,6)', 12],
      ['SUM(ROUND(A1:A5/3,1))', 50],
    ]))

  test('scalar fn at top level intersects its range args', () => {
    expect(ev('ROUND(A1:A5/3,1)', 'J1')).toBe(3.3)
    expect(ev('LEN(B1:B5)', 'J3')).toBe(6)
  })

  test('random', () => {
    for (let i = 0; i < 20; i++) {
      const r = ev('RAND()') as number
      expect(r >= 0 && r < 1).toBe(true)
      const k = ev('RANDBETWEEN(1,3)') as number
      expect([1, 2, 3]).toContain(k)
    }
    expect(ev('RANDBETWEEN(3,1)')).toEqual(E('#NUM!'))
  })
})

describe('text', () => {
  test('text functions', () =>
    checks([
      ['CONCATENATE("a",1,TRUE)', 'a1TRUE'],
      ['CONCAT(B1:B2)', 'appleBanana'],
      ['_xlfn.CONCAT("a",A1:A2)', 'a1020'],
      ['TEXTJOIN("-",TRUE,"a","",C4,"b")', 'a-b'],
      ['TEXTJOIN(",",FALSE,A1:A3)', '10,20,30'],
      ['TEXTJOIN(",",FALSE,"a",C4,"b")', 'a,,b'],
      ['LEFT("hello",2)', 'he'],
      ['LEFT("hello")', 'h'],
      ['LEFT("hello",-1)', E('#VALUE!')],
      ['RIGHT("hello",3)', 'llo'],
      ['RIGHT("hi",0)', ''],
      ['RIGHT("hi",5)', 'hi'],
      ['MID("hello",2,3)', 'ell'],
      ['MID("hello",9,3)', ''],
      ['MID("abc",0,1)', E('#VALUE!')],
      ['LEN("héllo")', 5],
      ['LEN(A1)', 2],
      ['UPPER("abc")', 'ABC'],
      ['LOWER("ABC")', 'abc'],
      ['PROPER("hello wORLD o\'neil 2-way")', "Hello World O'Neil 2-Way"],
      ['TRIM("  a   b  ")', 'a b'],
      ['SUBSTITUTE("a-b-c","-","+")', 'a+b+c'],
      ['SUBSTITUTE("a-b-c","-","+",2)', 'a-b+c'],
      ['SUBSTITUTE("a-b-c","-","+",5)', 'a-b-c'],
      ['REPLACE("abcdef",2,3,"X")', 'aXef'],
      ['REPT("ab",3)', 'ababab'],
      ['REPT("ab",-1)', E('#VALUE!')],
      ['FIND("b","abcb")', 2],
      ['FIND("b","abcb",3)', 4],
      ['FIND("B","abc")', E('#VALUE!')],
      ['SEARCH("B","abc")', 2],
      ['SEARCH("c?","abcd")', 3],
      ['SEARCH("~?","a?b")', 2],
      ['EXACT("a","A")', false],
      ['EXACT("a","a")', true],
      ['HYPERLINK("https://example.com")', 'https://example.com'],
      ['HYPERLINK("https://example.com","Site")', 'Site'],
      ['HYPERLINK("#Sheet2!A1",42)', 42],
      ['VALUE("1,234.5")', 1234.5],
      ['VALUE("50%")', 0.5],
      ['VALUE("1e3")', 1000],
      ['VALUE("abc")', E('#VALUE!')],
      ['CHAR(65)', 'A'],
      ['CODE("A")', 65],
      ['CODE("")', E('#VALUE!')],
    ]))

  test('TEXT formats', () =>
    checks([
      ['TEXT(1234.567,"#,##0.00")', '1,234.57'],
      ['TEXT(1234567,"#,##0")', '1,234,567'],
      ['TEXT(0.256,"0%")', '26%'],
      ['TEXT(0.256,"0.0%")', '25.6%'],
      ['TEXT(5,"0.00")', '5.00'],
      ['TEXT(2.5,"0")', '3'],
      ['TEXT(-5,"0")', '-5'],
      ['TEXT(-5,"0;(0)")', '(5)'],
      ['TEXT(42,"$#,##0.00")', '$42.00'],
      ['TEXT(7,"000")', '007'],
      ['TEXT(DATE(2024,3,5),"yyyy-mm-dd")', '2024-03-05'],
      ['TEXT(DATE(2024,3,5),"dd/mm/yyyy")', '05/03/2024'],
      ['TEXT(DATE(2024,3,5),"mm/dd/yyyy")', '03/05/2024'],
      ['TEXT(DATE(2024,3,5),"d mmm yy")', '5 Mar 24'],
      ['TEXT(DATE(2024,3,5),"dddd, mmmm d")', 'Tuesday, March 5'],
      ['TEXT(TIME(14,5,0),"hh:mm")', '14:05'],
      ['TEXT(TIME(14,5,9),"h:mm:ss AM/PM")', '2:05:09 PM'],
      ['TEXT("2024-01-15","yyyy/mm/dd")', '2024/01/15'],
      ['TEXT(1.5,"General")', '1.5'],
      ['TEXT("abc","0.00")', 'abc'],
    ]))
})

describe('dates', () => {
  test('construct & extract', () =>
    checks([
      ['DATE(2024,1,15)', 45306],
      ['DATE(2024,14,1)', 45689],
      ['DATE(2024,3,0)', 45351],
      ['DATE(1900,3,1)', 61],
      ['DATE(1900,1,1)', 1],
      ['DATE(124,1,15)', 45306],
      ['DATE(-1,1,1)', E('#NUM!')],
      ['YEAR(45306)', 2024],
      ['MONTH(45306)', 1],
      ['DAY(45306)', 15],
      ['MONTH(60)', 2],
      ['DAY(60)', 29],
      ['DAY(61)', 1],
      ['YEAR("2024-01-15")', 2024],
      ['YEAR(-1)', E('#NUM!')],
      ['TIME(12,0,0)', 0.5],
      ['TIME(25,0,0)', 1 / 24],
      ['HOUR(0.75)', 18],
      ['MINUTE(TIME(1,30,0))', 30],
      ['SECOND(TIME(1,2,3))', 3],
      ['HOUR("13:45")', 13],
      ['WEEKDAY(45306)', 2],
      ['WEEKDAY(45306,2)', 1],
      ['WEEKDAY(45306,3)', 0],
      ['WEEKDAY(1)', 1],
      ['EDATE(DATE(2024,1,31),1)', 45351],
      ['EDATE(DATE(2024,3,31),-1)', 45351],
      ['EOMONTH(DATE(2024,1,15),0)', 45322],
      ['EOMONTH(DATE(2024,1,15),1)', 45351],
      ['DATEDIF(DATE(2020,1,15),DATE(2024,3,10),"Y")', 4],
      ['DATEDIF(DATE(2020,1,15),DATE(2024,3,10),"M")', 49],
      ['DATEDIF(DATE(2020,1,15),DATE(2024,3,10),"D")', 45361 - 43845],
      ['DATEDIF(DATE(2020,1,15),DATE(2024,3,10),"MD")', 24],
      ['DATEDIF(DATE(2020,1,15),DATE(2024,3,10),"YM")', 1],
      ['DATEDIF(DATE(2020,1,15),DATE(2024,3,10),"YD")', 55],
      ['DATEDIF(DATE(2024,1,1),DATE(2020,1,1),"D")', E('#NUM!')],
      ['DAYS(DATE(2024,3,1),DATE(2024,2,1))', 29],
      ['DAYS("2024-03-01","2024-02-01")', 29],
      ['NETWORKDAYS(DATE(2024,1,1),DATE(2024,1,31))', 23],
      ['NETWORKDAYS(DATE(2024,1,1),DATE(2024,1,31),{45292,45293})', 21],
      ['NETWORKDAYS(DATE(2024,1,31),DATE(2024,1,1))', -23],
    ]))

  test('TODAY / NOW use provider clock', () => {
    expect(ev('TODAY()')).toBe(45306)
    expect(ev('NOW()')).toBeCloseTo(45306.5625, 9)
    expect(ev('HOUR(NOW())')).toBe(13)
  })

  test('1904 date system', () => {
    const p = makeProvider(DATA, { date1904: true })
    expect(ev('DATE(2024,1,15)', 'J1', p)).toBe(45306 - 1462)
    expect(ev('YEAR(0)', 'J1', p)).toBe(1904)
    expect(ev('TEXT(0,"yyyy-mm-dd")', 'J1', p)).toBe('1904-01-01')
    expect(ev('WEEKDAY(0)', 'J1', p)).toBe(6)
    // restored after evaluation
    expect(ev('DATE(2024,1,15)')).toBe(45306)
  })
})

describe('lookup & reference', () => {
  test('VLOOKUP / HLOOKUP', () =>
    checks([
      ['VLOOKUP(5,E1:F5,2,FALSE)', 'five'],
      ['VLOOKUP(6,E1:F5,2)', 'five'],
      ['VLOOKUP(6,E1:F5,2,TRUE)', 'five'],
      ['VLOOKUP(6,E1:F5,2,FALSE)', E('#N/A')],
      ['VLOOKUP(6,E1:F5,2,)', E('#N/A')],
      ['VLOOKUP(0,E1:F5,2)', E('#N/A')],
      ['VLOOKUP(100,E1:F5,2)', 'nine'],
      ['VLOOKUP(5,E1:F5,3,FALSE)', E('#REF!')],
      ['VLOOKUP(5,E1:F5,0,FALSE)', E('#VALUE!')],
      ['VLOOKUP("th*",F1:F5,1,FALSE)', 'three'],
      ['VLOOKUP("APPLE",B1:B5,1,FALSE)', 'apple'],
      ['VLOOKUP(9,E:F,2,FALSE)', 'nine'],
      ['HLOOKUP("b",{"a","b","c";1,2,3},2,FALSE)', 2],
      ['HLOOKUP("bb",{"a","b","c";1,2,3},2)', 2],
    ]))

  test('XLOOKUP', () =>
    checks([
      ['XLOOKUP("five",F1:F5,E1:E5)', 5],
      ['_xlfn.XLOOKUP("zzz",F1:F5,E1:E5,"none")', 'none'],
      ['XLOOKUP("zzz",F1:F5,E1:E5)', E('#N/A')],
      ['XLOOKUP(6,E1:E5,F1:F5,,-1)', 'five'],
      ['XLOOKUP(6,E1:E5,F1:F5,,1)', 'seven'],
      ['XLOOKUP("t*",F1:F5,E1:E5,,2)', 3],
      ['XLOOKUP("apple*",B1:B5,A1:A5,,2,-1)', 40],
      ['SUM(XLOOKUP(3,E1:E5,A1:B5))', 20],
      ['XLOOKUP(1,E1:E5,A1:A3)', E('#VALUE!')],
    ]))

  test('INDEX / MATCH / LOOKUP', () =>
    checks([
      ['INDEX(E1:F5,3,2)', 'five'],
      ['INDEX(A1:A5,2)', 20],
      ['INDEX(E1:F1,2)', 'one'],
      ['SUM(INDEX(E1:F5,0,1))', 25],
      ['SUM(INDEX(A1:E5,2,0))', 23],
      ['INDEX(E1:F5,6,1)', E('#REF!')],
      ['INDEX({1,2;3,4},2,1)', 3],
      ['INDEX((A1:A5,E1:E5),2,1,2)', 3],
      ['INDEX(A1:A5,MATCH("cherry",B1:B5,0))', 30],
      ['MATCH(5,E1:E5,0)', 3],
      ['MATCH(6,E1:E5)', 3],
      ['MATCH(6,E1:E5,1)', 3],
      ['MATCH(6,{9,7,5,3,1},-1)', 2],
      ['MATCH("ch*",B1:B5,0)', 3],
      ['MATCH("BANANA",B1:B5,0)', 2],
      ['MATCH(100,E1:E5,0)', E('#N/A')],
      ['MATCH(1,A1:B2,0)', E('#N/A')],
      ['LOOKUP(6,E1:E5,F1:F5)', 'five'],
      ['LOOKUP(6,E1:F5)', 'five'],
      ['LOOKUP(2,1/(LEFT(B1:B5,5)="apple"),A1:A5)', 40],
      ['LOOKUP(0,E1:E5,F1:F5)', E('#N/A')],
    ]))

  test('ROW / COLUMN / ROWS / COLUMNS / OFFSET / INDIRECT', () =>
    checks([
      ['ROW()', 20],
      ['ROW(C7)', 7],
      ['COLUMN()', 10],
      ['COLUMN(D1)', 4],
      ['ROWS(A1:B5)', 5],
      ['COLUMNS(A1:C1)', 3],
      ['ROWS(A:A)', MAX_ROW + 1],
      ['COLUMNS(1:1)', MAX_COL + 1],
      ['ROWS({1;2;3})', 3],
      ['SUM(OFFSET(A1,1,0,2,1))', 50],
      ['OFFSET(A1,2,0)', 30],
      ['OFFSET(A1,-1,0)', E('#REF!')],
      ['SUM(OFFSET(A3,0,0,-2))', 50],
      ['INDIRECT("A2")', 20],
      ['INDIRECT("Sheet2!A1")', 100],
      ['SUM(INDIRECT("A1:A3"))', 60],
      ['INDIRECT("\'My Sheet\'!A2")', 2],
      ['INDIRECT("R2C1",FALSE)', 20],
      ['INDIRECT("R[-18]C[-9]",FALSE)', 20],
      ['INDIRECT("bad!!")', E('#REF!')],
      ['INDIRECT("Nope!A1")', E('#REF!')],
      ['SUM(INDIRECT("A"&1):INDIRECT("A"&2))', 30],
    ]))
})

describe('misc evaluation', () => {
  test('_xlfn prefixes & case-insensitive names', () =>
    checks([
      ['_xlfn.IFERROR(1/0,7)', 7],
      ['_xlfn.STDEV.S(A1:A5)', 15.811388300841896],
      ['sum(a1:a2)', 30],
      ['_xlfn._xlws.NOPE(1)', E('#NAME?')],
      ['Table1[Col]', E('#NAME?')],
      ['[1]Sheet1!A1', E('#REF!')],
      ["'[1]My Sheet'!A1:B2", E('#REF!')],
      ['SINGLE(A1:A5)', E('#VALUE!')],
    ]))

  test('top-level implicit intersection through IF / IFERROR', () => {
    expect(ev('IF(A1:A5>20,"big","small")', 'J3')).toBe('big')
    expect(ev('IFERROR(A1:A5,0)', 'J4')).toBe(40)
    expect(ev('_xlfn.SINGLE(A1:A5)', 'J2')).toBe(20)
  })

  test('unknown sheet without hasSheet is passed to provider', () => {
    const p: CellProvider = { value: (s, r, c) => (s === 'X' && r === 0 && c === 0 ? 7 : null) }
    expect(evaluateFormula('X!A1*2', { provider: p, r: 0, c: 5 })).toBe(14)
    // no bounds(): whole column iterates all rows (slow path) but stays correct
    expect(evaluateFormula('COUNTA(X!A1:A100)', { provider: p, r: 0, c: 5 })).toBe(1)
  })
})

// ---------------------------------------------------------------- rewriting

describe('translateFormula', () => {
  test('relative/absolute shifting', () => {
    expect(translateFormula('A1+$B$1+B$2+$C3', 1, 1)).toBe('B2+$B$1+C$2+$C4')
    expect(translateFormula('SUM(A1:B2)', 2, 0)).toBe('SUM(A3:B4)')
    expect(translateFormula('SUM($A1:A$2)', 3, 3)).toBe('SUM($A$2:D4)')
    expect(translateFormula('Sheet2!A1', 0, 1)).toBe('Sheet2!B1')
    expect(translateFormula("'My Sheet'!A1*2", 1, 0)).toBe("'My Sheet'!A2*2")
    expect(translateFormula('SUM(A:A)+SUM(1:1)', 2, 1)).toBe('SUM(B:B)+SUM(3:3)')
    expect(translateFormula('SUM($A:$A)', 2, 1)).toBe('SUM($A:$A)')
    expect(translateFormula('"A1"&A1', 1, 0)).toBe('"A1"&A2')
    expect(translateFormula('LOG10(A1)+_xlfn.STDEV.S(B1:B3)', 1, 0)).toBe('LOG10(A2)+_xlfn.STDEV.S(B2:B4)')
    expect(translateFormula('=A1 + B1', 1, 0)).toBe('=A2 + B2')
    expect(translateFormula('[1]Sheet1!A1+SUM(Table1[Col])', 1, 0)).toBe('[1]Sheet1!A2+SUM(Table1[Col])')
    expect(translateFormula('"unterminated', 1, 0)).toBe('"unterminated')
  })

  test('out of sheet -> #REF!', () => {
    expect(translateFormula('A1', -1, 0)).toBe('#REF!')
    expect(translateFormula('Sheet2!A1+1', 0, -1)).toBe('Sheet2!#REF!+1')
    expect(translateFormula('XFD1', 0, 1)).toBe('#REF!')
    expect(translateFormula('A1048576', 1, 0)).toBe('#REF!')
  })
})

describe('shiftFormula', () => {
  const rows = (text: string, at: number, delta: number, extra: { sheet?: string; formulaSheet?: string } = {}) =>
    shiftFormula(text, { axis: 'row', at, delta, ...extra })
  const cols = (text: string, at: number, delta: number) => shiftFormula(text, { axis: 'col', at, delta })

  test('insert rows', () => {
    expect(rows('A1+A2+A5', 1, 2)).toBe('A1+A4+A7')
    expect(rows('SUM(A1:A5)', 1, 2)).toBe('SUM(A1:A7)')
    expect(rows('SUM(A1:A5)', 5, 2)).toBe('SUM(A1:A5)')
    expect(rows('SUM(A1:A5)', 0, 1)).toBe('SUM(A2:A6)')
    expect(rows('$A$3*2', 1, 2)).toBe('$A$5*2')
    expect(rows('SUM(A:A)', 0, 3)).toBe('SUM(A:A)')
    expect(rows('SUM(3:3)', 1, 2)).toBe('SUM(5:5)')
  })

  test('delete rows', () => {
    expect(rows('A1+A2+A5', 1, -2)).toBe('A1+#REF!+A3')
    expect(rows('SUM(A1:A5)', 1, -2)).toBe('SUM(A1:A3)')
    expect(rows('SUM(A2:A3)', 1, -2)).toBe('SUM(#REF!)')
    expect(rows('SUM(A3:A6)', 1, -2)).toBe('SUM(A2:A4)')
    expect(rows('SUM(A1:A2)', 1, -2)).toBe('SUM(A1:A1)')
    expect(rows('SUM(2:4)', 0, -2)).toBe('SUM(1:2)')
  })

  test('sheet scoping', () => {
    const op = { sheet: 'Sheet1', formulaSheet: 'Sheet1' }
    expect(rows('A5+Sheet2!A5+sheet1!A5', 1, 2, op)).toBe('A7+Sheet2!A5+sheet1!A7')
    // formula on Sheet2, rows inserted in Sheet1
    expect(rows("A5+Sheet1!A5+'Sheet1'!B5", 1, 2, { sheet: 'Sheet1', formulaSheet: 'Sheet2' })).toBe(
      "A5+Sheet1!A7+'Sheet1'!B7",
    )
    expect(rows('Sheet1!A2', 1, -1, { sheet: 'Sheet1', formulaSheet: 'Sheet2' })).toBe('Sheet1!#REF!')
  })

  test('insert / delete columns', () => {
    expect(cols('A1+B1', 0, 1)).toBe('B1+C1')
    expect(cols('SUM(A:B)', 0, 1)).toBe('SUM(B:C)')
    expect(cols('SUM(1:1)', 0, 1)).toBe('SUM(1:1)')
    expect(cols('SUM(A1:C1)', 1, -1)).toBe('SUM(A1:B1)')
    expect(cols('B1', 1, -1)).toBe('#REF!')
    expect(cols('C1*$C$1', 1, -1)).toBe('B1*$B$1')
    expect(cols('SUM(A1:C3)', 1, 2)).toBe('SUM(A1:E3)')
  })
})

describe('renameSheetInFormula', () => {
  test('renames qualifiers with quoting', () => {
    expect(renameSheetInFormula("Sheet1!A1+'Sheet1'!B2+A1", 'Sheet1', 'My Data')).toBe("'My Data'!A1+'My Data'!B2+A1")
    expect(renameSheetInFormula("SUM('Old Name'!A1:A3)", 'Old Name', 'New')).toBe('SUM(New!A1:A3)')
    expect(renameSheetInFormula('Other!A1+sheet1!$C$3', 'Sheet1', 'Data2')).toBe('Other!A1+Data2!$C$3')
    expect(renameSheetInFormula('Sheet1!A1', 'Sheet1', 'S2')).toBe("'S2'!A1")
    expect(renameSheetInFormula('Sheet1!#REF!', 'Sheet1', 'My Data')).toBe("'My Data'!#REF!")
    expect(renameSheetInFormula('"Sheet1!A1"', 'Sheet1', 'X')).toBe('"Sheet1!A1"')
    expect(renameSheetInFormula("'it''s'!A1", "it's", 'ok')).toBe('ok!A1')
    expect(renameSheetInFormula('Sheet1!A1', 'Sheet1', 'A1')).toBe("'A1'!A1")
  })
})

describe('references()', () => {
  test('cells and ranges', () => {
    expect(references(parseFormula('SUM(A1:B2)+Sheet2!$C$3*D:D'))).toEqual([
      { r0: 0, c0: 0, r1: 1, c1: 1 },
      { sheet: 'Sheet2', r: 2, c: 2, absR: true, absC: true },
      { r0: 0, c0: 3, r1: MAX_ROW, c1: 3 },
    ])
    expect(references(parseFormula("IF('My Sheet'!A1>0,2:3,B$7)"))).toEqual([
      { sheet: 'My Sheet', r: 0, c: 0, absR: false, absC: false },
      { r0: 1, c0: 0, r1: 2, c1: MAX_COL },
      { r: 6, c: 1, absR: true, absC: false },
    ])
    expect(references(parseFormula('Sheet2!A1:Sheet2!B3'))).toEqual([{ sheet: 'Sheet2', r0: 0, c0: 0, r1: 2, c1: 1 }])
    expect(references(parseFormula('A1:INDEX(B:B,2)'))).toEqual([
      { r: 0, c: 0, absR: false, absC: false },
      { r0: 0, c0: 1, r1: MAX_ROW, c1: 1 },
    ])
    expect(references(parseFormula('1+2'))).toEqual([])
  })
})
