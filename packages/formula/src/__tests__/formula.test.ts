import { describe, expect, it } from 'vitest'
import { compileFormula, evaluateAst, evaluatePlan, parseCellAddress, parseFormula, parseRangeAddress } from '../index.js'
import { ValueTag, type CellValue } from '@bilig/protocol'

describe('formula', () => {
  it('parses A1 addresses', () => {
    expect(parseCellAddress('B12')).toMatchObject({ row: 11, col: 1, text: 'B12' })
  })

  it('parses quoted sheet addresses', () => {
    expect(parseCellAddress("'My Sheet'!B12")).toMatchObject({
      sheetName: 'My Sheet',
      row: 11,
      col: 1,
      text: 'B12',
    })
  })

  it('normalizes ranges', () => {
    expect(parseRangeAddress('B2:A1')).toMatchObject({
      kind: 'cells',
      start: { text: 'A1' },
      end: { text: 'B2' },
    })
  })

  it('normalizes row and column ranges', () => {
    expect(parseRangeAddress('10:1')).toMatchObject({
      kind: 'rows',
      start: { text: '1' },
      end: { text: '10' },
    })
    expect(parseRangeAddress('C:A')).toMatchObject({
      kind: 'cols',
      start: { text: 'A' },
      end: { text: 'C' },
    })
  })

  it('compiles arithmetic formulas with wasm-safe mode', () => {
    const compiled = compileFormula('A1*2')
    expect(compiled.mode).toBe(1)
    expect([...compiled.symbolicRefs]).toEqual(['A1'])
    expect(compiled.maxStackDepth).toBeGreaterThan(0)
    expect(compiled.id).toBe(0)
    expect(compiled.depsPtr).toBe(0)
    expect(compiled.depsLen).toBe(0)
    expect(compiled.programOffset).toBe(0)
    expect(compiled.constNumberOffset).toBe(0)
  })

  it('compiles postfix percent arithmetic through the existing numeric pipeline', () => {
    const compiled = compileFormula('A1*10%')
    expect(compiled.mode).toBe(1)
    expect([...compiled.symbolicRefs]).toEqual(['A1'])
  })

  it('keeps pass-through cell refs on the wasm-safe path', () => {
    const compiled = compileFormula('A1')
    expect(compiled.mode).toBe(1)
  })

  it('compiles bounded aggregate formulas into the wasm-safe path', () => {
    const compiled = compileFormula('SUM(A1:B2)')
    expect(compiled.mode).toBe(1)
    expect([...compiled.symbolicRefs]).toEqual([])
    expect([...compiled.symbolicRanges]).toEqual(['A1:B2'])
    expect(compileFormula('COUNTBLANK(A1:B2)').mode).toBe(1)
  })

  it('compiles exact-parity logical and rounding builtins onto the wasm-safe path', () => {
    expect(compileFormula('AND(A1,TRUE)').mode).toBe(1)
    expect(compileFormula('OR(A1,FALSE)').mode).toBe(1)
    expect(compileFormula('NOT(A1)').mode).toBe(1)
    expect(compileFormula('DELTA(4,4)').mode).toBe(1)
    expect(compileFormula('GESTEP(4,2)').mode).toBe(1)
    expect(compileFormula('GAUSS(0)').mode).toBe(1)
    expect(compileFormula('PHI(0)').mode).toBe(1)
    expect(compileFormula('ROUND(A1,1)').mode).toBe(1)
    expect(compileFormula('FLOOR(A1,2)').mode).toBe(1)
    expect(compileFormula('CEILING(A1,2)').mode).toBe(1)
    expect(compileFormula('SIN(A1)').mode).toBe(1)
    expect(compileFormula('ATAN2(A1,A2)').mode).toBe(1)
    expect(compileFormula('LOG(A1,10)').mode).toBe(1)
    expect(compileFormula('PI()').mode).toBe(1)
  })

  it('compiles exact-parity info and date builtins onto the wasm-safe path', () => {
    expect(compileFormula('ISBLANK()').mode).toBe(1)
    expect(compileFormula('ISBLANK(A1)').mode).toBe(1)
    expect(compileFormula('ISNUMBER()').mode).toBe(1)
    expect(compileFormula('ISNUMBER(A1)').mode).toBe(1)
    expect(compileFormula('ISTEXT()').mode).toBe(1)
    expect(compileFormula('ISTEXT(A1)').mode).toBe(1)
    expect(compileFormula('T()').mode).toBe(1)
    expect(compileFormula('T("alpha")').mode).toBe(1)
    expect(compileFormula('N()').mode).toBe(1)
    expect(compileFormula('N(A1)').mode).toBe(1)
    expect(compileFormula('TYPE()').mode).toBe(1)
    expect(compileFormula('TYPE(A1)').mode).toBe(1)
    expect(compileFormula('LEN(A1)').mode).toBe(1)
    expect(compileFormula('DATE(2024,2,29)').mode).toBe(1)
    expect(compileFormula('TIME(12,30,0)').mode).toBe(1)
    expect(compileFormula('YEAR(A1)').mode).toBe(1)
    expect(compileFormula('MONTH(A1)').mode).toBe(1)
    expect(compileFormula('DAY(A1)').mode).toBe(1)
    expect(compileFormula('HOUR(A1)').mode).toBe(1)
    expect(compileFormula('MINUTE(A1)').mode).toBe(1)
    expect(compileFormula('SECOND(A1)').mode).toBe(1)
    expect(compileFormula('WEEKDAY(A1)').mode).toBe(1)
    expect(compileFormula('WEEKDAY(A1,2)').mode).toBe(1)
    expect(compileFormula('DAYS(A1,A2)').mode).toBe(1)
    expect(compileFormula('ISOWEEKNUM(A1)').mode).toBe(1)
    expect(compileFormula('TIMEVALUE("1:30 PM")').mode).toBe(1)
    expect(compileFormula('WEEKNUM(A1)').mode).toBe(1)
    expect(compileFormula('WORKDAY(A1,1)').mode).toBe(1)
    expect(compileFormula('WORKDAY(A1,1,B1)').mode).toBe(1)
    expect(compileFormula('NETWORKDAYS(A1,A2)').mode).toBe(1)
    expect(compileFormula('NETWORKDAYS(A1,A2,B1)').mode).toBe(1)
    expect(compileFormula('EDATE(A1,1)').mode).toBe(1)
    expect(compileFormula('EOMONTH(A1,1)').mode).toBe(1)
    expect(compileFormula('EXACT(A1,A2)').mode).toBe(1)
    expect(compileFormula('VALUE("42")').mode).toBe(1)
    expect(compileFormula('INT(A1)').mode).toBe(1)
    expect(compileFormula('ROUNDUP(A1,2)').mode).toBe(1)
    expect(compileFormula('ROUNDDOWN(A1,2)').mode).toBe(1)
    expect(compileFormula('LEFT(A1,2)').mode).toBe(1)
    expect(compileFormula('RIGHT(A1,2)').mode).toBe(1)
    expect(compileFormula('MID(A1,2,3)').mode).toBe(1)
    expect(compileFormula('TRIM(A1)').mode).toBe(1)
    expect(compileFormula('UPPER(A1)').mode).toBe(1)
    expect(compileFormula('LOWER(A1)').mode).toBe(1)
    expect(compileFormula('FIND("a",A1)').mode).toBe(1)
    expect(compileFormula('SEARCH("a",A1)').mode).toBe(1)
    expect(compileFormula('REPLACE(A1,2,3,"x")').mode).toBe(1)
    expect(compileFormula('SUBSTITUTE(A1,"a","b")').mode).toBe(1)
    expect(compileFormula('REPT(A1,3)').mode).toBe(1)
  })

  it('keeps LEN on the JS path when it depends on a range until the range-string bridge lands', () => {
    expect(compileFormula('LEN(A1:B2)').mode).toBe(0)
    expect(compileFormula('VALUE(A1)').mode).toBe(1)
  })

  it('routes volatile scalar builtins through the wasm path', () => {
    expect(compileFormula('TODAY()').mode).toBe(1)
    expect(compileFormula('NOW()').mode).toBe(1)
    expect(compileFormula('RAND()').mode).toBe(1)
  })

  it('routes native sequence spills through the wasm path, including numeric aggregate consumers', () => {
    expect(compileFormula('SEQUENCE(3,1,1,1)')).toMatchObject({ mode: 1, producesSpill: true })
    expect(compileFormula('SUM(SEQUENCE(A1,1,1,1))').mode).toBe(1)
    expect(compileFormula('AVG(SEQUENCE(A1,1,1,1))').mode).toBe(1)
    expect(compileFormula('MIN(SEQUENCE(A1,1,1,1))').mode).toBe(1)
    expect(compileFormula('MAX(SEQUENCE(A1,1,1,1))').mode).toBe(1)
    expect(compileFormula('COUNT(SEQUENCE(A1,1,1,1))').mode).toBe(1)
    expect(compileFormula('COUNTA(SEQUENCE(A1,1,1,1))').mode).toBe(1)
  })

  it('routes dynamic-array family builtins to the wasm path for numeric-compatible inputs', () => {
    expect(compileFormula('OFFSET(A1:B4,0,0,2,2)')).toMatchObject({ mode: 1, producesSpill: true })
    expect(compileFormula('TAKE(A1:B4,2)')).toMatchObject({ mode: 1, producesSpill: true })
    expect(compileFormula('DROP(A1:B4,1)')).toMatchObject({ mode: 1, producesSpill: true })
    expect(compileFormula('CHOOSECOLS(A1:B4,2)')).toMatchObject({ mode: 1, producesSpill: true })
    expect(compileFormula('CHOOSEROWS(A1:B4,2)')).toMatchObject({ mode: 1, producesSpill: true })
    expect(compileFormula('SORT(A1:B4)')).toMatchObject({ mode: 1, producesSpill: true })
    expect(compileFormula('SORTBY(A1:A4,A1:A4)')).toMatchObject({ mode: 1, producesSpill: true })
    expect(compileFormula('TOCOL(A1:B4)')).toMatchObject({ mode: 1, producesSpill: true })
    expect(compileFormula('TOROW(A1:B4)')).toMatchObject({ mode: 1, producesSpill: true })
    expect(compileFormula('WRAPROWS(A1:B4,2)')).toMatchObject({ mode: 1, producesSpill: true })
    expect(compileFormula('WRAPCOLS(A1:B4,2)')).toMatchObject({ mode: 1, producesSpill: true })
  })

  it('routes accelerated text-splitting formulas to the wasm path while keeping indirection helpers on JS', () => {
    expect(compileFormula('TEXTSPLIT(A1,",")')).toMatchObject({ mode: 1, producesSpill: true })
    expect(compileFormula('INDIRECT("A1")').mode).toBe(0)
    expect(compileFormula('FORMULA(A1)').mode).toBe(0)
    expect(compileFormula('GETPIVOTDATA("Sales",A1)').mode).toBe(0)
  })

  it('routes canonical grouped-array SUM fixtures onto the wasm path', () => {
    const groupBy = compileFormula('GROUPBY(A1:A5,C1:C5,SUM,3,1)')
    expect(groupBy.mode).toBe(1)
    expect(groupBy.producesSpill).toBe(true)
    expect([...groupBy.symbolicNames]).toEqual([])

    const pivotBy = compileFormula('PIVOTBY(A1:A5,B1:B5,C1:C5,SUM,3,1,0,1)')
    expect(pivotBy.mode).toBe(1)
    expect(pivotBy.producesSpill).toBe(true)
    expect([...pivotBy.symbolicNames]).toEqual([])
  })

  it('keeps broader workbook-shape grouping formulas on the JS-special path without inventing symbolic aggregate names', () => {
    const groupBy = compileFormula('GROUPBY(A1:A5,B1:B5,SUM)')
    expect(groupBy.mode).toBe(0)
    expect(groupBy.producesSpill).toBe(true)
    expect([...groupBy.symbolicNames]).toEqual([])

    const pivotBy = compileFormula('PIVOTBY(A1:A5,B1:B5,C1:C5,SUM)')
    expect(pivotBy.mode).toBe(0)
    expect(pivotBy.producesSpill).toBe(true)
    expect([...pivotBy.symbolicNames]).toEqual([])

    expect(compileFormula('MULTIPLE.OPERATIONS(A1,B1,C1)').mode).toBe(0)
  })

  it('routes accelerated array-shape helpers to the wasm path', () => {
    expect(compileFormula('EXPAND(A1:B2,3,3)')).toMatchObject({ mode: 1, producesSpill: true })
    expect(compileFormula('TRIMRANGE(A1:C4)')).toMatchObject({ mode: 1, producesSpill: true })
    expect(compileFormula('TRIMRANGE(EXPAND(A1:B2,3,3,0))')).toMatchObject({
      mode: 1,
      producesSpill: true,
    })
  })

  it('routes accelerated date and financial scalar helpers to the wasm path', () => {
    expect(compileFormula('DATEDIF(DATE(2020,1,15),DATE(2021,3,20),"YM")').mode).toBe(1)
    expect(compileFormula('DAYS360(DATE(2024,1,29),DATE(2024,3,31))').mode).toBe(1)
    expect(compileFormula('DAYS360(DATE(2024,1,29),DATE(2024,3,31),TRUE)').mode).toBe(1)
    expect(compileFormula('YEARFRAC(DATE(2024,1,1),DATE(2024,7,1),3)').mode).toBe(1)
    expect(compileFormula('FV(0.1,2,-100,-1000)').mode).toBe(1)
    expect(compileFormula('FVSCHEDULE(1000,0.09,0.11,0.1)').mode).toBe(1)
    expect(compileFormula('PV(0.1,2,-576.1904761904761)').mode).toBe(1)
    expect(compileFormula('PMT(0.1,2,1000)').mode).toBe(1)
    expect(compileFormula('NPER(0.1,-576.1904761904761,1000)').mode).toBe(1)
    expect(compileFormula('NPV(0.1,100,200,300)').mode).toBe(1)
    expect(compileFormula('RATE(48,-200,8000)').mode).toBe(1)
    expect(compileFormula('IPMT(0.1,1,2,1000)').mode).toBe(1)
    expect(compileFormula('PPMT(0.1,1,2,1000)').mode).toBe(1)
    expect(compileFormula('ISPMT(0.1,1,2,1000)').mode).toBe(1)
    expect(compileFormula('CUMIPMT(9%/12,30*12,125000,13,24,0)').mode).toBe(1)
    expect(compileFormula('CUMPRINC(9%/12,30*12,125000,13,24,0)').mode).toBe(1)
    expect(compileFormula('DB(10000,1000,5,1)').mode).toBe(1)
    expect(compileFormula('DDB(2400,300,10,2)').mode).toBe(1)
    expect(compileFormula('VDB(2400,300,10,1,3)').mode).toBe(1)
    expect(compileFormula('SLN(10000,1000,9)').mode).toBe(1)
    expect(compileFormula('SYD(10000,1000,9,1)').mode).toBe(1)
    expect(compileFormula('DISC(DATE(2023,1,1),DATE(2023,4,1),97,100,2)').mode).toBe(1)
    expect(compileFormula('INTRATE(DATE(2023,1,1),DATE(2023,4,1),1000,1030,2)').mode).toBe(1)
    expect(compileFormula('RECEIVED(DATE(2023,1,1),DATE(2023,4,1),1000,0.12,2)').mode).toBe(1)
    expect(compileFormula('PRICEDISC(DATE(2008,2,16),DATE(2008,3,1),0.0525,100,2)').mode).toBe(1)
    expect(compileFormula('YIELDDISC(DATE(2008,2,16),DATE(2008,3,1),99.795,100,2)').mode).toBe(1)
    expect(compileFormula('PRICEMAT(DATE(2008,2,15),DATE(2008,4,13),DATE(2007,11,11),0.061,0.061,0)').mode).toBe(1)
    expect(compileFormula('YIELDMAT(DATE(2008,3,15),DATE(2008,11,3),DATE(2007,11,8),0.0625,100.0123,0)').mode).toBe(1)
    expect(compileFormula('ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,1)').mode).toBe(1)
    expect(compileFormula('ODDFYIELD(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0575,84.5,100,2,0)').mode).toBe(1)
    expect(compileFormula('ODDLPRICE(DATE(2008,2,7),DATE(2008,6,15),DATE(2007,10,15),0.0375,0.0405,100,2,0)').mode).toBe(1)
    expect(compileFormula('ODDLYIELD(DATE(2008,4,20),DATE(2008,6,15),DATE(2007,12,24),0.0375,99.875,100,2,0)').mode).toBe(1)
    expect(compileFormula('COUPDAYBS(DATE(2007,1,25),DATE(2009,11,15),2,4)').mode).toBe(1)
    expect(compileFormula('COUPDAYS(DATE(2007,1,25),DATE(2009,11,15),2,4)').mode).toBe(1)
    expect(compileFormula('COUPDAYSNC(DATE(2007,1,25),DATE(2009,11,15),2,4)').mode).toBe(1)
    expect(compileFormula('COUPNCD(DATE(2007,1,25),DATE(2009,11,15),2,4)').mode).toBe(1)
    expect(compileFormula('COUPNUM(DATE(2007,1,25),DATE(2009,11,15),2,4)').mode).toBe(1)
    expect(compileFormula('COUPPCD(DATE(2007,1,25),DATE(2009,11,15),2,4)').mode).toBe(1)
    expect(compileFormula('PRICE(DATE(2008,2,15),DATE(2017,11,15),0.0575,0.065,100,2,0)').mode).toBe(1)
    expect(compileFormula('YIELD(DATE(2008,2,15),DATE(2016,11,15),0.0575,95.04287,100,2,0)').mode).toBe(1)
    expect(compileFormula('DURATION(DATE(2018,7,1),DATE(2048,1,1),0.08,0.09,2,1)').mode).toBe(1)
    expect(compileFormula('MDURATION(DATE(2008,1,1),DATE(2016,1,1),0.08,0.09,2,1)').mode).toBe(1)
    expect(compileFormula('TBILLPRICE(DATE(2008,3,31),DATE(2008,6,1),0.09)').mode).toBe(1)
    expect(compileFormula('TBILLYIELD(DATE(2008,3,31),DATE(2008,6,1),98.45)').mode).toBe(1)
    expect(compileFormula('TBILLEQ(DATE(2008,3,31),DATE(2008,6,1),0.0914)').mode).toBe(1)
    expect(compileFormula('EFFECT(0.12,12)').mode).toBe(1)
    expect(compileFormula('NOMINAL(0.12682503013196977,12)').mode).toBe(1)
    expect(compileFormula('PDURATION(0.1,100,121)').mode).toBe(1)
    expect(compileFormula('RRI(2,100,121)').mode).toBe(1)
    expect(compileFormula('IRR(A1:A6)').mode).toBe(1)
    expect(compileFormula('MIRR(A1:A6,10%,12%)').mode).toBe(1)
    expect(compileFormula('XNPV(0.09,A1:A5,B1:B5)').mode).toBe(1)
    expect(compileFormula('XIRR(A1:A5,B1:B5)').mode).toBe(1)
  })

  it('routes accelerated regression and covariance helpers to the wasm path', () => {
    expect(compileFormula('CORREL(A1:A3,B1:B3)').mode).toBe(1)
    expect(compileFormula('COVAR(A1:A3,B1:B3)').mode).toBe(1)
    expect(compileFormula('COVARIANCE.P(A1:A3,B1:B3)').mode).toBe(1)
    expect(compileFormula('COVARIANCE.S(A1:A3,B1:B3)').mode).toBe(1)
    expect(compileFormula('PEARSON(A1:A3,B1:B3)').mode).toBe(1)
    expect(compileFormula('INTERCEPT(A1:A3,B1:B3)').mode).toBe(1)
    expect(compileFormula('SLOPE(A1:A3,B1:B3)').mode).toBe(1)
    expect(compileFormula('RSQ(A1:A3,B1:B3)').mode).toBe(1)
    expect(compileFormula('STEYX(A1:A3,B1:B3)').mode).toBe(1)
    expect(compileFormula('FORECAST(4,A1:A3,B1:B3)').mode).toBe(1)
    expect(compileFormula('FORECAST.LINEAR(4,A1:A3,B1:B3)').mode).toBe(1)
    expect(compileFormula('TREND(A1:A3,B1:B3,C1:C2)')).toMatchObject({
      mode: 1,
      producesSpill: true,
    })
    expect(compileFormula('TREND(A1:A3,B1:B3,4)')).toMatchObject({ mode: 1, producesSpill: false })
    expect(compileFormula('GROWTH(A1:A3,B1:B3,C1:C2)')).toMatchObject({
      mode: 1,
      producesSpill: true,
    })
    expect(compileFormula('GROWTH(A1:A3,B1:B3,4)')).toMatchObject({
      mode: 1,
      producesSpill: false,
    })
    expect(compileFormula('LINEST(A1:A3,B1:B3)')).toMatchObject({ mode: 1, producesSpill: true })
    expect(compileFormula('LOGEST(A1:A3,B1:B3)')).toMatchObject({ mode: 1, producesSpill: true })
    expect(compileFormula('LINEST(A1:A3,B1:B3,TRUE,TRUE)')).toMatchObject({
      mode: 1,
      producesSpill: true,
    })
    expect(compileFormula('LOGEST(A1:A3,B1:B3,TRUE,TRUE)')).toMatchObject({
      mode: 1,
      producesSpill: true,
    })
  })

  it('routes accelerated rank helpers to the wasm path', () => {
    expect(compileFormula('RANK(20,A1:A4)').mode).toBe(1)
    expect(compileFormula('RANK.EQ(20,A1:A4)').mode).toBe(1)
    expect(compileFormula('RANK.AVG(20,A1:A4)').mode).toBe(1)
    expect(compileFormula('RANK.AVG(20,A1:A4,1)').mode).toBe(1)
  })

  it('routes accelerated mode and confidence helpers to the wasm path', () => {
    expect(compileFormula('MODE(A1:A6)').mode).toBe(1)
    expect(compileFormula('MODE.SNGL(A1:A6)').mode).toBe(1)
    expect(compileFormula('CONFIDENCE.NORM(0.05,1,100)').mode).toBe(1)
  })

  it('keeps the promoted statistical distribution surface on the wasm path', () => {
    const nativeFormulas = [
      'CONFIDENCE(0.05,1,100)',
      'CONFIDENCE.NORM(0.05,1,100)',
      'CONFIDENCE.T(0.05,1,100)',
      'CHIDIST(18.307,10)',
      'LEGACY.CHIDIST(18.307,10)',
      'CHIINV(0.01,10)',
      'CHISQ.DIST.RT(18.307,10)',
      'CHISQ.DIST(0.5,1,TRUE)',
      'CHISQ.INV.RT(0.01,10)',
      'CHISQ.INV(0.95,1)',
      'CHISQDIST(18.307,10)',
      'CHISQINV(0.01,10)',
      'LEGACY.CHIINV(0.01,10)',
      'CHISQ.TEST(A1:B2,C1:D2)',
      'CHITEST(A1:B2,C1:D2)',
      'LEGACY.CHITEST(A1:B2,C1:D2)',
      'F.TEST(A1:A3,B1:B3)',
      'FTEST(A1:A3,B1:B3)',
      'Z.TEST(A1:A3,2)',
      'ZTEST(A1:A3,2)',
      'BETA.DIST(2,8,10,TRUE,1,3)',
      'BETA.INV(0.6854705810117458,8,10,1,3)',
      'BETADIST(2,8,10,1,3)',
      'BETAINV(0.6854705810117458,8,10,1,3)',
      'F.DIST(1,2,3,TRUE)',
      'F.DIST.RT(1,2,3)',
      'FDIST(1,2,3)',
      'F.INV(0.5,2,3)',
      'F.INV.RT(0.5,2,3)',
      'FINV(0.5,2,3)',
      'LEGACY.FDIST(1,2,3)',
      'LEGACY.FINV(0.5,2,3)',
      'T.DIST(1,1,TRUE)',
      'T.DIST.RT(1,1)',
      'T.DIST.2T(1,1)',
      'T.INV(0.75,1)',
      'T.INV.2T(0.5,1)',
      'TDIST(1,1,2)',
      'TINV(0.5,1)',
      'T.TEST(A1:A3,B1:B3,2,1)',
      'TTEST(A1:A3,B1:B3,2,1)',
      'BINOMDIST(1,3,0.5,TRUE)',
      'BINOM.DIST(1,3,0.5,TRUE)',
      'BINOM.DIST.RANGE(10,0.5,3,6)',
      'CRITBINOM(10,0.5,0.8)',
      'BINOM.INV(10,0.5,0.8)',
      'HYPGEOMDIST(1,4,8,20)',
      'HYPGEOM.DIST(1,4,8,20,TRUE)',
      'NEGBINOMDIST(10,5,0.25)',
      'NEGBINOM.DIST(10,5,0.25,TRUE)',
    ] as const

    for (const formula of nativeFormulas) {
      expect(compileFormula(formula).mode).toBe(1)
    }
  })

  it('keeps the promoted financial helper surface on the wasm path', () => {
    const nativeFormulas = [
      'FV(0.1,2,-576.1904761904761)',
      'PV(0.1,2,-576.1904761904761)',
      'PMT(0.1,2,1000)',
      'NPER(0.1,-576.1904761904761,1000)',
      'RATE(48,-200,8000)',
      'NPV(0.08,100,200,300)',
      'IPMT(0.1,1,2,1000)',
      'PPMT(0.1,1,2,1000)',
      'ISPMT(0.1,1,2,1000)',
      'CUMIPMT(9%/12,30*12,125000,13,24,0)',
      'CUMPRINC(9%/12,30*12,125000,13,24,0)',
      'DISC(DATE(2023,1,1),DATE(2023,4,1),97,100,2)',
      'INTRATE(DATE(2023,1,1),DATE(2023,4,1),1000,1030,2)',
      'RECEIVED(DATE(2023,1,1),DATE(2023,4,1),1000,0.12,2)',
      'COUPDAYBS(DATE(2024,1,15),DATE(2025,1,15),2,0)',
      'COUPDAYS(DATE(2024,1,15),DATE(2025,1,15),2,0)',
      'COUPDAYSNC(DATE(2024,1,15),DATE(2025,1,15),2,0)',
      'COUPNCD(DATE(2024,1,15),DATE(2025,1,15),2,0)',
      'COUPNUM(DATE(2024,1,15),DATE(2025,1,15),2,0)',
      'COUPPCD(DATE(2024,1,15),DATE(2025,1,15),2,0)',
      'PRICEDISC(DATE(2008,2,16),DATE(2008,3,1),0.0525,100,2)',
      'YIELDDISC(DATE(2008,2,16),DATE(2008,3,1),99.795,100,2)',
      'PRICEMAT(DATE(2008,2,15),DATE(2008,4,13),DATE(2007,11,11),0.061,0.061,0)',
      'YIELDMAT(DATE(2008,3,15),DATE(2008,11,3),DATE(2007,11,8),0.0625,100.0123,0)',
      'ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,1)',
      'ODDFYIELD(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0575,84.5,100,2,0)',
      'ODDLPRICE(DATE(2008,2,7),DATE(2008,6,15),DATE(2007,10,15),0.0375,0.0405,100,2,0)',
      'ODDLYIELD(DATE(2008,4,20),DATE(2008,6,15),DATE(2007,12,24),0.0375,99.875,100,2,0)',
      'PRICE(DATE(2024,1,15),DATE(2026,1,15),0.05,0.04,100,2,0)',
      'YIELD(DATE(2024,1,15),DATE(2026,1,15),0.05,101,100,2,0)',
      'DURATION(DATE(2024,1,15),DATE(2026,1,15),0.05,0.04,2,0)',
      'MDURATION(DATE(2024,1,15),DATE(2026,1,15),0.05,0.04,2,0)',
      'TBILLPRICE(DATE(2008,3,31),DATE(2008,6,1),0.09)',
      'TBILLYIELD(DATE(2008,3,31),DATE(2008,6,1),98.45)',
      'TBILLEQ(DATE(2008,3,31),DATE(2008,6,1),0.0914)',
    ] as const

    for (const formula of nativeFormulas) {
      expect(compileFormula(formula).mode).toBe(1)
    }
  })

  it('routes accelerated order-statistics helpers to the wasm path', () => {
    expect(compileFormula('MEDIAN(A1:A8)').mode).toBe(1)
    expect(compileFormula('MEDIAN(A1:A4,9)').mode).toBe(1)
    expect(compileFormula('MODE.MULT(A1:A6)')).toMatchObject({ mode: 1, producesSpill: true })
    expect(compileFormula('FREQUENCY(A1:A6,C1:C3)')).toMatchObject({
      mode: 1,
      producesSpill: true,
    })
    expect(compileFormula('SMALL(A1:A8,3)').mode).toBe(1)
    expect(compileFormula('LARGE(A1:A8,2)').mode).toBe(1)
    expect(compileFormula('PERCENTILE(A1:A8,0.25)').mode).toBe(1)
    expect(compileFormula('PERCENTILE.INC(A1:A8,0.25)').mode).toBe(1)
    expect(compileFormula('PERCENTILE.EXC(A1:A8,0.25)').mode).toBe(1)
    expect(compileFormula('PERCENTRANK(A1:A8,8)').mode).toBe(1)
    expect(compileFormula('PERCENTRANK.INC(A1:A8,8)').mode).toBe(1)
    expect(compileFormula('PERCENTRANK.EXC(A1:A8,8)').mode).toBe(1)
    expect(compileFormula('PROB(A1:A4,B1:B4,2,3)').mode).toBe(1)
    expect(compileFormula('QUARTILE(A1:A8,1)').mode).toBe(1)
    expect(compileFormula('QUARTILE.INC(A1:A8,1)').mode).toBe(1)
    expect(compileFormula('QUARTILE.EXC(A1:A8,1)').mode).toBe(1)
    expect(compileFormula('TRIMMEAN(A1:A8,0.25)').mode).toBe(1)
  })

  it('routes accelerated chi-square inverse functions and aliases to the wasm path', () => {
    expect(compileFormula('CHIDIST(18.307,10)').mode).toBe(1)
    expect(compileFormula('LEGACY.CHIDIST(18.307,10)').mode).toBe(1)
    expect(compileFormula('CHISQDIST(18.307,10)').mode).toBe(1)
    expect(compileFormula('CHIINV(0.050001,10)').mode).toBe(1)
    expect(compileFormula('CHISQ.INV.RT(0.050001,10)').mode).toBe(1)
    expect(compileFormula('CHISQINV(0.050001,10)').mode).toBe(1)
    expect(compileFormula('LEGACY.CHIINV(0.050001,10)').mode).toBe(1)
    expect(compileFormula('CHISQ.INV(0.93,1)').mode).toBe(1)
  })

  it('routes accelerated chi-square test functions and aliases to the wasm path', () => {
    expect(compileFormula('CHISQ.TEST(A1:B3,D1:E3)').mode).toBe(1)
    expect(compileFormula('CHITEST(A1:B3,D1:E3)').mode).toBe(1)
    expect(compileFormula('LEGACY.CHITEST(A1:B3,D1:E3)').mode).toBe(1)
    expect(compileFormula('CHISQ.TEST(SEQUENCE(3,2),SEQUENCE(3,2))').mode).toBe(1)
  })

  it('routes accelerated beta and f distribution functions and aliases to the wasm path', () => {
    expect(compileFormula('BETA.DIST(2,8,10,TRUE,1,3)').mode).toBe(1)
    expect(compileFormula('BETADIST(2,8,10,1,3)').mode).toBe(1)
    expect(compileFormula('BETA.INV(0.6854705810117458,8,10,1,3)').mode).toBe(1)
    expect(compileFormula('BETAINV(0.6854705810117458,8,10,1,3)').mode).toBe(1)
    expect(compileFormula('F.DIST(15.2068649,6,4,TRUE)').mode).toBe(1)
    expect(compileFormula('F.DIST.RT(15.2068649,6,4)').mode).toBe(1)
    expect(compileFormula('FDIST(15.2068649,6,4)').mode).toBe(1)
    expect(compileFormula('LEGACY.FDIST(15.2068649,6,4)').mode).toBe(1)
    expect(compileFormula('F.INV(0.01,6,4)').mode).toBe(1)
    expect(compileFormula('F.INV.RT(0.01,6,4)').mode).toBe(1)
    expect(compileFormula('FINV(0.01,6,4)').mode).toBe(1)
    expect(compileFormula('LEGACY.FINV(0.01,6,4)').mode).toBe(1)
    expect(compileFormula('F.TEST(A1:A5,B1:B5)').mode).toBe(1)
    expect(compileFormula('FTEST(A1:A5,B1:B5)').mode).toBe(1)
    expect(compileFormula('Z.TEST(A1:A5,2,1)').mode).toBe(1)
    expect(compileFormula('ZTEST(A1:A5,2,1)').mode).toBe(1)
  })

  it('routes accelerated student-t distribution functions and aliases to the wasm path', () => {
    expect(compileFormula('T.DIST(1,1,TRUE)').mode).toBe(1)
    expect(compileFormula('T.DIST.RT(1,1)').mode).toBe(1)
    expect(compileFormula('T.DIST.2T(1,1)').mode).toBe(1)
    expect(compileFormula('TDIST(1,1,2)').mode).toBe(1)
    expect(compileFormula('T.INV(0.75,1)').mode).toBe(1)
    expect(compileFormula('T.INV.2T(0.5,1)').mode).toBe(1)
    expect(compileFormula('TINV(0.5,1)').mode).toBe(1)
    expect(compileFormula('CONFIDENCE.T(0.5,2,4)').mode).toBe(1)
    expect(compileFormula('GAMMA.INV(0.08030139707139418,3,2)').mode).toBe(1)
    expect(compileFormula('GAMMAINV(0.08030139707139418,3,2)').mode).toBe(1)
    expect(compileFormula('T.TEST(A1:A3,B1:B3,2,1)').mode).toBe(1)
    expect(compileFormula('TTEST(A1:A3,B1:B3,2,1)').mode).toBe(1)
  })

  it('routes accelerated array-shape and conditional aggregate builtins by public compile contract', () => {
    expect(compileFormula('TRANSPOSE(A1:B4)').mode).toBe(1)
    expect(compileFormula('HSTACK(A1:B2,C1:D2)').mode).toBe(1)
    expect(compileFormula('VSTACK(A1:B2,C1:D2)').mode).toBe(1)
    expect(compileFormula('AREAS(A1:B4)').mode).toBe(1)
    expect(compileFormula('ROWS(A1:B4)').mode).toBe(1)
    expect(compileFormula('COLUMNS(A1:B4)').mode).toBe(1)
    expect(compileFormula('ARRAYTOTEXT(A1:B4)').mode).toBe(1)
    expect(compileFormula('ARRAYTOTEXT(A1:B4,1)').mode).toBe(1)
    expect(compileFormula('MINIFS(A1:A4,B1:B4,">0")').mode).toBe(1)
    expect(compileFormula('MAXIFS(A1:A4,B1:B4,">0")').mode).toBe(1)

    expect(compileFormula('TAKE(A1,1)').mode).toBe(0)
    expect(compileFormula('DROP(A1,1)').mode).toBe(0)
    expect(compileFormula('CHOOSECOLS(A1,1)').mode).toBe(0)
    expect(compileFormula('CHOOSEROWS(A1,1)').mode).toBe(0)
    expect(compileFormula('SORT(A1)').mode).toBe(0)
    expect(compileFormula('TOCOL(A1)').mode).toBe(0)
    expect(compileFormula('TOROW(A1)').mode).toBe(0)
    expect(compileFormula('WRAPROWS(A1,2)').mode).toBe(0)
    expect(compileFormula('WRAPCOLS(A1,2)').mode).toBe(0)
    expect(compileFormula('LOOKUP(A1)').mode).toBe(0)
    expect(compileFormula('AREAS(A1)').mode).toBe(0)
    expect(compileFormula('ROWS(A1)').mode).toBe(0)
    expect(compileFormula('COLUMNS(A1)').mode).toBe(0)
    expect(compileFormula('ARRAYTOTEXT(A1:B4,1,2)').mode).toBe(0)
    expect(compileFormula('MINIFS(A1:A4,B1:B4)').mode).toBe(0)
    expect(compileFormula('MAXIFS(A1,B1:B4,">0")').mode).toBe(0)
    expect(compileFormula('SORTBY(A1:A4)').mode).toBe(0)
  })

  it('keeps row and column aggregate formulas on the JS path', () => {
    expect(compileFormula('SUM(A:A)').mode).toBe(1)
    expect(compileFormula('SUM(1:10)').mode).toBe(1)
  })

  it('keeps contextual metadata functions on the JS path without constant-folding them to errors', () => {
    const rowCompiled = compileFormula('ROW()')
    const sheetCompiled = compileFormula('SHEET("Sheet2")')

    expect(rowCompiled.mode).toBe(0)
    expect(sheetCompiled.mode).toBe(0)
    expect(
      evaluatePlan(rowCompiled.jsPlan, {
        sheetName: 'Sheet1',
        currentAddress: 'C7',
        resolveCell: (): CellValue => ({ tag: ValueTag.Empty }),
        resolveRange: (): CellValue[] => [],
      }),
    ).toEqual({ tag: ValueTag.Number, value: 7 })
    expect(
      evaluatePlan(sheetCompiled.jsPlan, {
        sheetName: 'Sheet1',
        resolveCell: (): CellValue => ({ tag: ValueTag.Empty }),
        resolveRange: (): CellValue[] => [],
        listSheetNames: (): string[] => ['Sheet1', 'Sheet2'],
      }),
    ).toEqual({ tag: ValueTag.Number, value: 2 })
  })

  it('compiles IF, IFERROR, IFNA, and NA onto the wasm-safe path alongside exact-parity logical formulas', () => {
    const compiled = compileFormula('IF(A1>0,A1*2,A2-1)')
    expect(compiled.mode).toBe(1)
    expect(compileFormula('IFERROR(A1,"missing")').mode).toBe(1)
    expect(compileFormula('IFNA(NA(),"missing")').mode).toBe(1)
    expect(compileFormula('NA()').mode).toBe(1)
    expect(compileFormula('COUNTIF(A1:A4,">0")').mode).toBe(1)
    expect(compileFormula('COUNTIFS(A1:A4,">0",B1:B4,"x")').mode).toBe(1)
    expect(compileFormula('SUMIF(A1:A4,">0",B1:B4)').mode).toBe(1)
    expect(compileFormula('SUMIFS(C1:C4,A1:A4,">0",B1:B4,"x")').mode).toBe(1)
    expect(compileFormula('AVERAGEIF(A1:A4,">0")').mode).toBe(1)
    expect(compileFormula('AVERAGEIFS(C1:C4,A1:A4,">0",B1:B4,"x")').mode).toBe(1)
    expect(compileFormula('SUMPRODUCT(A1:A3,B1:B3)').mode).toBe(1)
    expect(compileFormula('MATCH("pear",A1:A3,0)').mode).toBe(1)
    expect(compileFormula('XMATCH("pear",A1:A3)').mode).toBe(1)
    expect(compileFormula('XLOOKUP("pear",A1:A3,B1:B3)').mode).toBe(1)
    expect(compileFormula('INDEX(A1:B3,2,2)').mode).toBe(1)
    expect(compileFormula('VLOOKUP("pear",A1:B3,2,FALSE)').mode).toBe(1)
    expect(compileFormula('HLOOKUP("pear",A1:C2,2,FALSE)').mode).toBe(1)

    expect(compileFormula('AND(A1,TRUE)').mode).toBe(1)
    expect(compileFormula('OR(A1,FALSE)').mode).toBe(1)
    expect(compileFormula('NOT(A1)').mode).toBe(1)
    expect(compileFormula('ROUND(A1,-1)').mode).toBe(1)
    expect(compileFormula('FLOOR(A1,2)').mode).toBe(1)
    expect(compileFormula('CEILING(A1,2)').mode).toBe(1)
  })

  it('keeps unsupported candidate builtin arities on the JS path', () => {
    expect(compileFormula('IF(A1,1)').mode).toBe(0)
    expect(compileFormula('NOT(A1,A2)').mode).toBe(0)
    expect(compileFormula('COUNTIF(A:A,">0")').mode).toBe(0)
    expect(compileFormula('SUMIF(A1:A4,B1:B4,C1:C4,D1:D4)').mode).toBe(0)
    expect(compileFormula('SUMIFS(A1:A4,">0")').mode).toBe(0)
    expect(compileFormula('MATCH("pear",A1:B3,0)').mode).toBe(0)
    expect(compileFormula('XLOOKUP("pear",A1:B3,C1:D4)').mode).toBe(0)
    expect(compileFormula('VLOOKUP("pear",A1:B3,2,FALSE,1)').mode).toBe(0)
    expect(compileFormula('ROUND(A1,A2,A3)').mode).toBe(0)
    expect(compileFormula('FLOOR(A1,A2,A3)').mode).toBe(0)
    expect(compileFormula('CEILING(A1,A2,A3)').mode).toBe(0)
    expect(compileFormula('TIME(A1,A2)').mode).toBe(0)
    expect(compileFormula('WEEKDAY(A1,A2,A3)').mode).toBe(0)
    expect(compileFormula('SIN(A1,A2)').mode).toBe(0)
    expect(compileFormula('SWITCH(A1,1,"yes")').mode).toBe(1)
    expect(compileFormula('WORKDAY(A1,1,B1:B3)').mode).toBe(0)
    expect(compileFormula('NETWORKDAYS(A1,A2,B1:B3)').mode).toBe(0)
    expect(compileFormula('T.DIST(A1,2)').mode).toBe(0)
    expect(compileFormula('TEXT(1234.567,"#,##0.00")').mode).toBe(1)
    expect(compileFormula('PHONETIC(A1:B2)').mode).toBe(1)
    expect(compileFormula('TEXTJOIN(",",TRUE,A1,A2)').mode).toBe(1)
    expect(compileFormula('NUMBERVALUE("2.500,27",",",".")').mode).toBe(1)
    expect(compileFormula('VALUETOTEXT("alpha",1)').mode).toBe(1)
    expect(compileFormula('LENB("é")').mode).toBe(1)
    expect(compileFormula('CHAR(65)').mode).toBe(1)
    expect(compileFormula('CODE("A")').mode).toBe(1)
    expect(compileFormula('UNICODE("A")').mode).toBe(1)
    expect(compileFormula('UNICHAR(66)').mode).toBe(1)
    expect(compileFormula('CLEAN(CHAR(97)&CHAR(1)&CHAR(98))').mode).toBe(1)
    expect(compileFormula('ASC("ＡＢＣ　１２３")').mode).toBe(1)
    expect(compileFormula('JIS("ABC 123")').mode).toBe(1)
    expect(compileFormula('DBCS("ABC 123")').mode).toBe(1)
    expect(compileFormula('BAHTTEXT(1234)').mode).toBe(1)
    expect(compileFormula('DAVERAGE(A1:C5,"Yield",E1:E2)').mode).toBe(1)
    expect(compileFormula('DCOUNT(A1:C5,"Yield",E1:E2)').mode).toBe(1)
    expect(compileFormula('DCOUNTA(A1:C5,"Height",E1:E2)').mode).toBe(1)
    expect(compileFormula('DGET(A1:C5,"Height",F1:F2)').mode).toBe(1)
    expect(compileFormula('DMAX(A1:C5,"Yield",E1:E2)').mode).toBe(1)
    expect(compileFormula('DMIN(A1:C5,"Yield",E1:E2)').mode).toBe(1)
    expect(compileFormula('DPRODUCT(A1:C5,"Yield",E1:E2)').mode).toBe(1)
    expect(compileFormula('DSTDEV(A1:C5,"Yield",E1:E2)').mode).toBe(1)
    expect(compileFormula('DSTDEVP(A1:C5,"Yield",E1:E2)').mode).toBe(1)
    expect(compileFormula('DSUM(A1:C5,"Yield",E1:E2)').mode).toBe(1)
    expect(compileFormula('DVAR(A1:C5,"Yield",E1:E2)').mode).toBe(1)
    expect(compileFormula('DVARP(A1:C5,"Yield",E1:E2)').mode).toBe(1)
    expect(compileFormula('STANDARDIZE(1,0,1)').mode).toBe(1)
    expect(compileFormula('STDEV(A1:A4)').mode).toBe(1)
    expect(compileFormula('STDEVA(2,TRUE(),"skip")').mode).toBe(1)
    expect(compileFormula('VAR(A1:A4)').mode).toBe(1)
    expect(compileFormula('VARA(2,TRUE(),"skip")').mode).toBe(1)
    expect(compileFormula('SKEW(A1:A5)').mode).toBe(1)
    expect(compileFormula('KURT(A1:A5)').mode).toBe(1)
    expect(compileFormula('NORMDIST(1,0,1,TRUE)').mode).toBe(1)
    expect(compileFormula('NORMINV(0.8413447460685429,0,1)').mode).toBe(1)
    expect(compileFormula('NORMSDIST(1)').mode).toBe(1)
    expect(compileFormula('NORMSINV(0.001)').mode).toBe(1)
    expect(compileFormula('LOGINV(0.5,0,1)').mode).toBe(1)
    expect(compileFormula('LOGNORMDIST(1,0,1)').mode).toBe(1)
    expect(compileFormula('LEFTB("abcdef",2)').mode).toBe(1)
    expect(compileFormula('MIDB("abcdef",3,2)').mode).toBe(1)
    expect(compileFormula('RIGHTB("abcdef",3)').mode).toBe(1)
    expect(compileFormula('FINDB("d","abcdef",3)').mode).toBe(1)
    expect(compileFormula('SEARCHB("ph","alphabet")').mode).toBe(1)
    expect(compileFormula('REPLACEB("alphabet",3,2,"Z")').mode).toBe(1)
    expect(compileFormula('ADDRESS(12,3)').mode).toBe(1)
    expect(compileFormula('DOLLAR(-1234.5,1)').mode).toBe(1)
    expect(compileFormula('DOLLARDE(1.08,16)').mode).toBe(1)
    expect(compileFormula('DOLLARFR(1.5,16)').mode).toBe(1)
    expect(compileFormula('BASE(255,16,4)').mode).toBe(1)
    expect(compileFormula('DECIMAL("00FF",16)').mode).toBe(1)
    expect(compileFormula('BIN2DEC("1111111111")').mode).toBe(1)
    expect(compileFormula('BIN2HEX("1111111111")').mode).toBe(1)
    expect(compileFormula('BIN2OCT("1111111111")').mode).toBe(1)
    expect(compileFormula('DEC2BIN(10,8)').mode).toBe(1)
    expect(compileFormula('DEC2HEX(255,4)').mode).toBe(1)
    expect(compileFormula('DEC2OCT(15,4)').mode).toBe(1)
    expect(compileFormula('HEX2BIN("A",8)').mode).toBe(1)
    expect(compileFormula('HEX2DEC("FFFFFFFFFF")').mode).toBe(1)
    expect(compileFormula('HEX2OCT("F",4)').mode).toBe(1)
    expect(compileFormula('OCT2BIN("12",8)').mode).toBe(1)
    expect(compileFormula('OCT2DEC("17")').mode).toBe(1)
    expect(compileFormula('OCT2HEX("17",4)').mode).toBe(1)
    expect(compileFormula('CONVERT(6,"mi","km")').mode).toBe(1)
    expect(compileFormula('EUROCONVERT(1.2,"DEM","EUR")').mode).toBe(1)
    expect(compileFormula('EUROCONVERT(1,"FRF","DEM",TRUE,3)').mode).toBe(1)
    expect(compileFormula('BESSELI(1.5,1)').mode).toBe(1)
    expect(compileFormula('BESSELJ(1.9,2)').mode).toBe(1)
    expect(compileFormula('BESSELK(1.5,1)').mode).toBe(1)
    expect(compileFormula('BESSELY(2.5,1)').mode).toBe(1)
    expect(compileFormula('BITAND(6,3)').mode).toBe(1)
    expect(compileFormula('BITOR(6,3)').mode).toBe(1)
    expect(compileFormula('BITXOR(6,3)').mode).toBe(1)
    expect(compileFormula('BITLSHIFT(1,4)').mode).toBe(1)
    expect(compileFormula('BITRSHIFT(16,4)').mode).toBe(1)
    expect(compileFormula('USE.THE.COUNTIF(A1:A3,">0")').mode).toBe(1)
    expect(compileFormula('WORKDAY.INTL(A1,1)').mode).toBe(1)
    expect(compileFormula('NETWORKDAYS.INTL(A1,A2,7)').mode).toBe(1)
    expect(compileFormula('LET(x,2,x+3)').mode).toBe(1)
    expect(compileFormula('LET(x,A1*2,x+3)').mode).toBe(1)
    expect(compileFormula('TEXTBEFORE(A1,"-")').mode).toBe(1)
    expect(compileFormula('TEXTAFTER(A1,"-")').mode).toBe(1)
    expect(compileFormula('CHOOSE(2,"red","blue","green")').mode).toBe(1)
    expect(compileFormula('CHOOSE(1,A1:B2,C1:D2)')).toMatchObject({ mode: 1, producesSpill: true })
    expect(compileFormula('FILTER(A1:A4,A1:A4>2)')).toMatchObject({ mode: 1, producesSpill: true })
    expect(compileFormula('UNIQUE(A1:A4)')).toMatchObject({ mode: 1, producesSpill: true })
    expect(compileFormula('FILTER(A1:A4,B1:B4)')).toMatchObject({ mode: 1, producesSpill: true })
    expect(compileFormula('A1:A4>2')).toMatchObject({ mode: 1, producesSpill: true })
  })

  it('accelerates rewritten logical calls while keeping higher-order lambda families on the JS path', () => {
    expect(compileFormula('TRUE()').mode).toBe(1)
    expect(compileFormula('FALSE()').mode).toBe(1)
    expect(compileFormula('SEQUENCE(3,1,1,1)')).toMatchObject({ mode: 1, producesSpill: true })
    expect(compileFormula('IFS(A1>0,1,TRUE(),2)').mode).toBe(1)
    expect(compileFormula('SWITCH(A1,1,"yes","no")').mode).toBe(1)
    expect(compileFormula('XOR(A1>0,B1>0)').mode).toBe(1)

    expect(compileFormula('LAMBDA(x,x+1)(4)').mode).toBe(1)
    expect(compileFormula('LAMBDA(x,x+1)(A1)').mode).toBe(1)
    expect(compileFormula('LET(fn,LAMBDA(x,x+1),fn(4))').mode).toBe(1)
    expect(compileFormula('LAMBDA(x,y,IF(ISOMITTED(y),x,y))(4)').mode).toBe(0)
    expect(compileFormula('MAKEARRAY(2,2,LAMBDA(r,c,r+c))')).toMatchObject({
      mode: 1,
      producesSpill: true,
    })
    expect(compileFormula('MAP(A1:A3,LAMBDA(x,x*2))')).toMatchObject({
      mode: 1,
      producesSpill: true,
    })
    expect(compileFormula('MAP(A1:A3,B1:B3,LAMBDA(x,y,x+y))')).toMatchObject({
      mode: 1,
      producesSpill: true,
    })
    expect(compileFormula('SCAN(0,A1:A3,LAMBDA(acc,x,acc+x))')).toMatchObject({
      mode: 1,
      producesSpill: true,
    })
    expect(compileFormula('SCAN(1,A1:A3,LAMBDA(acc,x,acc*x))')).toMatchObject({
      mode: 1,
      producesSpill: true,
    })
    expect(compileFormula('BYROW(A1:B2,LAMBDA(r,SUM(r)))')).toMatchObject({
      mode: 1,
      producesSpill: true,
    })
    expect(compileFormula('BYROW(A1:B2,LAMBDA(r,AVERAGE(r)))')).toMatchObject({
      mode: 1,
      producesSpill: true,
    })
    expect(compileFormula('BYCOL(A1:B2,LAMBDA(c,SUM(c)))')).toMatchObject({
      mode: 1,
      producesSpill: true,
    })
    expect(compileFormula('BYCOL(A1:B2,LAMBDA(c,COUNTA(c)))')).toMatchObject({
      mode: 1,
      producesSpill: true,
    })
    expect(compileFormula('REDUCE(0,A1:A3,LAMBDA(acc,x,acc+x))').mode).toBe(1)
    expect(compileFormula('REDUCE(1,A1:A3,LAMBDA(acc,x,acc*x))').mode).toBe(1)
  })

  it('routes promoted scalar and reducer math families through wasm while keeping matrix-only families on JS', () => {
    expect(compileFormula('TRUNC(A1,1)').mode).toBe(1)
    expect(compileFormula('FLOOR.MATH(-5.5,2)').mode).toBe(1)
    expect(compileFormula('FLOOR.PRECISE(-5.5,2)').mode).toBe(1)
    expect(compileFormula('CEILING.MATH(-5.5,2)').mode).toBe(1)
    expect(compileFormula('CEILING.PRECISE(-5.5,2)').mode).toBe(1)
    expect(compileFormula('ISO.CEILING(-5.5,2)').mode).toBe(1)
    expect(compileFormula('MROUND(10,4)').mode).toBe(1)
    expect(compileFormula('PERMUT(5,3)').mode).toBe(1)
    expect(compileFormula('PERMUTATIONA(2,3)').mode).toBe(1)
    expect(compileFormula('SERIESSUM(2,1,2,1,2)').mode).toBe(1)
    expect(compileFormula('SQRTPI(2)').mode).toBe(1)
    expect(compileFormula('ACOSH(A1)').mode).toBe(1)
    expect(compileFormula('ASINH(A1)').mode).toBe(1)
    expect(compileFormula('ATANH(A1)').mode).toBe(1)
    expect(compileFormula('COT(A1)').mode).toBe(1)
    expect(compileFormula('SEC(A1)').mode).toBe(1)
    expect(compileFormula('PRODUCT(A1:A3)').mode).toBe(1)
    expect(compileFormula('SUMSQ(A1:A3)').mode).toBe(1)
    expect(compileFormula('GEOMEAN(A1:A3)').mode).toBe(1)
    expect(compileFormula('HARMEAN(A1:A3)').mode).toBe(1)
    expect(compileFormula('GCD(A1:A3)').mode).toBe(1)
    expect(compileFormula('LCM(A1:A3)').mode).toBe(1)
    expect(compileFormula('COMBIN(8,3)').mode).toBe(1)
    expect(compileFormula('QUOTIENT(7,3)').mode).toBe(1)
    expect(compileFormula('SUMXMY2(A1:A3,B1:B3)').mode).toBe(0)
    expect(compileFormula('MDETERM(A1:B2)').mode).toBe(0)
    expect(compileFormula('MMULT(A1:B2,C1:D2)')).toMatchObject({ mode: 0, producesSpill: true })
    expect(compileFormula('MINVERSE(A1:B2)')).toMatchObject({ mode: 0, producesSpill: true })
    expect(compileFormula('MUNIT(3)')).toMatchObject({ mode: 0, producesSpill: true })
    expect(compileFormula('RANDARRAY(2,2)')).toMatchObject({
      mode: 0,
      producesSpill: true,
      volatile: true,
    })
    expect(compileFormula('RANDBETWEEN(1,10)').volatile).toBe(true)
  })

  it('keeps unsupported promoted argument shapes on JS while preserving native routing for valid helpers', () => {
    expect(compileFormula('ATAN2(A1)').mode).toBe(0)
    expect(compileFormula('EXACT(A1)').mode).toBe(0)
    expect(compileFormula('CONVERT(A1:A2,"m","s")').mode).toBe(0)
    expect(compileFormula('EUROCONVERT(A1:A2,"DEM","EUR")').mode).toBe(0)
    expect(compileFormula('ISO.CEILING("bad")').mode).toBe(1)

    expect(compileFormula('BIN2HEX("111",4)').mode).toBe(1)
    expect(compileFormula('HEX2DEC("FF")').mode).toBe(1)
    expect(compileFormula('CHAR(65)').mode).toBe(1)
    expect(compileFormula('UNICODE("A")').mode).toBe(1)
    expect(compileFormula('PRODUCT(A1:A3,A4:A6)').mode).toBe(1)
    expect(compileFormula('GCD(18,24,30)').mode).toBe(1)
    expect(compileFormula('LCM(3,4,5)').mode).toBe(1)
  })

  it('evaluates AST against a context', () => {
    const ast = parseFormula('A1+A2')
    const context = {
      sheetName: 'Sheet1',
      resolveCell: (_sheet: string, address: string): CellValue => {
        if (address === 'A1') return { tag: ValueTag.Number, value: 2 }
        return { tag: ValueTag.Number, value: 3 }
      },
      resolveRange: (): CellValue[] => [],
    }
    const value = evaluateAst(ast, context)
    expect(value).toEqual({ tag: ValueTag.Number, value: 5 })
    expect(evaluatePlan(compileFormula('A1+A2').jsPlan, context)).toEqual({
      tag: ValueTag.Number,
      value: 5,
    })
    expect(evaluateAst(parseFormula('10%'), context)).toEqual({ tag: ValueTag.Number, value: 0.1 })
    expect(evaluateAst(parseFormula('(A1+A2)%'), context)).toEqual({
      tag: ValueTag.Number,
      value: 0.05,
    })
    expect(evaluatePlan(compileFormula('A1*10%').jsPlan, context)).toEqual({
      tag: ValueTag.Number,
      value: 0.2,
    })
  })

  it('parses quoted sheet references inside formulas', () => {
    const ast = parseFormula("'My Sheet'!A1+1")
    expect(ast).toMatchObject({
      kind: 'BinaryExpr',
      left: { kind: 'CellRef', sheetName: 'My Sheet', ref: 'A1' },
    })
  })

  it('compiles quoted sheet ranges into symbolic refs', () => {
    const compiled = compileFormula("SUM('My Sheet'!A1:B2)")
    expect([...compiled.symbolicRefs]).toEqual([])
    expect([...compiled.symbolicRanges]).toEqual(["'My Sheet'!A1:B2"])
  })

  it('parses quoted sheet column ranges inside formulas', () => {
    const ast = parseFormula("SUM('My Sheet'!A:A)")
    expect(ast).toMatchObject({
      kind: 'CallExpr',
      args: [{ kind: 'RangeRef', refKind: 'cols', sheetName: 'My Sheet', start: 'A', end: 'A' }],
    })
  })

  it('preserves absolute and mixed references in formulas', () => {
    const ast = parseFormula('SUM($A1,B$2,$C$3,$4:5,$D:$F)')
    expect(ast).toMatchObject({
      kind: 'CallExpr',
      callee: 'SUM',
      args: [
        { kind: 'CellRef', ref: '$A1' },
        { kind: 'CellRef', ref: 'B$2' },
        { kind: 'CellRef', ref: '$C$3' },
        { kind: 'RangeRef', refKind: 'rows', start: '$4', end: '5' },
        { kind: 'RangeRef', refKind: 'cols', start: '$D', end: '$F' },
      ],
    })
  })

  it('constant folds numeric expressions and prunes IF branches before binding', () => {
    const compiled = compileFormula('IF(TRUE, 1+2*3, A1)')
    expect(compiled.optimizedAst).toEqual({ kind: 'NumberLiteral', value: 7 })
    expect(compiled.deps).toEqual([])
    expect(compiled.jsPlan).toEqual([{ opcode: 'push-number', value: 7 }, { opcode: 'return' }])
  })

  it('flattens concat calls in the optimized AST', () => {
    const compiled = compileFormula('CONCAT(A1, CONCAT(B1, C1))')
    expect(compiled.optimizedAst).toMatchObject({
      kind: 'CallExpr',
      callee: 'CONCAT',
      args: [
        { kind: 'CellRef', ref: 'A1' },
        { kind: 'CellRef', ref: 'B1' },
        { kind: 'CellRef', ref: 'C1' },
      ],
    })
  })
})
