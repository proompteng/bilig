import { describe, expect, it } from 'vitest'
import { parseFormula } from '../index.js'
import {
  getNativeAxisAggregateCode,
  getNativeGroupedArrayKind,
  getNativeRunningFoldCode,
  isCellRangeNode,
  isCellVectorNode,
  isNativeMakearraySumLambda,
  isWasmSafeBuiltinArgs,
  isWasmSafeBuiltinArity,
} from '../binder-wasm-rules.js'
import type { FormulaNode } from '../ast.js'

describe('binder wasm rules', () => {
  it('detects native lambda patterns and cell range shapes', () => {
    expect(getNativeAxisAggregateCode(parseFormula('LAMBDA(x,SUM(x))'))).toBe(1)
    expect(getNativeAxisAggregateCode(parseFormula('LAMBDA(x,AVERAGE(x))'))).toBe(2)
    expect(getNativeAxisAggregateCode(parseFormula('LAMBDA(x,SUM(y))'))).toBeNull()

    expect(getNativeRunningFoldCode(parseFormula('LAMBDA(acc,x,acc+x)'))).toBe(1)
    expect(getNativeRunningFoldCode(parseFormula('LAMBDA(acc,x,x*acc)'))).toBe(2)
    expect(getNativeRunningFoldCode(parseFormula('LAMBDA(acc,x,acc-x)'))).toBeNull()

    expect(isNativeMakearraySumLambda(parseFormula('LAMBDA(r,c,r+c)'))).toBe(true)
    expect(isNativeMakearraySumLambda(parseFormula('LAMBDA(r,c,r-c)'))).toBe(false)

    expect(isCellRangeNode(parseFormula('A1:B3'))).toBe(true)
    expect(isCellRangeNode(parseFormula('A:B'))).toBe(false)
    expect(isCellVectorNode(parseFormula('A1:A4'))).toBe(true)
    expect(isCellVectorNode(parseFormula('A1:B4'))).toBe(false)
  })

  it('detects canonical grouped-array native shapes', () => {
    expect(getNativeGroupedArrayKind(parseFormula('GROUPBY(A1:A4,B1:B4,SUM,3,1)'))).toBe('groupby-sum-canonical')
    expect(getNativeGroupedArrayKind(parseFormula('GROUPBY(A1:B4,B1:B4,SUM,3,1)'))).toBeNull()
    expect(getNativeGroupedArrayKind(parseFormula('GROUPBY(A1:A4,B1:B4,AVERAGE,3,1)'))).toBeNull()
    expect(getNativeGroupedArrayKind(parseFormula('PIVOTBY(A1:A4,B1:B4,C1:C4,SUM,3,1,0,1)'))).toBe('pivotby-sum-canonical')
    expect(getNativeGroupedArrayKind(parseFormula('PIVOTBY(A1:A4,B1:B4,C1:D4,SUM,3,1,0,1)'))).toBeNull()
    expect(getNativeGroupedArrayKind(parseFormula('SUM(A1:A4)'))).toBeNull()
  })

  it('keeps representative arity and wasm-argument policies stable', () => {
    expect(isWasmSafeBuiltinArity('COUNTIFS', 4)).toBe(true)
    expect(isWasmSafeBuiltinArity('COUNTIFS', 3)).toBe(false)
    expect(isWasmSafeBuiltinArity('LOOKUP', 2)).toBe(true)
    expect(isWasmSafeBuiltinArity('LOOKUP', 4)).toBe(false)
    expect(isWasmSafeBuiltinArity('NA', 0)).toBe(true)
    expect(isWasmSafeBuiltinArity('NA', 1)).toBe(false)

    const deps = {
      isWasmSafe(node: FormulaNode, allowRange = false): boolean {
        if (
          node.kind === 'NumberLiteral' ||
          node.kind === 'BooleanLiteral' ||
          node.kind === 'StringLiteral' ||
          node.kind === 'ErrorLiteral'
        ) {
          return true
        }
        return node.kind === 'RangeRef' ? allowRange : false
      },
    }

    expect(isWasmSafeBuiltinArgs('COUNTIFS', parseFormula('COUNTIFS(A1:A4,1,B1:B4,2)').args, deps)).toBe(true)
    expect(isWasmSafeBuiltinArgs('LOOKUP', parseFormula('LOOKUP(1,A1:A4,B1:B4)').args, deps)).toBe(true)
    expect(isWasmSafeBuiltinArgs('LOOKUP', parseFormula('LOOKUP(A1:A2,A1:A4)').args, deps)).toBe(false)
    expect(isWasmSafeBuiltinArgs('TEXTJOIN', parseFormula('TEXTJOIN(",",TRUE,A1:A4)').args, deps)).toBe(true)
  })

  it('covers promoted wasm argument-shape policies', () => {
    const deps = {
      isWasmSafe(node: FormulaNode, allowRange = false): boolean {
        switch (node.kind) {
          case 'NumberLiteral':
          case 'BooleanLiteral':
          case 'StringLiteral':
          case 'ErrorLiteral':
          case 'CellRef':
            return true
          case 'RangeRef':
            return allowRange
          case 'BinaryExpr':
            return deps.isWasmSafe(node.left, true) && deps.isWasmSafe(node.right, true)
          case 'CallExpr':
            return ['SEQUENCE', 'TRUE', 'FALSE', 'NA'].includes(node.callee.toUpperCase()) && node.args.every((arg) => deps.isWasmSafe(arg))
          case 'ColumnRef':
          case 'InvokeExpr':
          case 'NameRef':
          case 'RowRef':
          case 'SpillRef':
          case 'StructuredRef':
          case 'UnaryExpr':
            return false
          default:
            return false
        }
      },
    }
    const expectArgs = (formula: string, expected: boolean): void => {
      const parsed = parseFormula(formula)
      if (parsed.kind !== 'CallExpr') {
        throw new Error(`Expected call expression for ${formula}`)
      }
      expect(isWasmSafeBuiltinArgs(parsed.callee.toUpperCase(), parsed.args, deps), formula).toBe(expected)
    }

    expectArgs('SUM(SEQUENCE(2))', true)
    expectArgs('CHOOSE(1,A1:A2,SEQUENCE(2))', true)
    expectArgs('IF(A1,1,2)', true)
    expectArgs('IF(A1:A2,1,2)', false)
    expectArgs('IF(A1,"text",2)', false)
    expectArgs('IFS(A1,1,FALSE(),2)', true)
    expectArgs('IFS(A1:A2,1,FALSE(),2)', false)
    expectArgs('COUNTIF(A1:A4,1)', true)
    expectArgs('COUNTIF(A1,1)', false)
    expectArgs('DCOUNT(A1:B4,1,A1:B2)', true)
    expectArgs('T.TEST(A1:A4,B1:B4,1,2)', true)
    expectArgs('SUMIF(A1:A4,">0",B1:B4)', true)
    expectArgs('SUMIFS(C1:C4,A1:A4,">0")', true)
    expectArgs('SUMPRODUCT(A1:A4,B1:B4)', true)
    expectArgs('MATCH(1,A1:A4,0)', true)
    expectArgs('CORREL(A1:A4,B1:B4)', true)
    expectArgs('FREQUENCY(A1:A4,B1:B4)', true)
    expectArgs('SMALL(A1:A4,1)', true)
    expectArgs('PERCENTRANK(A1:A4,1,3)', true)
    expectArgs('RANK(1,A1:A4,0)', true)
    expectArgs('FORECAST(1,A1:A4,B1:B4)', true)
    expectArgs('TREND(A1:A4,B1:B4,1,TRUE())', true)
    expectArgs('DAYS(A1,B1)', true)
    expectArgs('DAYS(A1:A2,B1)', false)
    expectArgs('EXPAND(A1:A2,2,2,0)', true)
    expectArgs('TEXT(A1,"0")', true)
    expectArgs('PHONETIC(A1:A2)', true)
    expectArgs('TEXTJOIN(",",TRUE(),A1:A4)', true)
    expectArgs('REPLACE("abc",1,1,"d")', true)
    expectArgs('TAKE(A1:B4,2)', true)
    expectArgs('FILTER(A1:B4,C1:C4,"")', true)
    expectArgs('UNIQUE(A1:B4,FALSE(),TRUE())', true)
    expectArgs('TRIMRANGE(A1:B4,1,1)', true)
    expectArgs('PROB(A1:A4,B1:B4,1,2)', true)
    expectArgs('TRIMMEAN(A1:A4,0.2)', true)
    expectArgs('LOOKUP(1,A1:A4,B1:B4)', true)
    expectArgs('TRANSPOSE(A1:B4)', true)
    expectArgs('HSTACK(A1:B4,C1:C4)', true)
    expectArgs('AREAS(A1:B4)', true)
    expectArgs('ARRAYTOTEXT(A1:B4,1)', true)
    expectArgs('MINIFS(C1:C4,A1:A4,">0")', true)
    expectArgs('IRR(A1:A4,0.1)', true)
    expectArgs('MIRR(A1:A4,0.1,0.2)', true)
    expectArgs('XNPV(0.1,A1:A4,B1:B4)', true)
    expectArgs('XIRR(A1:A4,B1:B4,0.1)', true)
    expectArgs('SORTBY(A1:B4,C1:C4,1)', true)
  })

  it('covers representative arity policy families', () => {
    const cases: Array<readonly [string, number, boolean]> = [
      ['TODAY', 0, true],
      ['TODAY', 1, false],
      ['IF', 3, true],
      ['IFS', 4, true],
      ['IFS', 3, false],
      ['DAYS360', 3, true],
      ['DISC', 5, true],
      ['COUPDAYBS', 4, true],
      ['DURATION', 6, true],
      ['ODDFPRICE', 8, true],
      ['TBILLPRICE', 3, true],
      ['PRICE', 7, true],
      ['WORKDAY.INTL', 4, true],
      ['DAVERAGE', 3, true],
      ['ADDRESS', 5, true],
      ['TEXTBEFORE', 6, true],
      ['BESSELI', 2, true],
      ['EUROCONVERT', 5, true],
      ['VALUETOTEXT', 2, true],
      ['DOLLAR', 3, true],
      ['QUOTIENT', 2, true],
      ['BASE', 3, true],
      ['DECIMAL', 2, true],
      ['DEC2BIN', 2, true],
      ['BITAND', 3, true],
      ['BITRSHIFT', 2, true],
      ['MATCH', 3, true],
      ['REPLACEB', 4, true],
      ['TREND', 4, true],
      ['XLOOKUP', 6, true],
      ['LEFT', 2, true],
      ['MID', 3, true],
      ['FIND', 3, true],
      ['T', 0, true],
      ['PERMUTATIONA', 2, true],
      ['STANDARDIZE', 3, true],
      ['ERF', 2, true],
      ['GAMMA', 1, true],
      ['GAMMAINV', 3, true],
      ['FDIST', 3, true],
      ['CHISQ.INV', 2, true],
      ['CHISQ.DIST', 3, true],
      ['BETAINV', 5, true],
      ['BETADIST', 5, true],
      ['BETA.DIST', 6, true],
      ['F.DIST', 4, true],
      ['TINV', 2, true],
      ['TDIST', 3, true],
      ['T.TEST', 4, true],
      ['FINV', 3, true],
      ['BINOM.DIST', 4, true],
      ['POISSON.DIST', 3, true],
      ['BINOM.DIST.RANGE', 4, true],
      ['HYPGEOMDIST', 4, true],
      ['HYPGEOM.DIST', 5, true],
      ['NORMDIST', 4, true],
      ['NORM.DIST', 4, true],
      ['NORM.INV', 3, true],
      ['NORM.S.DIST', 2, true],
      ['NORM.S.INV', 1, true],
      ['LOGNORM.DIST', 4, true],
      ['LOGNORM.INV', 3, true],
      ['FV', 5, true],
      ['RATE', 6, true],
      ['IPMT', 6, true],
      ['ISPMT', 4, true],
      ['CUMIPMT', 6, true],
      ['FVSCHEDULE', 2, true],
      ['SLN', 3, true],
      ['DB', 5, true],
      ['SYD', 4, true],
      ['VDB', 7, true],
      ['SWITCH', 3, true],
      ['FILTER', 3, true],
      ['WRAPROWS', 4, true],
      ['HSTACK', 2, true],
    ]

    for (const [callee, argc, expected] of cases) {
      expect(isWasmSafeBuiltinArity(callee, argc), `${callee}/${String(argc)}`).toBe(expected)
    }
  })
})
