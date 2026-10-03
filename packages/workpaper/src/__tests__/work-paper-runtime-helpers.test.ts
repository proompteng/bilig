import { describe, expect, it } from 'vitest'
import { ErrorCode, ValueTag } from '@bilig/protocol'
import { parseFormula } from '@bilig/formula'
import { WorkPaperInvalidArgumentsError } from '../work-paper-errors.js'
import {
  assertRange,
  cloneCellValue,
  collectFormulaNameRefs,
  emptyValue,
  errorValue,
  formulaHasRelativeReferences,
  isCellRange,
  isDeferredBatchLiteralContent,
  isFormulaContent,
  isWorkPaperSheetMatrix,
  makeInternalScopedName,
  makeNamedExpressionKey,
  matrixContainsFormulaContent,
  matrixValuesEqual,
  scalarFromResult,
  scalarValueFromLiteral,
  stripLeadingEquals,
  transformFormulaNode,
  tryEvaluateSimpleNamedExpression,
  tryEvaluateSimpleScalarFormulaBody,
  tryReadSimpleScalarFormulaBody,
  valuesEqual,
} from '../work-paper-runtime-helpers.js'
import type { WorkPaperCellRange } from '../work-paper-types.js'

describe('work paper runtime helpers', () => {
  it('converts literals, errors, and array results to scalar values', () => {
    expect(emptyValue()).toEqual({ tag: ValueTag.Empty })
    expect(errorValue(ErrorCode.Value)).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(scalarValueFromLiteral(null)).toEqual({ tag: ValueTag.Empty })
    expect(scalarValueFromLiteral(3)).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(scalarValueFromLiteral(true)).toEqual({ tag: ValueTag.Boolean, value: true })
    expect(scalarValueFromLiteral('text')).toEqual({ tag: ValueTag.String, value: 'text', stringId: 0 })
    expect(scalarFromResult({ kind: 'array', values: [{ tag: ValueTag.Number, value: 4 }], rows: 1, cols: 1 })).toEqual({
      tag: ValueTag.Number,
      value: 4,
    })
  })

  it('evaluates simple scalar formula bodies without invoking the engine', () => {
    expect(tryReadSimpleScalarFormulaBody(' = -0 ')).toBe('-0')
    expect(tryReadSimpleScalarFormulaBody('plain')).toBeUndefined()
    expect(tryEvaluateSimpleScalarFormulaBody('-0')).toEqual({ tag: ValueTag.Number, value: 0 })
    expect(tryEvaluateSimpleScalarFormulaBody('TRUE')).toEqual({ tag: ValueTag.Boolean, value: true })
    expect(tryEvaluateSimpleScalarFormulaBody('"a""b"')).toEqual({ tag: ValueTag.String, value: 'a"b', stringId: 0 })
    expect(tryEvaluateSimpleScalarFormulaBody('SUM(1,2)')).toBeUndefined()
    expect(tryEvaluateSimpleNamedExpression(' plain ')).toEqual({ tag: ValueTag.String, value: ' plain ', stringId: 0 })
    expect(tryEvaluateSimpleNamedExpression('=FALSE')).toEqual({ tag: ValueTag.Boolean, value: false })
  })

  it('compares scalar and matrix values by value shape', () => {
    const number = { tag: ValueTag.Number, value: 1 } as const
    const sameNumber = { tag: ValueTag.Number, value: 1 } as const
    const differentNumber = { tag: ValueTag.Number, value: 2 } as const

    expect(valuesEqual(number, sameNumber)).toBe(true)
    expect(valuesEqual(number, differentNumber)).toBe(false)
    expect(matrixValuesEqual([[number]], [[sameNumber]])).toBe(true)
    expect(matrixValuesEqual([[number]], [[differentNumber]])).toBe(false)
    expect(matrixValuesEqual(number, [[sameNumber]])).toBe(false)
  })

  it('normalizes names and detects formula-like matrix content', () => {
    expect(makeNamedExpressionKey(' rate ', 4)).toBe('4:RATE')
    expect(makeNamedExpressionKey(' rate ')).toBe('workbook:RATE')
    expect(makeInternalScopedName(4, 'rate')).toBe('__BILIG_WORKPAPER_SCOPE_4_RATE')
    expect(isFormulaContent(' =A1 ')).toBe(true)
    expect(isFormulaContent('=A1')).toBe(true)
    expect(isFormulaContent(' label = A1')).toBe(false)
    expect(isFormulaContent('A1')).toBe(false)
    expect(isDeferredBatchLiteralContent('=A1')).toBe(true)
    expect(isWorkPaperSheetMatrix([[null]])).toBe(true)
    expect(matrixContainsFormulaContent([[null, '=A1']])).toBe(true)
    expect(stripLeadingEquals('=A1 ')).toBe('A1')
    expect(stripLeadingEquals(' = A1 ')).toBe(' A1')
    expect(stripLeadingEquals(' value ')).toBe('value')
  })

  it('validates same-sheet ranges and range-like values', () => {
    const range: WorkPaperCellRange = {
      start: { sheet: 1, row: 0, col: 0 },
      end: { sheet: 1, row: 1, col: 1 },
    }

    expect(isCellRange(range)).toBe(true)
    expect(isCellRange({ sheet: 1, row: 0, col: 0 })).toBe(false)
    expect(() => assertRange(range)).not.toThrow()
    expect(() =>
      assertRange({
        start: { sheet: 1, row: 0, col: 0 },
        end: { sheet: 2, row: 0, col: 0 },
      }),
    ).toThrow(WorkPaperInvalidArgumentsError)
  })

  it('clones values and traverses formula ASTs', () => {
    const stringValue = { tag: ValueTag.String, value: 'hello', stringId: 8 } as const
    expect(cloneCellValue(stringValue)).toEqual(stringValue)
    expect(cloneCellValue(stringValue)).not.toBe(stringValue)

    const names = new Set<string>()
    collectFormulaNameRefs(parseFormula('SUM(Rate, Tax)'), names)
    expect([...names].toSorted()).toEqual(['Rate', 'Tax'])

    const transformed = transformFormulaNode(parseFormula('SUM(Rate)'), (node) => {
      return node.kind === 'NameRef' ? { ...node, name: 'Amount' } : node
    })
    const transformedNames = new Set<string>()
    collectFormulaNameRefs(transformed, transformedNames)
    expect([...transformedNames]).toEqual(['Amount'])

    expect(formulaHasRelativeReferences(parseFormula('A1'))).toBe(true)
    expect(formulaHasRelativeReferences(parseFormula('$A$1'))).toBe(false)
    expect(formulaHasRelativeReferences(parseFormula('SUM($A$1:$B$2)'))).toBe(false)
  })
})
