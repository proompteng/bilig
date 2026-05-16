import { parseFormula } from '@bilig/formula'
import { MAX_ROWS, ValueTag } from '@bilig/protocol'
import { describe, expect, it } from 'vitest'
import type { RuntimeDirectCriteriaDescriptor } from '../engine/runtime-state.js'
import {
  appendDirectCriteriaResultTransform,
  callName,
  flattenCriteriaProduct,
  resolveDirectCriteriaRange,
  staticCellValue,
} from '../engine/services/formula-binding-direct-criteria-ast.js'

describe('formula binding direct criteria AST helpers', () => {
  it('normalizes static literal cells without treating dynamic nodes as literals', () => {
    expect(staticCellValue(parseFormula('-12'))).toEqual({ tag: ValueTag.Number, value: -12 })
    expect(staticCellValue(parseFormula('"east"'))).toEqual({ tag: ValueTag.String, value: 'east', stringId: 0 })
    expect(staticCellValue(parseFormula('TRUE'))).toEqual({ tag: ValueTag.Boolean, value: true })
    expect(staticCellValue(parseFormula('A1'))).toBeUndefined()
  })

  it('flattens multiplication criteria products while preserving factor order', () => {
    expect(flattenCriteriaProduct(parseFormula('(A1:A3=1)*(B1:B3="x")*(C1:C3>0)')).map((node) => node.kind)).toEqual([
      'BinaryExpr',
      'BinaryExpr',
      'BinaryExpr',
    ])
  })

  it('resolves single-column cell and whole-column criteria ranges', () => {
    expect(resolveDirectCriteriaRange(parseFormula('Data!B2:B5'), 'Sheet1')).toEqual({
      sheetName: 'Data',
      rowStart: 1,
      rowEnd: 4,
      col: 1,
      length: 4,
    })
    expect(resolveDirectCriteriaRange(parseFormula('C:C'), 'Sheet1')).toEqual({
      sheetName: 'Sheet1',
      rowStart: 0,
      rowEnd: MAX_ROWS - 1,
      col: 2,
      length: MAX_ROWS,
    })
    expect(resolveDirectCriteriaRange(parseFormula('A1:B3'), 'Sheet1')).toBeUndefined()
  })

  it('normalizes call names and appends transforms without mutating the descriptor', () => {
    const descriptor = {
      aggregateKind: 'sum',
      criteriaPairs: [],
    } satisfies RuntimeDirectCriteriaDescriptor
    const transformed = appendDirectCriteriaResultTransform(descriptor, { kind: 'round', digits: { tag: ValueTag.Number, value: 2 } })

    expect(callName(parseFormula(' sumifs(A1:A3,B1:B3,1)'))).toBe('SUMIFS')
    expect(descriptor.resultTransforms).toBeUndefined()
    expect(transformed.resultTransforms).toEqual([{ kind: 'round', digits: { tag: ValueTag.Number, value: 2 } }])
  })
})
