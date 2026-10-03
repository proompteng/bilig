import { describe, expect, it } from 'vitest'
import { buildMatrixMutationPlan } from '../matrix-mutation-plan.js'

describe('buildMatrixMutationPlan', () => {
  it('emits clear and literal ops before formulas for mixed sheet imports', () => {
    const plan = buildMatrixMutationPlan({
      target: { sheet: 1, row: 0, col: 0 },
      content: [
        ['=A2+A3', 10, null],
        [20, '=B1*2', 'done'],
      ],
      rewriteFormula: (formula) => formula.slice(1),
    })

    expect(plan.potentialNewCells).toBe(5)
    expect(plan.canApplyFreshNumericAggregateMatrixInOnePass).toBe(false)
    expect(plan.leadingPotentialNewCells).toBe(2)
    expect(plan.formulaPotentialNewCells).toBe(2)
    expect(plan.trailingLiteralPotentialNewCells).toBe(1)
    expect(plan.refs).toEqual([
      { sheetId: 1, mutation: { kind: 'setCellValue', row: 0, col: 1, value: 10 } },
      { sheetId: 1, mutation: { kind: 'clearCell', row: 0, col: 2 } },
      { sheetId: 1, mutation: { kind: 'setCellValue', row: 1, col: 2, value: 'done' } },
      { sheetId: 1, mutation: { kind: 'setCellFormula', row: 0, col: 0, formula: 'A2+A3' } },
      { sheetId: 1, mutation: { kind: 'setCellFormula', row: 1, col: 1, formula: 'B1*2' } },
      { sheetId: 1, mutation: { kind: 'setCellValue', row: 1, col: 0, value: 20 } },
    ])
  })

  it('defers literals below the first formula in each column without scanning every formula', () => {
    const plan = buildMatrixMutationPlan({
      target: { sheet: 1, row: 0, col: 0 },
      content: [
        [1, '=A2+A3', 'top'],
        [2, 20, 'middle'],
        [3, 30, '=A1+B1'],
        [4, 40, 'bottom'],
      ],
      rewriteFormula: (formula) => formula.slice(1),
    })

    expect(plan.potentialNewCells).toBe(12)
    expect(plan.leadingPotentialNewCells).toBe(6)
    expect(plan.formulaPotentialNewCells).toBe(2)
    expect(plan.trailingLiteralPotentialNewCells).toBe(4)
    expect(plan.leadingRefs).toEqual([
      { sheetId: 1, mutation: { kind: 'setCellValue', row: 0, col: 0, value: 1 } },
      { sheetId: 1, mutation: { kind: 'setCellValue', row: 0, col: 2, value: 'top' } },
      { sheetId: 1, mutation: { kind: 'setCellValue', row: 1, col: 0, value: 2 } },
      { sheetId: 1, mutation: { kind: 'setCellValue', row: 1, col: 2, value: 'middle' } },
      { sheetId: 1, mutation: { kind: 'setCellValue', row: 2, col: 0, value: 3 } },
      { sheetId: 1, mutation: { kind: 'setCellValue', row: 3, col: 0, value: 4 } },
    ])
    expect(plan.formulaRefs).toEqual([
      { sheetId: 1, mutation: { kind: 'setCellFormula', row: 0, col: 1, formula: 'A2+A3' } },
      { sheetId: 1, mutation: { kind: 'setCellFormula', row: 2, col: 2, formula: 'A1+B1' } },
    ])
    expect(plan.trailingLiteralRefs).toEqual([
      { sheetId: 1, mutation: { kind: 'setCellValue', row: 1, col: 1, value: 20 } },
      { sheetId: 1, mutation: { kind: 'setCellValue', row: 2, col: 1, value: 30 } },
      { sheetId: 1, mutation: { kind: 'setCellValue', row: 3, col: 1, value: 40 } },
      { sheetId: 1, mutation: { kind: 'setCellValue', row: 3, col: 2, value: 'bottom' } },
    ])
  })

  it('defers explicit literal addresses without formatting every planned cell', () => {
    const plan = buildMatrixMutationPlan({
      target: { sheet: 1, row: 0, col: 0 },
      content: [
        [1, 2],
        [3, '=A1+B1'],
      ],
      deferLiteralAddresses: new Set(['B1']),
      rewriteFormula: (formula) => formula.slice(1),
    })

    expect(plan.leadingRefs).toEqual([
      { sheetId: 1, mutation: { kind: 'setCellValue', row: 0, col: 0, value: 1 } },
      { sheetId: 1, mutation: { kind: 'setCellValue', row: 1, col: 0, value: 3 } },
    ])
    expect(plan.formulaRefs).toEqual([{ sheetId: 1, mutation: { kind: 'setCellFormula', row: 1, col: 1, formula: 'A1+B1' } }])
    expect(plan.trailingLiteralRefs).toEqual([{ sheetId: 1, mutation: { kind: 'setCellValue', row: 0, col: 1, value: 2 } }])
  })

  it('can skip the combined ref list for phased matrix application', () => {
    const plan = buildMatrixMutationPlan({
      target: { sheet: 1, row: 0, col: 0 },
      content: [
        [1, '=A1'],
        [2, 3],
      ],
      includeCombinedRefs: false,
      rewriteFormula: (formula) => formula.slice(1),
    })

    expect(plan.refCount).toBe(4)
    expect(plan.refs).toEqual([])
    expect(plan.leadingRefs).toHaveLength(2)
    expect(plan.formulaRefs).toHaveLength(1)
    expect(plan.trailingLiteralRefs).toHaveLength(1)
  })

  it('tracks matrix dimension impact across ragged rows, formulas, and skipped blanks', () => {
    const plan = buildMatrixMutationPlan({
      target: { sheet: 7, row: 10, col: 3 },
      content: [[null, 1], ['=SEQUENCE(2)', null, 3], [], [null]],
      skipNulls: true,
      rewriteFormula: (formula) => formula.slice(1),
    })

    expect(plan.dimensionImpact).toEqual({
      hasDynamicFormula: true,
      maxClearCol: -1,
      maxClearRow: -1,
      maxSetCol: 5,
      maxSetRow: 11,
      sheetId: 7,
    })
  })

  it('detects dense fresh numeric plus formula-column matrices during planning', () => {
    const plan = buildMatrixMutationPlan({
      target: { sheet: 3, row: 20, col: 1 },
      content: Array.from({ length: 16 }, (_, row) => [row + 1, row + 2, `=SUM(B${row + 21}:C${row + 21})`]),
      includeCombinedRefs: false,
      rewriteFormula: (formula) => formula.slice(1),
    })

    expect(plan.canApplyFreshNumericAggregateMatrixInOnePass).toBe(true)
    expect(plan.leadingRefs).toHaveLength(32)
    expect(plan.formulaRefs).toHaveLength(16)
    expect(plan.trailingLiteralRefs).toHaveLength(0)
    expect(plan.refs).toEqual([])
  })

  it('rejects ragged numeric plus formula matrices as fresh aggregate candidates', () => {
    const plan = buildMatrixMutationPlan({
      target: { sheet: 3, row: 20, col: 1 },
      content: [
        [1, 2, '=SUM(B21:C21)'],
        [3, 4, 5, '=SUM(B22:D22)'],
      ],
      rewriteFormula: (formula) => formula.slice(1),
    })

    expect(plan.canApplyFreshNumericAggregateMatrixInOnePass).toBe(false)
  })
})
