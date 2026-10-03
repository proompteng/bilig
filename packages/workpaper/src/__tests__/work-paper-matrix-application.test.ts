import { describe, expect, it } from 'vitest'
import type { EngineCellMutationRef } from '@bilig/core'
import {
  applyWorkPaperMatrixContents,
  applyWorkPaperSerializedMatrix,
  type WorkPaperCellMutationApplyOptions,
} from '../work-paper-matrix-application.js'

describe('work-paper matrix application', () => {
  it('applies serialized clipboard formulas through one translated matrix mutation batch', () => {
    const applied: Array<{ refs: readonly EngineCellMutationRef[]; options: WorkPaperCellMutationApplyOptions }> = []
    let flushCount = 0

    applyWorkPaperSerializedMatrix({
      targetLeftCorner: { sheet: 1, row: 2, col: 1 },
      sourceAnchor: { sheet: 1, row: 0, col: 0 },
      serialized: [['=A1', 4]],
      flushPendingBatchOps: () => {
        flushCount += 1
      },
      applyRawContent: (address, content) => {
        throw new Error(`Serialized formula paste should not write ${address.row}:${address.col} cell-by-cell: ${String(content)}`)
      },
      applyCellMutationRefs: (refs, options) => {
        applied.push({ refs, options })
      },
      rewriteFormulaForStorage: (formula) => formula,
    })

    expect(flushCount).toBe(1)
    expect(applied).toHaveLength(1)
    expect(applied[0]?.options).toEqual({
      captureUndo: true,
      potentialNewCells: 2,
      source: 'local',
      returnUndoOps: false,
      reuseRefs: true,
    })
    expect(applied[0]?.refs).toEqual([
      { sheetId: 1, mutation: { kind: 'setCellValue', row: 2, col: 2, value: 4 } },
      { sheetId: 1, mutation: { kind: 'setCellFormula', row: 2, col: 1, formula: 'B3' } },
    ])
  })

  it('applies matrix contents in leading, formula, and trailing phases', () => {
    const applied: Array<{ refs: readonly EngineCellMutationRef[]; options: WorkPaperCellMutationApplyOptions }> = []
    let flushCount = 0

    applyWorkPaperMatrixContents({
      address: { sheet: 1, row: 0, col: 0 },
      content: [
        [1, '=A2+A3', 'top'],
        [2, 20, 'middle'],
        [3, 30, '=A1+B1'],
      ],
      options: { captureUndo: false },
      flushPendingBatchOps: () => {
        flushCount += 1
      },
      applyCellMutationRefs: (refs, options) => {
        applied.push({ refs, options })
      },
      rewriteFormulaForStorage: (formula) => formula,
    })

    expect(flushCount).toBe(1)
    expect(applied).toHaveLength(3)
    expect(applied.map((entry) => entry.options)).toEqual([
      { captureUndo: false, potentialNewCells: 5, source: 'restore', returnUndoOps: false, reuseRefs: true },
      { captureUndo: false, potentialNewCells: 2, source: 'restore', returnUndoOps: false, reuseRefs: true },
      { captureUndo: false, potentialNewCells: 2, source: 'restore', returnUndoOps: false, reuseRefs: true },
    ])
    expect(applied.map((entry) => entry.refs.map((ref) => ref.mutation.kind))).toEqual([
      ['setCellValue', 'setCellValue', 'setCellValue', 'setCellValue', 'setCellValue'],
      ['setCellFormula', 'setCellFormula'],
      ['setCellValue', 'setCellValue'],
    ])
  })

  it('applies ordered literal and formula matrices in one pass when no trailing literals are deferred', () => {
    const applied: Array<{ refs: readonly EngineCellMutationRef[]; options: WorkPaperCellMutationApplyOptions }> = []

    applyWorkPaperMatrixContents({
      address: { sheet: 1, row: 0, col: 0 },
      content: [
        [1, 2, '=A1+B1'],
        [3, 4, '=A2+B2'],
      ],
      flushPendingBatchOps: () => {},
      applyCellMutationRefs: (refs, options) => {
        applied.push({ refs, options })
      },
      rewriteFormulaForStorage: (formula) => formula,
    })

    expect(applied).toHaveLength(1)
    expect(applied[0]?.options).toEqual({
      potentialNewCells: 6,
      source: 'local',
      returnUndoOps: false,
      reuseRefs: true,
    })
    expect(applied[0]?.refs.map((ref) => ref.mutation.kind)).toEqual([
      'setCellValue',
      'setCellValue',
      'setCellValue',
      'setCellValue',
      'setCellFormula',
      'setCellFormula',
    ])
  })

  it('keeps large fresh numeric plus formula-column matrices in one operation', () => {
    const applied: Array<{ refs: readonly EngineCellMutationRef[]; options: WorkPaperCellMutationApplyOptions }> = []

    applyWorkPaperMatrixContents({
      address: { sheet: 1, row: 10, col: 0 },
      content: Array.from({ length: 16 }, (_, row) => [row + 1, row + 2, `=A${row + 11}+B${row + 11}`]),
      flushPendingBatchOps: () => {},
      applyCellMutationRefs: (refs, options) => {
        applied.push({ refs, options })
      },
      rewriteFormulaForStorage: (formula) => formula,
    })

    expect(applied).toHaveLength(1)
    expect(applied[0]?.options.potentialNewCells).toBe(48)
    expect(applied[0]?.refs).toHaveLength(48)
    expect(applied[0]?.refs.slice(0, 32).every((ref) => ref.mutation.kind === 'setCellValue')).toBe(true)
    expect(applied[0]?.refs.slice(32).every((ref) => ref.mutation.kind === 'setCellFormula')).toBe(true)
  })

  it('updates dimensions once after phased formula matrix writes', () => {
    const applied: Array<{ refs: readonly EngineCellMutationRef[]; options: WorkPaperCellMutationApplyOptions }> = []
    const dimensionUpdates: Array<readonly EngineCellMutationRef[]> = []

    applyWorkPaperMatrixContents({
      address: { sheet: 1, row: 0, col: 0 },
      content: [
        [1, '=A2+A3', 'top'],
        [2, 20, 'middle'],
        [3, 30, '=A1+B1'],
      ],
      flushPendingBatchOps: () => {},
      applyCellMutationRefs: (refs, options) => {
        applied.push({ refs, options })
      },
      isEvaluationSuspended: () => false,
      rewriteFormulaForStorage: (formula) => formula,
      updateSheetDimensionsAfterCellMutationRefs: (refs) => {
        dimensionUpdates.push(refs)
      },
    })

    expect(applied.map((entry) => entry.options.skipDimensionUpdate)).toEqual([true, true, true])
    expect(dimensionUpdates).toHaveLength(1)
    expect(dimensionUpdates[0]?.map((ref) => ref.mutation.kind)).toEqual([
      'setCellValue',
      'setCellValue',
      'setCellValue',
      'setCellValue',
      'setCellValue',
      'setCellFormula',
      'setCellFormula',
      'setCellValue',
      'setCellValue',
    ])
  })

  it('leaves suspended local formula matrix dimensions to queued mutation flush', () => {
    const applied: Array<{ refs: readonly EngineCellMutationRef[]; options: WorkPaperCellMutationApplyOptions }> = []
    let dimensionUpdateCount = 0

    applyWorkPaperMatrixContents({
      address: { sheet: 1, row: 0, col: 0 },
      content: [[1, '=A1']],
      flushPendingBatchOps: () => {},
      applyCellMutationRefs: (refs, options) => {
        applied.push({ refs, options })
      },
      isEvaluationSuspended: () => true,
      rewriteFormulaForStorage: (formula) => formula,
      updateSheetDimensionsAfterCellMutationRefs: () => {
        dimensionUpdateCount += 1
      },
    })

    expect(applied.map((entry) => entry.options.skipDimensionUpdate)).toEqual([undefined])
    expect(dimensionUpdateCount).toBe(0)
  })
})
