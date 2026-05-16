import { describe, expect, it } from 'vitest'
import type { EngineCellMutationRef } from '@bilig/core'
import {
  applyWorkPaperMatrixContents,
  applyWorkPaperSerializedMatrix,
  type WorkPaperCellMutationApplyOptions,
} from '../work-paper-matrix-application.js'
import type { RawCellContent, WorkPaperCellAddress } from '../work-paper-types.js'

describe('work-paper matrix application', () => {
  it('applies serialized clipboard formulas through raw content translation', () => {
    const rawWrites: Array<{ address: WorkPaperCellAddress; content: RawCellContent }> = []
    let flushCount = 0

    applyWorkPaperSerializedMatrix({
      targetLeftCorner: { sheet: 1, row: 2, col: 1 },
      sourceAnchor: { sheet: 1, row: 0, col: 0 },
      serialized: [['=A1', 4]],
      flushPendingBatchOps: () => {
        flushCount += 1
      },
      applyRawContent: (address, content) => {
        rawWrites.push({ address, content })
      },
      applyCellMutationRefs: () => {
        throw new Error('Formula clipboard content should use raw writes')
      },
      rewriteFormulaForStorage: (formula) => formula,
    })

    expect(flushCount).toBe(1)
    expect(rawWrites).toEqual([
      { address: { sheet: 1, row: 2, col: 1 }, content: '=B3' },
      { address: { sheet: 1, row: 2, col: 2 }, content: 4 },
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

    expect(applied.map((entry) => entry.options.skipDimensionUpdate)).toEqual([undefined, undefined])
    expect(dimensionUpdateCount).toBe(0)
  })
})
