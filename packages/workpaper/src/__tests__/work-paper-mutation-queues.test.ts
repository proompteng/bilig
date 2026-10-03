import { describe, expect, it } from 'vitest'
import type { EngineCellMutationRef } from '@bilig/core'
import { WorkPaperMutationQueues } from '../work-paper-mutation-queues.js'
import type { WorkPaperCellMutationApplyOptions } from '../work-paper-cell-mutation-refs.js'

function createRecordedQueues(): {
  readonly queues: WorkPaperMutationQueues
  readonly applied: {
    readonly refs: readonly EngineCellMutationRef[]
    readonly options: WorkPaperCellMutationApplyOptions
  }[]
  readonly dimensionUpdates: readonly EngineCellMutationRef[][]
} {
  const applied: {
    readonly refs: readonly EngineCellMutationRef[]
    readonly options: WorkPaperCellMutationApplyOptions
  }[] = []
  const dimensionUpdates: EngineCellMutationRef[][] = []
  const queues = new WorkPaperMutationQueues({
    applyCellMutationsAtWithOptions: (refs, options) => {
      applied.push({ refs, options })
    },
    updateSheetDimensionsAfterCellMutationRefs: (refs) => {
      dimensionUpdates.push([...refs])
    },
  })
  return { queues, applied, dimensionUpdates }
}

describe('work paper mutation queues', () => {
  it('flushes deferred literal batch mutations with known potential new cells', () => {
    const { queues, applied, dimensionUpdates } = createRecordedQueues()

    expect(
      queues.enqueueDeferredBatchLiteral({
        sheetId: 1,
        row: 0,
        col: 0,
        content: 10,
        cellIndex: undefined,
      }),
    ).toBe(true)
    expect(
      queues.enqueueDeferredBatchLiteral({
        sheetId: 1,
        row: 0,
        col: 1,
        content: '=A1',
        cellIndex: undefined,
      }),
    ).toBe(false)
    expect(queues.hasPendingBatchOps()).toBe(true)

    queues.flushPendingBatchOps()

    expect(queues.hasPendingBatchOps()).toBe(false)
    expect(applied).toEqual([
      {
        refs: [
          {
            sheetId: 1,
            mutation: { kind: 'setCellValue', row: 0, col: 0, value: 10 },
          },
        ],
        options: {
          captureUndo: true,
          potentialNewCells: 1,
          source: 'local',
          returnUndoOps: false,
          reuseRefs: true,
        },
      },
    ])
    expect(dimensionUpdates).toHaveLength(1)
  })

  it('skips dimension updates for known existing suspended literal mutations', () => {
    const { queues, dimensionUpdates } = createRecordedQueues()
    const refs: EngineCellMutationRef[] = [
      {
        sheetId: 2,
        cellIndex: 5,
        mutation: { kind: 'setCellValue', row: 1, col: 1, value: 12 },
      },
    ]

    queues.appendSuspendedCellMutationRefs(refs)
    queues.addSuspendedCellMutationPotentialNewCells(0)
    queues.flushSuspendedCellMutations()

    expect(dimensionUpdates).toEqual([])
  })

  it('updates dimensions for suspended mutations that may add cells', () => {
    const { queues, dimensionUpdates } = createRecordedQueues()
    const refs: EngineCellMutationRef[] = [
      {
        sheetId: 2,
        mutation: { kind: 'setCellValue', row: 4, col: 1, value: 12 },
      },
    ]

    queues.appendSuspendedCellMutationRefs(refs)
    queues.addSuspendedCellMutationPotentialNewCells(1)
    queues.flushSuspendedCellMutations()

    expect(dimensionUpdates).toEqual([refs])
  })

  it('flushes validated deferred literal mutations without reclassifying them', () => {
    const { queues, applied, dimensionUpdates } = createRecordedQueues()

    queues.enqueueValidatedDeferredBatchLiteral({
      sheetId: 3,
      row: 2,
      col: 1,
      content: 42,
      cellIndex: 11,
    })
    queues.enqueueValidatedDeferredBatchLiteral({
      sheetId: 3,
      row: 3,
      col: 1,
      content: null,
      cellIndex: 12,
    })

    queues.flushPendingBatchOps()

    expect(applied).toEqual([
      {
        refs: [
          {
            sheetId: 3,
            cellIndex: 11,
            mutation: { kind: 'setCellValue', row: 2, col: 1, value: 42 },
          },
          {
            sheetId: 3,
            cellIndex: 12,
            mutation: { kind: 'clearCell', row: 3, col: 1 },
          },
        ],
        options: {
          captureUndo: true,
          potentialNewCells: 0,
          source: 'local',
          returnUndoOps: false,
          reuseRefs: true,
        },
      },
    ])
    expect(dimensionUpdates).toHaveLength(1)
  })

  it('flushes validated suspended literal mutations without reclassifying them', () => {
    const { queues, applied, dimensionUpdates } = createRecordedQueues()

    queues.enqueueValidatedSuspendedLiteral({
      sheetId: 4,
      row: 5,
      col: 2,
      content: 21,
      cellIndex: 20,
    })
    queues.enqueueValidatedSuspendedLiteral({
      sheetId: 4,
      row: 6,
      col: 2,
      content: null,
      cellIndex: undefined,
    })

    queues.flushSuspendedCellMutations()

    expect(applied).toEqual([
      {
        refs: [
          {
            sheetId: 4,
            cellIndex: 20,
            mutation: { kind: 'setCellValue', row: 5, col: 2, value: 21 },
          },
          {
            sheetId: 4,
            mutation: { kind: 'clearCell', row: 6, col: 2 },
          },
        ],
        options: {
          captureUndo: true,
          potentialNewCells: 0,
          source: 'local',
          returnUndoOps: false,
          reuseRefs: true,
        },
      },
    ])
    expect(dimensionUpdates).toHaveLength(1)
  })
})
