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
})
