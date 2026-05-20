import type { EngineCellMutationRef } from '@bilig/core/headless-runtime'
import { applyQueuedWorkPaperCellMutationRefs, type WorkPaperCellMutationApplyOptions } from './work-paper-cell-mutation-refs.js'
import { tryEnqueueWorkPaperLiteralMutation } from './work-paper-literal-mutation-queue.js'
import type { RawCellContent } from './work-paper-types.js'

export interface WorkPaperMutationQueuesRuntime {
  readonly applyCellMutationsAtWithOptions: (refs: readonly EngineCellMutationRef[], options: WorkPaperCellMutationApplyOptions) => void
  readonly updateSheetDimensionsAfterCellMutationRefs: (refs: readonly EngineCellMutationRef[]) => void
}

export interface WorkPaperLiteralMutationQueueInput {
  readonly sheetId: number
  readonly row: number
  readonly col: number
  readonly content: RawCellContent
  readonly cellIndex: number | undefined
}

export class WorkPaperMutationQueues {
  private pendingBatchOps: EngineCellMutationRef[] = []
  private pendingBatchPotentialNewCells = 0
  private suspendedCellMutationRefs: EngineCellMutationRef[] = []
  private suspendedCellMutationPotentialNewCells = 0

  constructor(private readonly runtime: WorkPaperMutationQueuesRuntime) {}

  hasPendingBatchOps(): boolean {
    return this.pendingBatchOps.length > 0
  }

  appendSuspendedCellMutationRefs(refs: readonly EngineCellMutationRef[]): void {
    this.suspendedCellMutationRefs.push(...refs)
  }

  addSuspendedCellMutationPotentialNewCells(amount: number): void {
    this.suspendedCellMutationPotentialNewCells += amount
  }

  flushPendingBatchOps(): void {
    if (this.pendingBatchOps.length === 0) {
      return
    }
    const refs = this.pendingBatchOps
    const potentialNewCells = this.pendingBatchPotentialNewCells
    this.pendingBatchOps = []
    this.pendingBatchPotentialNewCells = 0
    applyQueuedWorkPaperCellMutationRefs({
      refs,
      potentialNewCells,
      applyCellMutationsAtWithOptions: this.runtime.applyCellMutationsAtWithOptions,
      updateSheetDimensionsAfterCellMutationRefs: this.runtime.updateSheetDimensionsAfterCellMutationRefs,
    })
  }

  flushSuspendedCellMutations(): void {
    if (this.suspendedCellMutationRefs.length === 0) {
      return
    }
    const refs = this.suspendedCellMutationRefs
    const potentialNewCells = this.suspendedCellMutationPotentialNewCells
    this.suspendedCellMutationRefs = []
    this.suspendedCellMutationPotentialNewCells = 0
    applyQueuedWorkPaperCellMutationRefs({
      refs,
      potentialNewCells,
      applyCellMutationsAtWithOptions: this.runtime.applyCellMutationsAtWithOptions,
      updateSheetDimensionsAfterCellMutationRefs: this.runtime.updateSheetDimensionsAfterCellMutationRefs,
    })
  }

  enqueueSuspendedLiteralMutation(input: WorkPaperLiteralMutationQueueInput): boolean {
    return tryEnqueueWorkPaperLiteralMutation({
      enabled: true,
      queue: this.suspendedCellMutationRefs,
      ...input,
      addPotentialNewCell: () => {
        this.suspendedCellMutationPotentialNewCells += 1
      },
    })
  }

  enqueueDeferredBatchLiteral(input: WorkPaperLiteralMutationQueueInput): boolean {
    return tryEnqueueWorkPaperLiteralMutation({
      enabled: true,
      queue: this.pendingBatchOps,
      ...input,
      addPotentialNewCell: () => {
        this.pendingBatchPotentialNewCells += 1
      },
    })
  }

  enqueueValidatedDeferredBatchLiteral(input: WorkPaperLiteralMutationQueueInput): void {
    this.pendingBatchOps.push({
      sheetId: input.sheetId,
      mutation:
        input.content === null
          ? { kind: 'clearCell', row: input.row, col: input.col }
          : { kind: 'setCellValue', row: input.row, col: input.col, value: input.content },
      ...(input.cellIndex !== undefined ? { cellIndex: input.cellIndex } : {}),
    })
    if (input.content !== null && input.cellIndex === undefined) {
      this.pendingBatchPotentialNewCells += 1
    }
  }
}
