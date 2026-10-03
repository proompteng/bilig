import type { EngineCellMutationRef, EngineExistingNumericCellMutationsRef } from '@bilig/core/headless-runtime'
import { applyQueuedWorkPaperCellMutationRefs, type WorkPaperCellMutationApplyOptions } from './work-paper-cell-mutation-refs.js'
import { tryEnqueueWorkPaperLiteralMutation } from './work-paper-literal-mutation-queue.js'
import type { RawCellContent } from './work-paper-types.js'

const INITIAL_EXISTING_NUMERIC_BATCH_CAPACITY = 256

export interface WorkPaperMutationQueuesRuntime {
  readonly applyCellMutationsAtWithOptions: (refs: readonly EngineCellMutationRef[], options: WorkPaperCellMutationApplyOptions) => void
  readonly applyExistingNumericCellMutationsAtWithOptions?: (
    refs: EngineExistingNumericCellMutationsRef,
    options: WorkPaperCellMutationApplyOptions,
  ) => boolean
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
  private pendingExistingNumericBatch: ExistingNumericMutationQueue | undefined
  private suspendedCellMutationRefs: EngineCellMutationRef[] = []
  private suspendedCellMutationPotentialNewCells = 0

  constructor(private readonly runtime: WorkPaperMutationQueuesRuntime) {}

  hasPendingBatchOps(): boolean {
    return this.pendingBatchOps.length > 0 || (this.pendingExistingNumericBatch?.length ?? 0) > 0
  }

  appendSuspendedCellMutationRefs(refs: readonly EngineCellMutationRef[]): void {
    this.suspendedCellMutationRefs.push(...refs)
  }

  addSuspendedCellMutationPotentialNewCells(amount: number): void {
    this.suspendedCellMutationPotentialNewCells += amount
  }

  flushPendingBatchOps(): void {
    if (this.pendingBatchOps.length === 0 && (this.pendingExistingNumericBatch?.length ?? 0) === 0) {
      return
    }
    if (this.pendingBatchOps.length === 0 && this.pendingExistingNumericBatch !== undefined) {
      const record = this.pendingExistingNumericBatch.toRecord()
      this.pendingExistingNumericBatch = undefined
      if (
        this.runtime.applyExistingNumericCellMutationsAtWithOptions?.(record, {
          captureUndo: true,
          potentialNewCells: 0,
          source: 'local',
          returnUndoOps: false,
        }) === true
      ) {
        return
      }
      this.pendingBatchOps = existingNumericMutationRecordToRefs(record)
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

  enqueueValidatedSuspendedLiteral(input: WorkPaperLiteralMutationQueueInput): void {
    this.suspendedCellMutationRefs.push({
      sheetId: input.sheetId,
      mutation:
        input.content === null
          ? { kind: 'clearCell', row: input.row, col: input.col }
          : { kind: 'setCellValue', row: input.row, col: input.col, value: input.content },
      ...(input.cellIndex !== undefined ? { cellIndex: input.cellIndex } : {}),
    })
    if (input.content !== null && input.cellIndex === undefined) {
      this.suspendedCellMutationPotentialNewCells += 1
    }
  }

  enqueueDeferredBatchLiteral(input: WorkPaperLiteralMutationQueueInput): boolean {
    if (this.tryEnqueueDeferredExistingNumeric(input)) {
      return true
    }
    this.materializePendingExistingNumericBatch()
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
    if (this.tryEnqueueDeferredExistingNumeric(input)) {
      return
    }
    this.materializePendingExistingNumericBatch()
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

  private tryEnqueueDeferredExistingNumeric(input: WorkPaperLiteralMutationQueueInput): boolean {
    if (this.pendingBatchOps.length !== 0 || typeof input.content !== 'number' || input.cellIndex === undefined) {
      return false
    }
    const batch = (this.pendingExistingNumericBatch ??= new ExistingNumericMutationQueue())
    batch.push(input.sheetId, input.cellIndex, input.row, input.col, input.content)
    return true
  }

  private materializePendingExistingNumericBatch(): void {
    if (this.pendingExistingNumericBatch === undefined) {
      return
    }
    this.pendingBatchOps.push(...this.pendingExistingNumericBatch.toRefs())
    this.pendingExistingNumericBatch = undefined
  }
}

class ExistingNumericMutationQueue {
  private sheetIds = new Uint32Array(0)
  private cellIndexPlusOnes = new Uint32Array(0)
  private rows = new Uint32Array(0)
  private cols = new Uint32Array(0)
  private numbers = new Float64Array(0)
  length = 0

  push(sheetId: number, cellIndex: number, row: number, col: number, value: number): void {
    this.ensureCapacity(this.length + 1)
    this.sheetIds[this.length] = sheetId
    this.cellIndexPlusOnes[this.length] = cellIndex + 1
    this.rows[this.length] = row
    this.cols[this.length] = col
    this.numbers[this.length] = value
    this.length += 1
  }

  toRecord(): EngineExistingNumericCellMutationsRef {
    const useFullCapacity = this.length === this.sheetIds.length
    return {
      sheetIds: useFullCapacity ? this.sheetIds : this.sheetIds.subarray(0, this.length),
      cellIndexPlusOnes: useFullCapacity ? this.cellIndexPlusOnes : this.cellIndexPlusOnes.subarray(0, this.length),
      rows: useFullCapacity ? this.rows : this.rows.subarray(0, this.length),
      cols: useFullCapacity ? this.cols : this.cols.subarray(0, this.length),
      numbers: useFullCapacity ? this.numbers : this.numbers.subarray(0, this.length),
      potentialNewCells: 0,
    }
  }

  toRefs(): EngineCellMutationRef[] {
    return existingNumericMutationRecordToRefs(this.toRecord())
  }

  private ensureCapacity(required: number): void {
    if (this.sheetIds.length >= required) {
      return
    }
    const nextCapacity = Math.max(INITIAL_EXISTING_NUMERIC_BATCH_CAPACITY, this.sheetIds.length * 2, required)
    const nextSheetIds = new Uint32Array(nextCapacity)
    const nextCellIndexPlusOnes = new Uint32Array(nextCapacity)
    const nextRows = new Uint32Array(nextCapacity)
    const nextCols = new Uint32Array(nextCapacity)
    const nextNumbers = new Float64Array(nextCapacity)
    nextSheetIds.set(this.sheetIds)
    nextCellIndexPlusOnes.set(this.cellIndexPlusOnes)
    nextRows.set(this.rows)
    nextCols.set(this.cols)
    nextNumbers.set(this.numbers)
    this.sheetIds = nextSheetIds
    this.cellIndexPlusOnes = nextCellIndexPlusOnes
    this.rows = nextRows
    this.cols = nextCols
    this.numbers = nextNumbers
  }
}

function existingNumericMutationRecordToRefs(record: EngineExistingNumericCellMutationsRef): EngineCellMutationRef[] {
  const refs = Array<EngineCellMutationRef>(record.sheetIds.length)
  for (let index = 0; index < refs.length; index += 1) {
    const cellIndexPlusOne = record.cellIndexPlusOnes[index]!
    refs[index] = {
      sheetId: record.sheetIds[index]!,
      ...(cellIndexPlusOne === 0 ? {} : { cellIndex: cellIndexPlusOne - 1 }),
      mutation: {
        kind: 'setCellValue',
        row: record.rows[index]!,
        col: record.cols[index]!,
        value: record.numbers[index] ?? 0,
      },
    }
  }
  return refs
}
