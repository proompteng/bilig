import type { SheetRecord } from '@bilig/core/headless-runtime'
import { formatAddress } from '@bilig/formula'
import { WorkPaperNoOperationToRedoError, WorkPaperNoOperationToUndoError, WorkPaperOperationError } from './work-paper-errors.js'
import { sourceRangeRef } from './work-paper-address-format.js'
import { applyWorkPaperHistoryOperation } from './work-paper-history-operations.js'
import type { WorkPaperHistoryRecord } from './work-paper-history.js'
import type { WorkPaperSheetDimensionCache } from './work-paper-sheet-dimension-cache.js'
import type { QueuedEvent } from './work-paper-tracked-event-helpers.js'
import type {
  WorkPaperAddressLike,
  WorkPaperCellAddress,
  WorkPaperCellRange,
  WorkPaperChange,
  WorkPaperDependencyRef,
} from './work-paper-types.js'
import { isCellRange } from './work-paper-runtime-helpers.js'
import { WorkPaperPublicSurface } from './work-paper-public-surface.js'

export abstract class WorkPaperRuntimeMetadataSurface extends WorkPaperPublicSurface {
  protected abstract readonly sheetDimensionCache: WorkPaperSheetDimensionCache

  protected abstract flushPendingBatchOps(): void
  protected abstract getUndoStack(): WorkPaperHistoryRecord[]
  protected abstract getRedoStack(): WorkPaperHistoryRecord[]
  protected abstract canUseTrackedMutationFastPath(): boolean
  protected abstract captureTrackedChangesWithoutVisibilityCache(
    mutate: () => void,
    options?: { readonly preservePendingTrackedPositions?: boolean },
  ): WorkPaperChange[]
  protected abstract captureChanges(
    semanticEvent: QueuedEvent | undefined,
    mutate: () => void,
    options?: { readonly preservePendingTrackedPositions?: boolean },
  ): WorkPaperChange[]
  protected abstract listSheetRecords(): readonly SheetRecord[]
  protected abstract sheetName(sheetId: number): string
  protected abstract a1(address: Pick<WorkPaperCellAddress, 'row' | 'col'>): string
  protected abstract getDirectPrecedentStrings(address: WorkPaperCellAddress): string[]
  protected abstract getDirectPrecedentRefs(address: WorkPaperCellAddress): WorkPaperDependencyRef[]
  protected abstract collectRangeDependencies(
    range: WorkPaperCellRange,
    readDependencies: (address: WorkPaperCellAddress) => readonly string[],
  ): WorkPaperDependencyRef[]
  protected abstract toDependencyRefs(values: readonly string[]): WorkPaperDependencyRef[]

  undo(): WorkPaperChange[] {
    this.assertNotDisposed()
    return this.applyHistoryOperation('undo')
  }

  redo(): WorkPaperChange[] {
    this.assertNotDisposed()
    return this.applyHistoryOperation('redo')
  }

  isThereSomethingToUndo(): boolean {
    this.assertNotDisposed()
    return this.getUndoStack().length > 0
  }

  isThereSomethingToRedo(): boolean {
    this.assertNotDisposed()
    return this.getRedoStack().length > 0
  }

  clearUndoStack(): void {
    this.assertNotDisposed()
    this.getUndoStack().length = 0
  }

  clearRedoStack(): void {
    this.assertNotDisposed()
    this.getRedoStack().length = 0
  }

  getCellDependents(address: WorkPaperAddressLike): WorkPaperDependencyRef[] {
    this.assertNotDisposed()
    this.flushPendingBatchOps()
    if (!isCellRange(address)) {
      return this.toDependencyRefs(this.engine.getDependents(this.sheetName(address.sheet), this.a1(address)).directDependents)
    }
    return this.collectRangeDependencies(
      address,
      (cellAddress) => this.engine.getDependents(this.sheetName(cellAddress.sheet), this.a1(cellAddress)).directDependents,
    )
  }

  getCellPrecedents(address: WorkPaperAddressLike): WorkPaperDependencyRef[] {
    this.assertNotDisposed()
    this.flushPendingBatchOps()
    if (!isCellRange(address)) {
      return this.getDirectPrecedentRefs(address)
    }
    return this.collectRangeDependencies(address, (cellAddress) => this.getDirectPrecedentStrings(cellAddress))
  }

  getSheetName(sheetId: number): string | undefined {
    this.assertNotDisposed()
    return this.engine.workbook.getSheetById(sheetId)?.name
  }

  getSheetNames(): string[] {
    this.assertNotDisposed()
    return this.listSheetRecords().map((sheet) => sheet.name)
  }

  getSheetId(name: string): number | undefined {
    this.assertNotDisposed()
    return this.engine.workbook.getSheet(name)?.id
  }

  doesSheetExist(name: string): boolean {
    this.assertNotDisposed()
    return this.engine.workbook.getSheet(name) !== undefined
  }

  countSheets(): number {
    this.assertNotDisposed()
    return this.listSheetRecords().length
  }

  moveCells(source: WorkPaperCellRange, target: WorkPaperCellAddress): WorkPaperChange[] {
    this.assertNotDisposed()
    if (!this.isItPossibleToMoveCells(source, target)) {
      throw new WorkPaperOperationError('Cells cannot be moved')
    }
    const sourceHeight = source.end.row - source.start.row
    const sourceWidth = source.end.col - source.start.col
    return this.captureChanges(undefined, () => {
      this.engine.moveRange(sourceRangeRef(this.sheetName(source.start.sheet), source), {
        sheetName: this.sheetName(target.sheet),
        startAddress: formatAddress(target.row, target.col),
        endAddress: formatAddress(target.row + sourceHeight, target.col + sourceWidth),
      })
      this.sheetDimensionCache.invalidate(source.start.sheet)
      this.sheetDimensionCache.invalidate(target.sheet)
    })
  }

  protected applyHistoryOperation(kind: 'undo' | 'redo'): WorkPaperChange[] {
    return applyWorkPaperHistoryOperation({
      getStack: () => (kind === 'undo' ? this.getUndoStack() : this.getRedoStack()),
      canUseTrackedMutationFastPath: () => this.canUseTrackedMutationFastPath(),
      captureTrackedChangesWithoutVisibilityCache: (mutate, options) => this.captureTrackedChangesWithoutVisibilityCache(mutate, options),
      captureChanges: (mutate, options) => this.captureChanges(undefined, mutate, options),
      applyOperation: () => (kind === 'undo' ? this.engine.undo() : this.engine.redo()),
      createMissingOperationError: () => (kind === 'undo' ? new WorkPaperNoOperationToUndoError() : new WorkPaperNoOperationToRedoError()),
      invalidateAllSheetDimensions: () => this.sheetDimensionCache.invalidateAll(),
    })
  }
}
