import { CellFlags, type EngineCellMutationRef, type SheetRecord } from '@bilig/core/headless-runtime'
import type { CellRangeRef, CellValue, LiteralInput, WorkbookSnapshot } from '@bilig/protocol'
import { formatAddress } from '@bilig/formula'
import { assertRange, valuesEqual } from './work-paper-runtime-helpers.js'
import { formatTrackedA1, sourceRangeRef } from './work-paper-address-format.js'
import {
  clearWorkPaperHistoryStacks,
  mergeWorkPaperUndoHistory,
  readWorkPaperHistoryStack,
  type WorkPaperHistoryRecord,
} from './work-paper-history.js'
import type { WorkPaperSheetDimensionCache } from './work-paper-sheet-dimension-cache.js'
import {
  WorkPaperEvaluationSuspendedError,
  WorkPaperNoSheetWithIdError,
  WorkPaperNoSheetWithNameError,
  WorkPaperOperationError,
} from './work-paper-errors.js'
import {
  captureWorkPaperNamedExpressionValueSnapshot,
  computeWorkPaperNamedExpressionChanges,
  type InternalNamedExpression,
  type WorkPaperNamedExpressionValueSnapshot,
} from './work-paper-named-expression-helpers.js'
import type { WorkPaperEngineEventTracker } from './work-paper-engine-event-tracker.js'
import type { WorkPaperMutationQueues } from './work-paper-mutation-queues.js'
import { applyWorkPaperCellMutationRefs } from './work-paper-cell-mutation-refs.js'
import {
  materializeTrackedIndexChangeSourcesWithMetadata,
  materializeTrackedIndexChangesWithMetadata,
} from './tracked-cell-index-changes.js'
import { collectWorkPaperRangeDependencies, toWorkPaperDependencyRefs } from './work-paper-dependency-refs.js'
import {
  captureWorkPaperVisibilitySnapshot,
  computeWorkPaperCellChangesFromVisibilitySnapshots,
  type VisibilitySnapshot,
} from './work-paper-visibility-snapshot.js'
import type { TrackedEngineEvent, TrackedPatch } from './tracked-engine-event-refs.js'
import {
  computeWorkPaperTrackedCellChangesFromEvents,
  tryReadTinyTrackedEventChangesWithoutVisibility as tryReadTinyTrackedEventChangesWithoutVisibilityFromReducer,
  type MaterializedTrackedEventChanges,
} from './work-paper-tracked-change-reducer.js'
import {
  TINY_TRACKED_CHANGE_LIMIT,
  readTrackedCellChange,
  readTrustedPhysicalTrackedChangeMetadata,
  readTinySortedPhysicalTrackedEventChanges as readTinySortedPhysicalTrackedEventChangesFromStore,
  trackedEventHasNoValueChanges,
  tryBuildDirectSingleLiteralTrackedChange,
  withEventChanges,
  type QueuedEvent,
} from './work-paper-tracked-event-helpers.js'
import { tryRenameWorkPaperSheetWithoutVisibilitySnapshots } from './work-paper-sheet-rename-fast-path.js'
import { WORKPAPER_PUBLIC_ERROR_NAMES } from './work-paper-config.js'
import type {
  RawCellContent,
  WorkPaperCellAddress,
  WorkPaperCellChange,
  WorkPaperCellRange,
  WorkPaperChange,
  WorkPaperDependencyRef,
  WorkPaperStats,
} from './work-paper-types.js'
import { WorkPaperRuntimeMetadataSurface } from './work-paper-runtime-metadata-surface.js'

type NamedExpressionValueSnapshot = WorkPaperNamedExpressionValueSnapshot
type RebuildValueSnapshot = Map<number, CellValue>
const FORMULA_REBUILD_VALUE_FLAGS = CellFlags.HasFormula | CellFlags.SpillChild | CellFlags.PivotOutput
export const EMPTY_NAMED_EXPRESSION_VALUES: NamedExpressionValueSnapshot = new Map()

function shouldPreferLazyPublicChanges(events: readonly TrackedEngineEvent[], shouldEmitValuesUpdated: boolean): boolean {
  if (!shouldEmitValuesUpdated) {
    return true
  }
  return events.some(
    (event) =>
      event.changedCellIndices.length > TINY_TRACKED_CHANGE_LIMIT &&
      event.invalidation !== 'full' &&
      event.patches === undefined &&
      !event.hasInvalidatedRanges &&
      !event.hasInvalidatedRows &&
      !event.hasInvalidatedColumns,
  )
}

export abstract class WorkPaperRuntimeSurface extends WorkPaperRuntimeMetadataSurface {
  protected abstract readonly namedExpressions: Map<string, InternalNamedExpression>
  protected abstract readonly engineEvents: WorkPaperEngineEventTracker
  protected abstract readonly mutationQueues: WorkPaperMutationQueues
  protected abstract override sheetDimensionCache: WorkPaperSheetDimensionCache
  protected abstract visibilityCache: VisibilitySnapshot | null
  protected abstract namedExpressionValueCache: NamedExpressionValueSnapshot | null
  protected abstract sheetRecordsCache: readonly SheetRecord[] | null
  protected abstract batchDepth: number
  protected abstract batchStartVisibility: VisibilitySnapshot | null
  protected abstract batchStartNamedValues: NamedExpressionValueSnapshot | null
  protected abstract batchUsesTrackedFastPath: boolean
  protected abstract batchUndoStackLength: number
  protected abstract evaluationSuspended: boolean
  protected abstract suspendedVisibility: VisibilitySnapshot | null
  protected abstract suspendedNamedValues: NamedExpressionValueSnapshot | null
  protected abstract suspendedUsesTrackedFastPath: boolean
  protected abstract queuedEvents: QueuedEvent[]
  protected abstract disposed: boolean

  protected abstract ensureEngineEventTracking(): void

  protected abstract evaluateNamedExpression(
    expression: InternalNamedExpression,
  ): ReturnType<WorkPaperRuntimeMetadataSurface['calculateFormula']>

  getStats(): WorkPaperStats {
    this.assertNotDisposed()
    return {
      batchDepth: this.batchDepth,
      evaluationSuspended: this.evaluationSuspended,
      lastMetrics: structuredClone(this.engine.getLastMetrics()),
    }
  }

  exportSnapshot(): WorkbookSnapshot {
    this.assertNotDisposed()
    this.engineEvents.materializePendingLazyChanges()
    return structuredClone(this.engine.exportSnapshot())
  }

  rebuildAndRecalculate(): WorkPaperChange[] {
    this.assertNotDisposed()
    this.engineEvents.materializePendingLazyChanges()
    if (this.shouldSuppressEvents()) {
      try {
        this.engine.recalculateNow()
        this.sheetDimensionCache.invalidateAll()
      } catch (error) {
        throw new WorkPaperOperationError(this.messageOf(error, 'Recalculation failed'))
      }
      return []
    }
    if (this.canUseTrackedMutationFastPath()) {
      this.ensureEngineEventTracking()
      const beforeFormulaValues = this.captureFormulaResultValueSnapshot()
      this.engineEvents.drain()
      try {
        this.engineEvents.withRetainedIndices(() => {
          this.engine.recalculateNow()
          this.sheetDimensionCache.invalidateAll()
        })
      } catch (error) {
        throw new WorkPaperOperationError(this.messageOf(error, 'Recalculation failed'))
      }
      const shouldEmitValuesUpdated = this.emitter.hasListeners('valuesUpdated')
      const events = this.engineEvents.drain()
      const changes = this.filterUnchangedRebuildChanges(
        this.computeTrackedChangesWithoutVisibilityCache(events, {
          preferLazyPublicChanges: shouldPreferLazyPublicChanges(events, shouldEmitValuesUpdated),
        }),
        beforeFormulaValues,
      )
      if (changes.length > 0 && shouldEmitValuesUpdated) {
        this.emitter.emitDetailed({ eventName: 'valuesUpdated', payload: { changes } })
      }
      return changes
    }
    const beforeVisibility = this.ensureVisibilityCache()
    const beforeNames = this.namedExpressions.size > 0 ? this.ensureNamedExpressionValueCache() : EMPTY_NAMED_EXPRESSION_VALUES
    this.ensureEngineEventTracking()
    this.engineEvents.drain()
    try {
      this.engine.recalculateNow()
      this.sheetDimensionCache.invalidateAll()
    } catch (error) {
      throw new WorkPaperOperationError(this.messageOf(error, 'Recalculation failed'))
    }
    const afterVisibility = this.captureVisibilitySnapshot()
    const afterNames = this.namedExpressions.size > 0 ? this.captureNamedExpressionValueSnapshot() : EMPTY_NAMED_EXPRESSION_VALUES
    this.visibilityCache = afterVisibility
    this.namedExpressionValueCache = afterNames
    const changes = [
      ...this.computeCellChanges(beforeVisibility, afterVisibility),
      ...this.computeNamedExpressionChanges(beforeNames, afterNames),
    ]
    if (changes.length > 0 && this.emitter.hasListeners('valuesUpdated')) {
      this.emitter.emitDetailed({ eventName: 'valuesUpdated', payload: { changes } })
    }
    return changes
  }

  batch(batchOperations: () => void): WorkPaperChange[] {
    this.assertNotDisposed()
    this.engineEvents.materializePendingLazyChanges()
    const isOutermost = this.batchDepth === 0
    if (isOutermost) {
      this.batchUsesTrackedFastPath = this.canUseTrackedMutationFastPath()
      if (this.batchUsesTrackedFastPath) {
        this.ensureEngineEventTracking()
        this.batchStartVisibility = null
        this.batchStartNamedValues = EMPTY_NAMED_EXPRESSION_VALUES
      } else {
        this.batchStartVisibility = this.ensureVisibilityCache()
        this.batchStartNamedValues = this.ensureNamedExpressionValueCache()
      }
      this.batchUndoStackLength = this.getUndoStack().length
      this.engineEvents.drain()
    }
    this.batchDepth += 1
    try {
      batchOperations()
    } finally {
      this.batchDepth -= 1
      if (isOutermost) {
        if (this.batchUsesTrackedFastPath) {
          this.engineEvents.withRetainedIndices(() => {
            this.flushPendingBatchOps()
          })
        } else {
          this.flushPendingBatchOps()
        }
        this.mergeUndoHistory(this.batchUndoStackLength)
      }
    }
    if (!isOutermost) {
      return []
    }
    const shouldEmitValuesUpdated = this.emitter.hasListeners('valuesUpdated')
    const events = this.batchUsesTrackedFastPath ? this.engineEvents.drain() : []
    const changes = this.batchUsesTrackedFastPath
      ? this.computeTrackedChangesWithoutVisibilityCache(events, {
          preferLazyPublicChanges: shouldPreferLazyPublicChanges(events, shouldEmitValuesUpdated),
        })
      : this.computeChangesAfterMutation(this.batchStartVisibility ?? new Map(), this.batchStartNamedValues ?? new Map())
    this.batchUsesTrackedFastPath = false
    this.batchStartVisibility = null
    this.batchStartNamedValues = null
    if (!this.evaluationSuspended) {
      this.flushQueuedEvents()
      if (changes.length > 0 && shouldEmitValuesUpdated) {
        this.emitter.emitDetailed({ eventName: 'valuesUpdated', payload: { changes } })
      }
    }
    return changes
  }

  suspendEvaluation(): void {
    this.assertNotDisposed()
    this.engineEvents.materializePendingLazyChanges()
    if (this.evaluationSuspended) {
      return
    }
    this.evaluationSuspended = true
    this.flushPendingBatchOps()
    if (this.visibilityCache === null && this.namedExpressions.size === 0) {
      this.suspendedVisibility = null
      this.suspendedNamedValues = EMPTY_NAMED_EXPRESSION_VALUES
      this.suspendedUsesTrackedFastPath = true
    } else {
      this.suspendedVisibility = this.ensureVisibilityCache()
      this.suspendedNamedValues = this.ensureNamedExpressionValueCache()
      this.suspendedUsesTrackedFastPath = false
    }
    this.engineEvents.drain()
    this.emitter.emitDetailed({ eventName: 'evaluationSuspended', payload: {} })
  }

  resumeEvaluation(): WorkPaperChange[] {
    this.assertNotDisposed()
    this.engineEvents.materializePendingLazyChanges()
    if (!this.evaluationSuspended) {
      return []
    }
    if (this.suspendedUsesTrackedFastPath) {
      this.ensureEngineEventTracking()
      this.engineEvents.withRetainedIndices(() => {
        this.flushSuspendedCellMutations()
      })
    } else {
      this.flushSuspendedCellMutations()
    }
    const shouldEmitValuesUpdated = this.emitter.hasListeners('valuesUpdated')
    const events = this.suspendedUsesTrackedFastPath ? this.engineEvents.drain() : []
    const changes = this.suspendedUsesTrackedFastPath
      ? this.computeTrackedChangesWithoutVisibilityCache(events, {
          preferLazyPublicChanges: shouldPreferLazyPublicChanges(events, shouldEmitValuesUpdated),
        })
      : this.computeChangesAfterMutation(this.suspendedVisibility ?? new Map(), this.suspendedNamedValues ?? new Map())
    this.evaluationSuspended = false
    this.suspendedVisibility = null
    this.suspendedNamedValues = null
    this.suspendedUsesTrackedFastPath = false
    this.flushQueuedEvents()
    this.emitter.emitDetailed({ eventName: 'evaluationResumed', payload: { changes } })
    if (changes.length > 0 && shouldEmitValuesUpdated) {
      this.emitter.emitDetailed({ eventName: 'valuesUpdated', payload: { changes } })
    }
    return changes
  }

  isEvaluationSuspended(): boolean {
    return this.evaluationSuspended
  }

  protected resetChangeTrackingCaches(): void {
    this.sheetRecordsCache = null
    this.visibilityCache = null
    this.namedExpressionValueCache = null
    this.engineEvents.clearEvents()
  }

  protected ensureVisibilityCache(): VisibilitySnapshot {
    if (!this.visibilityCache) {
      this.visibilityCache = this.captureVisibilitySnapshot()
    }
    return this.visibilityCache
  }

  protected ensureNamedExpressionValueCache(): NamedExpressionValueSnapshot {
    if (!this.namedExpressionValueCache) {
      this.namedExpressionValueCache =
        this.namedExpressions.size > 0 ? this.captureNamedExpressionValueSnapshot() : EMPTY_NAMED_EXPRESSION_VALUES
    }
    return this.namedExpressionValueCache
  }

  protected flushPendingBatchOps(): void {
    this.mutationQueues.flushPendingBatchOps()
  }

  protected applyCellMutationRefs(
    refs: readonly EngineCellMutationRef[],
    options: {
      captureUndo?: boolean
      potentialNewCells?: number
      source?: 'local' | 'restore'
      returnUndoOps?: boolean
      reuseRefs?: boolean
    },
  ): void {
    applyWorkPaperCellMutationRefs(this.runtimeAdapters.cellMutationApplyRuntime, refs, options)
  }

  protected flushSuspendedCellMutations(): void {
    this.mutationQueues.flushSuspendedCellMutations()
  }

  protected enqueueSuspendedLiteralMutation(
    sheetId: number,
    row: number,
    col: number,
    content: RawCellContent,
    cellIndex: number | undefined,
  ): boolean {
    return (
      this.evaluationSuspended &&
      this.mutationQueues.enqueueSuspendedLiteralMutation({
        sheetId,
        row,
        col,
        content,
        cellIndex,
      })
    )
  }

  protected enqueueDeferredBatchLiteral(
    sheetId: number,
    row: number,
    col: number,
    content: RawCellContent,
    cellIndex: number | undefined,
  ): boolean {
    return (
      this.batchDepth !== 0 &&
      !this.evaluationSuspended &&
      this.mutationQueues.enqueueDeferredBatchLiteral({
        sheetId,
        row,
        col,
        content,
        cellIndex,
      })
    )
  }

  protected enqueueValidatedDeferredBatchLiteral(
    sheetId: number,
    row: number,
    col: number,
    content: RawCellContent,
    cellIndex: number | undefined,
  ): boolean {
    if (this.batchDepth === 0 || this.evaluationSuspended) {
      return false
    }
    this.mutationQueues.enqueueValidatedDeferredBatchLiteral({
      sheetId,
      row,
      col,
      content,
      cellIndex,
    })
    return true
  }

  protected prepareReadableState(): void {
    this.assertNotDisposed()
    this.flushPendingBatchOps()
  }

  protected assertNotDisposed(): void {
    if (this.disposed) {
      throw new WorkPaperOperationError('Workbook has been disposed')
    }
  }

  protected assertReadable(): void {
    this.prepareReadableState()
    if (this.evaluationSuspended) {
      throw new WorkPaperEvaluationSuspendedError()
    }
  }

  protected sheetRecord(sheetId: number): SheetRecord {
    const sheet = this.engine.workbook.getSheetById(sheetId)
    if (!sheet) {
      throw new WorkPaperNoSheetWithIdError(sheetId)
    }
    return sheet
  }

  protected sheetName(sheetId: number): string {
    return this.sheetRecord(sheetId).name
  }

  protected capabilityContext() {
    return {
      config: this.config,
      requireSheet: (sheetId: number): void => {
        void this.sheetRecord(sheetId)
      },
      doesSheetExist: (sheetName: string): boolean => this.doesSheetExist(sheetName),
      getSheetIdByName: (sheetName: string): number | undefined => this.getSheetId(sheetName),
    }
  }

  protected requireSheetId(name: string): number {
    const sheetId = this.getSheetId(name)
    if (sheetId === undefined) {
      throw new WorkPaperNoSheetWithNameError(name)
    }
    return sheetId
  }

  protected a1(address: Pick<WorkPaperCellAddress, 'row' | 'col'>): string {
    return formatAddress(address.row, address.col)
  }

  protected trackedA1(row: number, col: number): string {
    return formatTrackedA1(row, col)
  }

  protected rangeRef(range: WorkPaperCellRange): CellRangeRef {
    assertRange(range)
    return sourceRangeRef(this.sheetName(range.start.sheet), range)
  }

  protected getDirectPrecedentStrings(address: WorkPaperCellAddress): string[] {
    const precedents = new Set<string>(this.engine.getDependencies(this.sheetName(address.sheet), this.a1(address)).directPrecedents)
    const formula = this.getCellFormula(address)
    if (formula) {
      this.getNamedExpressionsFromFormula(formula).forEach((name) => {
        precedents.add(name)
      })
    }
    return [...precedents]
  }

  protected getDirectPrecedentRefs(address: WorkPaperCellAddress): WorkPaperDependencyRef[] {
    return this.toDependencyRefs(this.getDirectPrecedentStrings(address))
  }

  protected listSheetRecords(): readonly SheetRecord[] {
    if (this.sheetRecordsCache) {
      return this.sheetRecordsCache
    }
    this.sheetRecordsCache = [...this.engine.workbook.sheetsByName.values()].toSorted(
      (left, right) => left.order - right.order || left.name.localeCompare(right.name),
    )
    return this.sheetRecordsCache
  }

  protected captureVisibilitySnapshot(): VisibilitySnapshot {
    return captureWorkPaperVisibilitySnapshot({
      sheets: this.listSheetRecords(),
      cellStore: this.engine.workbook.cellStore,
      strings: this.engine.strings,
    })
  }

  protected captureFormulaResultValueSnapshot(): RebuildValueSnapshot {
    const snapshot: RebuildValueSnapshot = new Map()
    const cellStore = this.engine.workbook.cellStore
    this.listSheetRecords().forEach((sheet) => {
      sheet.grid.forEachCellEntry((cellIndex) => {
        if (((cellStore.flags[cellIndex] ?? 0) & FORMULA_REBUILD_VALUE_FLAGS) === 0) {
          return
        }
        snapshot.set(
          cellIndex,
          cellStore.getValue(cellIndex, (id) => this.engine.strings.get(id)),
        )
      })
    })
    return snapshot
  }

  protected filterUnchangedRebuildChanges(changes: WorkPaperChange[], beforeFormulaValues: RebuildValueSnapshot): WorkPaperChange[] {
    if (changes.length === 0 || beforeFormulaValues.size === 0) {
      return changes
    }
    return changes.filter((change) => {
      if (change.kind !== 'cell') {
        return true
      }
      const cellIndex = this.engine.workbook.getCellIndexAt(change.address.sheet, change.address.row, change.address.col)
      if (cellIndex === undefined) {
        return true
      }
      const beforeValue = beforeFormulaValues.get(cellIndex)
      return beforeValue === undefined || !valuesEqual(beforeValue, change.newValue)
    })
  }

  protected captureNamedExpressionValueSnapshot(): NamedExpressionValueSnapshot {
    if (this.namedExpressions.size === 0) {
      return EMPTY_NAMED_EXPRESSION_VALUES
    }
    return captureWorkPaperNamedExpressionValueSnapshot(this.namedExpressions.values(), (expression) =>
      this.evaluateNamedExpression(expression),
    )
  }

  protected computeCellChanges(beforeVisibility: VisibilitySnapshot, afterVisibility: VisibilitySnapshot): WorkPaperChange[] {
    return computeWorkPaperCellChangesFromVisibilitySnapshots({
      beforeVisibility,
      afterVisibility,
      sheets: this.listSheetRecords(),
    })
  }

  protected materializeTrackedEventChanges(event: TrackedEngineEvent, lazy = false): MaterializedTrackedEventChanges {
    if (event.patches && event.patches.length > 0) {
      const cellPatches = event.patches.filter((patch): patch is Extract<TrackedPatch, { kind: 'cell' }> => patch.kind === 'cell')
      return { changes: cellPatches, canReusePublicChanges: false, ordered: false }
    }
    const trustedPhysicalMetadata =
      lazy && event.changedCellIndices instanceof Uint32Array
        ? readTrustedPhysicalTrackedChangeMetadata(event.changedCellIndices)
        : undefined
    const materialized = materializeTrackedIndexChangesWithMetadata(this.engine, event.changedCellIndices, {
      lazy,
      ...(event.explicitChangedCount !== undefined ? { explicitChangedCount: event.explicitChangedCount } : {}),
      ...trustedPhysicalMetadata,
    })
    if (lazy) {
      this.engineEvents.trackLazyChanges(materialized.changes)
    }
    return {
      changes: materialized.changes,
      canReusePublicChanges: true,
      ordered: materialized.ordered,
    }
  }

  protected readSingleTrackedCellChange(cellIndex: number): WorkPaperCellChange | undefined {
    return readTrackedCellChange({
      cellIndex,
      workbook: this.engine.workbook,
      strings: this.engine.strings,
      trackedA1: (row, col) => this.trackedA1(row, col),
    })
  }

  protected readTinySortedPhysicalTrackedEventChanges(event: TrackedEngineEvent): WorkPaperCellChange[] | null {
    return readTinySortedPhysicalTrackedEventChangesFromStore({
      event,
      workbook: this.engine.workbook,
      strings: this.engine.strings,
      trackedA1: (row, col) => this.trackedA1(row, col),
    })
  }

  protected tryReadTinyTrackedEventChangesWithoutVisibility(event: TrackedEngineEvent): WorkPaperChange[] | null {
    return tryReadTinyTrackedEventChangesWithoutVisibilityFromReducer({
      event,
      listSheets: () => this.listSheetRecords(),
      materializeTrackedEventChanges: (trackedEvent, lazy) => this.materializeTrackedEventChanges(trackedEvent, lazy),
      readSingleTrackedCellChange: (cellIndex) => this.readSingleTrackedCellChange(cellIndex),
      readTinySortedPhysicalTrackedEventChanges: (trackedEvent) => this.readTinySortedPhysicalTrackedEventChanges(trackedEvent),
      sheetOrder: (sheetId) => this.sheetRecord(sheetId).order,
    })
  }

  protected computeCellChangesFromTrackedEvents(
    beforeVisibility: VisibilitySnapshot,
    events: readonly TrackedEngineEvent[],
    updateVisibility = true,
    options: { readonly preferLazyPublicChanges?: boolean } = {},
  ): { changes: WorkPaperChange[]; nextVisibility: VisibilitySnapshot } | null {
    return computeWorkPaperTrackedCellChangesFromEvents({
      beforeVisibility,
      events,
      updateVisibility,
      ...(options.preferLazyPublicChanges !== undefined ? { preferLazyPublicChanges: options.preferLazyPublicChanges } : {}),
      listSheets: () => this.listSheetRecords(),
      materializeTrackedEventChanges: (event, lazy) => this.materializeTrackedEventChanges(event, lazy),
      materializeTrackedEventSources: (trackedEvents, sourceOptions) => {
        const materializedSources = materializeTrackedIndexChangeSourcesWithMetadata(this.engine, trackedEvents, {
          deferLazyDetach: true,
          ...(sourceOptions.preferLazyPublicChanges !== undefined ? { lazy: sourceOptions.preferLazyPublicChanges } : {}),
        })
        if (!materializedSources) {
          return null
        }
        this.engineEvents.trackLazyChanges(materializedSources.changes)
        return materializedSources
      },
      readSingleTrackedCellChange: (cellIndex) => this.readSingleTrackedCellChange(cellIndex),
      readTinySortedPhysicalTrackedEventChanges: (event) => this.readTinySortedPhysicalTrackedEventChanges(event),
      sheetOrder: (sheetId) => this.sheetRecord(sheetId).order,
    })
  }

  protected computeNamedExpressionChanges(
    beforeNames: NamedExpressionValueSnapshot,
    afterNames: NamedExpressionValueSnapshot,
  ): WorkPaperChange[] {
    return computeWorkPaperNamedExpressionChanges({
      beforeNames,
      afterNames,
      expressionsByKey: this.namedExpressions,
    })
  }

  protected canUseTrackedStructuralFastPath(): boolean {
    return this.batchDepth === 0 && !this.evaluationSuspended && this.visibilityCache === null && this.namedExpressions.size === 0
  }

  protected canUseTrackedMutationFastPath(): boolean {
    return this.batchDepth === 0 && !this.evaluationSuspended && this.visibilityCache === null && this.namedExpressions.size === 0
  }

  protected canUseNamedExpressionChangeFastPath(): boolean {
    return this.batchDepth === 0 && !this.evaluationSuspended && this.visibilityCache === null && !this.emitter.hasAnyListeners()
  }

  protected canUseMetadataOnlySheetRenameFastPath(): boolean {
    return this.batchDepth === 0 && !this.evaluationSuspended && this.visibilityCache === null && !this.emitter.hasAnyListeners()
  }

  protected downgradeTrackedBatchFastPath(): void {
    if (!this.batchUsesTrackedFastPath || this.batchDepth === 0) {
      return
    }
    this.batchStartVisibility = this.ensureVisibilityCache()
    this.batchStartNamedValues = this.namedExpressions.size > 0 ? this.ensureNamedExpressionValueCache() : EMPTY_NAMED_EXPRESSION_VALUES
    this.batchUsesTrackedFastPath = false
  }

  protected computeTrackedChangesWithoutVisibilityCache(
    events: readonly TrackedEngineEvent[],
    options: { readonly preferLazyPublicChanges?: boolean } = {},
  ): WorkPaperChange[] {
    if (events.length > 0 && events.every(trackedEventHasNoValueChanges)) {
      return []
    }
    if (events.length === 1) {
      const event = events[0]!
      if (!options.preferLazyPublicChanges || event.changedCellIndices.length <= TINY_TRACKED_CHANGE_LIMIT) {
        const tinyChanges = this.tryReadTinyTrackedEventChangesWithoutVisibility(event)
        if (tinyChanges) {
          return tinyChanges
        }
      }
      if (
        options.preferLazyPublicChanges &&
        event.invalidation !== 'full' &&
        event.patches === undefined &&
        !event.hasInvalidatedRanges &&
        !event.hasInvalidatedRows &&
        !event.hasInvalidatedColumns
      ) {
        const materialized = this.materializeTrackedEventChanges(event, true)
        if (materialized.canReusePublicChanges && materialized.ordered) {
          return materialized.changes
        }
      }
    }
    const fastPath = this.computeCellChangesFromTrackedEvents(new Map(), events, false, options)
    if (!fastPath) {
      throw new WorkPaperOperationError('Mutation emitted an unsupported invalidation pattern for tracked changes')
    }
    return fastPath.changes
  }

  protected tryRenameSheetWithoutVisibilitySnapshots(oldName: string, newName: string): WorkPaperChange[] | null {
    return tryRenameWorkPaperSheetWithoutVisibilitySnapshots(this.runtimeAdapters.sheetRenameFastPathRuntime, oldName, newName)
  }

  protected captureTrackedChangesWithoutVisibilityCache(
    mutate: () => void,
    options: {
      readonly preservePendingTrackedPositions?: boolean
      readonly singleLiteralChange?: {
        readonly address: WorkPaperCellAddress
        readonly cellIndex?: number
        readonly isPhysicalSheet: boolean
        readonly sheetName: string
        readonly value: LiteralInput
      }
    } = {},
  ): WorkPaperChange[] {
    this.assertNotDisposed()
    this.ensureEngineEventTracking()
    if (this.engineEvents.hasPendingLazyChanges) {
      this.engineEvents.materializePendingLazyChanges(
        options.preservePendingTrackedPositions === undefined ? {} : { preservePositions: options.preservePendingTrackedPositions },
      )
    }
    if (this.engineEvents.hasTrackedEvents) {
      this.engineEvents.drain()
    }
    try {
      this.engineEvents.withRetainedIndices(mutate)
    } catch (error) {
      if (error instanceof Error && WORKPAPER_PUBLIC_ERROR_NAMES.has(error.name)) {
        throw error
      }
      throw new WorkPaperOperationError(this.messageOf(error, 'Mutation failed'))
    }
    const events = this.engineEvents.drain()
    if (events.length > 0 && events.every(trackedEventHasNoValueChanges)) {
      return []
    }
    const shouldEmitValuesUpdated = this.emitter.hasListeners('valuesUpdated')
    const event = events.length === 1 ? events[0] : undefined
    const preferLazyPublicChanges = shouldPreferLazyPublicChanges(events, shouldEmitValuesUpdated)
    const shouldReadDirectSingleLiteralEagerly =
      event === undefined || event.changedCellIndices.length <= TINY_TRACKED_CHANGE_LIMIT || !preferLazyPublicChanges
    if (shouldReadDirectSingleLiteralEagerly) {
      const directSingleLiteralChanges = tryBuildDirectSingleLiteralTrackedChange({
        events,
        ...(options.singleLiteralChange !== undefined ? { expected: options.singleLiteralChange } : {}),
        cellStore: this.engine.workbook.cellStore,
        strings: this.engine.strings,
        trackedA1: (row, col) => this.trackedA1(row, col),
      })
      if (directSingleLiteralChanges) {
        if (directSingleLiteralChanges.length > 0 && shouldEmitValuesUpdated) {
          this.emitter.emitDetailed({ eventName: 'valuesUpdated', payload: { changes: directSingleLiteralChanges } })
        }
        return directSingleLiteralChanges
      }
    }
    const changes = this.computeTrackedChangesWithoutVisibilityCache(events, {
      preferLazyPublicChanges,
    })
    if (changes.length > 0 && shouldEmitValuesUpdated) {
      this.emitter.emitDetailed({ eventName: 'valuesUpdated', payload: { changes } })
    }
    return changes
  }

  protected batchStructuralChanges(batchOperations: () => void): WorkPaperChange[] {
    this.engineEvents.materializePendingLazyChanges()
    if (!this.canUseTrackedStructuralFastPath()) {
      this.downgradeTrackedBatchFastPath()
      return this.batch(batchOperations)
    }
    this.assertNotDisposed()
    this.ensureEngineEventTracking()
    const undoStackStart = this.getUndoStack().length
    this.engineEvents.drain()
    this.batchDepth += 1
    try {
      this.engineEvents.withRetainedIndices(batchOperations)
    } finally {
      this.batchDepth -= 1
      this.flushPendingBatchOps()
      this.mergeUndoHistory(undoStackStart)
    }
    const shouldEmitValuesUpdated = this.emitter.hasListeners('valuesUpdated')
    const events = this.engineEvents.drain()
    const changes =
      events.length > 0 && events.every(trackedEventHasNoValueChanges)
        ? []
        : this.computeTrackedChangesWithoutVisibilityCache(events, {
            preferLazyPublicChanges: shouldPreferLazyPublicChanges(events, shouldEmitValuesUpdated),
          })
    this.flushQueuedEvents()
    if (changes.length > 0 && shouldEmitValuesUpdated) {
      this.emitter.emitDetailed({ eventName: 'valuesUpdated', payload: { changes } })
    }
    return changes
  }

  protected computeChangesAfterMutation(
    beforeVisibility: VisibilitySnapshot,
    beforeNames: NamedExpressionValueSnapshot,
  ): WorkPaperChange[] {
    const hasNamedExpressions = this.namedExpressions.size > 0
    const afterNames = hasNamedExpressions ? this.captureNamedExpressionValueSnapshot() : EMPTY_NAMED_EXPRESSION_VALUES
    const fastPath = this.computeCellChangesFromTrackedEvents(beforeVisibility, this.engineEvents.drain())
    let cellChanges: WorkPaperChange[]
    if (fastPath) {
      cellChanges = fastPath.changes
      this.visibilityCache = fastPath.nextVisibility
    } else {
      const afterVisibility = this.captureVisibilitySnapshot()
      cellChanges = this.computeCellChanges(beforeVisibility, afterVisibility)
      this.visibilityCache = afterVisibility
    }
    this.namedExpressionValueCache = afterNames
    return hasNamedExpressions ? [...cellChanges, ...this.computeNamedExpressionChanges(beforeNames, afterNames)] : cellChanges
  }

  protected captureChanges(
    semanticEvent: QueuedEvent | undefined,
    mutate: () => void,
    options: { readonly preservePendingTrackedPositions?: boolean } = {},
  ): WorkPaperChange[] {
    this.assertNotDisposed()
    this.engineEvents.materializePendingLazyChanges(
      options.preservePendingTrackedPositions === undefined ? {} : { preservePositions: options.preservePendingTrackedPositions },
    )
    this.downgradeTrackedBatchFastPath()
    if (semanticEvent !== undefined) {
      this.flushPendingBatchOps()
    }
    if (this.shouldSuppressEvents()) {
      try {
        mutate()
      } catch (error) {
        if (error instanceof Error && WORKPAPER_PUBLIC_ERROR_NAMES.has(error.name)) {
          throw error
        }
        throw new WorkPaperOperationError(this.messageOf(error, 'Mutation failed'))
      }
      if (semanticEvent) {
        this.queuedEvents.push(semanticEvent)
      }
      return []
    }
    const beforeVisibility = this.ensureVisibilityCache()
    const beforeNames = this.namedExpressions.size > 0 ? this.ensureNamedExpressionValueCache() : EMPTY_NAMED_EXPRESSION_VALUES
    this.ensureEngineEventTracking()
    this.engineEvents.drain()
    try {
      mutate()
    } catch (error) {
      if (error instanceof Error && WORKPAPER_PUBLIC_ERROR_NAMES.has(error.name)) {
        throw error
      }
      throw new WorkPaperOperationError(this.messageOf(error, 'Mutation failed'))
    }
    const changes =
      semanticEvent === undefined
        ? this.computeChangesAfterMutation(beforeVisibility, beforeNames)
        : (() => {
            const afterVisibility = this.captureVisibilitySnapshot()
            const afterNames = this.captureNamedExpressionValueSnapshot()
            this.visibilityCache = afterVisibility
            this.namedExpressionValueCache = afterNames
            return [
              ...this.computeCellChanges(beforeVisibility, afterVisibility),
              ...this.computeNamedExpressionChanges(beforeNames, afterNames),
            ]
          })()
    if (semanticEvent) {
      const event = withEventChanges(semanticEvent, changes)
      if (this.shouldSuppressEvents()) {
        this.queuedEvents.push(event)
      } else {
        this.emitter.emitDetailed(event)
      }
    }
    if (!this.shouldSuppressEvents() && changes.length > 0 && this.emitter.hasListeners('valuesUpdated')) {
      this.emitter.emitDetailed({ eventName: 'valuesUpdated', payload: { changes } })
    }
    return changes
  }

  protected shouldSuppressEvents(): boolean {
    return this.batchDepth > 0 || this.evaluationSuspended
  }

  protected flushQueuedEvents(): void {
    const events = [...this.queuedEvents]
    this.queuedEvents.length = 0
    events.forEach((event) => {
      this.emitter.emitDetailed(event)
    })
  }

  protected getUndoStack(): WorkPaperHistoryRecord[] {
    return readWorkPaperHistoryStack(this.engine, 'undoStack')
  }

  protected getRedoStack(): WorkPaperHistoryRecord[] {
    return readWorkPaperHistoryStack(this.engine, 'redoStack')
  }

  protected clearHistoryStacks(): void {
    clearWorkPaperHistoryStacks(this.getUndoStack(), this.getRedoStack())
  }

  protected mergeUndoHistory(startIndex: number): void {
    mergeWorkPaperUndoHistory(this.getUndoStack(), startIndex, (sheetId) => this.getSheetName(sheetId))
  }

  protected nextSheetName(): string {
    let index = 1
    while (this.doesSheetExist(`Sheet${index}`)) {
      index += 1
    }
    return `Sheet${index}`
  }

  protected collectRangeDependencies(
    range: WorkPaperCellRange,
    readDependencies: (address: WorkPaperCellAddress) => readonly string[],
  ): WorkPaperDependencyRef[] {
    return collectWorkPaperRangeDependencies({
      range,
      readDependencies,
      resolver: {
        defaultSheetName: () => this.listSheetRecords()[0]!.name,
        requireSheetId: (name) => this.requireSheetId(name),
      },
    })
  }

  protected toDependencyRefs(values: readonly string[]): WorkPaperDependencyRef[] {
    return toWorkPaperDependencyRefs(values, {
      defaultSheetName: () => this.listSheetRecords()[0]!.name,
      requireSheetId: (name) => this.requireSheetId(name),
    })
  }
}
