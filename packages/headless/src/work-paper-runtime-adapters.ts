import type { EngineCellMutationRef, SheetRecord, SpreadsheetEngine } from '@bilig/core/headless-runtime'
import type { CellRangeRef, CellSnapshot } from '@bilig/protocol'
import { orderWorkPaperCellChanges } from './change-order.js'
import { isWorkPaperAxisOrderPossible, isWorkPaperAxisSwapPossible, isWorkPaperMoveAxisPossible } from './work-paper-capability-checks.js'
import {
  editWorkPaperAxisIntervals,
  moveWorkPaperAxis,
  setWorkPaperAxisOrder,
  swapWorkPaperAxisIndexes,
  type WorkPaperAxisEditRuntime,
} from './work-paper-axis-helpers.js'
import { setWorkPaperCellContents, type WorkPaperSetCellContentsRuntime } from './work-paper-cell-content-setter.js'
import type { applyWorkPaperCellMutationRefs, WorkPaperCellMutationApplyRuntime } from './work-paper-cell-mutation-refs.js'
import {
  trySetExistingLiteralWorkPaperCellContentsWithTrackedFastPath,
  trySetExistingNumericWorkPaperCellContentsWithTrackedFastPath,
  type WorkPaperExistingNumericFastPathRuntime,
} from './work-paper-existing-numeric-fast-path.js'
import {
  tryCaptureWorkPaperNamedExpressionChangeWithoutSnapshots,
  type WorkPaperNamedExpressionFastPathRuntime,
} from './work-paper-named-expression-fast-path.js'
import type { WorkPaperSheetRenameFastPathRuntime } from './work-paper-sheet-rename-fast-path.js'
import { createWorkPaperReadOperations, type WorkPaperReadOperations } from './work-paper-read-operations.js'
import { createWorkPaperSheetOperations, type WorkPaperSheetOperations } from './work-paper-sheet-operations.js'
import { createWorkPaperClipboardOperations, type WorkPaperClipboardOperations } from './work-paper-clipboard-operations.js'
import {
  createWorkPaperNamedExpressionOperations,
  type WorkPaperNamedExpressionOperations,
} from './work-paper-named-expression-operations.js'
import { createWorkPaperCapabilityOperations, type WorkPaperCapabilityOperations } from './work-paper-capability-operations.js'
import type { WorkPaperMutationQueues } from './work-paper-mutation-queues.js'
import type { WorkPaperEngineEventTracker } from './work-paper-engine-event-tracker.js'
import type { WorkPaperSheetDimensionCache } from './work-paper-sheet-dimension-cache.js'
import type { WorkPaperEmitter } from './work-paper-emitter.js'
import type { WorkPaperClipboardPayload } from './work-paper-clipboard.js'
import type { InternalNamedExpression, WorkPaperNamedExpressionValueSnapshot } from './work-paper-named-expression-helpers.js'
import type { VisibilitySnapshot } from './work-paper-visibility-snapshot.js'
import type { TrackedEngineEvent } from './tracked-engine-event-refs.js'
import type { QueuedEvent } from './work-paper-tracked-event-helpers.js'
import type {
  RawCellContent,
  SerializedWorkPaperNamedExpression,
  WorkPaperAxisInterval,
  WorkPaperCellAddress,
  WorkPaperCellChange,
  WorkPaperCellRange,
  WorkPaperChange,
  WorkPaperConfig,
  WorkPaperDetailedEventMap,
  WorkPaperSheet,
  WorkPaperSheetDimensions,
} from './work-paper-types.js'

export interface WorkPaperRuntimeAdapters {
  readonly cellMutationApplyRuntime: WorkPaperCellMutationApplyRuntime
  readonly sheetRenameFastPathRuntime: WorkPaperSheetRenameFastPathRuntime
  readonly namedExpressionFastPathRuntime: WorkPaperNamedExpressionFastPathRuntime
  readonly existingNumericFastPathRuntime: WorkPaperExistingNumericFastPathRuntime
  readonly setCellContentsRuntime: WorkPaperSetCellContentsRuntime
  readonly readOperations: WorkPaperReadOperations
  readonly sheetOperations: WorkPaperSheetOperations
  readonly clipboardOperations: WorkPaperClipboardOperations
  readonly capabilityOperations: WorkPaperCapabilityOperations
  readonly namedExpressionOperations: WorkPaperNamedExpressionOperations
  readonly axisEditRuntime: WorkPaperAxisEditRuntime
}

export interface CreateWorkPaperRuntimeAdaptersArgs {
  readonly getEngine: () => SpreadsheetEngine
  readonly getConfig: () => WorkPaperConfig
  readonly getBatchDepth: () => number
  readonly isEvaluationSuspended: () => boolean
  readonly isTrackedBatchFastPathActive: () => boolean
  readonly getNamedExpressionValueCache: () => WorkPaperNamedExpressionValueSnapshot | null
  readonly getClipboard: () => WorkPaperClipboardPayload | null
  readonly setClipboard: (clipboard: WorkPaperClipboardPayload | null) => void
  readonly getNamedExpression: (name: string, scope: number | undefined) => InternalNamedExpression | undefined
  readonly getNamedExpressionValues: () => IterableIterator<InternalNamedExpression>
  readonly setNamedExpressionRecord: (key: string, record: InternalNamedExpression) => void
  readonly deleteNamedExpressionRecord: (name: string, scope: number | undefined) => void
  readonly mutationQueues: WorkPaperMutationQueues
  readonly engineEvents: WorkPaperEngineEventTracker
  readonly getSheetDimensionCache: () => WorkPaperSheetDimensionCache
  readonly emitter: WorkPaperEmitter
  readonly assertNotDisposed: () => void
  readonly assertReadable: () => void
  readonly prepareReadableState: () => void
  readonly flushPendingBatchOps: () => void
  readonly downgradeTrackedBatchFastPath: () => void
  readonly sheetRecord: (sheetId: number) => SheetRecord
  readonly sheetName: (sheetId: number) => string
  readonly a1: (address: Pick<WorkPaperCellAddress, 'row' | 'col'>) => string
  readonly trackedA1: (row: number, col: number) => string
  readonly rangeRef: (range: WorkPaperCellRange) => CellRangeRef
  readonly listSheetRecords: () => readonly SheetRecord[]
  readonly requireSheetId: (name: string) => number
  readonly restorePublicFormula: (formula: string, ownerSheetId: number) => string
  readonly cellSnapshotToRawContent: (cell: CellSnapshot, ownerSheetId: number) => RawCellContent
  readonly canUseMetadataOnlySheetRenameFastPath: () => boolean
  readonly canUseNamedExpressionChangeFastPath: () => boolean
  readonly canUseTrackedMutationFastPath: () => boolean
  readonly canUseTrackedStructuralFastPath: () => boolean
  readonly getCapabilityContext: () => {
    readonly config: WorkPaperConfig
    readonly requireSheet: (sheetId: number) => void
    readonly doesSheetExist: (sheetName: string) => boolean
    readonly getSheetIdByName: (sheetName: string) => number | undefined
  }
  readonly getVisibleCellIndexInSheet: (sheet: SheetRecord, row: number, col: number) => number | undefined
  readonly enqueueSuspendedLiteralMutation: (
    sheetId: number,
    row: number,
    col: number,
    content: RawCellContent,
    cellIndex: number | undefined,
  ) => boolean
  readonly enqueueDeferredBatchLiteral: (
    sheetId: number,
    row: number,
    col: number,
    content: RawCellContent,
    cellIndex: number | undefined,
  ) => boolean
  readonly rewriteFormulaForStorage: (formula: string, ownerSheetId: number) => string
  readonly applyCellMutationRefs: (
    refs: readonly EngineCellMutationRef[],
    options: Parameters<typeof applyWorkPaperCellMutationRefs>[2],
  ) => void
  readonly captureTrackedChangesWithoutVisibilityCache: (
    mutate: () => void,
    options?: Parameters<WorkPaperExistingNumericFastPathRuntime['computeTrackedChangesWithoutVisibilityCache']>[1] & {
      readonly preservePendingTrackedPositions?: boolean
      readonly singleLiteralChange?: {
        readonly address: WorkPaperCellAddress
        readonly cellIndex?: number
        readonly isPhysicalSheet: boolean
        readonly sheetName: string
        readonly value: RawCellContent
      }
    },
  ) => WorkPaperChange[]
  readonly captureChanges: (event: QueuedEvent | undefined, mutate: () => void) => WorkPaperChange[]
  readonly computeChangesAfterMutation: (
    beforeVisibility: VisibilitySnapshot,
    beforeNames: WorkPaperNamedExpressionValueSnapshot,
  ) => WorkPaperChange[]
  readonly computeTrackedChangesWithoutVisibilityCache: (
    events: readonly TrackedEngineEvent[],
    options?: Parameters<WorkPaperExistingNumericFastPathRuntime['computeTrackedChangesWithoutVisibilityCache']>[1],
  ) => WorkPaperChange[]
  readonly readSingleTrackedCellChange: (cellIndex: number) => WorkPaperCellChange | undefined
  readonly nextSheetName: () => string
  readonly isItPossibleToAddSheet: (name: string) => boolean
  readonly ensureVisibilityCache: () => VisibilitySnapshot
  readonly ensureNamedExpressionValueCache: () => WorkPaperNamedExpressionValueSnapshot
  readonly clearSheetRecordsCache: () => void
  readonly cacheSheetDimensions: (sheetId: number, dimensions: WorkPaperSheetDimensions) => void
  readonly shouldSuppressEvents: () => boolean
  readonly queueSheetAddedEvent: (payload: WorkPaperDetailedEventMap['sheetAdded']) => void
  readonly isItPossibleToRemoveSheet: (sheetId: number) => boolean
  readonly isItPossibleToClearSheet: (sheetId: number) => boolean
  readonly getSheetDimensions: (sheetId: number) => WorkPaperSheetDimensions
  readonly isItPossibleToReplaceSheetContent: (sheetId: number, content: WorkPaperSheet) => boolean
  readonly replaceSheetContentInternal: (
    sheetId: number,
    content: WorkPaperSheet,
    options: { readonly duringInitialization: boolean },
  ) => void
  readonly tryRenameSheetWithoutVisibilitySnapshots: (oldName: string, newName: string) => WorkPaperChange[] | null
  readonly evaluateNamedExpression: (
    expression: InternalNamedExpression,
  ) => ReturnType<WorkPaperNamedExpressionFastPathRuntime['evaluateNamedExpression']>
  readonly toDefinedNameSnapshot: WorkPaperNamedExpressionFastPathRuntime['toDefinedNameSnapshot']
  readonly upsertNamedExpressionInternal: (
    expression: SerializedWorkPaperNamedExpression,
    options: { duringInitialization: boolean; skipValidation?: boolean },
  ) => void
  readonly namedExpressionRecord: (name: string, scope: number | undefined) => InternalNamedExpression
  readonly validateNamedExpression: (expressionName: string, expression: RawCellContent, scope?: number) => void
  readonly deleteDefinedName: (internalName: string) => void
  readonly isItPossibleToSetCellContents: (address: WorkPaperCellAddress, content?: RawCellContent | WorkPaperSheet) => boolean
  readonly applyMatrixContents: (address: WorkPaperCellAddress, content: WorkPaperSheet) => void
  readonly getRangeSerialized: (range: WorkPaperCellRange) => RawCellContent[][]
  readonly getRangeValues: (range: WorkPaperCellRange) => ReturnType<WorkPaperReadOperations['getRangeValues']>
  readonly batch: (operations: () => void) => WorkPaperChange[]
  readonly setCellContents: (address: WorkPaperCellAddress, content: RawCellContent | WorkPaperSheet) => WorkPaperChange[]
  readonly applySerializedMatrix: (
    targetLeftCorner: WorkPaperCellAddress,
    content: RawCellContent[][],
    sourceAnchor: WorkPaperCellAddress,
  ) => void
  readonly doesSheetIdExist: (sheetId: number) => boolean
  readonly hasNamedExpression: (expressionName: string, scope?: number) => boolean
  readonly canEditAxisIntervals: (
    axis: 'row' | 'column',
    mode: 'add' | 'remove',
    sheetId: number,
    intervals: readonly WorkPaperAxisInterval[],
  ) => boolean
  readonly batchStructuralChanges: (operations: () => void) => WorkPaperChange[]
  readonly captureAxisChange: (operations: () => void) => WorkPaperChange[]
  readonly applyAxisIntervalEdit: (axis: 'row' | 'column', mode: 'add' | 'remove', sheetId: number, start: number, amount: number) => void
  readonly applyAxisMove: (axis: 'row' | 'column', sheetId: number, start: number, count: number, target: number) => void
  readonly messageOf: (error: unknown, fallback: string) => string
}

export function createWorkPaperRuntimeAdapters(args: CreateWorkPaperRuntimeAdaptersArgs): WorkPaperRuntimeAdapters {
  const cellMutationApplyRuntime: WorkPaperCellMutationApplyRuntime = {
    isEvaluationSuspended: args.isEvaluationSuspended,
    appendSuspendedCellMutationRefs: (refs) => {
      args.mutationQueues.appendSuspendedCellMutationRefs(refs)
    },
    addSuspendedCellMutationPotentialNewCells: (amount) => {
      args.mutationQueues.addSuspendedCellMutationPotentialNewCells(amount)
    },
    applyCellMutationsAtWithOptions: (refs, options) => {
      args.getEngine().applyCellMutationsAtWithOptions(refs, options)
    },
    updateSheetDimensionsAfterCellMutationRefs: (refs) => args.getSheetDimensionCache().updateAfterCellMutationRefs(refs),
  }

  const sheetRenameFastPathRuntime: WorkPaperSheetRenameFastPathRuntime = {
    canUseMetadataOnlySheetRenameFastPath: args.canUseMetadataOnlySheetRenameFastPath,
    assertNotDisposed: args.assertNotDisposed,
    hasPendingLazyTrackedChanges: () => args.engineEvents.hasPendingLazyChanges,
    materializePendingLazyTrackedChanges: () => args.engineEvents.materializePendingLazyChanges(),
    isTrackedBatchFastPathActive: args.isTrackedBatchFastPathActive,
    downgradeTrackedBatchFastPath: args.downgradeTrackedBatchFastPath,
    hasTrackedEngineEvents: () => args.engineEvents.hasTrackedEvents,
    drainTrackedEngineEvents: () => {
      void args.engineEvents.drain()
    },
    clearTrackedEngineEvents: () => {
      args.engineEvents.clearEvents()
    },
    clearSheetRecordsCache: args.clearSheetRecordsCache,
    renameSheetMetadataOnly: (oldName, newName) => args.getEngine().renameSheetMetadataOnly(oldName, newName),
    renameSheet: (oldName, newName) => args.getEngine().renameSheet(oldName, newName),
    withEngineEventCaptureDisabled: (callback) => {
      args.engineEvents.withCaptureDisabled(callback)
    },
    messageOf: args.messageOf,
  }

  const namedExpressionFastPathRuntime: WorkPaperNamedExpressionFastPathRuntime = {
    canUseNamedExpressionChangeFastPath: args.canUseNamedExpressionChangeFastPath,
    assertNotDisposed: args.assertNotDisposed,
    materializePendingLazyTrackedChanges: () => args.engineEvents.materializePendingLazyChanges(),
    downgradeTrackedBatchFastPath: args.downgradeTrackedBatchFastPath,
    getCachedNamedExpressionValue: (key) => args.getNamedExpressionValueCache()?.get(key),
    setCachedNamedExpressionValue: (key, value) => {
      args.getNamedExpressionValueCache()?.set(key, value)
    },
    evaluateNamedExpression: args.evaluateNamedExpression,
    hasAnyListeners: () => args.emitter.hasAnyListeners(),
    toDefinedNameSnapshot: args.toDefinedNameSnapshot,
    upsertNumericDefinedNameFast: (name, value, numericValue) => args.getEngine().upsertNumericDefinedNameFast(name, value, numericValue),
    setNamedExpressionRecord: args.setNamedExpressionRecord,
    readSingleTrackedCellChange: args.readSingleTrackedCellChange,
    orderCellChanges: (changes, explicitChangedCount) => orderWorkPaperCellChanges(changes, args.listSheetRecords(), explicitChangedCount),
    drainTrackedEngineEvents: () => args.engineEvents.drain(),
    withRetainedTrackedEngineEventIndices: (callback) => {
      args.engineEvents.withRetainedIndices(callback)
    },
    upsertNamedExpressionInternal: (expression) => {
      args.upsertNamedExpressionInternal(expression, { duringInitialization: false })
    },
    namedExpressionRecord: args.namedExpressionRecord,
    computeTrackedChangesWithoutVisibilityCache: args.computeTrackedChangesWithoutVisibilityCache,
    messageOf: args.messageOf,
  }

  const existingNumericFastPathRuntime: WorkPaperExistingNumericFastPathRuntime = {
    canUseTrackedMutationFastPath: args.canUseTrackedMutationFastPath,
    getEngine: args.getEngine,
    hasPendingLazyTrackedChanges: () => args.engineEvents.hasPendingLazyChanges,
    materializePendingLazyTrackedChanges: () => args.engineEvents.materializePendingLazyChanges(),
    hasTrackedEngineEvents: () => args.engineEvents.hasTrackedEvents,
    drainTrackedEngineEvents: () => args.engineEvents.drain(),
    clearTrackedEngineEvents: () => {
      args.engineEvents.clearEvents()
    },
    getEngineEventCaptureEnabled: () => args.engineEvents.isCaptureEnabled,
    setEngineEventCaptureEnabled: (enabled) => {
      args.engineEvents.setCaptureEnabled(enabled)
    },
    hasPendingBatchOps: () => args.mutationQueues.hasPendingBatchOps(),
    flushPendingBatchOps: args.flushPendingBatchOps,
    messageOf: args.messageOf,
    trackedA1: args.trackedA1,
    orderChanges: (changes, explicitChangedCount) => orderWorkPaperCellChanges(changes, args.listSheetRecords(), explicitChangedCount),
    computeTrackedChangesWithoutVisibilityCache: args.computeTrackedChangesWithoutVisibilityCache,
    trackLazyChanges: (changes) => {
      args.engineEvents.trackLazyChanges(changes)
    },
    hasValuesUpdatedListeners: () => args.emitter.hasListeners('valuesUpdated'),
    emitValuesUpdated: (changes) => {
      args.emitter.emitDetailed({ eventName: 'valuesUpdated', payload: { changes } })
    },
  }

  const setCellContentsRuntime: WorkPaperSetCellContentsRuntime = {
    assertNotDisposed: args.assertNotDisposed,
    getConfig: args.getConfig,
    getEngine: args.getEngine,
    sheetRecord: args.sheetRecord,
    getVisibleCellIndexInSheet: args.getVisibleCellIndexInSheet,
    isEvaluationSuspended: args.isEvaluationSuspended,
    getBatchDepth: args.getBatchDepth,
    enqueueSuspendedLiteralMutation: args.enqueueSuspendedLiteralMutation,
    enqueueDeferredBatchLiteral: args.enqueueDeferredBatchLiteral,
    trySetExistingNumericCellContentsWithTrackedFastPath: (request) =>
      trySetExistingNumericWorkPaperCellContentsWithTrackedFastPath(existingNumericFastPathRuntime, request),
    trySetExistingLiteralCellContentsWithTrackedFastPath: (request) =>
      trySetExistingLiteralWorkPaperCellContentsWithTrackedFastPath(existingNumericFastPathRuntime, request),
    flushPendingBatchOps: args.flushPendingBatchOps,
    rewriteFormulaForStorage: args.rewriteFormulaForStorage,
    applyCellMutationRefs: args.applyCellMutationRefs,
    canUseTrackedMutationFastPath: args.canUseTrackedMutationFastPath,
    isTrackedBatchFastPathActive: args.isTrackedBatchFastPathActive,
    captureTrackedChangesWithoutVisibilityCache: args.captureTrackedChangesWithoutVisibilityCache,
    captureChanges: (mutate) => args.captureChanges(undefined, mutate),
    isItPossibleToSetCellContents: args.isItPossibleToSetCellContents,
    applyMatrixContents: args.applyMatrixContents,
  }

  const readOperations = createWorkPaperReadOperations({
    getEngine: args.getEngine,
    getSheetDimensionCache: args.getSheetDimensionCache,
    assertReadable: args.assertReadable,
    assertNotDisposed: args.assertNotDisposed,
    prepareReadableState: args.prepareReadableState,
    flushPendingBatchOps: args.flushPendingBatchOps,
    sheetRecord: args.sheetRecord,
    sheetName: args.sheetName,
    a1: args.a1,
    rangeRef: args.rangeRef,
    listSheetRecords: args.listSheetRecords,
    requireSheetId: args.requireSheetId,
    restorePublicFormula: args.restorePublicFormula,
    cellSnapshotToRawContent: args.cellSnapshotToRawContent,
  })

  const sheetOperations = createWorkPaperSheetOperations({
    assertNotDisposed: args.assertNotDisposed,
    materializePendingLazyChanges: () => args.engineEvents.materializePendingLazyChanges(),
    nextSheetName: args.nextSheetName,
    isItPossibleToAddSheet: args.isItPossibleToAddSheet,
    ensureVisibilityCache: args.ensureVisibilityCache,
    ensureNamedExpressionValueCache: args.ensureNamedExpressionValueCache,
    drainEngineEvents: () => {
      void args.engineEvents.drain()
    },
    createSheet: (name) => {
      args.getEngine().createSheet(name)
    },
    clearSheetRecordsCache: args.clearSheetRecordsCache,
    requireSheetId: args.requireSheetId,
    cacheSheetDimensions: args.cacheSheetDimensions,
    shouldSuppressEvents: args.shouldSuppressEvents,
    queueSheetAddedEvent: args.queueSheetAddedEvent,
    emitSheetAdded: (payload) => {
      args.emitter.emitDetailed({ eventName: 'sheetAdded', payload })
    },
    computeChangesAfterMutation: args.computeChangesAfterMutation,
    hasValuesUpdatedListeners: () => args.emitter.hasListeners('valuesUpdated'),
    emitValuesUpdated: (changes) => {
      args.emitter.emitDetailed({ eventName: 'valuesUpdated', payload: { changes } })
    },
    isItPossibleToRemoveSheet: args.isItPossibleToRemoveSheet,
    sheetName: args.sheetName,
    captureChanges: args.captureChanges,
    deleteSheet: (name) => {
      args.getEngine().deleteSheet(name)
    },
    invalidateSheetDimensions: (sheetId) => args.getSheetDimensionCache().invalidate(sheetId),
    isItPossibleToClearSheet: args.isItPossibleToClearSheet,
    getSheetDimensions: args.getSheetDimensions,
    clearRange: (range) => {
      args.getEngine().clearRange(range)
    },
    isItPossibleToReplaceSheetContent: args.isItPossibleToReplaceSheetContent,
    replaceSheetContentInternal: args.replaceSheetContentInternal,
    sheetRecord: args.sheetRecord,
    getSheetByName: (name) => args.getEngine().workbook.getSheet(name),
    tryRenameSheetWithoutVisibilitySnapshots: args.tryRenameSheetWithoutVisibilitySnapshots,
    renameSheet: (oldName, newName) => {
      args.getEngine().renameSheet(oldName, newName)
    },
  })

  const clipboardOperations = createWorkPaperClipboardOperations({
    assertReadable: args.assertReadable,
    assertNotDisposed: args.assertNotDisposed,
    getClipboard: args.getClipboard,
    setClipboard: args.setClipboard,
    getRangeSerialized: args.getRangeSerialized,
    getRangeValues: args.getRangeValues,
    batch: args.batch,
    setCellContents: args.setCellContents,
    captureChanges: (mutate) => args.captureChanges(undefined, mutate),
    applySerializedMatrix: args.applySerializedMatrix,
  })

  const capabilityOperations = createWorkPaperCapabilityOperations({
    assertNotDisposed: args.assertNotDisposed,
    getCapabilityContext: args.getCapabilityContext,
    doesSheetIdExist: args.doesSheetIdExist,
    validateNamedExpression: args.validateNamedExpression,
    hasNamedExpression: args.hasNamedExpression,
  })

  const namedExpressionOperations = createWorkPaperNamedExpressionOperations({
    assertReadable: args.assertReadable,
    getNamedExpression: args.getNamedExpression,
    getNamedExpressionValues: args.getNamedExpressionValues,
    evaluateNamedExpression: args.evaluateNamedExpression,
    isItPossibleToAddNamedExpression: (expressionName, expression, scope) =>
      capabilityOperations.isItPossibleToAddNamedExpression(expressionName, expression, scope),
    isItPossibleToRemoveNamedExpression: (expressionName, scope) =>
      capabilityOperations.isItPossibleToRemoveNamedExpression(expressionName, scope),
    validateNamedExpression: args.validateNamedExpression,
    tryCaptureNamedExpressionChangeWithoutSnapshots: (existing, expressionName, expression, scope, options) =>
      tryCaptureWorkPaperNamedExpressionChangeWithoutSnapshots(namedExpressionFastPathRuntime, {
        existing,
        expressionName,
        expression,
        ...(scope !== undefined ? { scope } : {}),
        ...(options !== undefined ? { options } : {}),
      }),
    captureChanges: args.captureChanges,
    upsertNamedExpressionInternal: args.upsertNamedExpressionInternal,
    deleteNamedExpressionRecord: args.deleteNamedExpressionRecord,
    deleteDefinedName: args.deleteDefinedName,
  })

  const axisEditRuntime: WorkPaperAxisEditRuntime = {
    canSwapAxisIndexes: (axis, sheetId, mappings) => isWorkPaperAxisSwapPossible(args.getCapabilityContext(), axis, sheetId, mappings),
    canSetAxisOrder: (axis, sheetId, order) => isWorkPaperAxisOrderPossible(args.getCapabilityContext(), axis, sheetId, order),
    canEditAxisIntervals: args.canEditAxisIntervals,
    canMoveAxis: (axis, sheetId, start, count, target) =>
      isWorkPaperMoveAxisPossible(args.getCapabilityContext(), axis, sheetId, start, count, target),
    canUseTrackedStructuralFastPath: args.canUseTrackedStructuralFastPath,
    isTrackedBatchFastPathActive: args.isTrackedBatchFastPathActive,
    batch: args.batch,
    batchStructuralChanges: args.batchStructuralChanges,
    captureAxisChange: args.captureAxisChange,
    captureTrackedStructuralChanges: args.captureTrackedChangesWithoutVisibilityCache,
    moveAxis: (axis, sheetId, start, count, target) => moveWorkPaperAxis(axisEditRuntime, axis, sheetId, start, count, target),
    applyAxisIntervalEdit: args.applyAxisIntervalEdit,
    applyAxisMove: args.applyAxisMove,
  }

  return {
    cellMutationApplyRuntime,
    sheetRenameFastPathRuntime,
    namedExpressionFastPathRuntime,
    existingNumericFastPathRuntime,
    setCellContentsRuntime,
    readOperations,
    sheetOperations,
    clipboardOperations,
    capabilityOperations,
    namedExpressionOperations,
    axisEditRuntime,
  }
}

export const workPaperRuntimeAdapterCommands = {
  editWorkPaperAxisIntervals,
  setWorkPaperAxisOrder,
  setWorkPaperCellContents,
  swapWorkPaperAxisIndexes,
}
