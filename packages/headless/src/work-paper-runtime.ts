import { SpreadsheetEngine, type EngineCellMutationRef, type SheetRecord } from '@bilig/core'
import { MAX_COLS, MAX_ROWS, type CellSnapshot, type CellValue, type WorkbookSnapshot } from '@bilig/protocol'
import {
  WorkPaperEvaluationTimeoutError,
  WorkPaperNamedExpressionDoesNotExistError,
  WorkPaperOperationError,
  WorkPaperSheetSizeLimitExceededError,
} from './work-paper-errors.js'
import {
  cloneConfig,
  canApplyRuntimeOnlyWorkPaperConfigUpdate,
  canReuseWorkPaperSnapshotRebuild,
  DEFAULT_CONFIG,
  validateWorkPaperConfig,
  WORKPAPER_CONFIG_KEYS,
} from './work-paper-config.js'
import { makeNamedExpressionKey } from './work-paper-runtime-helpers.js'
import { inspectSheetDimensionsWithinLimits, validateSheetWithinLimits } from './work-paper-sheet-inspection.js'
import { WorkPaperSheetDimensionCache } from './work-paper-sheet-dimension-cache.js'
import type { WorkPaperAxisIntervalEditMode, WorkPaperAxisKind } from './work-paper-axis-helpers.js'
import {
  createInternalNamedExpressionRecord,
  evaluateWorkPaperNamedExpression,
  validateWorkPaperNamedExpression,
  workPaperCellSnapshotToRawContent,
  workPaperNamedExpressionToDefinedNameSnapshot,
  type InternalNamedExpression,
  type WorkPaperNamedExpressionValueSnapshot,
} from './work-paper-named-expression-helpers.js'
import type {
  WorkPaperAxisInterval,
  WorkPaperCellAddress,
  WorkPaperCellRange,
  WorkPaperChange,
  WorkPaperConfig,
  WorkPaperSheet,
  WorkPaperSheets,
  WorkPaperInternals,
  RawCellContent,
  SerializedWorkPaperNamedExpression,
} from './work-paper-types.js'
import { WorkPaperEmitter } from './work-paper-emitter.js'
import { replaceWorkPaperSheetContent } from './work-paper-sheet-replacement.js'
import { restorePublicWorkPaperFormula, rewriteWorkPaperFormulaForStorage } from './work-paper-formula-rewrite.js'
import { applyWorkPaperMatrixContents, applyWorkPaperSerializedMatrix } from './work-paper-matrix-application.js'
import { cloneWorkPaperClipboardPayload, type WorkPaperClipboardPayload } from './work-paper-clipboard.js'
import type { QueuedEvent } from './work-paper-tracked-event-helpers.js'
import type { VisibilitySnapshot } from './work-paper-visibility-snapshot.js'
import {
  captureWorkPaperFunctionRegistry,
  clearWorkPaperFunctionBindings,
  type InternalFunctionBinding,
} from './work-paper-function-registry.js'
import { createWorkPaperInternals } from './work-paper-internals.js'
import {
  ensureWorkPaperCustomAdapterInstalled,
  getAllRegisteredWorkPaperFunctionPlugins,
  workPaperGlobalCustomFunctions,
} from './work-paper-static-registry.js'
import { initializeWorkPaperFromSheets, initializeWorkPaperFromSnapshot } from './work-paper-sheet-initialization.js'
import { buildWorkPaperRawCellMutation } from './work-paper-literal-mutation-queue.js'
import { getVisibleWorkPaperCellIndexInSheet } from './work-paper-cell-read.js'
import { WorkPaperMutationQueues } from './work-paper-mutation-queues.js'
import { WorkPaperEngineEventTracker } from './work-paper-engine-event-tracker.js'
import { createWorkPaperRuntimeAdapters } from './work-paper-runtime-adapters.js'
import { WorkPaperRuntimeSurface } from './work-paper-runtime-surface.js'
import type { WorkPaperHistoryRecord, WorkPaperHistoryTransactionRecord } from './work-paper-history.js'

type NamedExpressionValueSnapshot = WorkPaperNamedExpressionValueSnapshot

interface WorkPaperTransactionSnapshot {
  readonly clipboard: WorkPaperClipboardPayload | null
  readonly config: WorkPaperConfig
  readonly namedExpressions: readonly SerializedWorkPaperNamedExpression[]
  readonly redoStack: readonly WorkPaperHistoryRecord[]
  readonly sheets: WorkPaperSheets
  readonly undoStack: readonly WorkPaperHistoryRecord[]
}

let nextWorkbookId = 1

function workPaperEvaluationTimeoutErrorFrom(error: unknown): WorkPaperEvaluationTimeoutError | undefined {
  let current: unknown = error
  while (typeof current === 'object' && current !== null) {
    if (current instanceof WorkPaperEvaluationTimeoutError) {
      return current
    }
    const name = current instanceof Error ? current.name : undefined
    if (name === 'WorkPaperEvaluationTimeoutError' || name === 'EngineEvaluationTimeoutError') {
      const timeoutMs = Reflect.get(current, 'timeoutMs')
      return new WorkPaperEvaluationTimeoutError(typeof timeoutMs === 'number' ? timeoutMs : 0)
    }
    current = Reflect.get(current, 'cause')
  }
  return undefined
}

export class WorkPaper extends WorkPaperRuntimeSurface {
  readonly workbookId = nextWorkbookId++
  protected engine: SpreadsheetEngine
  protected readonly emitter = new WorkPaperEmitter()
  protected readonly namedExpressions = new Map<string, InternalNamedExpression>()
  protected readonly functionSnapshot = new Map<string, InternalFunctionBinding>()
  protected readonly functionAliasLookup = new Map<string, InternalFunctionBinding>()
  private readonly internalFunctionLookup = new Map<string, InternalFunctionBinding>()
  private readonly workPaperInternals: WorkPaperInternals
  protected readonly sheetDimensionCache: WorkPaperSheetDimensionCache
  protected config: WorkPaperConfig
  private clipboard: WorkPaperClipboardPayload | null = null
  protected visibilityCache: VisibilitySnapshot | null = null
  protected namedExpressionValueCache: NamedExpressionValueSnapshot | null = null
  protected sheetRecordsCache: readonly SheetRecord[] | null = null
  protected batchDepth = 0
  protected batchStartVisibility: VisibilitySnapshot | null = null
  protected batchStartNamedValues: NamedExpressionValueSnapshot | null = null
  protected batchUsesTrackedFastPath = false
  protected batchUndoStackLength = 0
  protected evaluationSuspended = false
  protected suspendedVisibility: VisibilitySnapshot | null = null
  protected suspendedNamedValues: NamedExpressionValueSnapshot | null = null
  protected suspendedUsesTrackedFastPath = false
  protected queuedEvents: QueuedEvent[] = []
  protected readonly engineEvents = new WorkPaperEngineEventTracker()
  protected disposed = false
  protected readonly mutationQueues = new WorkPaperMutationQueues({
    applyCellMutationsAtWithOptions: (refs, options) => {
      this.engine.applyCellMutationsAtWithOptions(refs, options)
    },
    updateSheetDimensionsAfterCellMutationRefs: (refs) => this.updateSheetDimensionsAfterCellMutationRefs(refs),
  })
  protected readonly runtimeAdapters = createWorkPaperRuntimeAdapters({
    getEngine: () => this.engine,
    getConfig: () => this.config,
    getBatchDepth: () => this.batchDepth,
    isEvaluationSuspended: () => this.evaluationSuspended,
    isTrackedBatchFastPathActive: () => this.batchUsesTrackedFastPath,
    getNamedExpressionValueCache: () => this.namedExpressionValueCache,
    getClipboard: () => this.clipboard,
    setClipboard: (clipboard) => {
      this.clipboard = clipboard
    },
    getNamedExpression: (name, scope) => this.namedExpressions.get(makeNamedExpressionKey(name, scope)),
    getNamedExpressionValues: () => this.namedExpressions.values(),
    setNamedExpressionRecord: (key, record) => {
      this.namedExpressions.set(key, record)
    },
    deleteNamedExpressionRecord: (name, scope) => {
      this.namedExpressions.delete(makeNamedExpressionKey(name, scope))
    },
    mutationQueues: this.mutationQueues,
    engineEvents: this.engineEvents,
    getSheetDimensionCache: () => this.sheetDimensionCache,
    emitter: this.emitter,
    assertNotDisposed: () => this.assertNotDisposed(),
    assertReadable: () => this.assertReadable(),
    prepareReadableState: () => this.prepareReadableState(),
    flushPendingBatchOps: () => this.flushPendingBatchOps(),
    downgradeTrackedBatchFastPath: () => this.downgradeTrackedBatchFastPath(),
    sheetRecord: (sheetId) => this.sheetRecord(sheetId),
    sheetName: (sheetId) => this.sheetName(sheetId),
    a1: (address) => this.a1(address),
    trackedA1: (row, col) => this.trackedA1(row, col),
    rangeRef: (range) => this.rangeRef(range),
    listSheetRecords: () => this.listSheetRecords(),
    requireSheetId: (name) => this.requireSheetId(name),
    restorePublicFormula: (formula, ownerSheetId) => this.restorePublicFormula(formula, ownerSheetId),
    cellSnapshotToRawContent: (cell, ownerSheetId) => this.cellSnapshotToRawContent(cell, ownerSheetId),
    canUseMetadataOnlySheetRenameFastPath: () => this.canUseMetadataOnlySheetRenameFastPath(),
    canUseNamedExpressionChangeFastPath: () => this.canUseNamedExpressionChangeFastPath(),
    canUseTrackedMutationFastPath: () => this.canUseTrackedMutationFastPath(),
    canUseTrackedStructuralFastPath: () => this.canUseTrackedStructuralFastPath(),
    getCapabilityContext: () => this.capabilityContext(),
    getVisibleCellIndexInSheet: (sheet, row, col) => this.getVisibleCellIndexInSheet(sheet, row, col),
    enqueueSuspendedLiteralMutation: (sheetId, row, col, content, cellIndex) =>
      this.enqueueSuspendedLiteralMutation(sheetId, row, col, content, cellIndex),
    enqueueDeferredBatchLiteral: (sheetId, row, col, content, cellIndex) =>
      this.enqueueDeferredBatchLiteral(sheetId, row, col, content, cellIndex),
    rewriteFormulaForStorage: (formula, ownerSheetId) => this.rewriteFormulaForStorage(formula, ownerSheetId),
    applyCellMutationRefs: (refs, options) => this.applyCellMutationRefs(refs, options),
    captureTrackedChangesWithoutVisibilityCache: (mutate, options) => this.captureTrackedChangesWithoutVisibilityCache(mutate, options),
    captureChanges: (event, mutate) => this.captureChanges(event, mutate),
    computeChangesAfterMutation: (beforeVisibility, beforeNames) => this.computeChangesAfterMutation(beforeVisibility, beforeNames),
    computeTrackedChangesWithoutVisibilityCache: (events, options) => this.computeTrackedChangesWithoutVisibilityCache(events, options),
    readSingleTrackedCellChange: (cellIndex) => this.readSingleTrackedCellChange(cellIndex),
    nextSheetName: () => this.nextSheetName(),
    isItPossibleToAddSheet: (name) => this.isItPossibleToAddSheet(name),
    ensureVisibilityCache: () => this.ensureVisibilityCache(),
    ensureNamedExpressionValueCache: () => this.ensureNamedExpressionValueCache(),
    clearSheetRecordsCache: () => {
      this.sheetRecordsCache = null
    },
    cacheSheetDimensions: (sheetId, dimensions) => this.sheetDimensionCache.cache(sheetId, dimensions),
    shouldSuppressEvents: () => this.shouldSuppressEvents(),
    queueSheetAddedEvent: (payload) => {
      this.queuedEvents.push({ eventName: 'sheetAdded', payload })
    },
    isItPossibleToRemoveSheet: (sheetId) => this.isItPossibleToRemoveSheet(sheetId),
    isItPossibleToClearSheet: (sheetId) => this.isItPossibleToClearSheet(sheetId),
    getSheetDimensions: (sheetId) => this.getSheetDimensions(sheetId),
    isItPossibleToReplaceSheetContent: (sheetId, content) => this.isItPossibleToReplaceSheetContent(sheetId, content),
    replaceSheetContentInternal: (sheetId, content, options) => this.replaceSheetContentInternal(sheetId, content, options),
    tryRenameSheetWithoutVisibilitySnapshots: (oldName, newName) => this.tryRenameSheetWithoutVisibilitySnapshots(oldName, newName),
    evaluateNamedExpression: (expression) => this.evaluateNamedExpression(expression),
    toDefinedNameSnapshot: (expression, scope) => this.toDefinedNameSnapshot(expression, scope),
    upsertNamedExpressionInternal: (expression, options) => this.upsertNamedExpressionInternal(expression, options),
    namedExpressionRecord: (name, scope) => this.namedExpressionRecord(name, scope),
    validateNamedExpression: (expressionName, expression, scope) => this.validateNamedExpression(expressionName, expression, scope),
    deleteDefinedName: (internalName) => {
      this.engine.deleteDefinedName(internalName)
    },
    isItPossibleToSetCellContents: (address, content) => this.isItPossibleToSetCellContents(address, content),
    applyMatrixContents: (address, content) => this.applyMatrixContents(address, content),
    getRangeSerialized: (range) => this.getRangeSerialized(range),
    getRangeValues: (range) => this.getRangeValues(range),
    batch: (operations) => this.batch(operations),
    setCellContents: (address, content) => this.setCellContents(address, content),
    applySerializedMatrix: (targetLeftCorner, content, sourceAnchor) => this.applySerializedMatrix(targetLeftCorner, content, sourceAnchor),
    doesSheetIdExist: (sheetId) => this.engine.workbook.getSheetById(sheetId) !== undefined,
    hasNamedExpression: (expressionName, scope) => this.namedExpressions.has(makeNamedExpressionKey(expressionName, scope)),
    canEditAxisIntervals: (axis, mode, sheetId, indexes) => this.canEditAxisIntervals(axis, mode, sheetId, indexes),
    batchStructuralChanges: (operations) => this.batchStructuralChanges(operations),
    captureAxisChange: (operations) => this.captureChanges(undefined, operations),
    applyAxisIntervalEdit: (axis, mode, sheetId, start, amount) => this.applyAxisIntervalEdit(axis, mode, sheetId, start, amount),
    applyAxisMove: (axis, sheetId, start, count, target) => this.applyAxisMove(axis, sheetId, start, count, target),
    messageOf: (error, fallback) => this.messageOf(error, fallback),
  })

  private getVisibleCellIndex(sheetId: number, row: number, col: number): number | undefined {
    const sheet = this.engine.workbook.getSheetById(sheetId)
    if (!sheet) {
      return undefined
    }
    return this.getVisibleCellIndexInSheet(sheet, row, col)
  }

  private getVisibleCellIndexInSheet(sheet: SheetRecord, row: number, col: number): number | undefined {
    return getVisibleWorkPaperCellIndexInSheet(sheet, row, col)
  }

  private constructor(configInput: WorkPaperConfig = {}) {
    super()
    ensureWorkPaperCustomAdapterInstalled()
    validateWorkPaperConfig(configInput)
    this.config = {
      ...cloneConfig(DEFAULT_CONFIG),
      ...cloneConfig(configInput),
    }
    this.engine = new SpreadsheetEngine({
      workbookName: 'Workbook',
      trackReplicaVersions: false,
      ...(this.config.useColumnIndex !== undefined ? { useColumnIndex: this.config.useColumnIndex } : {}),
      ...(this.config.evaluationTimeoutMs !== undefined ? { evaluationTimeoutMs: this.config.evaluationTimeoutMs } : {}),
    })
    this.sheetDimensionCache = new WorkPaperSheetDimensionCache(this.engine)
    this.sheetDimensionCache.invalidateAll()
    this.engineEvents.attach(this.engine)
    this.captureFunctionRegistry()
    this.workPaperInternals = createWorkPaperInternals({
      getCellDependents: (reference) => this.getCellDependents(reference),
      getCellPrecedents: (reference) => this.getCellPrecedents(reference),
      getRangeValues: (range) => this.getRangeValues(range),
      getRangeSerialized: (range) => this.getRangeSerialized(range),
      isCellPartOfArray: (address) => this.isCellPartOfArray(address),
      getCellFormula: (address) => this.getCellFormula(address),
      getSheetName: (sheetId) => this.getSheetName(sheetId),
      getSheetId: (name) => this.getSheetId(name),
      getSheetNames: () => this.getSheetNames(),
      countSheets: () => this.countSheets(),
      hasCellValueOrFormula: (address) => !this.isCellEmpty(address) || this.doesCellHaveFormula(address),
      getCellValue: (address) => this.getCellValue(address),
      recalculate: () => this.rebuildAndRecalculate(),
      calculateFormula: (formula, scope) => this.calculateFormula(formula, scope),
      getSheetDimensions: (sheetId) => this.getSheetDimensions(sheetId),
      normalizeFormula: (formula) => this.normalizeFormula(formula),
      validateFormula: (formula) => this.validateFormula(formula),
      getNamedExpressionsFromFormula: (formula) => this.getNamedExpressionsFromFormula(formula),
    })
  }

  get internals(): WorkPaperInternals {
    return this.workPaperInternals
  }

  static buildEmpty(configInput: WorkPaperConfig = {}, namedExpressions: readonly SerializedWorkPaperNamedExpression[] = []): WorkPaper {
    const workbook = new WorkPaper(configInput)
    workbook.engineEvents.withCaptureDisabled(() => {
      namedExpressions.forEach((expression) => {
        workbook.upsertNamedExpressionInternal(expression, { duringInitialization: true })
      })
    })
    workbook.clearHistoryStacks()
    workbook.resetChangeTrackingCaches()
    return workbook
  }

  static buildFromArray(
    sheet: WorkPaperSheet,
    configInput: WorkPaperConfig = {},
    namedExpressions: readonly SerializedWorkPaperNamedExpression[] = [],
  ): WorkPaper {
    return this.buildFromSheets({ Sheet1: sheet }, configInput, namedExpressions)
  }

  static buildFromSheets(
    sheets: WorkPaperSheets,
    configInput: WorkPaperConfig = {},
    namedExpressions: readonly SerializedWorkPaperNamedExpression[] = [],
  ): WorkPaper {
    const workbook = new WorkPaper(configInput)
    try {
      initializeWorkPaperFromSheets({
        engine: workbook.engine,
        config: workbook.config,
        sheets,
        namedExpressions,
        hasNamedExpressions: () => workbook.namedExpressions.size > 0,
        hasFunctionAliases: () => workbook.functionAliasLookup.size > 0 || workbook.internalFunctionLookup.size > 0,
        withEngineEventCaptureDisabled: (callback) => workbook.engineEvents.withCaptureDisabled(callback),
        upsertNamedExpression: (expression, options) => workbook.upsertNamedExpressionInternal(expression, options),
        rewriteFormulaForStorage: (formula, ownerSheetId) => workbook.rewriteFormulaForStorage(formula, ownerSheetId),
        requireSheetId: (name) => workbook.requireSheetId(name),
        cacheInitializedSheetDimensions: (sheetId, dimensions) => workbook.sheetDimensionCache.cacheInitialized(sheetId, dimensions),
        clearHistoryStacks: () => workbook.clearHistoryStacks(),
        resetChangeTrackingCaches: () => workbook.resetChangeTrackingCaches(),
      })
    } catch (error) {
      const timeoutError = workPaperEvaluationTimeoutErrorFrom(error)
      if (timeoutError) {
        throw timeoutError
      }
      throw error
    }
    return workbook
  }

  static buildFromSnapshot(snapshot: WorkbookSnapshot, configInput: WorkPaperConfig = {}): WorkPaper {
    const workbook = new WorkPaper(configInput)
    try {
      initializeWorkPaperFromSnapshot({
        engine: workbook.engine,
        config: workbook.config,
        snapshot,
        withEngineEventCaptureDisabled: (callback) => workbook.engineEvents.withCaptureDisabled(callback),
        requireSheetId: (name) => workbook.requireSheetId(name),
        cacheInitializedSheetDimensions: (sheetId, dimensions) => workbook.sheetDimensionCache.cacheInitialized(sheetId, dimensions),
        clearHistoryStacks: () => workbook.clearHistoryStacks(),
        resetChangeTrackingCaches: () => workbook.resetChangeTrackingCaches(),
      })
    } catch (error) {
      const timeoutError = workPaperEvaluationTimeoutErrorFrom(error)
      if (timeoutError) {
        throw timeoutError
      }
      throw error
    }
    return workbook
  }

  updateConfig(next: WorkPaperConfig): void {
    this.assertNotDisposed()
    this.engineEvents.materializePendingLazyChanges()
    const merged = {
      ...this.config,
      ...cloneConfig(next),
    }
    const changedKeys = WORKPAPER_CONFIG_KEYS.filter((key) => Object.hasOwn(next, key) && this.config[key] !== merged[key])
    if (changedKeys.length === 0) {
      return
    }
    validateWorkPaperConfig(merged)
    if (canApplyRuntimeOnlyWorkPaperConfigUpdate(changedKeys)) {
      this.applyRuntimeOnlyConfigUpdate(merged)
      return
    }
    this.rebuildWithConfig(merged)
  }

  transaction(operations: () => void): WorkPaperChange[] {
    this.assertNotDisposed()
    if (this.shouldSuppressEvents()) {
      throw new WorkPaperOperationError('WorkPaper transactions cannot run inside another suppressed mutation scope')
    }
    this.engineEvents.materializePendingLazyChanges()
    const snapshot = this.captureTransactionSnapshot()
    try {
      return this.batch(operations)
    } catch (error) {
      this.restoreTransactionSnapshot(snapshot)
      throw error
    }
  }

  private canEditAxisIntervals(
    axis: WorkPaperAxisKind,
    mode: WorkPaperAxisIntervalEditMode,
    sheetId: number,
    indexes: readonly WorkPaperAxisInterval[],
  ): boolean {
    if (axis === 'row') {
      return mode === 'add' ? this.isItPossibleToAddRows(sheetId, ...indexes) : this.isItPossibleToRemoveRows(sheetId, ...indexes)
    }
    return mode === 'add' ? this.isItPossibleToAddColumns(sheetId, ...indexes) : this.isItPossibleToRemoveColumns(sheetId, ...indexes)
  }

  private applyAxisIntervalEdit(
    axis: WorkPaperAxisKind,
    mode: WorkPaperAxisIntervalEditMode,
    sheetId: number,
    start: number,
    amount: number,
  ): void {
    if (axis === 'row') {
      if (mode === 'add') {
        this.engine.insertRows(this.sheetName(sheetId), start, amount)
      } else {
        this.engine.deleteRows(this.sheetName(sheetId), start, amount)
      }
    } else if (mode === 'add') {
      this.engine.insertColumns(this.sheetName(sheetId), start, amount)
    } else {
      this.engine.deleteColumns(this.sheetName(sheetId), start, amount)
    }
    this.sheetDimensionCache.updateAfterAxisIntervalEdit(axis, mode, sheetId, start, amount)
  }

  private applyAxisMove(axis: WorkPaperAxisKind, sheetId: number, start: number, count: number, target: number): void {
    if (axis === 'row') {
      this.engine.moveRows(this.sheetName(sheetId), start, count, target)
    } else {
      this.engine.moveColumns(this.sheetName(sheetId), start, count, target)
    }
    this.sheetDimensionCache.updateAfterAxisMove(axis, sheetId, start, count, target)
  }

  private updateSheetDimensionsAfterCellMutationRefs(refs: readonly EngineCellMutationRef[]): void {
    this.sheetDimensionCache.updateAfterCellMutationRefs(refs)
  }

  destroy(): void {
    this.dispose()
  }

  dispose(): void {
    if (this.disposed) {
      return
    }
    this.engineEvents.dispose()
    this.disposed = true
    this.emitter.clear()
    this.clearFunctionBindings()
    this.clipboard = null
    this.visibilityCache = null
    this.namedExpressionValueCache = null
    this.sheetDimensionCache.invalidateAll()
    this.queuedEvents = []
    this.namedExpressions.clear()
  }

  private applySerializedMatrix(
    targetLeftCorner: WorkPaperCellAddress,
    serialized: RawCellContent[][],
    sourceAnchor: WorkPaperCellAddress,
  ): void {
    applyWorkPaperSerializedMatrix({
      targetLeftCorner,
      serialized,
      sourceAnchor,
      flushPendingBatchOps: () => this.flushPendingBatchOps(),
      applyRawContent: (address, content) => this.applyRawContent(address, content),
      applyCellMutationRefs: (refs, options) => this.applyCellMutationRefs(refs, options),
      rewriteFormulaForStorage: (formula, ownerSheetId) => this.rewriteFormulaForStorage(formula, ownerSheetId),
    })
  }

  private captureTransactionSnapshot(): WorkPaperTransactionSnapshot {
    return {
      clipboard: cloneWorkPaperClipboardPayload(this.clipboard),
      config: cloneConfig(this.config),
      namedExpressions: this.getAllNamedExpressionsSerialized(),
      redoStack: cloneWorkPaperHistoryRecords(this.getRedoStack()),
      sheets: this.getAllSheetsSerialized(),
      undoStack: cloneWorkPaperHistoryRecords(this.getUndoStack()),
    }
  }

  private restoreTransactionSnapshot(snapshot: WorkPaperTransactionSnapshot): void {
    this.clearFunctionBindings({ preserveInternalFunctionLookup: true })
    this.namedExpressions.clear()
    this.engine = new SpreadsheetEngine({
      workbookName: 'Workbook',
      trackReplicaVersions: false,
      ...(snapshot.config.useColumnIndex !== undefined ? { useColumnIndex: snapshot.config.useColumnIndex } : {}),
      ...(snapshot.config.evaluationTimeoutMs !== undefined ? { evaluationTimeoutMs: snapshot.config.evaluationTimeoutMs } : {}),
    })
    this.engineEvents.attach(this.engine)
    this.config = cloneConfig(snapshot.config)
    this.captureFunctionRegistry()
    this.engineEvents.withCaptureDisabled(() => {
      initializeWorkPaperFromSheets({
        engine: this.engine,
        config: this.config,
        sheets: snapshot.sheets,
        namedExpressions: snapshot.namedExpressions,
        hasNamedExpressions: () => this.namedExpressions.size > 0,
        hasFunctionAliases: () => this.functionAliasLookup.size > 0 || this.internalFunctionLookup.size > 0,
        withEngineEventCaptureDisabled: (callback) => callback(),
        upsertNamedExpression: (expression, options) => this.upsertNamedExpressionInternal(expression, options),
        rewriteFormulaForStorage: (formula, ownerSheetId) => this.rewriteFormulaForStorage(formula, ownerSheetId),
        requireSheetId: (name) => this.requireSheetId(name),
        cacheInitializedSheetDimensions: (sheetId, dimensions) => this.sheetDimensionCache.cacheInitialized(sheetId, dimensions),
        clearHistoryStacks: () => this.clearHistoryStacks(),
        resetChangeTrackingCaches: () => this.resetChangeTrackingCaches(),
      })
    })
    this.resetTransactionRuntimeState()
    this.restoreHistoryStacks(snapshot)
    this.clipboard = cloneWorkPaperClipboardPayload(snapshot.clipboard)
  }

  private resetTransactionRuntimeState(): void {
    this.batchDepth = 0
    this.batchStartVisibility = null
    this.batchStartNamedValues = null
    this.batchUsesTrackedFastPath = false
    this.batchUndoStackLength = 0
    this.evaluationSuspended = false
    this.suspendedVisibility = null
    this.suspendedNamedValues = null
    this.suspendedUsesTrackedFastPath = false
    this.queuedEvents = []
    this.clearHistoryStacks()
    this.resetChangeTrackingCaches()
  }

  private restoreHistoryStacks(snapshot: WorkPaperTransactionSnapshot): void {
    const undoStack = this.getUndoStack()
    const redoStack = this.getRedoStack()
    undoStack.push(...cloneWorkPaperHistoryRecords(snapshot.undoStack))
    redoStack.push(...cloneWorkPaperHistoryRecords(snapshot.redoStack))
  }

  private applyMatrixContents(
    address: WorkPaperCellAddress,
    content: WorkPaperSheet,
    options: {
      captureUndo?: boolean
      deferLiteralAddresses?: ReadonlySet<string>
      skipNulls?: boolean
    } = {},
  ): void {
    applyWorkPaperMatrixContents({
      address,
      content,
      options,
      flushPendingBatchOps: () => this.flushPendingBatchOps(),
      applyCellMutationRefs: (refs, applyOptions) => this.applyCellMutationRefs(refs, applyOptions),
      rewriteFormulaForStorage: (formula, ownerSheetId) => this.rewriteFormulaForStorage(formula, ownerSheetId),
    })
  }

  private replaceSheetContentInternal(sheetId: number, content: WorkPaperSheet, options: { duringInitialization: boolean }): void {
    const dimensions = inspectSheetDimensionsWithinLimits(this.sheetName(sheetId), content, this.config)
    replaceWorkPaperSheetContent({
      sheetId,
      sheetName: this.sheetName(sheetId),
      content,
      duringInitialization: options.duringInitialization,
      listSpills: () => this.engine.workbook.listSpills(),
      getSheetDimensions: (nextSheetId) => this.getSheetDimensions(nextSheetId),
      clearRange: (input) => this.engine.clearRange(input),
      applyMatrixContents: (address, nextContent, applyOptions) => this.applyMatrixContents(address, nextContent, applyOptions),
      clearHistoryStacks: () => this.clearHistoryStacks(),
      getUndoStackLength: () => this.getUndoStack().length,
      mergeUndoHistory: (undoStackStart) => this.mergeUndoHistory(undoStackStart),
    })
    this.sheetDimensionCache.cacheInitialized(sheetId, dimensions)
  }

  private applyRawContent(address: WorkPaperCellAddress, content: RawCellContent): void {
    const cellIndex = this.getVisibleCellIndex(address.sheet, address.row, address.col)
    const mutation = buildWorkPaperRawCellMutation({
      row: address.row,
      col: address.col,
      content,
      rewriteFormulaForStorage: (formula) => this.rewriteFormulaForStorage(formula, address.sheet),
    })
    this.applyCellMutationRefs([{ sheetId: address.sheet, mutation, ...(cellIndex !== undefined ? { cellIndex } : {}) }], {
      captureUndo: true,
      potentialNewCells: content === null || cellIndex !== undefined ? 0 : 1,
      source: 'local',
      returnUndoOps: false,
      reuseRefs: true,
    })
  }

  protected createScratchWorkbook(config: WorkPaperConfig) {
    const temporaryWorkbook = new WorkPaper(config)
    return {
      engine: temporaryWorkbook.engine,
      registerNamedExpression: (expression: SerializedWorkPaperNamedExpression) => {
        temporaryWorkbook.upsertNamedExpressionInternal(expression, {
          duringInitialization: true,
        })
      },
      requireSheetId: (sheetName: string) => temporaryWorkbook.requireSheetId(sheetName),
      replaceSheetContent: (sheetId: number, sheet: WorkPaperSheet) => {
        temporaryWorkbook.replaceSheetContentInternal(sheetId, sheet, {
          duringInitialization: true,
        })
      },
      clearHistoryStacks: () => temporaryWorkbook.clearHistoryStacks(),
      applyRawContent: (address: WorkPaperCellAddress, content: RawCellContent) => temporaryWorkbook.applyRawContent(address, content),
      getRangeValues: (range: WorkPaperCellRange) => temporaryWorkbook.getRangeValues(range),
      getCellValue: (address: WorkPaperCellAddress) => temporaryWorkbook.getCellValue(address),
      dispose: () => temporaryWorkbook.dispose(),
    }
  }

  private captureFunctionRegistry(): void {
    captureWorkPaperFunctionRegistry({
      workbookId: this.workbookId,
      configFunctionPlugins: this.config.functionPlugins,
      plugins: getAllRegisteredWorkPaperFunctionPlugins(),
      functionSnapshot: this.functionSnapshot,
      functionAliasLookup: this.functionAliasLookup,
      internalFunctionLookup: this.internalFunctionLookup,
      globalCustomFunctions: workPaperGlobalCustomFunctions,
    })
  }

  private clearFunctionBindings(options: { preserveInternalFunctionLookup?: boolean } = {}): void {
    clearWorkPaperFunctionBindings({
      functionSnapshot: this.functionSnapshot,
      functionAliasLookup: this.functionAliasLookup,
      internalFunctionLookup: this.internalFunctionLookup,
      globalCustomFunctions: workPaperGlobalCustomFunctions,
      ...(options.preserveInternalFunctionLookup !== undefined
        ? { preserveInternalFunctionLookup: options.preserveInternalFunctionLookup }
        : {}),
    })
  }

  private validateCurrentSheetsWithinLimits(nextConfig: WorkPaperConfig): void {
    this.listSheetRecords().forEach((sheet) => {
      const dimensions = this.getSheetDimensions(sheet.id)
      if (dimensions.height > (nextConfig.maxRows ?? MAX_ROWS) || dimensions.width > (nextConfig.maxColumns ?? MAX_COLS)) {
        throw new WorkPaperSheetSizeLimitExceededError()
      }
    })
  }

  private applyRuntimeOnlyConfigUpdate(nextConfig: WorkPaperConfig): void {
    if (this.config.useColumnIndex !== nextConfig.useColumnIndex) {
      ;(this.engine as SpreadsheetEngine & { setUseColumnIndexEnabled(enabled: boolean): void }).setUseColumnIndexEnabled(
        nextConfig.useColumnIndex ?? false,
      )
    }
    if (this.config.evaluationTimeoutMs !== nextConfig.evaluationTimeoutMs) {
      this.engine.setEvaluationTimeoutMs(nextConfig.evaluationTimeoutMs)
    }
    this.config = cloneConfig(nextConfig)
  }

  private rebuildWithConfig(nextConfig: WorkPaperConfig): void {
    this.validateCurrentSheetsWithinLimits(nextConfig)
    const canReuseSnapshot = canReuseWorkPaperSnapshotRebuild(this.config, nextConfig)
    const snapshot = canReuseSnapshot ? this.engine.exportSnapshot() : null
    const serializedSheets = canReuseSnapshot ? null : this.getAllSheetsSerialized()
    if (serializedSheets) {
      Object.entries(serializedSheets).forEach(([sheetName, sheet]) => {
        validateSheetWithinLimits(sheetName, sheet, nextConfig)
      })
    }
    const serializedNamedExpressions = canReuseSnapshot ? null : this.getAllNamedExpressionsSerialized()
    const suspended = this.evaluationSuspended
    const clipboard = cloneWorkPaperClipboardPayload(this.clipboard)

    this.clearFunctionBindings({ preserveInternalFunctionLookup: true })
    if (!canReuseSnapshot) {
      this.namedExpressions.clear()
    }
    this.engine = new SpreadsheetEngine({
      workbookName: 'Workbook',
      trackReplicaVersions: false,
      ...(nextConfig.useColumnIndex !== undefined ? { useColumnIndex: nextConfig.useColumnIndex } : {}),
      ...(nextConfig.evaluationTimeoutMs !== undefined ? { evaluationTimeoutMs: nextConfig.evaluationTimeoutMs } : {}),
    })
    this.engineEvents.attach(this.engine)
    this.config = cloneConfig(nextConfig)
    this.captureFunctionRegistry()

    this.engineEvents.withCaptureDisabled(() => {
      if (snapshot) {
        this.engine.importSnapshot(snapshot)
      } else {
        try {
          initializeWorkPaperFromSheets({
            engine: this.engine,
            config: this.config,
            sheets: serializedSheets!,
            namedExpressions: serializedNamedExpressions!,
            hasNamedExpressions: () => this.namedExpressions.size > 0,
            hasFunctionAliases: () => this.functionAliasLookup.size > 0 || this.internalFunctionLookup.size > 0,
            withEngineEventCaptureDisabled: (callback) => callback(),
            upsertNamedExpression: (expression, options) => this.upsertNamedExpressionInternal(expression, options),
            rewriteFormulaForStorage: (formula, ownerSheetId) => this.rewriteFormulaForStorage(formula, ownerSheetId),
            requireSheetId: (name) => this.requireSheetId(name),
            cacheInitializedSheetDimensions: (sheetId, dimensions) => this.sheetDimensionCache.cacheInitialized(sheetId, dimensions),
            clearHistoryStacks: () => this.clearHistoryStacks(),
            resetChangeTrackingCaches: () => this.resetChangeTrackingCaches(),
          })
        } catch (error) {
          const timeoutError = workPaperEvaluationTimeoutErrorFrom(error)
          if (timeoutError) {
            throw timeoutError
          }
          throw error
        }
      }
    })
    this.clearHistoryStacks()
    this.resetChangeTrackingCaches()
    this.clipboard = clipboard
    if (suspended) {
      this.suspendedVisibility = this.ensureVisibilityCache()
      this.suspendedNamedValues = this.ensureNamedExpressionValueCache()
    }
  }

  private rewriteFormulaForStorage(formula: string, ownerSheetId: number): string {
    return rewriteWorkPaperFormulaForStorage({
      formula,
      ownerSheetId,
      namedExpressions: this.namedExpressions,
      functionAliasLookup: this.functionAliasLookup,
      messageOf: (error, fallback) => this.messageOf(error, fallback),
    })
  }

  private restorePublicFormula(formula: string, ownerSheetId: number): string {
    return restorePublicWorkPaperFormula({
      formula,
      ownerSheetId,
      namedExpressions: this.namedExpressions,
      internalFunctionLookup: this.internalFunctionLookup,
    })
  }

  private validateNamedExpression(expressionName: string, expression: RawCellContent, scope?: number): void {
    validateWorkPaperNamedExpression({
      expressionName,
      expression,
      ...(scope !== undefined ? { scope } : {}),
      requireScope: (sheetId) => {
        this.sheetRecord(sheetId)
      },
      messageOf: (error, fallback) => this.messageOf(error, fallback),
    })
  }

  private upsertNamedExpressionInternal(
    expression: SerializedWorkPaperNamedExpression,
    options: { duringInitialization: boolean; skipValidation?: boolean },
  ): void {
    if (options.skipValidation !== true) {
      this.validateNamedExpression(expression.name, expression.expression, expression.scope)
    }
    const record = createInternalNamedExpressionRecord(expression)
    this.namedExpressions.set(makeNamedExpressionKey(record.publicName, record.scope), record)
    this.engine.setDefinedName(record.internalName, this.toDefinedNameSnapshot(record.expression, record.scope))
    if (options.duringInitialization) {
      this.clearHistoryStacks()
    }
  }

  private toDefinedNameSnapshot(expression: RawCellContent, scope?: number) {
    return workPaperNamedExpressionToDefinedNameSnapshot({
      expression,
      ...(scope !== undefined ? { scope } : {}),
      defaultScopeId: this.listSheetRecords()[0]?.id ?? 1,
      rewriteFormulaForStorage: (formula, ownerSheetId) => this.rewriteFormulaForStorage(formula, ownerSheetId),
    })
  }

  private namedExpressionRecord(name: string, scope?: number): InternalNamedExpression {
    const direct = this.namedExpressions.get(makeNamedExpressionKey(name, scope))
    if (direct) {
      return direct
    }
    throw new WorkPaperNamedExpressionDoesNotExistError(name)
  }

  protected evaluateNamedExpression(expression: InternalNamedExpression): CellValue | CellValue[][] {
    return evaluateWorkPaperNamedExpression(expression, (formula, scope) => this.calculateFormula(formula, scope))
  }

  private cellSnapshotToRawContent(cell: CellSnapshot, ownerSheetId: number): RawCellContent {
    return workPaperCellSnapshotToRawContent({
      cell,
      ownerSheetId,
      restorePublicFormula: (formula, sheetId) => this.restorePublicFormula(formula, sheetId),
    })
  }

  protected messageOf(error: unknown, fallback: string): string {
    return error instanceof Error && error.message.length > 0 ? error.message : fallback
  }
}

function cloneWorkPaperHistoryRecords(records: readonly WorkPaperHistoryRecord[]): WorkPaperHistoryRecord[] {
  return records.map((record) => ({
    forward: cloneWorkPaperHistoryTransactionRecord(record.forward),
    inverse: cloneWorkPaperHistoryTransactionRecord(record.inverse),
  }))
}

function cloneWorkPaperHistoryTransactionRecord(record: WorkPaperHistoryTransactionRecord): WorkPaperHistoryTransactionRecord {
  switch (record.kind) {
    case 'ops':
      return {
        kind: 'ops',
        ops: record.ops.map((op) => structuredClone(op)),
        ...(record.potentialNewCells !== undefined ? { potentialNewCells: record.potentialNewCells } : {}),
      }
    case 'single-op':
      return {
        kind: 'single-op',
        op: structuredClone(record.op),
        ...(record.potentialNewCells !== undefined ? { potentialNewCells: record.potentialNewCells } : {}),
      }
    case 'single-existing-numeric-cell-mutation':
      return {
        kind: 'single-existing-numeric-cell-mutation',
        sheetId: record.sheetId,
        row: record.row,
        col: record.col,
        cellIndex: record.cellIndex,
        value: record.value,
        ...(record.potentialNewCells !== undefined ? { potentialNewCells: record.potentialNewCells } : {}),
      }
    case 'cell-mutations':
      return {
        kind: 'cell-mutations',
        refs: record.refs.map((ref) => ({
          sheetId: ref.sheetId,
          mutation: { ...ref.mutation },
        })),
        ...(record.potentialNewCells !== undefined ? { potentialNewCells: record.potentialNewCells } : {}),
      }
  }
}
