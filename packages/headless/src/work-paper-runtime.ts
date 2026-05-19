import type { EngineCellMutationRef, SheetRecord, SpreadsheetEngine } from '@bilig/core'
import { MAX_COLS, MAX_ROWS, type CellSnapshot, type CellValue, type WorkbookSnapshot } from '@bilig/protocol'
import {
  WorkPaperInvalidArgumentsError,
  WorkPaperNamedExpressionDoesNotExistError,
  WorkPaperOperationError,
  WorkPaperSheetError,
  WorkPaperSheetSizeLimitExceededError,
} from './work-paper-errors.js'
import {
  cloneConfig,
  canApplyRuntimeOnlyWorkPaperConfigUpdate,
  canReuseWorkPaperSnapshotRebuild,
  DEFAULT_CONFIG,
  normalizeConfiguredWorkPaperCalculationSettings,
  validateWorkPaperConfig,
  WORKPAPER_CONFIG_KEYS,
  WORKPAPER_PUBLIC_ERROR_NAMES,
} from './work-paper-config.js'
import { assertRowAndColumn, makeNamedExpressionKey } from './work-paper-runtime-helpers.js'
import {
  inspectSheetDimensionsWithinLimits,
  validateSheetWithinLimits,
  workPaperSheetHasDynamicSpillFormula,
} from './work-paper-sheet-inspection.js'
import { WorkPaperSheetDimensionCache } from './work-paper-sheet-dimension-cache.js'
import type { WorkPaperAxisIntervalEditMode, WorkPaperAxisKind } from './work-paper-axis-helpers.js'
import {
  createInternalNamedExpressionRecord,
  evaluateWorkPaperNamedExpression,
  trySimpleWorkPaperNamedExpressionDefinedNameSnapshot,
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
  hasRegisteredWorkPaperFunctionPlugins,
  workPaperGlobalCustomFunctions,
} from './work-paper-static-registry.js'
import {
  initializeWorkPaperFromSheetEntries,
  initializeWorkPaperFromSheets,
  initializeWorkPaperFromSnapshot,
} from './work-paper-sheet-initialization.js'
import { buildWorkPaperRawCellMutation } from './work-paper-literal-mutation-queue.js'
import { WorkPaperMutationQueues } from './work-paper-mutation-queues.js'
import { WorkPaperEngineEventTracker } from './work-paper-engine-event-tracker.js'
import { WorkPaperRuntimeFastPathBase } from './work-paper-runtime-fast-path-base.js'
import { cloneWorkPaperHistoryRecords } from './work-paper-history.js'
import { tryChangeSimpleNumericNamedExpressionFastPath } from './work-paper-named-expression-fast-path-runtime.js'
import {
  createWorkPaperEngine,
  workPaperEvaluationTimeoutErrorFrom,
  type WorkPaperTransactionSnapshot,
} from './work-paper-runtime-construction.js'

type NamedExpressionValueSnapshot = WorkPaperNamedExpressionValueSnapshot

let nextWorkbookId = 1
type MetadataRenameEngine = SpreadsheetEngine & {
  readonly renameSheetMetadataOnlyById?: (sheetId: number, newName: string) => boolean
}

type WorkPaperStructuralInsertEngine = SpreadsheetEngine & {
  insertRows(sheetName: string, start: number, count: number, options?: { readonly emitTracked?: boolean }): void
  insertColumns(sheetName: string, start: number, count: number, options?: { readonly emitTracked?: boolean }): void
}

export class WorkPaper extends WorkPaperRuntimeFastPathBase {
  readonly workbookId = nextWorkbookId++
  protected engine: SpreadsheetEngine
  protected readonly emitter = new WorkPaperEmitter()
  protected readonly namedExpressions = new Map<string, InternalNamedExpression>()
  protected readonly functionSnapshot = new Map<string, InternalFunctionBinding>()
  protected readonly functionAliasLookup = new Map<string, InternalFunctionBinding>()
  protected readonly internalFunctionLookup = new Map<string, InternalFunctionBinding>()
  private workPaperInternalsCache: WorkPaperInternals | undefined
  protected sheetDimensionCache: WorkPaperSheetDimensionCache
  protected config: WorkPaperConfig
  protected clipboard: WorkPaperClipboardPayload | null = null
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
  private engineEventsAttached = false
  protected disposed = false
  protected readonly mutationQueues = new WorkPaperMutationQueues({
    applyCellMutationsAtWithOptions: (refs, options) => {
      this.engine.applyCellMutationsAtWithOptions(refs, options)
    },
    updateSheetDimensionsAfterCellMutationRefs: (refs) => this.updateSheetDimensionsAfterCellMutationRefs(refs),
  })
  private constructor(configInput: WorkPaperConfig = {}) {
    super()
    validateWorkPaperConfig(configInput)
    this.config = {
      ...cloneConfig(DEFAULT_CONFIG),
      ...cloneConfig(configInput),
    }
    const hasFunctionPlugins = (this.config.functionPlugins?.length ?? 0) > 0 || hasRegisteredWorkPaperFunctionPlugins()
    if (hasFunctionPlugins) {
      ensureWorkPaperCustomAdapterInstalled()
    }
    this.engine = createWorkPaperEngine(this.config)
    this.sheetDimensionCache = new WorkPaperSheetDimensionCache(this.engine)
    if (hasFunctionPlugins) {
      this.captureFunctionRegistry()
    }
  }

  protected ensureEngineEventTracking(): void {
    if (!this.engineEventsAttached) {
      this.engineEvents.attach(this.engine)
      this.engineEventsAttached = true
    }
  }

  get internals(): WorkPaperInternals {
    this.workPaperInternalsCache ??= createWorkPaperInternals({
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
    return this.workPaperInternalsCache
  }

  override renameSheet(sheetId: number, nextName: string): WorkPaperChange[] {
    const sheet = this.sheetRecord(sheetId)
    const newName = nextName.trim()
    if (newName.length === 0) {
      throw new WorkPaperInvalidArgumentsError('Sheet name must be non-empty')
    }
    const existing = this.engine.workbook.getSheet(newName)
    if (existing && existing.id !== sheetId) {
      throw new WorkPaperSheetError(`Sheet '${sheetId}' cannot be renamed to '${nextName}'`)
    }

    const oldName = sheet.name
    const fastPathChanges = this.tryRenameSheetWithoutRuntimeAdapters(sheet, newName)
    if (fastPathChanges !== null) {
      return fastPathChanges
    }

    return this.captureChanges(
      {
        eventName: 'sheetRenamed',
        payload: {
          sheetId,
          oldName,
          newName,
        },
      },
      () => {
        this.engine.renameSheet(oldName, newName)
        this.sheetRecordsCache = null
      },
    )
  }

  override changeNamedExpression(
    expressionName: string,
    expression: RawCellContent,
    scope?: number,
    options?: Record<string, string | number | boolean>,
  ): WorkPaperChange[] {
    if (options === undefined) {
      const fastPathChanges = this.tryChangeSimpleNumericNamedExpression(expressionName, expression, scope)
      if (fastPathChanges !== null) {
        return fastPathChanges
      }
    }
    return super.changeNamedExpression(expressionName, expression, scope, options)
  }

  override addRows(sheetId: number, start: number, count?: number): WorkPaperChange[]
  override addRows(sheetId: number, ...indexes: readonly WorkPaperAxisInterval[]): WorkPaperChange[]
  override addRows(
    sheetId: number,
    startOrInterval: number | WorkPaperAxisInterval,
    countOrInterval?: number | WorkPaperAxisInterval,
    ...restIntervals: readonly WorkPaperAxisInterval[]
  ): WorkPaperChange[] {
    if (
      typeof startOrInterval === 'number' &&
      (countOrInterval === undefined || typeof countOrInterval === 'number') &&
      restIntervals.length === 0
    ) {
      return this.addSingleAxisIntervalWithoutRuntimeAdapters('row', sheetId, startOrInterval, countOrInterval ?? 1)
    }
    return this.editAxisIntervalsWithoutRuntimeAdapters('row', 'add', sheetId, startOrInterval, countOrInterval, restIntervals)
  }

  override removeRows(sheetId: number, start: number, count?: number): WorkPaperChange[]
  override removeRows(sheetId: number, ...indexes: readonly WorkPaperAxisInterval[]): WorkPaperChange[]
  override removeRows(
    sheetId: number,
    startOrInterval: number | WorkPaperAxisInterval,
    countOrInterval?: number | WorkPaperAxisInterval,
    ...restIntervals: readonly WorkPaperAxisInterval[]
  ): WorkPaperChange[] {
    return this.editAxisIntervalsWithoutRuntimeAdapters('row', 'remove', sheetId, startOrInterval, countOrInterval, restIntervals)
  }

  override addColumns(sheetId: number, start: number, count?: number): WorkPaperChange[]
  override addColumns(sheetId: number, ...indexes: readonly WorkPaperAxisInterval[]): WorkPaperChange[]
  override addColumns(
    sheetId: number,
    startOrInterval: number | WorkPaperAxisInterval,
    countOrInterval?: number | WorkPaperAxisInterval,
    ...restIntervals: readonly WorkPaperAxisInterval[]
  ): WorkPaperChange[] {
    if (
      typeof startOrInterval === 'number' &&
      (countOrInterval === undefined || typeof countOrInterval === 'number') &&
      restIntervals.length === 0
    ) {
      return this.addSingleAxisIntervalWithoutRuntimeAdapters('column', sheetId, startOrInterval, countOrInterval ?? 1)
    }
    return this.editAxisIntervalsWithoutRuntimeAdapters('column', 'add', sheetId, startOrInterval, countOrInterval, restIntervals)
  }

  override removeColumns(sheetId: number, start: number, count?: number): WorkPaperChange[]
  override removeColumns(sheetId: number, ...indexes: readonly WorkPaperAxisInterval[]): WorkPaperChange[]
  override removeColumns(
    sheetId: number,
    startOrInterval: number | WorkPaperAxisInterval,
    countOrInterval?: number | WorkPaperAxisInterval,
    ...restIntervals: readonly WorkPaperAxisInterval[]
  ): WorkPaperChange[] {
    return this.editAxisIntervalsWithoutRuntimeAdapters('column', 'remove', sheetId, startOrInterval, countOrInterval, restIntervals)
  }

  private tryChangeSimpleNumericNamedExpression(
    expressionName: string,
    expression: RawCellContent,
    scope: number | undefined,
  ): WorkPaperChange[] | null {
    return tryChangeSimpleNumericNamedExpressionFastPath({
      assertNotDisposed: () => this.assertNotDisposed(),
      canUseNamedExpressionChangeFastPath: () => this.canUseNamedExpressionChangeFastPath(),
      downgradeTrackedBatchFastPath: () => this.downgradeTrackedBatchFastPath(),
      engine: this.engine,
      listSheetRecords: () => this.listSheetRecords(),
      materializePendingLazyChanges: () => this.engineEvents.materializePendingLazyChanges(),
      messageOf: (error, fallback) => this.messageOf(error, fallback),
      namedExpressionValueCache: this.namedExpressionValueCache,
      namedExpressions: this.namedExpressions,
      publicErrorNames: WORKPAPER_PUBLIC_ERROR_NAMES,
      readSingleTrackedCellChange: (cellIndex) => this.readSingleTrackedCellChange(cellIndex),
      toDefinedNameSnapshot: (nextExpression, nextScope) => this.toDefinedNameSnapshot(nextExpression, nextScope),
      validateNamedExpression: (nextName, nextExpression, nextScope) => this.validateNamedExpression(nextName, nextExpression, nextScope),
      expressionName,
      expression,
      scope,
    })
  }

  private addSingleAxisIntervalWithoutRuntimeAdapters(
    axis: WorkPaperAxisKind,
    sheetId: number,
    start: number,
    amount: number,
  ): WorkPaperChange[] {
    this.assertNotDisposed()
    const sheet = this.sheetRecord(sheetId)
    const limit = axis === 'row' ? (this.config.maxRows ?? MAX_ROWS) : (this.config.maxColumns ?? MAX_COLS)
    assertRowAndColumn(start, 'start')
    assertRowAndColumn(amount, 'count')
    if (amount <= 0 || start + amount > limit) {
      throw new WorkPaperOperationError(`${axis === 'row' ? 'Rows' : 'Columns'} cannot be added`)
    }
    if (this.batchDepth === 0 && !this.evaluationSuspended && this.visibilityCache === null && this.namedExpressions.size === 0) {
      if (this.engineEvents.hasPendingLazyChanges) {
        this.engineEvents.materializePendingLazyChanges()
      }
      if (this.engineEvents.hasTrackedEvents) {
        this.engineEvents.drain()
      }
      try {
        const structuralInsertEngine = this.engine as WorkPaperStructuralInsertEngine
        if (axis === 'row') {
          structuralInsertEngine.insertRows(sheet.name, start, amount, { emitTracked: false })
        } else {
          structuralInsertEngine.insertColumns(sheet.name, start, amount, { emitTracked: false })
        }
        this.sheetDimensionCache.updateAfterAxisIntervalEdit(axis, 'add', sheet.id, start, amount)
      } catch (error) {
        if (error instanceof Error && WORKPAPER_PUBLIC_ERROR_NAMES.has(error.name)) {
          throw error
        }
        throw new WorkPaperOperationError(this.messageOf(error, 'Mutation failed'))
      }
      return []
    }
    if (this.batchUsesTrackedFastPath) {
      this.applyAxisIntervalEditForSheet(axis, 'add', sheet, start, amount)
      return []
    }
    return this.batchStructuralChanges(() => {
      this.applyAxisIntervalEditForSheet(axis, 'add', sheet, start, amount)
    })
  }

  private tryRenameSheetWithoutRuntimeAdapters(sheet: SheetRecord, newName: string): WorkPaperChange[] | null {
    if (!this.canUseMetadataOnlySheetRenameFastPath()) {
      return null
    }
    const oldName = sheet.name
    this.assertNotDisposed()
    if (this.engineEvents.hasPendingLazyChanges) {
      this.engineEvents.materializePendingLazyChanges()
    }
    if (this.batchUsesTrackedFastPath) {
      this.downgradeTrackedBatchFastPath()
    }
    if (!this.canUseMetadataOnlySheetRenameFastPath()) {
      return null
    }
    if (this.engineEvents.hasTrackedEvents) {
      this.engineEvents.drain()
    }
    try {
      const metadataRenameEngine = this.engine as MetadataRenameEngine
      const renamed =
        metadataRenameEngine.renameSheetMetadataOnlyById?.(sheet.id, newName) ?? this.engine.renameSheetMetadataOnly(oldName, newName)
      if (renamed) {
        this.sheetRecordsCache = null
        this.engineEvents.clearEvents()
      } else {
        this.engineEvents.withCaptureDisabled(() => {
          this.engine.renameSheet(oldName, newName)
          this.sheetRecordsCache = null
        })
      }
    } catch (error) {
      if (error instanceof Error && WORKPAPER_PUBLIC_ERROR_NAMES.has(error.name)) {
        throw error
      }
      throw new WorkPaperOperationError(this.messageOf(error, 'Mutation failed'))
    }
    if (this.engineEvents.hasTrackedEvents) {
      this.engineEvents.drain()
    }
    return []
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
        cacheInitializedSheetDimensions: (sheetId, dimensions, options) =>
          workbook.sheetDimensionCache.cacheInitialized(sheetId, dimensions, options),
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

  static buildFromSheetEntries(
    sheetEntries: readonly (readonly [string, WorkPaperSheet])[],
    configInput: WorkPaperConfig = {},
    namedExpressions: readonly SerializedWorkPaperNamedExpression[] = [],
  ): WorkPaper {
    const workbook = new WorkPaper(configInput)
    try {
      initializeWorkPaperFromSheetEntries({
        engine: workbook.engine,
        config: workbook.config,
        sheetEntries,
        namedExpressions,
        hasNamedExpressions: () => workbook.namedExpressions.size > 0,
        hasFunctionAliases: () => workbook.functionAliasLookup.size > 0 || workbook.internalFunctionLookup.size > 0,
        withEngineEventCaptureDisabled: (callback) => workbook.engineEvents.withCaptureDisabled(callback),
        upsertNamedExpression: (expression, options) => workbook.upsertNamedExpressionInternal(expression, options),
        rewriteFormulaForStorage: (formula, ownerSheetId) => workbook.rewriteFormulaForStorage(formula, ownerSheetId),
        requireSheetId: (name) => workbook.requireSheetId(name),
        cacheInitializedSheetDimensions: (sheetId, dimensions, options) =>
          workbook.sheetDimensionCache.cacheInitialized(sheetId, dimensions, options),
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
        cacheInitializedSheetDimensions: (sheetId, dimensions, options) =>
          workbook.sheetDimensionCache.cacheInitialized(sheetId, dimensions, options),
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

  protected applyCalculationSettings(settings: WorkPaperConfig['calculationSettings']): void {
    this.captureChanges(undefined, () => {
      const normalized = normalizeConfiguredWorkPaperCalculationSettings(settings, this.engine.getCalculationSettings())
      this.engine.setCalculationSettings(normalized ?? this.engine.getCalculationSettings())
      this.config = {
        ...this.config,
        ...(settings === undefined ? { calculationSettings: undefined } : { calculationSettings: structuredClone(settings) }),
      }
    })
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

  protected canEditAxisIntervals(
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

  protected applyAxisIntervalEdit(
    axis: WorkPaperAxisKind,
    mode: WorkPaperAxisIntervalEditMode,
    sheetId: number,
    start: number,
    amount: number,
    options: { readonly emitTracked?: boolean } = {},
  ): void {
    this.applyAxisIntervalEditForSheet(axis, mode, this.sheetRecord(sheetId), start, amount, options)
  }

  protected applyAxisIntervalEditForSheet(
    axis: WorkPaperAxisKind,
    mode: WorkPaperAxisIntervalEditMode,
    sheet: SheetRecord,
    start: number,
    amount: number,
    options: { readonly emitTracked?: boolean } = {},
  ): void {
    const sheetName = sheet.name
    const structuralInsertEngine = this.engine as WorkPaperStructuralInsertEngine
    if (axis === 'row') {
      if (mode === 'add') {
        structuralInsertEngine.insertRows(sheetName, start, amount, options)
      } else {
        this.engine.deleteRows(sheetName, start, amount)
      }
    } else if (mode === 'add') {
      structuralInsertEngine.insertColumns(sheetName, start, amount, options)
    } else {
      this.engine.deleteColumns(sheetName, start, amount)
    }
    this.sheetDimensionCache.updateAfterAxisIntervalEdit(axis, mode, sheet.id, start, amount)
  }

  protected applyAxisMove(axis: WorkPaperAxisKind, sheetId: number, start: number, count: number, target: number): void {
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

  protected applySerializedMatrix(
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
    this.replaceEngineForConfig(snapshot.config)
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
        cacheInitializedSheetDimensions: (sheetId, dimensions, options) =>
          this.sheetDimensionCache.cacheInitialized(sheetId, dimensions, options),
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

  protected applyMatrixContents(
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
      isEvaluationSuspended: () => this.evaluationSuspended,
      applyCellMutationRefs: (refs, applyOptions) => this.applyCellMutationRefs(refs, applyOptions),
      rewriteFormulaForStorage: (formula, ownerSheetId) => this.rewriteFormulaForStorage(formula, ownerSheetId),
      updateSheetDimensionsAfterCellMutationRefs: (refs) => this.updateSheetDimensionsAfterCellMutationRefs(refs),
      updateSheetDimensionsAfterMatrixMutationImpact: (impact) => this.sheetDimensionCache.updateAfterMatrixMutationImpact(impact),
    })
  }

  protected replaceSheetContentInternal(sheetId: number, content: WorkPaperSheet, options: { duringInitialization: boolean }): void {
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
    this.sheetDimensionCache.cacheInitialized(sheetId, dimensions, {
      mayResizeDynamically: workPaperSheetHasDynamicSpillFormula(content),
    })
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
    this.replaceEngineForConfig(nextConfig)
    this.config = cloneConfig(nextConfig)
    this.captureFunctionRegistry()

    this.engineEvents.withCaptureDisabled(() => {
      if (snapshot) {
        this.engine.importSnapshot(snapshot)
        const calculationSettings = normalizeConfiguredWorkPaperCalculationSettings(
          this.config.calculationSettings,
          this.engine.getCalculationSettings(),
        )
        if (calculationSettings !== undefined) {
          this.engine.setCalculationSettings(calculationSettings)
        }
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
            cacheInitializedSheetDimensions: (sheetId, dimensions, options) =>
              this.sheetDimensionCache.cacheInitialized(sheetId, dimensions, options),
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

  protected rewriteFormulaForStorage(formula: string, ownerSheetId: number): string {
    return rewriteWorkPaperFormulaForStorage({
      formula,
      ownerSheetId,
      namedExpressions: this.namedExpressions,
      functionAliasLookup: this.functionAliasLookup,
      messageOf: (error, fallback) => this.messageOf(error, fallback),
    })
  }

  private replaceEngineForConfig(config: WorkPaperConfig): void {
    this.engineEvents.detach()
    this.engineEventsAttached = false
    this.engine = createWorkPaperEngine(config)
    this.sheetDimensionCache = new WorkPaperSheetDimensionCache(this.engine)
  }

  protected restorePublicFormula(formula: string, ownerSheetId: number): string {
    return restorePublicWorkPaperFormula({
      formula,
      ownerSheetId,
      namedExpressions: this.namedExpressions,
      internalFunctionLookup: this.internalFunctionLookup,
    })
  }

  protected validateNamedExpression(expressionName: string, expression: RawCellContent, scope?: number): void {
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

  protected upsertNamedExpressionInternal(
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

  protected toDefinedNameSnapshot(expression: RawCellContent, scope?: number) {
    const simpleSnapshot = trySimpleWorkPaperNamedExpressionDefinedNameSnapshot(expression)
    if (simpleSnapshot !== undefined) {
      return simpleSnapshot
    }
    return workPaperNamedExpressionToDefinedNameSnapshot({
      expression,
      ...(scope !== undefined ? { scope } : {}),
      defaultScopeId: this.listSheetRecords()[0]?.id ?? 1,
      rewriteFormulaForStorage: (formula, ownerSheetId) => this.rewriteFormulaForStorage(formula, ownerSheetId),
    })
  }

  protected namedExpressionRecord(name: string, scope?: number): InternalNamedExpression {
    const direct = this.namedExpressions.get(makeNamedExpressionKey(name, scope))
    if (direct) {
      return direct
    }
    throw new WorkPaperNamedExpressionDoesNotExistError(name)
  }

  protected evaluateNamedExpression(expression: InternalNamedExpression): CellValue | CellValue[][] {
    return evaluateWorkPaperNamedExpression(expression, (formula, scope) => this.calculateFormula(formula, scope))
  }

  protected cellSnapshotToRawContent(cell: CellSnapshot, ownerSheetId: number): RawCellContent {
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
