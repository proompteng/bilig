import type { EngineCellMutationRef, SheetRecord, SpreadsheetEngine } from '@bilig/core/headless-runtime'
import { MAX_COLS, MAX_ROWS, type CellSnapshot, type CellValue, type WorkbookSnapshot } from '@bilig/protocol'
import {
  WorkPaperInvalidArgumentsError,
  WorkPaperNamedExpressionDoesNotExistError,
  WorkPaperOperationError,
  WorkPaperSheetError,
} from './work-paper-errors.js'
import { cloneConfig, DEFAULT_CONFIG, validateWorkPaperConfig, WORKPAPER_PUBLIC_ERROR_NAMES } from './work-paper-config.js'
import { assertRowAndColumn, makeNamedExpressionKey } from './work-paper-runtime-helpers.js'
import { inspectSheetDimensionsWithinLimits, workPaperSheetHasDynamicSpillFormula } from './work-paper-sheet-inspection.js'
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
import type { WorkPaperClipboardPayload } from './work-paper-clipboard.js'
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
import { WorkPaperRuntimeLifecycleBase } from './work-paper-runtime-lifecycle-base.js'
import { tryChangeSimpleNumericNamedExpressionFastPath } from './work-paper-named-expression-fast-path-runtime.js'
import { createWorkPaperEngine, workPaperEvaluationTimeoutErrorFrom } from './work-paper-runtime-construction.js'

type NamedExpressionValueSnapshot = WorkPaperNamedExpressionValueSnapshot

let nextWorkbookId = 1
const importedXlsxSourceBytes = Symbol.for('bilig.importedXlsxSourceBytes')
type ImportedXlsxSourceReader = {
  readonly byteLength: number
  readBytes(): Uint8Array
}
type SnapshotWithImportedXlsxSource = WorkbookSnapshot & {
  readonly [importedXlsxSourceBytes]?: Uint8Array | ImportedXlsxSourceReader
}
type WorkbookSnapshotWorkbook = WorkbookSnapshot['workbook']
type WorkbookSnapshotSheetMetadata = NonNullable<WorkbookSnapshot['sheets'][number]['metadata']>
type MetadataRenameEngine = SpreadsheetEngine & {
  readonly renameSheetMetadataOnlyById?: (sheetId: number, newName: string) => boolean
}

type WorkPaperStructuralInsertEngine = SpreadsheetEngine & {
  insertRows(sheetName: string, start: number, count: number, options?: { readonly emitTracked?: boolean }): void
  insertColumns(sheetName: string, start: number, count: number, options?: { readonly emitTracked?: boolean }): void
}

function isCloneRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function cloneSnapshotMetadataValue(value: unknown): unknown {
  if (!isCloneRecord(value)) {
    return value
  }
  if (Array.isArray(value)) {
    return value.map(cloneSnapshotMetadataValue)
  }
  if (ArrayBuffer.isView(value)) {
    return structuredClone(value)
  }

  const path = value['path']
  if (typeof path === 'string') {
    const storage = value['storage']
    const dataBase64 = value['dataBase64']
    const byteLength = value['byteLength']
    if (storage === 'base64' && typeof dataBase64 === 'string' && typeof byteLength === 'number') {
      return { path, storage, dataBase64, byteLength }
    }

    const xml = value['xml']
    if (typeof xml === 'string' && typeof value['readXml'] === 'function') {
      return { path, xml }
    }
  }

  const cloned: Record<string, unknown> = {}
  for (const key of Object.keys(value)) {
    const child = value[key]
    if (typeof child !== 'function') {
      cloned[key] = cloneSnapshotMetadataValue(child)
    }
  }
  return cloned
}

function isWorkbookSnapshotWorkbook(value: unknown): value is WorkbookSnapshotWorkbook {
  return isCloneRecord(value) && typeof value['name'] === 'string'
}

function cloneWorkbookSnapshotWorkbook(workbook: WorkbookSnapshotWorkbook): WorkbookSnapshotWorkbook {
  const cloned = cloneSnapshotMetadataValue(workbook)
  return isWorkbookSnapshotWorkbook(cloned) ? cloned : { name: workbook.name }
}

function isWorkbookSnapshotSheetMetadata(value: unknown): value is WorkbookSnapshotSheetMetadata {
  return isCloneRecord(value)
}

function cloneWorkbookSnapshotSheetMetadata(metadata: WorkbookSnapshotSheetMetadata): WorkbookSnapshotSheetMetadata {
  const cloned = cloneSnapshotMetadataValue(metadata)
  return isWorkbookSnapshotSheetMetadata(cloned) ? cloned : {}
}

function clonePreservedImportedSnapshot(snapshot: WorkbookSnapshot): WorkbookSnapshot {
  const cloned: WorkbookSnapshot = {
    version: snapshot.version,
    workbook: cloneWorkbookSnapshotWorkbook(snapshot.workbook),
    sheets: snapshot.sheets.map((sheet) => ({
      ...sheet,
      ...(sheet.metadata === undefined ? {} : { metadata: cloneWorkbookSnapshotSheetMetadata(sheet.metadata) }),
      cells: sheet.cells,
    })),
  }
  const sourceBytes = (snapshot as SnapshotWithImportedXlsxSource)[importedXlsxSourceBytes]
  if (sourceBytes !== undefined) {
    Object.defineProperty(cloned, importedXlsxSourceBytes, {
      configurable: true,
      enumerable: false,
      value: sourceBytes,
    })
  }
  return cloned
}

export class WorkPaper extends WorkPaperRuntimeLifecycleBase {
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
  protected engineEventsAttached = false
  protected disposed = false
  private preservedImportedSnapshot: WorkbookSnapshot | undefined
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

  override exportSnapshot(): WorkbookSnapshot {
    this.assertNotDisposed()
    this.engineEvents.materializePendingLazyChanges()
    if (this.preservedImportedSnapshot !== undefined) {
      return clonePreservedImportedSnapshot(this.preservedImportedSnapshot)
    }
    return structuredClone(this.engine.exportSnapshot())
  }

  override setCellContents(address: WorkPaperCellAddress, content: RawCellContent | WorkPaperSheet): WorkPaperChange[] {
    this.preservedImportedSnapshot = undefined
    return super.setCellContents(address, content)
  }

  protected override captureChanges(
    semanticEvent: QueuedEvent | undefined,
    mutate: () => void,
    options: { readonly preservePendingTrackedPositions?: boolean } = {},
  ): WorkPaperChange[] {
    this.preservedImportedSnapshot = undefined
    return super.captureChanges(semanticEvent, mutate, options)
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
    workbook.preservedImportedSnapshot = snapshot
    return workbook
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

  protected override captureFunctionRegistry(): void {
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

  protected override clearFunctionBindings(options: { preserveInternalFunctionLookup?: boolean } = {}): void {
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

  protected rewriteFormulaForStorage(formula: string, ownerSheetId: number): string {
    return rewriteWorkPaperFormulaForStorage({
      formula,
      ownerSheetId,
      namedExpressions: this.namedExpressions,
      functionAliasLookup: this.functionAliasLookup,
      messageOf: (error, fallback) => this.messageOf(error, fallback),
    })
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
