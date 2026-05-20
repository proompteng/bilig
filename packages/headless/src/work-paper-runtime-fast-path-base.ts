import {
  CellFlags,
  type EngineExistingNumericCellMutationResult,
  type SheetRecord,
  type SpreadsheetEngine,
} from '@bilig/core/headless-runtime'
import { MAX_COLS, MAX_ROWS, ValueTag, type CellSnapshot, type CellValue, type WorkbookDefinedNameValueSnapshot } from '@bilig/protocol'
import type { InternalFunctionBinding } from './work-paper-function-registry.js'
import { createWorkPaperRuntimeAdapters, type WorkPaperRuntimeAdapters } from './work-paper-runtime-adapters.js'
import { WorkPaperRuntimeSurface } from './work-paper-runtime-surface.js'
import {
  trySetExistingLiteralWorkPaperCellContentsDirectFastPath,
  type WorkPaperExistingLiteralDirectFastPathRuntime,
} from './work-paper-existing-literal-direct-fast-path.js'
import {
  trySetExistingLiteralWorkPaperCellContentsWithTrackedFastPath,
  trySetExistingNumericWorkPaperCellContentsWithTrackedFastPath,
  type WorkPaperExistingNumericFastPathRuntime,
} from './work-paper-existing-numeric-fast-path.js'
import { setWorkPaperCellContents, type WorkPaperSetCellContentsRuntime } from './work-paper-cell-content-setter.js'
import {
  editWorkPaperAxisIntervalsWithoutRuntimeAdapters,
  type WorkPaperStructuralAxisFastPathRuntime,
} from './work-paper-structural-axis-fast-path.js'
import { getVisibleWorkPaperCellIndexInSheet, readWorkPaperCellValue } from './work-paper-cell-read.js'
import {
  assertRowAndColumn,
  isDeferredBatchLiteralContent,
  isFormulaContent,
  isWorkPaperSheetMatrix,
  makeNamedExpressionKey,
} from './work-paper-runtime-helpers.js'
import { orderWorkPaperCellChanges } from './change-order.js'
import type { InternalNamedExpression, WorkPaperNamedExpressionValueSnapshot } from './work-paper-named-expression-helpers.js'
import type { WorkPaperEngineEventTracker } from './work-paper-engine-event-tracker.js'
import type { WorkPaperMutationQueues } from './work-paper-mutation-queues.js'
import type { WorkPaperSheetDimensionCache } from './work-paper-sheet-dimension-cache.js'
import type { WorkPaperEmitter } from './work-paper-emitter.js'
import type { WorkPaperClipboardPayload } from './work-paper-clipboard.js'
import { trackedEventFromExistingNumericMutationResult, type QueuedEvent } from './work-paper-tracked-event-helpers.js'
import { WORKPAPER_PUBLIC_ERROR_NAMES } from './work-paper-config.js'
import { WorkPaperOperationError } from './work-paper-errors.js'
import type {
  RawCellContent,
  SerializedWorkPaperNamedExpression,
  WorkPaperAxisInterval,
  WorkPaperCellAddress,
  WorkPaperCellChange,
  WorkPaperChange,
  WorkPaperConfig,
  WorkPaperSheet,
} from './work-paper-types.js'
import type { WorkPaperAxisIntervalEditMode, WorkPaperAxisKind } from './work-paper-axis-helpers.js'

const FAST_DIRECT_EXISTING_NUMERIC_LITERAL_FLAGS =
  CellFlags.HasFormula | CellFlags.JsOnly | CellFlags.InCycle | CellFlags.SpillChild | CellFlags.PivotOutput

type ExistingNumericMutationEngine = SpreadsheetEngine & {
  readonly tryApplyExistingNumericCellMutationAt?: (request: {
    readonly sheetId: number
    readonly row: number
    readonly col: number
    readonly cellIndex: number
    readonly value: number
    readonly emitTracked?: boolean
    readonly trustedExistingNumericLiteral?: boolean
    readonly oldNumericValue?: number
  }) => EngineExistingNumericCellMutationResult | null
}

export abstract class WorkPaperRuntimeFastPathBase extends WorkPaperRuntimeSurface {
  protected abstract override readonly emitter: WorkPaperEmitter
  protected abstract readonly internalFunctionLookup: Map<string, InternalFunctionBinding>
  protected abstract override readonly namedExpressions: Map<string, InternalNamedExpression>
  protected abstract override readonly engineEvents: WorkPaperEngineEventTracker
  protected abstract override readonly mutationQueues: WorkPaperMutationQueues
  protected abstract override sheetDimensionCache: WorkPaperSheetDimensionCache
  protected abstract clipboard: WorkPaperClipboardPayload | null
  protected abstract override config: WorkPaperConfig
  protected abstract override namedExpressionValueCache: WorkPaperNamedExpressionValueSnapshot | null
  protected abstract override sheetRecordsCache: readonly SheetRecord[] | null
  protected abstract override batchDepth: number
  protected abstract override batchUsesTrackedFastPath: boolean
  protected abstract override evaluationSuspended: boolean
  protected abstract override queuedEvents: QueuedEvent[]
  protected abstract override disposed: boolean

  private existingNumericFastPathRuntimeCache: WorkPaperExistingNumericFastPathRuntime | undefined
  private existingLiteralDirectFastPathRuntimeCache: WorkPaperExistingLiteralDirectFastPathRuntime | undefined
  private structuralAxisFastPathRuntimeCache: WorkPaperStructuralAxisFastPathRuntime | undefined
  private setCellContentsRuntimeCache: WorkPaperSetCellContentsRuntime | undefined
  private runtimeAdaptersCache: WorkPaperRuntimeAdapters | undefined

  protected abstract applyAxisIntervalEdit(
    axis: WorkPaperAxisKind,
    mode: WorkPaperAxisIntervalEditMode,
    sheetId: number,
    start: number,
    amount: number,
  ): void
  protected abstract applyAxisIntervalEditForSheet(
    axis: WorkPaperAxisKind,
    mode: WorkPaperAxisIntervalEditMode,
    sheet: SheetRecord,
    start: number,
    amount: number,
    options?: { readonly emitTracked?: boolean },
  ): void
  protected abstract applyAxisMove(axis: WorkPaperAxisKind, sheetId: number, start: number, count: number, target: number): void
  protected abstract applyMatrixContents(address: WorkPaperCellAddress, content: WorkPaperSheet): void
  protected abstract applySerializedMatrix(
    targetLeftCorner: WorkPaperCellAddress,
    serialized: RawCellContent[][],
    sourceAnchor: WorkPaperCellAddress,
  ): void
  protected abstract canEditAxisIntervals(
    axis: WorkPaperAxisKind,
    mode: WorkPaperAxisIntervalEditMode,
    sheetId: number,
    indexes: readonly WorkPaperAxisInterval[],
  ): boolean
  protected abstract cellSnapshotToRawContent(cell: CellSnapshot, ownerSheetId: number): RawCellContent
  protected abstract namedExpressionRecord(name: string, scope?: number): InternalNamedExpression
  protected abstract replaceSheetContentInternal(sheetId: number, content: WorkPaperSheet, options: { duringInitialization: boolean }): void
  protected abstract restorePublicFormula(formula: string, ownerSheetId: number): string
  protected abstract rewriteFormulaForStorage(formula: string, ownerSheetId: number): string
  protected abstract toDefinedNameSnapshot(expression: RawCellContent, scope?: number): WorkbookDefinedNameValueSnapshot
  protected abstract upsertNamedExpressionInternal(
    expression: SerializedWorkPaperNamedExpression,
    options: { duringInitialization: boolean; skipValidation?: boolean },
  ): void
  protected abstract validateNamedExpression(expressionName: string, expression: RawCellContent, scope?: number): void

  protected get runtimeAdapters(): WorkPaperRuntimeAdapters {
    this.runtimeAdaptersCache ??= this.createRuntimeAdapters()
    return this.runtimeAdaptersCache
  }

  protected getExistingLiteralDirectFastPathRuntime(): WorkPaperExistingLiteralDirectFastPathRuntime {
    this.existingLiteralDirectFastPathRuntimeCache ??= {
      clearTrackedEngineEvents: () => {
        this.engineEvents.clearEvents()
      },
      computeTrackedChangesWithoutVisibilityCache: (events, options) => this.computeTrackedChangesWithoutVisibilityCache(events, options),
      getBatchDepth: () => this.batchDepth,
      getConfig: () => this.config,
      getEngine: () => this.engine,
      hasNamedExpressions: () => this.namedExpressions.size !== 0,
      hasPendingBatchOps: () => this.mutationQueues.hasPendingBatchOps(),
      hasPendingLazyTrackedChanges: () => this.engineEvents.hasPendingLazyChanges,
      hasTrackedEngineEvents: () => this.engineEvents.hasTrackedEvents,
      hasValuesUpdatedListeners: () => this.emitter.hasListeners('valuesUpdated'),
      isDisposed: () => this.disposed,
      isEvaluationSuspended: () => this.evaluationSuspended,
      messageOf: (error, fallback) => this.messageOf(error, fallback),
      readSingleTrackedCellChange: (cellIndex) => this.readSingleTrackedCellChange(cellIndex),
      trackedA1: (row, col) => this.trackedA1(row, col),
    }
    return this.existingLiteralDirectFastPathRuntimeCache
  }

  protected getSetCellContentsRuntime(): WorkPaperSetCellContentsRuntime {
    this.setCellContentsRuntimeCache ??= {
      assertNotDisposed: () => this.assertNotDisposed(),
      getConfig: () => this.config,
      getEngine: () => this.engine,
      sheetRecord: (sheetId) => this.sheetRecord(sheetId),
      getVisibleCellIndexInSheet: (sheet, row, col) => this.getVisibleCellIndexInSheet(sheet, row, col),
      isEvaluationSuspended: () => this.evaluationSuspended,
      getBatchDepth: () => this.batchDepth,
      enqueueSuspendedLiteralMutation: (sheetId, row, col, content, cellIndex) =>
        this.enqueueSuspendedLiteralMutation(sheetId, row, col, content, cellIndex),
      enqueueDeferredBatchLiteral: (sheetId, row, col, content, cellIndex) =>
        this.enqueueDeferredBatchLiteral(sheetId, row, col, content, cellIndex),
      trySetExistingNumericCellContentsWithTrackedFastPath: (request) =>
        trySetExistingNumericWorkPaperCellContentsWithTrackedFastPath(this.getExistingNumericFastPathRuntime(), request),
      trySetExistingLiteralCellContentsWithTrackedFastPath: (request) =>
        trySetExistingLiteralWorkPaperCellContentsWithTrackedFastPath(this.getExistingNumericFastPathRuntime(), request),
      flushPendingBatchOps: () => this.flushPendingBatchOps(),
      rewriteFormulaForStorage: (formula, ownerSheetId) => this.rewriteFormulaForStorage(formula, ownerSheetId),
      applyCellMutationRefs: (refs, options) => this.applyCellMutationRefs(refs, options),
      canUseTrackedMutationFastPath: () => this.canUseTrackedMutationFastPath(),
      isTrackedBatchFastPathActive: () => this.batchUsesTrackedFastPath,
      captureTrackedChangesWithoutVisibilityCache: (mutate, options) => this.captureTrackedChangesWithoutVisibilityCache(mutate, options),
      captureChanges: (mutate) => this.captureChanges(undefined, mutate),
      isItPossibleToSetCellContents: (address, content) => this.isItPossibleToSetCellContents(address, content),
      applyMatrixContents: (address, content) => this.applyMatrixContents(address, content),
    }
    return this.setCellContentsRuntimeCache
  }

  protected getVisibleCellIndex(sheetId: number, row: number, col: number): number | undefined {
    const sheet = this.engine.workbook.getSheetById(sheetId)
    if (!sheet) {
      return undefined
    }
    return this.getVisibleCellIndexInSheet(sheet, row, col)
  }

  protected getVisibleCellIndexInSheet(sheet: SheetRecord, row: number, col: number): number | undefined {
    return getVisibleWorkPaperCellIndexInSheet(sheet, row, col)
  }

  private trySetExistingNumericCellContentsInlineDirectFastPath(
    address: WorkPaperCellAddress,
    content: RawCellContent | WorkPaperSheet,
  ): WorkPaperChange[] | null {
    if (
      typeof content !== 'number' ||
      this.disposed ||
      !Number.isInteger(address.row) ||
      !Number.isInteger(address.col) ||
      address.row < 0 ||
      address.col < 0 ||
      address.row >= (this.config.maxRows ?? MAX_ROWS) ||
      address.col >= (this.config.maxColumns ?? MAX_COLS) ||
      this.evaluationSuspended ||
      this.batchDepth !== 0 ||
      this.namedExpressions.size !== 0 ||
      this.emitter.hasListeners('valuesUpdated') ||
      this.engineEvents.hasPendingLazyChanges ||
      this.engineEvents.hasTrackedEvents ||
      this.mutationQueues.hasPendingBatchOps()
    ) {
      return null
    }

    const sheet = this.engine.workbook.getSheetById(address.sheet)
    if (!sheet || sheet.structureVersion !== 1) {
      return null
    }
    const cellIndex = sheet.grid.getPhysical(address.row, address.col)
    if (cellIndex === -1) {
      return null
    }

    const cellStore = this.engine.workbook.cellStore
    if (
      cellStore.sheetIds[cellIndex] !== address.sheet ||
      cellStore.rows[cellIndex] !== address.row ||
      cellStore.cols[cellIndex] !== address.col ||
      cellStore.tags[cellIndex] !== ValueTag.Number ||
      (cellStore.formulaIds[cellIndex] ?? 0) !== 0 ||
      ((cellStore.flags[cellIndex] ?? 0) & FAST_DIRECT_EXISTING_NUMERIC_LITERAL_FLAGS) !== 0
    ) {
      return null
    }

    const mutationEngine = this.engine as ExistingNumericMutationEngine
    if (typeof mutationEngine.tryApplyExistingNumericCellMutationAt !== 'function') {
      return null
    }

    let result: EngineExistingNumericCellMutationResult | null = null
    try {
      result = mutationEngine.tryApplyExistingNumericCellMutationAt({
        sheetId: address.sheet,
        row: address.row,
        col: address.col,
        cellIndex,
        value: content,
        emitTracked: false,
        trustedExistingNumericLiteral: true,
        oldNumericValue: cellStore.numbers[cellIndex] ?? 0,
      })
    } catch (error) {
      if (error instanceof Error && WORKPAPER_PUBLIC_ERROR_NAMES.has(error.name)) {
        throw error
      }
      throw new WorkPaperOperationError(this.messageOf(error, 'Mutation failed'))
    }
    if (!result) {
      return null
    }
    if (this.engineEvents.hasTrackedEvents) {
      this.engineEvents.clearEvents()
    }

    return (
      this.tryReadCompactExistingNumericDirectChanges(result, {
        address,
        cellIndex,
        content,
        sheet,
      }) ??
      this.computeTrackedChangesWithoutVisibilityCache([trackedEventFromExistingNumericMutationResult(result)], {
        preferLazyPublicChanges: true,
      })
    )
  }

  private tryReadCompactExistingNumericDirectChanges(
    result: EngineExistingNumericCellMutationResult,
    request: {
      readonly address: WorkPaperCellAddress
      readonly cellIndex: number
      readonly content: number
      readonly sheet: SheetRecord
    },
  ): WorkPaperChange[] | null {
    const changedCellIndices = result.changedCellIndices
    if (changedCellIndices !== undefined) {
      if (changedCellIndices.length > 4) {
        return null
      }
      const changes: WorkPaperChange[] = []
      for (let index = 0; index < changedCellIndices.length; index += 1) {
        const change = this.readSingleTrackedCellChange(changedCellIndices[index]!)
        if (change === undefined) {
          return null
        }
        changes.push(change)
      }
      return changes
    }

    const changedCellCount = result.changedCellCount ?? 0
    if (
      changedCellCount === 0 ||
      changedCellCount > 2 ||
      result.firstChangedCellIndex === undefined ||
      result.firstChangedCellIndex !== request.cellIndex
    ) {
      return null
    }

    const firstChange: WorkPaperCellChange = {
      kind: 'cell',
      address: { sheet: request.address.sheet, row: request.address.row, col: request.address.col },
      sheetName: request.sheet.name,
      a1: this.trackedA1(request.address.row, request.address.col),
      newValue: { tag: ValueTag.Number, value: request.content },
    }
    if (changedCellCount === 1) {
      return [firstChange]
    }
    if (result.secondChangedCellIndex === undefined) {
      return null
    }

    const cellStore = this.engine.workbook.cellStore
    const secondRow = result.secondChangedRow ?? cellStore.rows[result.secondChangedCellIndex]
    const secondCol = result.secondChangedCol ?? cellStore.cols[result.secondChangedCellIndex]
    if (secondRow === undefined || secondCol === undefined) {
      return null
    }
    const secondValue: CellValue =
      result.secondChangedValue ??
      (result.secondChangedNumericValue === undefined
        ? cellStore.getValue(result.secondChangedCellIndex, (stringId) => (stringId === 0 ? '' : this.engine.strings.get(stringId)))
        : { tag: ValueTag.Number, value: result.secondChangedNumericValue })
    return [
      firstChange,
      {
        kind: 'cell',
        address: { sheet: request.address.sheet, row: secondRow, col: secondCol },
        sheetName: request.sheet.name,
        a1: this.trackedA1(secondRow, secondCol),
        newValue: secondValue,
      },
    ]
  }

  private tryEnqueueDeferredBatchLiteralCellContents(
    address: WorkPaperCellAddress,
    content: RawCellContent | WorkPaperSheet,
  ): WorkPaperChange[] | null {
    if (this.batchDepth === 0 || this.evaluationSuspended || isWorkPaperSheetMatrix(content)) {
      return null
    }
    if (!isDeferredBatchLiteralContent(content) || isFormulaContent(content)) {
      return null
    }
    this.assertNotDisposed()
    const sheet = this.sheetRecord(address.sheet)
    assertRowAndColumn(address.row, 'address.row')
    assertRowAndColumn(address.col, 'address.col')
    if (address.row >= (this.config.maxRows ?? MAX_ROWS) || address.col >= (this.config.maxColumns ?? MAX_COLS)) {
      throw new WorkPaperOperationError('Cell contents cannot be set')
    }
    const visibleCellIndex = this.getVisibleCellIndexInSheet(sheet, address.row, address.col)
    return this.enqueueValidatedDeferredBatchLiteral(address.sheet, address.row, address.col, content, visibleCellIndex) ? [] : null
  }

  override getCellValue(address: WorkPaperCellAddress): CellValue {
    this.assertReadable()
    return readWorkPaperCellValue({ engine: this.engine, sheet: this.sheetRecord(address.sheet), address })
  }

  override setCellContents(address: WorkPaperCellAddress, content: RawCellContent | WorkPaperSheet): WorkPaperChange[] {
    const deferredBatchChanges = this.tryEnqueueDeferredBatchLiteralCellContents(address, content)
    if (deferredBatchChanges !== null) {
      return deferredBatchChanges
    }
    const inlineDirectChanges = this.trySetExistingNumericCellContentsInlineDirectFastPath(address, content)
    if (inlineDirectChanges !== null) {
      return inlineDirectChanges
    }
    const directChanges = trySetExistingLiteralWorkPaperCellContentsDirectFastPath(
      this.getExistingLiteralDirectFastPathRuntime(),
      address,
      content,
    )
    if (directChanges !== null) {
      return directChanges
    }
    return setWorkPaperCellContents(this.getSetCellContentsRuntime(), address, content)
  }

  protected editAxisIntervalsWithoutRuntimeAdapters(
    axis: WorkPaperAxisKind,
    mode: WorkPaperAxisIntervalEditMode,
    sheetId: number,
    startOrInterval: number | WorkPaperAxisInterval,
    countOrInterval: number | WorkPaperAxisInterval | undefined,
    restIntervals: readonly WorkPaperAxisInterval[],
  ): WorkPaperChange[] {
    return editWorkPaperAxisIntervalsWithoutRuntimeAdapters(
      this.getStructuralAxisFastPathRuntime(),
      axis,
      mode,
      sheetId,
      startOrInterval,
      countOrInterval,
      restIntervals,
    )
  }

  private createRuntimeAdapters(): WorkPaperRuntimeAdapters {
    return createWorkPaperRuntimeAdapters({
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
      applySerializedMatrix: (targetLeftCorner, content, sourceAnchor) =>
        this.applySerializedMatrix(targetLeftCorner, content, sourceAnchor),
      doesSheetIdExist: (sheetId) => this.engine.workbook.getSheetById(sheetId) !== undefined,
      hasNamedExpression: (expressionName, scope) => this.namedExpressions.has(makeNamedExpressionKey(expressionName, scope)),
      canEditAxisIntervals: (axis, mode, sheetId, indexes) => this.canEditAxisIntervals(axis, mode, sheetId, indexes),
      batchStructuralChanges: (operations) => this.batchStructuralChanges(operations),
      captureAxisChange: (operations) => this.captureChanges(undefined, operations),
      applyAxisIntervalEdit: (axis, mode, sheetId, start, amount) => this.applyAxisIntervalEdit(axis, mode, sheetId, start, amount),
      applyAxisMove: (axis, sheetId, start, count, target) => this.applyAxisMove(axis, sheetId, start, count, target),
      messageOf: (error, fallback) => this.messageOf(error, fallback),
    })
  }

  private getExistingNumericFastPathRuntime(): WorkPaperExistingNumericFastPathRuntime {
    this.existingNumericFastPathRuntimeCache ??= {
      canUseTrackedMutationFastPath: () => this.canUseTrackedMutationFastPath(),
      getEngine: () => this.engine,
      hasPendingLazyTrackedChanges: () => this.engineEvents.hasPendingLazyChanges,
      materializePendingLazyTrackedChanges: () => this.engineEvents.materializePendingLazyChanges(),
      hasTrackedEngineEvents: () => this.engineEvents.hasTrackedEvents,
      drainTrackedEngineEvents: () => this.engineEvents.drain(),
      clearTrackedEngineEvents: () => {
        this.engineEvents.clearEvents()
      },
      getEngineEventCaptureEnabled: () => this.engineEvents.isCaptureEnabled,
      setEngineEventCaptureEnabled: (enabled) => {
        this.engineEvents.setCaptureEnabled(enabled)
      },
      hasPendingBatchOps: () => this.mutationQueues.hasPendingBatchOps(),
      flushPendingBatchOps: () => this.flushPendingBatchOps(),
      messageOf: (error, fallback) => this.messageOf(error, fallback),
      trackedA1: (row, col) => this.trackedA1(row, col),
      orderChanges: (changes, explicitChangedCount) => orderWorkPaperCellChanges(changes, this.listSheetRecords(), explicitChangedCount),
      computeTrackedChangesWithoutVisibilityCache: (events, options) => this.computeTrackedChangesWithoutVisibilityCache(events, options),
      trackLazyChanges: (changes) => {
        this.engineEvents.trackLazyChanges(changes)
      },
      hasValuesUpdatedListeners: () => this.emitter.hasListeners('valuesUpdated'),
      emitValuesUpdated: (changes) => {
        this.emitter.emitDetailed({ eventName: 'valuesUpdated', payload: { changes } })
      },
    }
    return this.existingNumericFastPathRuntimeCache
  }

  private getStructuralAxisFastPathRuntime(): WorkPaperStructuralAxisFastPathRuntime {
    this.structuralAxisFastPathRuntimeCache ??= {
      applyAxisIntervalEditForSheet: (axis, mode, sheet, start, amount, options) =>
        this.applyAxisIntervalEditForSheet(axis, mode, sheet, start, amount, options),
      assertNotDisposed: () => this.assertNotDisposed(),
      batchStructuralChanges: (operations) => this.batchStructuralChanges(operations),
      canUseTrackedStructuralFastPath: () => this.canUseTrackedStructuralFastPath(),
      captureTrackedChangesWithoutVisibilityCache: (mutate) => this.captureTrackedChangesWithoutVisibilityCache(mutate),
      drainTrackedEngineEvents: () => {
        this.engineEvents.drain()
      },
      getBatchUsesTrackedFastPath: () => this.batchUsesTrackedFastPath,
      getConfig: () => this.config,
      hasPendingLazyTrackedChanges: () => this.engineEvents.hasPendingLazyChanges,
      hasTrackedEngineEvents: () => this.engineEvents.hasTrackedEvents,
      materializePendingLazyTrackedChanges: () => this.engineEvents.materializePendingLazyChanges(),
      messageOf: (error, fallback) => this.messageOf(error, fallback),
      sheetRecord: (sheetId) => this.sheetRecord(sheetId),
    }
    return this.structuralAxisFastPathRuntimeCache
  }
}
