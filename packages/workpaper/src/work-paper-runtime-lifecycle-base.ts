import type { SpreadsheetEngine } from '@bilig/core/headless-runtime'
import { MAX_COLS, MAX_ROWS } from '@bilig/protocol'
import {
  cloneConfig,
  canApplyRuntimeOnlyWorkPaperConfigUpdate,
  canReuseWorkPaperSnapshotRebuild,
  normalizeConfiguredWorkPaperCalculationSettings,
  validateWorkPaperConfig,
  WORKPAPER_CONFIG_KEYS,
} from './work-paper-config.js'
import { WorkPaperOperationError, WorkPaperSheetSizeLimitExceededError } from './work-paper-errors.js'
import { cloneWorkPaperClipboardPayload } from './work-paper-clipboard.js'
import { cloneWorkPaperHistoryRecords } from './work-paper-history.js'
import { WorkPaperSheetDimensionCache } from './work-paper-sheet-dimension-cache.js'
import { validateSheetWithinLimits } from './work-paper-sheet-inspection.js'
import { initializeWorkPaperFromSheets } from './work-paper-sheet-initialization.js'
import {
  createWorkPaperEngine,
  releaseWorkPaperEngine,
  workPaperEvaluationTimeoutErrorFrom,
  type WorkPaperTransactionSnapshot,
} from './work-paper-runtime-construction.js'
import { WorkPaperRuntimeFastPathBase } from './work-paper-runtime-fast-path-base.js'
import type { WorkPaperChange, WorkPaperConfig } from './work-paper-types.js'

export abstract class WorkPaperRuntimeLifecycleBase extends WorkPaperRuntimeFastPathBase {
  protected abstract engineEventsAttached: boolean

  protected abstract captureFunctionRegistry(): void
  protected abstract clearFunctionBindings(options?: { preserveInternalFunctionLookup?: boolean }): void

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
    this.rebuildWithConfigAtomically(merged)
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

  private rebuildWithConfigAtomically(nextConfig: WorkPaperConfig): void {
    const snapshot = this.captureTransactionSnapshot()
    try {
      this.rebuildWithConfig(nextConfig)
    } catch (error) {
      this.restoreTransactionSnapshot(snapshot)
      throw error
    }
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

  private replaceEngineForConfig(config: WorkPaperConfig): void {
    this.engineEvents.detach()
    this.engineEventsAttached = false
    releaseWorkPaperEngine(this.engine)
    this.engine = createWorkPaperEngine(config, { fresh: true })
    this.sheetDimensionCache = new WorkPaperSheetDimensionCache(this.engine)
  }
}
