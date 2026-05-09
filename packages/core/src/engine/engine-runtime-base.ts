import type { CellSnapshot, RecalcMetrics, SelectionState, SyncState, WorkbookSnapshot } from '@bilig/protocol'
import { Float64Arena, Uint32Arena } from '@bilig/formula'
import type { EngineOp, EngineOpBatch } from '@bilig/workbook-domain'
import { cellToCsvValue, serializeCsv } from '../csv.js'
import { CycleDetector } from '../cycle-detection.js'
import { EdgeArena, type EdgeSlice } from '../edge-arena.js'
import { EngineEventBus } from '../events.js'
import { FormulaTable } from '../formula-table.js'
import { createEngineCounters, type EngineCounters } from '../perf/engine-counters.js'
import { RangeRegistry } from '../range-registry.js'
import { batchOpOrder, compareOpOrder, createBatch, createReplicaState, type OpOrder, type ReplicaState } from '../replica-state.js'
import { RecalcScheduler } from '../scheduler.js'
import { StringPool } from '../string-pool.js'
import { WasmKernelFacade } from '../wasm-facade.js'
import { WorkbookStore } from '../workbook-store.js'
import {
  createEngineRuntimeState,
  createInitialRecalcMetrics,
  createInitialSelectionState,
  type EngineRuntimeState,
  type EngineSyncClientConnection,
  type RuntimeFormula,
  type SpreadsheetEngineOptions,
  type TransactionLogEntry,
  type U32,
} from './runtime-state.js'
import { createEngineServiceRuntime, runEngineEffect, type EngineServiceRuntime } from './live.js'
import { EngineEvaluationTimeoutError } from './errors.js'

export abstract class SpreadsheetEngineRuntimeBase {
  protected readonly performanceCounters: EngineCounters = createEngineCounters()
  readonly workbook: WorkbookStore
  readonly strings = new StringPool()
  readonly events = new EngineEventBus()
  protected readonly replicaState: ReplicaState
  readonly ranges = new RangeRegistry(this.performanceCounters)
  readonly scheduler = new RecalcScheduler(this.performanceCounters)
  readonly wasm = new WasmKernelFacade()

  protected readonly formulas: FormulaTable<RuntimeFormula>
  private readonly cycleDetector = new CycleDetector()
  private readonly edgeArena = new EdgeArena()
  private readonly programArena = new Uint32Arena()
  private readonly constantArena = new Float64Arena()
  private readonly rangeListArena = new Uint32Arena()
  protected reverseCellEdges: Array<EdgeSlice | undefined> = []
  private reverseRangeEdges: Array<EdgeSlice | undefined> = []
  private readonly reverseDefinedNameEdges = new Map<string, Set<number>>()
  private readonly reverseTableEdges = new Map<string, Set<number>>()
  protected readonly reverseSpillEdges = new Map<string, Set<number>>()
  protected readonly reverseAggregateColumnEdges = new Map<number, Set<number>>()
  protected readonly reverseExactLookupColumnEdges = new Map<number, EdgeSlice>()
  protected readonly reverseSortedLookupColumnEdges = new Map<number, EdgeSlice>()
  protected readonly pivotOutputOwners = new Map<number, string>()
  protected readonly batchListeners = new Set<(batch: EngineOpBatch) => void>()
  private readonly selectionListeners = new Set<() => void>()
  protected readonly entityVersions = new Map<string, OpOrder>()
  protected readonly sheetDeleteVersions = new Map<string, OpOrder>()
  protected selection: SelectionState = createInitialSelectionState()
  protected syncState: SyncState = 'local-only'
  protected syncClientConnection: EngineSyncClientConnection | null = null
  protected readonly undoStack: TransactionLogEntry[] = []
  protected readonly redoStack: TransactionLogEntry[] = []
  protected transactionReplayDepth = 0
  private useColumnIndexEnabled: boolean
  private evaluationTimeoutMs: number | undefined
  private evaluationBudgetActive = false
  private evaluationBudgetDeadlineMs = Number.POSITIVE_INFINITY
  private evaluationBudgetDepth = 0
  private evaluationBudgetStepCount = 0
  private dependencyBuildEpoch = 1
  private dependencyBuildSeen: U32 = new Uint32Array(128)
  private dependencyBuildCells: U32 = new Uint32Array(128)
  private dependencyBuildEntities: U32 = new Uint32Array(128)
  private dependencyBuildRanges: U32 = new Uint32Array(128)
  private dependencyBuildNewRanges: U32 = new Uint32Array(128)
  private symbolicRefBindings: U32 = new Uint32Array(128)
  private symbolicRangeBindings: U32 = new Uint32Array(128)
  private wasmProgramTargets: U32 = new Uint32Array(128)
  private wasmProgramOffsets: U32 = new Uint32Array(128)
  private wasmProgramLengths: U32 = new Uint32Array(128)
  private wasmConstantOffsets: U32 = new Uint32Array(128)
  private wasmConstantLengths: U32 = new Uint32Array(128)
  private wasmRangeOffsets: U32 = new Uint32Array(128)
  private wasmRangeLengths: U32 = new Uint32Array(128)
  private wasmRangeRowCounts: U32 = new Uint32Array(128)
  private wasmRangeColCounts: U32 = new Uint32Array(128)
  private topoIndegree: U32 = new Uint32Array(128)
  private topoQueue: U32 = new Uint32Array(128)
  protected batchMutationDepth = 0
  private wasmProgramSyncPending = false
  protected lastMetrics: RecalcMetrics = createInitialRecalcMetrics()
  protected readonly state: EngineRuntimeState
  protected readonly runtime: EngineServiceRuntime

  constructor(options: SpreadsheetEngineOptions = {}) {
    this.workbook = new WorkbookStore(options.workbookName ?? 'Workbook', this.performanceCounters)
    this.formulas = new FormulaTable(this.workbook.cellStore)
    this.replicaState = createReplicaState(options.replicaId ?? 'local')
    this.useColumnIndexEnabled = options.useColumnIndex ?? false
    this.evaluationTimeoutMs = options.evaluationTimeoutMs
    this.state = createEngineRuntimeState({
      workbook: this.workbook,
      strings: this.strings,
      events: this.events,
      ranges: this.ranges,
      scheduler: this.scheduler,
      wasm: this.wasm,
      formulas: this.formulas,
      replicaState: this.replicaState,
      entityVersions: this.entityVersions,
      sheetDeleteVersions: this.sheetDeleteVersions,
      batchListeners: this.batchListeners,
      selectionListeners: this.selectionListeners,
      undoStack: this.undoStack,
      redoStack: this.redoStack,
      counters: this.performanceCounters,
      trackReplicaVersions: options.trackReplicaVersions ?? true,
      getUseColumnIndex: () => this.useColumnIndexEnabled,
      setUseColumnIndex: (enabled) => {
        this.useColumnIndexEnabled = enabled
      },
      setEvaluationTimeoutMs: (timeoutMs) => {
        this.evaluationTimeoutMs = timeoutMs
      },
      beginEvaluationBudget: (startedAtMs) => {
        this.beginEvaluationBudget(startedAtMs)
      },
      endEvaluationBudget: () => {
        this.endEvaluationBudget()
      },
      checkEvaluationBudget: (stepCost) => {
        this.checkEvaluationBudget(stepCost)
      },
      getSelection: () => this.selection,
      setSelection: (selection) => {
        this.selection = selection
      },
      getSyncState: () => this.syncState,
      setSyncState: (state) => {
        this.syncState = state
      },
      getSyncClientConnection: () => this.syncClientConnection,
      setSyncClientConnection: (connection) => {
        this.syncClientConnection = connection
      },
      getTransactionReplayDepth: () => this.transactionReplayDepth,
      setTransactionReplayDepth: (depth) => {
        this.transactionReplayDepth = depth
      },
      getLastMetrics: () => this.lastMetrics,
      setLastMetrics: (metrics) => {
        this.lastMetrics = metrics
      },
    })
    this.runtime = createEngineServiceRuntime({
      state: this.state,
      getCellByIndex: (cellIndex) => this.getCellByIndex(cellIndex),
      exportSnapshot: () => this.exportSnapshot(),
      importSnapshot: (snapshot) => this.importSnapshot(snapshot),
      maintenance: {
        state: this.state,
        edgeArena: this.edgeArena,
        reverseState: {
          reverseCellEdges: this.reverseCellEdges,
          reverseRangeEdges: this.reverseRangeEdges,
          reverseDefinedNameEdges: this.reverseDefinedNameEdges,
          reverseTableEdges: this.reverseTableEdges,
          reverseSpillEdges: this.reverseSpillEdges,
          reverseAggregateColumnEdges: this.reverseAggregateColumnEdges,
          reverseExactLookupColumnEdges: this.reverseExactLookupColumnEdges,
          reverseSortedLookupColumnEdges: this.reverseSortedLookupColumnEdges,
        },
        pivotOutputOwners: this.pivotOutputOwners,
        setWasmProgramSyncPending: (next) => {
          this.wasmProgramSyncPending = next
        },
        resetWasmState: () => {
          this.wasm.resetStoreState()
        },
      },
      mutationSupport: {
        state: this.state,
        edgeArena: this.edgeArena,
        reverseState: {
          reverseCellEdges: this.reverseCellEdges,
          reverseRangeEdges: this.reverseRangeEdges,
        },
        getSelectionState: () => this.getSelectionState(),
        setSelection: (sheetName, address) => this.setSelection(sheetName, address),
      },
      formulaBinding: {
        state: this.state,
        edgeArena: this.edgeArena,
        programArena: this.programArena,
        constantArena: this.constantArena,
        rangeListArena: this.rangeListArena,
        reverseState: {
          reverseCellEdges: this.reverseCellEdges,
          reverseRangeEdges: this.reverseRangeEdges,
          reverseDefinedNameEdges: this.reverseDefinedNameEdges,
          reverseTableEdges: this.reverseTableEdges,
          reverseSpillEdges: this.reverseSpillEdges,
          reverseAggregateColumnEdges: this.reverseAggregateColumnEdges,
          reverseExactLookupColumnEdges: this.reverseExactLookupColumnEdges,
          reverseSortedLookupColumnEdges: this.reverseSortedLookupColumnEdges,
        },
        getDependencyBuildEpoch: () => this.dependencyBuildEpoch,
        setDependencyBuildEpoch: (next) => {
          this.dependencyBuildEpoch = next
        },
        getDependencyBuildSeen: () => this.dependencyBuildSeen,
        setDependencyBuildSeen: (next) => {
          this.dependencyBuildSeen = next
        },
        getDependencyBuildCells: () => this.dependencyBuildCells,
        setDependencyBuildCells: (next) => {
          this.dependencyBuildCells = next
        },
        getDependencyBuildEntities: () => this.dependencyBuildEntities,
        setDependencyBuildEntities: (next) => {
          this.dependencyBuildEntities = next
        },
        getDependencyBuildRanges: () => this.dependencyBuildRanges,
        setDependencyBuildRanges: (next) => {
          this.dependencyBuildRanges = next
        },
        getDependencyBuildNewRanges: () => this.dependencyBuildNewRanges,
        setDependencyBuildNewRanges: (next) => {
          this.dependencyBuildNewRanges = next
        },
        getSymbolicRefBindings: () => this.symbolicRefBindings,
        setSymbolicRefBindings: (next) => {
          this.symbolicRefBindings = next
        },
        getSymbolicRangeBindings: () => this.symbolicRangeBindings,
        setSymbolicRangeBindings: (next) => {
          this.symbolicRangeBindings = next
        },
      },
      formulaGraph: {
        state: this.state,
        cycleDetector: this.cycleDetector,
        programArena: this.programArena,
        constantArena: this.constantArena,
        rangeListArena: this.rangeListArena,
        getTopoIndegree: () => this.topoIndegree,
        setTopoIndegree: (next) => {
          this.topoIndegree = next
        },
        getTopoQueue: () => this.topoQueue,
        setTopoQueue: (next) => {
          this.topoQueue = next
        },
        getWasmProgramTargets: () => this.wasmProgramTargets,
        setWasmProgramTargets: (next) => {
          this.wasmProgramTargets = next
        },
        getWasmProgramOffsets: () => this.wasmProgramOffsets,
        setWasmProgramOffsets: (next) => {
          this.wasmProgramOffsets = next
        },
        getWasmProgramLengths: () => this.wasmProgramLengths,
        setWasmProgramLengths: (next) => {
          this.wasmProgramLengths = next
        },
        getWasmConstantOffsets: () => this.wasmConstantOffsets,
        setWasmConstantOffsets: (next) => {
          this.wasmConstantOffsets = next
        },
        getWasmConstantLengths: () => this.wasmConstantLengths,
        setWasmConstantLengths: (next) => {
          this.wasmConstantLengths = next
        },
        getWasmRangeOffsets: () => this.wasmRangeOffsets,
        setWasmRangeOffsets: (next) => {
          this.wasmRangeOffsets = next
        },
        getWasmRangeLengths: () => this.wasmRangeLengths,
        setWasmRangeLengths: (next) => {
          this.wasmRangeLengths = next
        },
        getWasmRangeRowCounts: () => this.wasmRangeRowCounts,
        setWasmRangeRowCounts: (next) => {
          this.wasmRangeRowCounts = next
        },
        getWasmRangeColCounts: () => this.wasmRangeColCounts,
        setWasmRangeColCounts: (next) => {
          this.wasmRangeColCounts = next
        },
        getBatchMutationDepth: () => this.batchMutationDepth,
        getWasmProgramSyncPending: () => this.wasmProgramSyncPending,
        setWasmProgramSyncPending: (next) => {
          this.wasmProgramSyncPending = next
        },
      },
      traversal: {
        state: this.state,
        edgeArena: this.edgeArena,
        reverseState: {
          reverseCellEdges: this.reverseCellEdges,
          reverseRangeEdges: this.reverseRangeEdges,
          reverseExactLookupColumnEdges: this.reverseExactLookupColumnEdges,
          reverseSortedLookupColumnEdges: this.reverseSortedLookupColumnEdges,
        },
      },
      cellToCsvValue: (cell) => cellToCsvValue(cell),
      serializeCsv: (rows) => serializeCsv(rows),
      pivotState: {
        pivotOutputOwners: this.pivotOutputOwners,
      },
      recalc: {
        state: this.state,
        getCellByIndex: (cellIndex) => this.getCellByIndex(cellIndex),
        exportSnapshot: () => this.exportSnapshot(),
        importSnapshot: (snapshot) => this.importSnapshot(snapshot),
        now: () => new Date(),
        random: () => Math.random(),
        performanceNow: () => performance.now(),
        forEachFormulaDependencyCell: (cellIndex, fn) => this.runtime.traversal.forEachFormulaDependencyCellNow(cellIndex, fn),
      },
      pivot: {
        state: {
          workbook: this.state.workbook,
          strings: this.state.strings,
          formulas: this.state.formulas,
          ranges: this.state.ranges,
          wasm: this.state.wasm,
          pivotOutputOwners: this.pivotOutputOwners,
        },
      },
      applyRemoteSnapshot: (snapshot) => {
        this.importSnapshot(snapshot)
      },
      operation: {
        state: this.state,
        reverseState: {
          reverseCellEdges: this.reverseCellEdges,
          reverseSpillEdges: this.reverseSpillEdges,
          reverseAggregateColumnEdges: this.reverseAggregateColumnEdges,
          reverseExactLookupColumnEdges: this.reverseExactLookupColumnEdges,
          reverseSortedLookupColumnEdges: this.reverseSortedLookupColumnEdges,
        },
        getBatchMutationDepth: () => this.batchMutationDepth,
        setBatchMutationDepth: (next) => {
          this.batchMutationDepth = next
        },
        materializePivot: (pivotRecord) => this.runtime.pivot.materializePivotNow(pivotRecord),
        refreshRangeDependencies: () => {
          return
        },
        materializeDeferredStructuralFormulaSources: () => {
          return
        },
        getChangedFormulaBuffer: () => new Uint32Array(),
        repairTopoRanks: () => false,
        getEntityDependents: () => new Uint32Array(),
        getSingleEntityDependent: () => -1,
        collectFormulaDependents: () => new Uint32Array(),
        noteExactLookupLiteralWrite: () => {
          return
        },
        noteAggregateLiteralWrite: () => {
          return
        },
        noteSortedLookupLiteralWrite: () => {
          return
        },
        sortedLookup: {
          findPreparedVectorMatch: () => ({ handled: false }),
        },
        exactLookup: {
          findPreparedVectorMatch: () => ({ handled: false }),
        },
        deferKernelSync: () => {
          return
        },
        invalidateExactLookupColumn: () => {
          return
        },
        invalidateAggregateColumn: () => {
          return
        },
        invalidateSortedLookupColumn: () => {
          return
        },
      },
    })
    if (!this.wasm.initSyncIfPossible()) {
      void this.wasm.init()
    }
  }

  renameSheetMetadataOnly(oldName: string, newName: string): boolean {
    const trimmedName = newName.trim()
    if (
      trimmedName.length === 0 ||
      oldName === trimmedName ||
      !this.workbook.getSheet(oldName) ||
      this.workbook.getSheet(trimmedName) ||
      this.syncClientConnection !== null ||
      this.batchListeners.size > 0 ||
      this.batchMutationDepth !== 0 ||
      this.transactionReplayDepth !== 0
    ) {
      return false
    }

    const renamedSheet = this.workbook.renameSheet(oldName, trimmedName)
    if (!renamedSheet) {
      return false
    }
    const op: EngineOp = { kind: 'renameSheet', oldName, newName: trimmedName }
    if (this.state.trackReplicaVersions) {
      const batch = createBatch(this.replicaState, [op])
      const order = batchOpOrder(batch, 0)
      this.entityVersions.set(`sheet:${oldName}`, order)
      this.entityVersions.set(`sheet:${trimmedName}`, order)
      this.sheetDeleteVersions.set(oldName, order)
      const renamedTombstone = this.sheetDeleteVersions.get(trimmedName)
      if (!renamedTombstone || compareOpOrder(order, renamedTombstone) > 0) {
        this.sheetDeleteVersions.delete(trimmedName)
      }
    }
    if (this.selection.sheetName === oldName) {
      this.selection = { ...this.selection, sheetName: trimmedName }
    }
    if (this.workbook.metadata.definedNames.size > 0) {
      runEngineEffect(this.runtime.maintenance.rewriteDefinedNamesForSheetRename(oldName, trimmedName))
    }
    this.runtime.binding.deferCellFormulasForSheetRenameNow(oldName, trimmedName)
    this.undoStack.push({
      forward: { kind: 'single-op', op },
      inverse: { kind: 'single-op', op: { kind: 'renameSheet', oldName: trimmedName, newName: oldName } },
    })
    this.redoStack.length = 0
    return true
  }

  setEvaluationTimeoutMs(timeoutMs: number | undefined): void {
    this.state.setEvaluationTimeoutMs(timeoutMs)
  }

  private beginEvaluationBudget(startedAtMs: number): void {
    const timeoutMs = this.evaluationTimeoutMs
    if (timeoutMs === undefined) {
      this.evaluationBudgetActive = false
      this.evaluationBudgetDepth = 0
      this.evaluationBudgetDeadlineMs = Number.POSITIVE_INFINITY
      this.evaluationBudgetStepCount = 0
      return
    }
    if (this.evaluationBudgetActive) {
      this.evaluationBudgetDepth += 1
      return
    }
    this.evaluationBudgetActive = true
    this.evaluationBudgetDepth = 1
    this.evaluationBudgetDeadlineMs = startedAtMs + timeoutMs
    this.evaluationBudgetStepCount = 0
  }

  private endEvaluationBudget(): void {
    if (!this.evaluationBudgetActive) {
      return
    }
    this.evaluationBudgetDepth = Math.max(0, this.evaluationBudgetDepth - 1)
    if (this.evaluationBudgetDepth > 0) {
      return
    }
    this.evaluationBudgetActive = false
    this.evaluationBudgetDepth = 0
    this.evaluationBudgetDeadlineMs = Number.POSITIVE_INFINITY
    this.evaluationBudgetStepCount = 0
  }

  private checkEvaluationBudget(stepCost = 1): void {
    if (!this.evaluationBudgetActive || this.evaluationTimeoutMs === undefined) {
      return
    }
    this.evaluationBudgetStepCount += Math.max(1, stepCost)
    if (this.evaluationTimeoutMs === 0 || performance.now() >= this.evaluationBudgetDeadlineMs) {
      throw new EngineEvaluationTimeoutError(this.evaluationTimeoutMs)
    }
  }

  abstract getSelectionState(): SelectionState
  abstract setSelection(
    sheetName: string,
    address: string | null,
    options?: {
      anchorAddress?: string | null
      range?: { startAddress: string; endAddress: string } | null
      editMode?: SelectionState['editMode']
    },
  ): void
  abstract getCellByIndex(cellIndex: number): CellSnapshot
  abstract exportSnapshot(): WorkbookSnapshot
  abstract importSnapshot(snapshot: WorkbookSnapshot): void

  protected executeLocalTransaction(ops: EngineOp[], potentialNewCells?: number): readonly EngineOp[] | null {
    if (ops.length === 0) {
      return null
    }
    return this.runtime.mutation.executeLocalNow(ops, potentialNewCells, { returnUndoOps: false })
  }
}
