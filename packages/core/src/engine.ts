import {
  type CellNumberFormatInput,
  type CellNumberFormatRecord,
  type CellRangeRef,
  type CellStyleField,
  type CellStylePatch,
  type CellStyleRecord,
  ErrorCode,
  MAX_COLS,
  MAX_ROWS,
  ValueTag,
  type CellSnapshot,
  type CellValue,
  type DependencySnapshot,
  type EngineEvent,
  type ExplainCellSnapshot,
  type LiteralInput,
  type RecalcMetrics,
  type SyncState,
  type SelectionState,
  type WorkbookAxisEntrySnapshot,
  type WorkbookCalculationSettingsSnapshot,
  type WorkbookDefinedNameValueSnapshot,
  type WorkbookFreezePaneSnapshot,
  type WorkbookPivotSnapshot,
  type WorkbookSortSnapshot,
  type WorkbookSnapshot,
} from "@bilig/protocol";
import {
  type FormulaNode,
  evaluatePlanResult,
  formatAddress,
  isArrayValue,
  parseCellAddress,
  parseRangeAddress,
  translateFormulaReferences,
} from "@bilig/formula";
import { Float64Arena, Uint32Arena } from "@bilig/formula/program-arena";
import type { EngineOp, EngineOpBatch } from "@bilig/workbook-domain";
import {
  createReplicaState,
  type OpOrder,
  type ReplicaState,
} from "./replica-state.js";
import { CellFlags } from "./cell-store.js";
import { CycleDetector } from "./cycle-detection.js";
import { EdgeArena, type EdgeSlice } from "./edge-arena.js";
import { entityPayload, isRangeEntity, makeCellEntity, makeRangeEntity } from "./entity-ids.js";
import { appendPackedCellIndex, growUint32 } from "./engine-buffer-utils.js";
import {
  definedNameValueToCellValue,
  definedNameValuesEqual,
  renameDefinedNameValueSheet,
} from "./engine-metadata-utils.js";
import { normalizeRange } from "./engine-range-utils.js";
import {
  buildFormatClearOps,
  buildFormatPatchOps,
  buildStyleClearOps,
  buildStylePatchOps,
} from "./engine-range-format-ops.js";
import {
  areCellValuesEqual,
  emptyValue,
  errorValue,
} from "./engine-value-utils.js";
import { EngineEventBus } from "./events.js";
import { FormulaTable } from "./formula-table.js";
import { RangeRegistry } from "./range-registry.js";
import { RecalcScheduler } from "./scheduler.js";
import {
  selectCellSnapshot,
  selectMetrics,
  selectSelectionState,
  selectViewportCells,
} from "./selectors.js";
import { StringPool } from "./string-pool.js";
import { WasmKernelFacade } from "./wasm-facade.js";
import {
  WorkbookStore,
  normalizeDefinedName,
  type WorkbookAxisMetadataRecord,
  type WorkbookCalculationSettingsRecord,
  type WorkbookDefinedNameRecord,
  type WorkbookFilterRecord,
  type WorkbookPivotRecord,
  type WorkbookPropertyRecord,
  type WorkbookSortRecord,
  type WorkbookSpillRecord,
  type WorkbookTableRecord,
  type WorkbookVolatileContextRecord,
} from "./workbook-store.js";
import { cellToCsvValue, serializeCsv } from "./csv.js";
import {
  createEngineRuntimeState,
  createInitialRecalcMetrics,
  createInitialSelectionState,
  type CommitOp,
  type EngineReplicaSnapshot,
  type EngineRuntimeState,
  type EngineSyncClient,
  type EngineSyncClientConnection,
  type PivotTableInput,
  type RuntimeFormula,
  type SpreadsheetEngineOptions,
  type SpillMaterialization,
  type TransactionLogEntry,
  type U32,
} from "./engine/runtime-state.js";
import {
  createEngineServiceRuntime,
  runEngineEffect,
  runEngineEffectPromise,
  type EngineServiceRuntime,
} from "./engine/live.js";

export type {
  CommitOp,
  EngineReplicaSnapshot,
  EngineSyncClient,
  EngineSyncClientConnection,
  SpreadsheetEngineOptions,
} from "./engine/runtime-state.js";

export class SpreadsheetEngine {
  readonly workbook: WorkbookStore;
  readonly strings = new StringPool();
  readonly events = new EngineEventBus();
  private readonly replicaState: ReplicaState;
  readonly ranges = new RangeRegistry();
  readonly scheduler = new RecalcScheduler();
  readonly wasm = new WasmKernelFacade();

  private readonly formulas: FormulaTable<RuntimeFormula>;
  private readonly cycleDetector = new CycleDetector();
  private readonly edgeArena = new EdgeArena();
  private readonly programArena = new Uint32Arena();
  private readonly constantArena = new Float64Arena();
  private readonly rangeListArena = new Uint32Arena();
  private reverseCellEdges: Array<EdgeSlice | undefined> = [];
  private reverseRangeEdges: Array<EdgeSlice | undefined> = [];
  private readonly reverseDefinedNameEdges = new Map<string, Set<number>>();
  private readonly reverseTableEdges = new Map<string, Set<number>>();
  private readonly reverseSpillEdges = new Map<string, Set<number>>();
  private readonly pivotOutputOwners = new Map<number, string>();
  private readonly batchListeners = new Set<(batch: EngineOpBatch) => void>();
  private readonly selectionListeners = new Set<() => void>();
  private readonly entityVersions = new Map<string, OpOrder>();
  private readonly sheetDeleteVersions = new Map<string, OpOrder>();
  private selection: SelectionState = createInitialSelectionState();
  private syncState: SyncState = "local-only";
  private syncClientConnection: EngineSyncClientConnection | null = null;
  private readonly undoStack: TransactionLogEntry[] = [];
  private readonly redoStack: TransactionLogEntry[] = [];
  private transactionReplayDepth = 0;
  private pendingKernelSync: U32 = new Uint32Array(128);
  private wasmBatch: U32 = new Uint32Array(128);
  private mutationRoots: U32 = new Uint32Array(128);
  private changedInputEpoch = 1;
  private changedInputSeen: U32 = new Uint32Array(128);
  private changedInputBuffer: U32 = new Uint32Array(128);
  private changedFormulaEpoch = 1;
  private changedFormulaSeen: U32 = new Uint32Array(128);
  private changedFormulaBuffer: U32 = new Uint32Array(128);
  private changedUnionEpoch = 1;
  private changedUnionSeen: U32 = new Uint32Array(128);
  private changedUnion: U32 = new Uint32Array(128);
  private materializedCellCount = 0;
  private materializedCells: U32 = new Uint32Array(128);
  private explicitChangedEpoch = 1;
  private explicitChangedSeen: U32 = new Uint32Array(128);
  private explicitChangedBuffer: U32 = new Uint32Array(128);
  private dependencyBuildEpoch = 1;
  private dependencyBuildSeen: U32 = new Uint32Array(128);
  private dependencyBuildCells: U32 = new Uint32Array(128);
  private dependencyBuildEntities: U32 = new Uint32Array(128);
  private dependencyBuildRanges: U32 = new Uint32Array(128);
  private dependencyBuildNewRanges: U32 = new Uint32Array(128);
  private symbolicRefBindings: U32 = new Uint32Array(128);
  private symbolicRangeBindings: U32 = new Uint32Array(128);
  private impactedFormulaEpoch = 1;
  private impactedFormulaSeen: U32 = new Uint32Array(128);
  private impactedFormulaBuffer: U32 = new Uint32Array(128);
  private wasmProgramTargets: U32 = new Uint32Array(128);
  private wasmProgramOffsets: U32 = new Uint32Array(128);
  private wasmProgramLengths: U32 = new Uint32Array(128);
  private wasmConstantOffsets: U32 = new Uint32Array(128);
  private wasmConstantLengths: U32 = new Uint32Array(128);
  private wasmRangeOffsets: U32 = new Uint32Array(128);
  private wasmRangeLengths: U32 = new Uint32Array(128);
  private wasmRangeRowCounts: U32 = new Uint32Array(128);
  private wasmRangeColCounts: U32 = new Uint32Array(128);
  private topoIndegree: U32 = new Uint32Array(128);
  private topoQueue: U32 = new Uint32Array(128);
  private topoFormulaBuffer: U32 = new Uint32Array(128);
  private topoEntityQueue: U32 = new Uint32Array(128);
  private topoFormulaSeenEpoch = 1;
  private topoRangeSeenEpoch = 1;
  private topoFormulaSeen: U32 = new Uint32Array(128);
  private topoRangeSeen: U32 = new Uint32Array(128);
  private batchMutationDepth = 0;
  private wasmProgramSyncPending = false;
  private lastMetrics: RecalcMetrics = createInitialRecalcMetrics();
  private readonly state: EngineRuntimeState;
  private readonly runtime: EngineServiceRuntime;

  constructor(options: SpreadsheetEngineOptions = {}) {
    this.workbook = new WorkbookStore(options.workbookName ?? "Workbook");
    this.formulas = new FormulaTable(this.workbook.cellStore);
    this.replicaState = createReplicaState(options.replicaId ?? "local");
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
      getSelection: () => this.selection,
      setSelection: (selection) => {
        this.selection = selection;
      },
      getSyncState: () => this.syncState,
      setSyncState: (state) => {
        this.syncState = state;
      },
      getSyncClientConnection: () => this.syncClientConnection,
      setSyncClientConnection: (connection) => {
        this.syncClientConnection = connection;
      },
      getTransactionReplayDepth: () => this.transactionReplayDepth,
      setTransactionReplayDepth: (depth) => {
        this.transactionReplayDepth = depth;
      },
      getLastMetrics: () => this.lastMetrics,
      setLastMetrics: (metrics) => {
        this.lastMetrics = metrics;
      },
    });
    this.runtime = createEngineServiceRuntime({
      state: this.state,
      getCellByIndex: (cellIndex) => this.getCellByIndex(cellIndex),
      exportSnapshot: () => this.exportSnapshot(),
      importSnapshot: (snapshot) => this.importSnapshot(snapshot),
      resetWorkbook: () => this.resetWorkbook(),
      captureSheetCellState: (sheetName) => this.captureSheetCellState(sheetName),
      captureRowRangeCellState: (sheetName, start, count) =>
        this.captureRowRangeCellState(sheetName, start, count),
      captureColumnRangeCellState: (sheetName, start, count) =>
        this.captureColumnRangeCellState(sheetName, start, count),
      restoreCellOps: (sheetName, address) => this.restoreCellOps(sheetName, address),
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
        },
        ensureCellTracked: (sheetName, address) => this.ensureCellTracked(sheetName, address),
        ensureCellTrackedByCoords: (sheetId, row, col) =>
          this.ensureCellTrackedByCoords(sheetId, row, col),
        forEachSheetCell: (sheetId, fn) => this.forEachSheetCell(sheetId, fn),
        markFormulaChanged: (cellIndex, count) => this.markFormulaChanged(cellIndex, count),
        resolveStructuredReference: (tableName, columnName) =>
          this.resolveStructuredReference(tableName, columnName),
        resolveSpillReference: (currentSheetName, sheetName, address) =>
          this.resolveSpillReference(currentSheetName, sheetName, address),
        getDependencyBuildEpoch: () => this.dependencyBuildEpoch,
        setDependencyBuildEpoch: (next) => {
          this.dependencyBuildEpoch = next;
        },
        getDependencyBuildSeen: () => this.dependencyBuildSeen,
        setDependencyBuildSeen: (next) => {
          this.dependencyBuildSeen = next;
        },
        getDependencyBuildCells: () => this.dependencyBuildCells,
        setDependencyBuildCells: (next) => {
          this.dependencyBuildCells = next;
        },
        getDependencyBuildEntities: () => this.dependencyBuildEntities,
        setDependencyBuildEntities: (next) => {
          this.dependencyBuildEntities = next;
        },
        getDependencyBuildRanges: () => this.dependencyBuildRanges,
        setDependencyBuildRanges: (next) => {
          this.dependencyBuildRanges = next;
        },
        getDependencyBuildNewRanges: () => this.dependencyBuildNewRanges,
        setDependencyBuildNewRanges: (next) => {
          this.dependencyBuildNewRanges = next;
        },
        getSymbolicRefBindings: () => this.symbolicRefBindings,
        setSymbolicRefBindings: (next) => {
          this.symbolicRefBindings = next;
        },
        getSymbolicRangeBindings: () => this.symbolicRangeBindings,
        setSymbolicRangeBindings: (next) => {
          this.symbolicRangeBindings = next;
        },
        scheduleWasmProgramSync: () => this.scheduleWasmProgramSync(),
      },
      formulaGraph: {
        state: this.state,
        cycleDetector: this.cycleDetector,
        edgeArena: this.edgeArena,
        programArena: this.programArena,
        constantArena: this.constantArena,
        rangeListArena: this.rangeListArena,
        reverseState: {
          reverseCellEdges: this.reverseCellEdges,
          reverseRangeEdges: this.reverseRangeEdges,
        },
        getTopoIndegree: () => this.topoIndegree,
        setTopoIndegree: (next) => {
          this.topoIndegree = next;
        },
        getTopoQueue: () => this.topoQueue,
        setTopoQueue: (next) => {
          this.topoQueue = next;
        },
        getTopoFormulaBuffer: () => this.topoFormulaBuffer,
        setTopoFormulaBuffer: (next) => {
          this.topoFormulaBuffer = next;
        },
        getTopoEntityQueue: () => this.topoEntityQueue,
        setTopoEntityQueue: (next) => {
          this.topoEntityQueue = next;
        },
        getTopoFormulaSeenEpoch: () => this.topoFormulaSeenEpoch,
        setTopoFormulaSeenEpoch: (next) => {
          this.topoFormulaSeenEpoch = next;
        },
        getTopoRangeSeenEpoch: () => this.topoRangeSeenEpoch,
        setTopoRangeSeenEpoch: (next) => {
          this.topoRangeSeenEpoch = next;
        },
        getTopoFormulaSeen: () => this.topoFormulaSeen,
        setTopoFormulaSeen: (next) => {
          this.topoFormulaSeen = next;
        },
        getTopoRangeSeen: () => this.topoRangeSeen,
        setTopoRangeSeen: (next) => {
          this.topoRangeSeen = next;
        },
        getWasmProgramTargets: () => this.wasmProgramTargets,
        setWasmProgramTargets: (next) => {
          this.wasmProgramTargets = next;
        },
        getWasmProgramOffsets: () => this.wasmProgramOffsets,
        setWasmProgramOffsets: (next) => {
          this.wasmProgramOffsets = next;
        },
        getWasmProgramLengths: () => this.wasmProgramLengths,
        setWasmProgramLengths: (next) => {
          this.wasmProgramLengths = next;
        },
        getWasmConstantOffsets: () => this.wasmConstantOffsets,
        setWasmConstantOffsets: (next) => {
          this.wasmConstantOffsets = next;
        },
        getWasmConstantLengths: () => this.wasmConstantLengths,
        setWasmConstantLengths: (next) => {
          this.wasmConstantLengths = next;
        },
        getWasmRangeOffsets: () => this.wasmRangeOffsets,
        setWasmRangeOffsets: (next) => {
          this.wasmRangeOffsets = next;
        },
        getWasmRangeLengths: () => this.wasmRangeLengths,
        setWasmRangeLengths: (next) => {
          this.wasmRangeLengths = next;
        },
        getWasmRangeRowCounts: () => this.wasmRangeRowCounts,
        setWasmRangeRowCounts: (next) => {
          this.wasmRangeRowCounts = next;
        },
        getWasmRangeColCounts: () => this.wasmRangeColCounts,
        setWasmRangeColCounts: (next) => {
          this.wasmRangeColCounts = next;
        },
        getBatchMutationDepth: () => this.batchMutationDepth,
        getWasmProgramSyncPending: () => this.wasmProgramSyncPending,
        setWasmProgramSyncPending: (next) => {
          this.wasmProgramSyncPending = next;
        },
      },
      readRangeCells: (range) => this.readRangeCells(range),
      toCellStateOps: (sheetName, address, snapshot, sourceSheetName, sourceAddress) =>
        this.toCellStateOps(sheetName, address, snapshot, sourceSheetName, sourceAddress),
      forEachFormulaDependencyCell: (cellIndex, fn) => this.forEachFormulaDependencyCell(cellIndex, fn),
      cellToCsvValue: (cell) => cellToCsvValue(cell),
      serializeCsv: (rows) => serializeCsv(rows),
      pivotState: {
        pivotOutputOwners: this.pivotOutputOwners,
      },
      ensureCellTrackedByCoords: (sheetId, row, col) =>
        this.ensureCellTrackedByCoords(sheetId, row, col),
      forEachSheetCell: (sheetId, fn) => this.forEachSheetCell(sheetId, fn),
      beginMutationCollection: () => this.beginMutationCollection(),
      markInputChanged: (cellIndex, count) => this.markInputChanged(cellIndex, count),
      markFormulaChanged: (cellIndex, count) => this.markFormulaChanged(cellIndex, count),
      markExplicitChanged: (cellIndex, count) => this.markExplicitChanged(cellIndex, count),
      composeMutationRoots: (changedInputCount, formulaChangedCount) =>
        this.composeMutationRoots(changedInputCount, formulaChangedCount),
      composeEventChanges: (recalculated, explicitChangedCount) =>
        this.composeEventChanges(recalculated, explicitChangedCount),
      getChangedInputBuffer: () => this.changedInputBuffer,
      unionChangedSets: (...sets) => this.unionChangedSets(...sets),
      composeChangedRootsAndOrdered: (changedRoots, ordered, orderedCount) =>
        this.composeChangedRootsAndOrdered(changedRoots, ordered, orderedCount),
      emptyChangedSet: () => this.changedUnion.subarray(0, 0),
      ensureRecalcScratchCapacity: (size) => this.ensureRecalcScratchCapacity(size),
      getPendingKernelSync: () => this.pendingKernelSync,
      getWasmBatch: () => this.wasmBatch,
      materializeSpill: (cellIndex, arrayValue) => this.materializeSpill(cellIndex, arrayValue),
      clearOwnedSpill: (cellIndex) => this.clearOwnedSpill(cellIndex),
      evaluateUnsupportedFormula: (cellIndex) => this.evaluateUnsupportedFormula(cellIndex),
      getEntityDependents: (entityId) => this.getEntityDependents(entityId),
      clearOwnedPivot: (pivot) => this.clearOwnedPivot(pivot),
      clearPivotForCell: (cellIndex) => this.clearPivotForCell(cellIndex),
      materializePivot: (pivot) => this.materializePivot(pivot),
      scheduleWasmProgramSync: () => this.scheduleWasmProgramSync(),
      flushWasmProgramSync: () => this.flushWasmProgramSync(),
      applyRemoteSnapshot: (snapshot) => {
        this.importSnapshot(snapshot);
      },
      operation: {
        state: this.state,
        reverseState: {
          reverseSpillEdges: this.reverseSpillEdges,
        },
        getSelectionState: () => this.getSelectionState(),
        setSelection: (sheetName, address) => this.setSelection(sheetName, address),
        rewriteDefinedNamesForSheetRename: (oldSheetName, newSheetName) =>
          this.rewriteDefinedNamesForSheetRename(oldSheetName, newSheetName),
        rewriteCellFormulasForSheetRename: (oldSheetName, newSheetName, formulaChangedCount) =>
          this.rewriteCellFormulasForSheetRename(
            oldSheetName,
            newSheetName,
            formulaChangedCount,
          ),
        rebindDefinedNameDependents: (names, formulaChangedCount) =>
          this.rebindDefinedNameDependents(names, formulaChangedCount),
        rebindTableDependents: (tableNames, formulaChangedCount) =>
          this.rebindTableDependents(tableNames, formulaChangedCount),
        rebindFormulaCells: (candidates, formulaChangedCount) =>
          this.rebindFormulaCells(candidates, formulaChangedCount),
        rebindFormulasForSheet: (sheetName, formulaChangedCount, candidates) =>
          this.rebindFormulasForSheet(sheetName, formulaChangedCount, candidates),
        removeSheetRuntime: (sheetName, explicitChangedCount) =>
          this.removeSheetRuntime(sheetName, explicitChangedCount),
        applyStructuralAxisOp: (op) => this.applyStructuralAxisOp(op),
        clearOwnedSpill: (cellIndex) => this.clearOwnedSpill(cellIndex),
        clearPivotForCell: (cellIndex) => this.clearPivotForCell(cellIndex),
        clearOwnedPivot: (pivot) => this.clearOwnedPivot(pivot),
        removeFormula: (cellIndex) => this.removeFormula(cellIndex),
        bindFormula: (cellIndex, ownerSheetName, source) =>
          this.bindFormula(cellIndex, ownerSheetName, source),
        setInvalidFormulaValue: (cellIndex) => this.setInvalidFormulaValue(cellIndex),
        beginMutationCollection: () => this.beginMutationCollection(),
        markInputChanged: (cellIndex, count) => this.markInputChanged(cellIndex, count),
        markFormulaChanged: (cellIndex, count) => this.markFormulaChanged(cellIndex, count),
        markVolatileFormulasChanged: (count) => this.markVolatileFormulasChanged(count),
        markSpillRootsChanged: (cellIndices, count) =>
          this.markSpillRootsChanged(cellIndices, count),
        markPivotRootsChanged: (cellIndices, count) =>
          this.markPivotRootsChanged(cellIndices, count),
        markExplicitChanged: (cellIndex, count) => this.markExplicitChanged(cellIndex, count),
        composeMutationRoots: (changedInputCount, formulaChangedCount) =>
          this.composeMutationRoots(changedInputCount, formulaChangedCount),
        composeEventChanges: (recalculated, explicitChangedCount) =>
          this.composeEventChanges(recalculated, explicitChangedCount),
        getChangedInputBuffer: () => this.changedInputBuffer,
        ensureCellTracked: (sheetName, address) => this.ensureCellTracked(sheetName, address),
        estimatePotentialNewCells: (ops) => this.estimatePotentialNewCells(ops),
        resetMaterializedCellScratch: (expectedSize) =>
          this.resetMaterializedCellScratch(expectedSize),
        syncDynamicRanges: (formulaChangedCount) => this.syncDynamicRanges(formulaChangedCount),
        rebuildTopoRanks: () => this.rebuildTopoRanks(),
        detectCycles: () => this.detectCycles(),
        recalculate: (changedRoots, kernelSyncRoots) =>
          this.recalculate(changedRoots, kernelSyncRoots),
        reconcilePivotOutputs: (baseChanged, forceAllPivots) =>
          this.reconcilePivotOutputs(baseChanged, forceAllPivots),
        flushWasmProgramSync: () => this.flushWasmProgramSync(),
        getBatchMutationDepth: () => this.batchMutationDepth,
        setBatchMutationDepth: (next) => {
          this.batchMutationDepth = next;
        },
      },
    });
    void this.wasm.init();
  }

  async ready(): Promise<void> {
    await this.wasm.init();
  }

  subscribe(listener: (event: EngineEvent) => void): () => void {
    return runEngineEffect(this.runtime.events.subscribe(listener));
  }

  subscribeCell(sheetName: string, address: string, listener: () => void): () => void {
    return runEngineEffect(this.runtime.events.subscribeCell(sheetName, address, listener));
  }

  subscribeCells(
    sheetName: string,
    addresses: readonly string[],
    listener: () => void,
  ): () => void {
    return runEngineEffect(this.runtime.events.subscribeCells(sheetName, addresses, listener));
  }

  subscribeBatches(listener: (batch: EngineOpBatch) => void): () => void {
    return runEngineEffect(this.runtime.events.subscribeBatches(listener));
  }

  subscribeSelection(listener: () => void): () => void {
    return runEngineEffect(this.runtime.selection.subscribe(listener));
  }

  getSelectionState(): SelectionState {
    return runEngineEffect(this.runtime.selection.getSelectionState());
  }

  setSelection(
    sheetName: string,
    address: string | null,
    options: {
      anchorAddress?: string | null;
      range?: { startAddress: string; endAddress: string } | null;
      editMode?: SelectionState["editMode"];
    } = {},
  ): void {
    runEngineEffect(this.runtime.selection.setSelection(sheetName, address, options));
  }

  getLastMetrics(): RecalcMetrics {
    return this.lastMetrics;
  }

  getSyncState(): SyncState {
    return this.syncState;
  }

  async connectSyncClient(client: EngineSyncClient): Promise<void> {
    await runEngineEffectPromise(this.runtime.sync.connectClient(client));
  }

  async disconnectSyncClient(): Promise<void> {
    await runEngineEffectPromise(this.runtime.sync.disconnectClient());
  }

  createSheet(name: string): void {
    this.executeLocalTransaction([
      { kind: "upsertSheet", name, order: this.workbook.sheetsByName.size },
    ]);
  }

  renameSheet(oldName: string, newName: string): void {
    const trimmedName = newName.trim();
    if (trimmedName.length === 0 || oldName === trimmedName) {
      return;
    }
    if (this.workbook.getSheet(trimmedName)) {
      return;
    }
    this.executeLocalTransaction([{ kind: "renameSheet", oldName, newName: trimmedName }]);
  }

  deleteSheet(name: string): void {
    this.executeLocalTransaction([{ kind: "deleteSheet", name }]);
  }

  setCellValue(sheetName: string, address: string, value: LiteralInput): CellValue {
    this.executeLocalTransaction([{ kind: "setCellValue", sheetName, address, value }]);
    return this.getCellValue(sheetName, address);
  }

  setCellFormula(sheetName: string, address: string, formula: string): CellValue {
    this.executeLocalTransaction([{ kind: "setCellFormula", sheetName, address, formula }]);
    return this.getCellValue(sheetName, address);
  }

  setCellFormat(sheetName: string, address: string, format: string | null): void {
    this.executeLocalTransaction([{ kind: "setCellFormat", sheetName, address, format }]);
  }

  setRangeNumberFormat(range: CellRangeRef, format: CellNumberFormatInput): void {
    const ops = buildFormatPatchOps(this.workbook, range, format);
    this.executeLocalTransaction(ops);
  }

  clearRangeNumberFormat(range: CellRangeRef): void {
    const ops = buildFormatClearOps(this.workbook, range);
    this.executeLocalTransaction(ops);
  }

  setRangeStyle(range: CellRangeRef, patch: CellStylePatch): void {
    const ops = buildStylePatchOps(this.workbook, range, patch);
    this.executeLocalTransaction(ops);
  }

  clearRangeStyle(range: CellRangeRef, fields?: readonly CellStyleField[]): void {
    const ops = buildStyleClearOps(this.workbook, range, fields);
    this.executeLocalTransaction(ops);
  }

  getCellStyle(styleId: string | undefined): CellStyleRecord | undefined {
    return this.workbook.getCellStyle(styleId);
  }

  getCellNumberFormat(id: string | undefined): CellNumberFormatRecord | undefined {
    return this.workbook.getCellNumberFormat(id);
  }

  setDefinedName(name: string, value: WorkbookDefinedNameValueSnapshot): void {
    const normalizedName = normalizeDefinedName(name);
    const previous = this.workbook.getDefinedName(normalizedName);
    const trimmedName = name.trim();
    if (previous?.name === trimmedName && definedNameValuesEqual(previous.value, value)) {
      return;
    }
    this.executeLocalTransaction([{ kind: "upsertDefinedName", name: trimmedName, value }]);
  }

  deleteDefinedName(name: string): boolean {
    if (!this.workbook.getDefinedName(name)) {
      return false;
    }
    this.executeLocalTransaction([{ kind: "deleteDefinedName", name }]);
    return true;
  }

  getDefinedName(name: string): WorkbookDefinedNameRecord | undefined {
    return this.workbook.getDefinedName(name);
  }

  getDefinedNames(): WorkbookDefinedNameRecord[] {
    return this.workbook.listDefinedNames();
  }

  setWorkbookMetadata(key: string, value: LiteralInput): void {
    const existing = this.workbook.getWorkbookProperty(key);
    if (existing?.value === value || (existing === undefined && value === null)) {
      return;
    }
    this.executeLocalTransaction([{ kind: "setWorkbookMetadata", key, value }]);
  }

  getWorkbookMetadata(key: string): WorkbookPropertyRecord | undefined {
    return this.workbook.getWorkbookProperty(key);
  }

  getWorkbookMetadataEntries(): WorkbookPropertyRecord[] {
    return this.workbook.listWorkbookProperties();
  }

  setCalculationSettings(settings: WorkbookCalculationSettingsSnapshot): void {
    const current = this.workbook.getCalculationSettings();
    const nextSettings = { compatibilityMode: "excel-modern" as const, ...settings };
    if (
      current.mode === nextSettings.mode &&
      current.compatibilityMode === nextSettings.compatibilityMode
    ) {
      return;
    }
    this.executeLocalTransaction([{ kind: "setCalculationSettings", settings: nextSettings }]);
  }

  getCalculationSettings(): WorkbookCalculationSettingsRecord {
    return this.workbook.getCalculationSettings();
  }

  getVolatileContext(): WorkbookVolatileContextRecord {
    return this.workbook.getVolatileContext();
  }

  recalculateNow(): number[] {
    return runEngineEffect(this.runtime.recalc.recalculateNow());
  }

  recalculateDifferential(): { js: CellSnapshot[]; wasm: CellSnapshot[]; drift: string[] } {
    return runEngineEffect(this.runtime.recalc.recalculateDifferential());
  }

  recalculateDirty(
    dirtyRegions: Array<{
      sheetName: string;
      rowStart: number;
      rowEnd: number;
      colStart: number;
      colEnd: number;
    }>,
  ): number[] {
    return runEngineEffect(this.runtime.recalc.recalculateDirty(dirtyRegions));
  }

  updateRowMetadata(
    sheetName: string,
    start: number,
    count: number,
    size: number | null,
    hidden: boolean | null,
  ): void {
    const existing = this.workbook.getRowMetadata(sheetName, start, count);
    if (existing?.size === size && existing.hidden === hidden) {
      return;
    }
    if (existing === undefined && size === null && hidden === null) {
      return;
    }
    this.executeLocalTransaction([
      { kind: "updateRowMetadata", sheetName, start, count, size, hidden },
    ]);
  }

  getRowMetadata(sheetName: string): WorkbookAxisMetadataRecord[] {
    return this.workbook.listRowMetadata(sheetName);
  }

  getRowAxisEntries(sheetName: string): WorkbookAxisEntrySnapshot[] {
    return this.workbook.listRowAxisEntries(sheetName);
  }

  insertRows(sheetName: string, start: number, count: number): void {
    if (count <= 0) {
      return;
    }
    this.executeLocalTransaction([{ kind: "insertRows", sheetName, start, count }]);
  }

  deleteRows(sheetName: string, start: number, count: number): void {
    if (count <= 0) {
      return;
    }
    this.executeLocalTransaction([{ kind: "deleteRows", sheetName, start, count }]);
  }

  moveRows(sheetName: string, start: number, count: number, target: number): void {
    if (count <= 0 || start === target) {
      return;
    }
    this.executeLocalTransaction([{ kind: "moveRows", sheetName, start, count, target }]);
  }

  updateColumnMetadata(
    sheetName: string,
    start: number,
    count: number,
    size: number | null,
    hidden: boolean | null,
  ): void {
    const existing = this.workbook.getColumnMetadata(sheetName, start, count);
    if (existing?.size === size && existing.hidden === hidden) {
      return;
    }
    if (existing === undefined && size === null && hidden === null) {
      return;
    }
    this.executeLocalTransaction([
      { kind: "updateColumnMetadata", sheetName, start, count, size, hidden },
    ]);
  }

  getColumnMetadata(sheetName: string): WorkbookAxisMetadataRecord[] {
    return this.workbook.listColumnMetadata(sheetName);
  }

  getColumnAxisEntries(sheetName: string): WorkbookAxisEntrySnapshot[] {
    return this.workbook.listColumnAxisEntries(sheetName);
  }

  insertColumns(sheetName: string, start: number, count: number): void {
    if (count <= 0) {
      return;
    }
    this.executeLocalTransaction([{ kind: "insertColumns", sheetName, start, count }]);
  }

  deleteColumns(sheetName: string, start: number, count: number): void {
    if (count <= 0) {
      return;
    }
    this.executeLocalTransaction([{ kind: "deleteColumns", sheetName, start, count }]);
  }

  moveColumns(sheetName: string, start: number, count: number, target: number): void {
    if (count <= 0 || start === target) {
      return;
    }
    this.executeLocalTransaction([{ kind: "moveColumns", sheetName, start, count, target }]);
  }

  setFreezePane(sheetName: string, rows: number, cols: number): void {
    const existing = this.workbook.getFreezePane(sheetName);
    if (existing?.rows === rows && existing.cols === cols) {
      return;
    }
    this.executeLocalTransaction([{ kind: "setFreezePane", sheetName, rows, cols }]);
  }

  clearFreezePane(sheetName: string): boolean {
    if (!this.workbook.getFreezePane(sheetName)) {
      return false;
    }
    this.executeLocalTransaction([{ kind: "clearFreezePane", sheetName }]);
    return true;
  }

  getFreezePane(sheetName: string): WorkbookFreezePaneSnapshot | undefined {
    return this.workbook.getFreezePane(sheetName);
  }

  setFilter(sheetName: string, range: CellRangeRef): void {
    const existing = this.workbook.getFilter(sheetName, range);
    if (existing) {
      return;
    }
    this.executeLocalTransaction([{ kind: "setFilter", sheetName, range: { ...range } }]);
  }

  clearFilter(sheetName: string, range: CellRangeRef): boolean {
    if (!this.workbook.getFilter(sheetName, range)) {
      return false;
    }
    this.executeLocalTransaction([{ kind: "clearFilter", sheetName, range: { ...range } }]);
    return true;
  }

  getFilters(sheetName: string): WorkbookFilterRecord[] {
    return this.workbook.listFilters(sheetName);
  }

  setSort(sheetName: string, range: CellRangeRef, keys: WorkbookSortSnapshot["keys"]): void {
    const existing = this.workbook.getSort(sheetName, range);
    const normalizedKeys = keys.map((key) => Object.assign({}, key));
    if (
      existing &&
      existing.keys.length === normalizedKeys.length &&
      existing.keys.every(
        (key, index) =>
          key.keyAddress === normalizedKeys[index]?.keyAddress &&
          key.direction === normalizedKeys[index]?.direction,
      )
    ) {
      return;
    }
    this.executeLocalTransaction([
      { kind: "setSort", sheetName, range: { ...range }, keys: normalizedKeys },
    ]);
  }

  clearSort(sheetName: string, range: CellRangeRef): boolean {
    if (!this.workbook.getSort(sheetName, range)) {
      return false;
    }
    this.executeLocalTransaction([{ kind: "clearSort", sheetName, range: { ...range } }]);
    return true;
  }

  getSorts(sheetName: string): WorkbookSortRecord[] {
    return this.workbook.listSorts(sheetName);
  }

  setTable(table: WorkbookTableRecord): void {
    const existing = this.workbook.getTable(table.name);
    if (
      existing &&
      existing.sheetName === table.sheetName &&
      existing.startAddress === table.startAddress &&
      existing.endAddress === table.endAddress &&
      existing.headerRow === table.headerRow &&
      existing.totalsRow === table.totalsRow &&
      existing.columnNames.length === table.columnNames.length &&
      existing.columnNames.every((name, index) => name === table.columnNames[index])
    ) {
      return;
    }
    this.executeLocalTransaction([
      {
        kind: "upsertTable",
        table: Object.assign({}, table, { columnNames: [...table.columnNames] }),
      },
    ]);
  }

  deleteTable(name: string): boolean {
    if (!this.workbook.getTable(name)) {
      return false;
    }
    this.executeLocalTransaction([{ kind: "deleteTable", name }]);
    return true;
  }

  getTable(name: string): WorkbookTableRecord | undefined {
    return this.workbook.getTable(name);
  }

  getTables(): WorkbookTableRecord[] {
    return this.workbook.listTables();
  }

  setSpillRange(sheetName: string, address: string, rows: number, cols: number): void {
    const existing = this.workbook.getSpill(sheetName, address);
    if (existing?.rows === rows && existing.cols === cols) {
      return;
    }
    this.executeLocalTransaction([{ kind: "upsertSpillRange", sheetName, address, rows, cols }]);
  }

  deleteSpillRange(sheetName: string, address: string): boolean {
    if (!this.workbook.getSpill(sheetName, address)) {
      return false;
    }
    this.executeLocalTransaction([{ kind: "deleteSpillRange", sheetName, address }]);
    return true;
  }

  getSpillRanges(): WorkbookSpillRecord[] {
    return this.workbook.listSpills();
  }

  setPivotTable(sheetName: string, address: string, definition: PivotTableInput): void {
    this.executeLocalTransaction([
      {
        kind: "upsertPivotTable",
        name: definition.name.trim(),
        sheetName,
        address,
        source: { ...definition.source },
        groupBy: [...definition.groupBy],
        values: definition.values.map((v) => Object.assign({}, v)),
        rows: 1,
        cols: Math.max(definition.groupBy.length + definition.values.length, 1),
      },
    ]);
  }

  deletePivotTable(sheetName: string, address: string): boolean {
    if (!this.workbook.getPivot(sheetName, address)) {
      return false;
    }
    this.executeLocalTransaction([{ kind: "deletePivotTable", sheetName, address }]);
    return true;
  }

  getPivotTable(sheetName: string, address: string): WorkbookPivotSnapshot | undefined {
    return this.workbook.getPivot(sheetName, address);
  }

  getPivotTables(): WorkbookPivotSnapshot[] {
    return this.workbook.listPivots();
  }

  clearCell(sheetName: string, address: string): void {
    this.executeLocalTransaction([{ kind: "clearCell", sheetName, address }]);
  }

  setRangeValues(range: CellRangeRef, values: readonly (readonly LiteralInput[])[]): void {
    runEngineEffect(this.runtime.mutation.setRangeValues(range, values));
  }

  setRangeFormulas(range: CellRangeRef, formulas: readonly (readonly string[])[]): void {
    runEngineEffect(this.runtime.mutation.setRangeFormulas(range, formulas));
  }

  clearRange(range: CellRangeRef): void {
    runEngineEffect(this.runtime.mutation.clearRange(range));
  }

  fillRange(source: CellRangeRef, target: CellRangeRef): void {
    runEngineEffect(this.runtime.mutation.fillRange(source, target));
  }

  copyRange(source: CellRangeRef, target: CellRangeRef): void {
    runEngineEffect(this.runtime.mutation.copyRange(source, target));
  }

  moveRange(source: CellRangeRef, target: CellRangeRef): void {
    runEngineEffect(this.runtime.mutation.moveRange(source, target));
  }

  pasteRange(source: CellRangeRef, target: CellRangeRef): void {
    runEngineEffect(this.runtime.mutation.copyRange(source, target));
  }

  undo(): boolean {
    return runEngineEffect(this.runtime.history.undo());
  }

  redo(): boolean {
    return runEngineEffect(this.runtime.history.redo());
  }

  exportSheetCsv(sheetName: string): string {
    return runEngineEffect(this.runtime.read.exportSheetCsv(sheetName));
  }

  importSheetCsv(sheetName: string, csv: string): void {
    runEngineEffect(this.runtime.mutation.importSheetCsv(sheetName, csv));
  }

  getCellValue(sheetName: string, address: string): CellValue {
    return runEngineEffect(this.runtime.read.getCellValue(sheetName, address));
  }

  getRangeValues(range: CellRangeRef): CellValue[][] {
    return runEngineEffect(this.runtime.read.getRangeValues(range));
  }

  getCell(sheetName: string, address: string): CellSnapshot {
    return runEngineEffect(this.runtime.read.getCell(sheetName, address));
  }

  getCellByIndex(cellIndex: number): CellSnapshot {
    return runEngineEffect(this.runtime.read.getCellByIndex(cellIndex));
  }

  getDependencies(sheetName: string, address: string): DependencySnapshot {
    return runEngineEffect(this.runtime.read.getDependencies(sheetName, address));
  }

  getDependents(sheetName: string, address: string): DependencySnapshot {
    return runEngineEffect(this.runtime.read.getDependents(sheetName, address));
  }

  explainCell(sheetName: string, address: string): ExplainCellSnapshot {
    return runEngineEffect(this.runtime.read.explainCell(sheetName, address));
  }

  exportSnapshot(): WorkbookSnapshot {
    return runEngineEffect(this.runtime.snapshot.exportWorkbook());
  }

  importSnapshot(snapshot: WorkbookSnapshot): void {
    runEngineEffect(this.runtime.snapshot.importWorkbook(snapshot));
  }

  private captureSheetCellState(sheetName: string): EngineOp[] {
    return runEngineEffect(this.runtime.structure.captureSheetCellState(sheetName));
  }

  private captureRowRangeCellState(sheetName: string, start: number, count: number): EngineOp[] {
    return runEngineEffect(this.runtime.structure.captureRowRangeCellState(sheetName, start, count));
  }

  private captureColumnRangeCellState(sheetName: string, start: number, count: number): EngineOp[] {
    return runEngineEffect(
      this.runtime.structure.captureColumnRangeCellState(sheetName, start, count),
    );
  }

  private applyStructuralAxisOp(
    op: Extract<
      EngineOp,
      {
        kind:
          | "insertRows"
          | "deleteRows"
          | "moveRows"
          | "insertColumns"
          | "deleteColumns"
          | "moveColumns";
      }
    >,
  ): { changedCellIndices: number[]; formulaCellIndices: number[] } {
    return runEngineEffect(this.runtime.structure.applyStructuralAxisOp(op));
  }

  private rewriteDefinedNamesForSheetRename(oldSheetName: string, newSheetName: string): void {
    this.workbook.listDefinedNames().forEach((record) => {
      const nextValue = renameDefinedNameValueSheet(record.value, oldSheetName, newSheetName);
      if (!definedNameValuesEqual(record.value, nextValue)) {
        this.workbook.setDefinedName(record.name, nextValue);
      }
    });
  }

  private rewriteCellFormulasForSheetRename(
    oldSheetName: string,
    newSheetName: string,
    formulaChangedCount: number,
  ): number {
    return runEngineEffect(
      this.runtime.binding.rewriteCellFormulasForSheetRename(
        oldSheetName,
        newSheetName,
        formulaChangedCount,
      ),
    );
  }

  private rebindFormulaCells(candidates: readonly number[], formulaChangedCount: number): number {
    return runEngineEffect(this.runtime.binding.rebindFormulaCells(candidates, formulaChangedCount));
  }

  private rebindDefinedNameDependents(
    names: readonly string[],
    formulaChangedCount: number,
  ): number {
    return runEngineEffect(
      this.runtime.binding.rebindDefinedNameDependents(names, formulaChangedCount),
    );
  }

  private rebindTableDependents(
    tableNames: readonly string[],
    formulaChangedCount: number,
  ): number {
    return runEngineEffect(
      this.runtime.binding.rebindTableDependents(tableNames, formulaChangedCount),
    );
  }

  private bindFormula(cellIndex: number, ownerSheetName: string, source: string): void {
    runEngineEffect(this.runtime.binding.bindFormula(cellIndex, ownerSheetName, source));
  }

  private reconcilePivotOutputs(baseChanged: U32, forceAllPivots = false): U32 {
    return runEngineEffect(this.runtime.recalc.reconcilePivotOutputs(baseChanged, forceAllPivots));
  }

  private materializePivot(pivot: WorkbookPivotRecord): number[] {
    return runEngineEffect(this.runtime.pivot.materializePivot(pivot));
  }

  private resolvePivotData(
    sheetName: string,
    address: string,
    dataField: string,
    filters: ReadonlyArray<{ field: string; item: CellValue }>,
  ): CellValue {
    return runEngineEffect(this.runtime.pivot.resolvePivotData(sheetName, address, dataField, filters));
  }

  private resolveMultipleOperations(request: {
    formulaSheetName: string;
    formulaAddress: string;
    rowCellSheetName: string;
    rowCellAddress: string;
    rowReplacementSheetName: string;
    rowReplacementAddress: string;
    columnCellSheetName?: string;
    columnCellAddress?: string;
    columnReplacementSheetName?: string;
    columnReplacementAddress?: string;
  }): CellValue {
    const replacements = new Map<string, { sheetName: string; address: string }>();
    replacements.set(
      this.referenceReplacementKey(request.rowCellSheetName, request.rowCellAddress),
      {
        sheetName: request.rowReplacementSheetName,
        address: request.rowReplacementAddress,
      },
    );
    if (
      request.columnCellSheetName &&
      request.columnCellAddress &&
      request.columnReplacementSheetName &&
      request.columnReplacementAddress
    ) {
      replacements.set(
        this.referenceReplacementKey(request.columnCellSheetName, request.columnCellAddress),
        {
          sheetName: request.columnReplacementSheetName,
          address: request.columnReplacementAddress,
        },
      );
    }
    return this.evaluateCellWithReferenceReplacements(
      request.formulaSheetName,
      request.formulaAddress,
      replacements,
      new Set<string>(),
    );
  }

  private referenceReplacementKey(sheetName: string, address: string): string {
    return `${sheetName.trim().toUpperCase()}!${address.trim().toUpperCase()}`;
  }

  private evaluateCellWithReferenceReplacements(
    sheetName: string,
    address: string,
    replacements: ReadonlyMap<string, { sheetName: string; address: string }>,
    visiting: Set<string>,
  ): CellValue {
    const replacementKey = this.referenceReplacementKey(sheetName, address);
    const replacement = replacements.get(replacementKey);
    if (replacement) {
      return this.evaluateCellWithReferenceReplacements(
        replacement.sheetName,
        replacement.address,
        replacements,
        visiting,
      );
    }

    const visitKey = this.referenceReplacementKey(sheetName, address);
    if (visiting.has(visitKey)) {
      return errorValue(ErrorCode.Cycle);
    }

    const cellIndex = this.workbook.getCellIndex(sheetName, address);
    if (cellIndex === undefined) {
      return emptyValue();
    }

    const formula = this.formulas.get(cellIndex);
    if (!formula) {
      return this.workbook.cellStore.getValue(cellIndex, (id) => this.strings.get(id));
    }

    visiting.add(visitKey);
    const evaluationContext = {
      sheetName,
      currentAddress: address,
      resolveCell: (targetSheetName, targetAddress) =>
        this.evaluateCellWithReferenceReplacements(
          targetSheetName,
          targetAddress,
          replacements,
          visiting,
        ),
      resolveRange: (targetSheetName, start, end, refKind) => {
        if (refKind !== "cells") {
          return [];
        }
        const range = parseRangeAddress(`${start}:${end}`, targetSheetName);
        if (range.kind !== "cells") {
          return [];
        }
        const values: CellValue[] = [];
        for (let row = range.start.row; row <= range.end.row; row += 1) {
          for (let col = range.start.col; col <= range.end.col; col += 1) {
            values.push(
              this.evaluateCellWithReferenceReplacements(
                targetSheetName,
                formatAddress(row, col),
                replacements,
                visiting,
              ),
            );
          }
        }
        return values;
      },
      resolveName: (name) => {
        const definedName = this.workbook.getDefinedName(name);
        if (!definedName) {
          return errorValue(ErrorCode.Name);
        }
        return definedNameValueToCellValue(definedName.value, this.strings);
      },
      resolveFormula: (targetSheetName: string, targetAddress: string) =>
        this.getCell(targetSheetName, targetAddress).formula,
      resolvePivotData: ({
        dataField,
        sheetName: pivotSheetName,
        address: pivotAddress,
        filters,
      }) => this.resolvePivotData(pivotSheetName, pivotAddress, dataField, filters),
      resolveMultipleOperations: (nested: {
        formulaSheetName: string;
        formulaAddress: string;
        rowCellSheetName: string;
        rowCellAddress: string;
        rowReplacementSheetName: string;
        rowReplacementAddress: string;
        columnCellSheetName?: string;
        columnCellAddress?: string;
        columnReplacementSheetName?: string;
        columnReplacementAddress?: string;
      }) => this.resolveMultipleOperations(nested),
      listSheetNames: () =>
        [...this.workbook.sheetsByName.values()]
          .toSorted((left, right) => left.order - right.order)
          .map((sheet) => sheet.name),
    } as Parameters<typeof evaluatePlanResult>[1];
    const result = evaluatePlanResult(formula.compiled.jsPlan, evaluationContext);
    visiting.delete(visitKey);
    return isArrayValue(result) ? (result.values[0] ?? emptyValue()) : result;
  }

  private clearOwnedPivot(pivot: WorkbookPivotRecord): number[] {
    return runEngineEffect(this.runtime.pivot.clearOwnedPivot(pivot));
  }

  private clearPivotForCell(cellIndex: number): number[] {
    return runEngineEffect(this.runtime.pivot.clearPivotForCell(cellIndex));
  }

  exportReplicaSnapshot(): EngineReplicaSnapshot {
    return runEngineEffect(this.runtime.snapshot.exportReplica());
  }

  importReplicaSnapshot(snapshot: EngineReplicaSnapshot): void {
    runEngineEffect(this.runtime.snapshot.importReplica(snapshot));
  }

  renderCommit(ops: CommitOp[]): void {
    runEngineEffect(this.runtime.mutation.renderCommit(ops));
  }

  applyRemoteBatch(batch: EngineOpBatch): boolean {
    return runEngineEffect(this.runtime.sync.applyRemoteBatch(batch));
  }

  captureUndoOps<T>(mutate: () => T): {
    result: T;
    undoOps: readonly EngineOp[] | null;
  } {
    return runEngineEffect(this.runtime.mutation.captureUndoOps(mutate));
  }

  applyOps(
    ops: readonly EngineOp[],
    options: {
      captureUndo?: boolean;
      potentialNewCells?: number;
    } = {},
  ): readonly EngineOp[] | null {
    return runEngineEffect(this.runtime.mutation.applyOps(ops, options));
  }

  private executeLocalTransaction(
    ops: EngineOp[],
    potentialNewCells?: number,
  ): readonly EngineOp[] | null {
    return runEngineEffect(this.runtime.mutation.executeLocal(ops, potentialNewCells));
  }

  private applyDerivedOp(
    op: Extract<
      EngineOp,
      { kind: "upsertSpillRange" | "deleteSpillRange" | "upsertPivotTable" | "deletePivotTable" }
    >,
  ): number[] {
    return runEngineEffect(this.runtime.operations.applyDerivedOp(op));
  }

  private restoreCellOps(sheetName: string, address: string): EngineOp[] {
    const cellIndex = this.workbook.getCellIndex(sheetName, address);
    if (cellIndex === undefined) {
      return [{ kind: "clearCell", sheetName, address }];
    }
    return this.toCellStateOps(sheetName, address, this.getCellByIndex(cellIndex)).filter(
      (op) => op.kind !== "setCellFormat",
    );
  }

  private readRangeCells(range: CellRangeRef): CellSnapshot[][] {
    const bounds = normalizeRange(range);
    const rows: CellSnapshot[][] = [];
    for (let row = bounds.startRow; row <= bounds.endRow; row += 1) {
      const cells: CellSnapshot[] = [];
      for (let col = bounds.startCol; col <= bounds.endCol; col += 1) {
        cells.push(this.getCell(range.sheetName, formatAddress(row, col)));
      }
      rows.push(cells);
    }
    return rows;
  }

  private toCellStateOps(
    sheetName: string,
    address: string,
    snapshot: CellSnapshot,
    sourceSheetName?: string,
    sourceAddress?: string,
  ): EngineOp[] {
    const ops: EngineOp[] = [];
    if (snapshot.formula !== undefined) {
      const translatedFormula =
        sourceSheetName && sourceAddress
          ? this.translateFormulaForTarget(
              snapshot.formula,
              sourceSheetName,
              sourceAddress,
              sheetName,
              address,
            )
          : snapshot.formula;
      ops.push({ kind: "setCellFormula", sheetName, address, formula: translatedFormula });
    } else {
      switch (snapshot.value.tag) {
        case ValueTag.Empty:
          ops.push({ kind: "clearCell", sheetName, address });
          break;
        case ValueTag.Number:
          ops.push({ kind: "setCellValue", sheetName, address, value: snapshot.value.value });
          break;
        case ValueTag.Boolean:
          ops.push({ kind: "setCellValue", sheetName, address, value: snapshot.value.value });
          break;
        case ValueTag.String:
          ops.push({ kind: "setCellValue", sheetName, address, value: snapshot.value.value });
          break;
        case ValueTag.Error:
          ops.push({ kind: "clearCell", sheetName, address });
          break;
      }
    }
    ops.push({
      kind: "setCellFormat",
      sheetName,
      address,
      format: snapshot.format ?? null,
    });
    return ops;
  }

  private translateFormulaForTarget(
    formula: string,
    sourceSheetName: string,
    sourceAddress: string,
    targetSheetName: string,
    targetAddress: string,
  ): string {
    const source = parseCellAddress(sourceAddress, sourceSheetName);
    const target = parseCellAddress(targetAddress, targetSheetName);
    return translateFormulaReferences(formula, target.row - source.row, target.col - source.col);
  }

  private setInvalidFormulaValue(cellIndex: number): void {
    runEngineEffect(this.runtime.binding.invalidateFormula(cellIndex));
  }

  private removeFormula(cellIndex: number): boolean {
    return runEngineEffect(this.runtime.binding.clearFormula(cellIndex));
  }

  private rebuildTopoRanks(): void {
    runEngineEffect(this.runtime.graph.rebuildTopoRanks());
  }

  private detectCycles(): void {
    runEngineEffect(this.runtime.graph.detectCycles());
  }

  private recalculate(
    changedRoots: readonly number[] | U32,
    kernelSyncRoots: readonly number[] | U32 = changedRoots,
  ): Uint32Array {
    return runEngineEffect(this.runtime.recalc.recalculate(changedRoots, kernelSyncRoots));
  }

  private ensureRecalcScratchCapacity(size: number): void {
    if (size > this.mutationRoots.length) {
      this.mutationRoots = growUint32(this.mutationRoots, size);
    }
    if (size > this.changedInputSeen.length) {
      this.changedInputSeen = growUint32(this.changedInputSeen, size);
    }
    if (size > this.changedInputBuffer.length) {
      this.changedInputBuffer = growUint32(this.changedInputBuffer, size);
    }
    if (size > this.changedFormulaSeen.length) {
      this.changedFormulaSeen = growUint32(this.changedFormulaSeen, size);
    }
    if (size > this.changedFormulaBuffer.length) {
      this.changedFormulaBuffer = growUint32(this.changedFormulaBuffer, size);
    }
    if (size > this.pendingKernelSync.length) {
      this.pendingKernelSync = growUint32(this.pendingKernelSync, size);
    }
    if (size > this.wasmBatch.length) {
      this.wasmBatch = growUint32(this.wasmBatch, size);
    }
    if (size > this.changedUnion.length) {
      this.changedUnion = growUint32(this.changedUnion, size);
    }
    if (size > this.changedUnionSeen.length) {
      this.changedUnionSeen = growUint32(this.changedUnionSeen, size);
    }
    if (size > this.explicitChangedSeen.length) {
      this.explicitChangedSeen = growUint32(this.explicitChangedSeen, size);
    }
    if (size > this.explicitChangedBuffer.length) {
      this.explicitChangedBuffer = growUint32(this.explicitChangedBuffer, size);
    }
    if (size > this.impactedFormulaSeen.length) {
      this.impactedFormulaSeen = growUint32(this.impactedFormulaSeen, size);
    }
    if (size > this.impactedFormulaBuffer.length) {
      this.impactedFormulaBuffer = growUint32(this.impactedFormulaBuffer, size);
    }
  }

  private evaluateUnsupportedFormula(cellIndex: number): number[] {
    const formula = this.formulas.get(cellIndex);
    const sheetName = this.workbook.getSheetNameById(this.workbook.cellStore.sheetIds[cellIndex]!);
    if (!formula || !sheetName) {
      return [];
    }

    const evaluationContext = {
      sheetName,
      currentAddress: this.workbook.getAddress(cellIndex),
      resolveCell: (targetSheetName, address) => this.readCellValue(targetSheetName, address),
      resolveRange: (targetSheetName, start, end, refKind) =>
        this.readRangeValues(targetSheetName, start, end, refKind),
      resolveName: (name) => {
        const definedName = this.workbook.getDefinedName(name);
        if (!definedName) {
          return errorValue(ErrorCode.Name);
        }
        return definedNameValueToCellValue(definedName.value, this.strings);
      },
      resolveFormula: (targetSheetName: string, address: string) =>
        this.getCell(targetSheetName, address).formula,
      resolvePivotData: ({ dataField, sheetName: pivotSheetName, address, filters }) =>
        this.resolvePivotData(pivotSheetName, address, dataField, filters),
      resolveMultipleOperations: ({
        formulaSheetName,
        formulaAddress,
        rowCellSheetName,
        rowCellAddress,
        rowReplacementSheetName,
        rowReplacementAddress,
        columnCellSheetName,
        columnCellAddress,
        columnReplacementSheetName,
        columnReplacementAddress,
      }: {
        formulaSheetName: string;
        formulaAddress: string;
        rowCellSheetName: string;
        rowCellAddress: string;
        rowReplacementSheetName: string;
        rowReplacementAddress: string;
        columnCellSheetName?: string;
        columnCellAddress?: string;
        columnReplacementSheetName?: string;
        columnReplacementAddress?: string;
      }) =>
        this.resolveMultipleOperations({
          formulaSheetName,
          formulaAddress,
          rowCellSheetName,
          rowCellAddress,
          rowReplacementSheetName,
          rowReplacementAddress,
          ...(columnCellSheetName ? { columnCellSheetName } : {}),
          ...(columnCellAddress ? { columnCellAddress } : {}),
          ...(columnReplacementSheetName ? { columnReplacementSheetName } : {}),
          ...(columnReplacementAddress ? { columnReplacementAddress } : {}),
        }),
      listSheetNames: () =>
        [...this.workbook.sheetsByName.values()]
          .toSorted((left, right) => left.order - right.order)
          .map((sheet) => sheet.name),
    } as Parameters<typeof evaluatePlanResult>[1];
    const result = evaluatePlanResult(formula.compiled.jsPlan, evaluationContext);

    const materialization = isArrayValue(result)
      ? this.materializeSpill(cellIndex, result)
      : {
          changedCellIndices: this.clearOwnedSpill(cellIndex),
          ownerValue: result,
        };

    this.workbook.cellStore.flags[cellIndex] =
      (this.workbook.cellStore.flags[cellIndex] ?? 0) &
      ~(CellFlags.SpillChild | CellFlags.PivotOutput);
    this.workbook.cellStore.setValue(
      cellIndex,
      materialization.ownerValue,
      materialization.ownerValue.tag === ValueTag.String
        ? this.strings.intern(materialization.ownerValue.value)
        : 0,
    );
    return materialization.changedCellIndices;
  }

  private resolveStructuredReference(
    tableName: string,
    columnName: string,
  ): FormulaNode | undefined {
    const table = this.workbook.getTable(tableName);
    if (!table) {
      return undefined;
    }
    const columnIndex = table.columnNames.findIndex(
      (name) => name.trim().toUpperCase() === columnName.trim().toUpperCase(),
    );
    if (columnIndex === -1) {
      return undefined;
    }
    const start = parseCellAddress(table.startAddress, table.sheetName);
    const end = parseCellAddress(table.endAddress, table.sheetName);
    const startRow = start.row + (table.headerRow ? 1 : 0);
    const endRow = end.row - (table.totalsRow ? 1 : 0);
    if (endRow < startRow) {
      return { kind: "ErrorLiteral", code: ErrorCode.Ref };
    }
    const column = start.col + columnIndex;
    return {
      kind: "RangeRef",
      refKind: "cells",
      sheetName: table.sheetName,
      start: formatAddress(startRow, column),
      end: formatAddress(endRow, column),
    };
  }

  private resolveSpillReference(
    currentSheetName: string,
    sheetName: string | undefined,
    address: string,
  ): FormulaNode | undefined {
    const targetSheetName = sheetName ?? currentSheetName;
    const spill = this.workbook.getSpill(targetSheetName, address);
    if (!spill) {
      return undefined;
    }
    const owner = parseCellAddress(address, targetSheetName);
    return {
      kind: "RangeRef",
      refKind: "cells",
      sheetName: targetSheetName,
      start: owner.text,
      end: formatAddress(owner.row + spill.rows - 1, owner.col + spill.cols - 1),
    };
  }

  private readCellValue(sheetName: string, address: string): CellValue {
    const cellIndex = this.workbook.getCellIndex(sheetName, address);
    if (cellIndex === undefined) {
      return emptyValue();
    }
    return this.workbook.cellStore.getValue(cellIndex, (id) => this.strings.get(id));
  }

  private readRangeValueMatrix(range: CellRangeRef): CellValue[][] {
    const bounds = normalizeRange(range);
    const width = bounds.endCol - bounds.startCol + 1;
    const height = bounds.endRow - bounds.startRow + 1;
    const rows = Array.from<CellValue[]>({ length: height });
    const sheet = this.workbook.getSheet(range.sheetName);

    for (let rowOffset = 0; rowOffset < height; rowOffset += 1) {
      const row = bounds.startRow + rowOffset;
      const values = Array.from<CellValue>({ length: width });
      for (let colOffset = 0; colOffset < width; colOffset += 1) {
        const col = bounds.startCol + colOffset;
        const cellIndex = sheet?.grid.get(row, col) ?? -1;
        values[colOffset] =
          cellIndex === -1
            ? emptyValue()
            : this.workbook.cellStore.getValue(cellIndex, (id) => this.strings.get(id));
      }
      rows[rowOffset] = values;
    }

    return rows;
  }

  private readRangeValues(
    sheetName: string,
    start: string,
    end: string,
    refKind: "cells" | "rows" | "cols",
  ): CellValue[] {
    if (refKind !== "cells") {
      return [];
    }
    const range = parseRangeAddress(`${start}:${end}`, sheetName);
    if (range.kind !== "cells") {
      return [];
    }
    const rows = this.readRangeValueMatrix({
      sheetName,
      startAddress: start,
      endAddress: end,
    });
    const values = Array.from<CellValue>({ length: rows.length * (rows[0]?.length ?? 0) });
    let valueIndex = 0;
    for (let rowIndex = 0; rowIndex < rows.length; rowIndex += 1) {
      const row = rows[rowIndex]!;
      for (let colIndex = 0; colIndex < row.length; colIndex += 1) {
        values[valueIndex] = row[colIndex]!;
        valueIndex += 1;
      }
    }
    return values;
  }

  private scheduleWasmProgramSync(): void {
    runEngineEffect(this.runtime.graph.scheduleWasmProgramSync());
  }

  private flushWasmProgramSync(): void {
    runEngineEffect(this.runtime.graph.flushWasmProgramSync());
  }

  private estimatePotentialNewCells(ops: readonly EngineOp[]): number {
    let count = 0;
    for (let index = 0; index < ops.length; index += 1) {
      const op = ops[index]!;
      if (
        op.kind === "setCellValue" ||
        op.kind === "setCellFormula" ||
        op.kind === "setCellFormat"
      ) {
        count += 1;
      }
    }
    return count;
  }

  private getEntityDependents(entityId: number): Uint32Array {
    const slice = this.getReverseEdgeSlice(entityId) ?? this.edgeArena.empty();
    return this.edgeArena.readView(slice);
  }

  private setReverseEdgeSlice(entityId: number, slice: EdgeSlice): void {
    const empty = slice.ptr < 0 || slice.len === 0;
    if (isRangeEntity(entityId)) {
      this.reverseRangeEdges[entityPayload(entityId)] = empty ? undefined : slice;
      return;
    }
    this.reverseCellEdges[entityPayload(entityId)] = empty ? undefined : slice;
  }

  private appendReverseEdge(entityId: number, dependentEntityId: number): void {
    const slice = this.getReverseEdgeSlice(entityId) ?? this.edgeArena.empty();
    this.setReverseEdgeSlice(entityId, this.edgeArena.appendUnique(slice, dependentEntityId));
  }

  private getReverseEdgeSlice(entityId: number): EdgeSlice | undefined {
    if (isRangeEntity(entityId)) {
      return this.reverseRangeEdges[entityPayload(entityId)];
    }
    return this.reverseCellEdges[entityPayload(entityId)];
  }

  private collectFormulaDependentsForEntityInto(entityId: number): number {
    this.topoFormulaSeenEpoch += 1;
    this.topoRangeSeenEpoch += 1;

    let entityQueueLength = 1;
    let formulaCount = 0;
    this.topoEntityQueue[0] = entityId;

    for (let queueIndex = 0; queueIndex < entityQueueLength; queueIndex += 1) {
      const currentEntity = this.topoEntityQueue[queueIndex]!;
      const dependents = this.getEntityDependents(currentEntity);
      for (let index = 0; index < dependents.length; index += 1) {
        const dependent = dependents[index]!;
        if (isRangeEntity(dependent)) {
          const rangeIndex = entityPayload(dependent);
          if (this.topoRangeSeen[rangeIndex] === this.topoRangeSeenEpoch) {
            continue;
          }
          this.topoRangeSeen[rangeIndex] = this.topoRangeSeenEpoch;
          this.topoEntityQueue[entityQueueLength] = dependent;
          entityQueueLength += 1;
          continue;
        }

        const formulaCellIndex = entityPayload(dependent);
        if (this.topoFormulaSeen[formulaCellIndex] === this.topoFormulaSeenEpoch) {
          continue;
        }
        this.topoFormulaSeen[formulaCellIndex] = this.topoFormulaSeenEpoch;
        this.topoFormulaBuffer[formulaCount] = formulaCellIndex;
        formulaCount += 1;
      }
    }

    return formulaCount;
  }

  private forEachFormulaDependencyCell(
    cellIndex: number,
    fn: (dependencyCellIndex: number) => void,
  ): void {
    const formula = this.formulas.get(cellIndex);
    if (!formula) {
      return;
    }
    for (let index = 0; index < formula.dependencyIndices.length; index += 1) {
      fn(formula.dependencyIndices[index]!);
    }
  }

  private removeSheetRuntime(
    sheetName: string,
    explicitChangedCount: number,
  ): { changedInputCount: number; formulaChangedCount: number; explicitChangedCount: number } {
    const sheet = this.workbook.getSheet(sheetName);
    if (!sheet) {
      return { changedInputCount: 0, formulaChangedCount: 0, explicitChangedCount };
    }

    const cellIndices: number[] = [];
    sheet.grid.forEachCell((cellIndex) => {
      cellIndices.push(cellIndex);
    });
    const impactedCount = this.collectImpactedFormulasForCells(cellIndices);

    let changedInputCount = 0;
    let formulaChangedCount = 0;
    cellIndices.forEach((cellIndex) => {
      this.removeFormula(cellIndex);
      this.setReverseEdgeSlice(makeCellEntity(cellIndex), this.edgeArena.empty());
      this.workbook.cellStore.setValue(cellIndex, emptyValue());
      this.workbook.cellStore.flags[cellIndex] =
        (this.workbook.cellStore.flags[cellIndex] ?? 0) | CellFlags.PendingDelete;
      changedInputCount = this.markInputChanged(cellIndex, changedInputCount);
      explicitChangedCount = this.markExplicitChanged(cellIndex, explicitChangedCount);
    });

    this.workbook.deleteSheet(sheetName);
    if (this.selection.sheetName === sheetName) {
      const nextSheet = [...this.workbook.sheetsByName.values()].toSorted(
        (left, right) => left.order - right.order,
      )[0];
      this.setSelection(nextSheet?.name ?? sheetName, "A1");
    }
    formulaChangedCount = this.rebindFormulasForSheet(
      sheetName,
      formulaChangedCount,
      this.impactedFormulaBuffer.subarray(0, impactedCount),
    );
    return { changedInputCount, formulaChangedCount, explicitChangedCount };
  }

  private rebindFormulasForSheet(
    sheetName: string,
    formulaChangedCount: number,
    candidates?: readonly number[] | U32,
  ): number {
    return runEngineEffect(
      this.runtime.binding.rebindFormulasForSheet(sheetName, formulaChangedCount, candidates),
    );
  }

  private collectImpactedFormulasForCells(cellIndices: readonly number[]): number {
    this.ensureRecalcScratchCapacity(this.workbook.cellStore.size + 1);
    this.impactedFormulaEpoch += 1;
    if (this.impactedFormulaEpoch === 0xffff_ffff) {
      this.impactedFormulaEpoch = 1;
      this.impactedFormulaSeen.fill(0);
    }

    let impactedCount = 0;
    for (let cellCursor = 0; cellCursor < cellIndices.length; cellCursor += 1) {
      const cellIndex = cellIndices[cellCursor]!;
      const dependentCount = this.collectFormulaDependentsForEntityInto(makeCellEntity(cellIndex));
      for (let dependentIndex = 0; dependentIndex < dependentCount; dependentIndex += 1) {
        const formulaCellIndex = this.topoFormulaBuffer[dependentIndex]!;
        if (this.impactedFormulaSeen[formulaCellIndex] === this.impactedFormulaEpoch) {
          continue;
        }
        this.impactedFormulaSeen[formulaCellIndex] = this.impactedFormulaEpoch;
        this.impactedFormulaBuffer[impactedCount] = formulaCellIndex;
        impactedCount += 1;
      }
    }

    return impactedCount;
  }

  private beginMutationCollection(): void {
    this.changedInputEpoch += 1;
    if (this.changedInputEpoch === 0xffff_ffff) {
      this.changedInputEpoch = 1;
      this.changedInputSeen.fill(0);
    }
    this.changedFormulaEpoch += 1;
    if (this.changedFormulaEpoch === 0xffff_ffff) {
      this.changedFormulaEpoch = 1;
      this.changedFormulaSeen.fill(0);
    }
    this.explicitChangedEpoch += 1;
    if (this.explicitChangedEpoch === 0xffff_ffff) {
      this.explicitChangedEpoch = 1;
      this.explicitChangedSeen.fill(0);
    }
    this.ensureRecalcScratchCapacity(this.workbook.cellStore.size + 1);
  }

  private markInputChanged(cellIndex: number, count: number): number {
    if (this.changedInputSeen[cellIndex] === this.changedInputEpoch) {
      return count;
    }
    this.changedInputSeen[cellIndex] = this.changedInputEpoch;
    this.changedInputBuffer[count] = cellIndex;
    return count + 1;
  }

  private markFormulaChanged(cellIndex: number, count: number): number {
    if (this.changedFormulaSeen[cellIndex] === this.changedFormulaEpoch) {
      return count;
    }
    this.changedFormulaSeen[cellIndex] = this.changedFormulaEpoch;
    this.changedFormulaBuffer[count] = cellIndex;
    return count + 1;
  }

  private markVolatileFormulasChanged(count: number): number {
    this.formulas.forEach((formula, cellIndex) => {
      if (!formula.compiled.volatile) {
        return;
      }
      count = this.markFormulaChanged(cellIndex, count);
    });
    return count;
  }

  private markSpillRootsChanged(cellIndices: readonly number[], count: number): number {
    for (let index = 0; index < cellIndices.length; index += 1) {
      count = this.markInputChanged(cellIndices[index]!, count);
    }
    return count;
  }

  private markPivotRootsChanged(cellIndices: readonly number[], count: number): number {
    for (let index = 0; index < cellIndices.length; index += 1) {
      count = this.markInputChanged(cellIndices[index]!, count);
    }
    return count;
  }

  private markExplicitChanged(cellIndex: number, count: number): number {
    if (this.explicitChangedSeen[cellIndex] === this.explicitChangedEpoch) {
      return count;
    }
    this.explicitChangedSeen[cellIndex] = this.explicitChangedEpoch;
    this.explicitChangedBuffer[count] = cellIndex;
    return count + 1;
  }

  private composeMutationRoots(changedInputCount: number, formulaChangedCount: number): U32 {
    const total = changedInputCount + formulaChangedCount;
    this.ensureRecalcScratchCapacity(total + 1);
    for (let index = 0; index < changedInputCount; index += 1) {
      this.mutationRoots[index] = this.changedInputBuffer[index]!;
    }
    for (let index = 0; index < formulaChangedCount; index += 1) {
      this.mutationRoots[changedInputCount + index] = this.changedFormulaBuffer[index]!;
    }
    return this.mutationRoots.subarray(0, total);
  }

  private composeEventChanges(recalculated: U32, explicitChangedCount: number): U32 {
    this.changedUnionEpoch += 1;
    if (this.changedUnionEpoch === 0xffff_ffff) {
      this.changedUnionEpoch = 1;
      this.changedUnionSeen.fill(0);
    }
    let changedCount = 0;

    for (let index = 0; index < explicitChangedCount; index += 1) {
      const cellIndex = this.explicitChangedBuffer[index]!;
      if (this.changedUnionSeen[cellIndex] === this.changedUnionEpoch) {
        continue;
      }
      this.changedUnionSeen[cellIndex] = this.changedUnionEpoch;
      this.changedUnion[changedCount] = cellIndex;
      changedCount += 1;
    }

    for (let index = 0; index < recalculated.length; index += 1) {
      const cellIndex = recalculated[index]!;
      if (this.changedUnionSeen[cellIndex] === this.changedUnionEpoch) {
        continue;
      }
      this.changedUnionSeen[cellIndex] = this.changedUnionEpoch;
      this.changedUnion[changedCount] = cellIndex;
      changedCount += 1;
    }

    return this.changedUnion.subarray(0, changedCount);
  }

  private unionChangedSets(...sets: Array<readonly number[] | U32>): U32 {
    this.changedUnionEpoch += 1;
    if (this.changedUnionEpoch === 0xffff_ffff) {
      this.changedUnionEpoch = 1;
      this.changedUnionSeen.fill(0);
    }
    let changedCount = 0;
    for (let setIndex = 0; setIndex < sets.length; setIndex += 1) {
      const set = sets[setIndex]!;
      for (let index = 0; index < set.length; index += 1) {
        const cellIndex = set[index]!;
        if (this.changedUnionSeen[cellIndex] === this.changedUnionEpoch) {
          continue;
        }
        this.changedUnionSeen[cellIndex] = this.changedUnionEpoch;
        this.changedUnion[changedCount] = cellIndex;
        changedCount += 1;
      }
    }
    return this.changedUnion.subarray(0, changedCount);
  }

  private composeChangedRootsAndOrdered(
    changedRoots: readonly number[] | U32,
    ordered: U32,
    orderedCount: number,
  ): U32 {
    this.changedUnionEpoch += 1;
    if (this.changedUnionEpoch === 0xffff_ffff) {
      this.changedUnionEpoch = 1;
      this.changedUnionSeen.fill(0);
    }
    let changedCount = 0;

    for (let index = 0; index < changedRoots.length; index += 1) {
      const cellIndex = changedRoots[index]!;
      if (this.changedUnionSeen[cellIndex] === this.changedUnionEpoch) {
        continue;
      }
      this.changedUnionSeen[cellIndex] = this.changedUnionEpoch;
      this.changedUnion[changedCount] = cellIndex;
      changedCount += 1;
    }
    for (let index = 0; index < orderedCount; index += 1) {
      const cellIndex = ordered[index]!;
      if (this.changedUnionSeen[cellIndex] === this.changedUnionEpoch) {
        continue;
      }
      this.changedUnionSeen[cellIndex] = this.changedUnionEpoch;
      this.changedUnion[changedCount] = cellIndex;
      changedCount += 1;
    }

    return this.changedUnion.subarray(0, changedCount);
  }

  private ensureCellTracked(sheetName: string, address: string): number {
    const ensured = this.workbook.ensureCellRecord(sheetName, address);
    if (ensured.created) {
      this.pushMaterializedCell(ensured.cellIndex);
    }
    return ensured.cellIndex;
  }

  private ensureCellTrackedByCoords(sheetId: number, row: number, col: number): number {
    const ensured = this.workbook.ensureCellAt(sheetId, row, col);
    if (ensured.created) {
      this.pushMaterializedCell(ensured.cellIndex);
    }
    return ensured.cellIndex;
  }

  private clearOwnedSpill(cellIndex: number): number[] {
    const sheetName = this.workbook.getSheetNameById(this.workbook.cellStore.sheetIds[cellIndex]!);
    const address = this.workbook.getAddress(cellIndex);
    const spill = this.workbook.getSpill(sheetName, address);
    if (!spill) {
      return [];
    }

    const owner = parseCellAddress(address, sheetName);
    const changedCellIndices: number[] = [];
    for (let rowOffset = 0; rowOffset < spill.rows; rowOffset += 1) {
      for (let colOffset = 0; colOffset < spill.cols; colOffset += 1) {
        if (rowOffset === 0 && colOffset === 0) {
          continue;
        }
        const childAddress = formatAddress(owner.row + rowOffset, owner.col + colOffset);
        const childIndex = this.workbook.getCellIndex(sheetName, childAddress);
        if (childIndex === undefined) {
          continue;
        }
        if (this.clearSpillChildCell(childIndex)) {
          changedCellIndices.push(childIndex);
        }
      }
    }
    changedCellIndices.push(
      ...this.applyDerivedOp({ kind: "deleteSpillRange", sheetName, address }),
    );
    return changedCellIndices;
  }

  private materializeSpill(
    cellIndex: number,
    arrayValue: { values: CellValue[]; rows: number; cols: number },
  ): SpillMaterialization {
    const changedCellIndices = this.clearOwnedSpill(cellIndex);
    const sheetId = this.workbook.cellStore.sheetIds[cellIndex]!;
    const sheetName = this.workbook.getSheetNameById(sheetId);
    const address = this.workbook.getAddress(cellIndex);
    const owner = parseCellAddress(address, sheetName);

    if (owner.row + arrayValue.rows > MAX_ROWS || owner.col + arrayValue.cols > MAX_COLS) {
      return { changedCellIndices, ownerValue: errorValue(ErrorCode.Spill) };
    }

    for (let rowOffset = 0; rowOffset < arrayValue.rows; rowOffset += 1) {
      for (let colOffset = 0; colOffset < arrayValue.cols; colOffset += 1) {
        if (rowOffset === 0 && colOffset === 0) {
          continue;
        }
        const targetAddress = formatAddress(owner.row + rowOffset, owner.col + colOffset);
        const targetIndex = this.workbook.getCellIndex(sheetName, targetAddress);
        if (targetIndex === undefined) {
          continue;
        }
        const targetValue = this.workbook.cellStore.getValue(targetIndex, (id) =>
          this.strings.get(id),
        );
        if (this.formulas.get(targetIndex) || targetValue.tag !== ValueTag.Empty) {
          return { changedCellIndices, ownerValue: errorValue(ErrorCode.Blocked) };
        }
      }
    }

    for (let rowOffset = 0; rowOffset < arrayValue.rows; rowOffset += 1) {
      for (let colOffset = 0; colOffset < arrayValue.cols; colOffset += 1) {
        if (rowOffset === 0 && colOffset === 0) {
          continue;
        }
        const targetIndex = this.ensureCellTrackedByCoords(
          sheetId,
          owner.row + rowOffset,
          owner.col + colOffset,
        );
        const valueIndex = rowOffset * arrayValue.cols + colOffset;
        const value = arrayValue.values[valueIndex] ?? emptyValue();
        if (this.setSpillChildValue(targetIndex, value)) {
          changedCellIndices.push(targetIndex);
        }
      }
    }

    if (arrayValue.rows > 1 || arrayValue.cols > 1) {
      changedCellIndices.push(
        ...this.applyDerivedOp({
          kind: "upsertSpillRange",
          sheetName,
          address,
          rows: arrayValue.rows,
          cols: arrayValue.cols,
        }),
      );
    }

    return {
      changedCellIndices,
      ownerValue: arrayValue.values[0] ?? emptyValue(),
    };
  }

  private clearSpillChildCell(cellIndex: number): boolean {
    const currentFlags = this.workbook.cellStore.flags[cellIndex] ?? 0;
    const currentValue = this.workbook.cellStore.getValue(cellIndex, (id) => this.strings.get(id));
    if (currentValue.tag === ValueTag.Empty && (currentFlags & CellFlags.SpillChild) === 0) {
      return false;
    }
    this.workbook.cellStore.setValue(cellIndex, emptyValue());
    this.workbook.cellStore.flags[cellIndex] = currentFlags & ~CellFlags.SpillChild;
    return true;
  }

  private setSpillChildValue(cellIndex: number, value: CellValue): boolean {
    const currentValue = this.workbook.cellStore.getValue(cellIndex, (id) => this.strings.get(id));
    const currentFlags = this.workbook.cellStore.flags[cellIndex] ?? 0;
    const nextFlags =
      (currentFlags & ~(CellFlags.HasFormula | CellFlags.JsOnly | CellFlags.InCycle)) |
      CellFlags.SpillChild;
    if (areCellValuesEqual(currentValue, value) && currentFlags === nextFlags) {
      return false;
    }
    this.workbook.cellStore.setValue(
      cellIndex,
      value,
      value.tag === ValueTag.String ? this.strings.intern(value.value) : 0,
    );
    this.workbook.cellStore.flags[cellIndex] = nextFlags;
    return true;
  }

  private forEachSheetCell(
    sheetId: number,
    fn: (cellIndex: number, row: number, col: number) => void,
  ): void {
    const sheet = this.workbook.getSheetById(sheetId);
    if (!sheet) {
      return;
    }
    sheet.grid.forEachCellEntry((cellIndex, row, col) => {
      fn(cellIndex, row, col);
    });
  }

  private syncDynamicRanges(formulaChangedCount: number): number {
    for (let index = 0; index < this.materializedCellCount; index += 1) {
      const cellIndex = this.materializedCells[index]!;
      const sheetId = this.workbook.cellStore.sheetIds[cellIndex] ?? 0;
      if (sheetId === 0) {
        continue;
      }
      const row = this.workbook.cellStore.rows[cellIndex] ?? 0;
      const col = this.workbook.cellStore.cols[cellIndex] ?? 0;
      const rangeIndices = this.ranges.addDynamicMember(sheetId, row, col, cellIndex);
      if (rangeIndices.length > 0) {
        this.scheduleWasmProgramSync();
      }
      for (let rangeCursor = 0; rangeCursor < rangeIndices.length; rangeCursor += 1) {
        const rangeIndex = rangeIndices[rangeCursor]!;
        const rangeEntity = makeRangeEntity(rangeIndex);
        this.appendReverseEdge(makeCellEntity(cellIndex), rangeEntity);
        const formulas = this.getEntityDependents(rangeEntity);
        for (let formulaCursor = 0; formulaCursor < formulas.length; formulaCursor += 1) {
          const formulaEntity = formulas[formulaCursor]!;
          if (isRangeEntity(formulaEntity)) {
            continue;
          }
          const formulaCellIndex = entityPayload(formulaEntity);
          const formula = this.formulas.get(formulaCellIndex);
          if (!formula) {
            continue;
          }
          const nextDependencyIndices = appendPackedCellIndex(formula.dependencyIndices, cellIndex);
          if (nextDependencyIndices !== formula.dependencyIndices) {
            formula.dependencyIndices = nextDependencyIndices;
            formulaChangedCount = this.markFormulaChanged(formulaCellIndex, formulaChangedCount);
          }
        }
      }
    }
    return formulaChangedCount;
  }

  private resetWorkbook(workbookName = "Workbook"): void {
    const previousBatchId = this.lastMetrics.batchId;
    this.workbook.reset(workbookName);
    this.formulas.clear();
    this.reverseCellEdges.length = 0;
    this.reverseRangeEdges.length = 0;
    this.reverseDefinedNameEdges.clear();
    this.reverseTableEdges.clear();
    this.reverseSpillEdges.clear();
    this.pivotOutputOwners.clear();
    this.ranges.reset();
    this.edgeArena.reset();
    this.entityVersions.clear();
    this.sheetDeleteVersions.clear();
    this.undoStack.length = 0;
    this.redoStack.length = 0;
    this.selection = {
      sheetName: "Sheet1",
      address: "A1",
      anchorAddress: "A1",
      range: { startAddress: "A1", endAddress: "A1" },
      editMode: "idle",
    };
    this.syncState = "local-only";
    this.lastMetrics = {
      batchId: previousBatchId,
      changedInputCount: 0,
      dirtyFormulaCount: 0,
      wasmFormulaCount: 0,
      jsFormulaCount: 0,
      rangeNodeVisits: 0,
      recalcMs: 0,
      compileMs: 0,
    };
    this.wasmProgramSyncPending = false;
    this.materializedCellCount = 0;
    this.scheduleWasmProgramSync();
  }

  private resetMaterializedCellScratch(expectedSize: number): void {
    this.materializedCellCount = 0;
    if (expectedSize > this.materializedCells.length) {
      this.materializedCells = growUint32(this.materializedCells, expectedSize);
    }
  }

  private pushMaterializedCell(cellIndex: number): void {
    const nextCount = this.materializedCellCount + 1;
    if (nextCount > this.materializedCells.length) {
      this.materializedCells = growUint32(this.materializedCells, nextCount);
    }
    this.materializedCells[this.materializedCellCount] = cellIndex;
    this.materializedCellCount = nextCount;
  }
}

export const selectors = {
  selectCellSnapshot,
  selectMetrics,
  selectSelectionState,
  selectViewportCells,
};
