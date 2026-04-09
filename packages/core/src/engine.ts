import {
  type CellNumberFormatInput,
  type CellNumberFormatRecord,
  type CellRangeRef,
  type CellStyleField,
  type CellStylePatch,
  type CellStyleRecord,
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
  formatAddress,
  parseCellAddress,
  translateFormulaReferences,
} from "@bilig/formula";
import { Float64Arena, Uint32Arena } from "@bilig/formula/program-arena";
import type { EngineOp, EngineOpBatch } from "@bilig/workbook-domain";
import {
  createReplicaState,
  type OpOrder,
  type ReplicaState,
} from "./replica-state.js";
import { CycleDetector } from "./cycle-detection.js";
import { EdgeArena, type EdgeSlice } from "./edge-arena.js";
import { entityPayload, isRangeEntity } from "./entity-ids.js";
import { growUint32 } from "./engine-buffer-utils.js";
import {
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
      mutationSupport: {
        state: this.state,
        edgeArena: this.edgeArena,
        reverseState: {
          reverseCellEdges: this.reverseCellEdges,
          reverseRangeEdges: this.reverseRangeEdges,
        },
        getSelectionState: () => this.getSelectionState(),
        setSelection: (sheetName, address) => this.setSelection(sheetName, address),
        ensureRecalcScratchCapacity: (size) => this.ensureRecalcScratchCapacity(size),
        collectFormulaDependentsForEntityInto: (entityId) =>
          this.collectFormulaDependentsForEntityInto(entityId),
        getTopoFormulaBuffer: () => this.topoFormulaBuffer,
        getChangedInputEpoch: () => this.changedInputEpoch,
        setChangedInputEpoch: (next) => {
          this.changedInputEpoch = next;
        },
        getChangedInputSeen: () => this.changedInputSeen,
        setChangedInputSeen: (next) => {
          this.changedInputSeen = next;
        },
        getChangedInputBuffer: () => this.changedInputBuffer,
        setChangedInputBuffer: (next) => {
          this.changedInputBuffer = next;
        },
        getChangedFormulaEpoch: () => this.changedFormulaEpoch,
        setChangedFormulaEpoch: (next) => {
          this.changedFormulaEpoch = next;
        },
        getChangedFormulaSeen: () => this.changedFormulaSeen,
        setChangedFormulaSeen: (next) => {
          this.changedFormulaSeen = next;
        },
        getChangedFormulaBuffer: () => this.changedFormulaBuffer,
        setChangedFormulaBuffer: (next) => {
          this.changedFormulaBuffer = next;
        },
        getChangedUnionEpoch: () => this.changedUnionEpoch,
        setChangedUnionEpoch: (next) => {
          this.changedUnionEpoch = next;
        },
        getChangedUnionSeen: () => this.changedUnionSeen,
        setChangedUnionSeen: (next) => {
          this.changedUnionSeen = next;
        },
        getChangedUnion: () => this.changedUnion,
        setChangedUnion: (next) => {
          this.changedUnion = next;
        },
        getMutationRoots: () => this.mutationRoots,
        setMutationRoots: (next) => {
          this.mutationRoots = next;
        },
        getMaterializedCellCount: () => this.materializedCellCount,
        setMaterializedCellCount: (next) => {
          this.materializedCellCount = next;
        },
        getMaterializedCells: () => this.materializedCells,
        setMaterializedCells: (next) => {
          this.materializedCells = next;
        },
        getExplicitChangedEpoch: () => this.explicitChangedEpoch,
        setExplicitChangedEpoch: (next) => {
          this.explicitChangedEpoch = next;
        },
        getExplicitChangedSeen: () => this.explicitChangedSeen,
        setExplicitChangedSeen: (next) => {
          this.explicitChangedSeen = next;
        },
        getExplicitChangedBuffer: () => this.explicitChangedBuffer,
        setExplicitChangedBuffer: (next) => {
          this.explicitChangedBuffer = next;
        },
        getImpactedFormulaEpoch: () => this.impactedFormulaEpoch,
        setImpactedFormulaEpoch: (next) => {
          this.impactedFormulaEpoch = next;
        },
        getImpactedFormulaSeen: () => this.impactedFormulaSeen,
        setImpactedFormulaSeen: (next) => {
          this.impactedFormulaSeen = next;
        },
        getImpactedFormulaBuffer: () => this.impactedFormulaBuffer,
        setImpactedFormulaBuffer: (next) => {
          this.impactedFormulaBuffer = next;
        },
      },
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
        forEachSheetCell: (sheetId, fn) => this.forEachSheetCell(sheetId, fn),
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
      forEachSheetCell: (sheetId, fn) => this.forEachSheetCell(sheetId, fn),
      recalc: {
        state: this.state,
        getCellByIndex: (cellIndex) => this.getCellByIndex(cellIndex),
        exportSnapshot: () => this.exportSnapshot(),
        importSnapshot: (snapshot) => this.importSnapshot(snapshot),
        ensureRecalcScratchCapacity: (size) => this.ensureRecalcScratchCapacity(size),
        getPendingKernelSync: () => this.pendingKernelSync,
        getWasmBatch: () => this.wasmBatch,
        getEntityDependents: (entityId) => this.getEntityDependents(entityId),
        now: () => new Date(),
        random: () => Math.random(),
        performanceNow: () => performance.now(),
      },
      getEntityDependents: (entityId) => this.getEntityDependents(entityId),
      pivot: {
        state: {
          workbook: this.state.workbook,
          strings: this.state.strings,
          formulas: this.state.formulas,
          ranges: this.state.ranges,
          wasm: this.state.wasm,
          pivotOutputOwners: this.pivotOutputOwners,
        },
        forEachSheetCell: (sheetId, fn) => this.forEachSheetCell(sheetId, fn),
      },
      applyRemoteSnapshot: (snapshot) => {
        this.importSnapshot(snapshot);
      },
      operation: {
        state: this.state,
        reverseState: {
          reverseSpillEdges: this.reverseSpillEdges,
        },
        rewriteDefinedNamesForSheetRename: (oldSheetName, newSheetName) =>
          this.rewriteDefinedNamesForSheetRename(oldSheetName, newSheetName),
        estimatePotentialNewCells: (ops) => this.estimatePotentialNewCells(ops),
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

  private rewriteDefinedNamesForSheetRename(oldSheetName: string, newSheetName: string): void {
    this.workbook.listDefinedNames().forEach((record) => {
      const nextValue = renameDefinedNameValueSheet(record.value, oldSheetName, newSheetName);
      if (!definedNameValuesEqual(record.value, nextValue)) {
        this.workbook.setDefinedName(record.name, nextValue);
      }
    });
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
    runEngineEffect(this.runtime.graph.scheduleWasmProgramSync());
  }
}

export const selectors = {
  selectCellSnapshot,
  selectMetrics,
  selectSelectionState,
  selectViewportCells,
};
