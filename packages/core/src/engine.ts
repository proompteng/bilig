import type {
  CellNumberFormatInput,
  CellNumberFormatRecord,
  CellRangeRef,
  CellStyleField,
  CellStylePatch,
  CellStyleRecord,
  CellSnapshot,
  CellValue,
  DependencySnapshot,
  EngineEvent,
  ExplainCellSnapshot,
  LiteralInput,
  RecalcMetrics,
  SelectionState,
  SyncState,
  WorkbookAxisEntrySnapshot,
  WorkbookCalculationSettingsSnapshot,
  WorkbookChartSnapshot,
  WorkbookCommentThreadSnapshot,
  WorkbookConditionalFormatSnapshot,
  WorkbookDataValidationSnapshot,
  WorkbookDefinedNameValueSnapshot,
  WorkbookFreezePaneSnapshot,
  WorkbookNoteSnapshot,
  WorkbookPivotSnapshot,
  WorkbookRangeProtectionSnapshot,
  WorkbookSheetProtectionSnapshot,
  WorkbookSortSnapshot,
  WorkbookSnapshot,
} from "@bilig/protocol";
import { Float64Arena, Uint32Arena, formatAddress, parseCellAddress } from "@bilig/formula";
import type { EngineOp, EngineOpBatch } from "@bilig/workbook-domain";
import type { EngineCellMutationRef } from "./cell-mutations-at.js";
import { createReplicaState, type OpOrder, type ReplicaState } from "./replica-state.js";
import { CycleDetector } from "./cycle-detection.js";
import { EdgeArena, type EdgeSlice } from "./edge-arena.js";
import { definedNameValuesEqual } from "./engine-metadata-utils.js";
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
  type WorkbookCommentThreadRecord,
  type WorkbookConditionalFormatRecord,
  type WorkbookDataValidationRecord,
  type WorkbookDefinedNameRecord,
  type WorkbookFilterRecord,
  type WorkbookPropertyRecord,
  type WorkbookRangeProtectionRecord,
  type WorkbookSheetProtectionRecord,
  type WorkbookSortRecord,
  type WorkbookSpillRecord,
  type WorkbookTableRecord,
  type WorkbookVolatileContextRecord,
  type WorkbookNoteRecord,
} from "./workbook-store.js";
import { cellToCsvValue, serializeCsv } from "./csv.js";
import { canonicalWorkbookRangeRef } from "./workbook-range-records.js";
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
  private dependencyBuildEpoch = 1;
  private dependencyBuildSeen: U32 = new Uint32Array(128);
  private dependencyBuildCells: U32 = new Uint32Array(128);
  private dependencyBuildEntities: U32 = new Uint32Array(128);
  private dependencyBuildRanges: U32 = new Uint32Array(128);
  private dependencyBuildNewRanges: U32 = new Uint32Array(128);
  private symbolicRefBindings: U32 = new Uint32Array(128);
  private symbolicRangeBindings: U32 = new Uint32Array(128);
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
      useColumnIndex: options.useColumnIndex ?? false,
      trackReplicaVersions: options.trackReplicaVersions ?? true,
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
      maintenance: {
        state: this.state,
        edgeArena: this.edgeArena,
        reverseState: {
          reverseCellEdges: this.reverseCellEdges,
          reverseRangeEdges: this.reverseRangeEdges,
          reverseDefinedNameEdges: this.reverseDefinedNameEdges,
          reverseTableEdges: this.reverseTableEdges,
          reverseSpillEdges: this.reverseSpillEdges,
        },
        pivotOutputOwners: this.pivotOutputOwners,
        setWasmProgramSyncPending: (next) => {
          this.wasmProgramSyncPending = next;
        },
        resetWasmState: () => {
          this.wasm.resetStoreState();
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
        },
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
        programArena: this.programArena,
        constantArena: this.constantArena,
        rangeListArena: this.rangeListArena,
        getTopoIndegree: () => this.topoIndegree,
        setTopoIndegree: (next) => {
          this.topoIndegree = next;
        },
        getTopoQueue: () => this.topoQueue,
        setTopoQueue: (next) => {
          this.topoQueue = next;
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
      traversal: {
        state: this.state,
        edgeArena: this.edgeArena,
        reverseState: {
          reverseCellEdges: this.reverseCellEdges,
          reverseRangeEdges: this.reverseRangeEdges,
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
        this.importSnapshot(snapshot);
      },
      operation: {
        state: this.state,
        reverseState: {
          reverseSpillEdges: this.reverseSpillEdges,
        },
        getBatchMutationDepth: () => this.batchMutationDepth,
        setBatchMutationDepth: (next) => {
          this.batchMutationDepth = next;
        },
        collectFormulaDependents: () => new Uint32Array(),
      },
    });
    if (!this.wasm.initSyncIfPossible()) {
      void this.wasm.init();
    }
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
    if (!this.state.trackReplicaVersions) {
      throw new Error(
        "Sync is unavailable when trackReplicaVersions is disabled; construct the engine with trackReplicaVersions enabled.",
      );
    }
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

  setCellValueAt(sheetId: number, row: number, col: number, value: LiteralInput): CellValue {
    const sheetName = this.workbook.getSheetById(sheetId)?.name;
    if (!sheetName) {
      throw new Error(`Unknown sheet id: ${sheetId}`);
    }
    const address = formatAddress(row, col);
    this.runtime.mutation.executeLocalCellMutationsAtNow(
      [{ sheetId, mutation: { kind: "setCellValue", row, col, value } }],
      1,
      { returnUndoOps: false },
    );
    return this.getCellValue(sheetName, address);
  }

  setCellFormula(sheetName: string, address: string, formula: string): CellValue {
    this.executeLocalTransaction([{ kind: "setCellFormula", sheetName, address, formula }]);
    return this.getCellValue(sheetName, address);
  }

  setCellFormulaAt(sheetId: number, row: number, col: number, formula: string): CellValue {
    const sheetName = this.workbook.getSheetById(sheetId)?.name;
    if (!sheetName) {
      throw new Error(`Unknown sheet id: ${sheetId}`);
    }
    const address = formatAddress(row, col);
    this.runtime.mutation.executeLocalCellMutationsAtNow(
      [{ sheetId, mutation: { kind: "setCellFormula", row, col, formula } }],
      1,
      { returnUndoOps: false },
    );
    return this.getCellValue(sheetName, address);
  }

  setCellFormat(sheetName: string, address: string, format: string | null): void {
    this.executeLocalTransaction([{ kind: "setCellFormat", sheetName, address, format }]);
  }

  clearCellAt(sheetId: number, row: number, col: number): void {
    this.runtime.mutation.executeLocalCellMutationsAtNow(
      [{ sheetId, mutation: { kind: "clearCell", row, col } }],
      0,
      { returnUndoOps: false },
    );
  }

  applyCellMutationsAt(
    refs: readonly EngineCellMutationRef[],
    potentialNewCells?: number,
  ): readonly EngineOp[] | null {
    return this.runtime.mutation.executeLocalCellMutationsAtNow(refs, potentialNewCells);
  }

  applyCellMutationsAtWithOptions(
    refs: readonly EngineCellMutationRef[],
    options: {
      captureUndo?: boolean;
      potentialNewCells?: number;
      source?: "local" | "restore";
      returnUndoOps?: boolean;
      reuseRefs?: boolean;
    } = {},
  ): readonly EngineOp[] | null {
    return this.runtime.mutation.applyCellMutationsAtNow(refs, options);
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

  setSheetProtection(protection: WorkbookSheetProtectionSnapshot): void {
    const existing = this.workbook.getSheetProtection(protection.sheetName);
    const normalized: WorkbookSheetProtectionSnapshot = {
      sheetName: protection.sheetName,
      ...(protection.hideFormulas !== undefined ? { hideFormulas: protection.hideFormulas } : {}),
    };
    if (existing && JSON.stringify(existing) === JSON.stringify(normalized)) {
      return;
    }
    this.executeLocalTransaction([{ kind: "setSheetProtection", protection: normalized }]);
  }

  clearSheetProtection(sheetName: string): boolean {
    if (!this.workbook.getSheetProtection(sheetName)) {
      return false;
    }
    this.executeLocalTransaction([{ kind: "clearSheetProtection", sheetName }]);
    return true;
  }

  getSheetProtection(sheetName: string): WorkbookSheetProtectionRecord | undefined {
    return this.workbook.getSheetProtection(sheetName);
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

  setDataValidation(validation: WorkbookDataValidationSnapshot): void {
    const existing = this.workbook.getDataValidation(validation.range.sheetName, validation.range);
    const normalized: WorkbookDataValidationSnapshot = {
      ...structuredClone(validation),
      range: canonicalWorkbookRangeRef(validation.range),
    };
    if (existing && JSON.stringify(existing) === JSON.stringify(normalized)) {
      return;
    }
    this.executeLocalTransaction([{ kind: "setDataValidation", validation: normalized }]);
  }

  clearDataValidation(sheetName: string, range: CellRangeRef): boolean {
    if (!this.workbook.getDataValidation(sheetName, range)) {
      return false;
    }
    this.executeLocalTransaction([{ kind: "clearDataValidation", sheetName, range: { ...range } }]);
    return true;
  }

  getDataValidation(
    sheetName: string,
    range: CellRangeRef,
  ): WorkbookDataValidationRecord | undefined {
    return this.workbook.getDataValidation(sheetName, range);
  }

  getDataValidations(sheetName: string): WorkbookDataValidationRecord[] {
    return this.workbook.listDataValidations(sheetName);
  }

  setConditionalFormat(format: WorkbookConditionalFormatSnapshot): void {
    const normalized: WorkbookConditionalFormatSnapshot = {
      ...structuredClone(format),
      id: format.id.trim(),
      range: canonicalWorkbookRangeRef(format.range),
    };
    const existing = this.workbook.getConditionalFormat(normalized.id);
    if (existing && JSON.stringify(existing) === JSON.stringify(normalized)) {
      return;
    }
    this.executeLocalTransaction([{ kind: "upsertConditionalFormat", format: normalized }]);
  }

  deleteConditionalFormat(id: string): boolean {
    const existing = this.workbook.getConditionalFormat(id);
    if (!existing) {
      return false;
    }
    this.executeLocalTransaction([
      {
        kind: "deleteConditionalFormat",
        id: existing.id,
        sheetName: existing.range.sheetName,
      },
    ]);
    return true;
  }

  getConditionalFormat(id: string): WorkbookConditionalFormatRecord | undefined {
    return this.workbook.getConditionalFormat(id);
  }

  getConditionalFormats(sheetName: string): WorkbookConditionalFormatRecord[] {
    return this.workbook.listConditionalFormats(sheetName);
  }

  setRangeProtection(protection: WorkbookRangeProtectionSnapshot): void {
    const normalized: WorkbookRangeProtectionSnapshot = {
      id: protection.id.trim(),
      range: canonicalWorkbookRangeRef(protection.range),
      ...(protection.hideFormulas !== undefined ? { hideFormulas: protection.hideFormulas } : {}),
    };
    const existing = this.workbook.getRangeProtection(normalized.id);
    if (existing && JSON.stringify(existing) === JSON.stringify(normalized)) {
      return;
    }
    this.executeLocalTransaction([{ kind: "upsertRangeProtection", protection: normalized }]);
  }

  deleteRangeProtection(id: string): boolean {
    const existing = this.workbook.getRangeProtection(id);
    if (!existing) {
      return false;
    }
    this.executeLocalTransaction([
      { kind: "deleteRangeProtection", id: existing.id, sheetName: existing.range.sheetName },
    ]);
    return true;
  }

  getRangeProtection(id: string): WorkbookRangeProtectionRecord | undefined {
    return this.workbook.getRangeProtection(id);
  }

  getRangeProtections(sheetName: string): WorkbookRangeProtectionRecord[] {
    return this.workbook.listRangeProtections(sheetName);
  }

  setCommentThread(thread: WorkbookCommentThreadSnapshot): void {
    const parsed = parseCellAddress(thread.address, thread.sheetName);
    const normalized: WorkbookCommentThreadSnapshot = {
      threadId: thread.threadId.trim(),
      sheetName: thread.sheetName,
      address: formatAddress(parsed.row, parsed.col),
      comments: thread.comments.map((comment) => ({
        id: comment.id.trim(),
        body: comment.body.trim(),
        ...(comment.authorUserId !== undefined ? { authorUserId: comment.authorUserId } : {}),
        ...(comment.authorDisplayName !== undefined
          ? { authorDisplayName: comment.authorDisplayName }
          : {}),
        ...(comment.createdAtUnixMs !== undefined
          ? { createdAtUnixMs: comment.createdAtUnixMs }
          : {}),
      })),
      ...(thread.resolved !== undefined ? { resolved: thread.resolved } : {}),
      ...(thread.resolvedByUserId !== undefined
        ? { resolvedByUserId: thread.resolvedByUserId }
        : {}),
      ...(thread.resolvedAtUnixMs !== undefined
        ? { resolvedAtUnixMs: thread.resolvedAtUnixMs }
        : {}),
    };
    const existing = this.workbook.getCommentThread(thread.sheetName, thread.address);
    if (existing && JSON.stringify(existing) === JSON.stringify(normalized)) {
      return;
    }
    this.executeLocalTransaction([{ kind: "upsertCommentThread", thread: normalized }]);
  }

  deleteCommentThread(sheetName: string, address: string): boolean {
    if (!this.workbook.getCommentThread(sheetName, address)) {
      return false;
    }
    this.executeLocalTransaction([{ kind: "deleteCommentThread", sheetName, address }]);
    return true;
  }

  getCommentThread(sheetName: string, address: string): WorkbookCommentThreadRecord | undefined {
    return this.workbook.getCommentThread(sheetName, address);
  }

  getCommentThreads(sheetName: string): WorkbookCommentThreadRecord[] {
    return this.workbook.listCommentThreads(sheetName);
  }

  setNote(note: WorkbookNoteSnapshot): void {
    const parsed = parseCellAddress(note.address, note.sheetName);
    const normalized: WorkbookNoteSnapshot = {
      sheetName: note.sheetName,
      address: formatAddress(parsed.row, parsed.col),
      text: note.text.trim(),
    };
    const existing = this.workbook.getNote(note.sheetName, note.address);
    if (existing && JSON.stringify(existing) === JSON.stringify(normalized)) {
      return;
    }
    this.executeLocalTransaction([{ kind: "upsertNote", note: normalized }]);
  }

  deleteNote(sheetName: string, address: string): boolean {
    if (!this.workbook.getNote(sheetName, address)) {
      return false;
    }
    this.executeLocalTransaction([{ kind: "deleteNote", sheetName, address }]);
    return true;
  }

  getNote(sheetName: string, address: string): WorkbookNoteRecord | undefined {
    return this.workbook.getNote(sheetName, address);
  }

  getNotes(sheetName: string): WorkbookNoteRecord[] {
    return this.workbook.listNotes(sheetName);
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

  setChart(chart: WorkbookChartSnapshot): void {
    const existing = this.workbook.getChart(chart.id);
    if (
      existing &&
      existing.sheetName === chart.sheetName &&
      existing.address === chart.address &&
      existing.chartType === chart.chartType &&
      existing.source.sheetName === chart.source.sheetName &&
      existing.source.startAddress === chart.source.startAddress &&
      existing.source.endAddress === chart.source.endAddress &&
      existing.rows === chart.rows &&
      existing.cols === chart.cols &&
      existing.seriesOrientation === chart.seriesOrientation &&
      existing.firstRowAsHeaders === chart.firstRowAsHeaders &&
      existing.firstColumnAsLabels === chart.firstColumnAsLabels &&
      existing.title === chart.title &&
      existing.legendPosition === chart.legendPosition
    ) {
      return;
    }
    this.executeLocalTransaction([{ kind: "upsertChart", chart: structuredClone(chart) }]);
  }

  deleteChart(id: string): boolean {
    if (!this.workbook.getChart(id)) {
      return false;
    }
    this.executeLocalTransaction([{ kind: "deleteChart", id }]);
    return true;
  }

  getChart(id: string): WorkbookChartSnapshot | undefined {
    return this.workbook.getChart(id);
  }

  getCharts(): WorkbookChartSnapshot[] {
    return this.workbook.listCharts();
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
      source?: "local" | "restore";
      trusted?: boolean;
    } = {},
  ): readonly EngineOp[] | null {
    return this.runtime.mutation.applyOpsNow(ops, options);
  }

  private executeLocalTransaction(
    ops: EngineOp[],
    potentialNewCells?: number,
  ): readonly EngineOp[] | null {
    return this.runtime.mutation.executeLocalNow(ops, potentialNewCells);
  }
}

export const selectors = {
  selectCellSnapshot,
  selectMetrics,
  selectSelectionState,
  selectViewportCells,
};
