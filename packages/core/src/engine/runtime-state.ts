import type {
  CellValue,
  RecalcMetrics,
  SelectionState,
  SyncState,
  WorkbookPivotSnapshot,
} from "@bilig/protocol";
import { compileFormula } from "@bilig/formula";
import type { EngineOp, EngineOpBatch } from "@bilig/workbook-domain";
import type { OpOrder, ReplicaSnapshot, ReplicaState, ReplicaVersionSnapshot } from "../replica-state.js";
import type { EdgeSlice } from "../edge-arena.js";
import type { EngineEventBus } from "../events.js";
import type { FormulaTable } from "../formula-table.js";
import type { RangeRegistry } from "../range-registry.js";
import type { RecalcScheduler } from "../scheduler.js";
import type { StringPool } from "../string-pool.js";
import type { WasmKernelFacade } from "../wasm-facade.js";
import type { WorkbookStore } from "../workbook-store.js";

export interface CommitOp {
  kind:
    | "upsertWorkbook"
    | "upsertSheet"
    | "renameSheet"
    | "deleteSheet"
    | "upsertCell"
    | "deleteCell";
  name?: string;
  oldName?: string;
  newName?: string;
  order?: number;
  sheetName?: string;
  addr?: string;
  value?: import("@bilig/protocol").LiteralInput;
  formula?: string;
  format?: string;
}

export interface SpreadsheetEngineOptions {
  workbookName?: string;
  replicaId?: string;
}

export interface EngineSyncClientConnection {
  send(batch: EngineOpBatch): void | Promise<void>;
  disconnect(): void | Promise<void>;
}

export interface EngineSyncClient {
  connect(handlers: {
    applyRemoteBatch(batch: EngineOpBatch): boolean;
    applyRemoteSnapshot?(snapshot: import("@bilig/protocol").WorkbookSnapshot): void;
    setState(state: SyncState): void;
  }): EngineSyncClientConnection | Promise<EngineSyncClientConnection>;
}

export interface EngineReplicaSnapshot {
  replica: ReplicaSnapshot;
  entityVersions: ReplicaVersionSnapshot[];
  sheetDeleteVersions: Array<{ sheetName: string; order: OpOrder }>;
}

export interface TransactionRecord {
  ops: EngineOp[];
  potentialNewCells?: number;
}

export interface TransactionLogEntry {
  forward: TransactionRecord;
  inverse: TransactionRecord;
}

export interface RuntimeFormula {
  cellIndex: number;
  source: string;
  compiled: ReturnType<typeof compileFormula>;
  dependencyIndices: Uint32Array;
  dependencyEntities: EdgeSlice;
  rangeDependencies: Uint32Array;
  runtimeProgram: Uint32Array;
  constants: Float64Array;
  programOffset: number;
  programLength: number;
  constNumberOffset: number;
  constNumberLength: number;
  rangeListOffset: number;
  rangeListLength: number;
}

export type U32 = Uint32Array;

export const UNRESOLVED_WASM_OPERAND = 0x00ff_ffff;

export interface MaterializedDependencies {
  dependencyIndices: Uint32Array;
  dependencyEntities: Uint32Array;
  rangeDependencies: Uint32Array;
  symbolicRangeIndices: U32;
  symbolicRangeCount: number;
  newRangeIndices: U32;
  newRangeCount: number;
}

export interface SpillMaterialization {
  changedCellIndices: number[];
  ownerValue: CellValue;
}

export type PivotTableInput = Omit<
  WorkbookPivotSnapshot,
  "sheetName" | "address" | "rows" | "cols"
>;

export interface RecalcVolatileState {
  nowSerial: number;
  randomValues: number[];
  randomCursor: number;
}

export interface EngineRuntimeState {
  readonly workbook: WorkbookStore;
  readonly strings: StringPool;
  readonly events: EngineEventBus;
  readonly ranges: RangeRegistry;
  readonly scheduler: RecalcScheduler;
  readonly wasm: WasmKernelFacade;
  readonly formulas: FormulaTable<RuntimeFormula>;
  readonly replicaState: ReplicaState;
  readonly entityVersions: Map<string, OpOrder>;
  readonly sheetDeleteVersions: Map<string, OpOrder>;
  readonly batchListeners: Set<(batch: EngineOpBatch) => void>;
  readonly selectionListeners: Set<() => void>;
  readonly undoStack: TransactionLogEntry[];
  readonly redoStack: TransactionLogEntry[];
  getSelection(): SelectionState;
  setSelection(selection: SelectionState): void;
  getSyncState(): SyncState;
  setSyncState(state: SyncState): void;
  getSyncClientConnection(): EngineSyncClientConnection | null;
  setSyncClientConnection(connection: EngineSyncClientConnection | null): void;
  getTransactionReplayDepth(): number;
  setTransactionReplayDepth(depth: number): void;
  getLastMetrics(): RecalcMetrics;
  setLastMetrics(metrics: RecalcMetrics): void;
}

export interface EngineRuntimeStateController {
  readonly workbook: WorkbookStore;
  readonly strings: StringPool;
  readonly events: EngineEventBus;
  readonly ranges: RangeRegistry;
  readonly scheduler: RecalcScheduler;
  readonly wasm: WasmKernelFacade;
  readonly formulas: FormulaTable<RuntimeFormula>;
  readonly replicaState: ReplicaState;
  readonly entityVersions: Map<string, OpOrder>;
  readonly sheetDeleteVersions: Map<string, OpOrder>;
  readonly batchListeners: Set<(batch: EngineOpBatch) => void>;
  readonly selectionListeners: Set<() => void>;
  readonly undoStack: TransactionLogEntry[];
  readonly redoStack: TransactionLogEntry[];
  getSelection(): SelectionState;
  setSelection(selection: SelectionState): void;
  getSyncState(): SyncState;
  setSyncState(state: SyncState): void;
  getSyncClientConnection(): EngineSyncClientConnection | null;
  setSyncClientConnection(connection: EngineSyncClientConnection | null): void;
  getTransactionReplayDepth(): number;
  setTransactionReplayDepth(depth: number): void;
  getLastMetrics(): RecalcMetrics;
  setLastMetrics(metrics: RecalcMetrics): void;
}

export function createEngineRuntimeState(
  controller: EngineRuntimeStateController,
): EngineRuntimeState {
  return controller;
}

export function createInitialSelectionState(): SelectionState {
  return {
    sheetName: "Sheet1",
    address: "A1",
    anchorAddress: "A1",
    range: { startAddress: "A1", endAddress: "A1" },
    editMode: "idle",
  };
}

export function createInitialRecalcMetrics(): RecalcMetrics {
  return {
    batchId: 0,
    changedInputCount: 0,
    dirtyFormulaCount: 0,
    wasmFormulaCount: 0,
    jsFormulaCount: 0,
    rangeNodeVisits: 0,
    recalcMs: 0,
    compileMs: 0,
  };
}
