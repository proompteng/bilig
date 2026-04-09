import { Effect, Exit, Cause } from "effect";
import type { CellSnapshot } from "@bilig/protocol";
import type { EngineRuntimeState } from "./runtime-state.js";
import {
  createEngineEventService,
  type EngineEventService,
} from "./services/event-service.js";
import {
  createEngineFormulaBindingService,
  type EngineFormulaBindingService,
} from "./services/formula-binding-service.js";
import {
  createEngineFormulaGraphService,
  type EngineFormulaGraphService,
} from "./services/formula-graph-service.js";
import {
  createEngineHistoryService,
  type EngineHistoryService,
} from "./services/history-service.js";
import {
  createEngineMutationService,
  type EngineMutationService,
} from "./services/mutation-service.js";
import {
  createEngineMutationSupportService,
  type EngineMutationSupportService,
} from "./services/mutation-support-service.js";
import {
  createEngineOperationService,
  type EngineOperationService,
} from "./services/operation-service.js";
import {
  createEnginePivotService,
  type EnginePivotService,
} from "./services/pivot-service.js";
import {
  createEngineReplicaSyncService,
  type EngineReplicaSyncService,
} from "./services/replica-sync-service.js";
import {
  createEngineReadService,
  type EngineReadService,
} from "./services/read-service.js";
import {
  createEngineRecalcService,
  type EngineRecalcService,
} from "./services/recalc-service.js";
import {
  createEngineSelectionService,
  type EngineSelectionService,
} from "./services/selection-service.js";
import {
  createEngineSnapshotService,
  type EngineSnapshotService,
} from "./services/snapshot-service.js";
import {
  createEngineStructureService,
  type EngineStructureService,
} from "./services/structure-service.js";

export interface EngineServiceRuntime {
  readonly events: EngineEventService;
  readonly selection: EngineSelectionService;
  readonly binding: EngineFormulaBindingService;
  readonly graph: EngineFormulaGraphService;
  readonly history: EngineHistoryService;
  readonly mutation: EngineMutationService;
  readonly support: EngineMutationSupportService;
  readonly operations: EngineOperationService;
  readonly pivot: EnginePivotService;
  readonly read: EngineReadService;
  readonly recalc: EngineRecalcService;
  readonly structure: EngineStructureService;
  readonly snapshot: EngineSnapshotService;
  readonly sync: EngineReplicaSyncService;
}

export function createEngineServiceRuntime(args: {
  readonly state: EngineRuntimeState;
  readonly getCellByIndex: (cellIndex: number) => CellSnapshot;
  readonly exportSnapshot: () => import("@bilig/protocol").WorkbookSnapshot;
  readonly importSnapshot: (snapshot: import("@bilig/protocol").WorkbookSnapshot) => void;
  readonly resetWorkbook: () => void;
  readonly mutationSupport: Parameters<typeof createEngineMutationSupportService>[0];
  readonly captureSheetCellState: (
    sheetName: string,
  ) => import("@bilig/workbook-domain").EngineOp[];
  readonly captureRowRangeCellState: (
    sheetName: string,
    start: number,
    count: number,
  ) => import("@bilig/workbook-domain").EngineOp[];
  readonly captureColumnRangeCellState: (
    sheetName: string,
    start: number,
    count: number,
  ) => import("@bilig/workbook-domain").EngineOp[];
  readonly restoreCellOps: (
    sheetName: string,
    address: string,
  ) => import("@bilig/workbook-domain").EngineOp[];
  readonly formulaBinding: Parameters<typeof createEngineFormulaBindingService>[0];
  readonly formulaGraph: Parameters<typeof createEngineFormulaGraphService>[0];
  readonly readRangeCells: (
    range: import("@bilig/protocol").CellRangeRef,
  ) => CellSnapshot[][];
  readonly toCellStateOps: (
    sheetName: string,
    address: string,
    snapshot: CellSnapshot,
    sourceSheetName?: string,
    sourceAddress?: string,
  ) => import("@bilig/workbook-domain").EngineOp[];
  readonly forEachFormulaDependencyCell: (
    cellIndex: number,
    fn: (dependencyCellIndex: number) => void,
  ) => void;
  readonly cellToCsvValue: (cell: CellSnapshot) => string;
  readonly serializeCsv: (rows: string[][]) => string;
  readonly pivotState: {
    readonly pivotOutputOwners: Map<number, string>;
  };
  readonly ensureCellTrackedByCoords: (sheetId: number, row: number, col: number) => number;
  readonly forEachSheetCell: (
    sheetId: number,
    fn: (cellIndex: number, row: number, col: number) => void,
  ) => void;
  readonly beginMutationCollection: () => void;
  readonly markInputChanged: (cellIndex: number, count: number) => number;
  readonly markFormulaChanged: (cellIndex: number, count: number) => number;
  readonly markExplicitChanged: (cellIndex: number, count: number) => number;
  readonly composeMutationRoots: (changedInputCount: number, formulaChangedCount: number) => import("./runtime-state.js").U32;
  readonly composeEventChanges: (
    recalculated: import("./runtime-state.js").U32,
    explicitChangedCount: number,
  ) => import("./runtime-state.js").U32;
  readonly getChangedInputBuffer: () => import("./runtime-state.js").U32;
  readonly unionChangedSets: (
    ...sets: Array<readonly number[] | import("./runtime-state.js").U32>
  ) => import("./runtime-state.js").U32;
  readonly composeChangedRootsAndOrdered: (
    changedRoots: readonly number[] | import("./runtime-state.js").U32,
    ordered: import("./runtime-state.js").U32,
    orderedCount: number,
  ) => import("./runtime-state.js").U32;
  readonly emptyChangedSet: () => import("./runtime-state.js").U32;
  readonly ensureRecalcScratchCapacity: (size: number) => void;
  readonly getPendingKernelSync: () => import("./runtime-state.js").U32;
  readonly getWasmBatch: () => import("./runtime-state.js").U32;
  readonly materializeSpill: (
    cellIndex: number,
    arrayValue: {
      values: import("@bilig/protocol").CellValue[];
      rows: number;
      cols: number;
    },
  ) => import("./runtime-state.js").SpillMaterialization;
  readonly clearOwnedSpill: (cellIndex: number) => number[];
  readonly evaluateUnsupportedFormula: (cellIndex: number) => number[];
  readonly getEntityDependents: (entityId: number) => Uint32Array;
  readonly clearOwnedPivot: (
    pivot: import("../workbook-store.js").WorkbookPivotRecord,
  ) => number[];
  readonly clearPivotForCell: (cellIndex: number) => number[];
  readonly materializePivot: (
    pivot: import("../workbook-store.js").WorkbookPivotRecord,
  ) => number[];
  readonly scheduleWasmProgramSync: () => void;
  readonly flushWasmProgramSync: () => void;
  readonly applyRemoteSnapshot: (snapshot: import("@bilig/protocol").WorkbookSnapshot) => void;
  readonly operation: Parameters<typeof createEngineOperationService>[0];
}): EngineServiceRuntime {
  const support = createEngineMutationSupportService(args.mutationSupport);
  const binding = createEngineFormulaBindingService(args.formulaBinding);
  const graph = createEngineFormulaGraphService(args.formulaGraph);
  const structure = createEngineStructureService({
    state: {
      workbook: args.state.workbook,
      formulas: args.state.formulas,
      pivotOutputOwners: args.pivotState.pivotOutputOwners,
    },
    getCellByIndex: args.getCellByIndex,
    toCellStateOps: args.toCellStateOps,
    removeFormula: (cellIndex) => runEngineEffect(binding.clearFormula(cellIndex)),
    clearOwnedPivot: (pivot) => args.clearOwnedPivot(pivot),
    rebuildAllFormulaBindings: () => runEngineEffect(binding.rebuildAllFormulaBindings()),
  });
  const read = createEngineReadService({
    state: args.state,
    forEachFormulaDependencyCell: args.forEachFormulaDependencyCell,
    getEntityDependents: args.getEntityDependents,
    cellToCsvValue: args.cellToCsvValue,
    serializeCsv: args.serializeCsv,
  });
  const recalc = createEngineRecalcService({
    state: args.state,
    getCellByIndex: args.getCellByIndex,
    exportSnapshot: args.exportSnapshot,
    importSnapshot: args.importSnapshot,
    beginMutationCollection: args.beginMutationCollection,
    markInputChanged: args.markInputChanged,
    markFormulaChanged: args.markFormulaChanged,
    markExplicitChanged: args.markExplicitChanged,
    composeMutationRoots: args.composeMutationRoots,
    composeEventChanges: args.composeEventChanges,
    unionChangedSets: args.unionChangedSets,
    composeChangedRootsAndOrdered: args.composeChangedRootsAndOrdered,
    emptyChangedSet: args.emptyChangedSet,
    ensureRecalcScratchCapacity: args.ensureRecalcScratchCapacity,
    getPendingKernelSync: args.getPendingKernelSync,
    getWasmBatch: args.getWasmBatch,
    getChangedInputBuffer: args.getChangedInputBuffer,
    materializeSpill: args.materializeSpill,
    clearOwnedSpill: args.clearOwnedSpill,
    evaluateUnsupportedFormula: args.evaluateUnsupportedFormula,
    materializePivot: (pivot) => args.materializePivot(pivot),
    getEntityDependents: args.getEntityDependents,
  });
  const operations = createEngineOperationService(args.operation);
  const mutation = createEngineMutationService({
    state: args.state,
    captureSheetCellState: args.captureSheetCellState,
    captureRowRangeCellState: args.captureRowRangeCellState,
    captureColumnRangeCellState: args.captureColumnRangeCellState,
    restoreCellOps: args.restoreCellOps,
    readRangeCells: args.readRangeCells,
    toCellStateOps: args.toCellStateOps,
    applyBatchNow: (batch, source, potentialNewCells) =>
      runEngineEffect(operations.applyBatch(batch, source, potentialNewCells)),
  });
  const history = createEngineHistoryService({
    state: args.state,
    executeTransaction: (transaction, source) =>
      runEngineEffect(mutation.executeTransaction(transaction, source)),
  });
  const pivot = createEnginePivotService({
    state: {
      workbook: args.state.workbook,
      strings: args.state.strings,
      formulas: args.state.formulas,
      ranges: args.state.ranges,
      wasm: args.state.wasm,
      pivotOutputOwners: args.pivotState.pivotOutputOwners,
    },
    ensureCellTrackedByCoords: args.ensureCellTrackedByCoords,
    forEachSheetCell: args.forEachSheetCell,
    scheduleWasmProgramSync: args.scheduleWasmProgramSync,
    flushWasmProgramSync: args.flushWasmProgramSync,
    applyDerivedOp: (op) => runEngineEffect(operations.applyDerivedOp(op)),
  });
  const snapshot = createEngineSnapshotService({
    state: args.state,
    getCellByIndex: args.getCellByIndex,
    resetWorkbook: args.resetWorkbook,
    executeRestoreTransaction: (transaction) =>
      runEngineEffect(mutation.executeTransaction(transaction, "restore")),
  });
  const sync = createEngineReplicaSyncService({
    state: args.state,
    applyRemoteBatchNow: (batch) => runEngineEffect(operations.applyBatch(batch, "remote")),
    applyRemoteSnapshot: args.applyRemoteSnapshot,
  });

  return {
    events: createEngineEventService(args.state),
    selection: createEngineSelectionService(args.state),
    binding,
    graph,
    history,
    support,
    read,
    recalc,
    structure,
    mutation,
    operations,
    pivot,
    snapshot,
    sync,
  };
}

export function runEngineEffect<Success, Failure>(
  effect: Effect.Effect<Success, Failure>,
): Success {
  const exit = Effect.runSyncExit(effect);
  if (Exit.isSuccess(exit)) {
    return exit.value;
  }
  throw Cause.squash(exit.cause);
}

export async function runEngineEffectPromise<Success, Failure>(
  effect: Effect.Effect<Success, Failure>,
): Promise<Success> {
  const exit = await Effect.runPromiseExit(effect);
  if (Exit.isSuccess(exit)) {
    return exit.value;
  }
  throw Cause.squash(exit.cause);
}
