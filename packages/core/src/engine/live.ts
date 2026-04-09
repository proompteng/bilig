import { Effect, Exit, Cause } from "effect";
import type { CellSnapshot } from "@bilig/protocol";
import type { EngineRuntimeState, TransactionRecord } from "./runtime-state.js";
import {
  createEngineEventService,
  type EngineEventService,
} from "./services/event-service.js";
import {
  createEngineHistoryService,
  type EngineHistoryService,
} from "./services/history-service.js";
import {
  createEngineMutationService,
  type EngineMutationService,
} from "./services/mutation-service.js";
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
  readonly history: EngineHistoryService;
  readonly mutation: EngineMutationService;
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
  readonly executeRestoreTransaction: (transaction: TransactionRecord) => void;
  readonly executeHistoryTransaction: (transaction: TransactionRecord, source: "history") => void;
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
  readonly applyBatchNow: (
    batch: import("@bilig/workbook-domain").EngineOpBatch,
    source: "local" | "restore" | "history",
    potentialNewCells?: number,
  ) => void;
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
  readonly getChangedInputBuffer: () => import("./runtime-state.js").U32;
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
  readonly removeFormula: (cellIndex: number) => boolean;
  readonly clearOwnedPivot: (
    pivot: import("../workbook-store.js").WorkbookPivotRecord,
  ) => number[];
  readonly materializePivot: (
    pivot: import("../workbook-store.js").WorkbookPivotRecord,
  ) => number[];
  readonly rebuildAllFormulaBindings: () => number[];
  readonly scheduleWasmProgramSync: () => void;
  readonly flushWasmProgramSync: () => void;
  readonly applyDerivedOp: (
    op: Extract<
      import("@bilig/workbook-domain").EngineOp,
      {
        kind:
          | "upsertSpillRange"
          | "deleteSpillRange"
          | "upsertPivotTable"
          | "deletePivotTable";
      }
    >,
  ) => number[];
  readonly applyRemoteBatchNow: (batch: import("@bilig/workbook-domain").EngineOpBatch) => void;
  readonly applyRemoteSnapshot: (snapshot: import("@bilig/protocol").WorkbookSnapshot) => void;
}): EngineServiceRuntime {
  const structure = createEngineStructureService({
    state: {
      workbook: args.state.workbook,
      formulas: args.state.formulas,
      pivotOutputOwners: args.pivotState.pivotOutputOwners,
    },
    getCellByIndex: args.getCellByIndex,
    toCellStateOps: args.toCellStateOps,
    removeFormula: (cellIndex) => args.removeFormula(cellIndex),
    clearOwnedPivot: (pivot) => args.clearOwnedPivot(pivot),
    rebuildAllFormulaBindings: () => args.rebuildAllFormulaBindings(),
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

  return {
    events: createEngineEventService(args.state),
    selection: createEngineSelectionService(args.state),
    history: createEngineHistoryService({
      state: args.state,
      executeTransaction: args.executeHistoryTransaction,
    }),
    read,
    recalc,
    structure,
    mutation: createEngineMutationService({
      state: args.state,
      captureSheetCellState: args.captureSheetCellState,
      captureRowRangeCellState: args.captureRowRangeCellState,
      captureColumnRangeCellState: args.captureColumnRangeCellState,
      restoreCellOps: args.restoreCellOps,
      applyBatchNow: args.applyBatchNow,
    }),
    pivot: createEnginePivotService({
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
      applyDerivedOp: args.applyDerivedOp,
    }),
    snapshot: createEngineSnapshotService({
      state: args.state,
      getCellByIndex: args.getCellByIndex,
      resetWorkbook: args.resetWorkbook,
      executeRestoreTransaction: args.executeRestoreTransaction,
    }),
    sync: createEngineReplicaSyncService({
      state: args.state,
      applyRemoteBatchNow: args.applyRemoteBatchNow,
      applyRemoteSnapshot: args.applyRemoteSnapshot,
    }),
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
