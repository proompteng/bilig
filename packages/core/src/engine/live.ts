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
  createEngineSelectionService,
  type EngineSelectionService,
} from "./services/selection-service.js";
import {
  createEngineSnapshotService,
  type EngineSnapshotService,
} from "./services/snapshot-service.js";

export interface EngineServiceRuntime {
  readonly events: EngineEventService;
  readonly selection: EngineSelectionService;
  readonly history: EngineHistoryService;
  readonly mutation: EngineMutationService;
  readonly pivot: EnginePivotService;
  readonly snapshot: EngineSnapshotService;
  readonly sync: EngineReplicaSyncService;
}

export function createEngineServiceRuntime(args: {
  readonly state: EngineRuntimeState;
  readonly getCellByIndex: (cellIndex: number) => CellSnapshot;
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
  return {
    events: createEngineEventService(args.state),
    selection: createEngineSelectionService(args.state),
    history: createEngineHistoryService({
      state: args.state,
      executeTransaction: args.executeHistoryTransaction,
    }),
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
