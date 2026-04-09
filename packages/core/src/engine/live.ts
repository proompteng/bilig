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
  readonly snapshot: EngineSnapshotService;
  readonly sync: EngineReplicaSyncService;
}

export function createEngineServiceRuntime(args: {
  readonly state: EngineRuntimeState;
  readonly getCellByIndex: (cellIndex: number) => CellSnapshot;
  readonly resetWorkbook: () => void;
  readonly executeRestoreTransaction: (transaction: TransactionRecord) => void;
  readonly executeHistoryTransaction: (transaction: TransactionRecord, source: "history") => void;
  readonly buildInverseOps: (ops: import("@bilig/workbook-domain").EngineOp[]) => import("@bilig/workbook-domain").EngineOp[];
  readonly applyBatchNow: (
    batch: import("@bilig/workbook-domain").EngineOpBatch,
    source: "local" | "restore" | "history",
    potentialNewCells?: number,
  ) => void;
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
      buildInverseOps: args.buildInverseOps,
      applyBatchNow: args.applyBatchNow,
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
