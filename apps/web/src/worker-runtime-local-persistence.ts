import type { EngineReplicaSnapshot, SpreadsheetEngine } from "@bilig/core";
import type {
  WorkbookLocalAuthoritativeDelta,
  WorkbookLocalStore,
  WorkbookStoredState,
} from "@bilig/storage-browser";
import type { WorkbookSnapshot } from "@bilig/protocol";
import { buildWorkbookLocalAuthoritativeBase } from "./worker-local-base.js";
import {
  buildWorkbookLocalProjectionOverlay,
  type ProjectionOverlayScope,
} from "./worker-local-overlay.js";

interface AuthoritativeStateSnapshotCache {
  getAuthoritativeSnapshot(input: {
    canReuseProjectionState: boolean;
    exportProjectionSnapshot: () => WorkbookSnapshot;
    exportAuthoritativeSnapshot: () => WorkbookSnapshot;
  }): WorkbookSnapshot;
  getAuthoritativeReplica(input: {
    canReuseProjectionState: boolean;
    exportProjectionReplica: () => EngineReplicaSnapshot;
    exportAuthoritativeReplica: () => EngineReplicaSnapshot;
  }): EngineReplicaSnapshot;
}

export function buildPersistedWorkerState(args: {
  snapshotCaches: AuthoritativeStateSnapshotCache;
  authoritativeEngine: SpreadsheetEngine;
  projectionEngine: SpreadsheetEngine;
  hasDedicatedAuthoritativeEngine: boolean;
  authoritativeRevision: number;
  appliedPendingLocalSeq: number;
}): WorkbookStoredState {
  return {
    snapshot: args.snapshotCaches.getAuthoritativeSnapshot({
      canReuseProjectionState:
        !args.hasDedicatedAuthoritativeEngine && args.appliedPendingLocalSeq === 0,
      exportProjectionSnapshot: () => args.projectionEngine.exportSnapshot(),
      exportAuthoritativeSnapshot: () => args.authoritativeEngine.exportSnapshot(),
    }),
    replica: args.snapshotCaches.getAuthoritativeReplica({
      canReuseProjectionState:
        !args.hasDedicatedAuthoritativeEngine && args.appliedPendingLocalSeq === 0,
      exportProjectionReplica: () => args.projectionEngine.exportReplicaSnapshot(),
      exportAuthoritativeReplica: () => args.authoritativeEngine.exportReplicaSnapshot(),
    }),
    authoritativeRevision: args.authoritativeRevision,
    appliedPendingLocalSeq: args.appliedPendingLocalSeq,
  };
}

export async function persistProjectionStateToLocalStore(args: {
  localStore: Pick<WorkbookLocalStore, "persistProjectionState">;
  state: WorkbookStoredState;
  authoritativeEngine: SpreadsheetEngine;
  projectionEngine: SpreadsheetEngine;
  projectionOverlayScope: ProjectionOverlayScope | null;
}): Promise<void> {
  await args.localStore.persistProjectionState({
    state: args.state,
    authoritativeBase: buildWorkbookLocalAuthoritativeBase(args.authoritativeEngine),
    projectionOverlay: buildWorkbookLocalProjectionOverlay({
      authoritativeEngine: args.authoritativeEngine,
      projectionEngine: args.projectionEngine,
      scope: args.projectionOverlayScope,
    }),
  });
}

export async function ingestAuthoritativeDeltaToLocalStore(args: {
  localStore: Pick<WorkbookLocalStore, "ingestAuthoritativeDelta">;
  state: WorkbookStoredState;
  authoritativeDelta: WorkbookLocalAuthoritativeDelta;
  authoritativeEngine: SpreadsheetEngine;
  projectionEngine: SpreadsheetEngine;
  projectionOverlayScope: ProjectionOverlayScope | null;
  removePendingMutationIds: readonly string[];
}): Promise<void> {
  await args.localStore.ingestAuthoritativeDelta({
    state: args.state,
    authoritativeDelta: args.authoritativeDelta,
    projectionOverlay: buildWorkbookLocalProjectionOverlay({
      authoritativeEngine: args.authoritativeEngine,
      projectionEngine: args.projectionEngine,
      scope: args.projectionOverlayScope,
    }),
    removePendingMutationIds: args.removePendingMutationIds,
  });
}
