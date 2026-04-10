import { describe, expect, it, vi } from "vitest";
import { SpreadsheetEngine } from "@bilig/core";
import type { WorkbookLocalAuthoritativeDelta } from "@bilig/storage-browser";
import {
  buildPersistedWorkerState,
  ingestAuthoritativeDeltaToLocalStore,
  persistProjectionStateToLocalStore,
} from "../worker-runtime-local-persistence.js";

async function createEngine(workbookName: string, replicaId: string): Promise<SpreadsheetEngine> {
  const engine = new SpreadsheetEngine({
    workbookName,
    replicaId,
  });
  await engine.ready();
  engine.createSheet("Sheet1");
  return engine;
}

describe("worker runtime local persistence", () => {
  it("reuses projection exports when authoritative state is shared and no pending mutations remain", async () => {
    const authoritativeEngine = await createEngine("reuse-doc", "browser:authoritative");
    const projectionEngine = await createEngine("reuse-doc", "browser:projection");
    const cache = {
      getAuthoritativeSnapshot: vi.fn((input) =>
        input.canReuseProjectionState
          ? input.exportProjectionSnapshot()
          : input.exportAuthoritativeSnapshot(),
      ),
      getAuthoritativeReplica: vi.fn((input) =>
        input.canReuseProjectionState
          ? input.exportProjectionReplica()
          : input.exportAuthoritativeReplica(),
      ),
    };

    buildPersistedWorkerState({
      snapshotCaches: cache,
      authoritativeEngine,
      projectionEngine,
      hasDedicatedAuthoritativeEngine: false,
      authoritativeRevision: 7,
      appliedPendingLocalSeq: 0,
    });

    expect(cache.getAuthoritativeSnapshot).toHaveBeenCalledWith(
      expect.objectContaining({ canReuseProjectionState: true }),
    );
    expect(cache.getAuthoritativeReplica).toHaveBeenCalledWith(
      expect.objectContaining({ canReuseProjectionState: true }),
    );
  });

  it("forces authoritative exports when pending mutations remain", async () => {
    const authoritativeEngine = await createEngine("authoritative-doc", "browser:authoritative");
    const projectionEngine = await createEngine("authoritative-doc", "browser:projection");
    const cache = {
      getAuthoritativeSnapshot: vi.fn((input) =>
        input.canReuseProjectionState
          ? input.exportProjectionSnapshot()
          : input.exportAuthoritativeSnapshot(),
      ),
      getAuthoritativeReplica: vi.fn((input) =>
        input.canReuseProjectionState
          ? input.exportProjectionReplica()
          : input.exportAuthoritativeReplica(),
      ),
    };

    const state = buildPersistedWorkerState({
      snapshotCaches: cache,
      authoritativeEngine,
      projectionEngine,
      hasDedicatedAuthoritativeEngine: true,
      authoritativeRevision: 11,
      appliedPendingLocalSeq: 3,
    });

    expect(cache.getAuthoritativeSnapshot).toHaveBeenCalledWith(
      expect.objectContaining({ canReuseProjectionState: false }),
    );
    expect(cache.getAuthoritativeReplica).toHaveBeenCalledWith(
      expect.objectContaining({ canReuseProjectionState: false }),
    );
    expect(state.authoritativeRevision).toBe(11);
    expect(state.appliedPendingLocalSeq).toBe(3);
  });

  it("writes projection persistence and authoritative deltas through the local store", async () => {
    const authoritativeEngine = await createEngine("persist-doc", "browser:authoritative");
    const projectionEngine = await createEngine("persist-doc", "browser:projection");
    projectionEngine.setCellValue("Sheet1", "A1", 42);
    const state = {
      snapshot: projectionEngine.exportSnapshot(),
      replica: projectionEngine.exportReplicaSnapshot(),
      authoritativeRevision: 2,
      appliedPendingLocalSeq: 0,
    };
    const persistProjectionState = vi.fn(async () => {});
    const ingestAuthoritativeDelta = vi.fn(async () => {});
    const authoritativeDelta: WorkbookLocalAuthoritativeDelta = {
      replaceAll: false,
      replacedSheetIds: [],
      sheets: [],
    };

    await persistProjectionStateToLocalStore({
      localStore: { persistProjectionState },
      state,
      authoritativeEngine,
      projectionEngine,
      projectionOverlayScope: null,
    });
    await ingestAuthoritativeDeltaToLocalStore({
      localStore: { ingestAuthoritativeDelta },
      state,
      authoritativeDelta,
      authoritativeEngine,
      projectionEngine,
      projectionOverlayScope: null,
      removePendingMutationIds: ["mutation-1"],
    });

    expect(persistProjectionState).toHaveBeenCalledTimes(1);
    expect(ingestAuthoritativeDelta).toHaveBeenCalledTimes(1);
    expect(ingestAuthoritativeDelta).toHaveBeenCalledWith(
      expect.objectContaining({
        state,
        authoritativeDelta,
        removePendingMutationIds: ["mutation-1"],
      }),
    );
  });
});
