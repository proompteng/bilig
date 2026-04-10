import { describe, expect, it } from "vitest";
import { SpreadsheetEngine } from "@bilig/core";
import { WorkerRuntimeSnapshotCaches } from "../worker-runtime-snapshot-caches.js";
import {
  ensureAuthoritativeEngine,
  installRestoredAuthoritativeState,
  readProjectedCellFromLocalStore,
  resolveAuthoritativeStateInput,
} from "../worker-runtime-engine-access.js";
import { createEmptyCellSnapshot } from "../worker-runtime-viewport-publisher.js";

async function createSnapshotPair() {
  const engine = new SpreadsheetEngine({
    workbookName: "worker-runtime-engine-access",
    replicaId: "replica-1",
  });
  await engine.ready();
  engine.createSheet("Sheet1");
  return {
    snapshot: engine.exportSnapshot(),
    replica: engine.exportReplicaSnapshot(),
  };
}

describe("worker runtime engine access", () => {
  it("restores authoritative state from local store snapshots", async () => {
    const caches = new WorkerRuntimeSnapshotCaches();
    const { snapshot, replica } = await createSnapshotPair();
    const restored: Array<{ snapshot: unknown; replica: unknown }> = [];

    const result = await resolveAuthoritativeStateInput({
      authoritativeStateSource: "localStore",
      localStore: {
        async loadState() {
          return {
            snapshot,
            replica,
            authoritativeRevision: 0,
            appliedPendingLocalSeq: 0,
          };
        },
      },
      snapshotCaches: caches,
      authoritativeEngine: null,
      installRestoredAuthoritativeState(restoredSnapshot, restoredReplica) {
        restored.push({ snapshot: restoredSnapshot, replica: restoredReplica });
        installRestoredAuthoritativeState(caches, restoredSnapshot, restoredReplica);
      },
    });

    expect(restored).toEqual([{ snapshot, replica }]);
    expect(result).toEqual({ snapshot, replica });
  });

  it("seeds authoritative caches when lazily creating an engine", async () => {
    const caches = new WorkerRuntimeSnapshotCaches();

    const engine = await ensureAuthoritativeEngine({
      authoritativeEngine: null,
      documentId: "worker-runtime-engine-access-doc",
      replicaId: "replica-2",
      snapshotCaches: caches,
      async resolveAuthoritativeStateInput() {
        return { snapshot: null, replica: null };
      },
    });

    const resolved = caches.resolveAuthoritativeState({
      exportSnapshot: null,
      exportReplica: null,
    });
    expect(resolved.snapshot).toEqual(engine.exportSnapshot());
    expect(resolved.replica).toEqual(engine.exportReplicaSnapshot());
  });

  it("reads and clones projected cells from the local viewport tile store", () => {
    const snapshot = {
      ...createEmptyCellSnapshot("Sheet1", "A1"),
      value: { tag: 2, value: 42 },
    };

    const result = readProjectedCellFromLocalStore({
      canReadLocalProjectionForViewport: true,
      localStore: {},
      viewportTileStore: {
        readViewport() {
          return { cells: [{ snapshot }] };
        },
      },
      sheetName: "Sheet1",
      address: "A1",
    });

    expect(result).toEqual(snapshot);
    expect(result).not.toBe(snapshot);
  });
});
