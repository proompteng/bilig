import { Effect } from "effect";
import { describe, expect, it, vi } from "vitest";
import { createBatch, createReplicaState, markBatchApplied } from "../replica-state.js";
import { createEngineReplicaSyncService } from "../engine/services/replica-sync-service.js";
import type { EngineSyncClientConnection } from "../engine/runtime-state.js";

describe("EngineReplicaSyncService", () => {
  it("deduplicates remote batches before invoking the mutation path", () => {
    let syncState = "local-only" as const;
    let connection: EngineSyncClientConnection | null = null;
    const replicaState = createReplicaState("local");
    const applied = vi.fn();
    const service = createEngineReplicaSyncService({
      state: {
        replicaState,
        getSyncState: () => syncState,
        setSyncState: (next) => {
          syncState = next;
        },
        getSyncClientConnection: () => connection,
        setSyncClientConnection: (next) => {
          connection = next;
        },
      },
      applyRemoteBatchNow: (batch) => {
        markBatchApplied(replicaState, batch);
        applied(batch);
      },
      applyRemoteSnapshot: () => {},
    });
    const remote = createReplicaState("remote");
    const batch = createBatch(remote, [{ kind: "upsertWorkbook", name: "book" }]);

    expect(Effect.runSync(service.applyRemoteBatch(batch))).toBe(true);
    expect(Effect.runSync(service.applyRemoteBatch(batch))).toBe(false);
    expect(applied).toHaveBeenCalledTimes(1);
  });
});
