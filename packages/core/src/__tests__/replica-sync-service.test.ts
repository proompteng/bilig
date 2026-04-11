import { Cause, Effect, Exit } from "effect";
import { describe, expect, it, vi } from "vitest";
import type { WorkbookSnapshot } from "@bilig/protocol";
import { createBatch, createReplicaState, markBatchApplied } from "../replica-state.js";
import { createEngineReplicaSyncService } from "../engine/services/replica-sync-service.js";
import type { EngineSyncClient, EngineSyncClientConnection } from "../engine/runtime-state.js";
import { EngineSyncError } from "../engine/errors.js";

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

  it("connects clients, tears down the previous connection, and forwards sync callbacks", async () => {
    let syncState: "local-only" | "syncing" | "live" | "offline" = "local-only";
    const existing = {
      disconnect: vi.fn(async () => {}),
    } satisfies EngineSyncClientConnection;
    let connection: EngineSyncClientConnection | null = existing;
    const replicaState = createReplicaState("local");
    const applied = vi.fn((batch) => {
      markBatchApplied(replicaState, batch);
    });
    const appliedSnapshots: WorkbookSnapshot[] = [];
    const nextConnection = {
      disconnect: vi.fn(async () => {}),
    } satisfies EngineSyncClientConnection;
    let handlers: Parameters<NonNullable<EngineSyncClient["connect"]>>[0] | undefined;
    const client: EngineSyncClient = {
      connect: vi.fn(async (nextHandlers) => {
        handlers = nextHandlers;
        return nextConnection;
      }),
    };
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
      applyRemoteBatchNow: applied,
      applyRemoteSnapshot: (snapshot) => {
        appliedSnapshots.push(snapshot);
      },
    });

    await Effect.runPromise(service.connectClient(client));

    expect(existing.disconnect).toHaveBeenCalledTimes(1);
    expect(syncState).toBe("live");
    expect(connection).toBe(nextConnection);
    expect(handlers).toBeDefined();

    const remote = createReplicaState("remote");
    const batch = createBatch(remote, [{ kind: "upsertWorkbook", name: "remote-book" }]);
    expect(handlers!.applyRemoteBatch(batch)).toBe(true);
    expect(applied).toHaveBeenCalledWith(batch);

    const snapshot: WorkbookSnapshot = {
      version: 1,
      workbook: { name: "remote-book" },
      sheets: [],
    };
    handlers!.applyRemoteSnapshot(snapshot);
    expect(appliedSnapshots).toEqual([snapshot]);
  });

  it("wraps connect, disconnect, and apply failures in EngineSyncError", async () => {
    let syncState: "local-only" | "syncing" | "live" = "local-only";
    let connection: EngineSyncClientConnection | null = {
      disconnect: async () => {
        throw new Error("disconnect failed");
      },
    };
    const replicaState = createReplicaState("local");
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
      applyRemoteBatchNow: () => {
        throw new Error("apply failed");
      },
      applyRemoteSnapshot: () => {},
    });

    const connectExit = await Effect.runPromiseExit(
      service.connectClient({
        connect: async () => {
          throw new Error("connect failed");
        },
      }),
    );
    expect(Exit.isFailure(connectExit)).toBe(true);
    if (!Exit.isFailure(connectExit)) {
      throw new Error("Expected failed connect exit");
    }
    expect(Cause.squash(connectExit.cause)).toMatchObject(
      new EngineSyncError({
        message: "Failed to connect sync client",
        cause: expect.any(Error),
      }),
    );

    connection = {
      disconnect: async () => {
        throw new Error("disconnect failed");
      },
    };
    const disconnectExit = await Effect.runPromiseExit(service.disconnectClient());
    expect(Exit.isFailure(disconnectExit)).toBe(true);
    if (!Exit.isFailure(disconnectExit)) {
      throw new Error("Expected failed disconnect exit");
    }
    expect(Cause.squash(disconnectExit.cause)).toMatchObject(
      new EngineSyncError({
        message: "Failed to disconnect sync client",
        cause: expect.any(Error),
      }),
    );

    const remote = createReplicaState("remote");
    const batch = createBatch(remote, [{ kind: "upsertWorkbook", name: "remote-book" }]);
    const applyExit = Effect.runSyncExit(service.applyRemoteBatch(batch));
    expect(Exit.isFailure(applyExit)).toBe(true);
    if (!Exit.isFailure(applyExit)) {
      throw new Error("Expected failed apply exit");
    }
    expect(Cause.squash(applyExit.cause)).toMatchObject(
      new EngineSyncError({
        message: "Failed to apply remote batch",
        cause: expect.any(Error),
      }),
    );
  });
});
