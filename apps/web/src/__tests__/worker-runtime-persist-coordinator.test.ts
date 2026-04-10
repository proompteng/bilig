import { describe, expect, it, vi } from "vitest";
import type { ProjectionOverlayScope } from "../worker-local-overlay.js";
import {
  resolveProjectionOverlayScopeForPersist,
  WorkerRuntimePersistCoordinator,
} from "../worker-runtime-persist-coordinator.js";

function createDeferredPromise() {
  let resolve!: () => void;
  let reject!: (error: unknown) => void;
  const promise = new Promise<void>((nextResolve, nextReject) => {
    resolve = nextResolve;
    reject = nextReject;
  });
  return { promise, resolve, reject };
}

describe("worker runtime persist coordinator", () => {
  it("creates an explicit empty overlay scope when no pending mutations remain", () => {
    const scope = resolveProjectionOverlayScopeForPersist({
      projectionOverlayScope: null,
      pendingMutationCount: 0,
    });

    expect(scope).not.toBeNull();
    expect(scope).toEqual({
      fullScan: false,
      cellAddressesBySheet: new Map(),
      rowAxisIndicesBySheet: new Map(),
      columnAxisIndicesBySheet: new Map(),
    });
  });

  it("preserves an existing overlay scope when one is already tracked", () => {
    const scope: ProjectionOverlayScope = {
      fullScan: false,
      cellAddressesBySheet: new Map([["Sheet1", new Set(["A1"])]]),
      rowAxisIndicesBySheet: new Map(),
      columnAxisIndicesBySheet: new Map(),
    };

    expect(
      resolveProjectionOverlayScopeForPersist({
        projectionOverlayScope: scope,
        pendingMutationCount: 0,
      }),
    ).toBe(scope);
  });

  it("returns null when pending mutations still require a projection overlay", () => {
    expect(
      resolveProjectionOverlayScopeForPersist({
        projectionOverlayScope: null,
        pendingMutationCount: 2,
      }),
    ).toBeNull();
  });

  it("re-runs persistence when another persist is queued while save is in flight", async () => {
    const firstPersist = createDeferredPromise();
    const saveState = vi
      .fn()
      .mockImplementationOnce(async () => await firstPersist.promise)
      .mockResolvedValueOnce(undefined);

    const coordinator = new WorkerRuntimePersistCoordinator({
      canPersistState: () => true,
      getLocalStore: () => ({ persistProjectionState: vi.fn() }),
      getAuthoritativeEngine: vi.fn(async () => ({ id: "authoritative" })),
      getProjectionEngine: vi.fn(async () => ({ id: "projection" })),
      buildPersistedState: vi.fn(() => ({ id: "persisted" })),
      getProjectionOverlayScope: vi.fn(() => null),
      saveState,
      markProjectionMatchesLocalStore: vi.fn(),
    });

    const first = coordinator.queuePersist();
    const second = coordinator.queuePersist();

    await new Promise((resolve) => {
      setTimeout(resolve, 0);
    });
    expect(saveState).toHaveBeenCalledTimes(1);
    firstPersist.resolve();
    await Promise.all([first, second]);

    expect(saveState).toHaveBeenCalledTimes(2);
  });

  it("does not mark the local projection as persisted when save fails", async () => {
    const markProjectionMatchesLocalStore = vi.fn();
    const coordinator = new WorkerRuntimePersistCoordinator({
      canPersistState: () => true,
      getLocalStore: () => ({ persistProjectionState: vi.fn() }),
      getAuthoritativeEngine: vi.fn(async () => ({ id: "authoritative" })),
      getProjectionEngine: vi.fn(async () => ({ id: "projection" })),
      buildPersistedState: vi.fn(() => ({ id: "persisted" })),
      getProjectionOverlayScope: vi.fn(() => null),
      saveState: vi.fn(async () => {
        throw new Error("persist failed");
      }),
      markProjectionMatchesLocalStore,
    });

    await expect(coordinator.queuePersist()).rejects.toThrow("persist failed");
    expect(markProjectionMatchesLocalStore).not.toHaveBeenCalled();
  });
});
