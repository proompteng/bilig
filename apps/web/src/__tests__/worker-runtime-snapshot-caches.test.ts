import { describe, expect, it, vi } from "vitest";
import { WorkerRuntimeSnapshotCaches } from "../worker-runtime-snapshot-caches.js";

describe("worker runtime snapshot caches", () => {
  it("reuses cached projection snapshots until invalidated", () => {
    const caches = new WorkerRuntimeSnapshotCaches();
    const exportSnapshot = vi
      .fn<() => { version: string }>()
      .mockReturnValueOnce({ version: "first" })
      .mockReturnValueOnce({ version: "second" });

    expect(caches.getProjectionSnapshot(exportSnapshot)).toEqual({ version: "first" });
    expect(caches.getProjectionSnapshot(exportSnapshot)).toEqual({ version: "first" });

    caches.invalidateProjectionSnapshot();

    expect(caches.getProjectionSnapshot(exportSnapshot)).toEqual({ version: "second" });
    expect(exportSnapshot).toHaveBeenCalledTimes(2);
  });

  it("reuses projection state for authoritative exports when no separate base is loaded", () => {
    const caches = new WorkerRuntimeSnapshotCaches();

    expect(
      caches.getAuthoritativeSnapshot({
        canReuseProjectionState: true,
        exportProjectionSnapshot: () => ({ source: "projection" }),
        exportAuthoritativeSnapshot: () => ({ source: "authoritative" }),
      }),
    ).toEqual({ source: "projection" });

    expect(
      caches.getAuthoritativeReplica({
        canReuseProjectionState: true,
        exportProjectionReplica: () => ({ source: "projection" }),
        exportAuthoritativeReplica: () => ({ source: "authoritative" }),
      }),
    ).toEqual({ source: "projection" });
  });
});
