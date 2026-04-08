import { describe, expect, it } from "vitest";
import { isEngineReplicaSnapshot } from "../guards.js";

describe("core guards", () => {
  it("accepts engine replica snapshots with the shipped shape", () => {
    expect(
      isEngineReplicaSnapshot({
        replica: { replicaId: "browser:test", counter: 1, appliedBatchIds: [] },
        entityVersions: [],
        sheetDeleteVersions: [],
      }),
    ).toBe(true);
  });

  it("rejects engine replica snapshots without version arrays", () => {
    expect(
      isEngineReplicaSnapshot({
        replica: { replicaId: "browser:test", counter: 1, appliedBatchIds: [] },
        entityVersions: null,
        sheetDeleteVersions: [],
      }),
    ).toBe(false);
  });
});
