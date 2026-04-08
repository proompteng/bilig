import { describe, expect, it } from "vitest";
import { isCommitOp, isCommitOps, isEngineReplicaSnapshot } from "../guards.js";

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

  it("accepts render commit ops with supported shapes", () => {
    expect(isCommitOp({ kind: "upsertCell", sheetName: "Sheet1", addr: "A1", value: 7 })).toBe(
      true,
    );
    expect(
      isCommitOps([
        { kind: "upsertWorkbook", name: "book" },
        { kind: "upsertSheet", name: "Sheet1", order: 0 },
        { kind: "deleteCell", sheetName: "Sheet1", addr: "A1" },
      ]),
    ).toBe(true);
  });

  it("rejects malformed render commit ops", () => {
    expect(isCommitOp({ kind: "upsertCell", sheetName: "Sheet1", addr: "A1" })).toBe(false);
    expect(isCommitOps([{ kind: "renameSheet", oldName: "A" }])).toBe(false);
  });
});
