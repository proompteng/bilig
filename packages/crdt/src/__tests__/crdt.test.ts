import { describe, expect, it } from "vitest";
import { compactLog, createBatch, createReplicaState, mergeBatches } from "../index.js";

describe("crdt", () => {
  it("orders batches deterministically", () => {
    const a = createReplicaState("a");
    const b = createReplicaState("b");
    const merged = mergeBatches([
      createBatch(b, [{ kind: "upsertSheet", name: "Sheet1", order: 0 }]),
      createBatch(a, [{ kind: "upsertSheet", name: "Sheet1", order: 0 }])
    ]);
    expect(merged.map((batch) => batch.replicaId)).toEqual(["a", "b"]);
  });

  it("deduplicates logs by id", () => {
    const state = createReplicaState("replica");
    const batch = createBatch(state, [{ kind: "upsertWorkbook", name: "book" }]);
    expect(compactLog([batch, batch])).toHaveLength(1);
  });
});
