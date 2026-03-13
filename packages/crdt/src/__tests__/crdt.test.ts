import { describe, expect, it } from "vitest";
import {
  compareOpOrder,
  compactLog,
  createBatch,
  createReplicaState,
  exportReplicaSnapshot,
  importReplicaSnapshot,
  mergeBatches
} from "../index.js";

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

  it("roundtrips replica snapshots", () => {
    const state = createReplicaState("replica");
    createBatch(state, [{ kind: "upsertWorkbook", name: "book" }]);
    createBatch(state, [{ kind: "upsertSheet", name: "Sheet1", order: 0 }]);

    const restored = importReplicaSnapshot(exportReplicaSnapshot(state));
    expect(restored.replicaId).toBe("replica");
    expect(restored.clock.counter).toBe(2);
    expect([...restored.appliedBatchIds]).toEqual(["replica:1", "replica:2"]);
  });

  it("orders operations deterministically inside competing batches", () => {
    const left = {
      counter: 4,
      replicaId: "a",
      batchId: "a:4",
      opIndex: 0
    };
    const right = {
      counter: 4,
      replicaId: "b",
      batchId: "b:4",
      opIndex: 0
    };

    expect(compareOpOrder(left, right)).toBeLessThan(0);
    expect(compareOpOrder(right, left)).toBeGreaterThan(0);
  });
});
