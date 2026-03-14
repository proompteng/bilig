import { describe, expect, it } from "vitest";
import { createBatch, createReplicaState } from "@bilig/crdt";
import { compactRelayEntries, type RelayEntry } from "../relay-queue.js";

function relayEntriesForSameTarget(entries: RelayEntry[]) {
  return entries.map((entry) => ({
    target: entry.target,
    batchId: entry.batch.id,
    ops: entry.batch.ops,
    deliverAt: entry.deliverAt
  }));
}

describe("relay queue compaction", () => {
  it("collapses repeated writes to the same cell down to the latest batch", () => {
    const state = createReplicaState("playground");
    const first = createBatch(state, [{ kind: "setCellValue", sheetName: "Sheet1", address: "A1", value: 15 }]);
    const second = createBatch(state, [{ kind: "setCellValue", sheetName: "Sheet1", address: "A1", value: 18 }]);

    const queue = compactRelayEntries([
      { target: "mirror", batch: first, deliverAt: 10 },
      { target: "mirror", batch: second, deliverAt: 20 }
    ]);

    expect(relayEntriesForSameTarget(queue)).toEqual([
      {
        target: "mirror",
        batchId: second.id,
        ops: second.ops,
        deliverAt: 20
      }
    ]);
  });

  it("does not compact independent targets together", () => {
    const state = createReplicaState("playground");
    const batch = createBatch(state, [{ kind: "setCellValue", sheetName: "Sheet1", address: "A1", value: 18 }]);

    const queue = compactRelayEntries([
      { target: "mirror", batch, deliverAt: 10 },
      { target: "primary", batch, deliverAt: 15 }
    ]);

    expect(relayEntriesForSameTarget(queue)).toEqual([
      {
        target: "mirror",
        batchId: batch.id,
        ops: batch.ops,
        deliverAt: 10
      },
      {
        target: "primary",
        batchId: batch.id,
        ops: batch.ops,
        deliverAt: 15
      }
    ]);
  });

  it("drops stale cell writes behind a later sheet delete for the same target", () => {
    const state = createReplicaState("playground");
    const setCell = createBatch(state, [{ kind: "setCellValue", sheetName: "Sheet1", address: "A1", value: 5 }]);
    const deleteSheet = createBatch(state, [{ kind: "deleteSheet", name: "Sheet1" }]);

    const queue = compactRelayEntries([
      { target: "mirror", batch: setCell, deliverAt: 10 },
      { target: "mirror", batch: deleteSheet, deliverAt: 30 }
    ]);

    expect(relayEntriesForSameTarget(queue)).toEqual([
      {
        target: "mirror",
        batchId: deleteSheet.id,
        ops: deleteSheet.ops,
        deliverAt: 30
      }
    ]);
  });

  it("retains recreate-after-delete batches and the latest write", () => {
    const state = createReplicaState("playground");
    const deleteSheet = createBatch(state, [{ kind: "deleteSheet", name: "Sheet1" }]);
    const recreateSheet = createBatch(state, [{ kind: "upsertSheet", name: "Sheet1", order: 0 }]);
    const writeCell = createBatch(state, [{ kind: "setCellValue", sheetName: "Sheet1", address: "A1", value: 9 }]);

    const queue = compactRelayEntries([
      { target: "mirror", batch: deleteSheet, deliverAt: 10 },
      { target: "mirror", batch: recreateSheet, deliverAt: 20 },
      { target: "mirror", batch: writeCell, deliverAt: 30 }
    ]);

    expect(relayEntriesForSameTarget(queue)).toEqual([
      {
        target: "mirror",
        batchId: recreateSheet.id,
        ops: recreateSheet.ops,
        deliverAt: 20
      },
      {
        target: "mirror",
        batchId: writeCell.id,
        ops: writeCell.ops,
        deliverAt: 30
      }
    ]);
  });
});
