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

  it("compacts stale entity writes down to the latest op", () => {
    const state = createReplicaState("replica");
    const createSheet = createBatch(state, [{ kind: "upsertSheet", name: "Sheet1", order: 0 }]);
    const firstWrite = createBatch(state, [{ kind: "setCellValue", sheetName: "Sheet1", address: "A1", value: 1 }]);
    const secondWrite = createBatch(state, [{ kind: "setCellValue", sheetName: "Sheet1", address: "A1", value: 2 }]);

    expect(compactLog([createSheet, firstWrite, secondWrite])).toEqual([
      createSheet,
      secondWrite
    ]);
  });

  it("drops stale cell writes that sit behind a later sheet tombstone", () => {
    const state = createReplicaState("replica");
    const createSheet = createBatch(state, [{ kind: "upsertSheet", name: "Sheet1", order: 0 }]);
    const firstWrite = createBatch(state, [{ kind: "setCellValue", sheetName: "Sheet1", address: "A1", value: 1 }]);
    const deleteSheet = createBatch(state, [{ kind: "deleteSheet", name: "Sheet1" }]);

    expect(compactLog([createSheet, firstWrite, deleteSheet])).toEqual([
      deleteSheet
    ]);
  });

  it("retains recreated sheets and newer cell writes after tombstones", () => {
    const state = createReplicaState("replica");
    const createSheet = createBatch(state, [{ kind: "upsertSheet", name: "Sheet1", order: 0 }]);
    const firstWrite = createBatch(state, [{ kind: "setCellValue", sheetName: "Sheet1", address: "A1", value: 1 }]);
    const deleteSheet = createBatch(state, [{ kind: "deleteSheet", name: "Sheet1" }]);
    const recreateSheet = createBatch(state, [{ kind: "upsertSheet", name: "Sheet1", order: 0 }]);
    const secondWrite = createBatch(state, [{ kind: "setCellValue", sheetName: "Sheet1", address: "A1", value: 9 }]);

    expect(compactLog([createSheet, firstWrite, deleteSheet, recreateSheet, secondWrite])).toEqual([
      recreateSheet,
      secondWrite
    ]);
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

  it("compacts workbook metadata entities by logical key", () => {
    const state = createReplicaState("replica");
    const define = createBatch(state, [{ kind: "upsertDefinedName", name: "TaxRate", value: 0.08 }]);
    const redefine = createBatch(state, [{ kind: "upsertDefinedName", name: "taxrate", value: 0.09 }]);
    const pivotCreate = createBatch(state, [{
      kind: "upsertPivotTable",
      name: "SalesByRegion",
      sheetName: "Pivot",
      address: "B2",
      source: { sheetName: "Data", startAddress: "A1", endAddress: "C4" },
      groupBy: ["Region"],
      values: [{ sourceColumn: "Amount", summarizeBy: "sum" }],
      rows: 3,
      cols: 2
    }]);
    const pivotDelete = createBatch(state, [{ kind: "deletePivotTable", sheetName: "Pivot", address: "B2" }]);

    expect(compactLog([define, redefine, pivotCreate, pivotDelete])).toEqual([
      redefine,
      pivotDelete
    ]);
  });
});
