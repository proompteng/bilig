import { describe, expect, it } from "vitest";
import {
  batchOpOrder,
  compareOpOrder,
  compactLog,
  createBatch,
  createReplicaState,
  exportReplicaSnapshot,
  hydrateReplicaState,
  importReplicaSnapshot,
  markBatchApplied,
  mergeBatches,
  shouldApplyBatch,
} from "../replica-state.js";

describe("replica-state", () => {
  it("orders batches deterministically", () => {
    const a = createReplicaState("a");
    const b = createReplicaState("b");
    const merged = mergeBatches([
      createBatch(b, [{ kind: "upsertSheet", name: "Sheet1", order: 0 }]),
      createBatch(a, [{ kind: "upsertSheet", name: "Sheet1", order: 0 }]),
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
    const firstWrite = createBatch(state, [
      { kind: "setCellValue", sheetName: "Sheet1", address: "A1", value: 1 },
    ]);
    const secondWrite = createBatch(state, [
      { kind: "setCellValue", sheetName: "Sheet1", address: "A1", value: 2 },
    ]);

    expect(compactLog([createSheet, firstWrite, secondWrite])).toEqual([createSheet, secondWrite]);
  });

  it("drops stale cell writes that sit behind a later sheet tombstone", () => {
    const state = createReplicaState("replica");
    const createSheet = createBatch(state, [{ kind: "upsertSheet", name: "Sheet1", order: 0 }]);
    const firstWrite = createBatch(state, [
      { kind: "setCellValue", sheetName: "Sheet1", address: "A1", value: 1 },
    ]);
    const deleteSheet = createBatch(state, [{ kind: "deleteSheet", name: "Sheet1" }]);

    expect(compactLog([createSheet, firstWrite, deleteSheet])).toEqual([deleteSheet]);
  });

  it("retains recreated sheets and newer cell writes after tombstones", () => {
    const state = createReplicaState("replica");
    const createSheet = createBatch(state, [{ kind: "upsertSheet", name: "Sheet1", order: 0 }]);
    const firstWrite = createBatch(state, [
      { kind: "setCellValue", sheetName: "Sheet1", address: "A1", value: 1 },
    ]);
    const deleteSheet = createBatch(state, [{ kind: "deleteSheet", name: "Sheet1" }]);
    const recreateSheet = createBatch(state, [{ kind: "upsertSheet", name: "Sheet1", order: 0 }]);
    const secondWrite = createBatch(state, [
      { kind: "setCellValue", sheetName: "Sheet1", address: "A1", value: 9 },
    ]);

    expect(compactLog([createSheet, firstWrite, deleteSheet, recreateSheet, secondWrite])).toEqual([
      recreateSheet,
      secondWrite,
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

  it("hydrates snapshots, trims export history, and respects apply eligibility", () => {
    const state = createReplicaState("replica");
    const first = createBatch(state, [{ kind: "upsertWorkbook", name: "book" }]);
    const second = createBatch(state, [{ kind: "upsertSheet", name: "Sheet1", order: 0 }]);
    const third = createBatch(state, [{ kind: "upsertSheet", name: "Sheet2", order: 1 }]);
    const exported = exportReplicaSnapshot(state, 2);

    expect(exported).toEqual({
      replicaId: "replica",
      counter: 3,
      appliedBatchIds: [second.id, third.id],
    });

    const restored = createReplicaState("stale");
    hydrateReplicaState(restored, exported);
    expect(restored.replicaId).toBe("replica");
    expect(restored.clock.counter).toBe(3);
    expect(shouldApplyBatch(restored, first)).toBe(true);
    expect(shouldApplyBatch(restored, second)).toBe(false);

    const external = {
      id: "remote:9",
      replicaId: "remote",
      clock: { counter: 9 },
      ops: [{ kind: "upsertWorkbook" as const, name: "remote-book" }],
    };
    markBatchApplied(restored, external);
    expect(restored.clock.counter).toBe(9);
    expect(restored.appliedBatchIds.has(external.id)).toBe(true);
  });

  it("orders operations deterministically inside competing batches", () => {
    const left = {
      counter: 4,
      replicaId: "a",
      batchId: "a:4",
      opIndex: 0,
    };
    const right = {
      counter: 4,
      replicaId: "b",
      batchId: "b:4",
      opIndex: 0,
    };

    expect(compareOpOrder(left, right)).toBeLessThan(0);
    expect(compareOpOrder(right, left)).toBeGreaterThan(0);
  });

  it("assigns monotonically increasing clocks within a replica", () => {
    const state = createReplicaState("replica");
    const b1 = createBatch(state, [{ kind: "upsertWorkbook", name: "book" }]);
    const b2 = createBatch(state, [{ kind: "upsertSheet", name: "Sheet1", order: 0 }]);

    expect(b1.clock.counter).toBe(1);
    expect(b2.clock.counter).toBe(2);
    expect(b1.id).toBe("replica:1");
    expect(b2.id).toBe("replica:2");
  });

  it("compacts workbook metadata entities by logical key", () => {
    const state = createReplicaState("replica");
    const define = createBatch(state, [
      { kind: "upsertDefinedName", name: "TaxRate", value: 0.08 },
    ]);
    const redefine = createBatch(state, [
      { kind: "upsertDefinedName", name: "taxrate", value: 0.09 },
    ]);
    const pivotCreate = createBatch(state, [
      {
        kind: "upsertPivotTable",
        name: "SalesByRegion",
        sheetName: "Pivot",
        address: "B2",
        source: { sheetName: "Data", startAddress: "A1", endAddress: "C4" },
        groupBy: ["Region"],
        values: [{ sourceColumn: "Amount", summarizeBy: "sum" }],
        rows: 3,
        cols: 2,
      },
    ]);
    const pivotDelete = createBatch(state, [
      { kind: "deletePivotTable", sheetName: "Pivot", address: "B2" },
    ]);

    expect(compactLog([define, redefine, pivotCreate, pivotDelete])).toEqual([
      redefine,
      pivotDelete,
    ]);
  });

  it("drops stale row and range metadata that sit behind a later sheet tombstone", () => {
    const state = createReplicaState("replica");
    const createSheet = createBatch(state, [{ kind: "upsertSheet", name: "Sheet1", order: 0 }]);
    const rowMetadata = createBatch(state, [
      {
        kind: "updateRowMetadata",
        sheetName: "Sheet1",
        start: 0,
        count: 2,
        size: 32,
        hidden: null,
      },
    ]);
    const styleRange = createBatch(state, [
      {
        kind: "setStyleRange",
        range: { sheetName: "Sheet1", startAddress: "A1", endAddress: "B2" },
        styleId: "style:body",
      },
    ]);
    const formatRange = createBatch(state, [
      {
        kind: "setFormatRange",
        range: { sheetName: "Sheet1", startAddress: "A1", endAddress: "B2" },
        formatId: "fmt:body",
      },
    ]);
    const deleteSheet = createBatch(state, [{ kind: "deleteSheet", name: "Sheet1" }]);

    expect(compactLog([createSheet, rowMetadata, styleRange, formatRange, deleteSheet])).toEqual([
      deleteSheet,
    ]);
  });

  it("keeps global style and number format entities while compacting them by id", () => {
    const state = createReplicaState("replica");
    const firstStyle = createBatch(state, [
      { kind: "upsertCellStyle", style: { id: "style:body", fill: { backgroundColor: "#fff" } } },
    ]);
    const latestStyle = createBatch(state, [
      { kind: "upsertCellStyle", style: { id: "style:body", fill: { backgroundColor: "#eee" } } },
    ]);
    const firstFormat = createBatch(state, [
      { kind: "upsertCellNumberFormat", format: { id: "fmt:body", code: "0.0", kind: "number" } },
    ]);
    const latestFormat = createBatch(state, [
      { kind: "upsertCellNumberFormat", format: { id: "fmt:body", code: "0.00", kind: "number" } },
    ]);

    expect(compactLog([firstStyle, latestStyle, firstFormat, latestFormat])).toEqual([
      latestStyle,
      latestFormat,
    ]);
  });

  it("compacts media entities by normalized id and drops sheet-barriered upserts", () => {
    const state = createReplicaState("replica");
    const createData = createBatch(state, [{ kind: "upsertSheet", name: "Data", order: 0 }]);
    const createDashboard = createBatch(state, [
      { kind: "upsertSheet", name: "Dashboard", order: 1 },
    ]);
    const chartUpsert = createBatch(state, [
      {
        kind: "upsertChart",
        chart: {
          id: " revenue ",
          sheetName: "Dashboard",
          address: "B2",
          source: { sheetName: "Data", startAddress: "A1", endAddress: "B4" },
          chartType: "line",
          rows: 4,
          cols: 5,
        },
      },
    ]);
    const imageUpsert = createBatch(state, [
      {
        kind: "upsertImage",
        image: {
          id: " logo ",
          sheetName: "Dashboard",
          address: "D2",
          sourceUrl: "https://example.com/logo.png",
          rows: 3,
          cols: 2,
        },
      },
    ]);
    const shapeUpsert = createBatch(state, [
      {
        kind: "upsertShape",
        shape: {
          id: " callout ",
          sheetName: "Dashboard",
          address: "F2",
          shapeType: "textBox",
          rows: 2,
          cols: 3,
        },
      },
    ]);
    const deleteData = createBatch(state, [{ kind: "deleteSheet", name: "Data" }]);
    const deleteDashboard = createBatch(state, [{ kind: "deleteSheet", name: "Dashboard" }]);
    const chartDelete = createBatch(state, [{ kind: "deleteChart", id: "REVENUE" }]);
    const imageDelete = createBatch(state, [{ kind: "deleteImage", id: "LOGO" }]);
    const shapeDelete = createBatch(state, [{ kind: "deleteShape", id: "CALLOUT" }]);

    expect(
      compactLog([
        createData,
        createDashboard,
        chartUpsert,
        imageUpsert,
        shapeUpsert,
        deleteData,
        deleteDashboard,
        chartDelete,
        imageDelete,
        shapeDelete,
      ]),
    ).toEqual([deleteData, deleteDashboard, chartDelete, imageDelete, shapeDelete]);

    expect(compareOpOrder(batchOpOrder(chartDelete, 0), batchOpOrder(imageDelete, 0))).toBeLessThan(
      0,
    );
  });

  it("covers remaining entity and sheet-barrier branches across structural metadata ops", () => {
    const state = createReplicaState("replica");
    const opsBatch = createBatch(state, [
      { kind: "setWorkbookMetadata", key: "Author", value: "greg" },
      { kind: "setCalculationSettings", settings: { mode: "manual" } },
      { kind: "setVolatileContext", context: { recalcEpoch: 2 } },
      { kind: "renameSheet", oldName: "Sheet1", newName: "Renamed", order: 0 },
      { kind: "insertRows", sheetName: "Sheet1", start: 0, count: 1 },
      { kind: "deleteRows", sheetName: "Sheet1", start: 2, count: 1 },
      { kind: "moveRows", sheetName: "Sheet1", start: 3, count: 1, target: 5 },
      { kind: "insertColumns", sheetName: "Sheet1", start: 0, count: 1 },
      { kind: "deleteColumns", sheetName: "Sheet1", start: 2, count: 1 },
      { kind: "moveColumns", sheetName: "Sheet1", start: 3, count: 1, target: 5 },
      {
        kind: "updateRowMetadata",
        sheetName: "Sheet1",
        start: 0,
        count: 2,
        size: 24,
        hidden: null,
      },
      {
        kind: "updateColumnMetadata",
        sheetName: "Sheet1",
        start: 0,
        count: 2,
        size: null,
        hidden: true,
      },
      { kind: "setFreezePane", sheetName: "Sheet1", rows: 1, cols: 1 },
      { kind: "clearFreezePane", sheetName: "Sheet1" },
      { kind: "setSheetProtection", protection: { sheetName: "Sheet1", hideFormulas: true } },
      { kind: "clearSheetProtection", sheetName: "Sheet1" },
      {
        kind: "setFilter",
        sheetName: "Sheet1",
        range: { sheetName: "Sheet1", startAddress: "A1", endAddress: "B3" },
      },
      {
        kind: "clearFilter",
        sheetName: "Sheet1",
        range: { sheetName: "Sheet1", startAddress: "A1", endAddress: "B3" },
      },
      {
        kind: "setSort",
        sheetName: "Sheet1",
        range: { sheetName: "Sheet1", startAddress: "A1", endAddress: "B3" },
        keys: [{ keyAddress: "A1", direction: "asc" }],
      },
      {
        kind: "clearSort",
        sheetName: "Sheet1",
        range: { sheetName: "Sheet1", startAddress: "A1", endAddress: "B3" },
      },
      {
        kind: "setDataValidation",
        validation: {
          range: { sheetName: "Sheet1", startAddress: "A1", endAddress: "B3" },
          rule: { kind: "list", values: ["Draft"] },
        },
      },
      {
        kind: "clearDataValidation",
        sheetName: "Sheet1",
        range: { sheetName: "Sheet1", startAddress: "A1", endAddress: "B3" },
      },
      {
        kind: "upsertConditionalFormat",
        format: {
          id: "cf-1",
          range: { sheetName: "Sheet1", startAddress: "A1", endAddress: "B3" },
          rule: { kind: "blanks" },
          style: {},
        },
      },
      { kind: "deleteConditionalFormat", id: "cf-1" },
      {
        kind: "upsertRangeProtection",
        protection: {
          id: "protect-1",
          range: { sheetName: "Sheet1", startAddress: "A1", endAddress: "B3" },
        },
      },
      { kind: "deleteRangeProtection", id: "protect-1" },
      {
        kind: "upsertCommentThread",
        thread: {
          threadId: "thread-1",
          sheetName: "Sheet1",
          address: "A1",
          comments: [{ id: "comment-1", body: "check" }],
        },
      },
      { kind: "deleteCommentThread", sheetName: "Sheet1", address: "A1" },
      { kind: "upsertNote", note: { sheetName: "Sheet1", address: "A2", text: "note" } },
      { kind: "deleteNote", sheetName: "Sheet1", address: "A2" },
      { kind: "setCellFormula", sheetName: "Sheet1", address: "B1", formula: "A1+1" },
      { kind: "setCellFormat", sheetName: "Sheet1", address: "B1", format: "0.00" },
      {
        kind: "upsertCellStyle",
        style: { id: "style:body", fill: { backgroundColor: "#fff" } },
      },
      {
        kind: "upsertCellNumberFormat",
        format: { id: "fmt:body", code: "0.00", kind: "number" },
      },
      {
        kind: "setStyleRange",
        range: { sheetName: "Sheet1", startAddress: "A1", endAddress: "B3" },
        styleId: "style:body",
      },
      {
        kind: "setFormatRange",
        range: { sheetName: "Sheet1", startAddress: "A1", endAddress: "B3" },
        formatId: "fmt:body",
      },
      { kind: "deleteDefinedName", name: "TaxRate" },
      { kind: "deleteTable", name: "Revenue" },
      { kind: "upsertSpillRange", sheetName: "Sheet1", address: "D4", rows: 2, cols: 2 },
      { kind: "deleteSpillRange", sheetName: "Sheet1", address: "D4" },
      { kind: "deletePivotTable", sheetName: "Sheet1", address: "F2" },
    ]);
    const deleteSheet = createBatch(state, [{ kind: "deleteSheet", name: "Sheet1" }]);

    const compacted = compactLog([opsBatch, deleteSheet]);
    const compactedKinds = compacted.flatMap((batch) => batch.ops.map((op) => op.kind));

    expect(compactedKinds).toContain("setWorkbookMetadata");
    expect(compactedKinds).toContain("setCalculationSettings");
    expect(compactedKinds).toContain("setVolatileContext");
    expect(compactedKinds).toContain("deleteDefinedName");
    expect(compactedKinds).toContain("deleteTable");
    expect(compactedKinds).toContain("upsertCellStyle");
    expect(compactedKinds).toContain("upsertCellNumberFormat");
    expect(compactedKinds).toContain("deleteSheet");
    expect(compactedKinds).not.toContain("setFreezePane");
    expect(compactedKinds).not.toContain("setCellFormula");
    expect(compactedKinds).not.toContain("deletePivotTable");
  });
});
