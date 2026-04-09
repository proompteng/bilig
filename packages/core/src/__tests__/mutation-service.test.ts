import { Effect } from "effect";
import { describe, expect, it } from "vitest";
import { createReplicaState } from "../replica-state.js";
import { SpreadsheetEngine } from "../engine.js";
import { createEngineMutationService } from "../engine/services/mutation-service.js";
import type { EngineMutationService } from "../engine/services/mutation-service.js";
import { WorkbookStore } from "../workbook-store.js";

function isEngineMutationService(value: unknown): value is EngineMutationService {
  if (typeof value !== "object" || value === null) {
    return false;
  }
  return (
    typeof Reflect.get(value, "executeLocal") === "function" &&
    typeof Reflect.get(value, "captureUndoOps") === "function" &&
    typeof Reflect.get(value, "renderCommit") === "function"
  );
}

function getMutationService(engine: SpreadsheetEngine): EngineMutationService {
  const runtime = Reflect.get(engine, "runtime");
  if (typeof runtime !== "object" || runtime === null) {
    throw new TypeError("Expected engine runtime");
  }
  const mutation = Reflect.get(runtime, "mutation");
  if (!isEngineMutationService(mutation)) {
    throw new TypeError("Expected engine mutation service");
  }
  return mutation;
}

describe("EngineMutationService", () => {
  it("clears redo history when a new local transaction lands", () => {
    const replicaState = createReplicaState("local");
    const workbook = new WorkbookStore("inverse");
    let replayDepth = 0;
    const batches: import("@bilig/workbook-domain").EngineOpBatch[] = [];
    const service = createEngineMutationService({
      state: {
        workbook,
        replicaState,
        undoStack: [],
        redoStack: [{ forward: { ops: [] }, inverse: { ops: [] } }],
        getTransactionReplayDepth: () => replayDepth,
        setTransactionReplayDepth: (next) => {
          replayDepth = next;
        },
      },
      captureSheetCellState: () => [],
      captureRowRangeCellState: () => [],
      captureColumnRangeCellState: () => [],
      restoreCellOps: () => [{ kind: "upsertWorkbook", name: "inverse" }],
      applyBatchNow: (batch) => {
        batches.push(batch);
      },
    });

    const undoOps = Effect.runSync(
      service.executeLocal([{ kind: "upsertWorkbook", name: "forward" }]),
    );

    expect(undoOps).toEqual([{ kind: "upsertWorkbook", name: "inverse" }]);
    expect(batches).toHaveLength(1);
    expect(service).toBeDefined();
  });

  it("captures a single local transaction and clones the undo ops", () => {
    const replicaState = createReplicaState("local");
    const workbook = new WorkbookStore("inverse");
    let replayDepth = 0;
    const state = {
      workbook,
      replicaState,
      undoStack: [] as Array<{ forward: { ops: unknown[] }; inverse: { ops: unknown[] } }>,
      redoStack: [] as Array<{ forward: { ops: unknown[] }; inverse: { ops: unknown[] } }>,
      getTransactionReplayDepth: () => replayDepth,
      setTransactionReplayDepth: (next: number) => {
        replayDepth = next;
      },
    };
    const service = createEngineMutationService({
      state,
      captureSheetCellState: () => [],
      captureRowRangeCellState: () => [],
      captureColumnRangeCellState: () => [],
      restoreCellOps: () => [{ kind: "upsertWorkbook", name: "inverse" }],
      applyBatchNow: () => {},
    });

    const captured = Effect.runSync(
      service.captureUndoOps(() =>
        Effect.runSync(service.executeLocal([{ kind: "upsertWorkbook", name: "forward" }])),
      ),
    );

    expect(captured.undoOps).toEqual([{ kind: "upsertWorkbook", name: "inverse" }]);
  });

  it("drops malformed render commit records instead of forwarding partial engine ops", () => {
    const replicaState = createReplicaState("local");
    const workbook = new WorkbookStore("spec");
    let replayDepth = 0;
    const batches: import("@bilig/workbook-domain").EngineOpBatch[] = [];
    const service = createEngineMutationService({
      state: {
        workbook,
        replicaState,
        undoStack: [],
        redoStack: [],
        getTransactionReplayDepth: () => replayDepth,
        setTransactionReplayDepth: (next) => {
          replayDepth = next;
        },
      },
      captureSheetCellState: () => [],
      captureRowRangeCellState: () => [],
      captureColumnRangeCellState: () => [],
      restoreCellOps: () => [],
      applyBatchNow: (batch) => {
        batches.push(batch);
      },
    });

    Effect.runSync(
      service.renderCommit([
        { kind: "renameSheet", oldName: "Old" },
        { kind: "upsertCell", sheetName: "Sheet1" },
        { kind: "upsertSheet", name: "Sheet1", order: 0 },
        { kind: "upsertCell", sheetName: "Sheet1", addr: "A1", value: 7, format: "0.00" },
      ]),
    );

    expect(batches).toHaveLength(1);
    expect(batches[0]?.ops).toEqual([
      { kind: "upsertSheet", name: "Sheet1", order: 0 },
      { kind: "setCellValue", sheetName: "Sheet1", address: "A1", value: 7 },
      { kind: "setCellFormat", sheetName: "Sheet1", address: "A1", format: "0.00" },
    ]);
  });

  it("captures sheet metadata and cells when building delete-sheet undo ops", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "undo-sheet" });
    await engine.ready();
    engine.createSheet("Sheet1");
    engine.setCellValue("Sheet1", "A1", 7);
    engine.setCellFormula("Sheet1", "B1", "A1*2");
    engine.setFreezePane("Sheet1", 1, 0);
    engine.setFilter("Sheet1", { sheetName: "Sheet1", startAddress: "A1", endAddress: "B3" });
    engine.setSort("Sheet1", { sheetName: "Sheet1", startAddress: "A1", endAddress: "B3" }, [
      { keyAddress: "B1", direction: "asc" },
    ]);
    engine.setTable({
      name: "Sales",
      sheetName: "Sheet1",
      startAddress: "A1",
      endAddress: "B3",
      columnNames: ["Amount", "Total"],
      headerRow: true,
      totalsRow: false,
    });

    const inverseOps = Effect.runSync(
      getMutationService(engine).executeLocal([{ kind: "deleteSheet", name: "Sheet1" }]),
    );

    expect(inverseOps).not.toBeNull();
    expect(inverseOps).toContainEqual({ kind: "upsertSheet", name: "Sheet1", order: 0 });
    expect(inverseOps).toContainEqual({
      kind: "setFreezePane",
      sheetName: "Sheet1",
      rows: 1,
      cols: 0,
    });
    expect(inverseOps).toContainEqual({
      kind: "setFilter",
      sheetName: "Sheet1",
      range: { sheetName: "Sheet1", startAddress: "A1", endAddress: "B3" },
    });
    expect(inverseOps).toContainEqual({
      kind: "setSort",
      sheetName: "Sheet1",
      range: { sheetName: "Sheet1", startAddress: "A1", endAddress: "B3" },
      keys: [{ keyAddress: "B1", direction: "asc" }],
    });
    expect(inverseOps).toContainEqual({
      kind: "upsertTable",
      table: {
        name: "Sales",
        sheetName: "Sheet1",
        startAddress: "A1",
        endAddress: "B3",
        columnNames: ["Amount", "Total"],
        headerRow: true,
        totalsRow: false,
      },
    });
    expect(inverseOps).toContainEqual({
      kind: "setCellValue",
      sheetName: "Sheet1",
      address: "A1",
      value: 7,
    });
    expect(inverseOps).toContainEqual({
      kind: "setCellFormula",
      sheetName: "Sheet1",
      address: "B1",
      formula: "A1*2",
    });
  });

  it("captures deleted row cells in reverse order-safe undo ops", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "undo-rows" });
    await engine.ready();
    engine.createSheet("Sheet1");
    engine.setCellValue("Sheet1", "A2", 10);
    engine.setCellFormula("Sheet1", "B2", "A2*3");
    engine.updateRowMetadata("Sheet1", 1, 1, 24, false);

    const inverseOps = Effect.runSync(
      getMutationService(engine).executeLocal([
        { kind: "deleteRows", sheetName: "Sheet1", start: 1, count: 1 },
      ]),
    );

    expect(inverseOps).not.toBeNull();
    expect(inverseOps).toContainEqual({
      kind: "insertRows",
      sheetName: "Sheet1",
      start: 1,
      count: 1,
      entries: [{ id: "row-1", index: 1, size: 24, hidden: false }],
    });
    expect(inverseOps).toContainEqual({
      kind: "setCellValue",
      sheetName: "Sheet1",
      address: "A2",
      value: 10,
    });
    expect(inverseOps).toContainEqual({
      kind: "setCellFormula",
      sheetName: "Sheet1",
      address: "B2",
      formula: "A2*3",
    });
  });
});
