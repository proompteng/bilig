import { Effect } from "effect";
import { describe, expect, it } from "vitest";
import { ValueTag, type CellSnapshot } from "@bilig/protocol";
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
    typeof Reflect.get(value, "copyRange") === "function" &&
    typeof Reflect.get(value, "importSheetCsv") === "function" &&
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

const EMPTY_CELL_SNAPSHOT: CellSnapshot = {
  sheetName: "Sheet1",
  address: "A1",
  value: { tag: ValueTag.Empty },
  flags: 0,
  version: 0,
};

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
      getCellByIndex: () => EMPTY_CELL_SNAPSHOT,
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
      getCellByIndex: () => EMPTY_CELL_SNAPSHOT,
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
      getCellByIndex: () => EMPTY_CELL_SNAPSHOT,
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

  it("preserves undo history for workbook and sheet render commits", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "before-render-commit" });
    await engine.ready();

    engine.renderCommit([
      { kind: "upsertWorkbook", name: "after-render-commit" },
      { kind: "upsertSheet", name: "Sheet1", order: 0 },
      { kind: "upsertCell", sheetName: "Sheet1", addr: "A1", value: 7 },
    ]);

    expect(engine.exportSnapshot()).toMatchObject({
      workbook: { name: "after-render-commit" },
      sheets: [
        {
          name: "Sheet1",
          cells: [{ address: "A1", value: 7 }],
        },
      ],
    });

    expect(engine.undo()).toBe(true);
    expect(engine.exportSnapshot()).toMatchObject({
      workbook: { name: "before-render-commit" },
      sheets: [],
    });
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

  it("fast-paths simple cell mutation history without restore callbacks", () => {
    const replicaState = createReplicaState("local");
    const workbook = new WorkbookStore("fast-history");
    const sheet = workbook.createSheet("Sheet1");
    const cell = workbook.ensureCellAt(sheet.id, 0, 0);
    workbook.cellStore.setValue(cell.cellIndex, { tag: ValueTag.Number, value: 7 }, 0);
    let replayDepth = 0;
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
      restoreCellOps: () => {
        throw new Error("restoreCellOps should not be used for simple cell history");
      },
      getCellByIndex: () => ({
        sheetName: "Sheet1",
        address: "A1",
        value: { tag: ValueTag.Number, value: 7 },
        flags: 0,
        version: 0,
      }),
      applyBatchNow: () => {},
    });

    const inverseOps = Effect.runSync(
      service.executeLocal([
        { kind: "setCellValue", sheetName: "Sheet1", address: "A1", value: 9 },
      ]),
    );

    expect(inverseOps).toEqual([
      { kind: "setCellValue", sheetName: "Sheet1", address: "A1", value: 7 },
    ]);
  });

  it("does not synthesize blank column identities in delete undo ops", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "undo-columns" });
    await engine.ready();
    engine.createSheet("Sheet1");

    const inverseOps = Effect.runSync(
      getMutationService(engine).executeLocal([
        { kind: "deleteColumns", sheetName: "Sheet1", start: 0, count: 1 },
      ]),
    );

    expect(inverseOps).toContainEqual({
      kind: "insertColumns",
      sheetName: "Sheet1",
      start: 0,
      count: 1,
      entries: [],
    });
  });

  it("copies ranges through the service and rewrites relative formulas", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "copy-range-service" });
    await engine.ready();
    engine.createSheet("Sheet1");
    engine.setCellValue("Sheet1", "A1", 5);
    engine.setCellValue("Sheet1", "A2", 9);
    engine.setCellFormula("Sheet1", "B1", "A1*2");

    Effect.runSync(
      getMutationService(engine).copyRange(
        { sheetName: "Sheet1", startAddress: "B1", endAddress: "B1" },
        { sheetName: "Sheet1", startAddress: "B2", endAddress: "B2" },
      ),
    );

    expect(engine.getCell("Sheet1", "B2").formula).toBe("A2*2");
    expect(engine.getCellValue("Sheet1", "B2")).toEqual({ tag: ValueTag.Number, value: 18 });
  });

  it("imports csv through the service and replaces prior sheet contents", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "csv-import-service" });
    await engine.ready();
    engine.createSheet("Sheet1");
    engine.setCellValue("Sheet1", "A1", 1);
    engine.setCellValue("Sheet1", "C3", 99);

    Effect.runSync(getMutationService(engine).importSheetCsv("Sheet1", '7,=A1*2\n"alpha,beta",'));

    expect(engine.getCellValue("Sheet1", "A1")).toEqual({ tag: ValueTag.Number, value: 7 });
    expect(engine.getCell("Sheet1", "B1").formula).toBe("A1*2");
    expect(engine.getCellValue("Sheet1", "B1")).toEqual({ tag: ValueTag.Number, value: 14 });
    expect(engine.getCell("Sheet1", "A2").value).toEqual({
      tag: ValueTag.String,
      value: "alpha,beta",
      stringId: 1,
    });
    expect(engine.getCellValue("Sheet1", "C3")).toEqual({ tag: ValueTag.Empty });
  });
});
