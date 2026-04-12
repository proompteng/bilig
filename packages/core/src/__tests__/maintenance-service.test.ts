/* eslint-disable typescript-eslint/no-unsafe-type-assertion -- error-path tests intentionally inject partial collaborators */
import { Effect } from "effect";
import { describe, expect, it } from "vitest";
import type { EngineOp } from "@bilig/workbook-domain";
import { SpreadsheetEngine } from "../engine.js";
import {
  createEngineMaintenanceService,
  type EngineMaintenanceService,
} from "../engine/services/maintenance-service.js";

function isEngineMaintenanceService(value: unknown): value is EngineMaintenanceService {
  if (typeof value !== "object" || value === null) {
    return false;
  }
  return (
    typeof Reflect.get(value, "estimatePotentialNewCells") === "function" &&
    typeof Reflect.get(value, "rewriteDefinedNamesForSheetRename") === "function" &&
    typeof Reflect.get(value, "resetWorkbook") === "function"
  );
}

function getMaintenanceService(engine: SpreadsheetEngine): EngineMaintenanceService {
  const runtime = Reflect.get(engine, "runtime");
  if (typeof runtime !== "object" || runtime === null) {
    throw new TypeError("Expected engine runtime");
  }
  const maintenance = Reflect.get(runtime, "maintenance");
  if (!isEngineMaintenanceService(maintenance)) {
    throw new TypeError("Expected engine maintenance service");
  }
  return maintenance;
}

describe("EngineMaintenanceService", () => {
  it("estimates potential new cells only for materializing ops", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "maintenance-estimate" });
    await engine.ready();

    const estimate = Effect.runSync(
      getMaintenanceService(engine).estimatePotentialNewCells([
        { kind: "setCellValue", sheetName: "Sheet1", address: "A1", value: 1 },
        { kind: "setCellFormula", sheetName: "Sheet1", address: "B1", formula: "A1+1" },
        { kind: "setCellFormat", sheetName: "Sheet1", address: "C1", format: "0.00" },
        { kind: "clearCell", sheetName: "Sheet1", address: "D1" },
      ] satisfies EngineOp[]),
    );

    expect(estimate).toBe(3);
  });

  it("rewrites defined names and resets workbook through the extracted maintenance boundary", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "maintenance-reset" });
    await engine.ready();
    engine.createSheet("Source");
    engine.setCellValue("Source", "A1", 10);
    engine.setDefinedName("SourceCell", "=Source!A1");
    engine.setSelection("Source", "B2");

    const maintenance = getMaintenanceService(engine);
    Effect.runSync(maintenance.rewriteDefinedNamesForSheetRename("Source", "Renamed"));

    expect(engine.getDefinedName("SourceCell")).toEqual({
      name: "SourceCell",
      value: "=Renamed!A1",
    });

    const previousBatchId = engine.getLastMetrics().batchId;
    Effect.runSync(maintenance.resetWorkbook("Reset"));

    expect(engine.workbook.workbookName).toBe("Reset");
    expect(engine.getDefinedNames()).toEqual([]);
    expect(engine.getSelectionState()).toEqual({
      sheetName: "Sheet1",
      address: "A1",
      anchorAddress: "A1",
      range: { startAddress: "A1", endAddress: "A1" },
      editMode: "idle",
    });
    expect(engine.getLastMetrics()).toMatchObject({
      batchId: previousBatchId,
      changedInputCount: 0,
      dirtyFormulaCount: 0,
      wasmFormulaCount: 0,
      jsFormulaCount: 0,
    });
  });

  it("captures sheet and range cell state through the extracted maintenance boundary", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "maintenance-capture" });
    await engine.ready();
    engine.createSheet("Sheet1");
    engine.setCellValue("Sheet1", "A1", 10);
    engine.setCellValue("Sheet1", "B2", 20);
    engine.setCellValue("Sheet1", "C3", 30);

    const maintenance = getMaintenanceService(engine);

    expect(
      Effect.runSync(maintenance.captureSheetCellState("Sheet1")).map((op) => op.kind),
    ).toEqual([
      "setCellValue",
      "setCellFormat",
      "setCellValue",
      "setCellFormat",
      "setCellValue",
      "setCellFormat",
    ]);
    expect(
      Effect.runSync(maintenance.captureRowRangeCellState("Sheet1", 0, 2)).map((op) => op.kind),
    ).toEqual(["setCellValue", "setCellFormat", "setCellValue", "setCellFormat"]);
    expect(
      Effect.runSync(maintenance.captureColumnRangeCellState("Sheet1", 0, 2)).map((op) => op.kind),
    ).toEqual(["setCellValue", "setCellFormat", "setCellValue", "setCellFormat"]);
  });

  it("wraps capture callback failures with maintenance service errors", () => {
    const service = createEngineMaintenanceService({
      state: {
        workbook: { listDefinedNames: () => [] },
        formulas: new Map(),
        ranges: { reset: () => undefined },
        entityVersions: new Map(),
        sheetDeleteVersions: new Map(),
        undoStack: [],
        redoStack: [],
        setSelection: () => undefined,
        setSyncState: () => undefined,
        getLastMetrics: () => ({ batchId: "batch-1" }),
        setLastMetrics: () => undefined,
      } as never,
      edgeArena: { reset: () => undefined } as never,
      reverseState: {
        reverseCellEdges: [],
        reverseRangeEdges: [],
        reverseDefinedNameEdges: new Map(),
        reverseTableEdges: new Map(),
        reverseSpillEdges: new Map(),
      },
      pivotOutputOwners: new Map(),
      captureSheetCellState: () => {
        throw new Error("sheet capture boom");
      },
      captureRowRangeCellState: () => {
        throw new Error("row capture boom");
      },
      captureColumnRangeCellState: () => {
        throw new Error("column capture boom");
      },
      setWasmProgramSyncPending: () => undefined,
      setMaterializedCellCount: () => undefined,
      scheduleWasmProgramSync: () => undefined,
      resetWasmState: () => undefined,
    });

    expect(() => Effect.runSync(service.captureSheetCellState("Sheet1"))).toThrow(
      "sheet capture boom",
    );
    expect(() => Effect.runSync(service.captureRowRangeCellState("Sheet1", 0, 1))).toThrow(
      "row capture boom",
    );
    expect(() => Effect.runSync(service.captureColumnRangeCellState("Sheet1", 0, 1))).toThrow(
      "column capture boom",
    );
  });

  it("wraps rename, estimate, and reset failures with maintenance service errors", () => {
    const service = createEngineMaintenanceService({
      state: {
        workbook: {
          listDefinedNames: () => {
            throw new Error("rename boom");
          },
          reset: () => {
            throw new Error("reset boom");
          },
        },
        formulas: new Map(),
        ranges: { reset: () => undefined },
        entityVersions: new Map(),
        sheetDeleteVersions: new Map(),
        undoStack: [],
        redoStack: [],
        setSelection: () => undefined,
        setSyncState: () => undefined,
        getLastMetrics: () => ({ batchId: "batch-1" }),
        setLastMetrics: () => undefined,
      } as never,
      edgeArena: { reset: () => undefined } as never,
      reverseState: {
        reverseCellEdges: [],
        reverseRangeEdges: [],
        reverseDefinedNameEdges: new Map(),
        reverseTableEdges: new Map(),
        reverseSpillEdges: new Map(),
      },
      pivotOutputOwners: new Map(),
      captureSheetCellState: () => [],
      captureRowRangeCellState: () => [],
      captureColumnRangeCellState: () => [],
      setWasmProgramSyncPending: () => undefined,
      setMaterializedCellCount: () => undefined,
      scheduleWasmProgramSync: () => undefined,
      resetWasmState: () => undefined,
    });
    const poisonedOps = {
      get length() {
        throw new Error("estimate boom");
      },
    } as unknown as EngineOp[];

    expect(() =>
      Effect.runSync(service.rewriteDefinedNamesForSheetRename("Sheet1", "Renamed")),
    ).toThrow("rename boom");
    expect(() => Effect.runSync(service.estimatePotentialNewCells(poisonedOps))).toThrow(
      "estimate boom",
    );
    expect(() => Effect.runSync(service.resetWorkbook("Reset"))).toThrow("reset boom");
  });
});
