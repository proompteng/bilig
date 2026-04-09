import { Effect } from "effect";
import { describe, expect, it } from "vitest";
import type { EngineOp } from "@bilig/workbook-domain";
import { SpreadsheetEngine } from "../engine.js";
import type { EngineMaintenanceService } from "../engine/services/maintenance-service.js";

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
});
