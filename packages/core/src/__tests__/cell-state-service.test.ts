import { Effect } from "effect";
import { describe, expect, it } from "vitest";
import { ValueTag } from "@bilig/protocol";
import { SpreadsheetEngine } from "../engine.js";
import type { EngineCellStateService } from "../engine/services/cell-state-service.js";

function isEngineCellStateService(value: unknown): value is EngineCellStateService {
  if (typeof value !== "object" || value === null) {
    return false;
  }
  return (
    typeof Reflect.get(value, "restoreCellOps") === "function" &&
    typeof Reflect.get(value, "readRangeCells") === "function" &&
    typeof Reflect.get(value, "toCellStateOps") === "function"
  );
}

function getCellStateService(engine: SpreadsheetEngine): EngineCellStateService {
  const runtime = Reflect.get(engine, "runtime");
  if (typeof runtime !== "object" || runtime === null) {
    throw new TypeError("Expected engine runtime");
  }
  const cellState = Reflect.get(runtime, "cellState");
  if (!isEngineCellStateService(cellState)) {
    throw new TypeError("Expected engine cell state service");
  }
  return cellState;
}

describe("EngineCellStateService", () => {
  it("translates relative formulas when materializing target cell ops", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "cell-state-formulas" });
    await engine.ready();
    engine.createSheet("Sheet1");
    engine.setCellValue("Sheet1", "A1", 2);
    engine.setCellValue("Sheet1", "A2", 5);
    engine.setCellFormula("Sheet1", "B1", "A1*2");

    const snapshot = engine.getCell("Sheet1", "B1");
    const ops = Effect.runSync(
      getCellStateService(engine).toCellStateOps(
        "Sheet1",
        "B2",
        snapshot,
        "Sheet1",
        "B1",
      ),
    );

    expect(ops).toContainEqual({
      kind: "setCellFormula",
      sheetName: "Sheet1",
      address: "B2",
      formula: "A2*2",
    });
  });

  it("restores inverse cell ops without duplicating format mutations", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "cell-state-restore" });
    await engine.ready();
    engine.createSheet("Sheet1");
    engine.setCellValue("Sheet1", "A1", 10);
    engine.setCellFormat("Sheet1", "A1", "0.00");

    const ops = Effect.runSync(getCellStateService(engine).restoreCellOps("Sheet1", "A1"));

    expect(ops).toContainEqual({
      kind: "setCellValue",
      sheetName: "Sheet1",
      address: "A1",
      value: 10,
    });
    expect(ops.some((op) => op.kind === "setCellFormat")).toBe(false);
    expect(engine.getCellValue("Sheet1", "A1")).toEqual({ tag: ValueTag.Number, value: 10 });
  });
});
