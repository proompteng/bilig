import { Effect } from "effect";
import { describe, expect, it } from "vitest";
import { ValueTag } from "@bilig/protocol";
import { SpreadsheetEngine } from "../engine.js";
import type { EngineFormulaBindingService } from "../engine/services/formula-binding-service.js";

function isEngineFormulaBindingService(value: unknown): value is EngineFormulaBindingService {
  if (typeof value !== "object" || value === null) {
    return false;
  }
  return (
    typeof Reflect.get(value, "bindFormula") === "function" &&
    typeof Reflect.get(value, "clearFormula") === "function" &&
    typeof Reflect.get(value, "rewriteCellFormulasForSheetRename") === "function"
  );
}

function getBindingService(engine: SpreadsheetEngine): EngineFormulaBindingService {
  const runtime = Reflect.get(engine, "runtime");
  if (typeof runtime !== "object" || runtime === null) {
    throw new TypeError("Expected engine runtime");
  }
  const binding = Reflect.get(runtime, "binding");
  if (!isEngineFormulaBindingService(binding)) {
    throw new TypeError("Expected engine formula binding service");
  }
  return binding;
}

describe("EngineFormulaBindingService", () => {
  it("clears reverse dependency edges when a formula is removed through the service", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "binding-clear" });
    await engine.ready();
    engine.createSheet("Sheet1");
    engine.setCellValue("Sheet1", "A1", 7);
    engine.setCellFormula("Sheet1", "B1", "A1*2");

    const formulaCellIndex = engine.workbook.getCellIndex("Sheet1", "B1");
    expect(formulaCellIndex).toBeDefined();
    expect(engine.getDependencies("Sheet1", "A1").directDependents).toContain("Sheet1!B1");

    Effect.runSync(getBindingService(engine).clearFormula(formulaCellIndex!));

    expect(engine.getCell("Sheet1", "B1").formula).toBeUndefined();
    expect(engine.getDependencies("Sheet1", "A1").directDependents).toEqual([]);
  });

  it("rewrites quoted sheet references on rename through the binding service", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "binding-rename" });
    await engine.ready();
    engine.createSheet("Q1's Data");
    engine.createSheet("Summary");
    engine.setCellValue("Q1's Data", "A1", 7);
    engine.setCellFormula("Summary", "A1", "'Q1''s Data'!A1*2");

    const renamed = engine.workbook.renameSheet("Q1's Data", "Q2's Data");
    expect(renamed).toBeTruthy();

    Effect.runSync(
      getBindingService(engine).rewriteCellFormulasForSheetRename("Q1's Data", "Q2's Data", 0),
    );

    expect(engine.getCell("Summary", "A1").formula).toBe("'Q2''s Data'!A1*2");
    expect(engine.getCellValue("Summary", "A1")).toEqual({ tag: ValueTag.Number, value: 14 });
  });

  it("binds repeated row-translated formulas through the service without changing results", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "binding-row-template" });
    await engine.ready();
    engine.createSheet("Sheet1");

    engine.setCellValue("Sheet1", "A1", 1);
    engine.setCellValue("Sheet1", "B1", 10);
    engine.setCellValue("Sheet1", "A2", 2);
    engine.setCellValue("Sheet1", "B2", 20);

    engine.setCellFormula("Sheet1", "E1", "A1+B1");
    engine.setCellFormula("Sheet1", "F1", "E1*2");
    engine.setCellFormula("Sheet1", "E2", "A2+B2");
    engine.setCellFormula("Sheet1", "F2", "E2*2");

    expect(engine.getCellValue("Sheet1", "E1")).toEqual({ tag: ValueTag.Number, value: 11 });
    expect(engine.getCellValue("Sheet1", "F1")).toEqual({ tag: ValueTag.Number, value: 22 });
    expect(engine.getCellValue("Sheet1", "E2")).toEqual({ tag: ValueTag.Number, value: 22 });
    expect(engine.getCellValue("Sheet1", "F2")).toEqual({ tag: ValueTag.Number, value: 44 });
  });

  it("preserves dependency wiring across formula rewrites with the same dependencies", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "binding-same-deps-rewrite" });
    await engine.ready();
    engine.createSheet("Sheet1");

    engine.setCellValue("Sheet1", "A1", 2);
    engine.setCellValue("Sheet1", "B1", 3);
    engine.setCellFormula("Sheet1", "C1", "A1+B1");
    engine.setCellFormula("Sheet1", "D1", "C1*2");

    engine.setCellFormula("Sheet1", "C1", "A1*B1");

    expect(engine.getCellValue("Sheet1", "C1")).toEqual({ tag: ValueTag.Number, value: 6 });
    expect(engine.getCellValue("Sheet1", "D1")).toEqual({ tag: ValueTag.Number, value: 12 });
    expect(engine.getDependencies("Sheet1", "A1").directDependents).toContain("Sheet1!C1");
    expect(engine.getDependencies("Sheet1", "C1").directDependents).toContain("Sheet1!D1");
  });
});
