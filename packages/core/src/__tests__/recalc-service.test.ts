import { Effect } from "effect";
import { describe, expect, it, vi } from "vitest";
import { FormulaMode, ValueTag } from "@bilig/protocol";
import { utcDateToExcelSerial } from "@bilig/formula";
import { SpreadsheetEngine } from "../engine.js";
import type { EngineRecalcService } from "../engine/services/recalc-service.js";
import type { RuntimeFormula } from "../engine/runtime-state.js";

function isEngineRecalcService(value: unknown): value is EngineRecalcService {
  if (typeof value !== "object" || value === null) {
    return false;
  }
  return (
    typeof Reflect.get(value, "recalculateNow") === "function" &&
    typeof Reflect.get(value, "recalculateDirty") === "function" &&
    typeof Reflect.get(value, "recalculateDifferential") === "function"
  );
}

function getRecalcService(engine: SpreadsheetEngine): EngineRecalcService {
  const runtime = Reflect.get(engine, "runtime");
  if (typeof runtime !== "object" || runtime === null) {
    throw new TypeError("Expected engine runtime");
  }
  const recalc = Reflect.get(runtime, "recalc");
  if (!isEngineRecalcService(recalc)) {
    throw new TypeError("Expected engine recalc service");
  }
  return recalc;
}

function isFormulaTable(
  value: unknown,
): value is { get(cellIndex: number): RuntimeFormula | undefined } {
  return typeof value === "object" && value !== null && typeof Reflect.get(value, "get") === "function";
}

function getFormulaTable(engine: SpreadsheetEngine): { get(cellIndex: number): RuntimeFormula | undefined } {
  const formulas = Reflect.get(engine, "formulas");
  if (!isFormulaTable(formulas)) {
    throw new TypeError("Expected engine formula table");
  }
  return formulas;
}

describe("EngineRecalcService", () => {
  it("performs dirty-region recalculation through the extracted service boundary", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "recalc-dirty" });
    await engine.ready();
    engine.createSheet("Sheet1");

    engine.setCellValue("Sheet1", "A1", 10);
    engine.setCellFormula("Sheet1", "B1", "A1*2");
    engine.setCellValue("Sheet1", "C1", 5);
    engine.setCellFormula("Sheet1", "D1", "C1+10");

    const a1Index = engine.workbook.ensureCell("Sheet1", "A1");
    const c1Index = engine.workbook.ensureCell("Sheet1", "C1");
    engine.workbook.cellStore.setValue(a1Index, { tag: ValueTag.Number, value: 50 });
    engine.workbook.cellStore.setValue(c1Index, { tag: ValueTag.Number, value: 100 });

    const changed = Effect.runSync(
      getRecalcService(engine).recalculateDirty([
        { sheetName: "Sheet1", rowStart: 0, rowEnd: 0, colStart: 0, colEnd: 0 },
      ]),
    );

    const b1Index = engine.workbook.getCellIndex("Sheet1", "B1");
    const d1Index = engine.workbook.getCellIndex("Sheet1", "D1");
    expect(changed).toContain(b1Index);
    expect(changed).not.toContain(d1Index);
    expect(engine.getCellValue("Sheet1", "B1")).toEqual({ tag: ValueTag.Number, value: 100 });
    expect(engine.getCellValue("Sheet1", "D1")).toEqual({ tag: ValueTag.Number, value: 15 });
  });

  it("recalculates all formulas through the extracted service boundary", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "recalc-now" });
    await engine.ready();
    engine.createSheet("Sheet1");
    engine.setCellValue("Sheet1", "A1", 10);
    engine.setCellFormula("Sheet1", "B1", "A1*2");
    const b1Index = engine.workbook.getCellIndex("Sheet1", "B1");
    const formula = b1Index === undefined ? undefined : getFormulaTable(engine).get(b1Index);
    if (!formula) {
      throw new TypeError("Expected B1 formula");
    }
    formula.compiled.mode = FormulaMode.JsOnly;

    const a1Index = engine.workbook.ensureCell("Sheet1", "A1");
    engine.workbook.cellStore.setValue(a1Index, { tag: ValueTag.Number, value: 25 });

    const changed = Effect.runSync(getRecalcService(engine).recalculateNow());

    expect(changed).toContain(b1Index);
    expect(engine.getCellValue("Sheet1", "B1")).toEqual({ tag: ValueTag.Number, value: 50 });
  });

  it("recalculates volatile formulas from the current clock and random inputs", async () => {
    vi.useFakeTimers();
    const randomSpy = vi.spyOn(Math, "random");
    vi.setSystemTime(new Date("2026-02-03T00:00:00Z"));
    randomSpy.mockReturnValue(0.125);

    const engine = new SpreadsheetEngine({ workbookName: "recalc-volatile" });
    await engine.ready();
    engine.createSheet("Sheet1");
    engine.setCellFormula("Sheet1", "A1", "TODAY()");
    engine.setCellFormula("Sheet1", "B1", "RAND()");

    vi.setSystemTime(new Date("2026-02-04T00:00:00Z"));
    randomSpy.mockReturnValue(0.75);

    const changed = Effect.runSync(getRecalcService(engine).recalculateNow());

    expect(changed).toContain(engine.workbook.getCellIndex("Sheet1", "A1"));
    expect(changed).toContain(engine.workbook.getCellIndex("Sheet1", "B1"));
    expect(engine.getCellValue("Sheet1", "A1")).toEqual({
      tag: ValueTag.Number,
      value: utcDateToExcelSerial(new Date("2026-02-04T00:00:00Z")),
    });
    expect(engine.getCellValue("Sheet1", "B1")).toEqual({
      tag: ValueTag.Number,
      value: 0.75,
    });
  });
});
