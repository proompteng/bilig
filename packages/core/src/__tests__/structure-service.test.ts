import { Effect } from "effect";
import { describe, expect, it } from "vitest";
import { SpreadsheetEngine } from "../engine.js";
import type { EngineStructureService } from "../engine/services/structure-service.js";

function isEngineStructureService(value: unknown): value is EngineStructureService {
  if (typeof value !== "object" || value === null) {
    return false;
  }
  return (
    typeof Reflect.get(value, "captureSheetCellState") === "function" &&
    typeof Reflect.get(value, "captureRowRangeCellState") === "function" &&
    typeof Reflect.get(value, "captureColumnRangeCellState") === "function" &&
    typeof Reflect.get(value, "applyStructuralAxisOp") === "function"
  );
}

function getStructureService(engine: SpreadsheetEngine): EngineStructureService {
  const runtime = Reflect.get(engine, "runtime");
  if (typeof runtime !== "object" || runtime === null) {
    throw new TypeError("Expected engine runtime");
  }
  const structure = Reflect.get(runtime, "structure");
  if (!isEngineStructureService(structure)) {
    throw new TypeError("Expected engine structure service");
  }
  return structure;
}

describe("EngineStructureService", () => {
  it("captures sheet cell state in row-major order for undo reconstruction", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "structure-capture" });
    await engine.ready();
    engine.createSheet("Sheet1");
    engine.setCellValue("Sheet1", "B2", 20);
    engine.setCellFormat("Sheet1", "B2", "0.00");
    engine.setCellFormula("Sheet1", "A3", "B2*2");

    const ops = Effect.runSync(getStructureService(engine).captureSheetCellState("Sheet1"));

    expect(ops).toEqual([
      { kind: "setCellValue", sheetName: "Sheet1", address: "B2", value: 20 },
      { kind: "setCellFormat", sheetName: "Sheet1", address: "B2", format: "0.00" },
      { kind: "setCellFormula", sheetName: "Sheet1", address: "A3", formula: "B2*2" },
    ]);
  });

  it("rewrites metadata-backed ranges and formula bindings across row inserts", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "structure-rewrite" });
    await engine.ready();
    engine.createSheet("Data");
    engine.createSheet("Pivot");
    engine.setRangeValues({ sheetName: "Data", startAddress: "A1", endAddress: "B4" }, [
      ["Region", "Sales"],
      ["East", 10],
      ["West", 7],
      ["East", 5],
    ]);
    engine.setDefinedName("SalesRange", "=Data!A1:B4");
    engine.setCellFormula("Pivot", "E2", "SUM(Data!B1:B4)");
    engine.setFreezePane("Data", 1, 0);
    engine.setFilter("Data", { sheetName: "Data", startAddress: "A1", endAddress: "B4" });
    engine.setSort("Data", { sheetName: "Data", startAddress: "A1", endAddress: "B4" }, [
      { keyAddress: "B1", direction: "asc" },
    ]);
    engine.setTable({
      name: "Sales",
      sheetName: "Data",
      startAddress: "A1",
      endAddress: "B4",
      columnNames: ["Region", "Sales"],
      headerRow: true,
      totalsRow: false,
    });
    engine.setPivotTable("Pivot", "B2", {
      name: "SalesPivot",
      source: { sheetName: "Data", startAddress: "A1", endAddress: "B4" },
      groupBy: ["Region"],
      values: [{ sourceColumn: "Sales", summarizeBy: "sum" }],
    });

    const result = Effect.runSync(
      getStructureService(engine).applyStructuralAxisOp({
        kind: "insertRows",
        sheetName: "Data",
        start: 0,
        count: 1,
      }),
    );

    expect(result.changedCellIndices.length).toBeGreaterThan(0);
    expect(engine.getDefinedName("SalesRange")).toEqual({
      name: "SalesRange",
      value: "=Data!A2:B5",
    });
    expect(engine.getCell("Pivot", "E2").formula).toBe("SUM(Data!B2:B5)");
    expect(engine.getFreezePane("Data")).toEqual({ sheetName: "Data", rows: 2, cols: 0 });
    expect(engine.getFilters("Data")).toEqual([
      { sheetName: "Data", range: { sheetName: "Data", startAddress: "A2", endAddress: "B5" } },
    ]);
    expect(engine.getSorts("Data")).toEqual([
      {
        sheetName: "Data",
        range: { sheetName: "Data", startAddress: "A2", endAddress: "B5" },
        keys: [{ keyAddress: "B2", direction: "asc" }],
      },
    ]);
    expect(engine.getTables()).toEqual([
      {
        name: "Sales",
        sheetName: "Data",
        startAddress: "A2",
        endAddress: "B5",
        columnNames: ["Region", "Sales"],
        headerRow: true,
        totalsRow: false,
      },
    ]);
    expect(engine.getPivotTable("Pivot", "B2")?.source).toEqual({
      sheetName: "Data",
      startAddress: "A2",
      endAddress: "B5",
    });
  });

  it("rewrites range-backed defined names across column inserts", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "structure-defined-range-rewrite" });
    await engine.ready();
    engine.createSheet("Data");
    engine.setRangeValues({ sheetName: "Data", startAddress: "A1", endAddress: "B3" }, [
      ["Qty", "Amount"],
      [1, 10],
      [2, 20],
    ]);
    engine.setDefinedName("SalesRange", {
      kind: "range-ref",
      sheetName: "Data",
      startAddress: "A1",
      endAddress: "B3",
    });
    engine.setCellFormula("Data", "C1", "SUM(SalesRange)");

    Effect.runSync(
      getStructureService(engine).applyStructuralAxisOp({
        kind: "insertColumns",
        sheetName: "Data",
        start: 0,
        count: 1,
      }),
    );

    expect(engine.getDefinedName("SalesRange")).toEqual({
      name: "SalesRange",
      value: {
        kind: "range-ref",
        sheetName: "Data",
        startAddress: "B1",
        endAddress: "C3",
      },
    });
    expect(engine.getCell("Data", "D1").formula).toBe("SUM(SalesRange)");
    expect(engine.getCellValue("Data", "D1")).toMatchObject({ tag: 1, value: 33 });
  });
});
