import { Effect } from "effect";
import { describe, expect, it } from "vitest";
import { makeCellEntity } from "../entity-ids.js";
import { SpreadsheetEngine } from "../engine.js";
import type { EngineTraversalService } from "../engine/services/traversal-service.js";

function isEngineTraversalService(value: unknown): value is EngineTraversalService {
  if (typeof value !== "object" || value === null) {
    return false;
  }
  return (
    typeof Reflect.get(value, "collectFormulaDependents") === "function" &&
    typeof Reflect.get(value, "forEachFormulaDependencyCell") === "function" &&
    typeof Reflect.get(value, "forEachSheetCell") === "function"
  );
}

function getTraversalService(engine: SpreadsheetEngine): EngineTraversalService {
  const runtime = Reflect.get(engine, "runtime");
  if (typeof runtime !== "object" || runtime === null) {
    throw new TypeError("Expected engine runtime");
  }
  const traversal = Reflect.get(runtime, "traversal");
  if (!isEngineTraversalService(traversal)) {
    throw new TypeError("Expected engine traversal service");
  }
  return traversal;
}

describe("EngineTraversalService", () => {
  it("collects formula dependents beyond the initial traversal scratch capacity", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "traversal-overflow" });
    await engine.ready();
    engine.createSheet("Sheet1");
    engine.setCellValue("Sheet1", "A1", 1);
    for (let row = 1; row <= 140; row += 1) {
      engine.setCellFormula("Sheet1", `B${row}`, "A1+1");
    }
    engine.setCellFormula("Sheet1", "D1", "SUM(A1:B10)");

    const a1Index = engine.workbook.getCellIndex("Sheet1", "A1");
    expect(a1Index).toBeDefined();

    const dependents = Effect.runSync(
      getTraversalService(engine).collectFormulaDependents(makeCellEntity(a1Index!)),
    );
    const dependentAddresses = [...dependents].map((cellIndex) =>
      engine.workbook.getQualifiedAddress(cellIndex),
    );

    expect(dependents.length).toBe(141);
    expect(new Set(dependentAddresses).size).toBe(141);
    expect(dependentAddresses).toContain("Sheet1!B1");
    expect(dependentAddresses).toContain("Sheet1!B140");
    expect(dependentAddresses).toContain("Sheet1!D1");
  });

  it("iterates formula dependencies and sheet cells through the extracted traversal boundary", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "traversal-iteration" });
    await engine.ready();
    engine.createSheet("Sheet1");
    engine.setCellValue("Sheet1", "A1", 3);
    engine.setCellFormula("Sheet1", "B1", "A1+2");
    engine.setCellValue("Sheet1", "C2", 9);

    const traversal = getTraversalService(engine);
    const b1Index = engine.workbook.getCellIndex("Sheet1", "B1");
    const sheetId = engine.workbook.getSheet("Sheet1")?.id;

    expect(b1Index).toBeDefined();
    expect(sheetId).toBeDefined();

    const dependencyAddresses: string[] = [];
    Effect.runSync(
      traversal.forEachFormulaDependencyCell(b1Index!, (dependencyCellIndex) => {
        dependencyAddresses.push(engine.workbook.getQualifiedAddress(dependencyCellIndex));
      }),
    );

    const sheetCells: string[] = [];
    Effect.runSync(
      traversal.forEachSheetCell(sheetId!, (cellIndex, row, col) => {
        sheetCells.push(`${engine.workbook.getQualifiedAddress(cellIndex)}@${row},${col}`);
      }),
    );

    expect(dependencyAddresses).toEqual(["Sheet1!A1"]);
    expect(sheetCells).toContain("Sheet1!A1@0,0");
    expect(sheetCells).toContain("Sheet1!B1@0,1");
    expect(sheetCells).toContain("Sheet1!C2@1,2");
  });
});
