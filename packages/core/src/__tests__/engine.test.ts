import { describe, expect, it } from "vitest";
import { SpreadsheetEngine } from "../index.js";
import { ValueTag } from "@bilig/protocol";

describe("SpreadsheetEngine", () => {
  it("recalculates simple formulas", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "spec" });
    await engine.ready();
    engine.createSheet("Sheet1");
    engine.setCellValue("Sheet1", "A1", 10);
    engine.setCellFormula("Sheet1", "B1", "A1*2");

    expect(engine.getCellValue("Sheet1", "B1")).toEqual({ tag: ValueTag.Number, value: 20 });

    engine.setCellValue("Sheet1", "A1", 12);
    expect(engine.getCellValue("Sheet1", "B1")).toEqual({ tag: ValueTag.Number, value: 24 });
  });

  it("supports cross-sheet references", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "spec" });
    await engine.ready();
    engine.createSheet("Sheet1");
    engine.createSheet("Sheet2");
    engine.setCellValue("Sheet1", "A1", 4);
    engine.setCellFormula("Sheet2", "B2", "Sheet1!A1*3");
    expect(engine.getCellValue("Sheet2", "B2")).toEqual({ tag: ValueTag.Number, value: 12 });
  });
});
