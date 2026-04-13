import { describe, expect, it } from "vitest";
import { ErrorCode, ValueTag } from "@bilig/protocol";
import { SpreadsheetEngine } from "../engine.js";

describe("direct lookup evaluation", () => {
  it("evaluates exact text and mixed MATCH/XMATCH formulas and refreshes after column writes", async () => {
    const engine = new SpreadsheetEngine({
      workbookName: "evaluation-direct-exact",
      useColumnIndex: true,
    });
    await engine.ready();
    engine.createSheet("Sheet1");

    engine.setCellValue("Sheet1", "A1", "pear");
    engine.setCellValue("Sheet1", "A2", "apple");
    engine.setCellValue("Sheet1", "A3", "pear");
    engine.setCellValue("Sheet1", "B2", "pear");
    engine.setCellValue("Sheet1", "B3", false);
    engine.setCellValue("Sheet1", "D1", "APPLE");
    engine.setCellValue("Sheet1", "D2", false);
    engine.setCellValue("Sheet1", "D3", "pear");

    engine.setCellFormula("Sheet1", "E1", "MATCH(D1,A1:A3,0)");
    engine.setCellFormula("Sheet1", "E2", "MATCH(D2,B1:B3,0)");
    engine.setCellFormula("Sheet1", "E3", "XMATCH(D3,A1:A3,0,-1)");

    expect(engine.getCellValue("Sheet1", "E1")).toEqual({ tag: ValueTag.Number, value: 2 });
    expect(engine.getCellValue("Sheet1", "E2")).toEqual({ tag: ValueTag.Number, value: 3 });
    expect(engine.getCellValue("Sheet1", "E3")).toEqual({ tag: ValueTag.Number, value: 3 });

    engine.setCellValue("Sheet1", "A3", "banana");
    expect(engine.getCellValue("Sheet1", "E3")).toEqual({ tag: ValueTag.Number, value: 1 });

    engine.setCellValue("Sheet1", "D1", false);
    expect(engine.getCellValue("Sheet1", "E1")).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.NA,
    });
  });

  it("evaluates approximate numeric MATCH formulas across uniform, refreshed, and descending columns", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "evaluation-direct-approx" });
    await engine.ready();
    engine.createSheet("Sheet1");

    engine.setCellValue("Sheet1", "A1", 1);
    engine.setCellValue("Sheet1", "A2", 3);
    engine.setCellValue("Sheet1", "A3", 5);
    engine.setCellValue("Sheet1", "B1", 9);
    engine.setCellValue("Sheet1", "B2", 7);
    engine.setCellValue("Sheet1", "B3", 5);
    engine.setCellValue("Sheet1", "D1", 4);
    engine.setCellValue("Sheet1", "D2", 6);

    engine.setCellFormula("Sheet1", "E1", "MATCH(D1,A1:A3,1)");
    engine.setCellFormula("Sheet1", "E2", "MATCH(D2,B1:B3,-1)");

    expect(engine.getCellValue("Sheet1", "E1")).toEqual({ tag: ValueTag.Number, value: 2 });
    expect(engine.getCellValue("Sheet1", "E2")).toEqual({ tag: ValueTag.Number, value: 2 });

    engine.setCellValue("Sheet1", "A2", 4);
    expect(engine.getCellValue("Sheet1", "E1")).toEqual({ tag: ValueTag.Number, value: 2 });

    engine.setCellValue("Sheet1", "D1", 0);
    expect(engine.getCellValue("Sheet1", "E1")).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.NA,
    });

    engine.setCellValue("Sheet1", "D2", 10);
    expect(engine.getCellValue("Sheet1", "E2")).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.NA,
    });
  });

  it("evaluates approximate text MATCH formulas and refreshes after text-column writes", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "evaluation-direct-approx-text" });
    await engine.ready();
    engine.createSheet("Sheet1");

    engine.setCellValue("Sheet1", "C1", "apple");
    engine.setCellValue("Sheet1", "C2", "banana");
    engine.setCellValue("Sheet1", "C3", "pear");
    engine.setCellValue("Sheet1", "D1", "peach");
    engine.setCellValue("Sheet1", "D2", 5);

    engine.setCellFormula("Sheet1", "E1", "MATCH(D1,C1:C3,1)");
    engine.setCellFormula("Sheet1", "E2", "MATCH(D2,C1:C3,1)");

    expect(engine.getCellValue("Sheet1", "E1")).toEqual({ tag: ValueTag.Number, value: 2 });
    expect(engine.getCellValue("Sheet1", "E2")).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.NA,
    });

    engine.setCellValue("Sheet1", "C2", "blueberry");
    expect(engine.getCellValue("Sheet1", "E1")).toEqual({ tag: ValueTag.Number, value: 2 });
  });
});
