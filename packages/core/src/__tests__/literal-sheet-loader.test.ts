import { describe, expect, it } from "vitest";
import { ValueTag } from "@bilig/protocol";
import { SpreadsheetEngine } from "../engine.js";
import { loadLiteralSheetIntoEmptySheet } from "../literal-sheet-loader.js";

describe("loadLiteralSheetIntoEmptySheet", () => {
  it("hydrates literal cells into a fresh sheet without mutation ops", () => {
    const engine = new SpreadsheetEngine({ workbookName: "literal-load" });
    engine.workbook.createSheet("Sheet1");
    const sheetId = engine.workbook.getSheet("Sheet1")!.id;

    const loaded = loadLiteralSheetIntoEmptySheet(engine.workbook, engine.strings, sheetId, [
      [1, "two", true],
      [null, 4, false],
    ]);

    expect(loaded).toBe(5);
    expect(engine.getCellValue("Sheet1", "A1")).toEqual({ tag: ValueTag.Number, value: 1 });
    expect(engine.getCellValue("Sheet1", "B1")).toEqual({
      tag: ValueTag.String,
      value: "two",
      stringId: expect.any(Number),
    });
    expect(engine.getCellValue("Sheet1", "C1")).toEqual({ tag: ValueTag.Boolean, value: true });
    expect(engine.getCellValue("Sheet1", "A2")).toEqual({ tag: ValueTag.Empty });
    expect(engine.getCellValue("Sheet1", "B2")).toEqual({ tag: ValueTag.Number, value: 4 });
    expect(engine.workbook.getSheet("Sheet1")!.columnVersions[0]).toBe(1);
    expect(engine.workbook.getSheet("Sheet1")!.columnVersions[1]).toBe(2);
    expect(engine.workbook.getSheet("Sheet1")!.columnVersions[2]).toBe(2);
  });

  it("keeps loaded literals usable as formula inputs", () => {
    const engine = new SpreadsheetEngine({ workbookName: "literal-formula-load" });
    engine.workbook.createSheet("Sheet1");
    const sheetId = engine.workbook.getSheet("Sheet1")!.id;

    loadLiteralSheetIntoEmptySheet(engine.workbook, engine.strings, sheetId, [[2], [5]]);
    engine.setCellFormula("Sheet1", "B1", "SUM(A1:A2)");

    expect(engine.getCellValue("Sheet1", "B1")).toEqual({ tag: ValueTag.Number, value: 7 });
  });

  it("can skip formula-like strings while bulk-loading the literal subset", () => {
    const engine = new SpreadsheetEngine({ workbookName: "literal-filtered-load" });
    engine.workbook.createSheet("Sheet1");
    const sheetId = engine.workbook.getSheet("Sheet1")!.id;

    const loaded = loadLiteralSheetIntoEmptySheet(
      engine.workbook,
      engine.strings,
      sheetId,
      [
        [1, "=A1+1", "ok"],
        [null, "=B1+1", false],
      ],
      (raw) => !(typeof raw === "string" && raw.startsWith("=")),
    );

    expect(loaded).toBe(4);
    expect(engine.getCellValue("Sheet1", "A1")).toEqual({ tag: ValueTag.Number, value: 1 });
    expect(engine.getCellValue("Sheet1", "B1")).toEqual({ tag: ValueTag.Empty });
    expect(engine.getCellValue("Sheet1", "C1")).toEqual({
      tag: ValueTag.String,
      value: "ok",
      stringId: expect.any(Number),
    });
    expect(engine.getCellValue("Sheet1", "B2")).toEqual({ tag: ValueTag.Empty });
    expect(engine.getCellValue("Sheet1", "C2")).toEqual({ tag: ValueTag.Boolean, value: false });
  });
});
