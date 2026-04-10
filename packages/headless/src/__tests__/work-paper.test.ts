import { describe, expect, it } from "vitest";
import { ValueTag } from "@bilig/protocol";

import {
  HeadlessWorkbook,
  WorkPaper,
  type WorkPaperCellAddress,
  type WorkPaperCellRange,
  type WorkPaperConfig,
} from "../index.js";

function cell(sheet: number, row: number, col: number): WorkPaperCellAddress {
  return { sheet, row, col };
}

describe("WorkPaper", () => {
  it("matches the published README usage example", () => {
    const workbook = WorkPaper.buildFromSheets({
      Sheet1: [[1, "=A1*2"]],
    });

    const sheetId = workbook.getSheetId("Sheet1")!;
    const value = workbook.getCellValue({ sheet: sheetId, row: 0, col: 1 });

    expect(value).toEqual({
      tag: ValueTag.Number,
      value: 2,
    });
  });

  it("is the canonical top-level alias for the headless workbook runtime", () => {
    const config: WorkPaperConfig = {
      useArrayArithmetic: true,
      useColumnIndex: true,
    };

    const workbook = WorkPaper.buildFromSheets(
      {
        Data: [[2, "=FILTER(A2:A5,A2:A5>A1)"], [1], [2], [3], [4]],
      },
      config,
    );

    const sheetId = workbook.getSheetId("Data")!;
    const spillRange: WorkPaperCellRange = {
      start: cell(sheetId, 0, 1),
      end: cell(sheetId, 1, 1),
    };

    expect(WorkPaper).toBe(HeadlessWorkbook);
    expect(workbook.getRangeValues(spillRange)).toEqual([
      [{ tag: ValueTag.Number, value: 3 }],
      [{ tag: ValueTag.Number, value: 4 }],
    ]);
    expect(workbook.isCellPartOfArray(cell(sheetId, 0, 1))).toBe(true);
  });

  it("supports a production-style headless workflow through the WorkPaper entrypoint", () => {
    const workbook = WorkPaper.buildFromSheets({
      Revenue: [[125], [250], [375], ["=SUM(A1:A3)"]],
    });
    const revenueId = workbook.getSheetId("Revenue")!;

    const copied = workbook.copy({
      start: cell(revenueId, 0, 0),
      end: cell(revenueId, 2, 0),
    });

    expect(copied).toEqual([
      [{ tag: ValueTag.Number, value: 125 }],
      [{ tag: ValueTag.Number, value: 250 }],
      [{ tag: ValueTag.Number, value: 375 }],
    ]);

    const changes = workbook.batch(() => {
      workbook.paste(cell(revenueId, 0, 1));
      workbook.setCellContents(cell(revenueId, 3, 1), "=SUM(B1:B3)");
    });

    expect(changes.length).toBeGreaterThan(0);
    expect(workbook.getCellValue(cell(revenueId, 3, 1))).toEqual({
      tag: ValueTag.Number,
      value: 750,
    });

    workbook.undo();
    expect(workbook.getCellSerialized(cell(revenueId, 3, 1))).toBeNull();
    workbook.redo();
    expect(workbook.getCellFormula(cell(revenueId, 3, 1))).toBe("=SUM(B1:B3)");
  });

  it("keeps compatibility adapters frozen and returns detached adapter results", () => {
    const workbook = WorkPaper.buildFromSheets({
      Data: [[1, "=A1*2"]],
    });
    const dataId = workbook.getSheetId("Data")!;

    expect(Object.isFrozen(workbook.internals)).toBe(true);
    expect(Object.isFrozen(workbook.graph)).toBe(true);
    expect(Object.isFrozen(workbook.rangeMapping)).toBe(true);
    expect(Object.isFrozen(workbook.arrayMapping)).toBe(true);
    expect(Object.isFrozen(workbook.sheetMapping)).toBe(true);
    expect(Object.isFrozen(workbook.addressMapping)).toBe(true);
    expect(Object.isFrozen(workbook.dependencyGraph)).toBe(true);
    expect(Object.isFrozen(workbook.evaluator)).toBe(true);
    expect(Object.isFrozen(workbook.columnSearch)).toBe(true);
    expect(Object.isFrozen(workbook.lazilyTransformingAstService)).toBe(true);

    const serialized = workbook.rangeMapping.getSerialized({
      start: cell(dataId, 0, 0),
      end: cell(dataId, 0, 1),
    });
    serialized[0][0] = 999;
    serialized[0][1] = "=A1*999";

    expect(workbook.getCellSerialized(cell(dataId, 0, 0))).toBe(1);
    expect(workbook.getCellSerialized(cell(dataId, 0, 1))).toBe("=A1*2");
  });
});
