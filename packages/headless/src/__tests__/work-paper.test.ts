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
  it("is the canonical top-level alias for the headless workbook runtime", () => {
    const config: WorkPaperConfig = {
      useArrayArithmetic: true,
      useColumnIndex: true,
    };

    const namedExpressions = [{ name: "Threshold", expression: "=2" }] as const;
    const workbook = WorkPaper.buildFromSheets(
      {
        Data: [[1], [2], [3], [4], ["=FILTER(A1:A4,A1:A4>Threshold)"]],
      },
      config,
      namedExpressions,
    );

    const sheetId = workbook.getSheetId("Data")!;
    const spillRange: WorkPaperCellRange = {
      start: cell(sheetId, 4, 0),
      end: cell(sheetId, 5, 0),
    };

    expect(WorkPaper).toBe(HeadlessWorkbook);
    expect(workbook.getRangeValues(spillRange)).toEqual([
      [{ tag: ValueTag.Number, value: 3 }],
      [{ tag: ValueTag.Number, value: 4 }],
    ]);
    expect(workbook.isCellPartOfArray(cell(sheetId, 4, 0))).toBe(true);
    expect(workbook.listNamedExpressions()).toEqual(["Threshold"]);
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
});
