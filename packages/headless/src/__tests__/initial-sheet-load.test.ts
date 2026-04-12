import { describe, expect, it, vi } from "vitest";
import { SpreadsheetEngine } from "@bilig/core";
import { ValueTag } from "@bilig/protocol";
import { WorkPaper } from "../index.js";

describe("initial mixed sheet load", () => {
  it("builds mixed sheets without routing formulas through restore cell mutations", () => {
    const restoreMutationSpy = vi.spyOn(
      SpreadsheetEngine.prototype,
      "applyCellMutationsAtWithOptions",
    );
    try {
      const workbook = WorkPaper.buildFromSheets({
        Bench: [
          [1, "=A1*2"],
          [2, "=A2*3"],
        ],
      });
      const sheetId = workbook.getSheetId("Bench")!;

      expect(workbook.getCellValue({ sheet: sheetId, row: 0, col: 1 })).toEqual({
        tag: ValueTag.Number,
        value: 2,
      });
      expect(workbook.getCellValue({ sheet: sheetId, row: 1, col: 1 })).toEqual({
        tag: ValueTag.Number,
        value: 6,
      });
      expect(restoreMutationSpy).not.toHaveBeenCalled();
    } finally {
      restoreMutationSpy.mockRestore();
    }
  });
});
