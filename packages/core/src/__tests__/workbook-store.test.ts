import { describe, expect, it } from "vitest";
import { createCellNumberFormatRecord } from "@bilig/protocol";
import { WorkbookStore } from "../workbook-store.js";

describe("WorkbookStore", () => {
  it("does not mutate existing style ranges when bulk style restoration includes an unknown style", () => {
    const workbook = new WorkbookStore("style-ranges");
    workbook.createSheet("Sheet1");
    workbook.upsertCellStyle({ id: "style-a", font: { bold: true } });
    workbook.setStyleRange(
      { sheetName: "Sheet1", startAddress: "A1", endAddress: "B2" },
      "style-a",
    );

    expect(() =>
      workbook.setStyleRanges("Sheet1", [
        {
          range: { sheetName: "Sheet1", startAddress: "C1", endAddress: "C2" },
          styleId: "style-missing",
        },
      ]),
    ).toThrow("Unknown cell style: style-missing");

    expect(workbook.listStyleRanges("Sheet1")).toEqual([
      {
        range: { sheetName: "Sheet1", startAddress: "A1", endAddress: "B2" },
        styleId: "style-a",
      },
    ]);
  });

  it("does not mutate existing format ranges when bulk format restoration includes an unknown format", () => {
    const workbook = new WorkbookStore("format-ranges");
    workbook.createSheet("Sheet1");
    workbook.upsertCellNumberFormat(createCellNumberFormatRecord("format-money", "$0.00"));
    workbook.setFormatRange(
      { sheetName: "Sheet1", startAddress: "A1", endAddress: "B2" },
      "format-money",
    );

    expect(() =>
      workbook.setFormatRanges("Sheet1", [
        {
          range: { sheetName: "Sheet1", startAddress: "C1", endAddress: "C2" },
          formatId: "format-missing",
        },
      ]),
    ).toThrow("Unknown cell number format: format-missing");

    expect(workbook.listFormatRanges("Sheet1")).toEqual([
      {
        range: { sheetName: "Sheet1", startAddress: "A1", endAddress: "B2" },
        formatId: "format-money",
      },
    ]);
  });
});
