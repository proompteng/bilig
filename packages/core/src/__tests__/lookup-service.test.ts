import { describe, expect, it } from "vitest";
import { ErrorCode, ValueTag, type CellValue } from "@bilig/protocol";
import { StringPool } from "../string-pool.js";
import { WorkbookStore } from "../workbook-store.js";
import { createEngineLookupService } from "../engine/services/lookup-service.js";

function setStoredCellValue(
  workbook: WorkbookStore,
  strings: StringPool,
  sheetName: string,
  address: string,
  value: CellValue,
): void {
  const cellIndex = workbook.ensureCell(sheetName, address);
  workbook.cellStore.setValue(
    cellIndex,
    value,
    value.tag === ValueTag.String ? strings.intern(value.value) : 0,
  );
}

describe("createEngineLookupService", () => {
  it("reuses primed exact column indexes and invalidates them when the column changes", () => {
    const workbook = new WorkbookStore("lookup-service");
    const strings = new StringPool();
    workbook.createSheet("Sheet1");

    setStoredCellValue(workbook, strings, "Sheet1", "A1", { tag: ValueTag.String, value: "pear" });
    setStoredCellValue(workbook, strings, "Sheet1", "A2", {
      tag: ValueTag.String,
      value: "apple",
    });
    setStoredCellValue(workbook, strings, "Sheet1", "A3", { tag: ValueTag.String, value: "pear" });

    const lookup = createEngineLookupService({
      state: {
        workbook,
        strings,
      },
    });

    lookup.primeExactColumnIndex({
      sheetName: "Sheet1",
      rowStart: 0,
      rowEnd: 2,
      col: 0,
    });

    expect(
      lookup.findExactVectorMatch({
        lookupValue: { tag: ValueTag.String, value: "PEAR" },
        sheetName: "Sheet1",
        start: "A1",
        end: "A3",
        startRow: 0,
        endRow: 2,
        startCol: 0,
        endCol: 0,
        searchMode: 1,
      }),
    ).toEqual({ handled: true, position: 1 });

    expect(
      lookup.findExactVectorMatch({
        lookupValue: { tag: ValueTag.String, value: "pear" },
        sheetName: "Sheet1",
        start: "A1",
        end: "A3",
        startRow: 0,
        endRow: 2,
        startCol: 0,
        endCol: 0,
        searchMode: -1,
      }),
    ).toEqual({ handled: true, position: 3 });

    setStoredCellValue(workbook, strings, "Sheet1", "A3", {
      tag: ValueTag.String,
      value: "banana",
    });

    expect(
      lookup.findExactVectorMatch({
        lookupValue: { tag: ValueTag.String, value: "pear" },
        sheetName: "Sheet1",
        start: "A1",
        end: "A3",
        startRow: 0,
        endRow: 2,
        startCol: 0,
        endCol: 0,
        searchMode: -1,
      }),
    ).toEqual({ handled: true, position: 1 });
  });

  it("handles parsed ranges and exact lookup edge cases across empty, boolean, number, and error cells", () => {
    const workbook = new WorkbookStore("lookup-service-edges");
    const strings = new StringPool();
    workbook.createSheet("Sheet1");

    workbook.ensureCell("Sheet1", "A2");
    setStoredCellValue(workbook, strings, "Sheet1", "A3", {
      tag: ValueTag.Empty,
      value: null,
    });
    setStoredCellValue(workbook, strings, "Sheet1", "A4", {
      tag: ValueTag.Boolean,
      value: false,
    });
    setStoredCellValue(workbook, strings, "Sheet1", "A5", {
      tag: ValueTag.Number,
      value: -0,
    });
    setStoredCellValue(workbook, strings, "Sheet1", "A6", {
      tag: ValueTag.Error,
      code: ErrorCode.Div0,
    });

    const lookup = createEngineLookupService({
      state: {
        workbook,
        strings,
      },
    });

    expect(
      lookup.findExactVectorMatch({
        lookupValue: { tag: ValueTag.Empty, value: null },
        sheetName: "Sheet1",
        start: "A1",
        end: "A5",
        searchMode: 1,
      }),
    ).toEqual({ handled: true, position: 1 });

    expect(
      lookup.findExactVectorMatch({
        lookupValue: { tag: ValueTag.Boolean, value: false },
        sheetName: "Sheet1",
        start: "A1",
        end: "A5",
        searchMode: 1,
      }),
    ).toEqual({ handled: true, position: 4 });

    expect(
      lookup.findExactVectorMatch({
        lookupValue: { tag: ValueTag.Number, value: -0 },
        sheetName: "Sheet1",
        start: "A1",
        end: "A6",
        searchMode: 1,
      }),
    ).toEqual({ handled: true, position: 5 });

    expect(
      lookup.findExactVectorMatch({
        lookupValue: { tag: ValueTag.Empty, value: null },
        sheetName: "Missing",
        start: "A1",
        end: "A3",
        searchMode: 1,
      }),
    ).toEqual({ handled: true, position: 1 });

    expect(
      lookup.findExactVectorMatch({
        lookupValue: { tag: ValueTag.Error, code: ErrorCode.Name },
        sheetName: "Sheet1",
        start: "A1",
        end: "A6",
        searchMode: 1,
      }),
    ).toEqual({ handled: false });

    expect(
      lookup.findExactVectorMatch({
        lookupValue: { tag: ValueTag.String, value: "pear" },
        sheetName: "Sheet1",
        start: "A1",
        end: "B2",
        searchMode: 1,
      }),
    ).toEqual({ handled: false });
  });
});
