import { describe, expect, it } from "vitest";
import { ValueTag } from "@bilig/protocol";
import { StringPool } from "../string-pool.js";
import { WorkbookStore } from "../workbook-store.js";
import { createExactColumnIndexService } from "../engine/services/exact-column-index-service.js";
import { createEngineRuntimeColumnStoreService } from "../engine/services/runtime-column-store-service.js";

function setStoredCellValue(
  workbook: WorkbookStore,
  strings: StringPool,
  sheetName: string,
  address: string,
  value: { tag: ValueTag.String; value: string } | { tag: ValueTag.Number; value: number },
): void {
  const cellIndex = workbook.ensureCell(sheetName, address);
  workbook.cellStore.setValue(
    cellIndex,
    value,
    value.tag === ValueTag.String ? strings.intern(value.value) : 0,
  );
}

describe("createExactColumnIndexService", () => {
  it("serves exact matches from a primed column index and invalidates by column version", () => {
    const workbook = new WorkbookStore("exact-index");
    const strings = new StringPool();
    workbook.createSheet("Sheet1");

    setStoredCellValue(workbook, strings, "Sheet1", "A1", { tag: ValueTag.String, value: "pear" });
    setStoredCellValue(workbook, strings, "Sheet1", "A2", {
      tag: ValueTag.String,
      value: "apple",
    });
    setStoredCellValue(workbook, strings, "Sheet1", "A3", { tag: ValueTag.String, value: "pear" });

    const runtimeColumnStore = createEngineRuntimeColumnStoreService({
      state: { workbook, strings },
    });
    const exact = createExactColumnIndexService({
      state: { workbook, strings },
      runtimeColumnStore,
    });

    exact.primeColumnIndex({ sheetName: "Sheet1", rowStart: 0, rowEnd: 2, col: 0 });

    expect(
      exact.findVectorMatch({
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

    setStoredCellValue(workbook, strings, "Sheet1", "A1", {
      tag: ValueTag.String,
      value: "banana",
    });

    expect(
      exact.findVectorMatch({
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
    ).toEqual({ handled: true, position: 3 });
  });
});
