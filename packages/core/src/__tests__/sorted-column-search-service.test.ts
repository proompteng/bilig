import { describe, expect, it } from "vitest";
import { ValueTag } from "@bilig/protocol";
import { StringPool } from "../string-pool.js";
import { WorkbookStore } from "../workbook-store.js";
import { createSortedColumnSearchService } from "../engine/services/sorted-column-search-service.js";
import { createEngineRuntimeColumnStoreService } from "../engine/services/runtime-column-store-service.js";

function setStoredNumber(
  workbook: WorkbookStore,
  strings: StringPool,
  address: string,
  value: number,
): void {
  const cellIndex = workbook.ensureCell("Sheet1", address);
  workbook.cellStore.setValue(cellIndex, { tag: ValueTag.Number, value }, 0);
  void strings;
}

describe("createSortedColumnSearchService", () => {
  it("serves approximate matches from a primed sorted column and invalidates by column version", () => {
    const workbook = new WorkbookStore("sorted-index");
    const strings = new StringPool();
    workbook.createSheet("Sheet1");

    [1, 3, 5, 7].forEach((value, index) => {
      setStoredNumber(workbook, strings, `A${index + 1}`, value);
    });

    const runtimeColumnStore = createEngineRuntimeColumnStoreService({
      state: { workbook, strings },
    });
    const sorted = createSortedColumnSearchService({
      state: { workbook, strings },
      runtimeColumnStore,
    });

    sorted.primeColumnIndex({ sheetName: "Sheet1", rowStart: 0, rowEnd: 3, col: 0 });

    expect(
      sorted.findVectorMatch({
        lookupValue: { tag: ValueTag.Number, value: 6 },
        sheetName: "Sheet1",
        start: "A1",
        end: "A4",
        startRow: 0,
        endRow: 3,
        startCol: 0,
        endCol: 0,
        matchMode: 1,
      }),
    ).toEqual({ handled: true, position: 3 });

    setStoredNumber(workbook, strings, "A3", 6);

    expect(
      sorted.findVectorMatch({
        lookupValue: { tag: ValueTag.Number, value: 6 },
        sheetName: "Sheet1",
        start: "A1",
        end: "A4",
        startRow: 0,
        endRow: 3,
        startCol: 0,
        endCol: 0,
        matchMode: 1,
      }),
    ).toEqual({ handled: true, position: 3 });
  });
});
