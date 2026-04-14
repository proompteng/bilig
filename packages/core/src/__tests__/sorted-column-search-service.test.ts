import { describe, expect, it, vi } from "vitest";
import { ErrorCode, ValueTag } from "@bilig/protocol";
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

function setStoredString(
  workbook: WorkbookStore,
  strings: StringPool,
  address: string,
  value: string,
): void {
  const cellIndex = workbook.ensureCell("Sheet1", address);
  workbook.cellStore.setValue(cellIndex, { tag: ValueTag.String, value }, strings.intern(value));
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

  it("updates cached sorted columns incrementally for monotonic literal writes without rematerializing the column", () => {
    const workbook = new WorkbookStore("sorted-index-incremental");
    const strings = new StringPool();
    workbook.createSheet("Sheet1");

    [1, 3, 5, 7].forEach((value, index) => {
      setStoredNumber(workbook, strings, `A${index + 1}`, value);
    });

    const runtimeColumnStore = createEngineRuntimeColumnStoreService({
      state: { workbook, strings },
    });
    const getColumnSliceSpy = vi.spyOn(runtimeColumnStore, "getColumnSlice");
    const sorted = createSortedColumnSearchService({
      state: { workbook, strings },
      runtimeColumnStore,
    });

    const prepared = sorted.prepareVectorLookup({
      sheetName: "Sheet1",
      rowStart: 0,
      rowEnd: 3,
      col: 0,
    });
    const sliceCallsAfterPrepare = getColumnSliceSpy.mock.calls.length;

    setStoredNumber(workbook, strings, "A4", 9);
    sorted.recordLiteralWrite({
      sheetName: "Sheet1",
      row: 3,
      col: 0,
      oldValue: { tag: ValueTag.Number, value: 7 },
      newValue: { tag: ValueTag.Number, value: 9 },
    });

    expect(
      sorted.findPreparedVectorMatch({
        lookupValue: { tag: ValueTag.Number, value: 8 },
        prepared,
        matchMode: 1,
      }),
    ).toEqual({ handled: true, position: 3 });
    expect(getColumnSliceSpy.mock.calls.length).toBe(sliceCallsAfterPrepare);
  });

  it("refreshes prepared lookups after structural remaps and rejects unsupported lookup shapes", () => {
    const workbook = new WorkbookStore("sorted-index-prepared");
    const strings = new StringPool();
    workbook.createSheet("Sheet1");

    [1, 3, 5].forEach((value, index) => {
      setStoredNumber(workbook, strings, `A${index + 1}`, value);
      setStoredNumber(workbook, strings, `B${index + 1}`, 7 - index * 2);
    });
    ["apple", "banana", "pear"].forEach((value, index) => {
      setStoredString(workbook, strings, `C${index + 1}`, value);
    });
    setStoredNumber(workbook, strings, "D1", 1);
    setStoredString(workbook, strings, "D2", "mixed");
    setStoredNumber(workbook, strings, "D3", 3);

    const runtimeColumnStore = createEngineRuntimeColumnStoreService({
      state: { workbook, strings },
    });
    const sorted = createSortedColumnSearchService({
      state: { workbook, strings },
      runtimeColumnStore,
    });

    const ascendingPrepared = sorted.prepareVectorLookup({
      sheetName: "Sheet1",
      rowStart: 0,
      rowEnd: 2,
      col: 0,
    });
    expect(
      sorted.findPreparedVectorMatch({
        lookupValue: { tag: ValueTag.Boolean, value: true },
        prepared: ascendingPrepared,
        matchMode: 1,
      }),
    ).toEqual({ handled: true, position: 1 });
    expect(
      sorted.findVectorMatch({
        lookupValue: { tag: ValueTag.Empty, value: null },
        sheetName: "Sheet1",
        start: "A1",
        end: "A3",
        matchMode: 1,
      }),
    ).toEqual({ handled: true, position: undefined });
    expect(
      sorted.findVectorMatch({
        lookupValue: { tag: ValueTag.Boolean, value: true },
        sheetName: "Sheet1",
        start: "A1",
        end: "A3",
        matchMode: 1,
      }),
    ).toEqual({ handled: true, position: 1 });
    expect(
      sorted.findPreparedVectorMatch({
        lookupValue: { tag: ValueTag.String, value: "bad" },
        prepared: ascendingPrepared,
        matchMode: 1,
      }),
    ).toEqual({ handled: false });

    const descendingPrepared = sorted.prepareVectorLookup({
      sheetName: "Sheet1",
      rowStart: 0,
      rowEnd: 2,
      col: 1,
    });
    const textPrepared = sorted.prepareVectorLookup({
      sheetName: "Sheet1",
      rowStart: 0,
      rowEnd: 2,
      col: 2,
    });
    expect(
      sorted.findPreparedVectorMatch({
        lookupValue: { tag: ValueTag.String, value: "peach" },
        prepared: textPrepared,
        matchMode: 1,
      }),
    ).toEqual({ handled: true, position: 2 });
    expect(
      sorted.findVectorMatch({
        lookupValue: { tag: ValueTag.Empty, value: null },
        sheetName: "Sheet1",
        start: "C1",
        end: "C3",
        matchMode: 1,
      }),
    ).toEqual({ handled: true, position: undefined });
    expect(
      sorted.findPreparedVectorMatch({
        lookupValue: { tag: ValueTag.Error, code: ErrorCode.Div0 },
        prepared: textPrepared,
        matchMode: 1,
      }),
    ).toEqual({ handled: false });
    expect(
      sorted.findVectorMatch({
        lookupValue: { tag: ValueTag.Number, value: 6 },
        sheetName: "Sheet1",
        start: "B1",
        end: "B3",
        matchMode: -1,
      }),
    ).toEqual({ handled: true, position: 1 });
    expect(
      sorted.findVectorMatch({
        lookupValue: { tag: ValueTag.Error, code: ErrorCode.Div0 },
        sheetName: "Sheet1",
        start: "C1",
        end: "C3",
        matchMode: 1,
      }),
    ).toEqual({ handled: false });

    workbook.deleteRows("Sheet1", 0, 1);
    workbook.remapSheetCells("Sheet1", "row", (index) => (index === 0 ? undefined : index - 1));

    expect(
      sorted.findPreparedVectorMatch({
        lookupValue: { tag: ValueTag.Empty, value: null },
        prepared: descendingPrepared,
        matchMode: -1,
      }),
    ).toEqual({ handled: true, position: 3 });
    expect(
      sorted.findPreparedVectorMatch({
        lookupValue: { tag: ValueTag.Empty, value: null },
        prepared: descendingPrepared,
        matchMode: -1,
      }),
    ).toEqual({ handled: true, position: 3 });
    expect(
      sorted.findPreparedVectorMatch({
        lookupValue: { tag: ValueTag.Number, value: 10 },
        prepared: descendingPrepared,
        matchMode: 1,
      }),
    ).toEqual({ handled: false });

    expect(
      sorted.findVectorMatch({
        lookupValue: { tag: ValueTag.Number, value: 2 },
        sheetName: "Sheet1",
        start: "D1",
        end: "D3",
        matchMode: 1,
      }),
    ).toEqual({ handled: false });
    expect(
      sorted.findVectorMatch({
        lookupValue: { tag: ValueTag.String, value: "pear" },
        sheetName: "Sheet1",
        start: "A1",
        end: "B2",
        matchMode: 1,
      }),
    ).toEqual({ handled: false });
  });
});
