import { describe, expect, it, vi } from "vitest";
import { ErrorCode, ValueTag, type CellValue } from "@bilig/protocol";
import { createRangeAggregateCacheService } from "../engine/services/range-aggregate-cache-service.js";
import type { RuntimeColumnSlice } from "../engine/services/runtime-column-store-service.js";
import { WorkbookStore } from "../workbook-store.js";

function makeSlice(
  values: readonly CellValue[],
  request: { sheetName: string; rowStart: number; rowEnd: number; col: number },
  columnVersion: number,
  structureVersion: number,
): RuntimeColumnSlice {
  const length = request.rowEnd - request.rowStart + 1;
  const tags = new Uint8Array(length);
  const numbers = new Float64Array(length);
  const stringIds = new Uint32Array(length);
  const errors = new Uint16Array(length);
  for (let offset = 0; offset < length; offset += 1) {
    const value = values[request.rowStart + offset] ?? { tag: ValueTag.Empty };
    tags[offset] = value.tag;
    if (value.tag === ValueTag.Number) {
      numbers[offset] = value.value;
    } else if (value.tag === ValueTag.Boolean) {
      numbers[offset] = value.value ? 1 : 0;
    } else if (value.tag === ValueTag.Error) {
      errors[offset] = value.code;
    }
  }
  return {
    sheetName: request.sheetName,
    rowStart: request.rowStart,
    rowEnd: request.rowEnd,
    col: request.col,
    length,
    columnVersion,
    structureVersion,
    sheetColumnVersions: Uint32Array.of(columnVersion),
    tags,
    numbers,
    stringIds,
    errors,
  };
}

describe("RangeAggregateCacheService", () => {
  it("extends cached prefixes incrementally instead of rebuilding from the start row", () => {
    const values: CellValue[] = [
      { tag: ValueTag.Number, value: 2 },
      { tag: ValueTag.Boolean, value: true },
      { tag: ValueTag.Empty },
      { tag: ValueTag.Error, code: ErrorCode.NA },
    ];
    const getColumnSlice = vi.fn(
      (request: { sheetName: string; rowStart: number; rowEnd: number; col: number }) =>
        makeSlice(values, request, 7, 3),
    );
    const workbook = new WorkbookStore("aggregate-cache");
    workbook.createSheet("Sheet1");
    const sheet = workbook.getSheet("Sheet1");
    if (!sheet) {
      throw new Error("expected Sheet1 to exist");
    }
    sheet.columnVersions = Uint32Array.of(7);
    sheet.structureVersion = 3;
    const service = createRangeAggregateCacheService({
      state: {
        workbook,
      },
      runtimeColumnStore: {
        getColumnSlice,
        readCellValue: () => ({ tag: ValueTag.Empty }),
        readRangeValues: () => [],
        normalizeStringId: () => "",
        normalizeLookupText: () => "",
      },
    });

    const first = service.getOrBuildPrefix({
      sheetName: "Sheet1",
      rowStart: 0,
      rowEnd: 1,
      col: 0,
    });
    expect(getColumnSlice).toHaveBeenNthCalledWith(1, {
      sheetName: "Sheet1",
      rowStart: 0,
      rowEnd: 1,
      col: 0,
    });
    expect(first.prefixSums).toEqual(Float64Array.from([2, 3]));
    expect(first.prefixCount).toEqual(Uint32Array.from([1, 2]));
    expect(first.prefixAverageCount).toEqual(Uint32Array.from([1, 2]));
    expect(first.prefixErrorCodes).toEqual(Uint16Array.from([0, 0]));

    const extended = service.getOrBuildPrefix({
      sheetName: "Sheet1",
      rowStart: 0,
      rowEnd: 3,
      col: 0,
    });
    expect(getColumnSlice).toHaveBeenNthCalledWith(2, {
      sheetName: "Sheet1",
      rowStart: 2,
      rowEnd: 3,
      col: 0,
    });
    expect(extended.prefixSums).toEqual(Float64Array.from([2, 3, 3, 3]));
    expect(extended.prefixCount).toEqual(Uint32Array.from([1, 2, 2, 2]));
    expect(extended.prefixAverageCount).toEqual(Uint32Array.from([1, 2, 3, 3]));
    expect(extended.prefixErrorCodes).toEqual(Uint16Array.from([0, 0, 0, ErrorCode.NA]));

    const reused = service.getOrBuildPrefix({
      sheetName: "Sheet1",
      rowStart: 0,
      rowEnd: 2,
      col: 0,
    });
    expect(reused).toBe(extended);
    expect(getColumnSlice).toHaveBeenCalledTimes(2);
  });

  it("rebuilds cached prefixes when the column version changes", () => {
    const values: CellValue[] = [{ tag: ValueTag.Number, value: 4 }];
    let columnVersion = 5;
    const getColumnSlice = vi.fn(
      (request: { sheetName: string; rowStart: number; rowEnd: number; col: number }) =>
        makeSlice(values, request, columnVersion, 1),
    );
    const workbook = new WorkbookStore("aggregate-cache-rebuild");
    workbook.createSheet("Sheet1");
    const sheet = workbook.getSheet("Sheet1");
    if (!sheet) {
      throw new Error("expected Sheet1 to exist");
    }
    sheet.columnVersions = Uint32Array.of(columnVersion);
    sheet.structureVersion = 1;
    const service = createRangeAggregateCacheService({
      state: { workbook },
      runtimeColumnStore: {
        getColumnSlice,
        readCellValue: () => ({ tag: ValueTag.Empty }),
        readRangeValues: () => [],
        normalizeStringId: () => "",
        normalizeLookupText: () => "",
      },
    });

    const first = service.getOrBuildPrefix({ sheetName: "Sheet1", rowStart: 0, rowEnd: 0, col: 0 });
    expect(first.prefixSums[0]).toBe(4);

    columnVersion = 6;
    sheet.columnVersions[0] = columnVersion;
    const rebuilt = service.getOrBuildPrefix({
      sheetName: "Sheet1",
      rowStart: 0,
      rowEnd: 0,
      col: 0,
    });

    expect(rebuilt).not.toBe(first);
    expect(getColumnSlice).toHaveBeenCalledTimes(2);
  });

  it("extends prefixes across number and boolean deltas", () => {
    const values: CellValue[] = [
      { tag: ValueTag.Number, value: 1 },
      { tag: ValueTag.Boolean, value: true },
      { tag: ValueTag.Number, value: 2 },
    ];
    const getColumnSlice = vi.fn(
      (request: { sheetName: string; rowStart: number; rowEnd: number; col: number }) =>
        makeSlice(values, request, 9, 2),
    );
    const workbook = new WorkbookStore("aggregate-cache-extend");
    workbook.createSheet("Sheet1");
    const sheet = workbook.getSheet("Sheet1");
    if (!sheet) {
      throw new Error("expected Sheet1 to exist");
    }
    sheet.columnVersions = Uint32Array.of(9);
    sheet.structureVersion = 2;
    const service = createRangeAggregateCacheService({
      state: { workbook },
      runtimeColumnStore: {
        getColumnSlice,
        readCellValue: () => ({ tag: ValueTag.Empty }),
        readRangeValues: () => [],
        normalizeStringId: () => "",
        normalizeLookupText: () => "",
      },
    });

    service.getOrBuildPrefix({ sheetName: "Sheet1", rowStart: 0, rowEnd: 0, col: 0 });
    const extended = service.getOrBuildPrefix({
      sheetName: "Sheet1",
      rowStart: 0,
      rowEnd: 2,
      col: 0,
    });

    expect(extended.prefixSums).toEqual(Float64Array.from([1, 2, 4]));
    expect(extended.prefixCount).toEqual(Uint32Array.from([1, 2, 3]));
  });
});
