import { describe, expect, it } from "vitest";
import {
  readViewportColumnWidths,
  readViewportHiddenColumns,
  readViewportHiddenRows,
  readViewportRowHeights,
} from "../worker-workbook-view-state.js";

describe("worker workbook view state", () => {
  it("returns stable empty viewport snapshots before the worker handle exists", () => {
    expect(readViewportColumnWidths(undefined, "Sheet1")).toBe(
      readViewportColumnWidths(undefined, "Sheet1"),
    );
    expect(readViewportRowHeights(undefined, "Sheet1")).toBe(
      readViewportRowHeights(undefined, "Sheet1"),
    );
    expect(readViewportHiddenColumns(undefined, "Sheet1")).toBe(
      readViewportHiddenColumns(undefined, "Sheet1"),
    );
    expect(readViewportHiddenRows(undefined, "Sheet1")).toBe(
      readViewportHiddenRows(undefined, "Sheet1"),
    );
  });

  it("passes through live viewport store snapshots when available", () => {
    const columnWidths = { 2: 128 } as const;
    const rowHeights = { 4: 36 } as const;
    const hiddenColumns = { 7: true } as const;
    const hiddenRows = { 9: true } as const;
    const workerHandle = {
      viewportStore: {
        getColumnWidths: () => columnWidths,
        getRowHeights: () => rowHeights,
        getHiddenColumns: () => hiddenColumns,
        getHiddenRows: () => hiddenRows,
      },
    };

    expect(readViewportColumnWidths(workerHandle, "Sheet1")).toBe(columnWidths);
    expect(readViewportRowHeights(workerHandle, "Sheet1")).toBe(rowHeights);
    expect(readViewportHiddenColumns(workerHandle, "Sheet1")).toBe(hiddenColumns);
    expect(readViewportHiddenRows(workerHandle, "Sheet1")).toBe(hiddenRows);
  });
});
