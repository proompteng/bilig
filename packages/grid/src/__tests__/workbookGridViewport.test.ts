import { describe, expect, test } from "vitest";
import { getGridMetrics } from "../gridMetrics.js";
import {
  resolveViewportScrollPosition,
  resolveVisibleRegionFromScroll,
} from "../workbookGridViewport.js";

describe("workbookGridViewport", () => {
  test("restores a saved viewport with row height overrides", () => {
    const gridMetrics = getGridMetrics();

    expect(
      resolveViewportScrollPosition({
        viewport: {
          rowStart: 3,
          colStart: 1,
        },
        sortedColumnWidthOverrides: [],
        sortedRowHeightOverrides: [[1, 34]],
        gridMetrics,
      }),
    ).toEqual({
      scrollLeft: gridMetrics.columnWidth,
      scrollTop: gridMetrics.rowHeight + 34 + gridMetrics.rowHeight,
    });
  });

  test("resolves the visible region using row height overrides", () => {
    const gridMetrics = getGridMetrics();

    expect(
      resolveVisibleRegionFromScroll({
        scrollLeft: 0,
        scrollTop: gridMetrics.rowHeight + 10,
        viewportWidth: 480,
        viewportHeight: 140,
        columnWidths: {},
        rowHeights: { 1: 34, 2: 28 },
        gridMetrics,
      }),
    ).toMatchObject({
      range: {
        x: 0,
        y: 1,
        width: 6,
        height: 6,
      },
      tx: 0,
      ty: 10,
      freezeRows: 0,
      freezeCols: 0,
    });
  });

  test("skips collapsed hidden rows when resolving the scroll anchor", () => {
    const gridMetrics = getGridMetrics();

    expect(
      resolveVisibleRegionFromScroll({
        scrollLeft: 0,
        scrollTop: gridMetrics.rowHeight,
        viewportWidth: 320,
        viewportHeight: 140,
        columnWidths: {},
        rowHeights: { 1: 0, 2: 0 },
        gridMetrics,
      }),
    ).toMatchObject({
      range: {
        y: 3,
      },
      ty: 0,
    });
  });

  test("keeps frozen rows and columns pinned when restoring viewport scroll", () => {
    const gridMetrics = getGridMetrics();

    expect(
      resolveViewportScrollPosition({
        viewport: {
          rowStart: 5,
          colStart: 4,
        },
        freezeRows: 2,
        freezeCols: 1,
        sortedColumnWidthOverrides: [
          [0, 120],
          [2, 132],
        ],
        sortedRowHeightOverrides: [
          [0, 30],
          [1, 28],
          [3, 26],
        ],
        gridMetrics,
      }),
    ).toEqual({
      scrollLeft: 104 + 132 + 104,
      scrollTop: 22 + 26 + 22,
    });
  });

  test("resolves the visible region after frozen rows and columns", () => {
    const gridMetrics = getGridMetrics();

    expect(
      resolveVisibleRegionFromScroll({
        scrollLeft: 104,
        scrollTop: 22,
        viewportWidth: 480,
        viewportHeight: 180,
        freezeRows: 1,
        freezeCols: 2,
        columnWidths: {},
        rowHeights: {},
        gridMetrics,
      }),
    ).toMatchObject({
      range: {
        x: 3,
        y: 2,
      },
      tx: 0,
      ty: 0,
      freezeRows: 1,
      freezeCols: 2,
    });
  });
});
