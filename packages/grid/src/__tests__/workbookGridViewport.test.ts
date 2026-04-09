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
    ).toEqual({
      range: {
        x: 0,
        y: 1,
        width: 6,
        height: 6,
      },
      tx: 0,
      ty: 10,
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
});
