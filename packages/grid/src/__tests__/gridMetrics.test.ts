import { describe, expect, test } from "vitest";
import {
  PRODUCT_COLUMN_WIDTH,
  PRODUCT_HEADER_HEIGHT,
  PRODUCT_ROW_HEIGHT,
  PRODUCT_ROW_MARKER_WIDTH,
  getGridMetrics,
  getVisibleColumnBounds,
  resolveColumnAtClientX,
  getResolvedColumnWidth,
} from "../gridMetrics.js";

describe("gridMetrics", () => {
  test("returns the product grid contract", () => {
    expect(getGridMetrics("product")).toEqual({
      columnWidth: PRODUCT_COLUMN_WIDTH,
      rowHeight: PRODUCT_ROW_HEIGHT,
      headerHeight: PRODUCT_HEADER_HEIGHT,
      rowMarkerWidth: PRODUCT_ROW_MARKER_WIDTH,
    });
  });

  test("resolves visible column bounds and pointer columns with overrides", () => {
    const columnWidths = { 1: 140, 2: 88 };
    const bounds = getVisibleColumnBounds(
      { x: 0, width: 4 },
      46,
      16384,
      columnWidths,
      PRODUCT_COLUMN_WIDTH,
    );

    expect(
      bounds.map((column) => ({ index: column.index, left: column.left, width: column.width })),
    ).toEqual([
      { index: 0, left: 46, width: PRODUCT_COLUMN_WIDTH },
      { index: 1, left: 46 + PRODUCT_COLUMN_WIDTH, width: 140 },
      { index: 2, left: 46 + PRODUCT_COLUMN_WIDTH + 140, width: 88 },
      { index: 3, left: 46 + PRODUCT_COLUMN_WIDTH + 140 + 88, width: PRODUCT_COLUMN_WIDTH },
    ]);

    expect(
      resolveColumnAtClientX(
        46 + 20,
        { x: 0, width: 4 },
        46,
        16384,
        columnWidths,
        PRODUCT_COLUMN_WIDTH,
      ),
    ).toBe(0);
    expect(
      resolveColumnAtClientX(
        46 + PRODUCT_COLUMN_WIDTH + 30,
        { x: 0, width: 4 },
        46,
        16384,
        columnWidths,
        PRODUCT_COLUMN_WIDTH,
      ),
    ).toBe(1);
    expect(
      resolveColumnAtClientX(
        46 + PRODUCT_COLUMN_WIDTH + 140 + 10,
        { x: 0, width: 4 },
        46,
        16384,
        columnWidths,
        PRODUCT_COLUMN_WIDTH,
      ),
    ).toBe(2);
    expect(getResolvedColumnWidth(columnWidths, 3, PRODUCT_COLUMN_WIDTH)).toBe(
      PRODUCT_COLUMN_WIDTH,
    );
  });
});
