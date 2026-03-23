import { describe, expect, test } from "vitest";
import { PRODUCT_COLUMN_WIDTH, getGridMetrics } from "../gridMetrics.js";
import {
  createPointerGeometry,
  resolveHeaderSelection,
  resolveHeaderSelectionForDrag,
  resolvePointerCell,
  type PointerGeometry,
  type VisibleRegionState,
} from "../gridPointer.js";

const gridMetrics = getGridMetrics("product");
const region: VisibleRegionState = {
  range: { x: 0, y: 0, width: 12, height: 24 },
  tx: 0,
  ty: 0,
};

function buildGeometry(): PointerGeometry {
  return createPointerGeometry(
    { left: 0, top: 33, right: 1068, bottom: 868 },
    region,
    {},
    gridMetrics,
  );
}

describe("gridPointer", () => {
  test("uses the product row contract for geometry", () => {
    const geometry = buildGeometry();
    expect(geometry.cellHeight).toBe(22);
    expect(geometry.dataTop).toBe(57);
    expect(geometry.dataLeft).toBe(46);
  });

  test("maps clicks in the upper half of a visible cell to that same cell", () => {
    const geometry = buildGeometry();
    const cell = resolvePointerCell({
      clientX: 46 + 4 * PRODUCT_COLUMN_WIDTH + Math.floor(PRODUCT_COLUMN_WIDTH / 2),
      clientY: 57 + 11 * 22 + 4,
      region,
      geometry,
      columnWidths: {},
      gridMetrics,
      selectedCell: [0, 0],
      selectedCellBounds: null,
      selectionRange: { x: 0, y: 0, width: 1, height: 1 },
      hasColumnSelection: false,
      hasRowSelection: false,
    });

    expect(cell).toEqual([4, 11]);
  });

  test("keeps the active single-cell selection when clicking its visible top border", () => {
    const geometry = buildGeometry();
    const cell = resolvePointerCell({
      clientX: 46 + 2 * PRODUCT_COLUMN_WIDTH + Math.floor(PRODUCT_COLUMN_WIDTH / 2),
      clientY: 57 + 4 * 22 - 1,
      region,
      geometry,
      columnWidths: {},
      gridMetrics,
      selectedCell: [2, 4],
      selectedCellBounds: {
        x: 46 + 2 * PRODUCT_COLUMN_WIDTH,
        y: 57 + 4 * 22,
        width: PRODUCT_COLUMN_WIDTH,
        height: 22,
      },
      selectionRange: { x: 2, y: 4, width: 1, height: 1 },
      hasColumnSelection: false,
      hasRowSelection: false,
    });

    expect(cell).toEqual([2, 4]);
  });

  test("resolves header selections and drag targets", () => {
    const geometry = buildGeometry();

    expect(
      resolveHeaderSelection(46 + PRODUCT_COLUMN_WIDTH + 10, 40, region, geometry, {}, gridMetrics),
    ).toEqual({
      kind: "column",
      index: 1,
    });

    expect(resolveHeaderSelection(20, 57 + 22 + 3, region, geometry, {}, gridMetrics)).toEqual({
      kind: "row",
      index: 1,
    });

    expect(
      resolveHeaderSelectionForDrag(
        "column",
        46 + 3 * PRODUCT_COLUMN_WIDTH + 10,
        70,
        region,
        geometry,
        {},
        gridMetrics,
      ),
    ).toEqual({
      kind: "column",
      index: 3,
    });

    expect(
      resolveHeaderSelectionForDrag("row", 20, 57 + 5 * 22 + 8, region, geometry, {}, gridMetrics),
    ).toEqual({
      kind: "row",
      index: 5,
    });
  });
});
