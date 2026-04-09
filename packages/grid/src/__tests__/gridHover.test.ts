import { describe, expect, test } from "vitest";
import { getGridMetrics } from "../gridMetrics.js";
import { createPointerGeometry, type VisibleRegionState } from "../gridPointer.js";
import { resolveGridHoverState, sameGridHoverState } from "../gridHover.js";

const gridMetrics = getGridMetrics();
const region: VisibleRegionState = {
  range: { x: 0, y: 0, width: 4, height: 4 },
  tx: 0,
  ty: 0,
};

function createGeometry() {
  return createPointerGeometry(
    { left: 0, top: 0, right: 480, bottom: 240 },
    region,
    {},
    {},
    gridMetrics,
  );
}

describe("gridHover", () => {
  test("resolves column resize hotspots before generic header hover", () => {
    const geometry = createGeometry();
    const state = resolveGridHoverState({
      clientX: 148,
      clientY: 12,
      region,
      geometry,
      columnWidths: {},
      rowHeights: {},
      defaultColumnWidth: gridMetrics.columnWidth,
      defaultRowHeight: gridMetrics.rowHeight,
      gridMetrics,
      selectedCell: [0, 0],
      selectedCellBounds: null,
      selectionRange: null,
      hasColumnSelection: false,
      hasRowSelection: false,
    });

    expect(state).toEqual({
      cell: null,
      header: { kind: "column", index: 0 },
      cursor: "col-resize",
    });
  });

  test("resolves header hover with pointer cursor", () => {
    const geometry = createGeometry();
    const state = resolveGridHoverState({
      clientX: 200,
      clientY: 12,
      region,
      geometry,
      columnWidths: {},
      rowHeights: {},
      defaultColumnWidth: gridMetrics.columnWidth,
      defaultRowHeight: gridMetrics.rowHeight,
      gridMetrics,
      selectedCell: [0, 0],
      selectedCellBounds: null,
      selectionRange: null,
      hasColumnSelection: false,
      hasRowSelection: false,
    });

    expect(state).toEqual({
      cell: null,
      header: { kind: "column", index: 1 },
      cursor: "pointer",
    });
  });

  test("resolves body cell hover with cell cursor", () => {
    const geometry = createGeometry();
    const state = resolveGridHoverState({
      clientX: 220,
      clientY: 60,
      region,
      geometry,
      columnWidths: {},
      rowHeights: {},
      defaultColumnWidth: gridMetrics.columnWidth,
      defaultRowHeight: gridMetrics.rowHeight,
      gridMetrics,
      selectedCell: [0, 0],
      selectedCellBounds: null,
      selectionRange: null,
      hasColumnSelection: false,
      hasRowSelection: false,
    });

    expect(state).toEqual({
      cell: [1, 1],
      header: null,
      cursor: "cell",
    });
  });

  test("compares hover state structurally", () => {
    expect(
      sameGridHoverState(
        { cell: [1, 2], header: null, cursor: "cell" },
        { cell: [1, 2], header: null, cursor: "cell" },
      ),
    ).toBe(true);
    expect(
      sameGridHoverState(
        { cell: [1, 2], header: null, cursor: "cell" },
        { cell: [1, 3], header: null, cursor: "cell" },
      ),
    ).toBe(false);
  });
});
