import { describe, expect, test } from "vitest";
import {
  isSelectionMoveHandleHit,
  resolveSelectionBounds,
  resolveSelectionMoveAnchorCell,
} from "../gridRangeMove.js";
import type { Rectangle } from "../gridTypes.js";

const CELL_WIDTH = 104;
const CELL_HEIGHT = 22;
const DATA_LEFT = 46;
const DATA_TOP = 24;

function getCellBounds(col: number, row: number): Rectangle {
  return {
    x: DATA_LEFT + col * CELL_WIDTH,
    y: DATA_TOP + row * CELL_HEIGHT,
    width: CELL_WIDTH,
    height: CELL_HEIGHT,
  };
}

describe("gridRangeMove", () => {
  test("resolves the hovered edge cell for selection-border dragging", () => {
    expect(
      resolveSelectionMoveAnchorCell(153, 48, { x: 1, y: 1, width: 2, height: 1 }, getCellBounds),
    ).toEqual([1, 1]);
  });

  test("does not resolve a drag start from the selection interior", () => {
    expect(
      resolveSelectionMoveAnchorCell(202, 57, { x: 1, y: 1, width: 2, height: 2 }, getCellBounds),
    ).toBeNull();
  });

  test("reports selection-border hits through the shared helper", () => {
    expect(
      isSelectionMoveHandleHit(153, 48, { x: 1, y: 1, width: 2, height: 1 }, getCellBounds),
    ).toBe(true);
    expect(
      isSelectionMoveHandleHit(202, 57, { x: 1, y: 1, width: 2, height: 2 }, getCellBounds),
    ).toBe(false);
  });

  test("resolves full selection bounds from the outer cells", () => {
    expect(resolveSelectionBounds({ x: 1, y: 1, width: 2, height: 2 }, getCellBounds)).toEqual({
      x: 150,
      y: 46,
      width: 208,
      height: 44,
    });
  });
});
