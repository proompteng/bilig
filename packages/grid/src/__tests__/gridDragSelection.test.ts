import { describe, expect, test } from "vitest";
import { createColumnSliceSelection, createGridSelection, createRowSliceSelection } from "../gridSelection.js";
import {
  resolveBodyDragSelection,
  resolveBodyPointerUpResult,
  resolveHeaderDragSelection
} from "../gridDragSelection.js";

describe("gridDragSelection", () => {
  test("resolves header drag selections into row and column slices", () => {
    expect(resolveHeaderDragSelection({ kind: "column", index: 1 }, 3, [4, 7])).toEqual({
      selection: createColumnSliceSelection(1, 3, 7),
      addr: "B8"
    });

    expect(resolveHeaderDragSelection({ kind: "row", index: 2 }, 5, [4, 7])).toEqual({
      selection: createRowSliceSelection(4, 2, 5),
      addr: "E3"
    });
  });

  test("builds rectangular range selections for body drags", () => {
    expect(resolveBodyDragSelection([2, 2], [4, 5])).toEqual({
      ...createGridSelection(2, 2),
      current: {
        cell: [2, 2],
        range: { x: 2, y: 2, width: 3, height: 4 },
        rangeStack: []
      }
    });
  });

  test("distinguishes drag completion from single body clicks", () => {
    expect(resolveBodyPointerUpResult([2, 2], [4, 5], true)).toEqual({
      addr: "C3",
      clickedCell: null,
      selection: {
        ...createGridSelection(2, 2),
        current: {
          cell: [2, 2],
          range: { x: 2, y: 2, width: 3, height: 4 },
          rangeStack: []
        }
      },
      shouldSetDragExpiry: true
    });

    expect(resolveBodyPointerUpResult([2, 2], [4, 5], false)).toEqual({
      addr: null,
      clickedCell: [2, 2],
      selection: null,
      shouldSetDragExpiry: false
    });
  });
});
