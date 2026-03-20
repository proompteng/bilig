import { describe, expect, test } from "vitest";
import { createColumnSliceSelection, createGridSelection, createRowSliceSelection } from "../gridSelection.js";
import { resolveActivatedCell, resolveSelectionChange } from "../gridSelectionSync.js";

describe("gridSelectionSync", () => {
  test("prefers pending drag cells when resolving activated cell", () => {
    expect(resolveActivatedCell([4, 4], [2, 5], null)).toEqual([2, 5]);
    expect(resolveActivatedCell([4, 4], null, [3, 6])).toEqual([3, 6]);
    expect(resolveActivatedCell([4, 4], null, null)).toEqual([4, 4]);
  });

  test("keeps row and column header selections anchored to the active sheet axis", () => {
    expect(
      resolveSelectionChange({
        nextSelection: createColumnSliceSelection(2, 4, 7),
        anchorCell: null,
        pointerCell: null,
        selectedCell: [5, 7]
      })
    ).toEqual({
      kind: "column",
      selection: createColumnSliceSelection(2, 4, 7),
      addr: "C8"
    });

    expect(
      resolveSelectionChange({
        nextSelection: createRowSliceSelection(5, 1, 3),
        anchorCell: null,
        pointerCell: null,
        selectedCell: [5, 7]
      })
    ).toEqual({
      kind: "row",
      selection: createRowSliceSelection(5, 1, 3),
      addr: "F2"
    });
  });

  test("clamps and preserves drag-corrected rectangular selections", () => {
    const nextSelection = createGridSelection(3, 3);
    nextSelection.current = {
      ...nextSelection.current,
      cell: [50_000, 2_000_000],
      range: { x: 50_000, y: 2_000_000, width: 5, height: 8 }
    };

    expect(
      resolveSelectionChange({
        nextSelection,
        anchorCell: null,
        pointerCell: null,
        selectedCell: [0, 0]
      })
    ).toEqual({
      kind: "cell",
      selection: {
        ...nextSelection,
        current: {
          ...nextSelection.current,
          cell: [16_383, 1_048_575],
          range: { x: 16_383, y: 1_048_575, width: 1, height: 1 }
        }
      },
      addr: "XFD1048576"
    });

    expect(
      resolveSelectionChange({
        nextSelection: createGridSelection(2, 2),
        anchorCell: [2, 2],
        pointerCell: [4, 5],
        selectedCell: [0, 0]
      })
    ).toEqual({
      kind: "cell",
      selection: {
        ...createGridSelection(2, 2),
        current: {
          cell: [2, 2],
          range: { x: 2, y: 2, width: 3, height: 4 },
          rangeStack: []
        },
        columns: createGridSelection(2, 2).columns,
        rows: createGridSelection(2, 2).rows
      },
      addr: "C3"
    });
  });
});
