import { describe, expect, test } from "vitest";
import { hasSelectionTargetChanged } from "../WorkbookGridSurface.js";
import { getGridMetrics } from "../gridMetrics.js";
import { resolveViewportScrollPosition } from "../workbookGridViewport.js";

describe("WorkbookGridSurface selection autoscroll", () => {
  test("autoscrolls on first selection target", () => {
    expect(
      hasSelectionTargetChanged(null, {
        sheetName: "Sheet1",
        col: 2,
        row: 4,
      }),
    ).toBe(true);
  });

  test("does not autoscroll again when the selection target is unchanged", () => {
    expect(
      hasSelectionTargetChanged(
        {
          sheetName: "Sheet1",
          col: 2,
          row: 4,
        },
        {
          sheetName: "Sheet1",
          col: 2,
          row: 4,
        },
      ),
    ).toBe(false);
  });

  test("autoscrolls when the selected cell changes", () => {
    expect(
      hasSelectionTargetChanged(
        {
          sheetName: "Sheet1",
          col: 2,
          row: 4,
        },
        {
          sheetName: "Sheet1",
          col: 3,
          row: 4,
        },
      ),
    ).toBe(true);
  });

  test("autoscrolls when the active sheet changes", () => {
    expect(
      hasSelectionTargetChanged(
        {
          sheetName: "Sheet1",
          col: 2,
          row: 4,
        },
        {
          sheetName: "Sheet2",
          col: 2,
          row: 4,
        },
      ),
    ).toBe(true);
  });

  test("restores a saved viewport to the recorded top-left cell", () => {
    expect(
      resolveViewportScrollPosition({
        viewport: {
          rowStart: 14,
          colStart: 3,
        },
        sortedColumnWidthOverrides: [],
        gridMetrics: getGridMetrics(),
      }),
    ).toEqual({
      scrollLeft: getGridMetrics().columnWidth * 3,
      scrollTop: getGridMetrics().rowHeight * 14,
    });
  });
});
