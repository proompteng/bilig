import { describe, expect, test } from "vitest";
import {
  clampSelectionRange,
  createColumnSliceSelection,
  createGridSelection,
  createRangeSelection,
  createRowSliceSelection,
  formatSelectionSummary,
  rectangleToAddresses
} from "../gridSelection.js";

describe("gridSelection", () => {
  test("formats single-cell and rectangular selections", () => {
    expect(formatSelectionSummary(createGridSelection(2, 4), "A1")).toBe("C5");

    const range = createRangeSelection(createGridSelection(1, 1), [1, 1], [3, 4]);
    expect(formatSelectionSummary(range, "A1")).toBe("B2:D5");
  });

  test("formats row and column slice selections", () => {
    expect(formatSelectionSummary(createColumnSliceSelection(1, 3, 0), "A1")).toBe("B:D");
    expect(formatSelectionSummary(createRowSliceSelection(0, 1, 3), "A1")).toBe("2:4");
  });

  test("clamps oversized ranges and converts them to addresses", () => {
    const clamped = clampSelectionRange({ x: -10, y: -20, width: 5, height: 8 });
    expect(clamped).toEqual({ x: 0, y: 0, width: 5, height: 8 });
    expect(rectangleToAddresses({ x: 1, y: 2, width: 3, height: 2 })).toEqual({
      startAddress: "B3",
      endAddress: "D4"
    });
  });
});
