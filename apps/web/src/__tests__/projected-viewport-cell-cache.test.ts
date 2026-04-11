import { describe, expect, it, vi } from "vitest";
import { ValueTag, type CellSnapshot } from "@bilig/protocol";
import { ProjectedViewportCellCache } from "../projected-viewport-cell-cache.js";

function countSheetCells(cache: ProjectedViewportCellCache, sheetName: string): number {
  let count = 0;
  cache.getSheet(sheetName)?.grid.forEachCellEntry(() => {
    count += 1;
  });
  return count;
}

function snapshot(address: string, value: number | string): CellSnapshot {
  return {
    sheetName: "Sheet1",
    address,
    value:
      typeof value === "number"
        ? { tag: ValueTag.Number, value }
        : { tag: ValueTag.String, value, stringId: 1 },
    flags: 0,
    version: 1,
  };
}

function columnLabel(columnIndex: number): string {
  let index = columnIndex + 1;
  let label = "";
  while (index > 0) {
    const remainder = (index - 1) % 26;
    label = String.fromCharCode(65 + remainder) + label;
    index = Math.floor((index - 1) / 26);
  }
  return label;
}

describe("ProjectedViewportCellCache", () => {
  it("tracks cell subscriptions and exposes sheet grid entries", () => {
    const cache = new ProjectedViewportCellCache();
    const listener = vi.fn();
    cache.subscribeCells("Sheet1", ["B2"], listener);

    cache.setCellSnapshot(snapshot("B2", "left"));

    expect(listener).toHaveBeenCalledTimes(1);
    expect(cache.getCell("Sheet1", "B2")).toMatchObject({
      value: { tag: ValueTag.String, value: "left" },
    });
    expect(countSheetCells(cache, "Sheet1")).toBe(1);
  });

  it("prunes to the cache cap after the last viewport unsubscribes while keeping pinned cells", () => {
    const cache = new ProjectedViewportCellCache({ maxCachedCellsPerSheet: 6000 });
    const untrackViewport = cache.trackViewport("Sheet1", {
      rowStart: 0,
      rowEnd: 600,
      colStart: 0,
      colEnd: 9,
    });
    const unsubscribeCell = cache.subscribeCells("Sheet1", ["A1"], () => undefined);

    for (let row = 0; row <= 600; row += 1) {
      for (let col = 0; col < 10; col += 1) {
        cache.setCellSnapshot(snapshot(`${columnLabel(col)}${row + 1}`, row * 10 + col));
      }
    }

    expect(countSheetCells(cache, "Sheet1")).toBe(6010);

    untrackViewport();

    expect(countSheetCells(cache, "Sheet1")).toBe(6000);
    expect(cache.peekCell("Sheet1", "A1")).toBeDefined();

    unsubscribeCell();
  });
});
