import { describe, expect, it } from "vitest";
import { SheetGrid } from "../sheet-grid.js";

describe("SheetGrid", () => {
  it("stores, retrieves, and clears sparse cells across blocks", () => {
    const grid = new SheetGrid();
    grid.set(0, 0, 4);
    grid.set(130, 33, 9);

    expect(grid.get(0, 0)).toBe(4);
    expect(grid.get(130, 33)).toBe(9);
    expect(grid.get(1, 1)).toBe(-1);

    grid.clear(0, 0);
    expect(grid.get(0, 0)).toBe(-1);
  });

  it("iterates visible cells by range and across all occupied blocks", () => {
    const grid = new SheetGrid();
    grid.set(0, 0, 1);
    grid.set(0, 2, 2);
    grid.set(200, 40, 3);

    const inRange: number[] = [];
    grid.forEachInRange(0, 0, 1, 2, (cellIndex) => inRange.push(cellIndex));
    expect(inRange).toEqual([1, 2]);

    const all: number[] = [];
    grid.forEachCell((cellIndex) => all.push(cellIndex));
    expect(all.toSorted((left, right) => left - right)).toEqual([1, 2, 3]);
  });
});
