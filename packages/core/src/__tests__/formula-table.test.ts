import { describe, expect, it } from "vitest";
import { CellStore } from "../cell-store.js";
import { FormulaTable } from "../formula-table.js";

describe("FormulaTable", () => {
  it("stores formulas in a compact vector and writes real formula ids into the cell store", () => {
    const store = new CellStore();
    const firstCell = store.allocate(0, 0, 0);
    const secondCell = store.allocate(0, 0, 5);
    const formulas = new FormulaTable<{ cellIndex: number; source: string }>(store);

    const firstId = formulas.set(secondCell, { cellIndex: secondCell, source: "A1*2" });
    const secondId = formulas.set(firstCell, { cellIndex: firstCell, source: "B1*3" });

    expect(firstId).toBe(1);
    expect(secondId).toBe(2);
    expect(store.formulaIds[secondCell]).toBe(1);
    expect(store.formulaIds[firstCell]).toBe(2);
    expect([...formulas.keys()]).toEqual([secondCell, firstCell]);
  });

  it("reuses freed formula slots instead of growing forever", () => {
    const store = new CellStore();
    const a = store.allocate(0, 0, 0);
    const b = store.allocate(0, 0, 1);
    const c = store.allocate(0, 0, 2);
    const formulas = new FormulaTable<{ cellIndex: number; source: string }>(store);

    formulas.set(a, { cellIndex: a, source: "1" });
    formulas.set(b, { cellIndex: b, source: "2" });
    const removed = formulas.delete(a);
    const reusedId = formulas.set(c, { cellIndex: c, source: "3" });

    expect(removed?.source).toBe("1");
    expect(reusedId).toBe(1);
    expect(store.formulaIds[a]).toBe(0);
    expect(store.formulaIds[c]).toBe(1);
    expect(formulas.size).toBe(2);
    expect([...formulas.values()].map((formula) => formula.source)).toEqual(["3", "2"]);
  });

  it("updates an existing formula record in place without allocating a new id", () => {
    const store = new CellStore();
    const cellIndex = store.allocate(0, 0, 0);
    const formulas = new FormulaTable<{ cellIndex: number; source: string }>(store);

    const originalId = formulas.set(cellIndex, { cellIndex, source: "A1*2" });
    const updatedId = formulas.set(cellIndex, { cellIndex, source: "A1*3" });

    expect(updatedId).toBe(originalId);
    expect(formulas.get(cellIndex)).toEqual({ cellIndex, source: "A1*3" });
    expect(formulas.size).toBe(1);
  });
});
