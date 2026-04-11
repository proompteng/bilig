import { describe, expect, it } from "vitest";
import { buildMatrixMutationPlan } from "../matrix-mutation-plan.js";

describe("buildMatrixMutationPlan", () => {
  it("emits clear and literal ops before formulas for mixed sheet imports", () => {
    const plan = buildMatrixMutationPlan({
      target: { sheet: 1, row: 0, col: 0 },
      content: [
        ["=A2+A3", 10, null],
        [20, "=B1*2", "done"],
      ],
      rewriteFormula: (formula) => formula.slice(1),
    });

    expect(plan.potentialNewCells).toBe(5);
    expect(plan.refs).toEqual([
      { sheetId: 1, mutation: { kind: "setCellValue", row: 0, col: 1, value: 10 } },
      { sheetId: 1, mutation: { kind: "clearCell", row: 0, col: 2 } },
      { sheetId: 1, mutation: { kind: "setCellValue", row: 1, col: 2, value: "done" } },
      { sheetId: 1, mutation: { kind: "setCellFormula", row: 0, col: 0, formula: "A2+A3" } },
      { sheetId: 1, mutation: { kind: "setCellFormula", row: 1, col: 1, formula: "B1*2" } },
      { sheetId: 1, mutation: { kind: "setCellValue", row: 1, col: 0, value: 20 } },
    ]);
  });

  it("defers literals below the first formula in each column without scanning every formula", () => {
    const plan = buildMatrixMutationPlan({
      target: { sheet: 1, row: 0, col: 0 },
      content: [
        [1, "=A2+A3", "top"],
        [2, 20, "middle"],
        [3, 30, "=A1+B1"],
        [4, 40, "bottom"],
      ],
      rewriteFormula: (formula) => formula.slice(1),
    });

    expect(plan.leadingRefs).toEqual([
      { sheetId: 1, mutation: { kind: "setCellValue", row: 0, col: 0, value: 1 } },
      { sheetId: 1, mutation: { kind: "setCellValue", row: 0, col: 2, value: "top" } },
      { sheetId: 1, mutation: { kind: "setCellValue", row: 1, col: 0, value: 2 } },
      { sheetId: 1, mutation: { kind: "setCellValue", row: 1, col: 2, value: "middle" } },
      { sheetId: 1, mutation: { kind: "setCellValue", row: 2, col: 0, value: 3 } },
      { sheetId: 1, mutation: { kind: "setCellValue", row: 3, col: 0, value: 4 } },
    ]);
    expect(plan.formulaRefs).toEqual([
      { sheetId: 1, mutation: { kind: "setCellFormula", row: 0, col: 1, formula: "A2+A3" } },
      { sheetId: 1, mutation: { kind: "setCellFormula", row: 2, col: 2, formula: "A1+B1" } },
    ]);
    expect(plan.trailingLiteralRefs).toEqual([
      { sheetId: 1, mutation: { kind: "setCellValue", row: 1, col: 1, value: 20 } },
      { sheetId: 1, mutation: { kind: "setCellValue", row: 2, col: 1, value: 30 } },
      { sheetId: 1, mutation: { kind: "setCellValue", row: 3, col: 1, value: 40 } },
      { sheetId: 1, mutation: { kind: "setCellValue", row: 3, col: 2, value: "bottom" } },
    ]);
  });
});
