import { describe, expect, it } from "vitest";
import { buildMatrixMutationPlan } from "../matrix-mutation-plan.js";

describe("buildMatrixMutationPlan", () => {
  it("emits clear and literal ops before formulas for mixed sheet imports", () => {
    const plan = buildMatrixMutationPlan({
      target: { sheet: 1, row: 0, col: 0 },
      targetSheetName: "Bench",
      content: [
        ["=A2+A3", 10, null],
        [20, "=B1*2", "done"],
      ],
      rewriteFormula: (formula) => formula.slice(1),
    });

    expect(plan.potentialNewCells).toBe(5);
    expect(plan.ops).toEqual([
      { kind: "setCellValue", sheetName: "Bench", address: "B1", value: 10 },
      { kind: "clearCell", sheetName: "Bench", address: "C1" },
      { kind: "setCellValue", sheetName: "Bench", address: "C2", value: "done" },
      { kind: "setCellFormula", sheetName: "Bench", address: "A1", formula: "A2+A3" },
      { kind: "setCellFormula", sheetName: "Bench", address: "B2", formula: "B1*2" },
      { kind: "setCellValue", sheetName: "Bench", address: "A2", value: 20 },
    ]);
  });
});
