import { describe, expect, it } from "vitest";
import { FormulaMode } from "@bilig/protocol";
import { SpreadsheetEngine } from "../engine.js";
import { fastPathFormulaArbitrary } from "../../../formula/src/__tests__/formula-fuzz-helpers.js";
import { runProperty } from "@bilig/test-fuzz";

const sheetName = "Sheet1";

function buildNumericGrid(rows: number, cols: number): number[][] {
  return Array.from({ length: rows }, (_rowValue, row) =>
    Array.from({ length: cols }, (_colValue, col) => row * cols + col + 1),
  );
}

describe("formula runtime differential fuzz", () => {
  it("keeps generated fast-path formulas in JS and wasm parity", async () => {
    await runProperty({
      suite: "core/formula-runtime/generated-differential",
      arbitrary: fastPathFormulaArbitrary,
      predicate: async (formula) => {
        const engine = new SpreadsheetEngine({
          workbookName: `fuzz-formula-diff-${formula.length}`,
          replicaId: "fuzz-formula-diff",
        });
        await engine.ready();
        engine.createSheet(sheetName);
        engine.setRangeValues(
          { sheetName, startAddress: "A1", endAddress: "F6" },
          buildNumericGrid(6, 6),
        );
        engine.setCellFormula(sheetName, "G1", formula);

        const explanation = engine.explainCell(sheetName, "G1");
        expect(explanation.mode).toBe(FormulaMode.WasmFastPath);

        const differential = engine.recalculateDifferential();
        expect(differential.drift).toEqual([]);

        const snapshot = engine.exportSnapshot();
        const restored = new SpreadsheetEngine({
          workbookName: snapshot.workbook.name,
          replicaId: "fuzz-formula-diff-restored",
        });
        await restored.ready();
        restored.importSnapshot(snapshot);

        expect(restored.getCellValue(sheetName, "G1")).toEqual(
          engine.getCellValue(sheetName, "G1"),
        );
      },
    });
  });
});
