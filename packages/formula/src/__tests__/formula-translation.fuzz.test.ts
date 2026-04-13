import { describe, expect, it } from "vitest";
import fc from "fast-check";
import { parseFormula } from "../parser.js";
import {
  renameFormulaSheetReferences,
  rewriteFormulaForStructuralTransform,
  serializeFormula,
  translateFormulaReferences,
} from "../translation.js";
import {
  renameScopedFormulaArbitrary,
  sheetNameArbitrary,
  validFormulaArbitrary,
} from "./formula-fuzz-helpers.js";
import { runProperty } from "@bilig/test-fuzz";

describe("formula translation fuzz", () => {
  it("reverses translated references back to the canonical formula", async () => {
    await runProperty({
      suite: "formula/translation/reference-reversal",
      arbitrary: fc
        .record({
          formula: validFormulaArbitrary,
          rowDelta: fc.integer({ min: 0, max: 4 }),
          colDelta: fc.integer({ min: 0, max: 2 }),
        })
        .filter((value) => value.rowDelta !== 0 || value.colDelta !== 0),
      predicate: ({ formula, rowDelta, colDelta }) => {
        const canonical = serializeFormula(parseFormula(formula));
        const translated = translateFormulaReferences(canonical, rowDelta, colDelta);
        const restored = translateFormulaReferences(translated, -rowDelta, -colDelta);
        expect(restored).toBe(canonical);
      },
    });
  });

  it("reverses insert transforms through matching delete transforms", async () => {
    await runProperty({
      suite: "formula/translation/structural-insert-delete-reversal",
      arbitrary: fc.record({
        formula: validFormulaArbitrary,
        axis: fc.constantFrom<"row" | "column">("row", "column"),
        start: fc.integer({ min: 0, max: 4 }),
        count: fc.integer({ min: 1, max: 2 }),
      }),
      predicate: ({ formula, axis, start, count }) => {
        const canonical = serializeFormula(parseFormula(formula));
        const inserted = rewriteFormulaForStructuralTransform(canonical, "Sheet1", "Sheet1", {
          kind: "insert",
          axis,
          start,
          count,
        });
        const restored = rewriteFormulaForStructuralTransform(inserted, "Sheet1", "Sheet1", {
          kind: "delete",
          axis,
          start,
          count,
        });
        expect(restored).toBe(canonical);
      },
    });
  });

  it("roundtrips quoted and unquoted sheet renames", async () => {
    await runProperty({
      suite: "formula/translation/sheet-rename-roundtrip",
      arbitrary: fc
        .record({
          oldSheetName: sheetNameArbitrary,
          newSheetName: sheetNameArbitrary,
        })
        .filter((value) => value.oldSheetName !== value.newSheetName)
        .chain(({ oldSheetName, newSheetName }) =>
          renameScopedFormulaArbitrary(oldSheetName, newSheetName).map((formula) => ({
            formula,
            oldSheetName,
            newSheetName,
          })),
        ),
      predicate: ({ formula, oldSheetName, newSheetName }) => {
        const canonical = serializeFormula(parseFormula(formula));
        const renamed = renameFormulaSheetReferences(canonical, oldSheetName, newSheetName);
        const restored = renameFormulaSheetReferences(renamed, newSheetName, oldSheetName);
        expect(restored).toBe(canonical);
      },
    });
  });
});
