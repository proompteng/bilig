import { describe, expect, it } from "vitest";
import fc from "fast-check";
import { parseFormula } from "../parser.js";
import { renameFormulaSheetReferences, serializeFormula } from "../translation.js";
import { renameScopedFormulaArbitrary, sheetNameArbitrary } from "./formula-fuzz-helpers.js";
import { runProperty } from "@bilig/test-fuzz";

describe("formula rename fuzz", () => {
  it("roundtrips quoted and unquoted sheet renames", async () => {
    await runProperty({
      suite: "formula/rename/sheet-roundtrip",
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
