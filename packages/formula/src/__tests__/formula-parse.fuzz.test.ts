import { describe, expect, it } from "vitest";
import { parseFormula } from "../parser.js";
import { serializeFormula } from "../translation.js";
import { invalidFormulaArbitrary, validFormulaArbitrary } from "./formula-fuzz-helpers.js";
import { runProperty } from "@bilig/test-fuzz";

describe("formula parse fuzz", () => {
  it("canonicalizes valid formulas through parse and serialize", async () => {
    await runProperty({
      suite: "formula/parse/canonicalization",
      arbitrary: validFormulaArbitrary,
      predicate: (formula) => {
        const canonical = serializeFormula(parseFormula(formula));
        expect(serializeFormula(parseFormula(canonical))).toBe(canonical);
      },
    });
  });

  it("rejects malformed formulas without crashing the parser", async () => {
    await runProperty({
      suite: "formula/parse/invalid-input",
      arbitrary: invalidFormulaArbitrary,
      predicate: (formula) => {
        expect(() => parseFormula(formula)).toThrow(/.+/);
      },
    });
  });
});
