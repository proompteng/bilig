import { describe, expect, it } from "vitest";
import fc from "fast-check";
import { parseFormula } from "../parser.js";
import { serializeFormula, translateFormulaReferences } from "../translation.js";
import { runProperty } from "@bilig/test-fuzz";

function quoteSheetName(sheetName: string): string {
  return sheetName.includes(" ") ? `'${sheetName}'` : sheetName;
}

const sheetNameArbitrary = fc.constantFrom("Sheet1", "My Sheet", "Data");
const columnArbitrary = fc.constantFrom("C", "D", "E", "F", "G", "H");
const rowArbitrary = fc.integer({ min: 5, max: 30 });
const anchoredCellReferenceArbitrary = fc
  .tuple(
    fc.boolean(),
    columnArbitrary,
    fc.boolean(),
    rowArbitrary,
    fc.option(sheetNameArbitrary, { nil: undefined }),
  )
  .map(([anchorColumn, column, anchorRow, row, sheetName]) => {
    const ref = `${anchorColumn ? "$" : ""}${column}${anchorRow ? "$" : ""}${row}`;
    return sheetName ? `${quoteSheetName(sheetName)}!${ref}` : ref;
  });
const rangeReferenceArbitrary = fc
  .tuple(
    fc.option(sheetNameArbitrary, { nil: undefined }),
    columnArbitrary,
    rowArbitrary,
    fc.integer({ min: 0, max: 1 }),
    fc.integer({ min: 0, max: 1 }),
  )
  .map(([sheetName, column, row, rowSpan, columnSpan]) => {
    const startColumn = column.charCodeAt(0);
    const endColumn = String.fromCharCode(startColumn + columnSpan);
    const start = `${column}${row}`;
    const end = `${endColumn}${row + rowSpan}`;
    const range = rowSpan === 0 && columnSpan === 0 ? start : `${start}:${end}`;
    return sheetName ? `${quoteSheetName(sheetName)}!${range}` : range;
  });
const scalarArbitrary = fc.oneof(
  anchoredCellReferenceArbitrary,
  rangeReferenceArbitrary,
  fc.integer({ min: -500, max: 500 }).map((value) => `${value}`),
  fc.constantFrom('"north"', '"sales"', '"ready"', "TRUE", "FALSE"),
);
const validFormulaArbitrary = fc.oneof(
  scalarArbitrary,
  fc
    .tuple(scalarArbitrary, fc.constantFrom("+", "-", "*", "/", "&"), scalarArbitrary)
    .map(([left, operator, right]) => `${left}${operator}${right}`),
  fc
    .tuple(
      fc.constantFrom("SUM", "MAX", "MIN", "PRODUCT"),
      fc.array(scalarArbitrary, { minLength: 1, maxLength: 3 }),
    )
    .map(([name, args]) => `${name}(${args.join(",")})`),
  fc
    .tuple(
      fc.constantFrom("SUM", "MAX"),
      fc.array(scalarArbitrary, { minLength: 1, maxLength: 2 }),
      fc.constantFrom("+", "-"),
      scalarArbitrary,
    )
    .map(([name, args, operator, right]) => `(${name}(${args.join(",")}))${operator}${right}`),
);
const invalidFormulaArbitrary = fc.constantFrom(
  "SUM(",
  "A1:B",
  "A1:2",
  "'Sheet 1'!1",
  "'Sheet 1'!$1",
  "SUM(A1,,B2)",
);

describe("formula fuzz", () => {
  it("canonicalizes valid formulas through parse and serialize", async () => {
    await runProperty({
      suite: "formula/canonicalization",
      arbitrary: validFormulaArbitrary,
      predicate: (formula) => {
        const canonical = serializeFormula(parseFormula(formula));
        expect(serializeFormula(parseFormula(canonical))).toBe(canonical);
      },
    });
  });

  it("reverses translated references back to the canonical formula", async () => {
    await runProperty({
      suite: "formula/translation-reversal",
      arbitrary: fc
        .record({
          formula: validFormulaArbitrary,
          rowDelta: fc.integer({ min: -4, max: 4 }),
          colDelta: fc.integer({ min: -2, max: 2 }),
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

  it("rejects malformed formulas without crashing the parser", async () => {
    await runProperty({
      suite: "formula/invalid-input",
      arbitrary: invalidFormulaArbitrary,
      predicate: (formula) => {
        expect(() => parseFormula(formula)).toThrow(/.+/);
      },
    });
  });
});
