import { describe, expect, it } from "vitest";
import fc from "fast-check";
import { ValueTag, type CellValue } from "@bilig/protocol";
import { evaluateAst } from "../js-evaluator.js";
import { parseFormula } from "../parser.js";
import {
  renameFormulaSheetReferences,
  rewriteFormulaForStructuralTransform,
  serializeFormula,
  translateFormulaReferences,
} from "../translation.js";
import { runProperty } from "@bilig/test-fuzz";

function quoteSheetName(sheetName: string): string {
  return sheetName.includes(" ") ? `'${sheetName}'` : sheetName;
}

function numberValue(value: number): CellValue {
  return { tag: ValueTag.Number, value };
}

function stringValue(value: string): CellValue {
  return { tag: ValueTag.String, value, stringId: 0 };
}

function booleanValue(value: boolean): CellValue {
  return { tag: ValueTag.Boolean, value };
}

function emptyValue(): CellValue {
  return { tag: ValueTag.Empty };
}

const sheetNameArbitrary = fc.constantFrom("Sheet1", "My Sheet", "Data");
const columnArbitrary = fc.constantFrom("A", "B", "C", "D", "E", "F", "G", "H");
const rowArbitrary = fc.integer({ min: 1, max: 30 });
const sheetQualifiedArbitrary = fc
  .option(sheetNameArbitrary, { nil: undefined })
  .map((sheetName) => (sheetName ? `${quoteSheetName(sheetName)}!` : ""));
const cellReferenceArbitrary = fc
  .tuple(fc.boolean(), columnArbitrary, fc.boolean(), rowArbitrary, sheetQualifiedArbitrary)
  .map(([anchorColumn, column, anchorRow, row, sheetPrefix]) => {
    return `${sheetPrefix}${anchorColumn ? "$" : ""}${column}${anchorRow ? "$" : ""}${row}`;
  });
const cellRangeReferenceArbitrary = fc
  .tuple(
    sheetQualifiedArbitrary,
    columnArbitrary,
    rowArbitrary,
    fc.integer({ min: 0, max: 1 }),
    fc.integer({ min: 0, max: 1 }),
  )
  .map(([sheetPrefix, column, row, rowSpan, columnSpan]) => {
    const startColumn = column.charCodeAt(0);
    const endColumn = String.fromCharCode(startColumn + columnSpan);
    const start = `${column}${row}`;
    const end = `${endColumn}${row + rowSpan}`;
    return `${sheetPrefix}${rowSpan === 0 && columnSpan === 0 ? start : `${start}:${end}`}`;
  });
const axisRangeArgumentArbitrary = fc.oneof(
  fc
    .tuple(fc.boolean(), columnArbitrary, sheetQualifiedArbitrary)
    .map(([anchorColumn, column, sheetPrefix]) => {
      const ref = `${anchorColumn ? "$" : ""}${column}`;
      return `${sheetPrefix}${ref}:${ref}`;
    }),
  fc
    .tuple(fc.boolean(), rowArbitrary, sheetQualifiedArbitrary)
    .map(([anchorRow, row, sheetPrefix]) => {
      const ref = `${anchorRow ? "$" : ""}${row}`;
      return `${sheetPrefix}${ref}:${ref}`;
    }),
);
const structuralFormulaReferenceArbitrary = fc.oneof(
  cellReferenceArbitrary,
  cellRangeReferenceArbitrary,
);
const scalarArbitrary = fc.oneof(
  structuralFormulaReferenceArbitrary,
  fc.integer({ min: -500, max: 500 }).map((value) => `${value}`),
  fc.constantFrom('"north"', '"sales"', '"ready"', "TRUE", "FALSE"),
);
const aggregateArgumentArbitrary = fc.oneof(scalarArbitrary, axisRangeArgumentArbitrary);
const validFormulaArbitrary = fc.oneof(
  scalarArbitrary,
  fc
    .tuple(scalarArbitrary, fc.constantFrom("+", "-", "*", "/", "&"), scalarArbitrary)
    .map(([left, operator, right]) => `${left}${operator}${right}`),
  fc
    .tuple(
      fc.constantFrom("SUM", "MAX", "MIN", "PRODUCT"),
      fc.array(aggregateArgumentArbitrary, { minLength: 1, maxLength: 3 }),
    )
    .map(([name, args]) => `${name}(${args.join(",")})`),
  fc
    .tuple(scalarArbitrary, scalarArbitrary, scalarArbitrary)
    .map(([condition, truthy, falsy]) => `IF(${condition},${truthy},${falsy})`),
);
function renameScopedFormulaArbitrary(
  oldSheetName: string,
  newSheetName: string,
): fc.Arbitrary<string> {
  const safeSheetChoices = ["Sheet1", "My Sheet", "Data", "Archive"].filter(
    (candidate) => candidate !== newSheetName,
  );
  const safeSheetArbitrary = fc.constantFrom(...safeSheetChoices);
  const renameScopedSheetQualifiedArbitrary = fc
    .oneof(safeSheetArbitrary, fc.constant(oldSheetName))
    .map((sheetName) => `${quoteSheetName(sheetName)}!`);
  const renameScopedCellReferenceArbitrary = fc
    .tuple(
      fc.boolean(),
      columnArbitrary,
      fc.boolean(),
      rowArbitrary,
      renameScopedSheetQualifiedArbitrary,
    )
    .map(([anchorColumn, column, anchorRow, row, sheetPrefix]) => {
      return `${sheetPrefix}${anchorColumn ? "$" : ""}${column}${anchorRow ? "$" : ""}${row}`;
    });
  const renameScopedRangeReferenceArbitrary = fc
    .tuple(
      renameScopedSheetQualifiedArbitrary,
      columnArbitrary,
      rowArbitrary,
      fc.integer({ min: 0, max: 1 }),
      fc.integer({ min: 0, max: 1 }),
    )
    .map(([sheetPrefix, column, row, rowSpan, columnSpan]) => {
      const startColumn = column.charCodeAt(0);
      const endColumn = String.fromCharCode(startColumn + columnSpan);
      const start = `${column}${row}`;
      const end = `${endColumn}${row + rowSpan}`;
      return `${sheetPrefix}${rowSpan === 0 && columnSpan === 0 ? start : `${start}:${end}`}`;
    });
  const renameScopedScalarArbitrary = fc.oneof(
    renameScopedCellReferenceArbitrary,
    renameScopedRangeReferenceArbitrary,
    fc.integer({ min: -500, max: 500 }).map((value) => `${value}`),
    fc.constantFrom('"north"', '"sales"', '"ready"', "TRUE", "FALSE"),
  );
  return fc.oneof(
    renameScopedScalarArbitrary,
    fc
      .tuple(
        renameScopedScalarArbitrary,
        fc.constantFrom("+", "-", "*", "/", "&"),
        renameScopedScalarArbitrary,
      )
      .map(([left, operator, right]) => `${left}${operator}${right}`),
    fc
      .tuple(
        fc.constantFrom("SUM", "MAX", "MIN", "PRODUCT"),
        fc.array(renameScopedScalarArbitrary, { minLength: 1, maxLength: 3 }),
      )
      .map(([name, args]) => `${name}(${args.join(",")})`),
  );
}
const invalidFormulaArbitrary = fc.constantFrom(
  "SUM(",
  "A1:B",
  "A1:2",
  "'Sheet 1'!1",
  "'Sheet 1'!$1",
  "SUM(A1,,B2)",
);

const evaluableCellReferenceArbitrary = fc.constantFrom(
  "A1",
  "B2",
  "C3",
  "Sheet2!B1",
  "Summary!C2",
);
const evaluableRangeReferenceArbitrary = fc.constantFrom("A1:B2", "Summary!A1:B2", "Sheet2!C1:D2");
const evaluableNumericScalarArbitrary = fc.oneof(
  evaluableCellReferenceArbitrary,
  fc.integer({ min: -20, max: 20 }).map((value) => `${value}`),
);
const evaluableTextScalarArbitrary = fc.constantFrom('"alpha"', '"beta"', '"42"', '"-12.5"');
const evaluableBooleanScalarArbitrary = fc.constantFrom("TRUE", "FALSE");
const evaluableFormulaArbitrary = fc.oneof(
  evaluableNumericScalarArbitrary,
  evaluableTextScalarArbitrary,
  evaluableBooleanScalarArbitrary,
  fc
    .tuple(
      evaluableNumericScalarArbitrary,
      fc.constantFrom("+", "-", "*", "/"),
      evaluableNumericScalarArbitrary,
    )
    .map(([left, operator, right]) => `${left}${operator}${right}`),
  fc
    .tuple(
      evaluableNumericScalarArbitrary,
      fc.constantFrom(">", ">=", "<", "<=", "=", "<>"),
      evaluableNumericScalarArbitrary,
    )
    .map(([left, operator, right]) => `${left}${operator}${right}`),
  fc
    .tuple(
      fc.constantFrom("SUM", "MAX", "MIN", "PRODUCT"),
      fc.array(fc.oneof(evaluableNumericScalarArbitrary, evaluableRangeReferenceArbitrary), {
        minLength: 1,
        maxLength: 3,
      }),
    )
    .map(([name, args]) => `${name}(${args.join(",")})`),
  fc
    .tuple(
      evaluableTextScalarArbitrary,
      fc.oneof(evaluableTextScalarArbitrary, evaluableCellReferenceArbitrary),
    )
    .map(([left, right]) => `${left}&${right}`),
  fc
    .tuple(
      fc.oneof(
        evaluableBooleanScalarArbitrary,
        evaluableNumericScalarArbitrary.map((value) => `${value}>0`),
      ),
      fc.oneof(evaluableNumericScalarArbitrary, evaluableTextScalarArbitrary),
      fc.oneof(evaluableNumericScalarArbitrary, evaluableTextScalarArbitrary),
    )
    .map(([condition, truthy, falsy]) => `IF(${condition},${truthy},${falsy})`),
  evaluableTextScalarArbitrary.map((value) => `VALUE(${value})`),
  fc
    .oneof(evaluableTextScalarArbitrary, evaluableCellReferenceArbitrary)
    .map((value) => `LEN(${value})`),
);

const evaluationContext = {
  sheetName: "Sheet1",
  currentAddress: "D9",
  resolveCell: (sheetName: string, address: string): CellValue => {
    switch (`${sheetName}!${address}`) {
      case "Sheet1!A1":
        return numberValue(4);
      case "Sheet1!B2":
        return numberValue(-3);
      case "Sheet1!C3":
        return stringValue("delta");
      case "Sheet2!B1":
        return numberValue(9);
      case "Summary!C2":
        return booleanValue(true);
      default:
        return emptyValue();
    }
  },
  resolveRange: (
    sheetName: string,
    start: string,
    end: string,
    refKind: "cells" | "rows" | "cols",
  ): CellValue[] => {
    switch (`${sheetName}!${start}:${end}:${refKind}`) {
      case "Sheet1!A1:B2:cells":
        return [numberValue(1), numberValue(2), numberValue(3), numberValue(4)];
      case "Summary!A1:B2:cells":
        return [numberValue(10), numberValue(20), numberValue(30), numberValue(40)];
      case "Sheet2!C1:D2:cells":
        return [numberValue(5), numberValue(6), numberValue(7), numberValue(8)];
      default:
        return [];
    }
  },
  listSheetNames: (): string[] => ["Sheet1", "Sheet2", "Summary"],
};

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
      suite: "formula/structural-insert-delete-reversal",
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
      suite: "formula/sheet-rename-roundtrip",
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

  it("keeps JS evaluation stable across canonicalization for coercion-heavy formulas", async () => {
    await runProperty({
      suite: "formula/evaluation-canonicalization-stability",
      arbitrary: evaluableFormulaArbitrary,
      predicate: (formula) => {
        const canonical = serializeFormula(parseFormula(formula));
        expect(evaluateAst(parseFormula(formula), evaluationContext)).toEqual(
          evaluateAst(parseFormula(canonical), evaluationContext),
        );
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
