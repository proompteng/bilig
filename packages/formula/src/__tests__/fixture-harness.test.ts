import { describe, expect, it } from "vitest";
import { ValueTag, type CellValue, type LiteralInput } from "@bilig/protocol";
import { excelTop50StarterFixtures, type ExcelExpectedValue } from "../../../excel-fixtures/src/index.js";
import { formatAddress, parseRangeAddress } from "../addressing.js";
import { compileFormula, evaluatePlan } from "../index.js";
import { getCompatibilityEntry } from "../compatibility.js";

const executableFixtures = excelTop50StarterFixtures.filter((fixture) => {
  const entry = getCompatibilityEntry(fixture.id);
  const hasVolatileCall = /\b(TODAY|NOW)\s*\(/i.test(fixture.formula);
  return (
    entry !== undefined &&
    (entry.status === "implemented-js" || entry.status === "implemented-js-and-wasm") &&
    fixture.family !== "volatile" &&
    !hasVolatileCall
  );
});

describe("excel fixture harness", () => {
  it("executes implemented Top 50 starter fixtures through the JS evaluator", () => {
    for (const fixture of executableFixtures) {
      expect(fixture.outputs).toHaveLength(1);
      const compiled = compileFormula(fixture.formula);
      const values = new Map<string, CellValue>();
      for (const input of fixture.inputs) {
        values.set(input.address.toUpperCase(), literalToCellValue(input.input));
      }

      const value = evaluatePlan(compiled.jsPlan, {
        sheetName: fixture.sheetName ?? "Sheet1",
        resolveCell: (_sheetName, address) => values.get(address.toUpperCase()) ?? { tag: ValueTag.Empty },
        resolveRange: (_sheetName, start, end, refKind) => {
          if (refKind !== "cells") {
            return [];
          }
          const range = parseRangeAddress(`${start}:${end}`);
          if (range.kind !== "cells") {
            return [];
          }
          const output: CellValue[] = [];
          for (let row = range.start.row; row <= range.end.row; row += 1) {
            for (let col = range.start.col; col <= range.end.col; col += 1) {
              output.push(values.get(formatAddress(row, col).toUpperCase()) ?? { tag: ValueTag.Empty });
            }
          }
          return output;
        }
      });

      expect(value, fixture.id).toEqual(expectedValueToCellValue(fixture.outputs[0]!.expected));
    }
  });
});

function literalToCellValue(input: LiteralInput): CellValue {
  if (input === null) {
    return { tag: ValueTag.Empty };
  }
  switch (typeof input) {
    case "number":
      return { tag: ValueTag.Number, value: input };
    case "boolean":
      return { tag: ValueTag.Boolean, value: input };
    case "string":
      return { tag: ValueTag.String, value: input, stringId: 0 };
  }
}

function expectedValueToCellValue(expected: ExcelExpectedValue): CellValue {
  switch (expected.kind) {
    case "empty":
      return { tag: ValueTag.Empty };
    case "number":
      return { tag: ValueTag.Number, value: expected.value };
    case "boolean":
      return { tag: ValueTag.Boolean, value: expected.value };
    case "string":
      return { tag: ValueTag.String, value: expected.value, stringId: 0 };
    case "error":
      return { tag: ValueTag.Error, code: expected.code };
  }
}
