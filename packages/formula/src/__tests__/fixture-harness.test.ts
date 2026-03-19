import { afterEach, describe, expect, it, vi } from "vitest";
import { ValueTag, type CellValue, type LiteralInput } from "@bilig/protocol";
import { excelTop100CanonicalFixtures, type ExcelExpectedValue } from "../../../excel-fixtures/src/index.js";
import { excelDateTimeFixtureSuite } from "../../../excel-fixtures/src/datetime-fixtures.js";
import { formatAddress, parseRangeAddress } from "../addressing.js";
import { compileFormula, evaluatePlan } from "../index.js";
import { getCompatibilityEntry } from "../compatibility.js";

const executableStatuses = new Set(["implemented-js", "implemented-js-and-wasm-shadow", "implemented-wasm-production"]);

const executableFixtures = excelTop100CanonicalFixtures.filter((fixture) => {
  const entry = getCompatibilityEntry(fixture.id);
  const hasVolatileCall = /\b(TODAY|NOW)\s*\(/i.test(fixture.formula);
  return entry !== undefined && executableStatuses.has(entry.status) && fixture.family !== "volatile" && !hasVolatileCall;
});

const executableDateTimeFixtures = (excelDateTimeFixtureSuite.cases ?? []).filter((fixture) => {
  const entry = getCompatibilityEntry(fixture.id);
  const hasVolatileCall = /\b(TODAY|NOW|RAND)\s*\(/i.test(fixture.formula);
  return entry !== undefined && executableStatuses.has(entry.status) && fixture.family !== "volatile" && !hasVolatileCall;
});

describe("excel fixture harness", () => {
  it("executes implemented canonical Top 100 fixtures through the JS evaluator", () => {
    for (const fixture of executableFixtures) {
      expect(evaluateFixture(fixture), fixture.id).toEqual(expectedValueToCellValue(fixture.outputs[0]!.expected));
    }
  });

  it("executes implemented date-time edge fixtures through the JS evaluator", () => {
    for (const fixture of executableDateTimeFixtures) {
      expect(evaluateFixture(fixture), fixture.id).toEqual(expectedValueToCellValue(fixture.outputs[0]!.expected));
    }
  });

  it("executes implemented volatile RAND fixtures deterministically through the JS evaluator", () => {
    const randomSpy = vi.spyOn(Math, "random").mockReturnValue(0.625);
    const randFixture = excelTop100CanonicalFixtures.find((fixture) => fixture.id === "volatile:rand-basic");

    expect(randFixture).toBeDefined();
    expect(getCompatibilityEntry("volatile:rand-basic")?.status).toBe("implemented-js");

    const compiled = compileFormula(randFixture!.formula);
    const value = evaluatePlan(compiled.jsPlan, {
      sheetName: randFixture!.sheetName ?? "Sheet1",
      resolveCell: () => ({ tag: ValueTag.Empty }),
      resolveRange: () => []
    });

    expect(randomSpy).toHaveBeenCalled();
    expect(value).toEqual(expectedValueToCellValue(randFixture!.outputs[0]!.expected));
  });
});

afterEach(() => {
  vi.restoreAllMocks();
});

function evaluateFixture(fixture: { formula: string; inputs: { address: string; input: LiteralInput }[]; outputs: { expected: ExcelExpectedValue }[]; sheetName?: string }): CellValue {
  expect(fixture.outputs).toHaveLength(1);
  const compiled = compileFormula(fixture.formula);
  const values = new Map<string, CellValue>();
  for (const input of fixture.inputs) {
    values.set(input.address.toUpperCase(), literalToCellValue(input.input));
  }

  return evaluatePlan(compiled.jsPlan, {
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
}

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
