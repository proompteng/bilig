import { afterEach, describe, expect, it, vi } from "vitest";
import { ValueTag, type CellValue, type LiteralInput } from "@bilig/protocol";
import { canonicalFormulaFixtures, type ExcelExpectedValue } from "../../../excel-fixtures/src/index.js";
import { excelDateTimeFixtureSuite } from "../../../excel-fixtures/src/datetime-fixtures.js";
import { formatAddress, parseRangeAddress } from "../addressing.js";
import { compileFormula, evaluatePlan } from "../index.js";
import { getCompatibilityEntry } from "../compatibility.js";

const executableStatuses = new Set(["implemented-js", "implemented-js-and-wasm-shadow", "implemented-wasm-production"]);

const executableFixtures = canonicalFormulaFixtures.filter((fixture) => {
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
  it("executes implemented canonical formula fixtures through the JS evaluator", () => {
    for (const fixture of executableFixtures) {
      expect(evaluateFixture(fixture)).toEqual(expectedValueToCellValue(firstOutput(fixture).expected));
    }
  });

  it("executes implemented date-time edge fixtures through the JS evaluator", () => {
    for (const fixture of executableDateTimeFixtures) {
      expect(evaluateFixture(fixture)).toEqual(expectedValueToCellValue(firstOutput(fixture).expected));
    }
  });

  it("executes implemented volatile RAND fixtures deterministically through the JS evaluator", () => {
    const randomSpy = vi.spyOn(Math, "random").mockReturnValue(0.625);
    const randFixture = canonicalFormulaFixtures.find((fixture) => fixture.id === "volatile:rand-basic");

    expect(randFixture).toBeDefined();
    expect(getCompatibilityEntry("volatile:rand-basic")?.status).toBe("implemented-js");

    const compiled = compileFormula(randFixture!.formula);
    const value = evaluatePlan(compiled.jsPlan, {
      sheetName: randFixture!.sheetName ?? "Sheet1",
      resolveCell: () => ({ tag: ValueTag.Empty }),
      resolveRange: () => []
    });

    expect(randomSpy).toHaveBeenCalled();
    expect(value).toEqual(expectedValueToCellValue(firstOutput(randFixture).expected));
  });

  it("executes implemented volatile TODAY and NOW fixtures against the captured UTC timestamp", () => {
    vi.useFakeTimers();
    vi.setSystemTime(new Date("2026-03-19T15:45:30.000Z"));

    const todayFixture = canonicalFormulaFixtures.find((fixture) => fixture.id === "date-time:today-volatile");
    const nowFixture = canonicalFormulaFixtures.find((fixture) => fixture.id === "date-time:now-volatile");

    expect(todayFixture).toBeDefined();
    expect(nowFixture).toBeDefined();
    expect(getCompatibilityEntry("date-time:today-volatile")?.status).toBe("implemented-js");
    expect(getCompatibilityEntry("date-time:now-volatile")?.status).toBe("implemented-js");

    const todayCompiled = compileFormula(todayFixture!.formula);
    const nowCompiled = compileFormula(nowFixture!.formula);
    const context = {
      sheetName: "Sheet1",
      resolveCell: () => ({ tag: ValueTag.Empty } as const),
      resolveRange: () => []
    };

    expect(evaluatePlan(todayCompiled.jsPlan, context)).toEqual(expectedValueToCellValue(firstOutput(todayFixture).expected));
    expect(evaluatePlan(nowCompiled.jsPlan, context)).toEqual(expectedValueToCellValue(firstOutput(nowFixture).expected));
  });

  it("executes the seeded logical backlog fixtures through the JS evaluator once they are promoted", () => {
    const ifFixture = canonicalFormulaFixtures.find((fixture) => fixture.id === "logical:if-condition-error");
    const ifnaFixture = canonicalFormulaFixtures.find((fixture) => fixture.id === "logical:ifna-catches-na-only");

    expect(ifFixture).toBeDefined();
    expect(ifnaFixture).toBeDefined();
    expect(getCompatibilityEntry("logical:if-condition-error")?.status).toBe("implemented-js");
    expect(getCompatibilityEntry("logical:ifna-catches-na-only")?.status).toBe("implemented-js");

    expect(evaluateFixture(ifFixture)).toEqual(expectedValueToCellValue(firstOutput(ifFixture).expected));
    expect(evaluateFixture(ifnaFixture)).toEqual(expectedValueToCellValue(firstOutput(ifnaFixture).expected));
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

function firstOutput(fixture: { outputs: { expected: ExcelExpectedValue }[] }): { expected: ExcelExpectedValue } {
  expect(fixture.outputs).toHaveLength(1);
  return fixture.outputs[0];
}

function literalToCellValue(input: LiteralInput): CellValue {
  if (input === null) {
    return { tag: ValueTag.Empty };
  }
  const inputType = typeof input;
  switch (inputType) {
    case "number":
      return { tag: ValueTag.Number, value: input };
    case "boolean":
      return { tag: ValueTag.Boolean, value: input };
    case "string":
      return { tag: ValueTag.String, value: input, stringId: 0 };
    case "bigint":
    case "function":
    case "object":
    case "symbol":
    case "undefined":
      throw new Error(`Unsupported literal input type: ${inputType}`);
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
