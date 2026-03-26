import { afterEach, describe, expect, it, vi } from "vitest";
import { ErrorCode, ValueTag, type CellValue, type LiteralInput } from "@bilig/protocol";
import {
  canonicalFormulaFixtures,
  canonicalWorkbookSemanticsFixtures,
  type ExcelExpectedValue,
  type ExcelFixtureCase,
} from "../../../excel-fixtures/src/index.js";
import { excelDateTimeFixtureSuite } from "../../../excel-fixtures/src/datetime-fixtures.js";
import { formatAddress, parseCellAddress, parseRangeAddress } from "../addressing.js";
import {
  compileFormula,
  compileFormulaAst,
  evaluatePlan,
  evaluatePlanResult,
  isArrayValue,
  parseFormula,
  type FormulaNode,
} from "../index.js";
import { getCompatibilityEntry } from "../compatibility.js";

const executableStatuses = new Set([
  "implemented-js",
  "implemented-js-and-wasm-shadow",
  "implemented-wasm-production",
]);

const executableFixtures = canonicalFormulaFixtures.filter((fixture) => {
  const entry = getCompatibilityEntry(fixture.id);
  const hasVolatileCall = /\b(TODAY|NOW)\s*\(/i.test(fixture.formula);
  return (
    entry !== undefined &&
    executableStatuses.has(entry.status) &&
    fixture.id !== "lookup-reference:offset-basic" &&
    fixture.family !== "volatile" &&
    !hasVolatileCall
  );
});

const executableWorkbookSemanticsFixtures = canonicalWorkbookSemanticsFixtures.filter((fixture) => {
  const entry = getCompatibilityEntry(fixture.id);
  return entry !== undefined && executableStatuses.has(entry.status);
});

const executableDateTimeFixtures = (excelDateTimeFixtureSuite.cases ?? []).filter((fixture) => {
  const entry = getCompatibilityEntry(fixture.id);
  const hasVolatileCall = /\b(TODAY|NOW|RAND)\s*\(/i.test(fixture.formula);
  return (
    entry !== undefined &&
    executableStatuses.has(entry.status) &&
    fixture.family !== "volatile" &&
    !hasVolatileCall
  );
});

describe("excel fixture harness", () => {
  it("executes implemented canonical formula fixtures through the JS evaluator", () => {
    for (const fixture of executableFixtures) {
      expectFixtureResult(fixture);
    }
  });

  it("executes workbook semantics fixtures through the JS evaluator", () => {
    for (const fixture of executableWorkbookSemanticsFixtures) {
      expectFixtureResult(fixture);
    }
  });

  it("executes implemented date-time edge fixtures through the JS evaluator", () => {
    for (const fixture of executableDateTimeFixtures) {
      expectFixtureResult(fixture);
    }
  });

  it("executes implemented volatile RAND fixtures deterministically through the JS evaluator", () => {
    const randomSpy = vi.spyOn(Math, "random").mockReturnValue(0.625);
    const randFixture = canonicalFormulaFixtures.find(
      (fixture) => fixture.id === "volatile:rand-basic",
    );

    expect(randFixture).toBeDefined();
    expect(getCompatibilityEntry("volatile:rand-basic")?.status).toBe(
      "implemented-wasm-production",
    );

    const compiled = compileFormula(randFixture!.formula);
    const value = evaluatePlan(compiled.jsPlan, {
      sheetName: randFixture!.sheetName ?? "Sheet1",
      resolveCell: () => ({ tag: ValueTag.Empty }),
      resolveRange: () => [],
    });

    expect(randomSpy).toHaveBeenCalled();
    expect(value).toEqual(expectedValueToCellValue(firstOutput(randFixture).expected));
  });

  it("executes implemented volatile TODAY and NOW fixtures against the captured UTC timestamp", () => {
    vi.useFakeTimers();
    vi.setSystemTime(new Date("2026-03-19T15:45:30.000Z"));

    const todayFixture = canonicalFormulaFixtures.find(
      (fixture) => fixture.id === "date-time:today-volatile",
    );
    const nowFixture = canonicalFormulaFixtures.find(
      (fixture) => fixture.id === "date-time:now-volatile",
    );

    expect(todayFixture).toBeDefined();
    expect(nowFixture).toBeDefined();
    expect(getCompatibilityEntry("date-time:today-volatile")?.status).toBe(
      "implemented-wasm-production",
    );
    expect(getCompatibilityEntry("date-time:now-volatile")?.status).toBe(
      "implemented-wasm-production",
    );

    const todayCompiled = compileFormula(todayFixture!.formula);
    const nowCompiled = compileFormula(nowFixture!.formula);
    const context = {
      sheetName: "Sheet1",
      resolveCell: () => ({ tag: ValueTag.Empty }) as const,
      resolveRange: () => [],
    };

    expect(evaluatePlan(todayCompiled.jsPlan, context)).toEqual(
      expectedValueToCellValue(firstOutput(todayFixture).expected),
    );
    expect(evaluatePlan(nowCompiled.jsPlan, context)).toEqual(
      expectedValueToCellValue(firstOutput(nowFixture).expected),
    );
  });

  it("executes the seeded logical backlog fixtures after native promotion", () => {
    const ifFixture = canonicalFormulaFixtures.find(
      (fixture) => fixture.id === "logical:if-condition-error",
    );
    const ifnaFixture = canonicalFormulaFixtures.find(
      (fixture) => fixture.id === "logical:ifna-catches-na-only",
    );

    expect(ifFixture).toBeDefined();
    expect(ifnaFixture).toBeDefined();
    expect(getCompatibilityEntry("logical:if-condition-error")?.status).toBe(
      "implemented-wasm-production",
    );
    expect(getCompatibilityEntry("logical:ifna-catches-na-only")?.status).toBe(
      "implemented-wasm-production",
    );

    expect(evaluateFixture(ifFixture)).toEqual(
      expectedValueToCellValue(firstOutput(ifFixture).expected),
    );
    expect(evaluateFixture(ifnaFixture)).toEqual(
      expectedValueToCellValue(firstOutput(ifnaFixture).expected),
    );
  });
});

afterEach(() => {
  vi.restoreAllMocks();
});

function expectFixtureResult(fixture: ExcelFixtureCase): void {
  const result = evaluateFixture(fixture);
  if (!isArrayValue(result)) {
    assertScalarFixtureResult(fixture, result);
    return;
  }
  assertArrayFixtureResult(fixture, result);
}

function assertArrayFixtureResult(
  fixture: ExcelFixtureCase,
  result: Extract<ReturnType<typeof evaluatePlanResult>, { kind: "array" }>,
): void {
  const defaultSheetName = fixture.sheetName ?? "Sheet1";
  const first = fixture.outputs[0];
  if (!first) {
    throw new Error(`Fixture ${fixture.id} is missing outputs`);
  }
  const firstSheetName = first.sheetName ?? defaultSheetName;
  const firstAddress = parseCellAddress(first.address, firstSheetName);
  let maxRow = firstAddress.row;
  let maxCol = firstAddress.col;
  const expectedByOffset = new Map<string, CellValue>();

  fixture.outputs.forEach((output) => {
    const sheetName = output.sheetName ?? defaultSheetName;
    expect(sheetName).toBe(firstSheetName);
    const address = parseCellAddress(output.address, sheetName);
    maxRow = Math.max(maxRow, address.row);
    maxCol = Math.max(maxCol, address.col);
    expectedByOffset.set(
      `${address.row - firstAddress.row}:${address.col - firstAddress.col}`,
      expectedValueToCellValue(output.expected),
    );
  });

  expect(result.rows).toBe(maxRow - firstAddress.row + 1);
  expect(result.cols).toBe(maxCol - firstAddress.col + 1);
  expect(result.values).toEqual(
    Array.from({ length: result.rows * result.cols }, (_entry, index) => {
      const rowOffset = Math.floor(index / result.cols);
      const colOffset = index % result.cols;
      return expectedByOffset.get(`${rowOffset}:${colOffset}`) ?? { tag: ValueTag.Empty };
    }),
  );
}

function assertScalarFixtureResult(fixture: ExcelFixtureCase, result: CellValue): void {
  expect(fixture.outputs).toHaveLength(1);
  expect(result).toEqual(expectedValueToCellValue(firstOutput(fixture).expected));
}

function evaluateFixture(fixture: ExcelFixtureCase): ReturnType<typeof evaluatePlanResult> {
  const defaultSheetName = fixture.sheetName ?? "Sheet1";
  const originalAst = parseFormula(fixture.formula);
  const resolvedAst = resolveFixtureMetadataReferences(originalAst, fixture, defaultSheetName);
  const compiled =
    resolvedAst === originalAst
      ? compileFormula(fixture.formula)
      : compileFormulaAst(fixture.formula, resolvedAst, { originalAst });
  const sheetValues = new Map<string, Map<string, CellValue>>();
  for (const input of fixture.inputs) {
    const sheetName = input.sheetName ?? defaultSheetName;
    let values = sheetValues.get(sheetName);
    if (!values) {
      values = new Map<string, CellValue>();
      sheetValues.set(sheetName, values);
    }
    values.set(input.address.toUpperCase(), literalToCellValue(input.input));
  }
  const definedNames = new Map<string, CellValue>(
    (fixture.definedNames ?? []).map((definedName) => [
      definedName.name.toUpperCase(),
      literalToCellValue(definedName.value),
    ]),
  );
  const hasSheet = (sheetName: string) =>
    sheetName === defaultSheetName || sheetValues.has(sheetName);

  return evaluatePlanResult(compiled.jsPlan, {
    sheetName: defaultSheetName,
    resolveCell: (sheetName, address) => {
      if (!hasSheet(sheetName)) {
        return { tag: ValueTag.Error, code: ErrorCode.Ref };
      }
      return sheetValues.get(sheetName)?.get(address.toUpperCase()) ?? { tag: ValueTag.Empty };
    },
    resolveRange: (sheetName, start, end, refKind) => {
      if (!hasSheet(sheetName)) {
        return [{ tag: ValueTag.Error, code: ErrorCode.Ref }];
      }
      if (refKind !== "cells") {
        return [];
      }
      const range = parseRangeAddress(`${start}:${end}`);
      if (range.kind !== "cells") {
        return [];
      }
      const output: CellValue[] = [];
      const sheetCells = sheetValues.get(sheetName);
      for (let row = range.start.row; row <= range.end.row; row += 1) {
        for (let col = range.start.col; col <= range.end.col; col += 1) {
          output.push(
            sheetCells?.get(formatAddress(row, col).toUpperCase()) ?? { tag: ValueTag.Empty },
          );
        }
      }
      return output;
    },
    resolveName: (name) =>
      definedNames.get(name.toUpperCase()) ?? { tag: ValueTag.Error, code: ErrorCode.Name },
  });
}

function resolveFixtureMetadataReferences(
  node: FormulaNode,
  fixture: ExcelFixtureCase,
  defaultSheetName: string,
  activeNames = new Set<string>(),
): FormulaNode {
  switch (node.kind) {
    case "NumberLiteral":
    case "BooleanLiteral":
    case "StringLiteral":
    case "ErrorLiteral":
    case "CellRef":
    case "RowRef":
    case "ColumnRef":
    case "RangeRef":
    case "SpillRef":
      return node;
    case "NameRef": {
      const normalized = node.name.trim().toUpperCase();
      if (activeNames.has(normalized)) {
        return { kind: "ErrorLiteral", code: ErrorCode.Cycle };
      }
      const definedName = fixture.definedNames?.find(
        (entry) => entry.name.trim().toUpperCase() === normalized,
      );
      if (!definedName) {
        return node;
      }
      const replacement = definedNameValueToFormulaNode(definedName.value);
      if (!replacement) {
        return node;
      }
      const nextActiveNames = new Set(activeNames);
      nextActiveNames.add(normalized);
      return resolveFixtureMetadataReferences(
        replacement,
        fixture,
        defaultSheetName,
        nextActiveNames,
      );
    }
    case "StructuredRef": {
      const replacement = resolveFixtureStructuredReference(
        node.tableName,
        node.columnName,
        fixture,
        defaultSheetName,
      );
      return replacement ?? { kind: "ErrorLiteral", code: ErrorCode.Ref };
    }
    case "UnaryExpr":
      return {
        ...node,
        argument: resolveFixtureMetadataReferences(
          node.argument,
          fixture,
          defaultSheetName,
          activeNames,
        ),
      };
    case "BinaryExpr":
      return {
        ...node,
        left: resolveFixtureMetadataReferences(node.left, fixture, defaultSheetName, activeNames),
        right: resolveFixtureMetadataReferences(node.right, fixture, defaultSheetName, activeNames),
      };
    case "CallExpr":
      return {
        ...node,
        args: node.args.map((arg) =>
          resolveFixtureMetadataReferences(arg, fixture, defaultSheetName, activeNames),
        ),
      };
    case "InvokeExpr":
      return {
        ...node,
        callee: resolveFixtureMetadataReferences(
          node.callee,
          fixture,
          defaultSheetName,
          activeNames,
        ),
        args: node.args.map((arg) =>
          resolveFixtureMetadataReferences(arg, fixture, defaultSheetName, activeNames),
        ),
      };
  }
}

function definedNameValueToFormulaNode(value: LiteralInput): FormulaNode | undefined {
  if (value === null) {
    return { kind: "ErrorLiteral", code: ErrorCode.Ref };
  }
  switch (typeof value) {
    case "number":
      return { kind: "NumberLiteral", value };
    case "boolean":
      return { kind: "BooleanLiteral", value };
    case "string":
      return value.startsWith("=")
        ? parseFormula(value.slice(1))
        : { kind: "StringLiteral", value };
    case "bigint":
    case "function":
    case "object":
    case "symbol":
    case "undefined":
      return undefined;
  }
}

function resolveFixtureStructuredReference(
  tableName: string,
  columnName: string,
  fixture: ExcelFixtureCase,
  defaultSheetName: string,
): FormulaNode | undefined {
  const table = fixture.tables?.find(
    (entry) => entry.name.trim().toUpperCase() === tableName.trim().toUpperCase(),
  );
  if (!table) {
    return undefined;
  }
  const columnIndex = table.columnNames.findIndex(
    (name) => name.trim().toUpperCase() === columnName.trim().toUpperCase(),
  );
  if (columnIndex === -1) {
    return undefined;
  }
  const sheetName = table.sheetName ?? defaultSheetName;
  const start = parseCellAddress(table.startAddress, sheetName);
  const end = parseCellAddress(table.endAddress, sheetName);
  const startRow = start.row + (table.headerRow ? 1 : 0);
  const endRow = end.row - (table.totalsRow ? 1 : 0);
  if (endRow < startRow) {
    return { kind: "ErrorLiteral", code: ErrorCode.Ref };
  }
  const column = start.col + columnIndex;
  return {
    kind: "RangeRef",
    refKind: "cells",
    sheetName,
    start: formatAddress(startRow, column),
    end: formatAddress(endRow, column),
  };
}

function firstOutput(fixture: { outputs: { expected: ExcelExpectedValue }[] }): {
  expected: ExcelExpectedValue;
} {
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
