import { describe, expect, it } from "vitest";
import { FormulaMode, ValueTag, type CellValue } from "@bilig/protocol";
import {
  canonicalFormulaFixtures,
  type ExcelExpectedValue,
  type ExcelFixtureCase,
} from "../../../excel-fixtures/src/index.js";
import { getCompatibilityEntry } from "../../../formula/src/compatibility.js";
import { SpreadsheetEngine } from "../index.js";

// These formulas still compile onto the wasm-capable path, but the engine intentionally
// reroutes them to specialized JS lookup handlers at bind time.
const runtimeJsOnlyFixtureIds = new Set([
  "lookup-reference:match-exact",
  "lookup-reference:xmatch-basic",
]);

const productionRuntimeFixtures = canonicalFormulaFixtures.filter((fixture) => {
  const entry = getCompatibilityEntry(fixture.id);
  return (
    entry?.wasmStatus === "production" &&
    (fixture.family === "text" || fixture.family === "lookup-reference") &&
    !runtimeJsOnlyFixtureIds.has(fixture.id) &&
    fixture.id !== "lookup-reference:offset-basic" &&
    fixture.definedNames === undefined &&
    fixture.tables === undefined &&
    fixture.multipleOperations === undefined
  );
});

const groupedArrayProductionFixtures = canonicalFormulaFixtures.filter((fixture) => {
  const entry = getCompatibilityEntry(fixture.id);
  return (
    entry?.wasmStatus === "production" &&
    (fixture.id === "dynamic-array:groupby-basic" || fixture.id === "dynamic-array:pivotby-basic")
  );
});

describe("formula runtime correctness", () => {
  it("keeps canonical text and lookup fixtures in oracle parity on the wasm path", async () => {
    expect(productionRuntimeFixtures.length).toBeGreaterThan(0);

    await Promise.all(
      productionRuntimeFixtures.map(async (fixture) => {
        try {
          await expectFixtureParity(fixture);
        } catch (error) {
          const message = error instanceof Error ? error.message : String(error);
          throw new Error(`Fixture ${fixture.id} failed: ${message}`, { cause: error });
        }
      }),
    );
  });

  it("keeps canonical grouped-array SUM fixtures in oracle parity on the wasm path", async () => {
    expect(groupedArrayProductionFixtures).toHaveLength(2);

    await Promise.all(
      groupedArrayProductionFixtures.map(async (fixture) => {
        try {
          await expectFixtureParity(fixture);
        } catch (error) {
          const message = error instanceof Error ? error.message : String(error);
          throw new Error(`Fixture ${fixture.id} failed: ${message}`, { cause: error });
        }
      }),
    );
  });
});

async function expectFixtureParity(fixture: ExcelFixtureCase): Promise<void> {
  const engine = new SpreadsheetEngine({ workbookName: fixture.id });
  await engine.ready();

  const sheetNames = new Set<string>();
  const defaultSheetName = fixture.sheetName ?? "Sheet1";
  sheetNames.add(defaultSheetName);

  fixture.inputs.forEach((input) => sheetNames.add(input.sheetName ?? defaultSheetName));
  fixture.outputs.forEach((output) => sheetNames.add(output.sheetName ?? defaultSheetName));

  for (const sheetName of sheetNames) {
    engine.createSheet(sheetName);
  }

  for (const input of fixture.inputs) {
    engine.setCellValue(input.sheetName ?? defaultSheetName, input.address, input.input);
  }

  const owner = fixture.outputs[0];
  if (!owner) {
    throw new Error(`Fixture ${fixture.id} is missing outputs`);
  }
  const ownerSheetName = owner.sheetName ?? defaultSheetName;
  engine.setCellFormula(ownerSheetName, owner.address, fixture.formula.replace(/^=/, ""));

  const explanation = engine.explainCell(ownerSheetName, owner.address);
  if (explanation.mode !== FormulaMode.WasmFastPath) {
    throw new Error(
      `Fixture ${fixture.id} expected wasm fast path, received ${String(explanation.mode)}`,
    );
  }

  const differential = engine.recalculateDifferential();
  if (differential.drift.length > 0) {
    throw new Error(
      `Fixture ${fixture.id} drifted between JS and wasm: ${JSON.stringify(differential.drift)}`,
    );
  }

  for (const output of fixture.outputs) {
    const actual = engine.getCellValue(output.sheetName ?? defaultSheetName, output.address);
    expectCellValueLike(actual, expectedValueToCellValue(output.expected));
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

function expectCellValueLike(actual: CellValue, expected: CellValue): void {
  expect(actual.tag).toBe(expected.tag);
  if (actual.tag === ValueTag.Number && expected.tag === ValueTag.Number) {
    expect(actual.value).toBeCloseTo(expected.value, 7);
    return;
  }
  if (actual.tag === ValueTag.Error && expected.tag === ValueTag.Error) {
    expect(actual.code).toBe(expected.code);
    return;
  }
  if (actual.tag === ValueTag.String && expected.tag === ValueTag.String) {
    expect(actual.value).toBe(expected.value);
    return;
  }
  if (actual.tag === ValueTag.Boolean && expected.tag === ValueTag.Boolean) {
    expect(actual.value).toBe(expected.value);
    return;
  }
  expect(actual).toEqual(expected);
}
