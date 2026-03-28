import { ErrorCode } from "@bilig/protocol";
import type {
  ExcelExpectedValue,
  ExcelFixtureCase,
  ExcelFixtureDefinedName,
  ExcelFixtureExpectedOutput,
  ExcelFixtureFamily,
  ExcelFixtureInputCell,
  ExcelFixtureMultipleOperationsMock,
} from "./index.js";

const excelFixtureIdPattern = /^[a-z][a-z0-9-]*:[a-z0-9-]+$/;

function createExcelFixtureId(family: ExcelFixtureFamily, slug: string): string {
  const normalizedSlug = slug.trim().toLowerCase();
  const id = `${family}:${normalizedSlug}`;
  if (!excelFixtureIdPattern.test(id)) {
    throw new Error(`Invalid Excel fixture id: ${id}`);
  }
  return id;
}

function numberExpected(value: number): ExcelExpectedValue {
  return { kind: "number", value };
}

function errorExpected(code: ErrorCode, display: string): ExcelExpectedValue {
  return { kind: "error", code, display };
}

function definedName(
  name: string,
  value: ExcelFixtureDefinedName["value"],
  note?: string,
): ExcelFixtureDefinedName {
  return note === undefined ? { name, value } : { name, value, note };
}

function input(
  address: string,
  value: ExcelFixtureInputCell["input"],
  options: { note?: string; sheetName?: string } = {},
): ExcelFixtureInputCell {
  const base: ExcelFixtureInputCell = { address, input: value };
  if (options.note !== undefined) {
    base.note = options.note;
  }
  if (options.sheetName !== undefined) {
    base.sheetName = options.sheetName;
  }
  return base;
}

function output(
  address: string,
  expected: ExcelFixtureExpectedOutput["expected"],
  note?: string,
): ExcelFixtureExpectedOutput {
  return note === undefined ? { address, expected } : { address, expected, note };
}

function fixture(
  family: ExcelFixtureFamily,
  slug: string,
  title: string,
  formula: string,
  inputs: ExcelFixtureInputCell[],
  outputs: ExcelFixtureExpectedOutput[],
  options: { notes?: string; definedNames?: ExcelFixtureDefinedName[] } = {},
): ExcelFixtureCase {
  const base: ExcelFixtureCase = {
    id: createExcelFixtureId(family, slug),
    family,
    title,
    formula,
    inputs,
    outputs,
    sheetName: "Sheet1",
  };
  if (options.notes !== undefined) {
    base.notes = options.notes;
  }
  if (options.definedNames !== undefined) {
    base.definedNames = options.definedNames;
  }
  return base;
}

function multipleOperations(
  config: ExcelFixtureMultipleOperationsMock,
): ExcelFixtureMultipleOperationsMock {
  return config;
}

const taxRate = definedName("TaxRate", 0.085);
const ratePack = [definedName("TaxRate", 0.085), definedName("FeeRate", 0.015)];

export const canonicalWorkbookSemanticsFixtures: readonly ExcelFixtureCase[] = [
  fixture(
    "names",
    "defined-name-case-insensitive",
    "Defined names resolve case-insensitively",
    "=taxrate*A1",
    [input("A1", 100)],
    [output("A2", numberExpected(8.5))],
    {
      definedNames: [taxRate],
      notes:
        "Workbook-level names follow the same case-insensitive lookup contract as the engine metadata store.",
    },
  ),
  fixture(
    "names",
    "defined-name-multi-scalar-pack",
    "Multiple workbook names participate in one scalar expression",
    "=TaxRate+FeeRate",
    [],
    [output("A1", numberExpected(0.1))],
    {
      definedNames: ratePack,
      notes:
        "Captures the common workbook metadata case where more than one scalar name is bound at once.",
    },
  ),
  fixture(
    "names",
    "defined-name-missing",
    "Missing workbook names surface #NAME?",
    "=MissingRate*A1",
    [input("A1", 100)],
    [output("A2", errorExpected(ErrorCode.Name, "#NAME?"))],
    {
      notes: "Keeps fixture coverage aligned with the JS evaluator and engine name-miss behavior.",
    },
  ),
  fixture(
    "arithmetic",
    "cross-sheet-multiply",
    "Cross-sheet scalar arithmetic",
    "=Sheet2!B1*3",
    [input("B1", 4, { sheetName: "Sheet2" })],
    [output("A1", numberExpected(12))],
    {
      notes:
        "Exercises qualified cell references without introducing dynamic-array or table metadata dependencies.",
    },
  ),
  fixture(
    "arithmetic",
    "cross-sheet-empty-cell-zero",
    "Cross-sheet empty cells coerce like blank worksheet cells",
    "=Sheet2!B1*3",
    [
      input("B1", null, {
        sheetName: "Sheet2",
        note: "Marks Sheet2 as present while leaving the referenced cell empty.",
      }),
    ],
    [output("A1", numberExpected(0))],
  ),
  fixture(
    "arithmetic",
    "missing-sheet-ref-error",
    "Missing cross-sheet cells surface #REF!",
    "=Sheet2!B1*3",
    [],
    [output("A1", errorExpected(ErrorCode.Ref, "#REF!"))],
    {
      notes:
        "The target worksheet is intentionally absent so the harness can capture unresolved-sheet error behavior.",
    },
  ),
  fixture(
    "aggregation",
    "cross-sheet-range-sum",
    "Cross-sheet range aggregation",
    "=SUM(Sheet2!A1:A2)",
    [input("A1", 2, { sheetName: "Sheet2" }), input("A2", 3, { sheetName: "Sheet2" })],
    [output("A1", numberExpected(5))],
    {
      notes: "Matches the native rebinding coverage that already exists in the core engine tests.",
    },
  ),
  fixture(
    "aggregation",
    "cross-sheet-empty-range-zero",
    "Cross-sheet empty ranges aggregate as zero once the sheet exists",
    "=SUM(Sheet2!A1:A2)",
    [input("A1", null, { sheetName: "Sheet2" }), input("A2", null, { sheetName: "Sheet2" })],
    [output("A1", numberExpected(0))],
  ),
  fixture(
    "aggregation",
    "missing-sheet-range-ref-error",
    "Missing cross-sheet ranges surface #REF!",
    "=SUM(Sheet2!A1:A2)",
    [],
    [output("A1", errorExpected(ErrorCode.Ref, "#REF!"))],
    {
      notes:
        "Range-qualified sheet misses stay JS-authoritative until a later sheet creation can rebind them.",
    },
  ),
  {
    id: createExcelFixtureId("lookup-reference", "multiple-operations-basic"),
    family: "lookup-reference",
    title: "MULTIPLE.OPERATIONS resolves workbook-aware what-if substitutions",
    formula: "=MULTIPLE.OPERATIONS(B5,B3,C4,B2,D2)",
    inputs: [input("B2", 1), input("B3", 2), input("C4", 5), input("D2", 3)],
    outputs: [output("A1", numberExpected(23))],
    sheetName: "Sheet1",
    notes:
      "The fixture harness stubs the what-if request contract while the engine integration test covers recursive workbook recomputation with real dependent formulas.",
    multipleOperations: multipleOperations({
      formulaAddress: "B5",
      rowCellAddress: "B3",
      rowReplacementAddress: "C4",
      columnCellAddress: "B2",
      columnReplacementAddress: "D2",
      result: numberExpected(23),
    }),
  },
];
