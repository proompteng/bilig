import { ErrorCode, ValueTag, type CellValue } from "@bilig/protocol";
import type { ExcelExpectedValue, ExcelFixtureCase, ExcelFixtureExpectedOutput, ExcelFixtureFamily, ExcelFixtureInputCell } from "./index.js";

export interface LogicalFixtureCase {
  id: string;
  functionName: "IF" | "IFERROR" | "IFNA" | "AND" | "OR" | "NOT" | "ISBLANK" | "ISNUMBER" | "ISTEXT";
  args: CellValue[];
  expected: CellValue;
  notes: string;
  source: "excel-support-docs" | "bilig-contract";
}

const excelFixtureIdPattern = /^[a-z][a-z0-9-]*:[a-z0-9-]+$/;

function createExcelFixtureId(family: ExcelFixtureFamily, slug: string): string {
  const normalizedSlug = slug.trim().toLowerCase();
  const id = `${family}:${normalizedSlug}`;
  if (!excelFixtureIdPattern.test(id)) {
    throw new Error(`Invalid Excel fixture id: ${id}`);
  }
  return id;
}

function booleanExpected(value: boolean): ExcelExpectedValue {
  return { kind: "boolean", value };
}

function stringExpected(value: string): ExcelExpectedValue {
  return { kind: "string", value };
}

function errorExpected(code: ErrorCode, display: string): ExcelExpectedValue {
  return { kind: "error", code, display };
}

function input(address: string, value: ExcelFixtureInputCell["input"], note?: string): ExcelFixtureInputCell {
  return note === undefined ? { address, input: value } : { address, input: value, note };
}

function output(address: string, expected: ExcelFixtureExpectedOutput["expected"], note?: string): ExcelFixtureExpectedOutput {
  return note === undefined ? { address, expected } : { address, expected, note };
}

function fixture(
  family: ExcelFixtureFamily,
  slug: string,
  title: string,
  formula: string,
  inputs: ExcelFixtureInputCell[],
  outputs: ExcelFixtureExpectedOutput[],
  notes?: string
): ExcelFixtureCase {
  const base = {
    id: createExcelFixtureId(family, slug),
    family,
    title,
    formula,
    inputs,
    outputs,
    sheetName: "Sheet1"
  };
  return notes === undefined ? base : { ...base, notes };
}

const empty = (): CellValue => ({ tag: ValueTag.Empty });
const bool = (value: boolean): CellValue => ({ tag: ValueTag.Boolean, value });
const num = (value: number): CellValue => ({ tag: ValueTag.Number, value });
const text = (value: string): CellValue => ({ tag: ValueTag.String, value, stringId: 0 });
const err = (code: ErrorCode): CellValue => ({ tag: ValueTag.Error, code });

export const logicalFixtureMetadata = {
  group: "logical-info",
  baseline: "Excel for the web desktop semantics, normalized into bilig CellValue tags",
  coverage: ["IF", "IFERROR", "IFNA", "AND", "OR", "NOT", "ISBLANK", "ISNUMBER", "ISTEXT"],
  notes: [
    "Error handling stays explicit in bilig CellValue form rather than formatted strings.",
    "Empty arguments are represented as ValueTag.Empty instead of worksheet references."
  ]
} as const;

export const logicalFixtureCases: readonly LogicalFixtureCase[] = [
  {
    id: "if-true-branch",
    functionName: "IF",
    args: [bool(true), text("yes"), text("no")],
    expected: text("yes"),
    notes: "Boolean TRUE selects the true branch.",
    source: "excel-support-docs"
  },
  {
    id: "if-condition-error",
    functionName: "IF",
    args: [err(ErrorCode.Ref), num(1), num(2)],
    expected: err(ErrorCode.Ref),
    notes: "Condition errors propagate instead of masking.",
    source: "bilig-contract"
  },
  {
    id: "iferror-catches-any-error",
    functionName: "IFERROR",
    args: [err(ErrorCode.Div0), text("fallback")],
    expected: text("fallback"),
    notes: "IFERROR replaces non-NA and NA errors alike.",
    source: "excel-support-docs"
  },
  {
    id: "ifna-catches-na-only",
    functionName: "IFNA",
    args: [err(ErrorCode.NA), text("missing")],
    expected: text("missing"),
    notes: "IFNA handles #N/A and leaves other errors alone.",
    source: "excel-support-docs"
  },
  {
    id: "and-false-on-empty",
    functionName: "AND",
    args: [bool(true), empty()],
    expected: bool(false),
    notes: "Current bilig scalar coercion treats empty as FALSE.",
    source: "bilig-contract"
  },
  {
    id: "or-true-branch",
    functionName: "OR",
    args: [empty(), bool(true)],
    expected: bool(true),
    notes: "Any truthy logical argument yields TRUE.",
    source: "excel-support-docs"
  },
  {
    id: "not-number",
    functionName: "NOT",
    args: [num(2)],
    expected: bool(false),
    notes: "Non-zero numbers coerce to TRUE before inversion.",
    source: "bilig-contract"
  },
  {
    id: "isblank-empty",
    functionName: "ISBLANK",
    args: [empty()],
    expected: bool(true),
    notes: "Only the Empty tag counts as blank.",
    source: "excel-support-docs"
  },
  {
    id: "isnumber-number",
    functionName: "ISNUMBER",
    args: [num(42)],
    expected: bool(true),
    notes: "Numbers are detected without coercing text.",
    source: "excel-support-docs"
  },
  {
    id: "istext-string",
    functionName: "ISTEXT",
    args: [text("hello")],
    expected: bool(true),
    notes: "Text detection is based on the string tag.",
    source: "excel-support-docs"
  }
];

export const canonicalLogicalFixtures: readonly ExcelFixtureCase[] = [
  fixture("logical", "if-true-branch", "IF true branch from boolean input", "=IF(A1,A2,A3)", [input("A1", true), input("A2", "yes"), input("A3", "no")], [output("A4", stringExpected("yes"))]),
  fixture("logical", "if-condition-error", "IF propagates condition error", "=IF(1/0,1,2)", [], [output("A1", errorExpected(ErrorCode.Div0, "#DIV/0!"))]),
  fixture("logical", "iferror-catches-any-error", "IFERROR catches any error", "=IFERROR(1/0,\"fallback\")", [], [output("A1", stringExpected("fallback"))]),
  fixture("logical", "ifna-catches-na-only", "IFNA catches #N/A only", "=IFNA(NA(),\"missing\")", [], [output("A1", stringExpected("missing"))]),
  fixture("logical", "and-false-on-empty", "AND treats empty as FALSE", "=AND(TRUE,A1)", [input("A1", null)], [output("A2", booleanExpected(false))]),
  fixture("logical", "or-true-branch", "OR returns TRUE when any operand is TRUE", "=OR(A1,TRUE)", [input("A1", null)], [output("A2", booleanExpected(true))]),
  fixture("logical", "not-number", "NOT coerces non-zero numbers to TRUE before inversion", "=NOT(2)", [], [output("A1", booleanExpected(false))]),
  fixture("information", "isblank-empty", "ISBLANK empty input", "=ISBLANK(A1)", [input("A1", null)], [output("A2", booleanExpected(true))]),
  fixture("information", "isnumber-number", "ISNUMBER numeric literal", "=ISNUMBER(42)", [], [output("A1", booleanExpected(true))]),
  fixture("information", "istext-string", "ISTEXT string literal", "=ISTEXT(\"hello\")", [], [output("A1", booleanExpected(true))])
];
