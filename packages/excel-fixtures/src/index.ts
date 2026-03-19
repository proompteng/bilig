import { ErrorCode, type CellValue, type LiteralInput } from "@bilig/protocol";
import { excelTop50StarterFixtures } from "./top50.js";

export const excelFixtureFamilies = [
  "arithmetic",
  "comparison",
  "logical",
  "aggregation",
  "math",
  "text",
  "date-time",
  "lookup-reference",
  "statistical",
  "information",
  "dynamic-array",
  "names",
  "tables",
  "structured-reference",
  "volatile",
  "lambda"
] as const;

export type ExcelFixtureFamily = typeof excelFixtureFamilies[number];

export const excelFixtureIdPattern = /^[a-z][a-z0-9-]*:[a-z0-9-]+$/;

export type ExcelExpectedValue =
  | { kind: "empty" }
  | { kind: "number"; value: number }
  | { kind: "boolean"; value: boolean }
  | { kind: "string"; value: string }
  | { kind: "error"; code: ErrorCode; display: string };

export interface ExcelFixtureInputCell {
  address: string;
  input: LiteralInput;
  note?: string;
}

export interface ExcelFixtureExpectedOutput {
  address: string;
  expected: ExcelExpectedValue;
  note?: string;
}

export interface ExcelFixtureCell {
  address: string;
  formula?: string;
  input?: LiteralInput;
  expected: ExcelExpectedValue | CellValue;
}

export interface ExcelFixtureSheet {
  name: string;
  cells?: ExcelFixtureCell[];
}

export interface ExcelFixtureCase {
  id: string;
  family: ExcelFixtureFamily;
  title: string;
  formula: string;
  notes?: string;
  sheetName?: string;
  inputs: ExcelFixtureInputCell[];
  outputs: ExcelFixtureExpectedOutput[];
}

export interface ExcelFixtureSuite {
  id: string;
  description: string;
  sheets: ExcelFixtureSheet[];
  cases?: ExcelFixtureCase[];
  excelBuild: string;
  capturedAt: string;
}

export function createExcelFixtureId(family: ExcelFixtureFamily, slug: string): string {
  const normalizedSlug = slug.trim().toLowerCase();
  if (!/^[a-z0-9-]+$/.test(normalizedSlug)) {
    throw new Error(`Invalid Excel fixture slug: ${slug}`);
  }
  const id = `${family}:${normalizedSlug}`;
  if (!excelFixtureIdPattern.test(id)) {
    throw new Error(`Invalid Excel fixture id: ${id}`);
  }
  return id;
}

export function isExcelFixtureId(value: string): boolean {
  return excelFixtureIdPattern.test(value);
}

export function emptyExpected(): ExcelExpectedValue {
  return { kind: "empty" };
}

export function numberExpected(value: number): ExcelExpectedValue {
  return { kind: "number", value };
}

export function booleanExpected(value: boolean): ExcelExpectedValue {
  return { kind: "boolean", value };
}

export function stringExpected(value: string): ExcelExpectedValue {
  return { kind: "string", value };
}

export function errorExpected(code: ErrorCode, display: string): ExcelExpectedValue {
  return { kind: "error", code, display };
}

export const excelFixtureSmokeSuite: ExcelFixtureSuite = {
  id: "top50-smoke",
  description: "Representative starter slice from the Top 50 Excel compatibility registry.",
  sheets: [{ name: "Sheet1" }],
  excelBuild: "Microsoft 365 / 2026-03-15",
  capturedAt: "2026-03-15T00:00:00.000Z",
  cases: excelTop50StarterFixtures.slice(0, 5)
};

export { excelTop50StarterFixtures } from "./top50.js";
export * from "./logical-fixtures.js";
export * from "./text-fixtures.js";
export * from "./datetime-fixtures.js";
