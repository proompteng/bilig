import { ErrorCode, type CellValue, type LiteralInput } from "@bilig/protocol";
import { excelDateTimeFixtureSuite } from "./datetime-fixtures.js";
import { canonicalLogicalFixtures } from "./logical-fixtures.js";
import { canonicalExpansionFixtures } from "./canonical-expansion-fixtures.js";
import { canonicalFoundationFixtures } from "./canonical-foundation-fixtures.js";
import { canonicalTextFixtures } from "./text-fixtures.js";
import { canonicalWorkbookSemanticsFixtures } from "./workbook-semantics-fixtures.js";

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
  "lambda",
] as const;

export type ExcelFixtureFamily = (typeof excelFixtureFamilies)[number];

export const excelFixtureIdPattern = /^[a-z][a-z0-9-]*:[a-z0-9-]+$/;

export type ExcelExpectedValue =
  | { kind: "empty" }
  | { kind: "number"; value: number }
  | { kind: "boolean"; value: boolean }
  | { kind: "string"; value: string }
  | { kind: "error"; code: ErrorCode; display: string };

export interface ExcelFixtureInputCell {
  address: string;
  sheetName?: string;
  input: LiteralInput;
  note?: string;
}

export interface ExcelFixtureDefinedName {
  name: string;
  value: LiteralInput;
  note?: string;
}

export interface ExcelFixtureTable {
  name: string;
  sheetName?: string;
  startAddress: string;
  endAddress: string;
  columnNames: string[];
  headerRow: boolean;
  totalsRow: boolean;
  note?: string;
}

export interface ExcelFixtureExpectedOutput {
  address: string;
  sheetName?: string;
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
  definedNames?: ExcelFixtureDefinedName[];
  tables?: ExcelFixtureTable[];
  inputs: ExcelFixtureInputCell[];
  outputs: ExcelFixtureExpectedOutput[];
}

export interface ExcelFixtureSuite {
  id: string;
  description: string;
  sheets: ExcelFixtureSheet[];
  cases?: readonly ExcelFixtureCase[];
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

function dedupeFixtures(fixtures: readonly ExcelFixtureCase[]): ExcelFixtureCase[] {
  const seen = new Set<string>();
  const output: ExcelFixtureCase[] = [];
  for (const fixture of fixtures) {
    if (seen.has(fixture.id)) {
      continue;
    }
    seen.add(fixture.id);
    output.push(fixture);
  }
  return output;
}

const canonicalCorpusExclusions = new Set<string>([
  "text:case-insensitive-compare",
  "information:value-error-display",
]);

const canonicalBaseFixtureIds = new Set<string>([
  "arithmetic:add-basic",
  "arithmetic:precedence-basic",
  "arithmetic:unary-negation",
  "arithmetic:division-basic",
  "arithmetic:power-basic",
  "arithmetic:percent-operator",
  "comparison:equality-number",
  "comparison:equality-text",
  "comparison:greater-than",
  "comparison:less-than-or-equal",
  "logical:if-basic",
  "logical:and-basic",
  "logical:or-basic",
  "logical:not-basic",
  "aggregation:sum-range",
  "aggregation:avg-range",
  "aggregation:min-range",
  "aggregation:max-range",
  "aggregation:count-range",
  "aggregation:counta-range",
  "math:abs-basic",
  "math:round-basic",
  "math:floor-basic",
  "math:ceiling-basic",
  "math:mod-basic",
  "text:concat-operator",
  "text:concat-function",
  "text:len-basic",
  "date-time:serial-addition",
  "date-time:date-constructor",
  "date-time:today-volatile",
  "lookup-reference:index-basic",
  "lookup-reference:match-exact",
  "lookup-reference:vlookup-exact",
  "lookup-reference:xlookup-exact",
  "statistical:averageif-basic",
  "statistical:countif-basic",
  "information:isblank-basic",
  "information:isnumber-basic",
  "information:istext-basic",
  "dynamic-array:sequence-spill",
  "dynamic-array:sequence-aggregate",
  "dynamic-array:filter-basic",
  "dynamic-array:unique-basic",
  "names:defined-name-scalar",
  "tables:table-total-row-sum",
  "structured-reference:table-column-ref",
  "volatile:rand-basic",
  "lambda:let-basic",
  "lambda:lambda-invoke",
  "lambda:map-basic",
  "logical:if-true-branch",
  "logical:if-condition-error",
  "logical:iferror-catches-any-error",
  "logical:ifna-catches-na-only",
  "logical:and-false-on-empty",
  "logical:or-true-branch",
  "logical:not-number",
  "information:isblank-empty",
  "information:isnumber-number",
  "information:istext-string",
  "text:len-counts-plain-string-length",
]);

const canonicalBaseFixtures = dedupeFixtures([
  ...canonicalFoundationFixtures,
  ...canonicalLogicalFixtures,
  ...canonicalTextFixtures,
  ...(excelDateTimeFixtureSuite.cases ?? []),
]).filter(
  (fixture) =>
    canonicalBaseFixtureIds.has(fixture.id) && !canonicalCorpusExclusions.has(fixture.id),
);

export const canonicalFormulaFixtures: readonly ExcelFixtureCase[] = dedupeFixtures([
  ...canonicalBaseFixtures,
  ...canonicalExpansionFixtures,
]);

export const workbookSemanticsFormulaFixtureSuite: ExcelFixtureSuite = {
  id: "workbook-semantics",
  description:
    "Extended workbook semantics fixture slice for names and cross-sheet reference behavior.",
  sheets: [{ name: "Sheet1" }, { name: "Sheet2" }],
  excelBuild: "Microsoft 365 / 2026-03-19",
  capturedAt: "2026-03-19T00:00:00.000Z",
  cases: canonicalWorkbookSemanticsFixtures,
};

export const canonicalFormulaSmokeSuite: ExcelFixtureSuite = {
  id: "canonical-smoke",
  description: "Representative smoke slice from the canonical formula compatibility corpus.",
  sheets: [{ name: "Sheet1" }],
  excelBuild: "Microsoft 365 / 2026-03-19",
  capturedAt: "2026-03-19T00:00:00.000Z",
  cases: canonicalFormulaFixtures.slice(0, 5),
};

function buildFamilySuite(
  id: string,
  description: string,
  families: readonly ExcelFixtureFamily[],
): ExcelFixtureSuite {
  return {
    id,
    description,
    sheets: [{ name: "Sheet1" }],
    excelBuild: "Microsoft 365 / 2026-03-19",
    capturedAt: "2026-03-19T00:00:00.000Z",
    cases: canonicalFormulaFixtures.filter((fixture) => families.includes(fixture.family)),
  };
}

export const textFormulaFixtureSuite = buildFamilySuite(
  "canonical-text",
  "Canonical formula corpus text-function fixture slice.",
  ["text"],
);

export const lookupReferenceFormulaFixtureSuite = buildFamilySuite(
  "canonical-lookup-reference",
  "Canonical formula corpus lookup/reference fixture slice.",
  ["lookup-reference", "statistical"],
);

export const dateTimeFormulaFixtureSuite = buildFamilySuite(
  "canonical-date-time",
  "Canonical formula corpus date/time and volatile fixture slice.",
  ["date-time", "volatile"],
);

export const dynamicArrayFormulaFixtureSuite = buildFamilySuite(
  "canonical-dynamic-array",
  "Canonical formula corpus dynamic-array fixture slice.",
  ["dynamic-array"],
);

export const namesTablesFormulaFixtureSuite = buildFamilySuite(
  "canonical-names-tables",
  "Canonical formula corpus names, tables, and structured-reference fixture slice.",
  ["names", "tables", "structured-reference"],
);

export const lambdaFormulaFixtureSuite = buildFamilySuite(
  "canonical-lambda",
  "Canonical formula corpus lambda fixture slice.",
  ["lambda"],
);

export { canonicalFoundationFixtures } from "./canonical-foundation-fixtures.js";
export { canonicalExpansionFixtures } from "./canonical-expansion-fixtures.js";
export { canonicalWorkbookSemanticsFixtures } from "./workbook-semantics-fixtures.js";
export * from "./logical-fixtures.js";
export * from "./text-fixtures.js";
export * from "./datetime-fixtures.js";
