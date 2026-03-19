import { ErrorCode } from "@bilig/protocol";
import type {
  ExcelExpectedValue,
  ExcelFixtureCase,
  ExcelFixtureExpectedOutput,
  ExcelFixtureFamily,
  ExcelFixtureInputCell
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

export const canonicalExpansionFixtures: readonly ExcelFixtureCase[] = [
  fixture("text", "exact-basic", "EXACT case-sensitive text comparison", "=EXACT(\"Alpha\",\"alpha\")", [], [
    output("A1", { kind: "boolean", value: false })
  ]),
  fixture("text", "left-basic", "LEFT basic prefix extraction", "=LEFT(\"alpha\",3)", [], [output("A1", stringExpected("alp"))]),
  fixture("text", "right-basic", "RIGHT basic suffix extraction", "=RIGHT(\"alpha\",2)", [], [output("A1", stringExpected("ha"))]),
  fixture("text", "mid-basic", "MID basic substring extraction", "=MID(\"alphabet\",2,3)", [], [output("A1", stringExpected("lph"))]),
  fixture("text", "trim-basic", "TRIM removes extra spaces", "=TRIM(\"  alpha   beta  \")", [], [output("A1", stringExpected("alpha beta"))]),
  fixture("text", "upper-basic", "UPPER uppercases text", "=UPPER(\"Alpha beta\")", [], [output("A1", stringExpected("ALPHA BETA"))]),
  fixture("text", "lower-basic", "LOWER lowercases text", "=LOWER(\"Alpha BETA\")", [], [output("A1", stringExpected("alpha beta"))]),
  fixture("text", "find-basic", "FIND basic case-sensitive search", "=FIND(\"ph\",\"alphabet\")", [], [output("A1", numberExpected(3))]),
  fixture("text", "search-basic", "SEARCH basic case-insensitive search", "=SEARCH(\"PH\",\"alphabet\")", [], [output("A1", numberExpected(3))]),
  fixture("text", "value-basic", "VALUE text-to-number coercion", "=VALUE(\"42\")", [], [output("A1", numberExpected(42))], "Canonical literal case now closes through constant folding; dynamic text-input VALUE semantics still route through the JS oracle."),
  fixture("lookup-reference", "xmatch-basic", "XMATCH exact match", "=XMATCH(\"pear\",A1:A3,0)", [input("A1", "apple"), input("A2", "pear"), input("A3", "plum")], [output("A4", numberExpected(2))]),
  fixture("lookup-reference", "hlookup-basic", "HLOOKUP exact match across header row", "=HLOOKUP(\"pear\",A1:C2,2,FALSE)", [input("A1", "apple"), input("B1", "pear"), input("C1", "plum"), input("A2", 10), input("B2", 20), input("C2", 30)], [output("D1", numberExpected(20))]),
  fixture("lookup-reference", "offset-basic", "OFFSET relative reference", "=OFFSET(A1,1,1)", [input("A1", 10), input("B2", 20)], [output("A2", numberExpected(20))], "Canonical corpus case; reference-returning semantics are still open."),
  fixture("dynamic-array", "take-basic", "TAKE returns leading rows", "=TAKE(A1:A4,2)", [input("A1", 1), input("A2", 2), input("A3", 3), input("A4", 4)], [output("B1", numberExpected(1)), output("B2", numberExpected(2))]),
  fixture("dynamic-array", "drop-basic", "DROP removes leading rows", "=DROP(A1:A4,2)", [input("A1", 1), input("A2", 2), input("A3", 3), input("A4", 4)], [output("B1", numberExpected(3)), output("B2", numberExpected(4))]),
  fixture("dynamic-array", "choosecols-basic", "CHOOSECOLS selects explicit columns", "=CHOOSECOLS(A1:C2,1,3)", [input("A1", 1), input("B1", 2), input("C1", 3), input("A2", 4), input("B2", 5), input("C2", 6)], [output("D1", numberExpected(1)), output("E1", numberExpected(3)), output("D2", numberExpected(4)), output("E2", numberExpected(6))]),
  fixture("dynamic-array", "chooserows-basic", "CHOOSEROWS selects explicit rows", "=CHOOSEROWS(A1:B3,1,3)", [input("A1", 1), input("B1", 2), input("A2", 3), input("B2", 4), input("A3", 5), input("B3", 6)], [output("D1", numberExpected(1)), output("E1", numberExpected(2)), output("D2", numberExpected(5)), output("E2", numberExpected(6))]),
  fixture("statistical", "sumif-basic", "SUMIF with scalar criteria", "=SUMIF(A1:A4,\">0\",B1:B4)", [input("A1", 2), input("A2", -1), input("A3", 4), input("A4", 0), input("B1", 10), input("B2", 20), input("B3", 30), input("B4", 40)], [output("C1", numberExpected(40))]),
  fixture("statistical", "sumifs-basic", "SUMIFS with paired criteria ranges", "=SUMIFS(C1:C4,A1:A4,\">0\",B1:B4,\"x\")", [input("A1", 2), input("A2", -1), input("A3", 4), input("A4", 7), input("B1", "x"), input("B2", "x"), input("B3", "y"), input("B4", "x"), input("C1", 10), input("C2", 20), input("C3", 30), input("C4", 40)], [output("D1", numberExpected(50))]),
  fixture("statistical", "averageifs-basic", "AVERAGEIFS with paired criteria ranges", "=AVERAGEIFS(C1:C4,A1:A4,\">0\",B1:B4,\"x\")", [input("A1", 2), input("A2", -1), input("A3", 4), input("A4", 7), input("B1", "x"), input("B2", "x"), input("B3", "y"), input("B4", "x"), input("C1", 10), input("C2", 20), input("C3", 30), input("C4", 40)], [output("D1", numberExpected(25))]),
  fixture("statistical", "countifs-basic", "COUNTIFS with paired criteria ranges", "=COUNTIFS(A1:A4,\">0\",B1:B4,\"x\")", [input("A1", 2), input("A2", -1), input("A3", 4), input("A4", 7), input("B1", "x"), input("B2", "x"), input("B3", "y"), input("B4", "x")], [output("C1", numberExpected(2))]),
  fixture("math", "sumproduct-basic", "SUMPRODUCT multiplies and sums aligned arrays", "=SUMPRODUCT(A1:A3,B1:B3)", [input("A1", 1), input("A2", 2), input("A3", 3), input("B1", 4), input("B2", 5), input("B3", 6)], [output("C1", numberExpected(32))]),
  fixture("math", "int-basic", "INT rounds toward negative infinity", "=INT(-3.1)", [], [output("A1", numberExpected(-4))]),
  fixture("math", "roundup-basic", "ROUNDUP rounds away from zero", "=ROUNDUP(12.341,2)", [], [output("A1", numberExpected(12.35))]),
  fixture("math", "rounddown-basic", "ROUNDDOWN rounds toward zero", "=ROUNDDOWN(12.349,2)", [], [output("A1", numberExpected(12.34))]),
  fixture("date-time", "now-volatile", "NOW volatile timestamp capture", "=NOW()", [], [output("A1", numberExpected(46100.65659722222))], "Oracle output captured for the canonical corpus plan; runtime normalization is still open."),
  fixture("date-time", "time-basic", "TIME constructs a fractional day serial", "=TIME(12,30,0)", [], [output("A1", numberExpected(0.5208333333333334))]),
  fixture("date-time", "hour-basic", "HOUR extracts the hour component", "=HOUR(A1)", [input("A1", 0.5208333333333334)], [output("A2", numberExpected(12))]),
  fixture("date-time", "minute-basic", "MINUTE extracts the minute component", "=MINUTE(A1)", [input("A1", 0.5208333333333334)], [output("A2", numberExpected(30))]),
  fixture("date-time", "second-basic", "SECOND extracts the second component", "=SECOND(A1)", [input("A1", 0.5208449074074074)], [output("A2", numberExpected(1))]),
  fixture("date-time", "weekday-basic", "WEEKDAY returns the default weekday number", "=WEEKDAY(DATE(2026,3,15))", [], [output("A1", numberExpected(1))]),
  fixture("dynamic-array", "sort-basic", "SORT reorders a vertical range", "=SORT(A1:A4)", [input("A1", 3), input("A2", 1), input("A3", 4), input("A4", 2)], [output("B1", numberExpected(1)), output("B2", numberExpected(2)), output("B3", numberExpected(3)), output("B4", numberExpected(4))]),
  fixture("dynamic-array", "sortby-basic", "SORTBY orders values by companion keys", "=SORTBY(A1:A3,B1:B3)", [input("A1", "pear"), input("A2", "apple"), input("A3", "plum"), input("B1", 2), input("B2", 1), input("B3", 3)], [output("C1", stringExpected("apple")), output("C2", stringExpected("pear")), output("C3", stringExpected("plum"))]),
  fixture("dynamic-array", "tocol-basic", "TOCOL flattens a matrix by columns", "=TOCOL(A1:B2)", [input("A1", 1), input("B1", 2), input("A2", 3), input("B2", 4)], [output("C1", numberExpected(1)), output("C2", numberExpected(3)), output("C3", numberExpected(2)), output("C4", numberExpected(4))]),
  fixture("dynamic-array", "torow-basic", "TOROW flattens a matrix by rows", "=TOROW(A1:B2)", [input("A1", 1), input("B1", 2), input("A2", 3), input("B2", 4)], [output("C1", numberExpected(1)), output("D1", numberExpected(2)), output("E1", numberExpected(3)), output("F1", numberExpected(4))]),
  fixture("dynamic-array", "wraprows-basic", "WRAPROWS wraps a vector into rows", "=WRAPROWS(A1:A4,2)", [input("A1", 1), input("A2", 2), input("A3", 3), input("A4", 4)], [output("B1", numberExpected(1)), output("C1", numberExpected(2)), output("B2", numberExpected(3)), output("C2", numberExpected(4))]),
  fixture("dynamic-array", "wrapcols-basic", "WRAPCOLS wraps a vector into columns", "=WRAPCOLS(A1:A4,2)", [input("A1", 1), input("A2", 2), input("A3", 3), input("A4", 4)], [output("B1", numberExpected(1)), output("B2", numberExpected(2)), output("C1", numberExpected(3)), output("C2", numberExpected(4))]),
  fixture("names", "defined-name-range", "Defined-name range reference", "=SUM(MyRange)", [input("A1", 10), input("A2", 12), input("A3", 15)], [output("A4", numberExpected(37))], "Requires workbook-level name metadata in the oracle workbook."),
  fixture("lambda", "byrow-basic", "BYROW applies a lambda to each row", "=BYROW(A1:B2,LAMBDA(r,SUM(r)))", [input("A1", 1), input("B1", 2), input("A2", 3), input("B2", 4)], [output("C1", numberExpected(3)), output("C2", numberExpected(7))])
];
