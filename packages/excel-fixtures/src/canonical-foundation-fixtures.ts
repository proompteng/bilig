import { ErrorCode } from "@bilig/protocol";
import type {
  ExcelExpectedValue,
  ExcelFixtureCase,
  ExcelFixtureFamily,
  ExcelFixtureExpectedOutput,
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

export const canonicalFoundationFixtures: ExcelFixtureCase[] = [
  fixture("arithmetic", "add-basic", "Addition", "=A1+A2", [input("A1", 2), input("A2", 5)], [output("A3", numberExpected(7))]),
  fixture("arithmetic", "precedence-basic", "Operator precedence", "=A1+A2*A3", [input("A1", 2), input("A2", 3), input("A3", 4)], [output("A4", numberExpected(14))]),
  fixture("arithmetic", "unary-negation", "Unary negation", "=-A1", [input("A1", 9)], [output("A2", numberExpected(-9))]),
  fixture("arithmetic", "division-basic", "Division", "=A1/A2", [input("A1", 12), input("A2", 3)], [output("A3", numberExpected(4))]),
  fixture("arithmetic", "power-basic", "Exponentiation", "=A1^A2", [input("A1", 2), input("A2", 3)], [output("A3", numberExpected(8))]),
  fixture("arithmetic", "percent-operator", "Percent postfix operator", "=A1*10%", [input("A1", 50)], [output("A2", numberExpected(5))], "Tracks unsupported postfix percent grammar until parity lands."),
  fixture("comparison", "equality-number", "Numeric equality", "=A1=A2", [input("A1", 7), input("A2", 7)], [output("A3", booleanExpected(true))]),
  fixture("comparison", "equality-text", "Case-insensitive text equality", "=\"hello\"=\"HELLO\"", [], [output("A1", booleanExpected(true))]),
  fixture("comparison", "greater-than", "Greater-than", "=A1>A2", [input("A1", 9), input("A2", 4)], [output("A3", booleanExpected(true))]),
  fixture("comparison", "less-than-or-equal", "Less-than-or-equal", "=A1<=A2", [input("A1", 4), input("A2", 4)], [output("A3", booleanExpected(true))]),
  fixture("logical", "if-basic", "IF true branch", "=IF(A1>0,\"yes\",\"no\")", [input("A1", 3)], [output("A2", stringExpected("yes"))]),
  fixture("logical", "and-basic", "AND", "=AND(A1,A2)", [input("A1", true), input("A2", true)], [output("A3", booleanExpected(true))]),
  fixture("logical", "or-basic", "OR", "=OR(A1,A2)", [input("A1", false), input("A2", true)], [output("A3", booleanExpected(true))]),
  fixture("logical", "not-basic", "NOT", "=NOT(A1)", [input("A1", false)], [output("A2", booleanExpected(true))]),
  fixture("aggregation", "sum-range", "SUM over range", "=SUM(A1:A3)", [input("A1", 2), input("A2", 3), input("A3", 4)], [output("A4", numberExpected(9))]),
  fixture("aggregation", "avg-range", "AVG over range", "=AVG(A1:A3)", [input("A1", 2), input("A2", 4), input("A3", 6)], [output("A4", numberExpected(4))]),
  fixture("aggregation", "min-range", "MIN over range", "=MIN(A1:A3)", [input("A1", 8), input("A2", 2), input("A3", 5)], [output("A4", numberExpected(2))]),
  fixture("aggregation", "max-range", "MAX over range", "=MAX(A1:A3)", [input("A1", 8), input("A2", 2), input("A3", 5)], [output("A4", numberExpected(8))]),
  fixture("aggregation", "count-range", "COUNT over range", "=COUNT(A1:A4)", [input("A1", 8), input("A2", 2), input("A3", null), input("A4", "x")], [output("A5", numberExpected(2))]),
  fixture("aggregation", "counta-range", "COUNTA over range", "=COUNTA(A1:A4)", [input("A1", 8), input("A2", 2), input("A3", null), input("A4", "x")], [output("A5", numberExpected(3))]),
  fixture("math", "abs-basic", "ABS", "=ABS(A1)", [input("A1", -7)], [output("A2", numberExpected(7))]),
  fixture("math", "round-basic", "ROUND", "=ROUND(A1,1)", [input("A1", 3.125)], [output("A2", numberExpected(3.1))]),
  fixture("math", "floor-basic", "FLOOR", "=FLOOR(A1,2)", [input("A1", 7)], [output("A2", numberExpected(6))]),
  fixture("math", "ceiling-basic", "CEILING", "=CEILING(A1,2)", [input("A1", 7)], [output("A2", numberExpected(8))]),
  fixture("math", "mod-basic", "MOD", "=MOD(A1,A2)", [input("A1", 10), input("A2", 3)], [output("A3", numberExpected(1))]),
  fixture("text", "concat-operator", "Concatenation operator", "=\"bi\"&\"lig\"", [], [output("A1", stringExpected("bilig"))]),
  fixture("text", "concat-function", "CONCAT function", "=CONCAT(\"bi\",\"lig\")", [], [output("A1", stringExpected("bilig"))]),
  fixture("text", "len-basic", "LEN", "=LEN(\"bilig\")", [], [output("A1", numberExpected(5))]),
  fixture("text", "case-insensitive-compare", "Text compare semantics", "=\"a\"=\"A\"", [], [output("A1", booleanExpected(true))]),
  fixture("date-time", "serial-addition", "Date serial addition", "=A1+7", [input("A1", 45292)], [output("A2", numberExpected(45299))], "Uses raw serials so the oracle stays stable across locale formatting."),
  fixture("date-time", "date-constructor", "DATE constructor", "=DATE(2026,3,15)", [], [output("A1", numberExpected(46096))], "Committed oracle result captured from Excel for the web."),
  fixture("date-time", "today-volatile", "TODAY volatile function", "=TODAY()", [], [output("A1", numberExpected(46100))], "Captured output is pinned to the current Excel-for-web oracle timestamp rather than treated as timeless."),
  fixture("lookup-reference", "index-basic", "INDEX exact cell lookup", "=INDEX(A1:B2,2,1)", [input("A1", 10), input("B1", 20), input("A2", 30), input("B2", 40)], [output("C1", numberExpected(30))]),
  fixture("lookup-reference", "match-exact", "MATCH exact", "=MATCH(\"pear\",A1:A3,0)", [input("A1", "apple"), input("A2", "pear"), input("A3", "plum")], [output("A4", numberExpected(2))]),
  fixture("lookup-reference", "vlookup-exact", "VLOOKUP exact", "=VLOOKUP(\"pear\",A1:B3,2,FALSE)", [input("A1", "apple"), input("B1", 10), input("A2", "pear"), input("B2", 20), input("A3", "plum"), input("B3", 30)], [output("C1", numberExpected(20))]),
  fixture("lookup-reference", "xlookup-exact", "XLOOKUP exact", "=XLOOKUP(\"pear\",A1:A3,B1:B3)", [input("A1", "apple"), input("B1", 10), input("A2", "pear"), input("B2", 20), input("A3", "plum"), input("B3", 30)], [output("C1", numberExpected(20))]),
  fixture("statistical", "averageif-basic", "AVERAGEIF", "=AVERAGEIF(A1:A4,\">0\")", [input("A1", 2), input("A2", 4), input("A3", -1), input("A4", 6)], [output("A5", numberExpected(4))]),
  fixture("statistical", "countif-basic", "COUNTIF", "=COUNTIF(A1:A4,\">0\")", [input("A1", 2), input("A2", 4), input("A3", -1), input("A4", 6)], [output("A5", numberExpected(3))]),
  fixture("information", "isblank-basic", "ISBLANK", "=ISBLANK(A1)", [input("A1", null)], [output("A2", booleanExpected(true))]),
  fixture("information", "isnumber-basic", "ISNUMBER", "=ISNUMBER(A1)", [input("A1", 42)], [output("A2", booleanExpected(true))]),
  fixture("information", "istext-basic", "ISTEXT", "=ISTEXT(A1)", [input("A1", "hello")], [output("A2", booleanExpected(true))]),
  fixture("dynamic-array", "sequence-spill", "SEQUENCE spill", "=SEQUENCE(3,1,1,1)", [], [output("A1", numberExpected(1)), output("A2", numberExpected(2)), output("A3", numberExpected(3))]),
  fixture("dynamic-array", "sequence-aggregate", "SUM over SEQUENCE", "=SUM(SEQUENCE(A1,1,1,1))", [input("A1", 3)], [output("B1", numberExpected(6))]),
  fixture("dynamic-array", "filter-basic", "FILTER", "=FILTER(A1:A4,A1:A4>2)", [input("A1", 1), input("A2", 3), input("A3", 2), input("A4", 4)], [output("B1", numberExpected(3)), output("B2", numberExpected(4))]),
  fixture("dynamic-array", "unique-basic", "UNIQUE", "=UNIQUE(A1:A4)", [input("A1", "A"), input("A2", "B"), input("A3", "A"), input("A4", "C")], [output("B1", stringExpected("A")), output("B2", stringExpected("B")), output("B3", stringExpected("C"))]),
  {
    ...fixture("names", "defined-name-scalar", "Defined name scalar reference", "=TaxRate*A1", [input("A1", 100)], [output("A2", numberExpected(8.5))], "Requires workbook-level defined names in the oracle workbook."),
    definedNames: [{ name: "TaxRate", value: 0.085 }]
  },
  fixture("tables", "table-total-row-sum", "Excel table total row SUM", "=SUM(Sales[Amount])", [input("A1", 10), input("A2", 12), input("A3", 15)], [output("A4", numberExpected(37))], "Table metadata must be part of the fixture workbook, not implied by plain cells."),
  fixture("structured-reference", "table-column-ref", "Structured table column reference", "=SUM(Sales[Amount])", [input("A1", 10), input("A2", 12), input("A3", 15)], [output("A4", numberExpected(37))], "Kept separate from table totals so the registry can track metadata-aware parsing independently."),
  fixture("volatile", "rand-basic", "RAND", "=RAND()", [], [output("A1", numberExpected(0.625))], "Volatile fixtures must record the exact captured output from the oracle workbook session."),
  fixture("lambda", "let-basic", "LET", "=LET(x,2,x+3)", [], [output("A1", numberExpected(5))]),
  fixture("lambda", "lambda-invoke", "LAMBDA invocation", "=LAMBDA(x,x+1)(4)", [], [output("A1", numberExpected(5))]),
  fixture("lambda", "map-basic", "MAP with LAMBDA", "=MAP(A1:A3,LAMBDA(x,x*2))", [input("A1", 1), input("A2", 2), input("A3", 3)], [output("B1", numberExpected(2)), output("B2", numberExpected(4)), output("B3", numberExpected(6))]),
  fixture("information", "value-error-display", "Visible #VALUE! error", "=1+\"x\"", [], [output("A1", errorExpected(ErrorCode.Value, "#VALUE!"))])
];
