import { ErrorCode } from "@bilig/protocol";
import type { ExcelExpectedValue, ExcelFixtureFamily, ExcelFixtureSuite } from "./index.js";

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

export const excelDateTimeFixtureSuite: ExcelFixtureSuite = {
  id: "datetime-serial-1900",
  description:
    "Focused Excel 1900-system serial date coverage for DATE, parts, TIME, weekday/time extraction, EDATE, and EOMONTH. TODAY/NOW are documented separately because they are volatile.",
  excelBuild: "Microsoft 365 / 2026-03-19",
  capturedAt: "2026-03-19T15:45:30.000Z",
  sheets: [{ name: "Sheet1" }],
  cases: [
    {
      id: createExcelFixtureId("date-time", "serial-addition"),
      family: "date-time",
      title: "Date serial addition",
      formula: "=A1+7",
      sheetName: "Sheet1",
      notes: "Real parity case using raw serial arithmetic rather than a seeded placeholder.",
      inputs: [{ address: "A1", input: 45292 }],
      outputs: [{ address: "A2", expected: numberExpected(45299) }],
    },
    {
      id: createExcelFixtureId("date-time", "date-constructor-leap-day"),
      family: "date-time",
      title: "DATE constructor leap day",
      formula: "=DATE(2024,2,29)",
      sheetName: "Sheet1",
      inputs: [],
      outputs: [{ address: "A1", expected: numberExpected(45351) }],
    },
    {
      id: createExcelFixtureId("date-time", "date-empty-year-coercion"),
      family: "date-time",
      title: "DATE coerces empty year to 1900",
      formula: "=DATE(A1,1,1)",
      sheetName: "Sheet1",
      inputs: [{ address: "A1", input: null }],
      outputs: [{ address: "A2", expected: numberExpected(1) }],
    },
    {
      id: createExcelFixtureId("date-time", "date-text-error"),
      family: "date-time",
      title: "DATE returns #VALUE! for text year",
      formula: "=DATE(A1,1,1)",
      sheetName: "Sheet1",
      inputs: [{ address: "A1", input: "bad" }],
      outputs: [{ address: "A2", expected: errorExpected(ErrorCode.Value, "#VALUE!") }],
    },
    {
      id: createExcelFixtureId("date-time", "year-month-day-parts"),
      family: "date-time",
      title: "YEAR, MONTH, and DAY from a serial",
      formula: "=YEAR(A1)+MONTH(A1)/100+DAY(A1)/10000",
      sheetName: "Sheet1",
      inputs: [{ address: "A1", input: 45351 }],
      outputs: [{ address: "A2", expected: numberExpected(2024.0229) }],
    },
    {
      id: createExcelFixtureId("date-time", "year-boolean-coercion"),
      family: "date-time",
      title: "YEAR coerces boolean serials",
      formula: "=YEAR(A1)",
      sheetName: "Sheet1",
      inputs: [{ address: "A1", input: true }],
      outputs: [{ address: "A2", expected: numberExpected(1900) }],
    },
    {
      id: createExcelFixtureId("date-time", "month-text-error"),
      family: "date-time",
      title: "MONTH returns #VALUE! for text input",
      formula: "=MONTH(A1)",
      sheetName: "Sheet1",
      inputs: [{ address: "A1", input: "bad" }],
      outputs: [{ address: "A2", expected: errorExpected(ErrorCode.Value, "#VALUE!") }],
    },
    {
      id: createExcelFixtureId("date-time", "day-leap-bug-serial"),
      family: "date-time",
      title: "DAY preserves Excel's 1900 leap-bug serial",
      formula: "=DAY(A1)",
      sheetName: "Sheet1",
      inputs: [{ address: "A1", input: 60 }],
      outputs: [{ address: "A2", expected: numberExpected(29) }],
    },
    {
      id: createExcelFixtureId("date-time", "time-basic"),
      family: "date-time",
      title: "TIME constructs a fractional day serial",
      formula: "=TIME(12,30,0)",
      sheetName: "Sheet1",
      inputs: [],
      outputs: [{ address: "A1", expected: numberExpected(0.5208333333333334) }],
    },
    {
      id: createExcelFixtureId("date-time", "hour-basic"),
      family: "date-time",
      title: "HOUR extracts the hour component",
      formula: "=HOUR(A1)",
      sheetName: "Sheet1",
      inputs: [{ address: "A1", input: 0.5208333333333334 }],
      outputs: [{ address: "A2", expected: numberExpected(12) }],
    },
    {
      id: createExcelFixtureId("date-time", "minute-basic"),
      family: "date-time",
      title: "MINUTE extracts the minute component",
      formula: "=MINUTE(A1)",
      sheetName: "Sheet1",
      inputs: [{ address: "A1", input: 0.5208333333333334 }],
      outputs: [{ address: "A2", expected: numberExpected(30) }],
    },
    {
      id: createExcelFixtureId("date-time", "second-basic"),
      family: "date-time",
      title: "SECOND extracts the second component",
      formula: "=SECOND(A1)",
      sheetName: "Sheet1",
      inputs: [{ address: "A1", input: 0.5208449074074074 }],
      outputs: [{ address: "A2", expected: numberExpected(1) }],
    },
    {
      id: createExcelFixtureId("date-time", "weekday-basic"),
      family: "date-time",
      title: "WEEKDAY returns default Sunday-first numbering",
      formula: "=WEEKDAY(DATE(2026,3,15))",
      sheetName: "Sheet1",
      inputs: [],
      outputs: [{ address: "A1", expected: numberExpected(1) }],
    },
    {
      id: createExcelFixtureId("date-time", "edate-month-shift"),
      family: "date-time",
      title: "EDATE month shift with leap-year clamp",
      formula: "=EDATE(A1,1)",
      sheetName: "Sheet1",
      inputs: [{ address: "A1", input: 45322 }],
      outputs: [{ address: "A2", expected: numberExpected(45351) }],
    },
    {
      id: createExcelFixtureId("date-time", "edate-boolean-coercion"),
      family: "date-time",
      title: "EDATE coerces boolean start dates",
      formula: "=EDATE(A1,1)",
      sheetName: "Sheet1",
      inputs: [{ address: "A1", input: true }],
      outputs: [{ address: "A2", expected: numberExpected(32) }],
    },
    {
      id: createExcelFixtureId("date-time", "edate-text-error"),
      family: "date-time",
      title: "EDATE returns #VALUE! for text start dates",
      formula: "=EDATE(A1,1)",
      sheetName: "Sheet1",
      inputs: [{ address: "A1", input: "bad" }],
      outputs: [{ address: "A2", expected: errorExpected(ErrorCode.Value, "#VALUE!") }],
    },
    {
      id: createExcelFixtureId("date-time", "eomonth-forward"),
      family: "date-time",
      title: "EOMONTH current and next month",
      formula: "=EOMONTH(A1,B1)",
      sheetName: "Sheet1",
      inputs: [
        { address: "A1", input: 45337 },
        { address: "B1", input: 1 },
      ],
      outputs: [{ address: "A2", expected: numberExpected(45382) }],
    },
    {
      id: createExcelFixtureId("date-time", "eomonth-boolean-coercion"),
      family: "date-time",
      title: "EOMONTH coerces boolean start dates",
      formula: "=EOMONTH(A1,1)",
      sheetName: "Sheet1",
      inputs: [{ address: "A1", input: true }],
      outputs: [{ address: "A2", expected: numberExpected(60) }],
    },
    {
      id: createExcelFixtureId("date-time", "eomonth-text-error"),
      family: "date-time",
      title: "EOMONTH returns #VALUE! for text offsets",
      formula: "=EOMONTH(A1,A2)",
      sheetName: "Sheet1",
      inputs: [
        { address: "A1", input: 1 },
        { address: "A2", input: "bad" },
      ],
      outputs: [{ address: "A3", expected: errorExpected(ErrorCode.Value, "#VALUE!") }],
    },
    {
      id: createExcelFixtureId("date-time", "datedif-ym"),
      family: "date-time",
      title: "DATEDIF returns remaining months after full years",
      formula: '=DATEDIF(DATE(2020,1,15),DATE(2021,3,20),"YM")',
      sheetName: "Sheet1",
      inputs: [],
      outputs: [{ address: "A1", expected: numberExpected(2) }],
    },
    {
      id: createExcelFixtureId("volatile", "today-captured-utc"),
      family: "volatile",
      title: "TODAY captured at suite timestamp",
      formula: "=TODAY()",
      sheetName: "Sheet1",
      notes:
        "Snapshot recorded against the suite capturedAt UTC timestamp for later oracle comparison.",
      inputs: [],
      outputs: [{ address: "A1", expected: numberExpected(46100) }],
    },
    {
      id: createExcelFixtureId("volatile", "now-captured-utc"),
      family: "volatile",
      title: "NOW captured at suite timestamp",
      formula: "=NOW()",
      sheetName: "Sheet1",
      notes:
        "Snapshot recorded against the suite capturedAt UTC timestamp for later oracle comparison.",
      inputs: [],
      outputs: [{ address: "A1", expected: numberExpected(46100.65659722222) }],
    },
    {
      id: createExcelFixtureId("volatile", "rand-captured"),
      family: "volatile",
      title: "RAND captured sample",
      formula: "=RAND()",
      sheetName: "Sheet1",
      notes:
        "Uses a fixed sampled output so fixture harnesses can assert deterministic JS behavior under a stubbed random source.",
      inputs: [],
      outputs: [{ address: "A1", expected: numberExpected(0.625) }],
    },
  ],
};
