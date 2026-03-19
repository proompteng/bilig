import { describe, expect, it } from "vitest";
import { ErrorCode, ValueTag } from "@bilig/protocol";
import {
  addMonthsToExcelDate,
  createNowBuiltin,
  createRandBuiltin,
  createTodayBuiltin,
  datetimeBuiltins,
  endOfMonthExcelDate,
  excelDatePartsToSerial,
  excelSerialToDateParts,
  utcDateToExcelSerial
} from "../builtins/datetime.js";
import { excelDateTimeFixtureSuite } from "../../../excel-fixtures/src/datetime-fixtures.js";

describe("datetime builtins", () => {
  it("converts between Excel serials and date parts in the 1900 system", () => {
    expect(excelDatePartsToSerial(1900, 1, 1)).toBe(1);
    expect(excelDatePartsToSerial(1900, 2, 29)).toBe(60);
    expect(excelDatePartsToSerial(1900, 3, 1)).toBe(61);
    expect(excelDatePartsToSerial(2024, 2, 29)).toBe(45351);

    expect(excelSerialToDateParts(60)).toEqual({ year: 1900, month: 2, day: 29 });
    expect(excelSerialToDateParts(61)).toEqual({ year: 1900, month: 3, day: 1 });
    expect(excelSerialToDateParts(45351)).toEqual({ year: 2024, month: 2, day: 29 });
  });

  it("supports DATE with Excel-style year and month/day normalization", () => {
    expect(datetimeBuiltins.DATE(
      { tag: ValueTag.Number, value: 2024 },
      { tag: ValueTag.Number, value: 2 },
      { tag: ValueTag.Number, value: 29 }
    )).toEqual({ tag: ValueTag.Number, value: 45351 });

    expect(datetimeBuiltins.DATE(
      { tag: ValueTag.Number, value: 24 },
      { tag: ValueTag.Number, value: 14 },
      { tag: ValueTag.Number, value: 1 }
    )).toEqual({ tag: ValueTag.Number, value: 9164 });

    expect(datetimeBuiltins.DATE(
      { tag: ValueTag.String, value: "2024", stringId: 1 },
      { tag: ValueTag.Number, value: 2 },
      { tag: ValueTag.Number, value: 29 }
    )).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });

    expect(datetimeBuiltins.DATE(
      { tag: ValueTag.Error, code: ErrorCode.Ref },
      { tag: ValueTag.Number, value: 2 },
      { tag: ValueTag.Number, value: 29 }
    )).toEqual({ tag: ValueTag.Error, code: ErrorCode.Ref });
  });

  it("extracts YEAR, MONTH, and DAY from serial inputs including the leap-year bug date", () => {
    expect(datetimeBuiltins.YEAR({ tag: ValueTag.Number, value: 45351 })).toEqual({
      tag: ValueTag.Number,
      value: 2024
    });
    expect(datetimeBuiltins.MONTH({ tag: ValueTag.Number, value: 45351.75 })).toEqual({
      tag: ValueTag.Number,
      value: 2
    });
    expect(datetimeBuiltins.DAY({ tag: ValueTag.Number, value: 60 })).toEqual({
      tag: ValueTag.Number,
      value: 29
    });

    expect(datetimeBuiltins.YEAR({ tag: ValueTag.String, value: "45351", stringId: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value
    });
  });

  it("supports TIME plus HOUR, MINUTE, SECOND, and WEEKDAY extraction", () => {
    const sundaySerial = excelDatePartsToSerial(2026, 3, 15)!;

    expect(datetimeBuiltins.TIME(
      { tag: ValueTag.Number, value: 12 },
      { tag: ValueTag.Number, value: 30 },
      { tag: ValueTag.Number, value: 0 }
    )).toEqual({ tag: ValueTag.Number, value: 0.5208333333333334 });

    expect(datetimeBuiltins.HOUR({ tag: ValueTag.Number, value: 0.5208333333333334 })).toEqual({
      tag: ValueTag.Number,
      value: 12
    });
    expect(datetimeBuiltins.MINUTE({ tag: ValueTag.Number, value: 0.5208333333333334 })).toEqual({
      tag: ValueTag.Number,
      value: 30
    });
    expect(datetimeBuiltins.SECOND({ tag: ValueTag.Number, value: 0.5208449074074074 })).toEqual({
      tag: ValueTag.Number,
      value: 1
    });
    expect(datetimeBuiltins.WEEKDAY({ tag: ValueTag.Number, value: sundaySerial })).toEqual({
      tag: ValueTag.Number,
      value: 1
    });
    expect(datetimeBuiltins.WEEKDAY(
      { tag: ValueTag.Number, value: sundaySerial },
      { tag: ValueTag.Number, value: 2 }
    )).toEqual({ tag: ValueTag.Number, value: 7 });
    expect(datetimeBuiltins.WEEKDAY(
      { tag: ValueTag.Number, value: sundaySerial },
      { tag: ValueTag.Number, value: 3 }
    )).toEqual({ tag: ValueTag.Number, value: 6 });
  });

  it("returns #VALUE for unsupported time-part coercions and weekday return types", () => {
    expect(datetimeBuiltins.TIME(
      { tag: ValueTag.Number, value: -1 },
      { tag: ValueTag.Number, value: 30 },
      { tag: ValueTag.Number, value: 0 }
    )).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });
    expect(datetimeBuiltins.HOUR({ tag: ValueTag.String, value: "12:30", stringId: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value
    });
    expect(datetimeBuiltins.WEEKDAY(
      { tag: ValueTag.Number, value: excelDatePartsToSerial(2026, 3, 15)! },
      { tag: ValueTag.Number, value: 99 }
    )).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });
  });

  it("creates deterministic TODAY and NOW builtins from injected UTC dates", () => {
    const fixedNow = new Date("2026-03-19T15:45:30.000Z");
    const TODAY = createTodayBuiltin(() => fixedNow);
    const NOW = createNowBuiltin(() => fixedNow);

    expect(TODAY()).toEqual({ tag: ValueTag.Number, value: 46100 });
    expect(NOW()).toEqual({ tag: ValueTag.Number, value: 46100.65659722222 });
    expect(utcDateToExcelSerial(fixedNow)).toBe(46100.65659722222);

    expect(TODAY({ tag: ValueTag.Number, value: 1 })).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });
    expect(NOW({ tag: ValueTag.Error, code: ErrorCode.NA })).toEqual({ tag: ValueTag.Error, code: ErrorCode.NA });
  });

  it("supports RAND with Excel-style numeric bounds and injectable randomness", () => {
    const RAND = createRandBuiltin(() => 0.625);
    const highRAND = createRandBuiltin(() => 2);
    const lowRAND = createRandBuiltin(() => -0.5);
    const invalidRAND = createRandBuiltin(() => Number.NaN);

    expect(RAND()).toEqual({ tag: ValueTag.Number, value: 0.625 });
    expect(highRAND()).toEqual({ tag: ValueTag.Number, value: 1 - Number.EPSILON });
    expect(lowRAND()).toEqual({ tag: ValueTag.Number, value: 0 });
    expect(invalidRAND()).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });
    expect(datetimeBuiltins.RAND({ tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value
    });
  });

  it("supports EDATE month shifting with end-of-month clamping", () => {
    expect(addMonthsToExcelDate(45322, 1)).toBe(45351);
    expect(addMonthsToExcelDate(45351, -1)).toBe(45320);

    expect(datetimeBuiltins.EDATE(
      { tag: ValueTag.Number, value: 45322 },
      { tag: ValueTag.Number, value: 1.9 }
    )).toEqual({ tag: ValueTag.Number, value: 45351 });

    expect(datetimeBuiltins.EDATE(
      { tag: ValueTag.String, value: "bad", stringId: 1 },
      { tag: ValueTag.Number, value: 1 }
    )).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });
  });

  it("supports EOMONTH end-of-month lookups", () => {
    expect(endOfMonthExcelDate(45337, 0)).toBe(45351);
    expect(endOfMonthExcelDate(45337, 1)).toBe(45382);

    expect(datetimeBuiltins.EOMONTH(
      { tag: ValueTag.Number, value: 45337 },
      { tag: ValueTag.Boolean, value: true }
    )).toEqual({ tag: ValueTag.Number, value: 45382 });

    expect(datetimeBuiltins.EOMONTH(
      { tag: ValueTag.Number, value: 45337 },
      { tag: ValueTag.Error, code: ErrorCode.Ref }
    )).toEqual({ tag: ValueTag.Error, code: ErrorCode.Ref });
  });

  it("ships a focused datetime fixture suite for later aggregation", () => {
    expect(excelDateTimeFixtureSuite.id).toBe("datetime-serial-1900");
    expect(excelDateTimeFixtureSuite.sheets).toEqual([{ name: "Sheet1" }]);
    expect(excelDateTimeFixtureSuite.cases).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          id: "date-time:serial-addition",
          outputs: [{ address: "A2", expected: { kind: "number", value: 45299 } }]
        }),
        expect.objectContaining({
          id: "date-time:date-empty-year-coercion",
          outputs: [{ address: "A2", expected: { kind: "number", value: 1 } }]
        }),
        expect.objectContaining({
          id: "date-time:date-text-error",
          outputs: [{ address: "A2", expected: { kind: "error", code: ErrorCode.Value, display: "#VALUE!" } }]
        }),
        expect.objectContaining({
          id: "date-time:date-constructor-leap-day",
          formula: "=DATE(2024,2,29)"
        }),
        expect.objectContaining({
          id: "date-time:year-boolean-coercion",
          outputs: [{ address: "A2", expected: { kind: "number", value: 1900 } }]
        }),
        expect.objectContaining({
          id: "date-time:month-text-error",
          outputs: [{ address: "A2", expected: { kind: "error", code: ErrorCode.Value, display: "#VALUE!" } }]
        }),
        expect.objectContaining({
          id: "date-time:day-leap-bug-serial",
          outputs: [{ address: "A2", expected: { kind: "number", value: 29 } }]
        }),
        expect.objectContaining({
          id: "date-time:time-basic",
          outputs: [{ address: "A1", expected: { kind: "number", value: 0.5208333333333334 } }]
        }),
        expect.objectContaining({
          id: "date-time:hour-basic",
          outputs: [{ address: "A2", expected: { kind: "number", value: 12 } }]
        }),
        expect.objectContaining({
          id: "date-time:minute-basic",
          outputs: [{ address: "A2", expected: { kind: "number", value: 30 } }]
        }),
        expect.objectContaining({
          id: "date-time:second-basic",
          outputs: [{ address: "A2", expected: { kind: "number", value: 1 } }]
        }),
        expect.objectContaining({
          id: "date-time:weekday-basic",
          outputs: [{ address: "A1", expected: { kind: "number", value: 1 } }]
        }),
        expect.objectContaining({
          id: "date-time:edate-month-shift",
          outputs: [{ address: "A2", expected: { kind: "number", value: 45351 } }]
        }),
        expect.objectContaining({
          id: "date-time:edate-boolean-coercion",
          outputs: [{ address: "A2", expected: { kind: "number", value: 32 } }]
        }),
        expect.objectContaining({
          id: "date-time:edate-text-error",
          outputs: [{ address: "A2", expected: { kind: "error", code: ErrorCode.Value, display: "#VALUE!" } }]
        }),
        expect.objectContaining({
          id: "date-time:eomonth-boolean-coercion",
          outputs: [{ address: "A2", expected: { kind: "number", value: 60 } }]
        }),
        expect.objectContaining({
          id: "date-time:eomonth-text-error",
          outputs: [{ address: "A3", expected: { kind: "error", code: ErrorCode.Value, display: "#VALUE!" } }]
        }),
        expect.objectContaining({
          id: "volatile:today-captured-utc",
          outputs: [{ address: "A1", expected: { kind: "number", value: 46100 } }]
        }),
        expect.objectContaining({
          id: "volatile:rand-captured",
          outputs: [{ address: "A1", expected: { kind: "number", value: 0.625 } }]
        })
      ])
    );
  });
});
