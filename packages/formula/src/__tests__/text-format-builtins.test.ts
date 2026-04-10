import { describe, expect, it } from "vitest";
import { ErrorCode, ValueTag, type CellValue } from "@bilig/protocol";
import { getTextBuiltin } from "../builtins/text.js";

describe("text format builtins", () => {
  it("supports TEXT for numeric, date-time, and text-section formatting", () => {
    const TEXT = getTextBuiltin("TEXT")!;

    expect(TEXT(number(1234.567), text("#,##0.00"))).toEqual(text("1,234.57"));
    expect(TEXT(number(0.1234), text("0.0%"))).toEqual(text("12.3%"));
    expect(TEXT(number(45356), text("yyyy-mm-dd"))).toEqual(text("2024-03-05"));
    expect(TEXT(number(0.5), text("h:mm AM/PM"))).toEqual(text("12:00 PM"));
    expect(TEXT(text("alpha"), text("prefix @"))).toEqual(text("prefix alpha"));
  });

  it("supports VALUE, NUMBERVALUE, and VALUETOTEXT conversions", () => {
    const VALUE = getTextBuiltin("VALUE")!;
    const NUMBERVALUE = getTextBuiltin("NUMBERVALUE")!;
    const VALUETOTEXT = getTextBuiltin("VALUETOTEXT")!;

    expect(VALUE(text(" 42 "))).toEqual(number(42));
    expect(NUMBERVALUE(text("2.500,27"), text(","), text("."))).toEqual(number(2500.27));
    expect(NUMBERVALUE(text("9%%"))).toEqual(number(0.0009));
    expect(VALUETOTEXT(number(42))).toEqual(text("42"));
    expect(VALUETOTEXT(text("alpha"), number(1))).toEqual(text('"alpha"'));
    expect(VALUETOTEXT(err(ErrorCode.Ref))).toEqual(text("#REF!"));
  });

  it("keeps validation and error propagation for text formatting builtins", () => {
    const TEXT = getTextBuiltin("TEXT")!;
    const VALUE = getTextBuiltin("VALUE")!;
    const NUMBERVALUE = getTextBuiltin("NUMBERVALUE")!;
    const VALUETOTEXT = getTextBuiltin("VALUETOTEXT")!;

    expect(TEXT()).toEqual(valueError());
    expect(TEXT(text("alpha"), text("0.00"))).toEqual(valueError());
    expect(TEXT(err(ErrorCode.Ref), text("0.00"))).toEqual(err(ErrorCode.Ref));
    expect(VALUE()).toEqual(valueError());
    expect(VALUE(err(ErrorCode.Name))).toEqual(err(ErrorCode.Name));
    expect(NUMBERVALUE(text("1.2.3"), text("."), text(","))).toEqual(valueError());
    expect(VALUETOTEXT(text("alpha"), number(2))).toEqual(valueError());
  });
});

function number(value: number): CellValue {
  return { tag: ValueTag.Number, value };
}

function text(value: string): CellValue {
  return { tag: ValueTag.String, value, stringId: 0 };
}

function err(code: ErrorCode): CellValue {
  return { tag: ValueTag.Error, code };
}

function valueError(): CellValue {
  return err(ErrorCode.Value);
}
