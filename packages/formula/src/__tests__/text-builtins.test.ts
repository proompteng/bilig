import { describe, expect, it } from "vitest";
import { ErrorCode, ValueTag, type CellValue } from "@bilig/protocol";
import { TEXT_FIXTURES } from "../../../excel-fixtures/src/text-fixtures.js";
import { getTextBuiltin } from "../builtins/text.js";

describe("text builtins", () => {
  it("matches the shared text fixture corpus", () => {
    for (const group of TEXT_FIXTURES) {
      const builtin = getTextBuiltin(group.builtin);
      expect(builtin).toBeTypeOf("function");

      for (const testCase of group.cases) {
        expect(builtin?.(...testCase.args)).toEqual(testCase.expected);
      }
    }
  });

  it("propagates existing error values", () => {
    const errorValue: CellValue = { tag: ValueTag.Error, code: ErrorCode.Ref };

    expect(getTextBuiltin("LEN")?.(errorValue)).toEqual(errorValue);
    expect(getTextBuiltin("CONCAT")?.(text("a"), errorValue, text("b"))).toEqual(errorValue);
    expect(getTextBuiltin("FIND")?.(text("a"), errorValue)).toEqual(errorValue);
    expect(getTextBuiltin("SEARCH")?.(text("a"), text("abc"), errorValue)).toEqual(errorValue);
  });

  it("returns #VALUE for invalid coercions and invalid positions", () => {
    expect(getTextBuiltin("LEFT")?.(text("abc"), text("bad"))).toEqual(valueError());
    expect(getTextBuiltin("RIGHT")?.(text("abc"), text("bad"))).toEqual(valueError());
    expect(getTextBuiltin("MID")?.(text("abc"), number(0), number(1))).toEqual(valueError());
    expect(getTextBuiltin("FIND")?.(text("z"), text("abc"))).toEqual(valueError());
    expect(getTextBuiltin("SEARCH")?.(text("a"), text("abc"), number(0))).toEqual(valueError());
    expect(getTextBuiltin("VALUE")?.(text("not-a-number"))).toEqual(valueError());
  });

  it("supports Excel-like SEARCH wildcards and FIND case sensitivity", () => {
    expect(getTextBuiltin("SEARCH")?.(text("b?d"), text("ABCD"))).toEqual(number(2));
    expect(getTextBuiltin("SEARCH")?.(text("a*d"), text("xaZZd"))).toEqual(number(2));
    expect(getTextBuiltin("SEARCH")?.(text("~*"), text("a*b"))).toEqual(number(2));
    expect(getTextBuiltin("FIND")?.(text("A"), text("abcA"))).toEqual(number(4));
    expect(getTextBuiltin("FIND")?.(text("A"), text("abca"))).toEqual(valueError());
  });

  it("keeps TRIM scoped to ASCII spaces", () => {
    expect(getTextBuiltin("TRIM")?.(text("  alpha   beta  "))).toEqual(text("alpha beta"));
    expect(getTextBuiltin("TRIM")?.(text("\t alpha  beta \t"))).toEqual(text("\t alpha beta \t"));
  });

  it("supports EXACT as case-sensitive text equality", () => {
    expect(getTextBuiltin("EXACT")?.(text("Alpha"), text("Alpha"))).toEqual({
      tag: ValueTag.Boolean,
      value: true,
    });
    expect(getTextBuiltin("EXACT")?.(text("Alpha"), text("alpha"))).toEqual({
      tag: ValueTag.Boolean,
      value: false,
    });
    expect(getTextBuiltin("EXACT")?.(number(42), text("42"))).toEqual({
      tag: ValueTag.Boolean,
      value: true,
    });
  });

  it("supports VALUE numeric coercion for trimmed text, booleans, and empties", () => {
    expect(getTextBuiltin("VALUE")?.(text(" 42 "))).toEqual(number(42));
    expect(getTextBuiltin("VALUE")?.({ tag: ValueTag.Boolean, value: true })).toEqual(number(1));
    expect(getTextBuiltin("VALUE")?.({ tag: ValueTag.Empty })).toEqual(number(0));
  });

  it("supports TEXTBEFORE search modes, negative instances, and fallbacks", () => {
    const TEXTBEFORE = getTextBuiltin("TEXTBEFORE")!;

    expect(TEXTBEFORE(text("alpha-beta-gamma"), text("-"))).toEqual(text("alpha"));
    expect(TEXTBEFORE(text("alpha-beta-gamma"), text("-"), number(-1))).toEqual(text("alpha-beta"));
    expect(TEXTBEFORE(text("Alpha-beta"), text("a"), number(2), number(1))).toEqual(text("Alph"));
    expect(
      TEXTBEFORE(text("alpha"), text("-"), number(1), number(0), number(0), text("fallback")),
    ).toEqual(text("fallback"));
    expect(
      TEXTBEFORE(text("alpha-beta"), text("-"), number(-5), number(0), number(1), text("edge")),
    ).toEqual(text("edge"));
  });

  it("supports REPLACE, SUBSTITUTE, and REPT", () => {
    expect(getTextBuiltin("REPLACE")?.(text("alphabet"), number(3), number(2), text("Z"))).toEqual(
      text("alZabet"),
    );
    expect(getTextBuiltin("SUBSTITUTE")?.(text("banana"), text("an"), text("oo"))).toEqual(
      text("booooa"),
    );
    expect(
      getTextBuiltin("SUBSTITUTE")?.(text("banana"), text("an"), text("oo"), number(2)),
    ).toEqual(text("banooa"));
    expect(getTextBuiltin("REPT")?.(text("xo"), number(3))).toEqual(text("xoxoxo"));
  });

  it("returns REPLACE, SUBSTITUTE, and REPT validation errors when arguments are invalid", () => {
    expect(getTextBuiltin("REPLACE")?.(text("abc"), number(0), number(1), text("z"))).toEqual(
      valueError(),
    );
    expect(getTextBuiltin("SUBSTITUTE")?.(text("abc"), text(""), text("z"))).toEqual(valueError());
    expect(getTextBuiltin("SUBSTITUTE")?.(text("abc"), text("a"), text("z"), number(0))).toEqual(
      valueError(),
    );
    expect(getTextBuiltin("REPT")?.(text("abc"), number(-1))).toEqual(valueError());
    expect(getTextBuiltin("REPT")?.(err(ErrorCode.Ref), number(2))).toEqual(err(ErrorCode.Ref));
  });

  it("returns TEXTBEFORE validation and lookup errors when arguments are invalid", () => {
    const TEXTBEFORE = getTextBuiltin("TEXTBEFORE")!;

    expect(TEXTBEFORE()).toEqual(valueError());
    expect(TEXTBEFORE(text("alpha"), text(""))).toEqual(valueError());
    expect(TEXTBEFORE(text("alpha"), text("-"), number(0))).toEqual(valueError());
    expect(TEXTBEFORE(text("alpha"), text("-"), number(1), number(2))).toEqual(valueError());
    expect(TEXTBEFORE(text("alpha"), text("-"), number(1), number(0), number(0))).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.NA,
    });
    expect(TEXTBEFORE(err(ErrorCode.Ref), text("-"))).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    });
    expect(getTextBuiltin("VALUE")?.()).toEqual(valueError());
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
  return { tag: ValueTag.Error, code: ErrorCode.Value };
}
