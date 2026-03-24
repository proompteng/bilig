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

  it("supports ENCODEURL for URL-safe formatting", () => {
    expect(getTextBuiltin("ENCODEURL")?.(text("https://example.com/a b"))).toEqual(
      text("https://example.com/a%20b"),
    );
    expect(getTextBuiltin("ENCODEURL")?.(text("a&b=c?x=y"))).toEqual(text("a&b=c?x=y"));
    expect(getTextBuiltin("ENCODEURL")?.(err(ErrorCode.Ref))).toEqual(err(ErrorCode.Ref));
  });

  it("supports FINDB, LEFTB, MIDB, and RIGHTB with byte boundaries", () => {
    expect(getTextBuiltin("FINDB")?.(text("b"), text("abc"))).toEqual(number(2));
    expect(getTextBuiltin("FINDB")?.(text("d"), text("abcd"), number(3))).toEqual(number(4));
    expect(getTextBuiltin("LEFTB")?.(text("abcdef"), number(2))).toEqual(text("ab"));
    expect(getTextBuiltin("MIDB")?.(text("abcdef"), number(3), number(2))).toEqual(text("cd"));
    expect(getTextBuiltin("RIGHTB")?.(text("abcdef"), number(3))).toEqual(text("def"));
  });

  it("validates FINDB/LEFTB/MIDB/RIGHTB argument bounds as VALUE errors", () => {
    expect(getTextBuiltin("FINDB")?.(text("x"), text("abc"), number(0))).toEqual(valueError());
    expect(getTextBuiltin("FINDB")?.(text("x"), text("abc"), number(5))).toEqual(valueError());
    expect(getTextBuiltin("LEFTB")?.(text("abc"), number(-1))).toEqual(valueError());
    expect(getTextBuiltin("MIDB")?.(text("abc"), number(0), number(2))).toEqual(valueError());
    expect(getTextBuiltin("RIGHTB")?.(text("abc"), number(-1))).toEqual(valueError());
  });

  it("supports CLEAN, CONCATENATE, and PROPER text cleanup functions", () => {
    expect(getTextBuiltin("CLEAN")?.(text("a\u0001b\u007fc"))).toEqual(text("abc"));
    expect(getTextBuiltin("CONCATENATE")?.(text("a"), number(1), text("b"))).toEqual(text("a1b"));
    expect(getTextBuiltin("PROPER")?.(text("hello world"))).toEqual(text("Hello World"));
    expect(getTextBuiltin("PROPER")?.(text("hELLO, wORLD"))).toEqual(text("Hello, World"));
    expect(getTextBuiltin("CONCATENATE")?.()).toEqual(valueError());
  });

  it("supports TEXTAFTER search modes, negative instances, and fallback values", () => {
    const TEXTAFTER = getTextBuiltin("TEXTAFTER")!;

    expect(TEXTAFTER(text("alpha-beta-gamma"), text("-"))).toEqual(text("beta-gamma"));
    expect(TEXTAFTER(text("alpha-beta-gamma"), text("-"), number(2))).toEqual(text("gamma"));
    expect(TEXTAFTER(text("Alpha-beta"), text("A"), number(1), number(1))).toEqual(
      text("lpha-beta"),
    );
    expect(
      TEXTAFTER(text("alpha"), text("-"), number(1), number(0), number(0), text("fallback")),
    ).toEqual(text("fallback"));
    expect(TEXTAFTER(text("alpha-beta-gamma"), text("-"), number(-1))).toEqual(text("gamma"));
    expect(TEXTAFTER(err(ErrorCode.Ref), text("-"))).toEqual(err(ErrorCode.Ref));
  });

  it("supports TEXTJOIN delimiter, ignore_empty behavior, and error propagation", () => {
    const TEXTJOIN = getTextBuiltin("TEXTJOIN")!;

    expect(
      TEXTJOIN(text(","), { tag: ValueTag.Boolean, value: false }, text("a"), text("b")),
    ).toEqual(text("a,b"));
    expect(
      TEXTJOIN(
        text("|"),
        { tag: ValueTag.Boolean, value: false },
        text("a"),
        { tag: ValueTag.Empty },
        text("b"),
      ),
    ).toEqual(text("a||b"));
    expect(
      TEXTJOIN(
        text("|"),
        { tag: ValueTag.Boolean, value: false },
        text("a"),
        err(ErrorCode.Ref),
        text("b"),
      ),
    ).toEqual(err(ErrorCode.Ref));
    expect(
      TEXTJOIN(text("-"), { tag: ValueTag.Boolean, value: true }, text("a"), text(""), text("b")),
    ).toEqual(text("a-b"));
    expect(TEXTJOIN()).toEqual(valueError());
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

  it("supports CHAR, CODE, UNICODE, and UNICHAR scalar conversions", () => {
    expect(getTextBuiltin("CHAR")?.(number(65))).toEqual(text("A"));
    expect(getTextBuiltin("CHAR")?.(number(10.8))).toEqual(text("\n"));
    expect(getTextBuiltin("CHAR")?.(number(0))).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });

    expect(getTextBuiltin("CODE")?.(text("A"))).toEqual(number(65));
    expect(getTextBuiltin("CODE")?.(text("😀"))).toEqual(number(0x1f600));
    expect(getTextBuiltin("CODE")?.(text(""))).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });

    expect(getTextBuiltin("UNICODE")?.(text("A"))).toEqual(number(65));
    expect(getTextBuiltin("UNICODE")?.(text("😀"))).toEqual(number(0x1f600));

    expect(getTextBuiltin("UNICHAR")?.(number(66))).toEqual(text("B"));
    expect(getTextBuiltin("UNICHAR")?.(text("66"))).toEqual(text("B"));
    expect(getTextBuiltin("UNICHAR")?.(number(-1))).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });
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
