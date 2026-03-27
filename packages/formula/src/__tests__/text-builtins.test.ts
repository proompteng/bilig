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
    expect(getTextBuiltin("FIND")?.()).toEqual(valueError());
    expect(getTextBuiltin("SEARCH")?.(text("a"))).toEqual(valueError());
    expect(getTextBuiltin("ENCODEURL")?.()).toEqual(valueError());
    expect(getTextBuiltin("FINDB")?.()).toEqual(valueError());
  });

  it("supports Excel-like SEARCH wildcards and FIND case sensitivity", () => {
    expect(getTextBuiltin("SEARCH")?.(text("b?d"), text("ABCD"))).toEqual(number(2));
    expect(getTextBuiltin("SEARCH")?.(text("a*d"), text("xaZZd"))).toEqual(number(2));
    expect(getTextBuiltin("SEARCH")?.(text("~*"), text("a*b"))).toEqual(number(2));
    expect(getTextBuiltin("FIND")?.(text("A"), text("abcA"))).toEqual(number(4));
    expect(getTextBuiltin("FIND")?.(text("A"), text("abca"))).toEqual(valueError());
    expect(getTextBuiltin("FIND")?.(text("A"), text("abc"), err(ErrorCode.Ref))).toEqual(
      err(ErrorCode.Ref),
    );
    expect(getTextBuiltin("SEARCH")?.(text("A"), text("abc"), err(ErrorCode.Name))).toEqual(
      err(ErrorCode.Name),
    );
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
    expect(getTextBuiltin("LENB")?.(text("abc"))).toEqual(number(3));
    expect(getTextBuiltin("LENB")?.(text("é"))).toEqual(number(2));
    expect(getTextBuiltin("LEFTB")?.(text("abcdef"), number(2))).toEqual(text("ab"));
    expect(getTextBuiltin("MIDB")?.(text("abcdef"), number(3), number(2))).toEqual(text("cd"));
    expect(getTextBuiltin("RIGHTB")?.(text("abcdef"), number(3))).toEqual(text("def"));
    expect(getTextBuiltin("SEARCHB")?.(text("ph"), text("alphabet"))).toEqual(number(3));
    expect(getTextBuiltin("SEARCHB")?.(text("b?d"), text("ABCD"))).toEqual(number(2));
    expect(getTextBuiltin("REPLACEB")?.(text("alphabet"), number(3), number(2), text("Z"))).toEqual(
      text("alZabet"),
    );
  });

  it("validates FINDB/LEFTB/MIDB/RIGHTB argument bounds as VALUE errors", () => {
    expect(getTextBuiltin("FINDB")?.(text("x"), text("abc"), number(0))).toEqual(valueError());
    expect(getTextBuiltin("FINDB")?.(text("x"), text("abc"), number(5))).toEqual(valueError());
    expect(getTextBuiltin("LEFTB")?.(text("abc"), number(-1))).toEqual(valueError());
    expect(getTextBuiltin("MIDB")?.(text("abc"), number(0), number(2))).toEqual(valueError());
    expect(getTextBuiltin("RIGHTB")?.(text("abc"), number(-1))).toEqual(valueError());
    expect(getTextBuiltin("LEFTB")?.()).toEqual(valueError());
    expect(getTextBuiltin("MIDB")?.(text("abc"), number(1))).toEqual(valueError());
    expect(getTextBuiltin("RIGHTB")?.()).toEqual(valueError());
    expect(getTextBuiltin("SEARCHB")?.(text("a"))).toEqual(valueError());
    expect(getTextBuiltin("SEARCHB")?.(text("a"), text("abc"), number(0))).toEqual(valueError());
    expect(getTextBuiltin("REPLACEB")?.(text("abc"), number(0), number(1), text("z"))).toEqual(
      valueError(),
    );
    expect(getTextBuiltin("LEFTB")?.(err(ErrorCode.Ref), number(1))).toEqual(err(ErrorCode.Ref));
    expect(getTextBuiltin("MIDB")?.(text("abc"), err(ErrorCode.Name), number(1))).toEqual(
      err(ErrorCode.Name),
    );
    expect(getTextBuiltin("RIGHTB")?.(text("abc"), err(ErrorCode.Div0))).toEqual(
      err(ErrorCode.Div0),
    );
    expect(getTextBuiltin("LENB")?.(err(ErrorCode.Ref))).toEqual(err(ErrorCode.Ref));
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
    expect(TEXTAFTER(text("alpha"), text("-"), number(-1), number(0), number(1))).toEqual(
      err(ErrorCode.NA),
    );
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
    expect(
      TEXTJOIN(text("-"), { tag: ValueTag.Boolean, value: true }, text("a"), text(""), text("")),
    ).toEqual(text("a"));
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
    expect(TEXTBEFORE(text("alpha"), text("-"), number(-1), number(0), number(1))).toEqual(
      err(ErrorCode.NA),
    );
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

  it("supports NUMBERVALUE locale-independent numeric parsing", () => {
    const NUMBERVALUE = getTextBuiltin("NUMBERVALUE")!;

    expect(NUMBERVALUE(text("2.500,27"), text(","), text("."))).toEqual(number(2500.27));
    expect(NUMBERVALUE(text(" 3 000 "))).toEqual(number(3000));
    expect(NUMBERVALUE(text("9%%"))).toEqual(number(0.0009));
    expect(NUMBERVALUE(text(""))).toEqual(number(0));
    expect(NUMBERVALUE(text("1,2,3"), text(","), text("."))).toEqual(valueError());
    expect(NUMBERVALUE(text("1.2.3"), text("."), text(","))).toEqual(valueError());
  });

  it("supports REGEXTEST, REGEXREPLACE, and REGEXEXTRACT modes", () => {
    const REGEXTEST = getTextBuiltin("REGEXTEST")!;
    const REGEXREPLACE = getTextBuiltin("REGEXREPLACE")!;
    const REGEXEXTRACT = getTextBuiltin("REGEXEXTRACT")!;

    expect(REGEXTEST(text("Alpha-42"), text("[a-z]+-[0-9]+"), number(1))).toEqual({
      tag: ValueTag.Boolean,
      value: true,
    });
    expect(REGEXTEST(text("Alpha-42"), text("^[a-z]+-[0-9]+$"))).toEqual({
      tag: ValueTag.Boolean,
      value: false,
    });

    expect(REGEXREPLACE(text("a1 b2 c3"), text("[0-9]"), text("X"))).toEqual(text("aX bX cX"));
    expect(REGEXREPLACE(text("a1 b2 c3"), text("[0-9]"), text("X"), number(2))).toEqual(
      text("a1 bX c3"),
    );
    expect(REGEXREPLACE(text("abc123"), text("([a-z]+)([0-9]+)"), text("$2-$1"))).toEqual(
      text("123-abc"),
    );

    expect(REGEXEXTRACT(text("DylanWilliams"), text("[A-Z][a-z]+"))).toEqual(text("Dylan"));
    expect(REGEXEXTRACT(text("a1 b2 c3"), text("[a-z][0-9]"), number(1))).toEqual({
      kind: "array",
      rows: 3,
      cols: 1,
      values: [text("a1"), text("b2"), text("c3")],
    });
    expect(REGEXEXTRACT(text("abc-123"), text("([a-z]+)-([0-9]+)"), number(2))).toEqual({
      kind: "array",
      rows: 1,
      cols: 2,
      values: [text("abc"), text("123")],
    });
    expect(REGEXEXTRACT(text("abc"), text("[0-9]+"))).toEqual(err(ErrorCode.NA));
    expect(REGEXTEST(text("abc"), text("("))).toEqual(valueError());
  });

  it("supports VALUETOTEXT concise and strict string rendering", () => {
    const VALUETOTEXT = getTextBuiltin("VALUETOTEXT")!;

    expect(VALUETOTEXT(number(42))).toEqual(text("42"));
    expect(VALUETOTEXT(text("alpha"))).toEqual(text("alpha"));
    expect(VALUETOTEXT(text("alpha"), number(1))).toEqual(text('"alpha"'));
    expect(VALUETOTEXT({ tag: ValueTag.Boolean, value: true })).toEqual(text("TRUE"));
    expect(VALUETOTEXT(err(ErrorCode.Ref))).toEqual(text("#REF!"));
    expect(VALUETOTEXT(text("alpha"), number(2))).toEqual(valueError());
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
    expect(getTextBuiltin("VALUE")?.(err(ErrorCode.Name))).toEqual(err(ErrorCode.Name));
    expect(getTextBuiltin("TEXTAFTER")?.()).toEqual(valueError());
    expect(getTextBuiltin("TEXTAFTER")?.(text("alpha"), text(""))).toEqual(valueError());
    expect(getTextBuiltin("TEXTAFTER")?.(text("alpha"), text("-"), number(1), number(2))).toEqual(
      valueError(),
    );
    expect(
      getTextBuiltin("TEXTAFTER")?.(text("alpha"), text("-"), number(1), number(0), number(0)),
    ).toEqual(err(ErrorCode.NA));
  });

  it("covers additional text coercion and edge cases", () => {
    const LEN = getTextBuiltin("LEN")!;
    const VALUE = getTextBuiltin("VALUE")!;
    const REPT = getTextBuiltin("REPT")!;
    const SUBSTITUTE = getTextBuiltin("SUBSTITUTE")!;
    const TEXTJOIN = getTextBuiltin("TEXTJOIN")!;

    // coerceText
    expect(LEN({ tag: ValueTag.Empty })).toEqual(number(0));
    expect(LEN({ tag: ValueTag.Boolean, value: true })).toEqual(number(4)); // "TRUE"
    expect(LEN({ tag: ValueTag.Boolean, value: false })).toEqual(number(5)); // "FALSE"

    // coerceNumber
    expect(VALUE(text("   "))).toEqual(number(0));
    expect(VALUE({ tag: ValueTag.Boolean, value: false })).toEqual(number(0));

    // REPT zero count
    expect(REPT(text("abc"), number(0))).toEqual(text(""));

    // SUBSTITUTE instance not found
    expect(SUBSTITUTE(text("abc"), text("z"), text("x"))).toEqual(text("abc"));
    expect(SUBSTITUTE(text("abc"), text("a"), text("x"), number(2))).toEqual(text("abc"));

    // TEXTJOIN with mixed empty values
    expect(
      TEXTJOIN(
        text(","),
        { tag: ValueTag.Boolean, value: true },
        text("a"),
        { tag: ValueTag.Empty },
        text(""),
        text("b"),
      ),
    ).toEqual(text("a,b"));
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
