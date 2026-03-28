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
    expect(getTextBuiltin("FINDB")?.(err(ErrorCode.Ref), text("abc"))).toEqual(err(ErrorCode.Ref));
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
    expect(getTextBuiltin("LOWER")?.()).toEqual(valueError());
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
    expect(getTextBuiltin("MIDB")?.(text("abc"), number(1), text("bad"))).toEqual(valueError());
    expect(getTextBuiltin("RIGHTB")?.(text("abc"), number(-1))).toEqual(valueError());
    expect(getTextBuiltin("LEFTB")?.()).toEqual(valueError());
    expect(getTextBuiltin("MIDB")?.(text("abc"), number(1))).toEqual(valueError());
    expect(getTextBuiltin("RIGHTB")?.()).toEqual(valueError());
    expect(getTextBuiltin("SEARCHB")?.(text("a"))).toEqual(valueError());
    expect(getTextBuiltin("SEARCHB")?.(text("a"), text("abc"), number(0))).toEqual(valueError());
    expect(getTextBuiltin("SEARCHB")?.(text("a"), text("abc"), number(5))).toEqual(valueError());
    expect(getTextBuiltin("REPLACEB")?.(text("abc"), number(0), number(1), text("z"))).toEqual(
      valueError(),
    );
    expect(getTextBuiltin("REPLACEB")?.(text("abc"), number(1), number(-1), text("z"))).toEqual(
      valueError(),
    );
    expect(getTextBuiltin("REPLACEB")?.(text("abc"), number(1), text("bad"), text("z"))).toEqual(
      valueError(),
    );
    expect(getTextBuiltin("REPLACEB")?.(text("abc"), number(1), number(1))).toEqual(valueError());
    expect(
      getTextBuiltin("REPLACEB")?.(err(ErrorCode.Ref), number(1), number(1), text("z")),
    ).toEqual(err(ErrorCode.Ref));
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

  it("supports BAHTTEXT for Thai baht currency text", () => {
    expect(getTextBuiltin("BAHTTEXT")?.(number(1234))).toEqual(text("หนึ่งพันสองร้อยสามสิบสี่บาทถ้วน"));
    expect(getTextBuiltin("BAHTTEXT")?.(number(21.25))).toEqual(text("ยี่สิบเอ็ดบาทยี่สิบห้าสตางค์"));
    expect(getTextBuiltin("BAHTTEXT")?.(text("bad"))).toEqual(valueError());
  });

  it("supports TEXT for numeric, percent, date, and text-section formatting", () => {
    const TEXT = getTextBuiltin("TEXT")!;

    expect(TEXT(number(1234.567), text("#,##0.00"))).toEqual(text("1,234.57"));
    expect(TEXT(number(0.1234), text("0.0%"))).toEqual(text("12.3%"));
    expect(TEXT(number(1234.567), text("0.00E+00"))).toEqual(text("1.23E+03"));
    expect(TEXT(number(45356), text("yyyy-mm-dd"))).toEqual(text("2024-03-05"));
    expect(TEXT(number(0.5), text("h:mm AM/PM"))).toEqual(text("12:00 PM"));
    expect(TEXT(text("alpha"), text("prefix @"))).toEqual(text("prefix alpha"));
    expect(TEXT(text("alpha"), text('"literal"'))).toEqual(text("literal"));
  });

  it("returns #VALUE! for unsupported TEXT coercions and missing args", () => {
    const TEXT = getTextBuiltin("TEXT")!;

    expect(TEXT()).toEqual(valueError());
    expect(TEXT(text("alpha"), text("0.00"))).toEqual(valueError());
    expect(TEXT(err(ErrorCode.Ref), text("0.00"))).toEqual(err(ErrorCode.Ref));
    expect(TEXT(number(42), err(ErrorCode.Name))).toEqual(err(ErrorCode.Name));
  });

  it("supports PHONETIC as a scalar text passthrough", () => {
    const PHONETIC = getTextBuiltin("PHONETIC")!;

    expect(PHONETIC(text("カタカナ"))).toEqual(text("カタカナ"));
    expect(PHONETIC(number(42))).toEqual(text("42"));
    expect(PHONETIC()).toEqual(valueError());
    expect(PHONETIC(err(ErrorCode.Ref))).toEqual(err(ErrorCode.Ref));
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
      TEXTJOIN(text("|"), { tag: ValueTag.Boolean, value: false }, text("a"), text(""), text("b")),
    ).toEqual(text("a||b"));
    expect(
      TEXTJOIN(text("|"), { tag: ValueTag.Boolean, value: true }, undefined, text("a")),
    ).toEqual(text("a"));
    expect(TEXTJOIN(text("|"), text("bad"), text("a"))).toEqual(valueError());
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
      TEXTJOIN(text("|"), { tag: ValueTag.Boolean, value: false }, text("a"), text(""), text("")),
    ).toEqual(text("a||"));
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

  it("supports ASC, JIS, and DBCS width conversions", () => {
    expect(getTextBuiltin("ASC")?.(text("ＡＢＣ　１２３"))).toEqual(text("ABC 123"));
    expect(getTextBuiltin("ASC")?.(text("ガギグゲゴ"))).toEqual(text("ｶﾞｷﾞｸﾞｹﾞｺﾞ"));
    expect(getTextBuiltin("JIS")?.(text("ABC 123"))).toEqual(text("ＡＢＣ　１２３"));
    expect(getTextBuiltin("JIS")?.(text("ｶﾞｷﾞｸﾞｹﾞｺﾞ"))).toEqual(text("ガギグゲゴ"));
    expect(getTextBuiltin("DBCS")?.(text("ABC 123"))).toEqual(text("ＡＢＣ　１２３"));
    expect(getTextBuiltin("ASC")?.()).toEqual(valueError());
    expect(getTextBuiltin("JIS")?.(err(ErrorCode.Ref))).toEqual(err(ErrorCode.Ref));
    expect(getTextBuiltin("DBCS")?.(err(ErrorCode.Name))).toEqual(err(ErrorCode.Name));
  });

  it("supports NUMBERVALUE locale-independent numeric parsing", () => {
    const NUMBERVALUE = getTextBuiltin("NUMBERVALUE")!;

    expect(NUMBERVALUE()).toEqual(valueError());
    expect(NUMBERVALUE(err(ErrorCode.Name))).toEqual(err(ErrorCode.Name));
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

  it("covers remaining VALUETOTEXT and regex validation branches", () => {
    const VALUETOTEXT = getTextBuiltin("VALUETOTEXT")!;
    const REGEXTEST = getTextBuiltin("REGEXTEST")!;
    const REGEXREPLACE = getTextBuiltin("REGEXREPLACE")!;
    const REGEXEXTRACT = getTextBuiltin("REGEXEXTRACT")!;

    expect(VALUETOTEXT()).toEqual(valueError());
    expect(VALUETOTEXT(number(42), text("bad"))).toEqual(valueError());

    expect(REGEXTEST()).toEqual(valueError());
    expect(REGEXTEST(text("abc"), text("a"), number(2))).toEqual(valueError());
    expect(REGEXTEST(err(ErrorCode.Ref), text("a"))).toEqual(err(ErrorCode.Ref));

    expect(REGEXREPLACE()).toEqual(valueError());
    expect(REGEXREPLACE(err(ErrorCode.Ref), text("a"), text("x"))).toEqual(err(ErrorCode.Ref));
    expect(REGEXREPLACE(text("abc"), text("a"), text("x"), text("bad"))).toEqual(valueError());
    expect(REGEXREPLACE(text("abc"), text("a"), text("x"), number(1), number(9))).toEqual(
      valueError(),
    );
    expect(REGEXREPLACE(text("abc"), text("("), text("x"))).toEqual(valueError());
    expect(REGEXREPLACE(text("abc"), text("z"), text("x"), number(1))).toEqual(text("abc"));
    expect(
      REGEXREPLACE(text("abc123"), text("([a-z]+)([0-9]+)"), text("$2-$1"), number(3)),
    ).toEqual(text("abc123"));
    expect(REGEXREPLACE(text("abc"), text("([a-z]+)([0-9]+)?"), text("$2-$1"), number(1))).toEqual(
      text("-abc"),
    );
    expect(REGEXREPLACE(text("a1 b2 c3"), text("[0-9]"), text("X"), number(-1))).toEqual(
      text("a1 b2 cX"),
    );

    expect(REGEXEXTRACT()).toEqual(valueError());
    expect(REGEXEXTRACT(text("abc"), text("([a-z]+)"), number(3))).toEqual(valueError());
    expect(REGEXEXTRACT(text("abc"), text("([a-z]+)"), number(0), number(9))).toEqual(valueError());
    expect(REGEXEXTRACT(text("abc"), text("("))).toEqual(valueError());
    expect(REGEXEXTRACT(text("abc"), text("("), number(1))).toEqual(valueError());
    expect(REGEXEXTRACT(text("abc"), text("[0-9]+"), number(1))).toEqual(err(ErrorCode.NA));
    expect(REGEXEXTRACT(text("abc"), text("[a-z]+"), number(2))).toEqual(err(ErrorCode.NA));
    expect(REGEXEXTRACT(err(ErrorCode.Name), text("[a-z]+"))).toEqual(err(ErrorCode.Name));
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
