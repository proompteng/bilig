import { describe, expect, it } from "vitest";
import { ErrorCode, ValueTag, type CellValue } from "@bilig/protocol";
import { getTextBuiltin } from "../builtins/text.js";

describe("text search builtins", () => {
  it("supports search, regex, and delimiter helpers", () => {
    const FIND = getTextBuiltin("FIND")!;
    const SEARCHB = getTextBuiltin("SEARCHB")!;
    const REGEXREPLACE = getTextBuiltin("REGEXREPLACE")!;
    const REGEXEXTRACT = getTextBuiltin("REGEXEXTRACT")!;
    const TEXTAFTER = getTextBuiltin("TEXTAFTER")!;
    const TEXTBEFORE = getTextBuiltin("TEXTBEFORE")!;
    const TEXTJOIN = getTextBuiltin("TEXTJOIN")!;

    expect(FIND(text("ha"), text("alphabet"))).toEqual(number(4));
    expect(SEARCHB(text("é"), text("café"))).toEqual(number(4));
    expect(REGEXREPLACE(text("abc123"), text("([a-z]+)([0-9]+)"), text("$2-$1"))).toEqual(
      text("123-abc"),
    );
    expect(REGEXEXTRACT(text("abc-123"), text("([a-z]+)-([0-9]+)"), number(2))).toEqual({
      kind: "array",
      rows: 1,
      cols: 2,
      values: [text("abc"), text("123")],
    });
    expect(TEXTAFTER(text("alpha-beta-gamma"), text("-"), number(2))).toEqual(text("gamma"));
    expect(TEXTBEFORE(text("alpha-beta-gamma"), text("-"), number(-1))).toEqual(
      text("alpha-beta"),
    );
    expect(
      TEXTJOIN(text("|"), { tag: ValueTag.Boolean, value: false }, text("a"), text(""), text("b")),
    ).toEqual(text("a||b"));
  });

  it("keeps validation, wildcard escaping, and explicit error propagation for search helpers", () => {
    const FIND = getTextBuiltin("FIND")!;
    const TEXTJOIN = getTextBuiltin("TEXTJOIN")!;
    const REGEXTEST = getTextBuiltin("REGEXTEST")!;
    const TEXTBEFORE = getTextBuiltin("TEXTBEFORE")!;
    const SEARCH = getTextBuiltin("SEARCH")!;

    expect(FIND()).toEqual(valueError());
    expect(REGEXTEST(text("abc"), text("a"), number(2))).toEqual(valueError());
    expect(TEXTJOIN()).toEqual(valueError());
    expect(TEXTBEFORE(err(ErrorCode.Ref), text("-"))).toEqual(err(ErrorCode.Ref));
    expect(SEARCH(text("north~*"), text("north* zone"))).toEqual(number(1));
    expect(SEARCH(text("north*"), text("northwest zone"))).toEqual(number(1));
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
