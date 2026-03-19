import { describe, expect, it } from "vitest";
import { ErrorCode, ValueTag, type CellValue } from "@bilig/protocol";
import { TEXT_FIXTURES } from "../../../excel-fixtures/src/text-fixtures.js";
import { getTextBuiltin } from "../builtins/text.js";

describe("text builtins", () => {
  it("matches the shared text fixture corpus", () => {
    for (const group of TEXT_FIXTURES) {
      const builtin = getTextBuiltin(group.builtin);
      expect(builtin, `${group.builtin} should exist`).toBeTypeOf("function");

      for (const testCase of group.cases) {
        expect(builtin?.(...testCase.args), `${group.builtin}: ${testCase.name}`).toEqual(testCase.expected);
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
    expect(getTextBuiltin("EXACT")?.(text("Alpha"), text("Alpha"))).toEqual({ tag: ValueTag.Boolean, value: true });
    expect(getTextBuiltin("EXACT")?.(text("Alpha"), text("alpha"))).toEqual({ tag: ValueTag.Boolean, value: false });
    expect(getTextBuiltin("EXACT")?.(number(42), text("42"))).toEqual({ tag: ValueTag.Boolean, value: true });
  });
});

function number(value: number): CellValue {
  return { tag: ValueTag.Number, value };
}

function text(value: string): CellValue {
  return { tag: ValueTag.String, value, stringId: 0 };
}

function valueError(): CellValue {
  return { tag: ValueTag.Error, code: ErrorCode.Value };
}
