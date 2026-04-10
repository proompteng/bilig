import { describe, expect, it } from "vitest";
import { ErrorCode, ValueTag, type CellValue } from "@bilig/protocol";
import { getTextBuiltin } from "../builtins/text.js";

describe("text core builtins", () => {
  it("should support core cleanup and composition helpers", () => {
    // Arrange
    const CLEAN = getTextBuiltin("CLEAN")!;
    const CONCAT = getTextBuiltin("CONCAT")!;
    const CONCATENATE = getTextBuiltin("CONCATENATE")!;
    const PROPER = getTextBuiltin("PROPER")!;
    const EXACT = getTextBuiltin("EXACT")!;
    const TRIM = getTextBuiltin("TRIM")!;

    // Act
    const cleaned = CLEAN(text("a\u0001b\u007fc"));
    const concatenated = CONCAT(text("a"), number(1), text("b"));
    const concatenatedLegacy = CONCATENATE(text("x"), text("y"));
    const proper = PROPER(text("hELLO, wORLD"));
    const exact = EXACT(text("Alpha"), text("Alpha"));
    const trimmed = TRIM(text("  alpha   beta  "));

    // Assert
    expect(cleaned).toEqual(text("abc"));
    expect(concatenated).toEqual(text("a1b"));
    expect(concatenatedLegacy).toEqual(text("xy"));
    expect(proper).toEqual(text("Hello, World"));
    expect(exact).toEqual(bool(true));
    expect(trimmed).toEqual(text("alpha beta"));
  });

  it("should support localization and baht text helpers", () => {
    // Arrange
    const ASC = getTextBuiltin("ASC")!;
    const JIS = getTextBuiltin("JIS")!;
    const DBCS = getTextBuiltin("DBCS")!;
    const BAHTTEXT = getTextBuiltin("BAHTTEXT")!;
    const PHONETIC = getTextBuiltin("PHONETIC")!;

    // Act
    const asc = ASC(text("ＡＢＣ　１２３"));
    const jis = JIS(text("ABC 123"));
    const dbcs = DBCS(text("ｶﾞｷﾞｸﾞ"));
    const baht = BAHTTEXT(number(1234));
    const phonetic = PHONETIC(text("カタカナ"));

    // Assert
    expect(asc).toEqual(text("ABC 123"));
    expect(jis).toEqual(text("ＡＢＣ　１２３"));
    expect(dbcs).toEqual(text("ガギグ"));
    expect(baht).toEqual(text("หนึ่งพันสองร้อยสามสิบสี่บาทถ้วน"));
    expect(phonetic).toEqual(text("カタカナ"));
  });

  it("should return value errors for missing required args and keep explicit errors", () => {
    // Arrange
    const CONCAT = getTextBuiltin("CONCAT")!;
    const PHONETIC = getTextBuiltin("PHONETIC")!;
    const BAHTTEXT = getTextBuiltin("BAHTTEXT")!;

    // Act
    const concatMissing = CONCAT();
    const phoneticMissing = PHONETIC();
    const bahtBad = BAHTTEXT(text("bad"));
    const concatError = CONCAT(err(ErrorCode.Ref), text("x"));

    // Assert
    expect(concatMissing).toEqual(valueError());
    expect(phoneticMissing).toEqual(valueError());
    expect(bahtBad).toEqual(valueError());
    expect(concatError).toEqual(err(ErrorCode.Ref));
  });
});

// Helpers
function number(value: number): CellValue {
  return { tag: ValueTag.Number, value };
}

function text(value: string): CellValue {
  return { tag: ValueTag.String, value, stringId: 0 };
}

function bool(value: boolean): CellValue {
  return { tag: ValueTag.Boolean, value };
}

function err(code: ErrorCode): CellValue {
  return { tag: ValueTag.Error, code };
}

function valueError(): CellValue {
  return err(ErrorCode.Value);
}
