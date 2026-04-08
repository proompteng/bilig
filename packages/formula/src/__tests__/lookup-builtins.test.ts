import { describe, expect, it } from "vitest";
import { ErrorCode, ValueTag, type CellValue } from "@bilig/protocol";
import { getLookupBuiltin, type RangeBuiltinArgument } from "../builtins/lookup.js";

const num = (value: number): CellValue => ({ tag: ValueTag.Number, value });
const text = (value: string): CellValue => ({ tag: ValueTag.String, value, stringId: 0 });
const bool = (value: boolean): CellValue => ({ tag: ValueTag.Boolean, value });
const err = (code: ErrorCode): CellValue => ({ tag: ValueTag.Error, code });

function cellRange(values: CellValue[], rows: number, cols: number): RangeBuiltinArgument {
  return { kind: "range", refKind: "cells", values, rows, cols };
}

describe("lookup builtins", () => {
  it("supports MATCH across one-dimensional cell ranges", () => {
    const MATCH = getLookupBuiltin("MATCH")!;
    expect(MATCH(text("b"), cellRange([text("a"), text("b"), text("c")], 3, 1), num(0))).toEqual(
      num(2),
    );
    expect(MATCH(num(4), cellRange([num(1), num(3), num(5)], 3, 1), num(1))).toEqual(num(2));
    expect(MATCH(num(3), cellRange([num(5), num(3), num(1)], 3, 1), num(-1))).toEqual(num(2));
    expect(MATCH(text("z"), cellRange([text("a"), text("b")], 2, 1), num(0))).toEqual(
      err(ErrorCode.NA),
    );
    expect(
      MATCH(text("x"), cellRange([text("a"), text("b"), text("c"), text("d")], 2, 2), num(0)),
    ).toEqual(err(ErrorCode.NA));
  });

  it("supports AREAS, ARRAYTOTEXT, ROWS, COLUMNS, and CORREL", () => {
    const AREAS = getLookupBuiltin("AREAS")!;
    const ARRAYTOTEXT = getLookupBuiltin("ARRAYTOTEXT")!;
    const ROWS = getLookupBuiltin("ROWS")!;
    const COLUMNS = getLookupBuiltin("COLUMNS")!;
    const CORREL = getLookupBuiltin("CORREL")!;

    const matrix = cellRange([num(1), text("x"), num(3), num(4)], 2, 2);

    expect(AREAS(matrix)).toEqual(num(1));
    expect(ROWS(matrix)).toEqual(num(2));
    expect(COLUMNS(matrix)).toEqual(num(2));
    expect(ARRAYTOTEXT(matrix)).toEqual(text("1\tx;3\t4"));
    expect(ARRAYTOTEXT(matrix, num(1))).toEqual(text('{1, "x";3, 4}'));
    expect(
      CORREL(cellRange([num(1), num(2), num(3)], 3, 1), cellRange([num(1), num(2), num(3)], 3, 1)),
    ).toEqual(num(1));
    expect(
      CORREL(cellRange([num(1), num(2), num(3)], 3, 1), cellRange([num(1), num(2)], 2, 1)),
    ).toEqual(err(ErrorCode.Value));
  });

  it("maps USE.THE.COUNTIF to the COUNTIF lookup implementation", () => {
    const COUNTIF = getLookupBuiltin("COUNTIF")!;
    const alias = getLookupBuiltin("USE.THE.COUNTIF")!;
    const sample = cellRange([num(1), num(0), num(3)], 3, 1);

    expect(alias).toBe(COUNTIF);
    expect(alias(sample, text(">0"))).toEqual(num(2));
  });

  it("supports database aggregation builtins on matching records", () => {
    const DAVERAGE = getLookupBuiltin("DAVERAGE")!;
    const DCOUNT = getLookupBuiltin("DCOUNT")!;
    const DCOUNTA = getLookupBuiltin("DCOUNTA")!;
    const DGET = getLookupBuiltin("DGET")!;
    const DMAX = getLookupBuiltin("DMAX")!;
    const DMIN = getLookupBuiltin("DMIN")!;
    const DPRODUCT = getLookupBuiltin("DPRODUCT")!;
    const DSTDEV = getLookupBuiltin("DSTDEV")!;
    const DSTDEVP = getLookupBuiltin("DSTDEVP")!;
    const DSUM = getLookupBuiltin("DSUM")!;
    const DVAR = getLookupBuiltin("DVAR")!;
    const DVARP = getLookupBuiltin("DVARP")!;

    const database = cellRange(
      [
        text("Age"),
        text("Height"),
        text("Yield"),
        num(10),
        num(100),
        num(5),
        num(12),
        num(110),
        num(7),
        num(12),
        num(120),
        num(9),
        num(15),
        num(130),
        num(11),
      ],
      5,
      3,
    );
    const ageIsTwelve = cellRange([text("Age"), num(12)], 2, 1);
    const ageIsFifteen = cellRange([text("Age"), num(15)], 2, 1);

    expect(DAVERAGE(database, text("Yield"), ageIsTwelve)).toEqual(num(8));
    expect(DCOUNT(database, text("Yield"), ageIsTwelve)).toEqual(num(2));
    expect(DCOUNT(database, { tag: ValueTag.Empty }, ageIsTwelve)).toEqual(num(2));
    expect(DCOUNTA(database, text("Height"), ageIsTwelve)).toEqual(num(2));
    expect(DCOUNTA(database, { tag: ValueTag.Empty }, ageIsTwelve)).toEqual(num(2));
    expect(DGET(database, text("Height"), ageIsFifteen)).toEqual(num(130));
    expect(DMAX(database, text("Yield"), ageIsTwelve)).toEqual(num(9));
    expect(DMIN(database, text("Yield"), ageIsTwelve)).toEqual(num(7));
    expect(DPRODUCT(database, text("Yield"), ageIsTwelve)).toEqual(num(63));
    expect(DSUM(database, text("Yield"), ageIsTwelve)).toEqual(num(16));
    expect(DVAR(database, text("Yield"), ageIsTwelve)).toEqual(num(2));
    expect(DVARP(database, text("Yield"), ageIsTwelve)).toEqual(num(1));

    const dstdev = DSTDEV(database, text("Yield"), ageIsTwelve);
    if (dstdev.tag !== ValueTag.Number) {
      throw new Error("DSTDEV should return a number");
    }
    expect(dstdev.value).toBeCloseTo(Math.SQRT2, 12);

    const dstdevp = DSTDEVP(database, text("Yield"), ageIsTwelve);
    if (dstdevp.tag !== ValueTag.Number) {
      throw new Error("DSTDEVP should return a number");
    }
    expect(dstdevp.value).toBeCloseTo(1, 12);

    expect(DGET(database, text("Yield"), ageIsTwelve)).toEqual(err(ErrorCode.Value));
    expect(DAVERAGE(database, text("Missing"), ageIsTwelve)).toEqual(err(ErrorCode.Value));
    expect(
      DCOUNT(database, text("Yield"), cellRange([text("Age"), err(ErrorCode.Ref)], 2, 1)),
    ).toEqual(err(ErrorCode.Ref));
  });

  it("covers database builtin validation and empty-match edge cases", () => {
    const DAVERAGE = getLookupBuiltin("DAVERAGE")!;
    const DCOUNT = getLookupBuiltin("DCOUNT")!;
    const DGET = getLookupBuiltin("DGET")!;
    const DMAX = getLookupBuiltin("DMAX")!;
    const DPRODUCT = getLookupBuiltin("DPRODUCT")!;
    const DSTDEVP = getLookupBuiltin("DSTDEVP")!;
    const DVAR = getLookupBuiltin("DVAR")!;
    const DVARP = getLookupBuiltin("DVARP")!;

    const database = cellRange(
      [
        text("Age"),
        text("Height"),
        text("Yield"),
        num(10),
        num(100),
        num(5),
        num(12),
        num(110),
        num(7),
        num(12),
        num(120),
        num(9),
        num(15),
        num(130),
        num(11),
      ],
      5,
      3,
    );
    const ageIsTwelve = cellRange([text("Age"), num(12)], 2, 1);
    const ageMissing = cellRange([text("Age"), num(99)], 2, 1);

    expect(
      DGET(database, cellRange([text("Height")], 1, 1), cellRange([text("Age"), num(15)], 2, 1)),
    ).toEqual(num(130));
    expect(DCOUNT(database, text(""), ageIsTwelve)).toEqual(num(2));
    expect(DAVERAGE(database, text(""), ageIsTwelve)).toEqual(err(ErrorCode.Value));
    expect(DCOUNT(database, cellRange([text("Yield"), text("Height")], 1, 2), ageIsTwelve)).toEqual(
      err(ErrorCode.Value),
    );
    expect(DMAX(database, bool(true), ageIsTwelve)).toEqual(err(ErrorCode.Value));
    expect(DMAX(database, num(0), ageIsTwelve)).toEqual(err(ErrorCode.Value));
    expect(DCOUNT(database, text("Yield"), cellRange([text("Age")], 1, 1))).toEqual(
      err(ErrorCode.Value),
    );
    expect(DCOUNT(database, text("Yield"), cellRange([err(ErrorCode.Ref), num(12)], 2, 1))).toEqual(
      err(ErrorCode.Ref),
    );
    expect(
      DCOUNT(database, text("Yield"), cellRange([text("Age"), err(ErrorCode.Name)], 2, 1)),
    ).toEqual(err(ErrorCode.Name));
    expect(
      DCOUNT(database, text("Yield"), cellRange([{ tag: ValueTag.Empty }, num(12)], 2, 1)),
    ).toEqual(num(0));
    expect(DCOUNT(database, text("Yield"), cellRange([text("Missing"), num(12)], 2, 1))).toEqual(
      num(0),
    );
    expect(DAVERAGE(database, text("Yield"), ageMissing)).toEqual(err(ErrorCode.Div0));
    expect(DMAX(database, text("Yield"), ageMissing)).toEqual(num(0));
    expect(DPRODUCT(database, text("Yield"), ageMissing)).toEqual(num(0));
    expect(DSTDEVP(database, text("Yield"), ageMissing)).toEqual(err(ErrorCode.Div0));
    expect(
      DVAR(database, text("Yield"), cellRange([text("Age"), err(ErrorCode.Ref)], 2, 1)),
    ).toEqual(err(ErrorCode.Ref));
    expect(DVARP(database, text("Missing"), ageIsTwelve)).toEqual(err(ErrorCode.Value));
  });

  it("supports COVAR, COVARIANCE.P, COVARIANCE.S, AVEDEV, and DEVSQ", () => {
    const COVAR = getLookupBuiltin("COVAR")!;
    const COVARP = getLookupBuiltin("COVARIANCE.P")!;
    const COVARS = getLookupBuiltin("COVARIANCE.S")!;
    const AVEDEV = getLookupBuiltin("AVEDEV")!;
    const DEVSQ = getLookupBuiltin("DEVSQ")!;

    const first = cellRange([num(1), num(2), num(3)], 3, 1);
    const second = cellRange([num(4), num(5), num(6)], 3, 1);
    expect(COVAR(first, second)).toEqual(num(2 / 3));
    expect(COVARP(first, second)).toEqual(num(2 / 3));
    expect(COVARS(first, second)).toEqual(num(1));

    expect(
      COVAR(
        cellRange([num(1), num(2), num(3), num(4)], 2, 2),
        cellRange([num(1), num(2), num(3)], 3, 1),
      ),
    ).toEqual(err(ErrorCode.Value));

    expect(COVARS(cellRange([num(2)], 1, 1), cellRange([num(4)], 1, 1))).toEqual(
      err(ErrorCode.Div0),
    );

    expect(AVEDEV(num(1), num(2), num(3))).toEqual(num(2 / 3));
    expect(DEVSQ(num(1), num(2), num(3))).toEqual(num(2));
    expect(AVEDEV(cellRange([text("bad")], 1, 1))).toEqual(err(ErrorCode.Value));
    expect(DEVSQ(cellRange([text("bad")], 1, 1))).toEqual(err(ErrorCode.Value));
  });

  it("supports CHISQ.TEST and legacy aliases across actual and expected matrices", () => {
    const CHISQ_TEST = getLookupBuiltin("CHISQ.TEST")!;
    const CHITEST = getLookupBuiltin("CHITEST")!;
    const LEGACY_CHITEST = getLookupBuiltin("LEGACY.CHITEST")!;

    const actual = cellRange([num(58), num(35), num(11), num(25), num(10), num(23)], 3, 2);
    const expected = cellRange(
      [num(45.35), num(47.65), num(17.56), num(18.44), num(16.09), num(16.91)],
      3,
      2,
    );

    const chisqTest = CHISQ_TEST(actual, expected);
    if (chisqTest.tag !== ValueTag.Number) {
      throw new Error("CHISQ.TEST should return a number");
    }
    expect(chisqTest.value).toBeCloseTo(0.0003082, 7);
    const chiTest = CHITEST(actual, expected);
    if (chiTest.tag !== ValueTag.Number) {
      throw new Error("CHITEST should return a number");
    }
    expect(chiTest.value).toBeCloseTo(0.0003082, 7);
    const legacyChiTest = LEGACY_CHITEST(actual, expected);
    if (legacyChiTest.tag !== ValueTag.Number) {
      throw new Error("LEGACY.CHITEST should return a number");
    }
    expect(legacyChiTest.value).toBeCloseTo(0.0003082, 7);

    expect(CHISQ_TEST(cellRange([num(1), num(2)], 2, 1), cellRange([num(1)], 1, 1))).toEqual(
      err(ErrorCode.NA),
    );
    expect(CHISQ_TEST(num(1), num(1))).toEqual(err(ErrorCode.NA));
    expect(
      CHISQ_TEST(cellRange([num(1), num(2)], 2, 1), cellRange([num(1), num(0)], 2, 1)),
    ).toEqual(err(ErrorCode.Div0));
  });

  it("supports F.TEST and Z.TEST legacy aliases on numeric samples", () => {
    const F_TEST = getLookupBuiltin("F.TEST")!;
    const FTEST = getLookupBuiltin("FTEST")!;
    const Z_TEST = getLookupBuiltin("Z.TEST")!;
    const ZTEST = getLookupBuiltin("ZTEST")!;

    const first = cellRange([num(6), num(7), num(9), num(15), num(21)], 5, 1);
    const second = cellRange([num(20), num(28), num(31), num(38), num(40)], 5, 1);
    const zSample = cellRange([num(1), num(2), num(3), num(4), num(5)], 5, 1);

    const fTest = F_TEST(first, second);
    if (fTest.tag !== ValueTag.Number) {
      throw new Error("F.TEST should return a number");
    }
    expect(fTest.value).toBeCloseTo(0.648317846786175, 12);

    const legacyFTest = FTEST(first, second);
    if (legacyFTest.tag !== ValueTag.Number) {
      throw new Error("FTEST should return a number");
    }
    expect(legacyFTest.value).toBeCloseTo(0.648317846786175, 12);

    const zTest = Z_TEST(zSample, num(2), num(1));
    if (zTest.tag !== ValueTag.Number) {
      throw new Error("Z.TEST should return a number");
    }
    expect(zTest.value).toBeCloseTo(0.012673617875446075, 12);

    const legacyZTest = ZTEST(zSample, num(2), num(1));
    if (legacyZTest.tag !== ValueTag.Number) {
      throw new Error("ZTEST should return a number");
    }
    expect(legacyZTest.value).toBeCloseTo(0.012673617875446075, 12);

    expect(F_TEST(cellRange([num(1), text("x")], 2, 1), cellRange([num(1)], 1, 1))).toEqual(
      err(ErrorCode.Div0),
    );
    expect(F_TEST(cellRange([num(3), num(3)], 2, 1), cellRange([num(1), num(2)], 2, 1))).toEqual(
      err(ErrorCode.Div0),
    );
    expect(Z_TEST(zSample, num(2), num(0))).toEqual(err(ErrorCode.Div0));
  });

  it("supports T.TEST across paired, equal-variance, and Welch modes", () => {
    const T_TEST = getLookupBuiltin("T.TEST")!;
    const TTEST = getLookupBuiltin("TTEST")!;

    const pairedFirst = cellRange([num(1), num(2), num(4)], 3, 1);
    const pairedSecond = cellRange([num(1), num(3), num(3)], 3, 1);
    expect(T_TEST(pairedFirst, pairedSecond, num(2), num(1))).toEqual(num(1));
    expect(TTEST(pairedFirst, pairedSecond, num(2), num(1))).toEqual(num(1));

    const independentFirst = cellRange([num(6), num(7), num(9), num(15), num(21)], 5, 1);
    const independentSecond = cellRange([num(20), num(28), num(31), num(38), num(40)], 5, 1);
    const equalVariance = T_TEST(independentFirst, independentSecond, num(2), num(2));
    if (equalVariance.tag !== ValueTag.Number) {
      throw new Error("T.TEST equal-variance mode should return a number");
    }
    expect(equalVariance.value).toBeCloseTo(0.0025154774780675737, 12);

    const welch = T_TEST(independentFirst, independentSecond, num(2), num(3));
    if (welch.tag !== ValueTag.Number) {
      throw new Error("T.TEST Welch mode should return a number");
    }
    expect(welch.value).toBeGreaterThan(equalVariance.value);
    expect(welch.value).toBeLessThan(0.01);

    expect(T_TEST(independentFirst, independentSecond, num(3), num(2))).toEqual(
      err(ErrorCode.Value),
    );
    expect(T_TEST(independentFirst, independentSecond, num(2), num(4))).toEqual(
      err(ErrorCode.Value),
    );
    expect(
      T_TEST(cellRange([num(1), num(2)], 2, 1), cellRange([num(1)], 1, 1), num(2), num(1)),
    ).toEqual(err(ErrorCode.NA));
  });

  it("supports LOOKUP, TRANSPOSE, HSTACK, VSTACK, and PEARSON", () => {
    const LOOKUP = getLookupBuiltin("LOOKUP")!;
    const TRANSPOSE = getLookupBuiltin("TRANSPOSE")!;
    const HSTACK = getLookupBuiltin("HSTACK")!;
    const VSTACK = getLookupBuiltin("VSTACK")!;
    const PEARSON = getLookupBuiltin("PEARSON")!;

    const lookupValues = cellRange([num(1), num(2), num(3)], 3, 1);
    const resultValues = cellRange([num(10), num(20), num(30)], 3, 1);

    expect(LOOKUP(num(2), lookupValues, resultValues)).toEqual(num(20));
    expect(LOOKUP(num(4), lookupValues, resultValues)).toEqual(num(30));
    expect(
      LOOKUP(
        text("not-found"),
        cellRange([text("a"), text("b"), text("c")], 3, 1),
        cellRange([num(1), num(2), num(3)], 3, 1),
      ),
    ).toEqual(err(ErrorCode.NA));

    expect(TRANSPOSE(cellRange([num(1), num(2), num(3), num(4), num(5), num(6)], 2, 3))).toEqual({
      kind: "array",
      rows: 3,
      cols: 2,
      values: [num(1), num(4), num(2), num(5), num(3), num(6)],
    });
    expect(TRANSPOSE(num(7))).toEqual(num(7));

    expect(
      HSTACK(
        cellRange([num(1), num(2), num(3)], 3, 1),
        cellRange([text("a"), text("b")], 1, 2),
        num(99),
      ),
    ).toEqual({
      kind: "array",
      rows: 3,
      cols: 4,
      values: [
        num(1),
        text("a"),
        text("b"),
        num(99),
        num(2),
        text("a"),
        text("b"),
        num(99),
        num(3),
        text("a"),
        text("b"),
        num(99),
      ],
    });

    expect(
      VSTACK(
        cellRange([text("x"), text("y")], 1, 2),
        cellRange([num(3), num(4), num(5), num(6)], 2, 2),
        num(7),
      ),
    ).toEqual({
      kind: "array",
      rows: 4,
      cols: 2,
      values: [text("x"), text("y"), num(3), num(4), num(5), num(6), num(7), num(7)],
    });

    expect(PEARSON(lookupValues, resultValues)).toEqual(num(1));
    expect(PEARSON(cellRange([num(1)], 1, 1), cellRange([num(2)], 1, 1))).toEqual(
      err(ErrorCode.Div0),
    );
  });

  it("validates TOCOL and TOROW control arguments", () => {
    const TOCOL = getLookupBuiltin("TOCOL")!;
    const TOROW = getLookupBuiltin("TOROW")!;
    const matrix = cellRange([num(1), num(2), num(3), num(4)], 2, 2);

    expect(TOCOL(matrix, num(2))).toEqual(err(ErrorCode.Value));
    expect(TOCOL(matrix, num(0), text("bad"))).toEqual(err(ErrorCode.Value));
    expect(TOROW(matrix, num(2))).toEqual(err(ErrorCode.Value));
    expect(TOROW(matrix, num(0), text("bad"))).toEqual(err(ErrorCode.Value));
  });

  it("supports INDEX over cell ranges", () => {
    const INDEX = getLookupBuiltin("INDEX")!;
    const matrix = cellRange([num(10), num(11), num(20), num(21)], 2, 2);
    expect(INDEX(matrix, num(2), num(1))).toEqual(num(20));
    expect(INDEX(cellRange([text("a"), text("b"), text("c")], 1, 3), num(2))).toEqual(text("b"));
    expect(INDEX(matrix, num(3), num(1))).toEqual(err(ErrorCode.Ref));
    expect(INDEX(matrix, text("oops"))).toEqual(err(ErrorCode.Value));
  });

  it("supports exact and approximate VLOOKUP", () => {
    const VLOOKUP = getLookupBuiltin("VLOOKUP")!;
    const table = cellRange([text("a"), num(10), text("b"), num(20), text("c"), num(30)], 3, 2);

    expect(VLOOKUP(text("b"), table, num(2), bool(false))).toEqual(num(20));
    expect(VLOOKUP(text("bb"), table, num(2), bool(true))).toEqual(num(20));
    expect(VLOOKUP(text("z"), table, num(2), bool(false))).toEqual(err(ErrorCode.NA));
    expect(VLOOKUP(text("a"), table, num(3), bool(false))).toEqual(err(ErrorCode.Value));
  });

  it("supports exact XLOOKUP and conditional aggregates", () => {
    const XLOOKUP = getLookupBuiltin("XLOOKUP")!;
    const XMATCH = getLookupBuiltin("XMATCH")!;
    const HLOOKUP = getLookupBuiltin("HLOOKUP")!;
    const COUNTIF = getLookupBuiltin("COUNTIF")!;
    const COUNTIFS = getLookupBuiltin("COUNTIFS")!;
    const SUMIF = getLookupBuiltin("SUMIF")!;
    const SUMIFS = getLookupBuiltin("SUMIFS")!;
    const AVERAGEIF = getLookupBuiltin("AVERAGEIF")!;
    const AVERAGEIFS = getLookupBuiltin("AVERAGEIFS")!;
    const MINIFS = getLookupBuiltin("MINIFS")!;
    const MAXIFS = getLookupBuiltin("MAXIFS")!;
    const SUMPRODUCT = getLookupBuiltin("SUMPRODUCT")!;

    expect(
      XLOOKUP(
        text("pear"),
        cellRange([text("apple"), text("pear"), text("plum")], 3, 1),
        cellRange([num(10), num(20), num(30)], 3, 1),
      ),
    ).toEqual(num(20));

    expect(
      XLOOKUP(
        text("missing"),
        cellRange([text("apple"), text("pear"), text("plum")], 3, 1),
        cellRange([num(10), num(20), num(30)], 3, 1),
        text("fallback"),
      ),
    ).toEqual(text("fallback"));

    expect(COUNTIF(cellRange([num(2), num(4), num(-1), num(6)], 4, 1), text(">0"))).toEqual(num(3));

    expect(
      COUNTIFS(
        cellRange([num(2), num(4), num(-1), num(6)], 4, 1),
        text(">0"),
        cellRange([text("a"), text("a"), text("b"), text("a")], 4, 1),
        text("a"),
      ),
    ).toEqual(num(3));

    expect(SUMIF(cellRange([num(2), num(4), num(-1), num(6)], 4, 1), text(">0"))).toEqual(num(12));

    expect(
      SUMIFS(
        cellRange([num(10), num(20), num(30), num(40)], 4, 1),
        cellRange([num(2), num(4), num(-1), num(6)], 4, 1),
        text(">0"),
        cellRange([text("a"), text("a"), text("b"), text("a")], 4, 1),
        text("a"),
      ),
    ).toEqual(num(70));

    expect(AVERAGEIF(cellRange([num(2), num(4), num(-1), num(6)], 4, 1), text(">0"))).toEqual(
      num(4),
    );

    expect(
      AVERAGEIFS(
        cellRange([num(10), num(20), num(30), num(40)], 4, 1),
        cellRange([num(2), num(4), num(-1), num(6)], 4, 1),
        text(">0"),
        cellRange([text("a"), text("a"), text("b"), text("a")], 4, 1),
        text("a"),
      ),
    ).toEqual(num((10 + 20 + 40) / 3));

    expect(
      MINIFS(
        cellRange([num(10), { tag: ValueTag.Empty }, num(30), num(5)], 4, 1),
        cellRange([num(2), num(4), num(-1), num(6)], 4, 1),
        text(">0"),
        cellRange([text("a"), text("a"), text("b"), text("a")], 4, 1),
        text("a"),
      ),
    ).toEqual(num(5));

    expect(
      MAXIFS(
        cellRange([num(10), text("skip"), num(30), num(5)], 4, 1),
        cellRange([num(2), num(4), num(-1), num(6)], 4, 1),
        text(">0"),
        cellRange([text("a"), text("a"), text("b"), text("a")], 4, 1),
        text("a"),
      ),
    ).toEqual(num(10));

    expect(
      MINIFS(
        cellRange([{ tag: ValueTag.Empty }, text("skip")], 2, 1),
        cellRange([text("a"), text("a")], 2, 1),
        text("a"),
      ),
    ).toEqual(num(0));

    expect(
      MAXIFS(cellRange([num(1), num(2)], 2, 1), cellRange([num(1)], 1, 1), text(">0")),
    ).toEqual(err(ErrorCode.Value));

    expect(
      SUMPRODUCT(
        cellRange([num(1), num(2), num(3)], 3, 1),
        cellRange([num(4), num(5), num(6)], 3, 1),
      ),
    ).toEqual(num(32));

    expect(
      XMATCH(text("pear"), cellRange([text("apple"), text("pear"), text("plum")], 3, 1)),
    ).toEqual(num(2));

    expect(
      HLOOKUP(
        text("pear"),
        cellRange([text("apple"), text("pear"), text("plum"), num(10), num(20), num(30)], 2, 3),
        num(2),
        bool(false),
      ),
    ).toEqual(num(20));
  });

  it("covers conditional aggregate validation and error branches", () => {
    const COUNTIF = getLookupBuiltin("COUNTIF")!;
    const COUNTIFS = getLookupBuiltin("COUNTIFS")!;
    const SUMIF = getLookupBuiltin("SUMIF")!;
    const SUMIFS = getLookupBuiltin("SUMIFS")!;
    const AVERAGEIF = getLookupBuiltin("AVERAGEIF")!;
    const AVERAGEIFS = getLookupBuiltin("AVERAGEIFS")!;

    const values = cellRange([num(2), num(4), num(-1)], 3, 1);
    const otherValues = cellRange([num(10), text("skip"), num(30)], 3, 1);

    expect(COUNTIF(num(2), text(">0"))).toEqual(err(ErrorCode.Value));
    expect(COUNTIF(values, cellRange([text(">0")], 1, 1))).toEqual(err(ErrorCode.Value));
    expect(COUNTIF(values, err(ErrorCode.Name))).toEqual(err(ErrorCode.Name));

    expect(COUNTIFS()).toEqual(err(ErrorCode.Value));
    expect(COUNTIFS(values)).toEqual(err(ErrorCode.Value));
    expect(COUNTIFS(num(1), text(">0"))).toEqual(err(ErrorCode.Value));
    expect(COUNTIFS(values, cellRange([text(">0")], 1, 1))).toEqual(err(ErrorCode.Value));
    expect(COUNTIFS(values, err(ErrorCode.Name))).toEqual(err(ErrorCode.Name));
    expect(
      COUNTIFS(values, text(">0"), cellRange([text("a"), text("b")], 2, 1), text("a")),
    ).toEqual(err(ErrorCode.Value));

    expect(SUMIF(num(2), text(">0"))).toEqual(err(ErrorCode.Value));
    expect(SUMIF(values, text(">0"), num(2))).toEqual(err(ErrorCode.Value));
    expect(SUMIF(values, text(">0"), cellRange([num(10), num(20)], 2, 1))).toEqual(
      err(ErrorCode.Value),
    );
    expect(SUMIF(values, cellRange([text(">0")], 1, 1))).toEqual(err(ErrorCode.Value));
    expect(SUMIF(values, err(ErrorCode.Name))).toEqual(err(ErrorCode.Name));

    expect(SUMIFS(num(10))).toEqual(err(ErrorCode.Value));
    expect(SUMIFS(values)).toEqual(err(ErrorCode.Value));
    expect(SUMIFS(num(10), values, text(">0"))).toEqual(err(ErrorCode.Value));
    expect(SUMIFS(values, values, cellRange([text(">0")], 1, 1))).toEqual(err(ErrorCode.Value));
    expect(SUMIFS(values, values, err(ErrorCode.Name))).toEqual(err(ErrorCode.Name));
    expect(SUMIFS(values, cellRange([text("a"), text("b")], 2, 1), text("a"))).toEqual(
      err(ErrorCode.Value),
    );

    expect(AVERAGEIF(num(2), text(">0"))).toEqual(err(ErrorCode.Value));
    expect(AVERAGEIF(values, text(">0"), num(2))).toEqual(err(ErrorCode.Value));
    expect(AVERAGEIF(values, text(">0"), cellRange([num(10), num(20)], 2, 1))).toEqual(
      err(ErrorCode.Value),
    );
    expect(AVERAGEIF(values, cellRange([text(">0")], 1, 1))).toEqual(err(ErrorCode.Value));
    expect(AVERAGEIF(values, err(ErrorCode.Name))).toEqual(err(ErrorCode.Name));
    expect(AVERAGEIF(values, text("<-100"))).toEqual(err(ErrorCode.Div0));
    expect(AVERAGEIF(values, text(">0"), otherValues)).toEqual(num(10));

    expect(AVERAGEIFS(num(10))).toEqual(err(ErrorCode.Value));
    expect(AVERAGEIFS(values)).toEqual(err(ErrorCode.Value));
    expect(AVERAGEIFS(num(10), values, text(">0"))).toEqual(err(ErrorCode.Value));
    expect(AVERAGEIFS(values, values, cellRange([text(">0")], 1, 1))).toEqual(err(ErrorCode.Value));
    expect(AVERAGEIFS(values, values, err(ErrorCode.Name))).toEqual(err(ErrorCode.Name));
    expect(AVERAGEIFS(values, cellRange([text("a"), text("b")], 2, 1), text("a"))).toEqual(
      err(ErrorCode.Value),
    );
    expect(AVERAGEIFS(values, values, text("<-100"))).toEqual(err(ErrorCode.Div0));
  });

  it("supports OFFSET, TAKE, and DROP shape transformations", () => {
    const OFFSET = getLookupBuiltin("OFFSET")!;
    const TAKE = getLookupBuiltin("TAKE")!;
    const DROP = getLookupBuiltin("DROP")!;

    const matrix = cellRange([num(1), num(2), num(3), num(4), num(5), num(6)], 3, 2);

    expect(OFFSET(matrix, num(1), num(0), num(1), num(1))).toEqual(num(3));
    expect(OFFSET(matrix, num(-1), num(-1), num(1), num(1))).toEqual(num(6));
    expect(TAKE(matrix, num(2), num(1))).toEqual({
      kind: "array",
      rows: 2,
      cols: 1,
      values: [num(1), num(3)],
    });
    expect(TAKE(matrix, num(-2), num(2))).toEqual({
      kind: "array",
      rows: 2,
      cols: 2,
      values: [num(3), num(4), num(5), num(6)],
    });
    expect(DROP(cellRange([num(1), num(2), num(3), num(4)], 4, 1), num(2))).toEqual({
      kind: "array",
      rows: 2,
      cols: 1,
      values: [num(3), num(4)],
    });
    expect(DROP(cellRange([num(1), num(2), num(3), num(4)], 1, 4), num(0), num(2))).toEqual({
      kind: "array",
      rows: 1,
      cols: 2,
      values: [num(3), num(4)],
    });
    expect(DROP(cellRange([num(1), num(2), num(3), num(4)], 1, 4), num(4))).toEqual(
      err(ErrorCode.Value),
    );
    expect(DROP(cellRange([num(1), num(2), num(3), num(4)], 4, 1), num(4))).toEqual(
      err(ErrorCode.Value),
    );
  });

  it("supports CHOOSECOLS and CHOOSEROWS extraction", () => {
    const CHOOSECOLS = getLookupBuiltin("CHOOSECOLS")!;
    const CHOOSEROWS = getLookupBuiltin("CHOOSEROWS")!;
    const matrix = cellRange([num(1), num(2), num(3), num(4), num(5), num(6)], 3, 2);

    expect(CHOOSECOLS(matrix, num(2))).toEqual({
      kind: "array",
      rows: 3,
      cols: 1,
      values: [num(2), num(4), num(6)],
    });
    expect(CHOOSECOLS(matrix, num(2), num(1))).toEqual({
      kind: "array",
      rows: 3,
      cols: 2,
      values: [num(2), num(1), num(4), num(3), num(6), num(5)],
    });
    expect(CHOOSEROWS(matrix, num(3), num(1))).toEqual({
      kind: "array",
      rows: 2,
      cols: 2,
      values: [num(5), num(6), num(1), num(2)],
    });
  });

  it("supports SORT and SORTBY ordering", () => {
    const SORT = getLookupBuiltin("SORT")!;
    const SORTBY = getLookupBuiltin("SORTBY")!;

    expect(SORT(cellRange([num(3), num(1), num(2)], 3, 1))).toEqual({
      kind: "array",
      rows: 3,
      cols: 1,
      values: [num(1), num(2), num(3)],
    });
    expect(SORT(cellRange([num(3), num(1), num(2), num(4)], 2, 2), num(1))).toEqual({
      kind: "array",
      rows: 2,
      cols: 2,
      values: [num(2), num(4), num(3), num(1)],
    });

    expect(
      SORT(
        cellRange([num(9), num(2), num(8), num(1), num(5), num(7)], 2, 3),
        num(2),
        num(1),
        bool(true),
      ),
    ).toEqual({
      kind: "array",
      rows: 2,
      cols: 3,
      values: [num(9), num(2), num(8), num(1), num(5), num(7)],
    });

    expect(
      SORTBY(
        cellRange([text("pear"), text("apple"), text("plum")], 3, 1),
        cellRange([num(2), num(1), num(3)], 3, 1),
      ),
    ).toEqual({
      kind: "array",
      rows: 3,
      cols: 1,
      values: [text("apple"), text("pear"), text("plum")],
    });
    expect(
      SORTBY(
        cellRange([num(2), num(1), num(3)], 3, 1),
        cellRange([num(5), num(1), num(3)], 3, 1),
        num(-1),
      ),
    ).toEqual({
      kind: "array",
      rows: 3,
      cols: 1,
      values: [num(2), num(3), num(1)],
    });
  });

  it("supports TOCOL and TOROW flattening modes", () => {
    const TOCOL = getLookupBuiltin("TOCOL")!;
    const TOROW = getLookupBuiltin("TOROW")!;
    const matrix = cellRange([num(1), num(2), num(3), num(4)], 2, 2);

    expect(TOCOL(matrix)).toEqual({
      kind: "array",
      rows: 4,
      cols: 1,
      values: [num(1), num(3), num(2), num(4)],
    });
    expect(TOROW(matrix)).toEqual({
      kind: "array",
      rows: 1,
      cols: 4,
      values: [num(1), num(2), num(3), num(4)],
    });
  });

  it("covers TOCOL and TOROW argument edge cases", () => {
    const TOCOL = getLookupBuiltin("TOCOL")!;
    const TOROW = getLookupBuiltin("TOROW")!;
    const rowRefRange: RangeBuiltinArgument = {
      kind: "range",
      refKind: "rows",
      rows: 1,
      cols: 1,
      values: [num(1)],
    };

    expect(TOCOL(num(1))).toEqual({
      kind: "array",
      rows: 1,
      cols: 1,
      values: [num(1)],
    });
    expect(TOCOL(rowRefRange)).toEqual(err(ErrorCode.Value));
    expect(
      TOCOL(cellRange([num(1), num(2), num(3), num(4)], 2, 2), cellRange([num(1)], 1, 1)),
    ).toEqual(err(ErrorCode.Value));
    expect(TOCOL(cellRange([num(1), num(2), num(3), num(4)], 2, 2), err(ErrorCode.Name))).toEqual(
      err(ErrorCode.Name),
    );
    expect(
      TOCOL(cellRange([num(1), num(2), num(3), num(4)], 2, 2), num(0), err(ErrorCode.Ref)),
    ).toEqual(err(ErrorCode.Ref));
    expect(TOROW(num(1))).toEqual({
      kind: "array",
      rows: 1,
      cols: 1,
      values: [num(1)],
    });
    expect(TOROW(rowRefRange)).toEqual(err(ErrorCode.Value));
    expect(
      TOROW(cellRange([num(1), num(2), num(3), num(4)], 2, 2), cellRange([num(1)], 1, 1)),
    ).toEqual(err(ErrorCode.Value));
    expect(TOROW(cellRange([num(1), num(2), num(3), num(4)], 2, 2), err(ErrorCode.Name))).toEqual(
      err(ErrorCode.Name),
    );
    expect(
      TOROW(cellRange([num(1), num(2), num(3), num(4)], 2, 2), num(0), err(ErrorCode.Ref)),
    ).toEqual(err(ErrorCode.Ref));
    expect(TOROW(cellRange([num(1), num(2), num(3), num(4)], 2, 2), num(2))).toEqual(
      err(ErrorCode.Value),
    );
    expect(TOROW(cellRange([num(1), num(2), num(3), num(4)], 2, 2), num(0), text("bad"))).toEqual(
      err(ErrorCode.Value),
    );
  });

  it("supports WRAPROWS and WRAPCOLS packing", () => {
    const WRAPROWS = getLookupBuiltin("WRAPROWS")!;
    const WRAPCOLS = getLookupBuiltin("WRAPCOLS")!;
    const vector = cellRange([num(1), num(2), num(3), num(4), num(5)], 5, 1);
    const rowRefRange: RangeBuiltinArgument = {
      kind: "range",
      refKind: "rows",
      rows: 1,
      cols: 1,
      values: [num(1)],
    };

    expect(WRAPROWS(vector, num(2))).toEqual({
      kind: "array",
      rows: 3,
      cols: 2,
      values: [num(1), num(2), num(3), num(4), num(5), err(ErrorCode.NA)],
    });
    expect(WRAPCOLS(vector, num(2))).toEqual({
      kind: "array",
      rows: 2,
      cols: 3,
      values: [num(1), num(3), num(5), num(2), num(4), err(ErrorCode.NA)],
    });
    expect(WRAPCOLS(vector, num(2), text("pad"), bool(true))).toEqual({
      kind: "array",
      rows: 2,
      cols: 3,
      values: [num(1), num(3), num(5), num(2), num(4), text("pad")],
    });

    expect(WRAPROWS(vector, num(0))).toEqual(err(ErrorCode.Value));
    expect(WRAPROWS(vector, num(2), cellRange([num(1)], 1, 1))).toEqual(err(ErrorCode.Value));
    expect(WRAPROWS(vector, err(ErrorCode.Name))).toEqual(err(ErrorCode.Name));
    expect(WRAPROWS(vector, num(2), err(ErrorCode.NA))).toEqual(err(ErrorCode.NA));
    expect(WRAPROWS(vector, num(2), text("pad"), err(ErrorCode.Ref))).toEqual(err(ErrorCode.Ref));
    expect(WRAPROWS(rowRefRange, num(2))).toEqual(err(ErrorCode.Value));
    expect(WRAPROWS(vector, num(2), text("pad"), text("bad"))).toEqual(err(ErrorCode.Value));
    expect(WRAPCOLS(vector, err(ErrorCode.Name))).toEqual(err(ErrorCode.Name));
    expect(WRAPCOLS(vector, num(2), err(ErrorCode.NA))).toEqual(err(ErrorCode.NA));
    expect(WRAPCOLS(vector, num(2), text("pad"), err(ErrorCode.Ref))).toEqual(err(ErrorCode.Ref));
    expect(WRAPCOLS(vector, num(0))).toEqual(err(ErrorCode.Value));
    expect(WRAPCOLS(rowRefRange, num(2))).toEqual(err(ErrorCode.Value));
    expect(WRAPCOLS(vector, cellRange([num(1)], 1, 1), text("pad"))).toEqual(err(ErrorCode.Value));
    expect(WRAPCOLS(vector, num(2), text("pad"), text("bad"))).toEqual(err(ErrorCode.Value));
  });

  it("covers remaining SUMPRODUCT validation branches", () => {
    const SUMPRODUCT = getLookupBuiltin("SUMPRODUCT")!;

    expect(SUMPRODUCT()).toEqual(err(ErrorCode.Value));
    expect(SUMPRODUCT(num(1), cellRange([num(2)], 1, 1))).toEqual(err(ErrorCode.Value));
    expect(SUMPRODUCT(cellRange([num(1), num(2)], 2, 1), cellRange([num(3)], 1, 1))).toEqual(
      err(ErrorCode.Value),
    );
  });

  it("covers COUNTIFS and SUMIFS error branches", () => {
    const COUNTIFS = getLookupBuiltin("COUNTIFS")!;
    const SUMIFS = getLookupBuiltin("SUMIFS")!;

    expect(COUNTIFS(cellRange([num(1), num(2)], 2, 1))).toEqual(err(ErrorCode.Value));
    expect(
      COUNTIFS(cellRange([num(1), num(2)], 2, 1), text(">1"), cellRange([num(2)], 1, 1), text("2")),
    ).toEqual(err(ErrorCode.Value));
    expect(
      SUMIFS(
        cellRange([num(1), num(2)], 2, 1),
        cellRange([num(1), num(2)], 2, 1),
        text(">1"),
        cellRange([num(3)], 1, 1),
        text("3"),
      ),
    ).toEqual(err(ErrorCode.Value));
  });

  it("covers remaining database passthrough and SUMPRODUCT scalar validation branches", () => {
    const DSTDEVP = getLookupBuiltin("DSTDEVP")!;
    const DSUM = getLookupBuiltin("DSUM")!;
    const SUMPRODUCT = getLookupBuiltin("SUMPRODUCT")!;

    const database = cellRange(
      [text("Age"), text("Yield"), num(10), num(5), num(12), num(7)],
      3,
      2,
    );

    expect(
      DSTDEVP(database, text("Yield"), cellRange([text("Age"), err(ErrorCode.Ref)], 2, 1)),
    ).toEqual(err(ErrorCode.Ref));
    expect(
      DSUM(database, text("Yield"), cellRange([text("Age"), err(ErrorCode.Name)], 2, 1)),
    ).toEqual(err(ErrorCode.Name));
    expect(SUMPRODUCT(num(1), num(2))).toEqual(err(ErrorCode.Value));
  });

  it("ignores blocked database criteria rows with missing or blank headers", () => {
    const DCOUNT = getLookupBuiltin("DCOUNT")!;
    const DSUM = getLookupBuiltin("DSUM")!;

    const database = cellRange(
      [text("Age"), text("Yield"), num(10), num(5), num(12), num(7), num(12), num(9)],
      4,
      2,
    );

    expect(
      DCOUNT(
        database,
        text("Yield"),
        cellRange([text("Age"), text("Missing"), num(12), num(1)], 2, 2),
      ),
    ).toEqual(num(0));
    expect(
      DSUM(
        database,
        text("Yield"),
        cellRange([{ tag: ValueTag.Empty }, text("Yield"), num(1), num(7)], 2, 2),
      ),
    ).toEqual(num(0));
  });

  it("covers UNIQUE by-column/row modes with duplicate and error branches", () => {
    const UNIQUE = getLookupBuiltin("UNIQUE")!;

    expect(
      UNIQUE(
        cellRange([num(1), num(1), num(2), num(1), num(1), num(2)], 2, 3),
        bool(true),
        bool(true),
      ),
    ).toEqual({
      kind: "array",
      rows: 2,
      cols: 1,
      values: [num(2), num(2)],
    });

    expect(UNIQUE(cellRange([num(1), num(2), num(3), num(4)], 2, 2), bool(true))).toEqual({
      kind: "array",
      rows: 2,
      cols: 2,
      values: [num(1), num(2), num(3), num(4)],
    });

    expect(
      UNIQUE(cellRange([num(1), err(ErrorCode.Ref), num(3), num(4)], 2, 2), bool(true)),
    ).toEqual(err(ErrorCode.Value));

    expect(UNIQUE(cellRange([num(1), num(2), err(ErrorCode.Name), num(4)], 2, 2))).toEqual(
      err(ErrorCode.Value),
    );
  });

  it("covers criteria matching with error values and invalid operand types", () => {
    const COUNTIF = getLookupBuiltin("COUNTIF")!;

    expect(COUNTIF(cellRange([err(ErrorCode.Ref), num(1), num(2)], 3, 1), text(">0"))).toEqual(
      num(2),
    );
    expect(COUNTIF(cellRange([num(1), num(2), num(3)], 3, 1), text(">=a"))).toEqual(num(0));
    expect(COUNTIF(cellRange([num(3), num(1), num(4)], 3, 1), text("<2"))).toEqual(num(1));
  });

  it("covers boundary behavior for lookup reshaping helpers", () => {
    const OFFSET = getLookupBuiltin("OFFSET")!;
    const TAKE = getLookupBuiltin("TAKE")!;
    const DROP = getLookupBuiltin("DROP")!;
    const CHOOSECOLS = getLookupBuiltin("CHOOSECOLS")!;
    const CHOOSEROWS = getLookupBuiltin("CHOOSEROWS")!;

    const matrix = cellRange([num(1), num(2), num(3), num(4)], 2, 2);
    const column = cellRange([num(1), num(2), num(3)], 3, 1);

    expect(OFFSET(matrix, num(0), num(0))).toEqual({
      kind: "array",
      rows: 2,
      cols: 2,
      values: matrix.values,
    });
    expect(OFFSET(matrix, num(0), num(0), num(1), num(1), num(2))).toEqual(err(ErrorCode.Value));

    expect(TAKE(column)).toEqual({
      kind: "array",
      rows: 3,
      cols: 1,
      values: column.values,
    });
    expect(TAKE(column, num(0))).toEqual(err(ErrorCode.Value));

    expect(DROP(matrix, num(0), num(0))).toEqual({
      kind: "array",
      rows: 2,
      cols: 2,
      values: matrix.values,
    });
    expect(DROP(matrix, num(5))).toEqual(err(ErrorCode.Value));

    expect(CHOOSECOLS(matrix, num(3))).toEqual(err(ErrorCode.Value));
    expect(CHOOSEROWS(matrix, num(3))).toEqual(err(ErrorCode.Value));
  });

  it("supports FILTER and UNIQUE dynamic-array results", () => {
    const FILTER = getLookupBuiltin("FILTER")!;
    const UNIQUE = getLookupBuiltin("UNIQUE")!;

    const filtered = FILTER(
      cellRange([num(1), num(3), num(2), num(4)], 4, 1),
      cellRange([bool(false), bool(true), bool(false), bool(true)], 4, 1),
    );
    expect(filtered).toEqual({
      kind: "array",
      rows: 2,
      cols: 1,
      values: [num(3), num(4)],
    });

    const unique = UNIQUE(cellRange([text("A"), text("B"), text("A"), text("C")], 4, 1));
    expect(unique).toEqual({
      kind: "array",
      rows: 3,
      cols: 1,
      values: [text("A"), text("B"), text("C")],
    });
  });

  it("covers FILTER fallbacks and UNIQUE row and column modes", () => {
    const FILTER = getLookupBuiltin("FILTER")!;
    const UNIQUE = getLookupBuiltin("UNIQUE")!;

    expect(
      FILTER(
        cellRange([num(1), num(2), num(3), num(4), num(5), num(6)], 2, 3),
        cellRange([bool(true), bool(false), bool(true)], 1, 3),
      ),
    ).toEqual({
      kind: "array",
      rows: 2,
      cols: 2,
      values: [num(1), num(3), num(4), num(6)],
    });

    expect(
      FILTER(
        cellRange([num(1), num(2)], 2, 1),
        cellRange([bool(false), bool(false)], 2, 1),
        text("empty"),
      ),
    ).toEqual(text("empty"));

    expect(
      FILTER(cellRange([num(1), num(2)], 2, 1), cellRange([text("bad"), bool(true)], 2, 1)),
    ).toEqual(err(ErrorCode.Value));

    expect(
      FILTER(
        cellRange([num(1), num(2)], 2, 1),
        cellRange([bool(true), bool(false), bool(true)], 3, 1),
      ),
    ).toEqual(err(ErrorCode.Value));

    expect(
      FILTER(cellRange([num(1), num(2)], 2, 1), cellRange([err(ErrorCode.Ref), bool(true)], 2, 1)),
    ).toEqual(err(ErrorCode.Ref));

    expect(
      UNIQUE(
        cellRange([text("A"), text("b"), text("a"), text("C"), text("c")], 5, 1),
        bool(false),
        bool(true),
      ),
    ).toEqual({
      kind: "array",
      rows: 1,
      cols: 1,
      values: [text("b")],
    });

    expect(
      UNIQUE(
        cellRange(
          [text("A"), text("B"), text("A"), text("C"), num(1), num(2), num(1), num(3)],
          2,
          4,
        ),
        bool(true),
      ),
    ).toEqual({
      kind: "array",
      rows: 2,
      cols: 3,
      values: [text("A"), text("B"), text("C"), num(1), num(2), num(3)],
    });

    expect(
      UNIQUE(
        cellRange(
          [text("A"), text("B"), text("A"), text("C"), num(1), num(2), num(1), num(3)],
          2,
          4,
        ),
        bool(true),
        bool(true),
      ),
    ).toEqual({
      kind: "array",
      rows: 2,
      cols: 2,
      values: [text("B"), text("C"), num(2), num(3)],
    });

    expect(
      UNIQUE(cellRange([text("A"), num(1), text("a"), num(1), text("B"), num(2)], 3, 2)),
    ).toEqual({
      kind: "array",
      rows: 2,
      cols: 2,
      values: [text("A"), num(1), text("B"), num(2)],
    });

    expect(UNIQUE(cellRange([err(ErrorCode.Name)], 1, 1))).toEqual(err(ErrorCode.Name));
    expect(UNIQUE(cellRange([num(1), num(2)], 2, 1), text("bad"))).toEqual(err(ErrorCode.Value));
  });

  it("covers FILTER horizontal validation and UNIQUE argument validation branches", () => {
    const FILTER = getLookupBuiltin("FILTER")!;
    const UNIQUE = getLookupBuiltin("UNIQUE")!;

    expect(
      FILTER(
        cellRange([num(1), num(2), num(3), num(4)], 2, 2),
        cellRange([err(ErrorCode.Ref), bool(true)], 1, 2),
      ),
    ).toEqual(err(ErrorCode.Ref));
    expect(
      FILTER(
        cellRange([num(1), num(2), num(3), num(4)], 2, 2),
        cellRange([text("bad"), bool(true)], 1, 2),
      ),
    ).toEqual(err(ErrorCode.Value));
    expect(
      FILTER(
        cellRange([num(1), num(2), num(3), num(4)], 2, 2),
        cellRange([bool(false), bool(false)], 1, 2),
        cellRange([num(0)], 1, 1),
      ),
    ).toEqual(err(ErrorCode.Value));

    expect(UNIQUE(err(ErrorCode.Name))).toEqual(err(ErrorCode.Value));
    expect(UNIQUE(cellRange([num(1), num(2)], 2, 1), cellRange([bool(true)], 1, 1))).toEqual(
      err(ErrorCode.Value),
    );
    expect(UNIQUE(cellRange([num(1), num(2)], 2, 1), err(ErrorCode.Ref))).toEqual(
      err(ErrorCode.Ref),
    );
    expect(UNIQUE(cellRange([num(1), num(2)], 2, 1), bool(false), err(ErrorCode.NA))).toEqual(
      err(ErrorCode.NA),
    );
    expect(
      UNIQUE(
        cellRange([text("A"), num(1), text("A"), num(1), text("B"), num(2)], 3, 2),
        bool(false),
        bool(true),
      ),
    ).toEqual({
      kind: "array",
      rows: 1,
      cols: 2,
      values: [text("B"), num(2)],
    });
  });

  it("supports matrix and extended numeric lookup builtins", () => {
    const SUMX2MY2 = getLookupBuiltin("SUMX2MY2")!;
    const SUMX2PY2 = getLookupBuiltin("SUMX2PY2")!;
    const SUMXMY2 = getLookupBuiltin("SUMXMY2")!;
    const MDETERM = getLookupBuiltin("MDETERM")!;
    const MINVERSE = getLookupBuiltin("MINVERSE")!;
    const MMULT = getLookupBuiltin("MMULT")!;
    const PERCENTOF = getLookupBuiltin("PERCENTOF")!;

    expect(SUMX2MY2(cellRange([num(2), num(3)], 2, 1), cellRange([num(1), num(1)], 2, 1))).toEqual(
      num(11),
    );
    expect(SUMX2PY2(cellRange([num(2), num(3)], 2, 1), cellRange([num(1), num(1)], 2, 1))).toEqual(
      num(15),
    );
    expect(SUMXMY2(cellRange([num(2), num(3)], 2, 1), cellRange([num(1), num(1)], 2, 1))).toEqual(
      num(5),
    );

    expect(MDETERM(cellRange([num(1), num(2), num(3), num(4)], 2, 2))).toEqual(num(-2));
    expect(MDETERM(cellRange([num(1), num(2), num(3)], 3, 1))).toEqual(err(ErrorCode.Value));

    const inverse = MINVERSE(cellRange([num(4), num(7), num(2), num(6)], 2, 2));
    expect(inverse).toMatchObject({ kind: "array", rows: 2, cols: 2 });
    if (!(inverse && "kind" in inverse && inverse.kind === "array")) {
      throw new Error("expected MINVERSE to return an array");
    }
    expect(inverse.values.map((value) => value.tag)).toEqual([
      ValueTag.Number,
      ValueTag.Number,
      ValueTag.Number,
      ValueTag.Number,
    ]);
    expect(inverse.values.map((value) => value.value)).toEqual([
      expect.closeTo(0.6, 12),
      expect.closeTo(-0.7, 12),
      expect.closeTo(-0.2, 12),
      expect.closeTo(0.4, 12),
    ]);
    expect(MINVERSE(cellRange([num(1), num(2), num(2), num(4)], 2, 2))).toEqual(
      err(ErrorCode.Value),
    );

    expect(
      MMULT(
        cellRange([num(1), num(2), num(3), num(4)], 2, 2),
        cellRange([num(5), num(6), num(7), num(8)], 2, 2),
      ),
    ).toEqual({
      kind: "array",
      rows: 2,
      cols: 2,
      values: [num(19), num(22), num(43), num(50)],
    });
    expect(
      MMULT(
        cellRange([num(1), num(2), num(3), num(4)], 2, 2),
        cellRange([num(5), num(6), num(7)], 3, 1),
      ),
    ).toEqual(err(ErrorCode.Value));

    expect(
      PERCENTOF(cellRange([num(2), num(3)], 2, 1), cellRange([num(10), num(10)], 2, 1)),
    ).toEqual(num(0.25));
    expect(PERCENTOF(cellRange([num(2)], 1, 1), cellRange([num(0)], 1, 1))).toEqual(
      err(ErrorCode.Div0),
    );
    expect(SUMXMY2(err(ErrorCode.Ref), cellRange([num(1)], 1, 1))).toEqual(err(ErrorCode.Value));
  });

  it("covers matrix helper validation and percent-of error branches", () => {
    const SUMX2PY2 = getLookupBuiltin("SUMX2PY2")!;
    const MDETERM = getLookupBuiltin("MDETERM")!;
    const MINVERSE = getLookupBuiltin("MINVERSE")!;
    const MMULT = getLookupBuiltin("MMULT")!;
    const PERCENTOF = getLookupBuiltin("PERCENTOF")!;

    expect(MINVERSE(err(ErrorCode.Ref))).toEqual(err(ErrorCode.Value));
    expect(MINVERSE(cellRange([num(1), num(2), num(3), num(4), num(5), num(6)], 2, 3))).toEqual(
      err(ErrorCode.Value),
    );

    expect(MMULT(err(ErrorCode.Name), cellRange([num(1)], 1, 1))).toEqual(err(ErrorCode.Value));
    expect(MMULT(cellRange([num(1)], 1, 1), err(ErrorCode.Ref))).toEqual(err(ErrorCode.Value));
    expect(MMULT(cellRange([num(1), num(2)], 1, 2), cellRange([num(1), num(2)], 1, 2))).toEqual(
      err(ErrorCode.Value),
    );

    expect(PERCENTOF(err(ErrorCode.Name), cellRange([num(10)], 1, 1))).toEqual(
      err(ErrorCode.Value),
    );
    expect(PERCENTOF(cellRange([num(1)], 1, 1), err(ErrorCode.Ref))).toEqual(err(ErrorCode.Value));
    expect(SUMX2PY2(err(ErrorCode.Ref), cellRange([num(1)], 1, 1))).toEqual(err(ErrorCode.Value));
    expect(SUMX2PY2(cellRange([num(1)], 1, 1), err(ErrorCode.Name))).toEqual(err(ErrorCode.Value));
    expect(SUMX2PY2(cellRange([num(1)], 1, 1), cellRange([num(1), num(2)], 2, 1))).toEqual(
      err(ErrorCode.Value),
    );
    expect(MDETERM(err(ErrorCode.Ref))).toEqual(err(ErrorCode.Value));
  });

  it("covers conditional criteria parsing variants", () => {
    const COUNTIF = getLookupBuiltin("COUNTIF")!;
    const SUMIF = getLookupBuiltin("SUMIF")!;

    expect(COUNTIF(cellRange([num(1), num(2), num(3)], 3, 1), num(2))).toEqual(num(1));
    expect(COUNTIF(cellRange([num(1), num(2), num(3)], 3, 1), text("<>2"))).toEqual(num(2));
    expect(COUNTIF(cellRange([num(1), num(2), num(3)], 3, 1), text(">=2"))).toEqual(num(2));
    expect(COUNTIF(cellRange([num(1), num(2), num(3)], 3, 1), text("<=2"))).toEqual(num(2));
    expect(COUNTIF(cellRange([bool(true), bool(false), bool(true)], 3, 1), text("=TRUE"))).toEqual(
      num(2),
    );
    expect(COUNTIF(cellRange([text(""), text("x"), text("")], 3, 1), text("="))).toEqual(num(2));
    expect(
      SUMIF(
        cellRange([text("a"), text("b"), text("c")], 3, 1),
        text("<>b"),
        cellRange([num(1), num(2), num(3)], 3, 1),
      ),
    ).toEqual(num(4));
  });

  it("covers FILTER column selection and UNIQUE row and column de-duplication branches", () => {
    const FILTER = getLookupBuiltin("FILTER")!;
    const UNIQUE = getLookupBuiltin("UNIQUE")!;

    expect(
      FILTER(
        cellRange([text("A"), text("B"), text("C"), text("D")], 2, 2),
        cellRange([bool(true), bool(false)], 1, 2),
      ),
    ).toEqual({
      kind: "array",
      rows: 2,
      cols: 1,
      values: [text("A"), text("C")],
    });
    expect(
      FILTER(
        cellRange([text("A"), text("B"), text("C"), text("D")], 2, 2),
        cellRange([text("bad"), bool(true)], 1, 2),
      ),
    ).toEqual(err(ErrorCode.Value));
    expect(
      FILTER(
        cellRange([text("A"), text("B"), text("C"), text("D")], 2, 2),
        cellRange([bool(false), bool(false)], 1, 2),
        text("empty"),
      ),
    ).toEqual(text("empty"));
    expect(
      FILTER(
        cellRange([text("A"), text("B"), text("C"), text("D")], 2, 2),
        cellRange([bool(false), bool(false)], 1, 2),
        cellRange([text("x")], 1, 1),
      ),
    ).toEqual(err(ErrorCode.Value));
    expect(FILTER(err(ErrorCode.Ref), cellRange([bool(true)], 1, 1))).toEqual(err(ErrorCode.Value));
    expect(FILTER(cellRange([text("A")], 1, 1), err(ErrorCode.Name))).toEqual(err(ErrorCode.Value));

    expect(
      UNIQUE(
        cellRange(
          [text("A"), text("A"), text("B"), text("C"), num(1), num(1), num(2), num(3)],
          2,
          4,
        ),
        bool(true),
        bool(true),
      ),
    ).toEqual({
      kind: "array",
      rows: 2,
      cols: 2,
      values: [text("B"), text("C"), num(2), num(3)],
    });
    expect(
      UNIQUE(
        cellRange([text("A"), num(1), text("A"), num(1), text("B"), num(2)], 3, 2),
        bool(false),
        bool(false),
      ),
    ).toEqual({
      kind: "array",
      rows: 2,
      cols: 2,
      values: [text("A"), num(1), text("B"), num(2)],
    });
  });

  it("covers SUMPRODUCT and the remaining matrix helper validation paths", () => {
    const SUMPRODUCT = getLookupBuiltin("SUMPRODUCT")!;
    const SUMX2MY2 = getLookupBuiltin("SUMX2MY2")!;
    const SUMXMY2 = getLookupBuiltin("SUMXMY2")!;
    const MDETERM = getLookupBuiltin("MDETERM")!;

    expect(SUMPRODUCT()).toEqual(err(ErrorCode.Value));
    expect(SUMPRODUCT(err(ErrorCode.Ref), cellRange([num(1)], 1, 1))).toEqual(err(ErrorCode.Value));
    expect(
      SUMPRODUCT(cellRange([num(2), num(3)], 2, 1), cellRange([num(4), num(5)], 2, 1)),
    ).toEqual(num(23));
    expect(
      SUMPRODUCT(cellRange([num(2), num(3)], 2, 1), cellRange([num(4), num(5), num(6)], 3, 1)),
    ).toEqual(err(ErrorCode.Value));

    expect(SUMX2MY2(err(ErrorCode.Ref), cellRange([num(1)], 1, 1))).toEqual(err(ErrorCode.Value));
    expect(SUMX2MY2(cellRange([num(1)], 1, 1), err(ErrorCode.Name))).toEqual(err(ErrorCode.Value));
    expect(SUMX2MY2(cellRange([num(1)], 1, 1), cellRange([num(1), num(2)], 2, 1))).toEqual(
      err(ErrorCode.Value),
    );
    expect(SUMXMY2(cellRange([num(1)], 1, 1), err(ErrorCode.Ref))).toEqual(err(ErrorCode.Value));
    expect(SUMXMY2(cellRange([num(1)], 1, 1), cellRange([num(1), num(2)], 2, 1))).toEqual(
      err(ErrorCode.Value),
    );

    expect(MDETERM(cellRange([], 0, 0))).toEqual(err(ErrorCode.Value));
  });

  it("covers AVERAGEIFS, MINIFS, and MAXIFS validation and zero-match branches", () => {
    const AVERAGEIFS = getLookupBuiltin("AVERAGEIFS")!;
    const MINIFS = getLookupBuiltin("MINIFS")!;
    const MAXIFS = getLookupBuiltin("MAXIFS")!;

    expect(AVERAGEIFS(err(ErrorCode.Ref), cellRange([num(1)], 1, 1), text(">0"))).toEqual(
      err(ErrorCode.Value),
    );
    expect(
      AVERAGEIFS(cellRange([text("skip")], 1, 1), cellRange([num(1)], 1, 1), text(">0")),
    ).toEqual(err(ErrorCode.Div0));
    expect(MINIFS(err(ErrorCode.Ref), cellRange([num(1)], 1, 1), text(">0"))).toEqual(
      err(ErrorCode.Value),
    );
    expect(MINIFS(cellRange([num(1)], 1, 1), err(ErrorCode.Name), text(">0"))).toEqual(
      err(ErrorCode.Value),
    );
    expect(MAXIFS(err(ErrorCode.Ref), cellRange([num(1)], 1, 1), text(">0"))).toEqual(
      err(ErrorCode.Value),
    );
    expect(MAXIFS(cellRange([num(1)], 1, 1), err(ErrorCode.Name), text(">0"))).toEqual(
      err(ErrorCode.Value),
    );
  });

  it("covers remaining matrix, sort, and criteria edge cases", () => {
    const MINVERSE = getLookupBuiltin("MINVERSE")!;
    const SORT = getLookupBuiltin("SORT")!;
    const UNIQUE = getLookupBuiltin("UNIQUE")!;
    const HSTACK = getLookupBuiltin("HSTACK")!;
    const VSTACK = getLookupBuiltin("VSTACK")!;
    const COUNTIF = getLookupBuiltin("COUNTIF")!;
    const TOCOL = getLookupBuiltin("TOCOL")!;
    const TOROW = getLookupBuiltin("TOROW")!;

    // Singular matrix for MINVERSE
    expect(MINVERSE(cellRange([num(1), num(2), num(2), num(4)], 2, 2))).toEqual(
      err(ErrorCode.Value),
    );

    // SORT by column
    const matrix = cellRange([num(3), num(1), num(4), num(2)], 2, 2);
    expect(SORT(matrix, num(1), num(1), bool(true))).toEqual({
      kind: "array",
      rows: 2,
      cols: 2,
      values: [num(1), num(3), num(2), num(4)],
    });

    // UNIQUE exactlyOnce on 2D
    const matrix2 = cellRange([num(1), num(2), num(1), num(2), num(3), num(4)], 3, 2);
    expect(UNIQUE(matrix2, bool(false), bool(true))).toEqual({
      kind: "array",
      rows: 1,
      cols: 2,
      values: [num(3), num(4)],
    });

    // HSTACK/VSTACK single row/col expansion
    expect(HSTACK(cellRange([num(1)], 1, 1), cellRange([num(2), num(3)], 2, 1))).toEqual({
      kind: "array",
      rows: 2,
      cols: 2,
      values: [num(1), num(2), num(1), num(3)],
    });
    expect(VSTACK(cellRange([num(1)], 1, 1), cellRange([num(2), num(3)], 1, 2))).toEqual({
      kind: "array",
      rows: 2,
      cols: 2,
      values: [num(1), num(1), num(2), num(3)],
    });
    expect(VSTACK()).toEqual(err(ErrorCode.Value));
    expect(
      VSTACK({
        kind: "range",
        refKind: "rows",
        rows: 1,
        cols: 1,
        values: [num(1)],
      }),
    ).toEqual(err(ErrorCode.Value));
    expect(
      VSTACK(
        cellRange([num(1), num(2), num(3), num(4)], 2, 2),
        cellRange([num(5), num(6), num(7)], 1, 3),
      ),
    ).toEqual(err(ErrorCode.Value));

    // matchesCriteria operators
    const range = cellRange([num(1), num(2), num(3), num(4)], 4, 1);
    expect(COUNTIF(range, text("<>2"))).toEqual(num(3));
    expect(COUNTIF(range, text("<=2"))).toEqual(num(2));
    expect(COUNTIF(range, text(">=3"))).toEqual(num(2));
    expect(COUNTIF(range, text("<3"))).toEqual(num(2));

    // TOCOL/TOROW ignoreEmpty
    const sparse = cellRange([num(1), { tag: ValueTag.Empty }, num(2)], 3, 1);
    expect(TOCOL(sparse, num(1))).toEqual({
      kind: "array",
      rows: 2,
      cols: 1,
      values: [num(1), num(2)],
    });
    expect(TOROW(sparse, num(1))).toEqual({
      kind: "array",
      rows: 1,
      cols: 2,
      values: [num(1), num(2)],
    });
  });

  it("supports native cash-flow rate helpers on numeric ranges", () => {
    const IRR = getLookupBuiltin("IRR")!;
    const MIRR = getLookupBuiltin("MIRR")!;
    const XNPV = getLookupBuiltin("XNPV")!;
    const XIRR = getLookupBuiltin("XIRR")!;

    const irrValues = cellRange(
      [num(-70000), num(12000), num(15000), num(18000), num(21000), num(26000)],
      6,
      1,
    );
    const mirrValues = cellRange(
      [num(-120000), num(39000), num(30000), num(21000), num(37000), num(46000)],
      6,
      1,
    );
    const xValues = cellRange([num(-10000), num(2750), num(4250), num(3250), num(2750)], 5, 1);
    const xDates = cellRange([num(39448), num(39508), num(39751), num(39859), num(39904)], 5, 1);

    const irr = IRR(irrValues);
    if (irr.tag !== ValueTag.Number) throw new Error(`Expected number result, received ${irr.tag}`);
    expect(irr.value).toBeCloseTo(0.08663094803653162, 12);

    const mirr = MIRR(mirrValues, num(0.1), num(0.12));
    if (mirr.tag !== ValueTag.Number)
      throw new Error(`Expected number result, received ${mirr.tag}`);
    expect(mirr.value).toBeCloseTo(0.1260941303659051, 12);

    const xnpv = XNPV(num(0.09), xValues, xDates);
    if (xnpv.tag !== ValueTag.Number)
      throw new Error(`Expected number result, received ${xnpv.tag}`);
    expect(xnpv.value).toBeCloseTo(2086.647602031535, 9);

    const xirr = XIRR(xValues, xDates);
    if (xirr.tag !== ValueTag.Number)
      throw new Error(`Expected number result, received ${xirr.tag}`);
    expect(xirr.value).toBeCloseTo(0.37336253351883136, 12);

    expect(IRR(cellRange([num(5), num(7)], 2, 1))).toEqual(err(ErrorCode.Value));
    expect(MIRR(cellRange([num(5), num(7)], 2, 1), num(0.1), num(0.12))).toEqual(
      err(ErrorCode.Div0),
    );
    expect(XNPV(num(0.09), xValues, cellRange([num(39448), num(39508)], 2, 1))).toEqual(
      err(ErrorCode.Value),
    );
    expect(
      XNPV(
        num(0.09),
        xValues,
        cellRange([num(39448), num(39508), num(39400), num(39859), num(39904)], 5, 1),
      ),
    ).toEqual(err(ErrorCode.Value));
    expect(
      XIRR(xValues, cellRange([num(39448), num(39508), num(39400), num(39859), num(39904)], 5, 1)),
    ).toEqual(err(ErrorCode.Value));
    expect(XIRR(xValues, xDates, text("bad"))).toEqual(err(ErrorCode.Value));
  });

  it("covers remaining lookup, database, and cash-flow validation branches", () => {
    const AREAS = getLookupBuiltin("AREAS")!;
    const ARRAYTOTEXT = getLookupBuiltin("ARRAYTOTEXT")!;
    const ROWS = getLookupBuiltin("ROWS")!;
    const COLUMNS = getLookupBuiltin("COLUMNS")!;
    const DCOUNT = getLookupBuiltin("DCOUNT")!;
    const IRR = getLookupBuiltin("IRR")!;
    const MIRR = getLookupBuiltin("MIRR")!;
    const XNPV = getLookupBuiltin("XNPV")!;
    const XIRR = getLookupBuiltin("XIRR")!;
    const XLOOKUP = getLookupBuiltin("XLOOKUP")!;
    const XMATCH = getLookupBuiltin("XMATCH")!;

    const database = cellRange(
      [
        text("Age"),
        text("Height"),
        text("Yield"),
        num(10),
        num(100),
        num(5),
        num(12),
        num(110),
        num(7),
        num(12),
        num(120),
        num(9),
      ],
      4,
      3,
    );
    const ageCriteria = cellRange([text("Age"), num(12)], 2, 1);
    const ageCriteriaWithBlankClause = cellRange(
      [text("Age"), text("Yield"), num(12), { tag: ValueTag.Empty }],
      2,
      2,
    );

    expect(DCOUNT(num(1), text("Yield"), ageCriteria)).toEqual(err(ErrorCode.Value));
    expect(DCOUNT(database, text("Yield"), num(1))).toEqual(err(ErrorCode.Value));
    expect(DCOUNT(cellRange([], 0, 0), text("Yield"), ageCriteria)).toEqual(err(ErrorCode.Value));
    expect(DCOUNT(database, text("Yield"), ageCriteriaWithBlankClause)).toEqual(num(2));

    expect(AREAS(num(1))).toEqual(err(ErrorCode.Value));
    expect(ROWS(num(1))).toEqual(err(ErrorCode.Value));
    expect(COLUMNS(num(1))).toEqual(err(ErrorCode.Value));
    expect(ARRAYTOTEXT(cellRange([err(ErrorCode.Ref)], 1, 1))).toEqual(err(ErrorCode.Value));

    const xValues = cellRange([num(-10000), num(2750), num(4250), num(3250), num(2750)], 5, 1);
    const xDates = cellRange([num(39448), num(39508), num(39751), num(39859), num(39904)], 5, 1);
    const mirrValues = cellRange(
      [num(-120000), num(39000), num(30000), num(21000), num(37000), num(46000)],
      6,
      1,
    );

    expect(IRR(cellRange([err(ErrorCode.Name), num(1)], 2, 1))).toEqual(err(ErrorCode.Name));
    expect(
      MIRR(
        { kind: "range", refKind: "rows", rows: 1, cols: 2, values: [num(-1), num(2)] },
        num(0.1),
        num(0.12),
      ),
    ).toEqual(err(ErrorCode.Value));
    expect(MIRR(mirrValues, num(-1), num(0.12))).toEqual(err(ErrorCode.Div0));
    expect(XNPV(err(ErrorCode.Name), xValues, xDates)).toEqual(err(ErrorCode.Name));
    expect(
      XNPV(
        num(0.09),
        { kind: "range", refKind: "rows", rows: 1, cols: 2, values: [num(-1), num(2)] },
        cellRange([num(39448), num(39508)], 2, 1),
      ),
    ).toEqual(err(ErrorCode.Value));
    expect(
      XNPV(
        num(0.09),
        cellRange([err(ErrorCode.Ref), num(2)], 2, 1),
        cellRange([num(39448), num(39508)], 2, 1),
      ),
    ).toEqual(err(ErrorCode.Ref));
    expect(
      XNPV(
        num(0.09),
        cellRange([num(-1), num(2)], 2, 1),
        cellRange([num(39448), { tag: ValueTag.Number, value: Number.POSITIVE_INFINITY }], 2, 1),
      ),
    ).toEqual(err(ErrorCode.Value));
    expect(XIRR(xValues, xDates, cellRange([num(0.1)], 1, 1))).toEqual(err(ErrorCode.Value));

    const duplicateLookup = cellRange([text("pear"), text("apple"), text("pear")], 3, 1);
    const duplicateReturn = cellRange([num(10), num(20), num(30)], 3, 1);

    expect(
      XLOOKUP(text("pear"), duplicateLookup, duplicateReturn, text("fallback"), num(0), num(-1)),
    ).toEqual(num(30));
    expect(
      XLOOKUP(text("pear"), duplicateLookup, duplicateReturn, text("fallback"), num(1), num(1)),
    ).toEqual(err(ErrorCode.Value));
    expect(XMATCH(text("pear"), duplicateLookup, num(0), num(-1))).toEqual(num(3));
    expect(XMATCH(text("pear"), duplicateLookup, num(2), num(1))).toEqual(err(ErrorCode.Value));
  });
});
