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

  it("supports MEDIAN, SMALL, LARGE, RANK, and RANK.EQ", () => {
    const MEDIAN = getLookupBuiltin("MEDIAN")!;
    const SMALL = getLookupBuiltin("SMALL")!;
    const LARGE = getLookupBuiltin("LARGE")!;
    const RANK = getLookupBuiltin("RANK")!;
    const RANKEQ = getLookupBuiltin("RANK.EQ")!;

    const sample = cellRange([num(1), num(4), num(2), num(4), num(3)], 5, 1);

    expect(MEDIAN(sample)).toEqual(num(3));
    expect(MEDIAN(num(7))).toEqual(num(7));
    expect(MEDIAN(cellRange([num(1), num(2), num(3), num(4)], 2, 2))).toEqual(num(2.5));
    expect(MEDIAN(cellRange([text("bad"), num(1)], 1, 2))).toEqual(err(ErrorCode.Value));

    expect(SMALL(sample, num(1))).toEqual(num(1));
    expect(SMALL(sample, num(4))).toEqual(num(4));
    expect(LARGE(sample, num(1))).toEqual(num(4));
    expect(LARGE(sample, num(4))).toEqual(num(2));
    expect(SMALL(sample, num(0))).toEqual(err(ErrorCode.Value));
    expect(LARGE(sample, num(6))).toEqual(err(ErrorCode.Value));

    expect(RANK(num(4), sample)).toEqual(num(1));
    expect(RANK(num(1), sample)).toEqual(num(5));
    expect(RANK(num(3), sample, num(1))).toEqual(num(3));
    expect(RANKEQ(num(4), sample)).toEqual(num(1));
    expect(RANKEQ(num(8), sample)).toEqual(err(ErrorCode.NA));
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
});
