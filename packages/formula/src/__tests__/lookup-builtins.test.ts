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
      values: [num(1), num(3), num(4), num(2)],
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

    expect(TOCOL(num(1))).toEqual({
      kind: "array",
      rows: 1,
      cols: 1,
      values: [num(1)],
    });
    expect(TOROW(num(1))).toEqual({
      kind: "array",
      rows: 1,
      cols: 1,
      values: [num(1)],
    });
    expect(
      TOROW(cellRange([num(1), num(2), num(3), num(4)], 2, 2), cellRange([num(1)], 1, 1)),
    ).toEqual(err(ErrorCode.Value));
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
    expect(WRAPROWS(vector, num(2), text("pad"), text("bad"))).toEqual(err(ErrorCode.Value));
    expect(WRAPCOLS(vector, num(0))).toEqual(err(ErrorCode.Value));
    expect(WRAPCOLS(vector, cellRange([num(1)], 1, 1), text("pad"))).toEqual(err(ErrorCode.Value));
    expect(WRAPCOLS(vector, num(2), text("pad"), text("bad"))).toEqual(err(ErrorCode.Value));
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
});
