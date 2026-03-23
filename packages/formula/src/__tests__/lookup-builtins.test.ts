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
