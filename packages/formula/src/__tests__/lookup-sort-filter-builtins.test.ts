import { describe, expect, it } from "vitest";
import { ErrorCode, ValueTag, type CellValue } from "@bilig/protocol";
import { getLookupBuiltin, type RangeBuiltinArgument } from "../builtins/lookup.js";

const num = (value: number): CellValue => ({ tag: ValueTag.Number, value });
const bool = (value: boolean): CellValue => ({ tag: ValueTag.Boolean, value });
const text = (value: string): CellValue => ({ tag: ValueTag.String, value, stringId: 0 });
const err = (code: ErrorCode): CellValue => ({ tag: ValueTag.Error, code });

function cellRange(values: CellValue[], rows: number, cols: number): RangeBuiltinArgument {
  return { kind: "range", refKind: "cells", values, rows, cols };
}

describe("lookup sort/filter builtins", () => {
  it("supports sorting, filtering, and uniqueness helpers", () => {
    const SORT = getLookupBuiltin("SORT")!;
    const SORTBY = getLookupBuiltin("SORTBY")!;
    const FILTER = getLookupBuiltin("FILTER")!;
    const UNIQUE = getLookupBuiltin("UNIQUE")!;

    expect(SORT(cellRange([num(3), num(1), num(2)], 3, 1))).toEqual({
      kind: "array",
      rows: 3,
      cols: 1,
      values: [num(1), num(2), num(3)],
    });
    expect(
      SORTBY(
        cellRange([text("b"), text("a"), text("c")], 3, 1),
        cellRange([num(2), num(1), num(3)], 3, 1),
      ),
    ).toEqual({
      kind: "array",
      rows: 3,
      cols: 1,
      values: [text("a"), text("b"), text("c")],
    });
    expect(
      FILTER(
        cellRange([text("north"), text("south"), text("east"), text("west")], 4, 1),
        cellRange([bool(true), bool(false), bool(true), bool(false)], 4, 1),
      ),
    ).toEqual({
      kind: "array",
      rows: 2,
      cols: 1,
      values: [text("north"), text("east")],
    });
    expect(UNIQUE(cellRange([text("A"), text("a"), text("B"), text("A")], 4, 1))).toEqual({
      kind: "array",
      rows: 2,
      cols: 1,
      values: [text("A"), text("B")],
    });
  });

  it("preserves explicit errors and missing-arg validation in sort helpers", () => {
    const SORT = getLookupBuiltin("SORT")!;
    const SORTBY = getLookupBuiltin("SORTBY")!;

    expect(Reflect.apply(SORT, undefined, [])).toEqual(err(ErrorCode.Value));
    expect(Reflect.apply(SORTBY, undefined, [undefined, cellRange([num(1)], 1, 1)])).toEqual(
      err(ErrorCode.Value),
    );
    expect(SORT(cellRange([num(3), num(1), num(2)], 3, 1), err(ErrorCode.Ref))).toEqual(
      err(ErrorCode.Ref),
    );
    expect(SORTBY(cellRange([num(3), num(1), num(2)], 3, 1), err(ErrorCode.Name))).toEqual(
      err(ErrorCode.Name),
    );
  });
});
