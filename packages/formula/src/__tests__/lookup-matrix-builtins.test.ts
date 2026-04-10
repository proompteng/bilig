import { describe, expect, it } from "vitest";
import { ErrorCode, ValueTag, type CellValue } from "@bilig/protocol";
import { getLookupBuiltin, type RangeBuiltinArgument } from "../builtins/lookup.js";

const num = (value: number): CellValue => ({ tag: ValueTag.Number, value });
const err = (code: ErrorCode): CellValue => ({ tag: ValueTag.Error, code });

function cellRange(values: CellValue[], rows: number, cols: number): RangeBuiltinArgument {
  return { kind: "range", refKind: "cells", values, rows, cols };
}

describe("lookup matrix builtins", () => {
  it("evaluates matrix and pairwise numeric helpers", () => {
    const SUMPRODUCT = getLookupBuiltin("SUMPRODUCT")!;
    const SUMX2MY2 = getLookupBuiltin("SUMX2MY2")!;
    const MINVERSE = getLookupBuiltin("MINVERSE")!;
    const MMULT = getLookupBuiltin("MMULT")!;
    const PERCENTOF = getLookupBuiltin("PERCENTOF")!;

    expect(
      SUMPRODUCT(cellRange([num(2), num(3)], 2, 1), cellRange([num(4), num(5)], 2, 1)),
    ).toEqual(num(23));
    expect(SUMX2MY2(cellRange([num(2), num(3)], 2, 1), cellRange([num(1), num(1)], 2, 1))).toEqual(
      num(11),
    );

    const inverse = MINVERSE(cellRange([num(4), num(7), num(2), num(6)], 2, 2));
    expect(inverse).toMatchObject({ kind: "array", rows: 2, cols: 2 });
    if (!("values" in inverse)) {
      throw new Error("expected MINVERSE to spill an array");
    }
    expect(inverse.values).toMatchObject([
      { tag: ValueTag.Number, value: expect.closeTo(0.6, 12) },
      { tag: ValueTag.Number, value: expect.closeTo(-0.7, 12) },
      { tag: ValueTag.Number, value: expect.closeTo(-0.2, 12) },
      { tag: ValueTag.Number, value: expect.closeTo(0.4, 12) },
    ]);

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
    expect(PERCENTOF(cellRange([num(2), num(3)], 2, 1), cellRange([num(10), num(10)], 2, 1))).toEqual(
      num(0.25),
    );
  });

  it("returns value errors for missing required matrix arguments instead of throwing", () => {
    const SUMX2MY2 = getLookupBuiltin("SUMX2MY2")!;
    const MDETERM = getLookupBuiltin("MDETERM")!;
    const MMULT = getLookupBuiltin("MMULT")!;
    const PERCENTOF = getLookupBuiltin("PERCENTOF")!;

    expect(Reflect.apply(SUMX2MY2, undefined, [cellRange([num(1)], 1, 1)])).toEqual(
      err(ErrorCode.Value),
    );
    expect(Reflect.apply(MDETERM, undefined, [])).toEqual(err(ErrorCode.Value));
    expect(Reflect.apply(MMULT, undefined, [cellRange([num(1)], 1, 1)])).toEqual(
      err(ErrorCode.Value),
    );
    expect(Reflect.apply(PERCENTOF, undefined, [cellRange([num(1)], 1, 1)])).toEqual(
      err(ErrorCode.Value),
    );
  });
});
