import { describe, expect, it } from "vitest";
import { ErrorCode, ValueTag, type CellValue } from "@bilig/protocol";
import { getLookupBuiltin, type RangeBuiltinArgument } from "../builtins/lookup.js";

const num = (value: number): CellValue => ({ tag: ValueTag.Number, value });
const err = (code: ErrorCode): CellValue => ({ tag: ValueTag.Error, code });

function cellRange(values: CellValue[], rows: number, cols: number): RangeBuiltinArgument {
  return { kind: "range", refKind: "cells", values, rows, cols };
}

describe("lookup hypothesis builtins", () => {
  it("supports legacy and modern hypothesis test wrappers", () => {
    const CHITEST = getLookupBuiltin("CHITEST")!;
    const FTEST = getLookupBuiltin("FTEST")!;
    const ZTEST = getLookupBuiltin("ZTEST")!;
    const TTEST = getLookupBuiltin("TTEST")!;

    expect(
      CHITEST(
        cellRange([num(10), num(20), num(20), num(40)], 2, 2),
        cellRange([num(15), num(15), num(15), num(45)], 2, 2),
      ),
    ).toMatchObject({ tag: ValueTag.Number, value: expect.any(Number) });
    expect(
      FTEST(
        cellRange([num(1), num(2), num(3)], 3, 1),
        cellRange([num(2), num(4), num(6)], 3, 1),
      ),
    ).toMatchObject({ tag: ValueTag.Number, value: expect.any(Number) });
    expect(ZTEST(cellRange([num(1), num(2), num(3)], 3, 1), num(2))).toMatchObject({
      tag: ValueTag.Number,
      value: expect.any(Number),
    });
    expect(
      TTEST(
        cellRange([num(1), num(2), num(3)], 3, 1),
        cellRange([num(2), num(4), num(6)], 3, 1),
        num(2),
        num(2),
      ),
    ).toMatchObject({ tag: ValueTag.Number, value: expect.any(Number) });
  });

  it("returns value errors for missing required hypothesis test args", () => {
    const CHITEST = getLookupBuiltin("CHITEST")!;
    const FTEST = getLookupBuiltin("FTEST")!;
    const ZTEST = getLookupBuiltin("ZTEST")!;
    const TTEST = getLookupBuiltin("TTEST")!;

    expect(Reflect.apply(CHITEST, undefined, [cellRange([num(1)], 1, 1)])).toEqual(
      err(ErrorCode.Value),
    );
    expect(Reflect.apply(FTEST, undefined, [cellRange([num(1)], 1, 1)])).toEqual(
      err(ErrorCode.Value),
    );
    expect(Reflect.apply(ZTEST, undefined, [cellRange([num(1)], 1, 1)])).toEqual(
      err(ErrorCode.Value),
    );
    expect(
      Reflect.apply(TTEST, undefined, [cellRange([num(1)], 1, 1), cellRange([num(2)], 1, 1)]),
    ).toEqual(err(ErrorCode.Value));
  });
});
