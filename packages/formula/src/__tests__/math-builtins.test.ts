import { describe, expect, it } from "vitest";
import { ErrorCode, ValueTag, type CellValue } from "@bilig/protocol";
import { getBuiltin } from "../builtins.js";

const num = (value: number): CellValue => ({ tag: ValueTag.Number, value });
const str = (value: string, stringId = 1): CellValue => ({ tag: ValueTag.String, value, stringId });
const valueError = { tag: ValueTag.Error, code: ErrorCode.Value } as const;

describe("math builtins", () => {
  it("rejects invalid scalar coercions instead of defaulting to zero or base 10", () => {
    expect(getBuiltin("ACOT")?.(str("bad"))).toEqual(valueError);
    expect(getBuiltin("ACOTH")?.(str("bad"))).toEqual(valueError);
    expect(getBuiltin("COT")?.(str("bad"))).toEqual(valueError);
    expect(getBuiltin("SEC")?.(str("bad"))).toEqual(valueError);
    expect(getBuiltin("LOG")?.(num(100), str("bad"))).toEqual(valueError);
  });

  it("keeps representative rounding, combinatoric, and bitwise behavior intact", () => {
    expect(getBuiltin("ROUNDUP")?.(num(12.341), num(2))).toEqual(num(12.35));
    expect(getBuiltin("COMBINA")?.(num(4), num(3))).toEqual(num(20));
    expect(getBuiltin("MROUND")?.(num(10), num(3))).toEqual(num(9));
    expect(getBuiltin("BITLSHIFT")?.(num(3), num(2))).toEqual(num(12));
  });
});
