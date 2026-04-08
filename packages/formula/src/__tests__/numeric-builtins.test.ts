import { describe, expect, it } from "vitest";
import { ErrorCode, ValueTag, type CellValue } from "@bilig/protocol";
import {
  buildIdentityMatrix,
  collectNumericArgs,
  collectStatNumericArgs,
  createNumericBuiltinHelpers,
  doubleFactorialValue,
  evenValue,
  factorialValue,
  gcdPair,
  lcmPair,
  oddValue,
  roundDownToDigits,
  roundToDigits,
  roundTowardZero,
  roundUpToDigits,
} from "../builtins/numeric.js";

function toNumber(value: CellValue): number | undefined {
  switch (value.tag) {
    case ValueTag.Number:
      return value.value;
    case ValueTag.Boolean:
      return value.value ? 1 : 0;
    case ValueTag.Empty:
      return 0;
    case ValueTag.String:
    case ValueTag.Error:
      return undefined;
  }
}

const numberResult = (value: number): CellValue => ({ tag: ValueTag.Number, value });
const valueError = (): CellValue => ({ tag: ValueTag.Error, code: ErrorCode.Value });
const numericResultOrError = (value: number): CellValue =>
  Number.isFinite(value) ? numberResult(value) : valueError();

describe("numeric builtin helpers", () => {
  it("handles directed rounding with positive and negative digit counts", () => {
    expect(roundToDigits(1499, -2)).toBe(1500);
    expect(roundUpToDigits(-12.341, 2)).toBe(-12.35);
    expect(roundDownToDigits(-12.341, 2)).toBe(-12.34);
    expect(roundTowardZero(-1234.56, -2)).toBe(-1200);
  });

  it("builds arithmetic helpers with injected coercion and error behavior", () => {
    const helpers = createNumericBuiltinHelpers({
      toNumber,
      numberResult,
      valueError,
      numericResultOrError,
    });

    expect(
      helpers.roundWith(
        { tag: ValueTag.Number, value: 12.34 },
        { tag: ValueTag.Number, value: -1 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 10 });
    expect(
      helpers.floorWith({ tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Number, value: 0 }),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Div0,
    });
    expect(
      helpers.ceilingWith(
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.String, value: "bad", stringId: 1 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });
    expect(helpers.unaryMath({ tag: ValueTag.Empty }, Math.sqrt)).toEqual({
      tag: ValueTag.Number,
      value: 0,
    });
    expect(
      helpers.binaryMath(
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Boolean, value: true },
        Math.pow,
      ),
    ).toEqual({
      tag: ValueTag.Number,
      value: 2,
    });
  });

  it("collects numeric aggregates with the expected coercion rules", () => {
    expect(
      collectNumericArgs(
        [
          { tag: ValueTag.Number, value: 3 },
          { tag: ValueTag.Boolean, value: true },
          { tag: ValueTag.Empty },
          { tag: ValueTag.String, value: "skip", stringId: 1 },
        ],
        toNumber,
      ),
    ).toEqual([3, 1, 0]);
    expect(
      collectStatNumericArgs([
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Boolean, value: false },
        { tag: ValueTag.Empty },
        { tag: ValueTag.String, value: "skip", stringId: 1 },
      ]),
    ).toEqual([3, 0]);
  });

  it("computes factorial and divisor helpers on truncated values", () => {
    expect(factorialValue(5.9)).toBe(120);
    expect(doubleFactorialValue(6.8)).toBe(48);
    expect(gcdPair(54.9, -24.1)).toBe(6);
    expect(lcmPair(-4.9, 6.2)).toBe(12);
    expect(evenValue(-3)).toBe(-4);
    expect(oddValue(-2)).toBe(-3);
  });

  it("builds identity matrices with injected numeric cells", () => {
    expect(buildIdentityMatrix(3, numberResult)).toEqual({
      kind: "array",
      rows: 3,
      cols: 3,
      values: [
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 1 },
      ],
    });
  });
});
