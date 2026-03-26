import { describe, expect, it } from "vitest";
import { BuiltinId, ErrorCode, ValueTag } from "@bilig/protocol";
import { getBuiltin, getBuiltinId } from "../builtins.js";
import type { ArrayValue } from "../runtime-values.js";
import {
  placeholderBuiltinNames,
  protocolPlaceholderBuiltinNames,
} from "../builtins/placeholder.js";

describe("formula builtins", () => {
  it("supports CHOOSE, COUNTBLANK, and bitwise builtins", () => {
    const CHOOSE = getBuiltin("CHOOSE")!;
    const COUNTBLANK = getBuiltin("COUNTBLANK")!;
    const BITAND = getBuiltin("BITAND")!;
    const BITOR = getBuiltin("BITOR")!;
    const BITXOR = getBuiltin("BITXOR")!;
    const BITLSHIFT = getBuiltin("BITLSHIFT")!;
    const BITRSHIFT = getBuiltin("BITRSHIFT")!;

    expect(
      CHOOSE(
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.String, value: "zero", stringId: 1 },
        {
          tag: ValueTag.Number,
          value: 10,
        },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 10 });
    expect(
      CHOOSE(
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Boolean, value: true },
        {
          tag: ValueTag.Number,
          value: 20,
        },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });
    expect(CHOOSE({ tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });

    expect(
      COUNTBLANK(
        { tag: ValueTag.Empty },
        { tag: ValueTag.String, value: "x", stringId: 1 },
        { tag: ValueTag.Empty },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 2 });

    expect(
      BITAND(
        { tag: ValueTag.Number, value: 6 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 1 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 0 });
    expect(
      BITOR(
        { tag: ValueTag.Number, value: 6 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 1 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 7 });
    expect(BITXOR({ tag: ValueTag.Number, value: 6 }, { tag: ValueTag.Number, value: 3 })).toEqual({
      tag: ValueTag.Number,
      value: 5,
    });
    expect(
      BITLSHIFT({ tag: ValueTag.Number, value: 1 }, { tag: ValueTag.Number, value: 4 }),
    ).toEqual({ tag: ValueTag.Number, value: 16 });
    expect(
      BITRSHIFT({ tag: ValueTag.Number, value: 16 }, { tag: ValueTag.Number, value: 4 }),
    ).toEqual({ tag: ValueTag.Number, value: 1 });
    expect(
      BITLSHIFT({ tag: ValueTag.Number, value: 1 }, { tag: ValueTag.String, value: "bad" }),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });
  });

  it("supports numeric aggregates and error propagation", () => {
    const sum = getBuiltin("SUM");
    const avg = getBuiltin("AVG");
    const mod = getBuiltin("MOD");

    expect(
      sum?.(
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Boolean, value: true },
        { tag: ValueTag.Empty },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 3 });

    expect(
      avg?.(
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.String, value: "skip", stringId: 1 },
        { tag: ValueTag.Empty },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 1 });

    expect(
      sum?.({ tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Error, code: ErrorCode.Ref }),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Ref });

    expect(mod?.({ tag: ValueTag.Number, value: 10 }, { tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Div0,
    });
  });

  it("supports boolean and string builtins and builtin ids", () => {
    expect(
      getBuiltin("AND")?.({ tag: ValueTag.Number, value: 1 }, { tag: ValueTag.Empty }),
    ).toEqual({ tag: ValueTag.Boolean, value: false });

    expect(
      getBuiltin("OR")?.({ tag: ValueTag.Empty }, { tag: ValueTag.Boolean, value: true }),
    ).toEqual({ tag: ValueTag.Boolean, value: true });

    expect(getBuiltin("NOT")?.({ tag: ValueTag.Boolean, value: false })).toEqual({
      tag: ValueTag.Boolean,
      value: true,
    });

    expect(getBuiltin("LEN")?.({ tag: ValueTag.Boolean, value: true })).toEqual({
      tag: ValueTag.Number,
      value: 4,
    });

    expect(
      getBuiltin("CONCAT")?.(
        { tag: ValueTag.String, value: "alpha", stringId: 1 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Empty },
      ),
    ).toEqual({ tag: ValueTag.String, value: "alpha2", stringId: 0 });

    expect(
      getBuiltin("EXACT")?.(
        { tag: ValueTag.String, value: "Alpha", stringId: 1 },
        { tag: ValueTag.String, value: "Alpha", stringId: 2 },
      ),
    ).toEqual({ tag: ValueTag.Boolean, value: true });

    expect(
      getBuiltin("EXACT")?.(
        { tag: ValueTag.String, value: "Alpha", stringId: 1 },
        { tag: ValueTag.String, value: "alpha", stringId: 2 },
      ),
    ).toEqual({ tag: ValueTag.Boolean, value: false });

    expect(
      getBuiltin("LEFT")?.(
        { tag: ValueTag.String, value: "alpha", stringId: 1 },
        { tag: ValueTag.Number, value: 3 },
      ),
    ).toEqual({ tag: ValueTag.String, value: "alp", stringId: 0 });

    expect(
      getBuiltin("TEXTBEFORE")?.(
        { tag: ValueTag.String, value: "alpha-beta", stringId: 1 },
        { tag: ValueTag.String, value: "-", stringId: 2 },
      ),
    ).toEqual({ tag: ValueTag.String, value: "alpha", stringId: 0 });

    expect(
      getBuiltin("IFERROR")?.(
        { tag: ValueTag.Error, code: ErrorCode.Div0 },
        { tag: ValueTag.String, value: "fallback", stringId: 1 },
      ),
    ).toEqual({ tag: ValueTag.String, value: "fallback", stringId: 1 });

    expect(
      getBuiltin("DATE")?.(
        { tag: ValueTag.Number, value: 2026 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 15 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 46096 });

    expect(
      getBuiltin("AVERAGE")?.(
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 4 },
        { tag: ValueTag.Number, value: 6 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 4 });

    expect(getBuiltinId("sum")).toBe(BuiltinId.Sum);
    expect(getBuiltinId("concat")).toBe(BuiltinId.Concat);
    expect(getBuiltinId("")).toBeUndefined();
    expect(getBuiltin("missing")).toBeUndefined();
  });

  it("supports the remaining scalar numeric builtins and conditional defaults", () => {
    expect(
      getBuiltin("MIN")?.(
        { tag: ValueTag.Empty },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: -1 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: -1 });

    expect(
      getBuiltin("MAX")?.(
        { tag: ValueTag.String, value: "skip", stringId: 1 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Boolean, value: true },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 3 });

    expect(
      getBuiltin("COUNT")?.(
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Boolean, value: false },
        { tag: ValueTag.String, value: "skip", stringId: 1 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 2 });

    expect(
      getBuiltin("COUNTA")?.(
        { tag: ValueTag.Empty },
        { tag: ValueTag.String, value: "x", stringId: 1 },
        { tag: ValueTag.Boolean, value: false },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 2 });

    expect(getBuiltin("ABS")?.({ tag: ValueTag.Number, value: -3.4 })).toEqual({
      tag: ValueTag.Number,
      value: 3.4,
    });
    expect(getBuiltin("INT")?.({ tag: ValueTag.Number, value: -3.1 })).toEqual({
      tag: ValueTag.Number,
      value: -4,
    });
    expect(getBuiltin("ROUND")?.({ tag: ValueTag.Number, value: 3.6 })).toEqual({
      tag: ValueTag.Number,
      value: 4,
    });
    expect(
      getBuiltin("ROUNDUP")?.(
        { tag: ValueTag.Number, value: 3.145 },
        { tag: ValueTag.Number, value: 2 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 3.15 });
    expect(
      getBuiltin("ROUNDDOWN")?.(
        { tag: ValueTag.Number, value: -3.145 },
        { tag: ValueTag.Number, value: 2 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: -3.14 });
    expect(
      getBuiltin("ROUND")?.(
        { tag: ValueTag.Number, value: 3.145 },
        { tag: ValueTag.Number, value: 2 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 3.15 });
    expect(getBuiltin("FLOOR")?.({ tag: ValueTag.Number, value: 3.6 })).toEqual({
      tag: ValueTag.Number,
      value: 3,
    });
    expect(
      getBuiltin("FLOOR")?.({ tag: ValueTag.Number, value: 7 }, { tag: ValueTag.Number, value: 2 }),
    ).toEqual({ tag: ValueTag.Number, value: 6 });
    expect(getBuiltin("CEILING")?.({ tag: ValueTag.Number, value: 3.1 })).toEqual({
      tag: ValueTag.Number,
      value: 4,
    });
    expect(
      getBuiltin("CEILING")?.(
        { tag: ValueTag.Number, value: 7 },
        { tag: ValueTag.Number, value: 2 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 8 });

    expect(
      getBuiltin("IF")?.(
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.String, value: "truthy", stringId: 1 },
      ),
    ).toEqual({ tag: ValueTag.String, value: "truthy", stringId: 1 });

    expect(getBuiltin("IF")?.({ tag: ValueTag.Empty }, { tag: ValueTag.Number, value: 1 })).toEqual(
      { tag: ValueTag.Empty },
    );
  });

  it("supports expanded math and numeric utility builtins", () => {
    expect(getBuiltin("SIN")?.({ tag: ValueTag.Number, value: Math.PI / 2 })).toEqual({
      tag: ValueTag.Number,
      value: 1,
    });
    expect(getBuiltin("COS")?.({ tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Number,
      value: 1,
    });
    expect(
      getBuiltin("POWER")?.({ tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Number, value: 3 }),
    ).toEqual({ tag: ValueTag.Number, value: 8 });
    expect(getBuiltin("LOG")?.({ tag: ValueTag.Number, value: 1000 })).toEqual({
      tag: ValueTag.Number,
      value: 3,
    });
    expect(getBuiltin("SIGN")?.({ tag: ValueTag.Number, value: -9 })).toEqual({
      tag: ValueTag.Number,
      value: -1,
    });
    expect(
      getBuiltin("TRUNC")?.(
        { tag: ValueTag.Number, value: -3.98 },
        { tag: ValueTag.Number, value: 1 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: -3.9 });
    expect(
      getBuiltin("CEILING.MATH")?.(
        { tag: ValueTag.Number, value: -5.5 },
        { tag: ValueTag.Number, value: 2 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: -4 });
    expect(
      getBuiltin("FLOOR.PRECISE")?.(
        { tag: ValueTag.Number, value: -5.5 },
        { tag: ValueTag.Number, value: 2 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: -6 });
    expect(getBuiltin("FACT")?.({ tag: ValueTag.Number, value: 5 })).toEqual({
      tag: ValueTag.Number,
      value: 120,
    });
    expect(
      getBuiltin("COMBIN")?.(
        { tag: ValueTag.Number, value: 5 },
        { tag: ValueTag.Number, value: 2 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 10 });
    expect(
      getBuiltin("GCD")?.({ tag: ValueTag.Number, value: 18 }, { tag: ValueTag.Number, value: 24 }),
    ).toEqual({ tag: ValueTag.Number, value: 6 });
    expect(
      getBuiltin("LCM")?.({ tag: ValueTag.Number, value: 6 }, { tag: ValueTag.Number, value: 8 }),
    ).toEqual({ tag: ValueTag.Number, value: 24 });
    expect(
      getBuiltin("MROUND")?.(
        { tag: ValueTag.Number, value: 10 },
        { tag: ValueTag.Number, value: 3 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 9 });
    expect(
      getBuiltin("PRODUCT")?.(
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 4 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 24 });
    expect(
      getBuiltin("QUOTIENT")?.(
        { tag: ValueTag.Number, value: 7 },
        { tag: ValueTag.Number, value: 2 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 3 });
    expect(
      getBuiltin("SUMSQ")?.({ tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Number, value: 3 }),
    ).toEqual({ tag: ValueTag.Number, value: 13 });
    expect(
      getBuiltin("SERIESSUM")?.(
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 2 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 18 });
    expect(
      getBuiltin("BASE")?.(
        { tag: ValueTag.Number, value: 31 },
        { tag: ValueTag.Number, value: 16 },
        { tag: ValueTag.Number, value: 4 },
      ),
    ).toEqual({ tag: ValueTag.String, value: "001F", stringId: 0 });
    expect(
      getBuiltin("DECIMAL")?.(
        { tag: ValueTag.String, value: "1F", stringId: 1 },
        { tag: ValueTag.Number, value: 16 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 31 });
    expect(getBuiltin("ROMAN")?.({ tag: ValueTag.Number, value: 14 })).toEqual({
      tag: ValueTag.String,
      value: "XIV",
      stringId: 0,
    });
    expect(getBuiltin("ARABIC")?.({ tag: ValueTag.String, value: "XIV", stringId: 1 })).toEqual({
      tag: ValueTag.Number,
      value: 14,
    });

    expect(getBuiltin("MUNIT")?.({ tag: ValueTag.Number, value: 2 })).toEqual({
      kind: "array",
      rows: 2,
      cols: 2,
      values: [
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 1 },
      ],
    });

    const randArray = getBuiltin("RANDARRAY")?.(
      { tag: ValueTag.Number, value: 2 },
      { tag: ValueTag.Number, value: 2 },
      { tag: ValueTag.Number, value: 3 },
      { tag: ValueTag.Number, value: 7 },
      { tag: ValueTag.Boolean, value: true },
    );
    expect(randArray).toMatchObject({ kind: "array", rows: 2, cols: 2 });
    if (!(randArray && "kind" in randArray && randArray.kind === "array")) {
      throw new Error("expected RANDARRAY to return an array");
    }
    for (const value of randArray.values) {
      expect(value.tag).toBe(ValueTag.Number);
      expect(value.value).toBeGreaterThanOrEqual(3);
      expect(value.value).toBeLessThanOrEqual(7);
    }
  });

  it("covers math builtin edge cases and aggregate variants", () => {
    expect(
      getBuiltin("CEILING.PRECISE")?.({ tag: ValueTag.String, value: "bad", stringId: 1 }),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });
    expect(
      getBuiltin("ISO.CEILING")?.(
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 0 },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });
    expect(
      getBuiltin("ROUNDUP")?.(
        { tag: ValueTag.Number, value: 5 },
        { tag: ValueTag.String, value: "bad", stringId: 1 },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });
    expect(
      getBuiltin("TRUNC")?.(
        { tag: ValueTag.Number, value: 5 },
        { tag: ValueTag.String, value: "bad", stringId: 1 },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });
    expect(getBuiltin("EVEN")?.({ tag: ValueTag.String, value: "bad", stringId: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });
    expect(getBuiltin("ODD")?.({ tag: ValueTag.String, value: "bad", stringId: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });
    expect(getBuiltin("LN")?.({ tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });
    expect(getBuiltin("SQRT")?.({ tag: ValueTag.Number, value: -1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });
    expect(getBuiltin("COT")?.({ tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Div0,
    });
    expect(getBuiltin("CSC")?.({ tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Div0,
    });
    expect(getBuiltin("FACT")?.({ tag: ValueTag.Number, value: -1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });
    expect(getBuiltin("FACTDOUBLE")?.({ tag: ValueTag.Number, value: 6 })).toEqual({
      tag: ValueTag.Number,
      value: 48,
    });
    expect(
      getBuiltin("COMBIN")?.(
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 3 },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });
    expect(
      getBuiltin("COMBINA")?.(
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 2 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 0 });
    expect(getBuiltin("GCD")?.()).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });
    expect(getBuiltin("LCM")?.()).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });
    expect(
      getBuiltin("MROUND")?.(
        { tag: ValueTag.Number, value: -10 },
        { tag: ValueTag.Number, value: 3 },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });
    expect(
      getBuiltin("MULTINOMIAL")?.(
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: -1 },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });
    expect(
      getBuiltin("QUOTIENT")?.(
        { tag: ValueTag.String, value: "bad", stringId: 1 },
        { tag: ValueTag.Number, value: 1 },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });
    expect(
      getBuiltin("QUOTIENT")?.(
        { tag: ValueTag.Number, value: 5 },
        { tag: ValueTag.Number, value: 0 },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Div0 });
    expect(
      getBuiltin("RANDBETWEEN")?.(
        { tag: ValueTag.Number, value: 5 },
        { tag: ValueTag.Number, value: 3 },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });
    expect(
      getBuiltin("BASE")?.({ tag: ValueTag.Number, value: -1 }, { tag: ValueTag.Number, value: 2 }),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });
    expect(
      getBuiltin("DECIMAL")?.(
        { tag: ValueTag.String, value: "1Z", stringId: 1 },
        { tag: ValueTag.Number, value: 10 },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });
    expect(getBuiltin("ROMAN")?.({ tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });
    expect(getBuiltin("ARABIC")?.({ tag: ValueTag.String, value: "IIII", stringId: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });
    expect(
      getBuiltin("RANDARRAY")?.(
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 5 },
        { tag: ValueTag.Number, value: 3 },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });
    expect(
      getBuiltin("RANDARRAY")?.(
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 1 },
      ),
    ).toMatchObject({ kind: "array", rows: 2, cols: 2 });
    expect(getBuiltin("MUNIT")?.({ tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });
    expect(
      getBuiltin("SERIESSUM")?.(
        { tag: ValueTag.String, value: "bad", stringId: 1 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 1 },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });
    expect(getBuiltin("SQRTPI")?.({ tag: ValueTag.String, value: "bad", stringId: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });
    expect(
      getBuiltin("SUBTOTAL")?.(
        { tag: ValueTag.Number, value: 9 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 3 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 5 });
    expect(getBuiltin("SUBTOTAL")?.({ tag: ValueTag.String, value: "bad", stringId: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });
    expect(
      getBuiltin("AGGREGATE")?.(
        { tag: ValueTag.Number, value: 6 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 4 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 24 });
    expect(
      getBuiltin("AGGREGATE")?.(
        { tag: ValueTag.Number, value: 99 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 2 },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });
    expect(getBuiltin("ARABIC")?.({ tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });
  });

  it("covers address, formatting, and mean helper builtins", () => {
    expect(
      getBuiltin("MAXA")?.(
        { tag: ValueTag.Boolean, value: true },
        { tag: ValueTag.String, value: "skip", stringId: 1 },
        { tag: ValueTag.Empty },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 1 });
    expect(
      getBuiltin("MINA")?.(
        { tag: ValueTag.Boolean, value: true },
        { tag: ValueTag.Boolean, value: false },
        { tag: ValueTag.String, value: "skip", stringId: 1 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 0 });

    expect(
      getBuiltin("ADDRESS")?.(
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 28 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.String, value: "O'Brien", stringId: 1 },
      ),
    ).toEqual({ tag: ValueTag.String, value: "'O''Brien'!$AB2", stringId: 0 });
    expect(
      getBuiltin("ADDRESS")?.(
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 28 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 2 },
      ),
    ).toEqual({ tag: ValueTag.String, value: "R2C[28]", stringId: 0 });
    expect(
      getBuiltin("ADDRESS")?.(
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 28 },
        { tag: ValueTag.Number, value: 5 },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });
    expect(
      getBuiltin("ADDRESS")?.(
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 28 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Empty },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });

    expect(
      getBuiltin("DOLLAR")?.(
        { tag: ValueTag.Number, value: -1234.567 },
        { tag: ValueTag.Number, value: 1 },
      ),
    ).toEqual({ tag: ValueTag.String, value: "-$1,234.6", stringId: 0 });
    expect(
      getBuiltin("DOLLAR")?.(
        { tag: ValueTag.Number, value: 1234.567 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Boolean, value: true },
      ),
    ).toEqual({ tag: ValueTag.String, value: "$1234.6", stringId: 0 });
    expect(
      getBuiltin("FIXED")?.(
        { tag: ValueTag.Number, value: 1234.567 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Boolean, value: true },
      ),
    ).toEqual({ tag: ValueTag.String, value: "1234.6", stringId: 0 });
    expect(
      getBuiltin("FIXED")?.(
        { tag: ValueTag.Number, value: 1234.567 },
        { tag: ValueTag.String, value: "bad", stringId: 1 },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });
    expect(
      getBuiltin("DOLLARDE")?.(
        { tag: ValueTag.Number, value: 1.08 },
        { tag: ValueTag.Number, value: 16 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 1.5 });
    expect(
      getBuiltin("DOLLARFR")?.(
        { tag: ValueTag.Number, value: 1.5 },
        { tag: ValueTag.Number, value: 16 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 1.08 });
    expect(
      getBuiltin("DOLLARFR")?.(
        { tag: ValueTag.Number, value: 1.5 },
        { tag: ValueTag.Number, value: 0 },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });

    expect(
      getBuiltin("GEOMEAN")?.(
        { tag: ValueTag.Number, value: 4 },
        { tag: ValueTag.Number, value: 1 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 2 });
    expect(
      getBuiltin("GEOMEAN")?.(
        { tag: ValueTag.Number, value: -1 },
        { tag: ValueTag.Number, value: 4 },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });
    expect(
      getBuiltin("HARMEAN")?.(
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 4 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 3 / 1.75 });
    expect(
      getBuiltin("HARMEAN")?.(
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 0 },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });
  });

  it("covers extended trigonometric and precise rounding builtins", () => {
    expect(getBuiltin("SINH")?.({ tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Number,
      value: 0,
    });
    expect(getBuiltin("COSH")?.({ tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Number,
      value: 1,
    });
    expect(getBuiltin("TANH")?.({ tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Number,
      value: 0,
    });
    expect(getBuiltin("ASINH")?.({ tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Number,
      value: 0,
    });
    expect(getBuiltin("ACOSH")?.({ tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.Number,
      value: 0,
    });
    expect(getBuiltin("ATANH")?.({ tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Number,
      value: 0,
    });
    expect(getBuiltin("ACOT")?.({ tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.Number,
      value: Math.PI / 4,
    });
    expect(getBuiltin("ACOT")?.({ tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Number,
      value: Math.PI / 2,
    });
    expect(getBuiltin("ACOTH")?.({ tag: ValueTag.Number, value: 2 })).toEqual({
      tag: ValueTag.Number,
      value: 0.5 * Math.log(3),
    });
    expect(getBuiltin("COTH")?.({ tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Div0,
    });
    expect(getBuiltin("CSCH")?.({ tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Div0,
    });
    expect(getBuiltin("SEC")?.({ tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Number,
      value: 1,
    });
    expect(getBuiltin("SECH")?.({ tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Number,
      value: 1,
    });
    expect(getBuiltin("SIGN")?.({ tag: ValueTag.String, value: "bad", stringId: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });
    expect(
      getBuiltin("FLOOR.MATH")?.(
        { tag: ValueTag.Number, value: -5.5 },
        { tag: ValueTag.Number, value: 2 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: -6 });
    expect(
      getBuiltin("FLOOR.MATH")?.(
        { tag: ValueTag.Number, value: -5.5 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 1 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: -4 });
    expect(
      getBuiltin("CEILING.MATH")?.(
        { tag: ValueTag.Number, value: -5.5 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 1 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: -6 });
    expect(
      getBuiltin("CEILING.PRECISE")?.(
        { tag: ValueTag.Number, value: 5.1 },
        { tag: ValueTag.Number, value: 2 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 6 });
    expect(
      getBuiltin("ISO.CEILING")?.(
        { tag: ValueTag.Number, value: 5.1 },
        { tag: ValueTag.Number, value: 2 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 6 });
  });

  it("supports ACCRINT, ACCRINTM, AMORDEGRC, and AMORLINC", () => {
    const issue = getBuiltin("DATE")?.(
      { tag: ValueTag.Number, value: 2020 },
      { tag: ValueTag.Number, value: 2 },
      { tag: ValueTag.Number, value: 1 },
    );
    const firstInterest = getBuiltin("DATE")?.(
      { tag: ValueTag.Number, value: 2020 },
      { tag: ValueTag.Number, value: 11 },
      { tag: ValueTag.Number, value: 30 },
    );
    const settlement = getBuiltin("DATE")?.(
      { tag: ValueTag.Number, value: 2020 },
      { tag: ValueTag.Number, value: 12 },
      { tag: ValueTag.Number, value: 31 },
    );
    const cost = { tag: ValueTag.Number, value: 2000 };
    const salvage = { tag: ValueTag.Number, value: 10 };
    const period = { tag: ValueTag.Number, value: 4 };
    const rate = { tag: ValueTag.Number, value: 0.1 };
    const basis = { tag: ValueTag.Number, value: 0 };

    expect(issue?.tag).toBe(ValueTag.Number);
    expect(firstInterest?.tag).toBe(ValueTag.Number);
    expect(settlement?.tag).toBe(ValueTag.Number);

    const firstAccrual = getBuiltin("ACCRINT")?.(
      issue,
      firstInterest,
      settlement,
      rate,
      { tag: ValueTag.Number, value: 1000 },
      { tag: ValueTag.Number, value: 2 },
      basis,
    );
    expect(firstAccrual).toMatchObject({ tag: ValueTag.Number });
    expect(firstAccrual?.tag === ValueTag.Number ? firstAccrual.value : Number.NaN).toBeCloseTo(
      91.66666666666667,
      12,
    );

    const omittedBasisAccrual = getBuiltin("ACCRINT")?.(
      issue,
      firstInterest,
      settlement,
      rate,
      { tag: ValueTag.Number, value: 1000 },
      { tag: ValueTag.Number, value: 2 },
    );
    expect(omittedBasisAccrual).toMatchObject({ tag: ValueTag.Number });
    expect(
      omittedBasisAccrual?.tag === ValueTag.Number ? omittedBasisAccrual.value : Number.NaN,
    ).toBeCloseTo(91.66666666666667, 12);

    const fullAccrual = getBuiltin("ACCRINT")?.(
      issue,
      firstInterest,
      settlement,
      rate,
      { tag: ValueTag.Number, value: 1000 },
      { tag: ValueTag.Number, value: 2 },
      basis,
    );
    const shortAccrual = getBuiltin("ACCRINT")?.(
      issue,
      firstInterest,
      settlement,
      rate,
      { tag: ValueTag.Number, value: 1000 },
      { tag: ValueTag.Number, value: 2 },
      basis,
      { tag: ValueTag.Boolean, value: false },
    );
    expect(fullAccrual).toMatchObject({ tag: ValueTag.Number });
    expect(shortAccrual).toMatchObject({ tag: ValueTag.Number });
    const shortAccrualValue =
      shortAccrual?.tag === ValueTag.Number ? shortAccrual.value : Number.NaN;
    const fullAccrualValue = fullAccrual?.tag === ValueTag.Number ? fullAccrual.value : Number.NaN;
    expect(shortAccrualValue).toBeLessThan(fullAccrualValue);

    const maturityAccrual = getBuiltin("ACCRINTM")?.(issue, settlement, rate, undefined, basis);
    expect(maturityAccrual).toMatchObject({ tag: ValueTag.Number });
    expect(
      maturityAccrual?.tag === ValueTag.Number ? maturityAccrual.value : Number.NaN,
    ).toBeCloseTo(91.66666666666667, 12);

    expect(getBuiltin("AMORLINC")?.(cost, issue, settlement, salvage, period, rate, basis)).toEqual(
      { tag: ValueTag.Number, value: 200 },
    );

    expect(
      getBuiltin("AMORDEGRC")?.(cost, issue, settlement, salvage, period, rate, basis),
    ).toEqual({ tag: ValueTag.Number, value: 163 });

    expect(
      getBuiltin("ACCRINT")?.(
        issue,
        settlement,
        issue,
        rate,
        { tag: ValueTag.Number, value: 1000 },
        { tag: ValueTag.Number, value: 2 },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });

    expect(
      getBuiltin("ACCRINT")?.(
        issue,
        settlement,
        settlement,
        rate,
        { tag: ValueTag.Number, value: 1000 },
        { tag: ValueTag.Number, value: 3 },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });
  });

  it("covers ACCRINT and ACCRINTM basis variants and invalid argument branches", () => {
    const issue = getBuiltin("DATE")?.(
      { tag: ValueTag.Number, value: 2020 },
      { tag: ValueTag.Number, value: 1 },
      { tag: ValueTag.Number, value: 1 },
    );
    const firstInterest = getBuiltin("DATE")?.(
      { tag: ValueTag.Number, value: 2020 },
      { tag: ValueTag.Number, value: 2 },
      { tag: ValueTag.Number, value: 1 },
    );
    const settlement = getBuiltin("DATE")?.(
      { tag: ValueTag.Number, value: 2021 },
      { tag: ValueTag.Number, value: 1 },
      { tag: ValueTag.Number, value: 1 },
    );
    const rate = { tag: ValueTag.Number, value: 0.08 };
    const par = { tag: ValueTag.Number, value: 1000 };
    const frequency = { tag: ValueTag.Number, value: 2 };

    expect(issue?.tag).toBe(ValueTag.Number);
    expect(firstInterest?.tag).toBe(ValueTag.Number);
    expect(settlement?.tag).toBe(ValueTag.Number);

    for (const basis of [0, 1, 2, 3, 4]) {
      expect(
        getBuiltin("ACCRINT")?.(issue, firstInterest, settlement, rate, par, frequency, {
          tag: ValueTag.Number,
          value: basis,
        }),
      ).toMatchObject({ tag: ValueTag.Number });
      expect(
        getBuiltin("ACCRINTM")?.(issue, settlement, rate, par, {
          tag: ValueTag.Number,
          value: basis,
        }),
      ).toMatchObject({ tag: ValueTag.Number });
    }

    expect(
      getBuiltin("ACCRINT")?.(issue, firstInterest, settlement, rate, par, frequency, {
        tag: ValueTag.Number,
        value: 5,
      }),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });

    expect(
      getBuiltin("ACCRINTM")?.(issue, settlement, rate, par, { tag: ValueTag.Number, value: 5 }),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });
  });

  it("covers AMORLINC and AMORDEGRC branch-heavy scenarios", () => {
    const cost = { tag: ValueTag.Number, value: 1000 };
    const datePurchased = getBuiltin("DATE")?.(
      { tag: ValueTag.Number, value: 2020 },
      { tag: ValueTag.Number, value: 1 },
      { tag: ValueTag.Number, value: 1 },
    );
    const firstPeriod = getBuiltin("DATE")?.(
      { tag: ValueTag.Number, value: 2021 },
      { tag: ValueTag.Number, value: 1 },
      { tag: ValueTag.Number, value: 1 },
    );
    const basis = { tag: ValueTag.Number, value: 0 };

    expect(datePurchased?.tag).toBe(ValueTag.Number);
    expect(firstPeriod?.tag).toBe(ValueTag.Number);

    expect(
      getBuiltin("AMORLINC")?.(
        cost,
        datePurchased,
        firstPeriod,
        { tag: ValueTag.Number, value: 25 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 0.15 },
        basis,
      ),
    ).toMatchObject({ tag: ValueTag.Number, value: 150 });

    expect(
      getBuiltin("AMORLINC")?.(
        cost,
        datePurchased,
        firstPeriod,
        { tag: ValueTag.Number, value: 25 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 0.15 },
        basis,
      ),
    ).toMatchObject({ tag: ValueTag.Number, value: 150 });

    expect(
      getBuiltin("AMORLINC")?.(
        cost,
        datePurchased,
        firstPeriod,
        { tag: ValueTag.Number, value: 25 },
        { tag: ValueTag.Number, value: 6 },
        { tag: ValueTag.Number, value: 0.15 },
        basis,
      ),
    ).toMatchObject({ tag: ValueTag.Number, value: 75 });

    expect(
      getBuiltin("AMORLINC")?.(
        cost,
        datePurchased,
        firstPeriod,
        { tag: ValueTag.Number, value: 25 },
        { tag: ValueTag.Number, value: 7 },
        { tag: ValueTag.Number, value: 0.15 },
        basis,
      ),
    ).toEqual({ tag: ValueTag.Number, value: 0 });

    expect(
      getBuiltin("AMORDEGRC")?.(
        { tag: ValueTag.Number, value: 1000 },
        datePurchased,
        firstPeriod,
        { tag: ValueTag.Number, value: 10 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 0.2 },
        basis,
      ),
    ).toMatchObject({ tag: ValueTag.Number, value: 240 });

    expect(
      getBuiltin("AMORDEGRC")?.(
        { tag: ValueTag.Number, value: 1000 },
        datePurchased,
        firstPeriod,
        { tag: ValueTag.Number, value: 10 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 0.3 },
        basis,
      ),
    ).toMatchObject({ tag: ValueTag.Number, value: 247 });

    expect(
      getBuiltin("AMORDEGRC")?.(
        { tag: ValueTag.Number, value: 1000 },
        datePurchased,
        firstPeriod,
        { tag: ValueTag.Number, value: 10 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 0.5 },
        basis,
      ),
    ).toMatchObject({ tag: ValueTag.Number, value: 250 });

    expect(
      getBuiltin("AMORDEGRC")?.(
        { tag: ValueTag.Number, value: 1000 },
        datePurchased,
        firstPeriod,
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 1.2 },
        basis,
      ),
    ).toEqual({ tag: ValueTag.Number, value: 0 });

    expect(
      getBuiltin("AMORDEGRC")?.(
        { tag: ValueTag.Number, value: 100 },
        datePurchased,
        firstPeriod,
        { tag: ValueTag.Number, value: 200 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 0.1 },
        basis,
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });
  });

  it("covers combinatorics, product, quotient, and financial validation edge branches", () => {
    expect(
      getBuiltin("COMBINA")?.(
        { tag: ValueTag.Number, value: 4 },
        { tag: ValueTag.Number, value: 0 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 1 });
    expect(
      getBuiltin("COMBINA")?.(
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 3 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 0 });
    expect(
      getBuiltin("COMBINA")?.(
        { tag: ValueTag.String, value: "bad", stringId: 1 },
        { tag: ValueTag.Number, value: 1 },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });

    expect(
      getBuiltin("GCD")?.(
        { tag: ValueTag.Number, value: 54 },
        { tag: ValueTag.Number, value: 24.9 },
        { tag: ValueTag.Number, value: 6 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 6 });
    expect(
      getBuiltin("LCM")?.(
        { tag: ValueTag.Number, value: 4 },
        { tag: ValueTag.Number, value: 6 },
        { tag: ValueTag.Number, value: 3.8 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 12 });
    expect(
      getBuiltin("MROUND")?.(
        { tag: ValueTag.Number, value: 10 },
        { tag: ValueTag.Number, value: 0 },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });
    expect(
      getBuiltin("MROUND")?.(
        { tag: ValueTag.Number, value: 10 },
        { tag: ValueTag.Number, value: 4 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 12 });
    expect(
      getBuiltin("MULTINOMIAL")?.(
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 3 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 60 });
    expect(getBuiltin("PRODUCT")?.()).toEqual({ tag: ValueTag.Number, value: 0 });
    expect(
      getBuiltin("PRODUCT")?.(
        { tag: ValueTag.Error, code: ErrorCode.Ref },
        { tag: ValueTag.Number, value: 2 },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Ref });
    expect(
      getBuiltin("QUOTIENT")?.(
        { tag: ValueTag.Number, value: 7 },
        { tag: ValueTag.Number, value: 2 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 3 });

    const issue = getBuiltin("DATE")?.(
      { tag: ValueTag.Number, value: 2020 },
      { tag: ValueTag.Number, value: 1 },
      { tag: ValueTag.Number, value: 1 },
    );
    const settlement = getBuiltin("DATE")?.(
      { tag: ValueTag.Number, value: 2020 },
      { tag: ValueTag.Number, value: 12 },
      { tag: ValueTag.Number, value: 31 },
    );
    expect(issue?.tag).toBe(ValueTag.Number);
    expect(settlement?.tag).toBe(ValueTag.Number);
    expect(
      getBuiltin("AMORDEGRC")?.(
        { tag: ValueTag.String, value: "bad", stringId: 2 },
        issue,
        settlement,
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 0.1 },
        { tag: ValueTag.Number, value: 0 },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });
    expect(
      getBuiltin("AMORLINC")?.(
        { tag: ValueTag.Number, value: 1000 },
        issue,
        settlement,
        { tag: ValueTag.String, value: "bad", stringId: 3 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 0.1 },
        { tag: ValueTag.Number, value: 0 },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });
  });

  it("covers bitwise, integer, and rounding validation branches", () => {
    expect(getBuiltin("BITXOR")?.({ tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });
    expect(
      getBuiltin("BITXOR")?.(
        { tag: ValueTag.String, value: "bad", stringId: 1 },
        { tag: ValueTag.Number, value: 1 },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });
    expect(
      getBuiltin("BITXOR")?.(
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.String, value: "bad", stringId: 2 },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });
    expect(
      getBuiltin("BITLSHIFT")?.(
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.String, value: "bad", stringId: 3 },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });
    expect(
      getBuiltin("BITRSHIFT")?.(
        { tag: ValueTag.String, value: "bad", stringId: 4 },
        { tag: ValueTag.Number, value: 1 },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });
    expect(getBuiltin("INT")?.({ tag: ValueTag.String, value: "bad", stringId: 5 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });
    expect(
      getBuiltin("ROUNDUP")?.(
        { tag: ValueTag.Number, value: 12.34 },
        { tag: ValueTag.String, value: "bad", stringId: 6 },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });
    expect(
      getBuiltin("ROUNDDOWN")?.(
        { tag: ValueTag.Number, value: 12.34 },
        { tag: ValueTag.String, value: "bad", stringId: 7 },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });
    expect(
      getBuiltin("TRUNC")?.(
        { tag: ValueTag.Number, value: 12.34 },
        { tag: ValueTag.String, value: "bad", stringId: 8 },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });
    expect(getBuiltin("TRUNC")?.({ tag: ValueTag.Number, value: -12.34 })).toEqual({
      tag: ValueTag.Number,
      value: -12,
    });
  });

  it("covers ceiling, floor, parity, factorial, and combinatoric branches", () => {
    expect(
      getBuiltin("FLOOR.MATH")?.(
        { tag: ValueTag.Number, value: -5.5 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 0 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: -6 });
    expect(
      getBuiltin("FLOOR.MATH")?.(
        { tag: ValueTag.Number, value: -5.5 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 1 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: -4 });
    expect(
      getBuiltin("FLOOR.PRECISE")?.(
        { tag: ValueTag.Number, value: -5.5 },
        { tag: ValueTag.Number, value: 2 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: -6 });
    expect(
      getBuiltin("CEILING.MATH")?.(
        { tag: ValueTag.Number, value: -5.5 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 0 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: -4 });
    expect(
      getBuiltin("CEILING.MATH")?.(
        { tag: ValueTag.Number, value: -5.5 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 1 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: -6 });
    expect(
      getBuiltin("CEILING.PRECISE")?.(
        { tag: ValueTag.Number, value: -5.5 },
        { tag: ValueTag.Number, value: 2 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: -4 });
    expect(
      getBuiltin("ISO.CEILING")?.(
        { tag: ValueTag.Number, value: -5.5 },
        { tag: ValueTag.Number, value: 2 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: -4 });
    expect(
      getBuiltin("CEILING.PRECISE")?.(
        { tag: ValueTag.String, value: "bad", stringId: 9 },
        { tag: ValueTag.Number, value: 2 },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });
    expect(
      getBuiltin("ISO.CEILING")?.(
        { tag: ValueTag.Number, value: 4 },
        { tag: ValueTag.Number, value: 0 },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });

    expect(getBuiltin("BITAND")?.({ tag: ValueTag.Number, value: 3 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });
    expect(
      getBuiltin("BITAND")?.(
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.String, value: "bad", stringId: 10 },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });
    expect(getBuiltin("BITOR")?.({ tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });
    expect(
      getBuiltin("BITOR")?.(
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.String, value: "bad", stringId: 11 },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });

    expect(getBuiltin("EVEN")?.({ tag: ValueTag.Number, value: -3 })).toEqual({
      tag: ValueTag.Number,
      value: -4,
    });
    expect(getBuiltin("ODD")?.({ tag: ValueTag.Number, value: -2 })).toEqual({
      tag: ValueTag.Number,
      value: -3,
    });
    expect(getBuiltin("FACT")?.({ tag: ValueTag.Number, value: -1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });
    expect(getBuiltin("FACTDOUBLE")?.({ tag: ValueTag.Number, value: -3 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });
    expect(
      getBuiltin("COMBIN")?.(
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 4 },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });
    expect(
      getBuiltin("COMBINA")?.(
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 2 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 6 });
    expect(getBuiltin("GCD")?.()).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });
    expect(getBuiltin("LCM")?.()).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });
  });

  it("covers logarithmic, hyperbolic, and sign-related math edge branches", () => {
    expect(getBuiltin("HARMEAN")?.()).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });
    expect(
      getBuiltin("HARMEAN")?.(
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: -1 },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });

    expect(getBuiltin("LOG10")?.({ tag: ValueTag.Number, value: 1000 })).toEqual({
      tag: ValueTag.Number,
      value: 3,
    });
    expect(getBuiltin("LOG10")?.({ tag: ValueTag.Number, value: -1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });
    expect(
      getBuiltin("LOG")?.({ tag: ValueTag.Number, value: 8 }, { tag: ValueTag.Number, value: 2 }),
    ).toEqual({ tag: ValueTag.Number, value: 3 });
    expect(getBuiltin("LOG")?.({ tag: ValueTag.Number, value: 100 })).toEqual({
      tag: ValueTag.Number,
      value: 2,
    });
    expect(
      getBuiltin("ACOT")?.({
        tag: ValueTag.Number,
        value: 0,
      }),
    ).toEqual({ tag: ValueTag.Number, value: Math.PI / 2 });
    expect(getBuiltin("ACOTH")?.({ tag: ValueTag.Number, value: 0.5 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });
    expect(getBuiltin("COT")?.({ tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Div0,
    });
    expect(getBuiltin("COTH")?.({ tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Div0,
    });
    expect(getBuiltin("CSC")?.({ tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Div0,
    });
    expect(getBuiltin("CSCH")?.({ tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Div0,
    });
    expect(getBuiltin("SECH")?.({ tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Number,
      value: 1,
    });
    expect(getBuiltin("SIGN")?.({ tag: ValueTag.Number, value: -42 })).toEqual({
      tag: ValueTag.Number,
      value: -1,
    });
    expect(getBuiltin("SIGN")?.({ tag: ValueTag.String, value: "bad", stringId: 12 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });
  });

  it("covers direct trig, exponential, and positive rounding builtin paths", () => {
    expect(getBuiltin("SIN")?.({ tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Number,
      value: 0,
    });
    expect(getBuiltin("COS")?.({ tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Number,
      value: 1,
    });
    expect(getBuiltin("TAN")?.({ tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Number,
      value: 0,
    });
    expect(getBuiltin("ASIN")?.({ tag: ValueTag.Number, value: 1 })).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(Math.PI / 2, 12),
    });
    expect(getBuiltin("ACOS")?.({ tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.Number,
      value: 0,
    });
    expect(getBuiltin("ATAN")?.({ tag: ValueTag.Number, value: 1 })).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(Math.PI / 4, 12),
    });
    expect(
      getBuiltin("ATAN2")?.({ tag: ValueTag.Number, value: 1 }, { tag: ValueTag.Number, value: 1 }),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(Math.PI / 4, 12),
    });
    expect(getBuiltin("DEGREES")?.({ tag: ValueTag.Number, value: Math.PI })).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(180, 12),
    });
    expect(getBuiltin("RADIANS")?.({ tag: ValueTag.Number, value: 180 })).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(Math.PI, 12),
    });
    expect(getBuiltin("EXP")?.({ tag: ValueTag.Number, value: 1 })).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(Math.E, 12),
    });
    expect(getBuiltin("LN")?.({ tag: ValueTag.Number, value: Math.E })).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(1, 12),
    });
    expect(
      getBuiltin("POWER")?.({ tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Number, value: 3 }),
    ).toEqual({ tag: ValueTag.Number, value: 8 });
    expect(getBuiltin("SQRT")?.({ tag: ValueTag.Number, value: 9 })).toEqual({
      tag: ValueTag.Number,
      value: 3,
    });
    expect(getBuiltin("PI")?.()).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(Math.PI, 12),
    });
    expect(getBuiltin("SINH")?.({ tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Number,
      value: 0,
    });
    expect(getBuiltin("COSH")?.({ tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Number,
      value: 1,
    });
    expect(getBuiltin("TANH")?.({ tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Number,
      value: 0,
    });
    expect(getBuiltin("ASINH")?.({ tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Number,
      value: 0,
    });
    expect(getBuiltin("ACOSH")?.({ tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.Number,
      value: 0,
    });
    expect(getBuiltin("ATANH")?.({ tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Number,
      value: 0,
    });

    expect(
      getBuiltin("FLOOR.MATH")?.(
        { tag: ValueTag.Number, value: 5.5 },
        { tag: ValueTag.Number, value: 2 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 4 });
    expect(
      getBuiltin("CEILING.MATH")?.(
        { tag: ValueTag.Number, value: 5.5 },
        { tag: ValueTag.Number, value: 2 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 6 });
    expect(
      getBuiltin("FLOOR.PRECISE")?.(
        { tag: ValueTag.Number, value: 4 },
        { tag: ValueTag.Number, value: 0 },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });
    expect(
      getBuiltin("CEILING.PRECISE")?.(
        { tag: ValueTag.Number, value: 4 },
        { tag: ValueTag.Number, value: 0 },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });
    expect(
      getBuiltin("BITAND")?.(
        { tag: ValueTag.String, value: "bad", stringId: 13 },
        { tag: ValueTag.Number, value: 1 },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });
    expect(
      getBuiltin("BITOR")?.(
        { tag: ValueTag.String, value: "bad", stringId: 14 },
        { tag: ValueTag.Number, value: 1 },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });
  });

  it("covers AVERAGEA and SUBTOTAL aggregate dispatch branches", () => {
    expect(
      getBuiltin("AVERAGEA")?.(
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.String, value: "skip", stringId: 15 },
        { tag: ValueTag.Boolean, value: true },
        { tag: ValueTag.Empty },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 0.75 });

    expect(
      getBuiltin("SUBTOTAL")?.(
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 4 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 3 });
    expect(
      getBuiltin("SUBTOTAL")?.(
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Boolean, value: true },
        { tag: ValueTag.String, value: "skip", stringId: 16 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 2 });
    expect(
      getBuiltin("SUBTOTAL")?.(
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Empty },
        { tag: ValueTag.String, value: "skip", stringId: 17 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 2 });
    expect(
      getBuiltin("SUBTOTAL")?.(
        { tag: ValueTag.Number, value: 4 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 4 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 4 });
    expect(
      getBuiltin("SUBTOTAL")?.(
        { tag: ValueTag.Number, value: 5 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 4 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 2 });
    expect(
      getBuiltin("SUBTOTAL")?.(
        { tag: ValueTag.Number, value: 6 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 4 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 8 });
    expect(
      getBuiltin("SUBTOTAL")?.(
        { tag: ValueTag.Number, value: 7 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 4 },
      ),
    ).toMatchObject({ tag: ValueTag.Number, value: expect.closeTo(Math.sqrt(2), 12) });
    expect(
      getBuiltin("SUBTOTAL")?.(
        { tag: ValueTag.Number, value: 8 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 4 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 1 });
    expect(
      getBuiltin("SUBTOTAL")?.(
        { tag: ValueTag.Number, value: 10 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 4 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 2 });
    expect(
      getBuiltin("SUBTOTAL")?.(
        { tag: ValueTag.Number, value: 11 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 4 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 1 });
  });

  it("covers aggregate aliases and formatting validation branches", () => {
    expect(getBuiltin("AVERAGEA")?.({ tag: ValueTag.Error, code: ErrorCode.Div0 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Div0,
    });
    expect(getBuiltin("AVERAGE")?.()).toEqual({ tag: ValueTag.Number, value: 0 });
    expect(getBuiltin("AVG")?.()).toEqual({ tag: ValueTag.Number, value: 0 });
    expect(getBuiltin("AVERAGE")?.({ tag: ValueTag.Error, code: ErrorCode.Value })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });
    expect(getBuiltin("AVG")?.({ tag: ValueTag.Error, code: ErrorCode.Name })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Name,
    });
    expect(getBuiltin("MAXA")?.({ tag: ValueTag.Error, code: ErrorCode.Ref })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    });
    expect(getBuiltin("MINA")?.({ tag: ValueTag.Error, code: ErrorCode.NA })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.NA,
    });

    expect(
      getBuiltin("ADDRESS")?.(
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 5 },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });
    expect(
      getBuiltin("ADDRESS")?.(
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Boolean, value: true },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });
    expect(
      getBuiltin("ADDRESS")?.(
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Empty },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });
    expect(
      getBuiltin("DOLLAR")?.(
        { tag: ValueTag.Number, value: Number.POSITIVE_INFINITY },
        { tag: ValueTag.Number, value: 2 },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });
    expect(
      getBuiltin("DOLLAR")?.(
        { tag: ValueTag.Number, value: 10 },
        { tag: ValueTag.Number, value: 1.5 },
      ),
    ).toEqual({ tag: ValueTag.String, value: "$10.0", stringId: 0 });
    expect(
      getBuiltin("DOLLARDE")?.(
        { tag: ValueTag.Number, value: 1.5 },
        { tag: ValueTag.Number, value: 0 },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });
    expect(
      getBuiltin("DOLLARDE")?.(
        { tag: ValueTag.Number, value: 1.6 },
        { tag: ValueTag.Number, value: 2 },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });
  });

  it("supports ADDRESS and DOLLAR formatting edge cases", () => {
    const ADDRESS = getBuiltin("ADDRESS")!;
    expect(
      ADDRESS({ tag: ValueTag.Number, value: 12 }, { tag: ValueTag.Number, value: 3 }),
    ).toEqual({
      tag: ValueTag.String,
      value: "$C$12",
      stringId: 0,
    });
    expect(
      ADDRESS(
        { tag: ValueTag.Number, value: 7 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 2 },
      ),
    ).toEqual({
      tag: ValueTag.String,
      value: "B$7",
      stringId: 0,
    });
    expect(
      ADDRESS(
        { tag: ValueTag.Number, value: 4 },
        { tag: ValueTag.Number, value: 5 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 2 },
      ),
    ).toEqual({
      tag: ValueTag.String,
      value: "R4C5",
      stringId: 0,
    });

    expect(
      getBuiltin("DOLLAR")?.(
        { tag: ValueTag.Number, value: -1234.5 },
        { tag: ValueTag.Number, value: 1 },
      ),
    ).toEqual({
      tag: ValueTag.String,
      value: "-$1,234.5",
      stringId: 0,
    });
    expect(
      getBuiltin("DOLLAR")?.(
        { tag: ValueTag.Number, value: 1234.56 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 1 },
      ),
    ).toEqual({
      tag: ValueTag.String,
      value: "$1235",
      stringId: 0,
    });
  });

  it("covers the new type, statistical, distribution, and combinatoric builtins", () => {
    const T = getBuiltin("T")!;
    const N = getBuiltin("N")!;
    const TYPE = getBuiltin("TYPE")!;
    const DELTA = getBuiltin("DELTA")!;
    const GESTEP = getBuiltin("GESTEP")!;
    const GAUSS = getBuiltin("GAUSS")!;
    const PHI = getBuiltin("PHI")!;
    const STANDARDIZE = getBuiltin("STANDARDIZE")!;
    const MODE = getBuiltin("MODE")!;
    const MODE_SNGL = getBuiltin("MODE.SNGL")!;
    const STDEV = getBuiltin("STDEV")!;
    const STDEV_S = getBuiltin("STDEV.S")!;
    const STDEVP = getBuiltin("STDEVP")!;
    const STDEV_P = getBuiltin("STDEV.P")!;
    const STDEVA = getBuiltin("STDEVA")!;
    const STDEVPA = getBuiltin("STDEVPA")!;
    const VAR = getBuiltin("VAR")!;
    const VAR_S = getBuiltin("VAR.S")!;
    const VARP = getBuiltin("VARP")!;
    const VAR_P = getBuiltin("VAR.P")!;
    const VARA = getBuiltin("VARA")!;
    const VARPA = getBuiltin("VARPA")!;
    const SKEW = getBuiltin("SKEW")!;
    const SKEW_P = getBuiltin("SKEW.P")!;
    const KURT = getBuiltin("KURT")!;
    const NORMDIST = getBuiltin("NORMDIST")!;
    const NORM_DIST = getBuiltin("NORM.DIST")!;
    const NORMINV = getBuiltin("NORMINV")!;
    const NORM_INV = getBuiltin("NORM.INV")!;
    const NORMSDIST = getBuiltin("NORMSDIST")!;
    const NORM_S_DIST = getBuiltin("NORM.S.DIST")!;
    const NORMSINV = getBuiltin("NORMSINV")!;
    const NORM_S_INV = getBuiltin("NORM.S.INV")!;
    const LOGINV = getBuiltin("LOGINV")!;
    const LOGNORMDIST = getBuiltin("LOGNORMDIST")!;
    const LOGNORM_DIST = getBuiltin("LOGNORM.DIST")!;
    const LOGNORM_INV = getBuiltin("LOGNORM.INV")!;
    const CONFIDENCE_NORM = getBuiltin("CONFIDENCE.NORM")!;
    const PERMUT = getBuiltin("PERMUT")!;
    const PERMUTATIONA = getBuiltin("PERMUTATIONA")!;

    expect(T({ tag: ValueTag.String, value: "alpha", stringId: 1 })).toEqual({
      tag: ValueTag.String,
      value: "alpha",
      stringId: 1,
    });
    expect(T({ tag: ValueTag.Number, value: 42 })).toEqual({ tag: ValueTag.Empty });
    expect(T({ tag: ValueTag.Error, code: ErrorCode.Ref })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    });

    expect(N({ tag: ValueTag.Boolean, value: true })).toEqual({ tag: ValueTag.Number, value: 1 });
    expect(N({ tag: ValueTag.String, value: "alpha", stringId: 2 })).toEqual({
      tag: ValueTag.Number,
      value: 0,
    });
    expect(N({ tag: ValueTag.Error, code: ErrorCode.Name })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Name,
    });

    expect(TYPE({ tag: ValueTag.Number, value: 1 })).toEqual({ tag: ValueTag.Number, value: 1 });
    expect(TYPE({ tag: ValueTag.String, value: "alpha", stringId: 3 })).toEqual({
      tag: ValueTag.Number,
      value: 2,
    });
    expect(TYPE({ tag: ValueTag.Boolean, value: true })).toEqual({
      tag: ValueTag.Number,
      value: 4,
    });
    expect(TYPE({ tag: ValueTag.Error, code: ErrorCode.Value })).toEqual({
      tag: ValueTag.Number,
      value: 16,
    });
    const arrayValue: ArrayValue = {
      kind: "array",
      rows: 1,
      cols: 1,
      values: [{ tag: ValueTag.Number, value: 1 }],
    };
    expect(TYPE(arrayValue)).toEqual({
      tag: ValueTag.Number,
      value: 64,
    });

    expect(DELTA({ tag: ValueTag.Number, value: 4 }, { tag: ValueTag.Number, value: 4 })).toEqual({
      tag: ValueTag.Number,
      value: 1,
    });
    expect(DELTA({ tag: ValueTag.Number, value: 4 })).toEqual({ tag: ValueTag.Number, value: 0 });
    expect(DELTA({ tag: ValueTag.String, value: "bad", stringId: 4 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });

    expect(GESTEP({ tag: ValueTag.Number, value: 4 }, { tag: ValueTag.Number, value: 2 })).toEqual({
      tag: ValueTag.Number,
      value: 1,
    });
    expect(GESTEP({ tag: ValueTag.Number, value: -1 })).toEqual({ tag: ValueTag.Number, value: 0 });
    expect(
      GESTEP(
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.String, value: "bad", stringId: 5 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });

    expect(GAUSS({ tag: ValueTag.Number, value: 0 })).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0, 8),
    });
    expect(PHI({ tag: ValueTag.Number, value: 0 })).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(1 / Math.sqrt(2 * Math.PI), 12),
    });
    expect(
      STANDARDIZE(
        { tag: ValueTag.Number, value: 42 },
        { tag: ValueTag.Number, value: 40 },
        { tag: ValueTag.Number, value: 2 },
      ),
    ).toEqual({
      tag: ValueTag.Number,
      value: 1,
    });
    expect(
      STANDARDIZE(
        { tag: ValueTag.Number, value: 42 },
        { tag: ValueTag.Number, value: 40 },
        { tag: ValueTag.Number, value: 0 },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });

    expect(
      MODE(
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 2 },
      ),
    ).toEqual({
      tag: ValueTag.Number,
      value: 3,
    });
    expect(
      MODE(
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 1 },
      ),
    ).toEqual({
      tag: ValueTag.Number,
      value: 1,
    });
    expect(
      MODE_SNGL(
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 3 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.NA,
    });
    expect(MODE({ tag: ValueTag.Error, code: ErrorCode.Ref })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    });
    expect(MODE_SNGL({ tag: ValueTag.Error, code: ErrorCode.Div0 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Div0,
    });

    expect(
      STDEV(
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 4 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(Math.sqrt(5 / 3), 12),
    });
    expect(
      STDEV_S(
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Boolean, value: true },
        { tag: ValueTag.String, value: "skip", stringId: 6 },
      ),
    ).toEqual({
      tag: ValueTag.Number,
      value: 0,
    });
    expect(
      STDEVP(
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 4 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(Math.sqrt(1.25), 12),
    });
    expect(STDEV_P({ tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Number, value: 2 })).toEqual(
      {
        tag: ValueTag.Number,
        value: 0,
      },
    );
    expect(
      STDEVA(
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Boolean, value: true },
        { tag: ValueTag.String, value: "skip", stringId: 7 },
      ),
    ).toEqual({
      tag: ValueTag.Number,
      value: 1,
    });
    expect(
      STDEVPA(
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Boolean, value: true },
        { tag: ValueTag.String, value: "skip", stringId: 8 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(Math.sqrt(2 / 3), 12),
    });
    expect(STDEV({ tag: ValueTag.Error, code: ErrorCode.Name })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Name,
    });
    expect(STDEV_S({ tag: ValueTag.Error, code: ErrorCode.Ref })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    });
    expect(STDEVP({ tag: ValueTag.Error, code: ErrorCode.Value })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });
    expect(STDEV_P({ tag: ValueTag.Error, code: ErrorCode.Div0 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Div0,
    });
    expect(STDEVA({ tag: ValueTag.Error, code: ErrorCode.Num })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    });
    expect(STDEVPA({ tag: ValueTag.Error, code: ErrorCode.NA })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.NA,
    });

    expect(
      VAR(
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 4 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(5 / 3, 12),
    });
    expect(
      VAR_S({ tag: ValueTag.Number, value: 1 }, { tag: ValueTag.Number, value: 2 }),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.5, 12),
    });
    expect(
      VARP(
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 4 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(1.25, 12),
    });
    expect(VAR_P({ tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Number, value: 2 })).toEqual({
      tag: ValueTag.Number,
      value: 0,
    });
    expect(
      VARA(
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Boolean, value: true },
        { tag: ValueTag.String, value: "skip", stringId: 9 },
      ),
    ).toEqual({
      tag: ValueTag.Number,
      value: 1,
    });
    expect(
      VARPA(
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Boolean, value: true },
        { tag: ValueTag.String, value: "skip", stringId: 10 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(2 / 3, 12),
    });
    expect(VAR({ tag: ValueTag.Error, code: ErrorCode.Name })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Name,
    });
    expect(VAR_S({ tag: ValueTag.Error, code: ErrorCode.Ref })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    });
    expect(VARP({ tag: ValueTag.Error, code: ErrorCode.Value })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });
    expect(VAR_P({ tag: ValueTag.Error, code: ErrorCode.Div0 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Div0,
    });
    expect(VARA({ tag: ValueTag.Error, code: ErrorCode.Num })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    });
    expect(VARPA({ tag: ValueTag.Error, code: ErrorCode.NA })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.NA,
    });

    expect(
      SKEW(
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 4 },
        { tag: ValueTag.Number, value: 5 },
        { tag: ValueTag.Number, value: 6 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0, 12),
    });
    expect(
      SKEW_P(
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 4 },
        { tag: ValueTag.Number, value: 5 },
        { tag: ValueTag.Number, value: 6 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0, 12),
    });
    expect(
      KURT(
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 4 },
        { tag: ValueTag.Number, value: 5 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(-1.2, 12),
    });
    expect(
      KURT(
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 3 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });
    expect(SKEW({ tag: ValueTag.Error, code: ErrorCode.Name })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Name,
    });
    expect(SKEW_P({ tag: ValueTag.Error, code: ErrorCode.Ref })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    });
    expect(KURT({ tag: ValueTag.Error, code: ErrorCode.Div0 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Div0,
    });

    expect(
      NORMDIST(
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Boolean, value: true },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.8413447460685429, 7),
    });
    expect(
      NORM_DIST(
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Boolean, value: false },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.24197072451914337, 12),
    });
    expect(
      NORMINV(
        { tag: ValueTag.Number, value: 0.8413447460685429 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 1 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(1, 8),
    });
    expect(
      NORM_INV(
        { tag: ValueTag.Number, value: 0.8413447460685429 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 1 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(1, 8),
    });
    expect(NORMSDIST({ tag: ValueTag.Number, value: 1 })).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.8413447460685429, 7),
    });
    expect(
      NORM_S_DIST({ tag: ValueTag.Number, value: 1 }, { tag: ValueTag.Boolean, value: false }),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.24197072451914337, 12),
    });
    expect(NORMSINV({ tag: ValueTag.Number, value: 0.001 })).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(-3.090232306167813, 8),
    });
    expect(NORM_S_INV({ tag: ValueTag.Number, value: 0.999 })).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(3.090232306167813, 8),
    });
    expect(NORMSINV({ tag: ValueTag.Number, value: 0.5 })).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0, 12),
    });
    expect(
      NORMINV(
        { tag: ValueTag.String, value: "bad", stringId: 11 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 1 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });
    expect(
      NORM_S_DIST(
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.String, value: "bad", stringId: 12 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });
    expect(NORMSINV({ tag: ValueTag.String, value: "bad", stringId: 13 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });
    expect(
      LOGINV(
        { tag: ValueTag.Number, value: 0.5 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 0 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });
    expect(
      LOGINV(
        { tag: ValueTag.Number, value: 0.5 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 1 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(1, 12),
    });
    expect(
      LOGNORMDIST(
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 1 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.5, 8),
    });
    expect(
      LOGNORM_DIST(
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Boolean, value: false },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(1 / Math.sqrt(2 * Math.PI), 12),
    });
    expect(
      LOGNORM_INV(
        { tag: ValueTag.Number, value: 0.5 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 1 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(1, 12),
    });
    expect(
      CONFIDENCE_NORM(
        { tag: ValueTag.Number, value: 0.05 },
        { tag: ValueTag.Number, value: 1.5 },
        { tag: ValueTag.Number, value: 100 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.2939945976810081, 9),
    });
    expect(
      CONFIDENCE_NORM(
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 1.5 },
        { tag: ValueTag.Number, value: 100 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });
    expect(
      LOGNORMDIST(
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 1 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });

    expect(PERMUT({ tag: ValueTag.Number, value: 5 }, { tag: ValueTag.Number, value: 3 })).toEqual({
      tag: ValueTag.Number,
      value: 60,
    });
    expect(
      PERMUTATIONA({ tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Number, value: 3 }),
    ).toEqual({
      tag: ValueTag.Number,
      value: 8,
    });
    expect(
      PERMUTATIONA(
        { tag: ValueTag.String, value: "bad", stringId: 14 },
        { tag: ValueTag.Number, value: 3 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });
    expect(PERMUT({ tag: ValueTag.Number, value: 3 }, { tag: ValueTag.Number, value: 4 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });

    expect(getBuiltinId("norm.dist")).toBe(BuiltinId.NormDist);
    expect(getBuiltinId("norm.s.inv")).toBe(BuiltinId.NormSInv);
    expect(getBuiltinId("confidence.norm")).toBe(BuiltinId.ConfidenceNorm);
    expect(getBuiltinId("permutationa")).toBe(BuiltinId.Permutationa);
  });

  it("supports the new statistical distribution builtins and aliases", () => {
    const ERF = getBuiltin("ERF")!;
    const ERF_PRECISE = getBuiltin("ERF.PRECISE")!;
    const ERFC = getBuiltin("ERFC")!;
    const ERFC_PRECISE = getBuiltin("ERFC.PRECISE")!;
    const FISHER = getBuiltin("FISHER")!;
    const FISHERINV = getBuiltin("FISHERINV")!;
    const GAMMALN = getBuiltin("GAMMALN")!;
    const GAMMALN_PRECISE = getBuiltin("GAMMALN.PRECISE")!;
    const GAMMA = getBuiltin("GAMMA")!;
    const CONFIDENCE = getBuiltin("CONFIDENCE")!;
    const EXPONDIST = getBuiltin("EXPONDIST")!;
    const EXPON_DIST = getBuiltin("EXPON.DIST")!;
    const POISSON = getBuiltin("POISSON")!;
    const POISSON_DIST = getBuiltin("POISSON.DIST")!;
    const WEIBULL = getBuiltin("WEIBULL")!;
    const WEIBULL_DIST = getBuiltin("WEIBULL.DIST")!;
    const GAMMADIST = getBuiltin("GAMMADIST")!;
    const GAMMA_DIST = getBuiltin("GAMMA.DIST")!;
    const CHIDIST = getBuiltin("CHIDIST")!;
    const CHISQ_DIST_RT = getBuiltin("CHISQ.DIST.RT")!;
    const CHISQ_DIST = getBuiltin("CHISQ.DIST")!;
    const BINOMDIST = getBuiltin("BINOMDIST")!;
    const BINOM_DIST = getBuiltin("BINOM.DIST")!;
    const BINOM_DIST_RANGE = getBuiltin("BINOM.DIST.RANGE")!;
    const CRITBINOM = getBuiltin("CRITBINOM")!;
    const BINOM_INV = getBuiltin("BINOM.INV")!;
    const HYPGEOMDIST = getBuiltin("HYPGEOMDIST")!;
    const HYPGEOM_DIST = getBuiltin("HYPGEOM.DIST")!;
    const NEGBINOMDIST = getBuiltin("NEGBINOMDIST")!;
    const NEGBINOM_DIST = getBuiltin("NEGBINOM.DIST")!;

    expect(ERF({ tag: ValueTag.Number, value: 1 })).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.8427006897475899, 7),
    });
    expect(
      ERF({ tag: ValueTag.Number, value: 0 }, { tag: ValueTag.Number, value: 1 }),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.8427006897475899, 7),
    });
    expect(ERF_PRECISE({ tag: ValueTag.Number, value: 1 })).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.8427006897475899, 7),
    });
    expect(ERFC({ tag: ValueTag.Number, value: 1 })).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.15729931025241006, 7),
    });
    expect(ERFC_PRECISE({ tag: ValueTag.Number, value: 1 })).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.15729931025241006, 7),
    });
    expect(ERF({ tag: ValueTag.String, value: "bad", stringId: 15 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });
    expect(FISHER({ tag: ValueTag.Number, value: 0.5 })).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.5493061443340549, 12),
    });
    expect(FISHERINV({ tag: ValueTag.Number, value: 0.5493061443340549 })).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.5, 12),
    });
    expect(FISHERINV({ tag: ValueTag.String, value: "bad", stringId: 16 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });
    expect(GAMMALN({ tag: ValueTag.Number, value: 5 })).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(Math.log(24), 12),
    });
    expect(GAMMALN_PRECISE({ tag: ValueTag.Number, value: 5 })).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(Math.log(24), 12),
    });
    expect(GAMMA({ tag: ValueTag.Number, value: 5 })).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(24, 10),
    });
    expect(
      CONFIDENCE(
        { tag: ValueTag.Number, value: 0.05 },
        { tag: ValueTag.Number, value: 1.5 },
        { tag: ValueTag.Number, value: 100 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.2939945976810081, 9),
    });
    expect(
      EXPONDIST(
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Boolean, value: false },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.2706705664732254, 12),
    });
    expect(
      EXPON_DIST(
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Boolean, value: true },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.8646647167633873, 12),
    });
    expect(
      EXPONDIST(
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Boolean, value: false },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });
    expect(
      POISSON(
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 2.5 },
        { tag: ValueTag.Boolean, value: false },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.21376301724973648, 12),
    });
    expect(
      POISSON_DIST(
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 2.5 },
        { tag: ValueTag.Boolean, value: true },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.7575761331330662, 12),
    });
    expect(
      POISSON(
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: -1 },
        { tag: ValueTag.Boolean, value: false },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });
    expect(
      WEIBULL(
        { tag: ValueTag.Number, value: 1.5 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Boolean, value: false },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.2596002610238016, 12),
    });
    expect(
      WEIBULL_DIST(
        { tag: ValueTag.Number, value: 1.5 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Boolean, value: true },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.22119921692859512, 12),
    });
    expect(
      WEIBULL(
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 0.5 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Boolean, value: false },
      ),
    ).toEqual({ tag: ValueTag.Number, value: Number.POSITIVE_INFINITY });
    expect(
      WEIBULL(
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Boolean, value: false },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });
    expect(
      GAMMADIST(
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Boolean, value: false },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.09196986029286061, 12),
    });
    expect(
      GAMMA_DIST(
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Boolean, value: true },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.08030139707139418, 12),
    });
    expect(
      HYPGEOM_DIST(
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 4 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 10 },
        { tag: ValueTag.Boolean, value: false },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.5, 12),
    });
    expect(
      NEGBINOM_DIST(
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 0.5 },
        { tag: ValueTag.Boolean, value: false },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.1875, 12),
    });
    expect(
      CHIDIST({ tag: ValueTag.Number, value: 3 }, { tag: ValueTag.Number, value: 4 }),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.5578254003710748, 12),
    });
    expect(
      CHISQ_DIST_RT({ tag: ValueTag.Number, value: 3 }, { tag: ValueTag.Number, value: 4 }),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.5578254003710748, 12),
    });
    expect(
      CHISQ_DIST(
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 4 },
        { tag: ValueTag.Boolean, value: true },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.4421745996289252, 12),
    });
    expect(
      BINOMDIST(
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 4 },
        { tag: ValueTag.Number, value: 0.5 },
        { tag: ValueTag.Boolean, value: false },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.375, 12),
    });
    expect(
      BINOM_DIST(
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 4 },
        { tag: ValueTag.Number, value: 0.5 },
        { tag: ValueTag.Boolean, value: true },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.6875, 12),
    });
    expect(
      BINOM_DIST_RANGE(
        { tag: ValueTag.Number, value: 6 },
        { tag: ValueTag.Number, value: 0.5 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 4 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.78125, 12),
    });
    expect(
      CRITBINOM(
        { tag: ValueTag.Number, value: 6 },
        { tag: ValueTag.Number, value: 0.5 },
        { tag: ValueTag.Number, value: 0.7 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 4 });
    expect(
      CRITBINOM(
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 0.5 },
        { tag: ValueTag.Number, value: 0.999999999999 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 1 });
    expect(
      BINOM_INV(
        { tag: ValueTag.Number, value: 6 },
        { tag: ValueTag.Number, value: 0.5 },
        { tag: ValueTag.Number, value: 0.7 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 4 });
    expect(
      HYPGEOMDIST(
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 4 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 10 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.5, 12),
    });
    expect(
      HYPGEOM_DIST(
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 4 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 10 },
        { tag: ValueTag.Boolean, value: true },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(2 / 3, 12),
    });
    expect(
      NEGBINOMDIST(
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 0.5 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.1875, 12),
    });
    expect(
      NEGBINOM_DIST(
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 0.5 },
        { tag: ValueTag.Boolean, value: true },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.5, 12),
    });

    expect(FISHER({ tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });
    expect(
      GAMMADIST(
        { tag: ValueTag.Number, value: -1 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Boolean, value: false },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });
    expect(
      GAMMA_DIST(
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Boolean, value: true },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });
    expect(GAMMA({ tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });
    expect(
      CHIDIST({ tag: ValueTag.Number, value: -1 }, { tag: ValueTag.Number, value: 4 }),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });
    expect(
      CHISQ_DIST(
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Boolean, value: true },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });
    expect(
      BINOMDIST(
        { tag: ValueTag.Number, value: 5 },
        { tag: ValueTag.Number, value: 4 },
        { tag: ValueTag.Number, value: 0.5 },
        { tag: ValueTag.Boolean, value: false },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });
    expect(
      BINOM_DIST_RANGE(
        { tag: ValueTag.Number, value: 4 },
        { tag: ValueTag.Number, value: 0.5 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 2 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });
    expect(
      CRITBINOM(
        { tag: ValueTag.Number, value: 6 },
        { tag: ValueTag.Number, value: 0.5 },
        { tag: ValueTag.Number, value: 1 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });
    expect(
      HYPGEOMDIST(
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 4 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 0 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });
    expect(
      HYPGEOM_DIST(
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 4 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 10 },
        { tag: ValueTag.String, value: "bad", stringId: 1 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });
    expect(
      NEGBINOMDIST(
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 1.5 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });
    expect(
      NEGBINOM_DIST(
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 0.5 },
        { tag: ValueTag.String, value: "bad", stringId: 1 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });
    expect(getBuiltinId("gamma.dist")).toBe(BuiltinId.GammaDist);
    expect(getBuiltinId("negbinom.dist")).toBe(BuiltinId.NegbinomDist);
    expect(getBuiltinId("binom.inv")).toBe(BuiltinId.BinomInv);
  });

  it("covers the new financial builtins and their error branches", () => {
    const EFFECT = getBuiltin("EFFECT")!;
    const NOMINAL = getBuiltin("NOMINAL")!;
    const PDURATION = getBuiltin("PDURATION")!;
    const RRI = getBuiltin("RRI")!;
    const FV = getBuiltin("FV")!;
    const PV = getBuiltin("PV")!;
    const PMT = getBuiltin("PMT")!;
    const NPER = getBuiltin("NPER")!;
    const NPV = getBuiltin("NPV")!;
    const IPMT = getBuiltin("IPMT")!;
    const PPMT = getBuiltin("PPMT")!;
    const ISPMT = getBuiltin("ISPMT")!;

    expect(
      EFFECT({ tag: ValueTag.Number, value: 0.12 }, { tag: ValueTag.Number, value: 12 }),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.12682503013196977, 12),
    });
    expect(
      NOMINAL(
        { tag: ValueTag.Number, value: 0.12682503013196977 },
        { tag: ValueTag.Number, value: 12 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.12, 12),
    });
    expect(
      PDURATION(
        { tag: ValueTag.Number, value: 0.1 },
        { tag: ValueTag.Number, value: 100 },
        { tag: ValueTag.Number, value: 121 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(2, 12),
    });
    expect(
      RRI(
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 100 },
        { tag: ValueTag.Number, value: 121 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.1, 12),
    });

    expect(
      FV(
        { tag: ValueTag.Number, value: 0.1 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: -100 },
        { tag: ValueTag.Number, value: -1000 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(1420, 12),
    });
    expect(
      PV(
        { tag: ValueTag.Number, value: 0.1 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: -100 },
        { tag: ValueTag.Number, value: 1420 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(-1000, 12),
    });
    expect(
      PMT(
        { tag: ValueTag.Number, value: 0.1 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 1000 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(-576.1904761904761, 12),
    });
    expect(
      NPER(
        { tag: ValueTag.Number, value: 0.1 },
        { tag: ValueTag.Number, value: -576.1904761904761 },
        { tag: ValueTag.Number, value: 1000 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(2, 12),
    });
    expect(
      NPV(
        { tag: ValueTag.Number, value: 0.1 },
        { tag: ValueTag.Number, value: 100 },
        { tag: ValueTag.Number, value: 100 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(173.55371900826447, 12),
    });
    expect(
      IPMT(
        { tag: ValueTag.String, value: "bad", stringId: 17 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 1000 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });
    expect(
      IPMT(
        { tag: ValueTag.Number, value: 0.1 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 1000 },
      ),
    ).toEqual({
      tag: ValueTag.Number,
      value: -100,
    });
    expect(
      IPMT(
        { tag: ValueTag.Number, value: 0.1 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 1000 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 1 },
      ),
    ).toEqual({
      tag: ValueTag.Number,
      value: 0,
    });
    expect(
      PPMT(
        { tag: ValueTag.Number, value: 0.1 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 1000 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(-476.19047619047615, 12),
    });
    expect(
      ISPMT(
        { tag: ValueTag.Number, value: 0.1 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 1000 },
      ),
    ).toEqual({
      tag: ValueTag.Number,
      value: -50,
    });

    expect(
      EFFECT({ tag: ValueTag.Number, value: 0.1 }, { tag: ValueTag.Number, value: 0 }),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });
    expect(
      NOMINAL({ tag: ValueTag.Number, value: -1 }, { tag: ValueTag.Number, value: 12 }),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });
    expect(
      PDURATION(
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 100 },
        { tag: ValueTag.Number, value: 121 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });
    expect(
      RRI(
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 100 },
        { tag: ValueTag.Number, value: 121 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });
    expect(
      PMT(
        { tag: ValueTag.Number, value: 0.1 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 1000 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });
    expect(
      NPER(
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 1000 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });
    expect(
      NPV(
        { tag: ValueTag.String, value: "bad", stringId: 11 },
        { tag: ValueTag.Number, value: 100 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });
    expect(
      IPMT(
        { tag: ValueTag.Number, value: 0.1 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 1000 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });
    expect(
      PPMT(
        { tag: ValueTag.Number, value: 0.1 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 1000 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 2 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });
    expect(
      ISPMT(
        { tag: ValueTag.Number, value: 0.1 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 1000 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });
  });

  it("registers protocol-declared placeholder builtins as blocked", () => {
    for (const name of placeholderBuiltinNames) {
      expect(getBuiltin(name)?.()).toEqual({ tag: ValueTag.Error, code: ErrorCode.Blocked });
    }

    for (const name of protocolPlaceholderBuiltinNames) {
      expect(getBuiltinId(name.toLowerCase())).toBeDefined();
    }

    expect(getBuiltinId("sin")).toBe(BuiltinId.Sin);
    expect(getBuiltinId("weeknum")).toBe(BuiltinId.Weeknum);
    expect(getBuiltinId("rept")).toBe(BuiltinId.Rept);
    expect(getBuiltinId("filter")).toBe(BuiltinId.Filter);
    expect(getBuiltinId("let")).toBeUndefined();
    expect(getBuiltinId("textjoin")).toBeUndefined();
  });
});
