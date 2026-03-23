import { describe, expect, it } from "vitest";
import { BuiltinId, ErrorCode, ValueTag } from "@bilig/protocol";
import { getBuiltin, getBuiltinId } from "../builtins.js";
import { placeholderBuiltinNames, protocolPlaceholderBuiltinNames } from "../builtins/placeholder.js";

describe("formula builtins", () => {
  it("supports numeric aggregates and error propagation", () => {
    const sum = getBuiltin("SUM");
    const avg = getBuiltin("AVG");
    const mod = getBuiltin("MOD");

    expect(sum?.(
      { tag: ValueTag.Number, value: 2 },
      { tag: ValueTag.Boolean, value: true },
      { tag: ValueTag.Empty }
    )).toEqual({ tag: ValueTag.Number, value: 3 });

    expect(avg?.(
      { tag: ValueTag.Number, value: 2 },
      { tag: ValueTag.String, value: "skip", stringId: 1 },
      { tag: ValueTag.Empty }
    )).toEqual({ tag: ValueTag.Number, value: 1 });

    expect(sum?.(
      { tag: ValueTag.Number, value: 2 },
      { tag: ValueTag.Error, code: ErrorCode.Ref }
    )).toEqual({ tag: ValueTag.Error, code: ErrorCode.Ref });

    expect(mod?.(
      { tag: ValueTag.Number, value: 10 },
      { tag: ValueTag.Number, value: 0 }
    )).toEqual({ tag: ValueTag.Error, code: ErrorCode.Div0 });
  });

  it("supports boolean and string builtins and builtin ids", () => {
    expect(getBuiltin("AND")?.(
      { tag: ValueTag.Number, value: 1 },
      { tag: ValueTag.Empty }
    )).toEqual({ tag: ValueTag.Boolean, value: false });

    expect(getBuiltin("OR")?.(
      { tag: ValueTag.Empty },
      { tag: ValueTag.Boolean, value: true }
    )).toEqual({ tag: ValueTag.Boolean, value: true });

    expect(getBuiltin("NOT")?.({ tag: ValueTag.Boolean, value: false })).toEqual({
      tag: ValueTag.Boolean,
      value: true
    });

    expect(getBuiltin("LEN")?.({ tag: ValueTag.Boolean, value: true })).toEqual({
      tag: ValueTag.Number,
      value: 4
    });

    expect(getBuiltin("CONCAT")?.(
      { tag: ValueTag.String, value: "alpha", stringId: 1 },
      { tag: ValueTag.Number, value: 2 },
      { tag: ValueTag.Empty }
    )).toEqual({ tag: ValueTag.String, value: "alpha2", stringId: 0 });

    expect(getBuiltin("EXACT")?.(
      { tag: ValueTag.String, value: "Alpha", stringId: 1 },
      { tag: ValueTag.String, value: "Alpha", stringId: 2 }
    )).toEqual({ tag: ValueTag.Boolean, value: true });

    expect(getBuiltin("EXACT")?.(
      { tag: ValueTag.String, value: "Alpha", stringId: 1 },
      { tag: ValueTag.String, value: "alpha", stringId: 2 }
    )).toEqual({ tag: ValueTag.Boolean, value: false });

    expect(getBuiltin("LEFT")?.(
      { tag: ValueTag.String, value: "alpha", stringId: 1 },
      { tag: ValueTag.Number, value: 3 }
    )).toEqual({ tag: ValueTag.String, value: "alp", stringId: 0 });

    expect(getBuiltin("TEXTBEFORE")?.(
      { tag: ValueTag.String, value: "alpha-beta", stringId: 1 },
      { tag: ValueTag.String, value: "-", stringId: 2 }
    )).toEqual({ tag: ValueTag.String, value: "alpha", stringId: 0 });

    expect(getBuiltin("IFERROR")?.(
      { tag: ValueTag.Error, code: ErrorCode.Div0 },
      { tag: ValueTag.String, value: "fallback", stringId: 1 }
    )).toEqual({ tag: ValueTag.String, value: "fallback", stringId: 1 });

    expect(getBuiltin("DATE")?.(
      { tag: ValueTag.Number, value: 2026 },
      { tag: ValueTag.Number, value: 3 },
      { tag: ValueTag.Number, value: 15 }
    )).toEqual({ tag: ValueTag.Number, value: 46096 });

    expect(getBuiltin("AVERAGE")?.(
      { tag: ValueTag.Number, value: 2 },
      { tag: ValueTag.Number, value: 4 },
      { tag: ValueTag.Number, value: 6 }
    )).toEqual({ tag: ValueTag.Number, value: 4 });

    expect(getBuiltinId("sum")).toBe(BuiltinId.Sum);
    expect(getBuiltinId("concat")).toBe(BuiltinId.Concat);
    expect(getBuiltinId("")).toBeUndefined();
    expect(getBuiltin("missing")).toBeUndefined();
  });

  it("supports the remaining scalar numeric builtins and conditional defaults", () => {
    expect(getBuiltin("MIN")?.(
      { tag: ValueTag.Empty },
      { tag: ValueTag.Number, value: 3 },
      { tag: ValueTag.Number, value: -1 }
    )).toEqual({ tag: ValueTag.Number, value: -1 });

    expect(getBuiltin("MAX")?.(
      { tag: ValueTag.String, value: "skip", stringId: 1 },
      { tag: ValueTag.Number, value: 3 },
      { tag: ValueTag.Boolean, value: true }
    )).toEqual({ tag: ValueTag.Number, value: 3 });

    expect(getBuiltin("COUNT")?.(
      { tag: ValueTag.Number, value: 3 },
      { tag: ValueTag.Boolean, value: false },
      { tag: ValueTag.String, value: "skip", stringId: 1 }
    )).toEqual({ tag: ValueTag.Number, value: 2 });

    expect(getBuiltin("COUNTA")?.(
      { tag: ValueTag.Empty },
      { tag: ValueTag.String, value: "x", stringId: 1 },
      { tag: ValueTag.Boolean, value: false }
    )).toEqual({ tag: ValueTag.Number, value: 2 });

    expect(getBuiltin("ABS")?.({ tag: ValueTag.Number, value: -3.4 })).toEqual({
      tag: ValueTag.Number,
      value: 3.4
    });
    expect(getBuiltin("INT")?.({ tag: ValueTag.Number, value: -3.1 })).toEqual({
      tag: ValueTag.Number,
      value: -4
    });
    expect(getBuiltin("ROUND")?.({ tag: ValueTag.Number, value: 3.6 })).toEqual({
      tag: ValueTag.Number,
      value: 4
    });
    expect(getBuiltin("ROUNDUP")?.(
      { tag: ValueTag.Number, value: 3.145 },
      { tag: ValueTag.Number, value: 2 }
    )).toEqual({ tag: ValueTag.Number, value: 3.15 });
    expect(getBuiltin("ROUNDDOWN")?.(
      { tag: ValueTag.Number, value: -3.145 },
      { tag: ValueTag.Number, value: 2 }
    )).toEqual({ tag: ValueTag.Number, value: -3.14 });
    expect(getBuiltin("ROUND")?.(
      { tag: ValueTag.Number, value: 3.145 },
      { tag: ValueTag.Number, value: 2 }
    )).toEqual({ tag: ValueTag.Number, value: 3.15 });
    expect(getBuiltin("FLOOR")?.({ tag: ValueTag.Number, value: 3.6 })).toEqual({
      tag: ValueTag.Number,
      value: 3
    });
    expect(getBuiltin("FLOOR")?.(
      { tag: ValueTag.Number, value: 7 },
      { tag: ValueTag.Number, value: 2 }
    )).toEqual({ tag: ValueTag.Number, value: 6 });
    expect(getBuiltin("CEILING")?.({ tag: ValueTag.Number, value: 3.1 })).toEqual({
      tag: ValueTag.Number,
      value: 4
    });
    expect(getBuiltin("CEILING")?.(
      { tag: ValueTag.Number, value: 7 },
      { tag: ValueTag.Number, value: 2 }
    )).toEqual({ tag: ValueTag.Number, value: 8 });

    expect(getBuiltin("IF")?.(
      { tag: ValueTag.Number, value: 1 },
      { tag: ValueTag.String, value: "truthy", stringId: 1 }
    )).toEqual({ tag: ValueTag.String, value: "truthy", stringId: 1 });

    expect(getBuiltin("IF")?.(
      { tag: ValueTag.Empty },
      { tag: ValueTag.Number, value: 1 }
    )).toEqual({ tag: ValueTag.Empty });
  });

  it("supports expanded math and numeric utility builtins", () => {
    expect(getBuiltin("SIN")?.({ tag: ValueTag.Number, value: Math.PI / 2 })).toEqual({ tag: ValueTag.Number, value: 1 });
    expect(getBuiltin("COS")?.({ tag: ValueTag.Number, value: 0 })).toEqual({ tag: ValueTag.Number, value: 1 });
    expect(getBuiltin("POWER")?.(
      { tag: ValueTag.Number, value: 2 },
      { tag: ValueTag.Number, value: 3 }
    )).toEqual({ tag: ValueTag.Number, value: 8 });
    expect(getBuiltin("LOG")?.({ tag: ValueTag.Number, value: 1000 })).toEqual({ tag: ValueTag.Number, value: 3 });
    expect(getBuiltin("SIGN")?.({ tag: ValueTag.Number, value: -9 })).toEqual({ tag: ValueTag.Number, value: -1 });
    expect(getBuiltin("TRUNC")?.(
      { tag: ValueTag.Number, value: -3.98 },
      { tag: ValueTag.Number, value: 1 }
    )).toEqual({ tag: ValueTag.Number, value: -3.9 });
    expect(getBuiltin("CEILING.MATH")?.(
      { tag: ValueTag.Number, value: -5.5 },
      { tag: ValueTag.Number, value: 2 }
    )).toEqual({ tag: ValueTag.Number, value: -4 });
    expect(getBuiltin("FLOOR.PRECISE")?.(
      { tag: ValueTag.Number, value: -5.5 },
      { tag: ValueTag.Number, value: 2 }
    )).toEqual({ tag: ValueTag.Number, value: -6 });
    expect(getBuiltin("FACT")?.({ tag: ValueTag.Number, value: 5 })).toEqual({ tag: ValueTag.Number, value: 120 });
    expect(getBuiltin("COMBIN")?.(
      { tag: ValueTag.Number, value: 5 },
      { tag: ValueTag.Number, value: 2 }
    )).toEqual({ tag: ValueTag.Number, value: 10 });
    expect(getBuiltin("GCD")?.(
      { tag: ValueTag.Number, value: 18 },
      { tag: ValueTag.Number, value: 24 }
    )).toEqual({ tag: ValueTag.Number, value: 6 });
    expect(getBuiltin("LCM")?.(
      { tag: ValueTag.Number, value: 6 },
      { tag: ValueTag.Number, value: 8 }
    )).toEqual({ tag: ValueTag.Number, value: 24 });
    expect(getBuiltin("MROUND")?.(
      { tag: ValueTag.Number, value: 10 },
      { tag: ValueTag.Number, value: 3 }
    )).toEqual({ tag: ValueTag.Number, value: 9 });
    expect(getBuiltin("PRODUCT")?.(
      { tag: ValueTag.Number, value: 2 },
      { tag: ValueTag.Number, value: 3 },
      { tag: ValueTag.Number, value: 4 }
    )).toEqual({ tag: ValueTag.Number, value: 24 });
    expect(getBuiltin("QUOTIENT")?.(
      { tag: ValueTag.Number, value: 7 },
      { tag: ValueTag.Number, value: 2 }
    )).toEqual({ tag: ValueTag.Number, value: 3 });
    expect(getBuiltin("SUMSQ")?.(
      { tag: ValueTag.Number, value: 2 },
      { tag: ValueTag.Number, value: 3 }
    )).toEqual({ tag: ValueTag.Number, value: 13 });
    expect(getBuiltin("SERIESSUM")?.(
      { tag: ValueTag.Number, value: 2 },
      { tag: ValueTag.Number, value: 1 },
      { tag: ValueTag.Number, value: 2 },
      { tag: ValueTag.Number, value: 1 },
      { tag: ValueTag.Number, value: 2 }
    )).toEqual({ tag: ValueTag.Number, value: 18 });
    expect(getBuiltin("BASE")?.(
      { tag: ValueTag.Number, value: 31 },
      { tag: ValueTag.Number, value: 16 },
      { tag: ValueTag.Number, value: 4 }
    )).toEqual({ tag: ValueTag.String, value: "001F", stringId: 0 });
    expect(getBuiltin("DECIMAL")?.(
      { tag: ValueTag.String, value: "1F", stringId: 1 },
      { tag: ValueTag.Number, value: 16 }
    )).toEqual({ tag: ValueTag.Number, value: 31 });
    expect(getBuiltin("ROMAN")?.({ tag: ValueTag.Number, value: 14 })).toEqual({
      tag: ValueTag.String,
      value: "XIV",
      stringId: 0
    });
    expect(getBuiltin("ARABIC")?.({ tag: ValueTag.String, value: "XIV", stringId: 1 })).toEqual({
      tag: ValueTag.Number,
      value: 14
    });

    expect(getBuiltin("MUNIT")?.({ tag: ValueTag.Number, value: 2 })).toEqual({
      kind: "array",
      rows: 2,
      cols: 2,
      values: [
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 1 }
      ]
    });

    const randArray = getBuiltin("RANDARRAY")?.(
      { tag: ValueTag.Number, value: 2 },
      { tag: ValueTag.Number, value: 2 },
      { tag: ValueTag.Number, value: 3 },
      { tag: ValueTag.Number, value: 7 },
      { tag: ValueTag.Boolean, value: true }
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
    expect(getBuiltin("CEILING.PRECISE")?.({ tag: ValueTag.String, value: "bad", stringId: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value
    });
    expect(getBuiltin("ISO.CEILING")?.(
      { tag: ValueTag.Number, value: 2 },
      { tag: ValueTag.Number, value: 0 }
    )).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });
    expect(getBuiltin("ROUNDUP")?.(
      { tag: ValueTag.Number, value: 5 },
      { tag: ValueTag.String, value: "bad", stringId: 1 }
    )).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });
    expect(getBuiltin("TRUNC")?.(
      { tag: ValueTag.Number, value: 5 },
      { tag: ValueTag.String, value: "bad", stringId: 1 }
    )).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });
    expect(getBuiltin("EVEN")?.({ tag: ValueTag.String, value: "bad", stringId: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value
    });
    expect(getBuiltin("ODD")?.({ tag: ValueTag.String, value: "bad", stringId: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value
    });
    expect(getBuiltin("LN")?.({ tag: ValueTag.Number, value: 0 })).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });
    expect(getBuiltin("SQRT")?.({ tag: ValueTag.Number, value: -1 })).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });
    expect(getBuiltin("COT")?.({ tag: ValueTag.Number, value: 0 })).toEqual({ tag: ValueTag.Error, code: ErrorCode.Div0 });
    expect(getBuiltin("CSC")?.({ tag: ValueTag.Number, value: 0 })).toEqual({ tag: ValueTag.Error, code: ErrorCode.Div0 });
    expect(getBuiltin("FACT")?.({ tag: ValueTag.Number, value: -1 })).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });
    expect(getBuiltin("FACTDOUBLE")?.({ tag: ValueTag.Number, value: 6 })).toEqual({ tag: ValueTag.Number, value: 48 });
    expect(getBuiltin("COMBIN")?.(
      { tag: ValueTag.Number, value: 2 },
      { tag: ValueTag.Number, value: 3 }
    )).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });
    expect(getBuiltin("COMBINA")?.(
      { tag: ValueTag.Number, value: 0 },
      { tag: ValueTag.Number, value: 2 }
    )).toEqual({ tag: ValueTag.Number, value: 0 });
    expect(getBuiltin("GCD")?.()).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });
    expect(getBuiltin("LCM")?.()).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });
    expect(getBuiltin("MROUND")?.(
      { tag: ValueTag.Number, value: -10 },
      { tag: ValueTag.Number, value: 3 }
    )).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });
    expect(getBuiltin("MULTINOMIAL")?.(
      { tag: ValueTag.Number, value: 1 },
      { tag: ValueTag.Number, value: -1 }
    )).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });
    expect(getBuiltin("QUOTIENT")?.(
      { tag: ValueTag.String, value: "bad", stringId: 1 },
      { tag: ValueTag.Number, value: 1 }
    )).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });
    expect(getBuiltin("QUOTIENT")?.(
      { tag: ValueTag.Number, value: 5 },
      { tag: ValueTag.Number, value: 0 }
    )).toEqual({ tag: ValueTag.Error, code: ErrorCode.Div0 });
    expect(getBuiltin("RANDBETWEEN")?.(
      { tag: ValueTag.Number, value: 5 },
      { tag: ValueTag.Number, value: 3 }
    )).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });
    expect(getBuiltin("BASE")?.(
      { tag: ValueTag.Number, value: -1 },
      { tag: ValueTag.Number, value: 2 }
    )).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });
    expect(getBuiltin("DECIMAL")?.(
      { tag: ValueTag.String, value: "1Z", stringId: 1 },
      { tag: ValueTag.Number, value: 10 }
    )).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });
    expect(getBuiltin("ROMAN")?.({ tag: ValueTag.Number, value: 0 })).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });
    expect(getBuiltin("ARABIC")?.({ tag: ValueTag.String, value: "IIII", stringId: 1 })).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });
    expect(getBuiltin("RANDARRAY")?.(
      { tag: ValueTag.Number, value: 1 },
      { tag: ValueTag.Number, value: 1 },
      { tag: ValueTag.Number, value: 5 },
      { tag: ValueTag.Number, value: 3 }
    )).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });
    expect(getBuiltin("RANDARRAY")?.(
      { tag: ValueTag.Number, value: 2 },
      { tag: ValueTag.Number, value: 2 },
      { tag: ValueTag.Number, value: 1 },
      { tag: ValueTag.Number, value: 3 },
      { tag: ValueTag.Number, value: 1 }
    )).toMatchObject({ kind: "array", rows: 2, cols: 2 });
    expect(getBuiltin("MUNIT")?.({ tag: ValueTag.Number, value: 0 })).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });
    expect(getBuiltin("SERIESSUM")?.(
      { tag: ValueTag.String, value: "bad", stringId: 1 },
      { tag: ValueTag.Number, value: 1 },
      { tag: ValueTag.Number, value: 2 },
      { tag: ValueTag.Number, value: 1 }
    )).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });
    expect(getBuiltin("SQRTPI")?.({ tag: ValueTag.String, value: "bad", stringId: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value
    });
    expect(getBuiltin("SUBTOTAL")?.(
      { tag: ValueTag.Number, value: 9 },
      { tag: ValueTag.Number, value: 2 },
      { tag: ValueTag.Number, value: 3 }
    )).toEqual({ tag: ValueTag.Number, value: 5 });
    expect(getBuiltin("SUBTOTAL")?.({ tag: ValueTag.String, value: "bad", stringId: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value
    });
    expect(getBuiltin("AGGREGATE")?.(
      { tag: ValueTag.Number, value: 6 },
      { tag: ValueTag.Number, value: 0 },
      { tag: ValueTag.Number, value: 2 },
      { tag: ValueTag.Number, value: 3 },
      { tag: ValueTag.Number, value: 4 }
    )).toEqual({ tag: ValueTag.Number, value: 24 });
    expect(getBuiltin("AGGREGATE")?.(
      { tag: ValueTag.Number, value: 99 },
      { tag: ValueTag.Number, value: 0 },
      { tag: ValueTag.Number, value: 2 }
    )).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });
    expect(getBuiltin("ARABIC")?.({ tag: ValueTag.Number, value: 1 })).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });
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
    expect(getBuiltinId("filter")).toBeUndefined();
    expect(getBuiltinId("let")).toBeUndefined();
    expect(getBuiltinId("textjoin")).toBeUndefined();
  });
});
