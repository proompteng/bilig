import { describe, expect, it } from "vitest";
import { BuiltinId, ErrorCode, ValueTag } from "@bilig/protocol";
import { getBuiltin, getBuiltinId } from "../builtins.js";

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
});
