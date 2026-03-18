import { describe, expect, it } from "vitest";
import { ErrorCode, ValueTag, type CellValue } from "@bilig/protocol";
import { getBuiltin, getBuiltinId } from "../builtins.js";
import { evaluateAst } from "../js-evaluator.js";
import { parseFormula } from "../parser.js";

describe("formula builtins and JS evaluator", () => {
  it("covers scalar builtins and builtin id lookup", () => {
    const SUM = getBuiltin("sum")!;
    const AVG = getBuiltin("AVG")!;
    const MOD = getBuiltin("MOD")!;
    const LEN = getBuiltin("LEN")!;
    const CONCAT = getBuiltin("CONCAT")!;
    const IF = getBuiltin("IF")!;
    const AND = getBuiltin("AND")!;
    const OR = getBuiltin("OR")!;
    const NOT = getBuiltin("NOT")!;

    expect(getBuiltinId("sum")).toBeDefined();
    expect(getBuiltinId("")).toBeUndefined();
    expect(getBuiltin("missing")).toBeUndefined();

    expect(SUM({ tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Boolean, value: true }, { tag: ValueTag.Empty })).toEqual({
      tag: ValueTag.Number,
      value: 3
    });
    expect(
      AVG(
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.String, value: "ignored", stringId: 0 },
        { tag: ValueTag.Number, value: 4 }
      )
    ).toEqual({ tag: ValueTag.Number, value: 3 });
    expect(MOD({ tag: ValueTag.Number, value: 8 }, { tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Div0
    });
    expect(LEN({ tag: ValueTag.Boolean, value: true })).toEqual({ tag: ValueTag.Number, value: 4 });
    expect(
      CONCAT(
        { tag: ValueTag.String, value: "hello", stringId: 0 },
        { tag: ValueTag.Empty },
        { tag: ValueTag.Number, value: 7 }
      )
    ).toEqual({ tag: ValueTag.String, value: "hello7", stringId: 0 });
    expect(IF({ tag: ValueTag.Number, value: 0 }, { tag: ValueTag.Number, value: 1 }, { tag: ValueTag.Number, value: 2 })).toEqual({
      tag: ValueTag.Number,
      value: 2
    });
    expect(AND({ tag: ValueTag.Number, value: 1 }, { tag: ValueTag.Boolean, value: true })).toEqual({
      tag: ValueTag.Boolean,
      value: true
    });
    expect(OR({ tag: ValueTag.Empty }, { tag: ValueTag.Boolean, value: true })).toEqual({
      tag: ValueTag.Boolean,
      value: true
    });
    expect(NOT({ tag: ValueTag.Empty })).toEqual({ tag: ValueTag.Boolean, value: true });
  });

  it("evaluates range, concat, unary, comparison, builtin, and error paths", () => {
    const num = (value: number): CellValue => ({ tag: ValueTag.Number, value });
    const empty = (): CellValue => ({ tag: ValueTag.Empty });
    const text = (value: string): CellValue => ({ tag: ValueTag.String, value, stringId: 0 });
    const context = {
      sheetName: "Sheet1",
      resolveCell: (sheetName: string, address: string): CellValue => {
        if (sheetName === "Sheet2" && address === "B1") {
          return num(9);
        }
        if (address === "A1") {
          return num(4);
        }
        if (address === "A2") {
          return text("x");
        }
        return empty();
      },
      resolveRange: (_sheetName: string, start: string, end: string, refKind: "cells" | "rows" | "cols"): CellValue[] => {
        if (refKind === "cells" && start === "A1" && end === "B2") {
          return [num(1), num(2), num(3)];
        }
        if (refKind === "rows") {
          return [num(6)];
        }
        return [];
      }
    };

    expect(evaluateAst(parseFormula("A1&\"!\""), context)).toEqual({
      tag: ValueTag.String,
      value: "4!",
      stringId: 0
    });
    expect(evaluateAst(parseFormula("A2=\"X\""), context)).toEqual({ tag: ValueTag.Boolean, value: true });
    expect(evaluateAst(parseFormula("\"b\">\"A\""), context)).toEqual({ tag: ValueTag.Boolean, value: true });
    expect(evaluateAst(parseFormula("-A1"), context)).toEqual({ tag: ValueTag.Number, value: -4 });
    expect(evaluateAst(parseFormula("A1>=Sheet2!B1"), context)).toEqual({ tag: ValueTag.Boolean, value: false });
    expect(evaluateAst(parseFormula("SUM(A1:B2)"), context)).toEqual({ tag: ValueTag.Number, value: 6 });
    expect(evaluateAst(parseFormula("SUM(1:1)"), context)).toEqual({ tag: ValueTag.Number, value: 6 });
    expect(evaluateAst(parseFormula("A1/0"), context)).toEqual({ tag: ValueTag.Error, code: ErrorCode.Div0 });
    expect(evaluateAst(parseFormula("A2+1"), context)).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });
    expect(evaluateAst(parseFormula("MissingFn(A1)"), context)).toEqual({ tag: ValueTag.Error, code: ErrorCode.Name });
    expect(evaluateAst(parseFormula("A:A"), context)).toEqual({ tag: ValueTag.Empty });
  });
});
