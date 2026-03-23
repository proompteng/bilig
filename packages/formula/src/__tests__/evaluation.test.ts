import { describe, expect, it } from "vitest";
import { ErrorCode, ValueTag, type CellValue } from "@bilig/protocol";
import { getBuiltin, getBuiltinId } from "../builtins.js";
import { evaluateAst, evaluateAstResult } from "../js-evaluator.js";
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
    expect(evaluateAst(parseFormula("SIN(A1)"), context)).toEqual({ tag: ValueTag.Number, value: Math.sin(4) });
    expect(evaluateAst(parseFormula("POWER(2,3)"), context)).toEqual({ tag: ValueTag.Number, value: 8 });
    expect(evaluateAst(parseFormula("TRUNC(-3.98,1)"), context)).toEqual({ tag: ValueTag.Number, value: -3.9 });
    expect(evaluateAst(parseFormula("PRODUCT(2,3,4)"), context)).toEqual({ tag: ValueTag.Number, value: 24 });
    expect(evaluateAst(parseFormula("SUMSQ(2,3)"), context)).toEqual({ tag: ValueTag.Number, value: 13 });
    expect(evaluateAst(parseFormula("SIGN(-2)"), context)).toEqual({ tag: ValueTag.Number, value: -1 });
    expect(evaluateAst(parseFormula("MDETERM(C1:D2)"), {
      ...context,
      resolveRange: (_sheetName: string, start: string, end: string): CellValue[] =>
        start === "C1" && end === "D2"
          ? [
              { tag: ValueTag.Number, value: 1 },
              { tag: ValueTag.Number, value: 2 },
              { tag: ValueTag.Number, value: 3 },
              { tag: ValueTag.Number, value: 4 }
            ]
          : context.resolveRange(_sheetName, start, end, "cells")
    })).toEqual({ tag: ValueTag.Number, value: -2 });
    expect(evaluateAst(parseFormula("DAYS(46101,46094)"), context)).toEqual({ tag: ValueTag.Number, value: 7 });
    expect(evaluateAst(parseFormula("WEEKNUM(46096)"), context)).toEqual({ tag: ValueTag.Number, value: 12 });
    expect(evaluateAst(parseFormula("WORKDAY(46094,1,46097)"), context)).toEqual({ tag: ValueTag.Number, value: 46098 });
    expect(evaluateAst(parseFormula("NETWORKDAYS(46094,46101,46097,46101)"), context)).toEqual({ tag: ValueTag.Number, value: 4 });
    expect(evaluateAst(parseFormula("REPLACE(\"alphabet\",3,2,\"Z\")"), context)).toEqual({
      tag: ValueTag.String,
      value: "alZabet",
      stringId: 0
    });
    expect(evaluateAst(parseFormula("SUBSTITUTE(\"banana\",\"an\",\"oo\",2)"), context)).toEqual({
      tag: ValueTag.String,
      value: "banooa",
      stringId: 0
    });
    expect(evaluateAst(parseFormula("REPT(\"x\",3)"), context)).toEqual({
      tag: ValueTag.String,
      value: "xxx",
      stringId: 0
    });
    expect(evaluateAst(parseFormula("T.DIST(0.1,2,TRUE)"), context)).toEqual({ tag: ValueTag.Error, code: ErrorCode.Blocked });
    expect(evaluateAst(parseFormula("TEXTJOIN(\",\",TRUE,\"a\",\"b\")"), context)).toEqual({ tag: ValueTag.Error, code: ErrorCode.Blocked });
    expect(evaluateAst(parseFormula("LET(x,2,x+3)"), context)).toEqual({ tag: ValueTag.Number, value: 5 });
    expect(evaluateAst(parseFormula("LAMBDA(x,x+1)(4)"), context)).toEqual({ tag: ValueTag.Number, value: 5 });
    expect(evaluateAst(parseFormula("LET(fn,LAMBDA(x,x*2),fn(4))"), context)).toEqual({ tag: ValueTag.Number, value: 8 });
    expect(evaluateAst(parseFormula("TRUE()"), context)).toEqual({ tag: ValueTag.Boolean, value: true });
    expect(evaluateAst(parseFormula("FALSE()"), context)).toEqual({ tag: ValueTag.Boolean, value: false });
    expect(evaluateAst(parseFormula("IFS(A1>3,\"big\",TRUE(),\"small\")"), context)).toEqual({
      tag: ValueTag.String,
      value: "big",
      stringId: 0
    });
    expect(evaluateAst(parseFormula("SWITCH(A1,4,\"four\",\"other\")"), context)).toEqual({
      tag: ValueTag.String,
      value: "four",
      stringId: 0
    });
    expect(evaluateAst(parseFormula("XOR(TRUE(),FALSE(),TRUE())"), context)).toEqual({ tag: ValueTag.Boolean, value: false });
    expect(evaluateAst(parseFormula("TEXTBEFORE(\"alpha-beta\",\"-\")"), context)).toEqual({
      tag: ValueTag.String,
      value: "alpha",
      stringId: 0
    });
    expect(evaluateAst(parseFormula("MissingFn(A1)"), context)).toEqual({ tag: ValueTag.Error, code: ErrorCode.Name });
    expect(evaluateAst(parseFormula("A:A"), context)).toEqual({ tag: ValueTag.Empty });
  });

  it("resolves scalar defined names through the JS evaluator", () => {
    const value = evaluateAst(parseFormula("TaxRate*A1"), {
      sheetName: "Sheet1",
      resolveCell: (_sheetName: string, address: string): CellValue =>
        address === "A1" ? { tag: ValueTag.Number, value: 100 } : { tag: ValueTag.Empty },
      resolveRange: (): CellValue[] => [],
      resolveName: (name: string): CellValue =>
        name.toUpperCase() === "TAXRATE" ? { tag: ValueTag.Number, value: 0.085 } : { tag: ValueTag.Error, code: ErrorCode.Name }
    });

    expect(value).toEqual({ tag: ValueTag.Number, value: 8.5 });
  });

  it("preserves sequence array results while flattening them for scalar consumers", () => {
    const context = {
      sheetName: "Sheet1",
      resolveCell: (): CellValue => ({ tag: ValueTag.Empty }),
      resolveRange: (): CellValue[] => []
    };

    expect(evaluateAstResult(parseFormula("SEQUENCE(3,1,1,1)"), context)).toEqual({
      kind: "array",
      rows: 3,
      cols: 1,
      values: [
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 3 }
      ]
    });
    expect(evaluateAst(parseFormula("SUM(SEQUENCE(3,1,1,1))"), context)).toEqual({ tag: ValueTag.Number, value: 6 });

    expect(evaluateAstResult(parseFormula("FILTER(A1:A4,A1:A4>2)"), {
      ...context,
      resolveRange: (): CellValue[] => [
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 4 }
      ]
    })).toEqual({
      kind: "array",
      rows: 2,
      cols: 1,
      values: [
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 4 }
      ]
    });

    expect(evaluateAstResult(parseFormula("UNIQUE(A1:A4)"), {
      ...context,
      resolveRange: (): CellValue[] => [
        { tag: ValueTag.String, value: "A", stringId: 0 },
        { tag: ValueTag.String, value: "B", stringId: 0 },
        { tag: ValueTag.String, value: "A", stringId: 0 },
        { tag: ValueTag.String, value: "C", stringId: 0 }
      ]
    })).toEqual({
      kind: "array",
      rows: 3,
      cols: 1,
      values: [
        { tag: ValueTag.String, value: "A", stringId: 0 },
        { tag: ValueTag.String, value: "B", stringId: 0 },
        { tag: ValueTag.String, value: "C", stringId: 0 }
      ]
    });

    expect(evaluateAstResult(parseFormula("MAKEARRAY(2,2,LAMBDA(r,c,r+c))"), context)).toEqual({
      kind: "array",
      rows: 2,
      cols: 2,
      values: [
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 4 }
      ]
    });

    expect(evaluateAstResult(parseFormula("MUNIT(3)"), context)).toEqual({
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
        { tag: ValueTag.Number, value: 1 }
      ]
    });

    expect(evaluateAstResult(parseFormula("MMULT(A1:B2,C1:D2)"), {
      ...context,
      resolveRange: (_sheetName: string, start: string, end: string): CellValue[] => {
        if (start === "A1" && end === "B2") {
          return [
            { tag: ValueTag.Number, value: 1 },
            { tag: ValueTag.Number, value: 2 },
            { tag: ValueTag.Number, value: 3 },
            { tag: ValueTag.Number, value: 4 }
          ];
        }
        return [
          { tag: ValueTag.Number, value: 5 },
          { tag: ValueTag.Number, value: 6 },
          { tag: ValueTag.Number, value: 7 },
          { tag: ValueTag.Number, value: 8 }
        ];
      }
    })).toEqual({
      kind: "array",
      rows: 2,
      cols: 2,
      values: [
        { tag: ValueTag.Number, value: 19 },
        { tag: ValueTag.Number, value: 22 },
        { tag: ValueTag.Number, value: 43 },
        { tag: ValueTag.Number, value: 50 }
      ]
    });

    expect(evaluateAst(parseFormula("SUMXMY2(A1:A2,B1:B2)"), {
      ...context,
      resolveRange: (_sheetName: string, start: string, _end: string): CellValue[] =>
        start === "A1"
          ? [
              { tag: ValueTag.Number, value: 1 },
              { tag: ValueTag.Number, value: 2 }
            ]
          : [
              { tag: ValueTag.Number, value: 3 },
              { tag: ValueTag.Number, value: 4 }
            ]
    })).toEqual({ tag: ValueTag.Number, value: 8 });

    const randomGrid = evaluateAstResult(parseFormula("RANDARRAY(2,2,3,7,TRUE())"), context);
    expect(randomGrid).toMatchObject({ kind: "array", rows: 2, cols: 2 });
    if (randomGrid.kind !== "array") {
      throw new Error("expected RANDARRAY to return an array");
    }
    for (const value of randomGrid.values) {
      expect(value.tag).toBe(ValueTag.Number);
      expect(value.value).toBeGreaterThanOrEqual(3);
      expect(value.value).toBeLessThanOrEqual(7);
    }

    expect(evaluateAstResult(parseFormula("MAP(A1:A3,LAMBDA(x,x*2))"), {
      ...context,
      resolveRange: (): CellValue[] => [
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 3 }
      ]
    })).toEqual({
      kind: "array",
      rows: 3,
      cols: 1,
      values: [
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 4 },
        { tag: ValueTag.Number, value: 6 }
      ]
    });

    expect(evaluateAst(parseFormula("REDUCE(0,A1:A3,LAMBDA(acc,x,acc+x))"), {
      ...context,
      resolveRange: (): CellValue[] => [
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 3 }
      ]
    })).toEqual({ tag: ValueTag.Number, value: 6 });

    expect(evaluateAstResult(parseFormula("SCAN(0,A1:A3,LAMBDA(acc,x,acc+x))"), {
      ...context,
      resolveRange: (): CellValue[] => [
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 3 }
      ]
    })).toEqual({
      kind: "array",
      rows: 3,
      cols: 1,
      values: [
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 6 }
      ]
    });

    expect(evaluateAstResult(parseFormula("BYROW(A1:B2,LAMBDA(r,SUM(r)))"), {
      ...context,
      resolveRange: (): CellValue[] => [
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 4 }
      ]
    })).toEqual({
      kind: "array",
      rows: 2,
      cols: 1,
      values: [
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 7 }
      ]
    });

    expect(evaluateAstResult(parseFormula("BYCOL(A1:B2,LAMBDA(c,SUM(c)))"), {
      ...context,
      resolveRange: (): CellValue[] => [
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 4 }
      ]
    })).toEqual({
      kind: "array",
      rows: 1,
      cols: 2,
      values: [
        { tag: ValueTag.Number, value: 4 },
        { tag: ValueTag.Number, value: 6 }
      ]
    });
  });
});
