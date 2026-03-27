import { describe, expect, it } from "vitest";
import { ErrorCode, ValueTag, type CellValue } from "@bilig/protocol";
import {
  evaluatePlan,
  evaluatePlanResult,
  lowerToPlan,
  optimizeFormula,
  parseFormula,
} from "../index.js";
import type { FormulaNode } from "../ast.js";

const context = {
  sheetName: "Sheet1",
  resolveCell: (_sheetName: string, address: string): CellValue => {
    switch (address) {
      case "A1":
        return { tag: ValueTag.Number, value: 2 };
      case "B1":
        return { tag: ValueTag.Number, value: 3 };
      default:
        return { tag: ValueTag.Empty };
    }
  },
  resolveRange: (_sheetName: string, start: string, end: string): CellValue[] => {
    if (start === "A1" && end === "B2") {
      return [
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Boolean, value: true },
        { tag: ValueTag.Empty },
      ];
    }
    return [];
  },
};

describe("js evaluator", () => {
  it("evaluates direct plans for ranges, jumps, and fallback stack handling", () => {
    expect(
      evaluatePlan(
        [
          { opcode: "push-range", start: "A1", end: "B2", refKind: "cells" },
          { opcode: "call", callee: "SUM", argc: 1 },
          { opcode: "return" },
        ],
        context,
      ),
    ).toEqual({ tag: ValueTag.Number, value: 6 });

    expect(
      evaluatePlan(
        [{ opcode: "push-range", start: "A1", end: "B2", refKind: "cells" }, { opcode: "return" }],
        context,
      ),
    ).toEqual({ tag: ValueTag.Number, value: 2 });

    expect(
      evaluatePlan(
        [
          { opcode: "push-number", value: 0 },
          { opcode: "jump-if-false", target: 4 },
          { opcode: "push-number", value: 1 },
          { opcode: "jump", target: 5 },
          { opcode: "push-number", value: 2 },
          { opcode: "return" },
        ],
        context,
      ),
    ).toEqual({ tag: ValueTag.Number, value: 2 });

    expect(
      evaluatePlan([{ opcode: "call", callee: "SUM", argc: 1 }, { opcode: "return" }], context),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });

    expect(
      evaluatePlan(
        [{ opcode: "call", callee: "DOES_NOT_EXIST", argc: 0 }, { opcode: "return" }],
        context,
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Name });

    expect(evaluatePlan([{ opcode: "push-number", value: 9 }], context)).toEqual({
      tag: ValueTag.Number,
      value: 9,
    });
  });

  it("keeps range shape for lookup/reference builtins", () => {
    expect(
      evaluatePlan(
        [
          { opcode: "push-number", value: 3 },
          { opcode: "push-range", start: "A1", end: "A4", refKind: "cells" },
          { opcode: "push-number", value: 0 },
          { opcode: "call", callee: "MATCH", argc: 3 },
          { opcode: "return" },
        ],
        {
          ...context,
          resolveRange: (_sheetName: string, start: string, end: string): CellValue[] => {
            if (start === "A1" && end === "A4") {
              return [num(2), num(3), num(4), num(5)];
            }
            return [];
          },
        },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 2 });

    expect(
      evaluatePlan(
        [
          { opcode: "push-range", start: "A1", end: "B2", refKind: "cells" },
          { opcode: "push-number", value: 2 },
          { opcode: "push-number", value: 2 },
          { opcode: "call", callee: "INDEX", argc: 3 },
          { opcode: "return" },
        ],
        context,
      ),
    ).toEqual({ tag: ValueTag.Empty });
  });

  it("lowers row and column refs into NaN sentinels for the JS path", () => {
    expect(lowerToPlan({ kind: "RowRef", ref: "3" } as FormulaNode)).toEqual([
      { opcode: "push-number", value: Number.NaN },
      { opcode: "return" },
    ]);
    expect(lowerToPlan({ kind: "ColumnRef", ref: "C" } as FormulaNode)).toEqual([
      { opcode: "push-number", value: Number.NaN },
      { opcode: "return" },
    ]);
  });

  it("resolves scalar defined names and returns #NAME? when missing", () => {
    expect(
      evaluatePlan(
        [
          { opcode: "push-name", name: "TaxRate" },
          { opcode: "push-number", value: 1 },
          { opcode: "binary", operator: "+" },
          { opcode: "return" },
        ],
        {
          ...context,
          resolveName: (name: string): CellValue =>
            name === "TaxRate"
              ? { tag: ValueTag.Number, value: 0.5 }
              : { tag: ValueTag.Error, code: ErrorCode.Name },
        },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 1.5 });

    expect(
      evaluatePlan([{ opcode: "push-name", name: "MissingName" }, { opcode: "return" }], context),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Name,
    });
  });

  it("supports LET scopes in lowered plans", () => {
    expect(evaluatePlan(lowerToPlan(parseFormula("LET(x,2,x+3)")), context)).toEqual({
      tag: ValueTag.Number,
      value: 5,
    });
  });

  it("supports lambda invocation and lambda-array helpers in lowered plans", () => {
    expect(evaluatePlan(lowerToPlan(parseFormula("LAMBDA(x,x+1)(4)")), context)).toEqual({
      tag: ValueTag.Number,
      value: 5,
    });
    expect(
      evaluatePlan(lowerToPlan(parseFormula("LAMBDA(x,y,IF(ISOMITTED(y),x,y))(4)")), context),
    ).toEqual({
      tag: ValueTag.Number,
      value: 4,
    });
    expect(
      evaluatePlan(lowerToPlan(parseFormula("LAMBDA(x,y,IF(ISOMITTED(y),x,y))(4,9)")), context),
    ).toEqual({
      tag: ValueTag.Number,
      value: 9,
    });

    expect(evaluatePlan(lowerToPlan(parseFormula("LET(fn,LAMBDA(x,x+1),fn(4))")), context)).toEqual(
      {
        tag: ValueTag.Number,
        value: 5,
      },
    );

    expect(evaluatePlan(lowerToPlan(parseFormula("LAMBDA(x,ISOMITTED(x))()")), context)).toEqual({
      tag: ValueTag.Boolean,
      value: true,
    });
    expect(evaluatePlan(lowerToPlan(parseFormula("LAMBDA(x,SUM(x))()")), context)).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });
    expect(
      evaluatePlan(lowerToPlan(parseFormula("LAMBDA(x,IF(ISOMITTED(x),9,x))(4)")), context),
    ).toEqual({
      tag: ValueTag.Number,
      value: 4,
    });
  });

  it("evaluates higher-order array helpers on the JS plan runtime", () => {
    expect(
      evaluatePlan(lowerToPlan(parseFormula("MAKEARRAY(2,2,LAMBDA(r,c,r+c))")), context),
    ).toEqual(num(2));
    expect(evaluatePlan(lowerToPlan(parseFormula("MAP(A1:B2,LAMBDA(x,x+1))")), context)).toEqual(
      num(3),
    );
    expect(
      evaluatePlan(lowerToPlan(parseFormula("BYROW(A1:B2,LAMBDA(r,SUM(r)))")), context),
    ).toEqual(num(5));
    expect(
      evaluatePlan(lowerToPlan(parseFormula("BYCOL(A1:B2,LAMBDA(c,SUM(c)))")), context),
    ).toEqual(num(3));
    expect(
      evaluatePlan(lowerToPlan(parseFormula("REDUCE(0,A1:B2,LAMBDA(a,x,a+x))")), context),
    ).toEqual(num(6));
    expect(
      evaluatePlan(lowerToPlan(parseFormula("SCAN(0,A1:B2,LAMBDA(a,x,a+x))")), context),
    ).toEqual(num(2));
  });

  it("trims outer empty rows and columns for TRIMRANGE", () => {
    const trimContext = {
      ...context,
      resolveRange: (_sheetName: string, start: string, end: string): CellValue[] => {
        if (start === "A1" && end === "D4") {
          return [
            empty(),
            empty(),
            empty(),
            empty(),
            empty(),
            num(1),
            num(2),
            empty(),
            empty(),
            num(3),
            empty(),
            empty(),
            empty(),
            empty(),
            empty(),
            empty(),
          ];
        }
        if (start === "F1" && end === "G2") {
          return [empty(), empty(), empty(), empty()];
        }
        return [];
      },
    };

    expect(evaluatePlanResult(lowerToPlan(parseFormula("TRIMRANGE(A1:D4)")), trimContext)).toEqual({
      kind: "array",
      rows: 2,
      cols: 2,
      values: [num(1), num(2), num(3), empty()],
    });
    expect(
      evaluatePlanResult(lowerToPlan(parseFormula("TRIMRANGE(A1:D4,1,1)")), trimContext),
    ).toEqual({
      kind: "array",
      rows: 3,
      cols: 3,
      values: [num(1), num(2), empty(), num(3), empty(), empty(), empty(), empty(), empty()],
    });
    expect(evaluatePlanResult(lowerToPlan(parseFormula("TRIMRANGE(F1:G2)")), trimContext)).toEqual({
      kind: "array",
      rows: 1,
      cols: 1,
      values: [empty()],
    });
    expect(evaluatePlan(lowerToPlan(parseFormula("TRIMRANGE(A1:D4,4)")), trimContext)).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });
  });

  it("covers special-call rewrites and evaluator guard rails", () => {
    expect(evaluatePlan(lowerToPlan(parseFormula("IFS(FALSE,1,TRUE,2)")), context)).toEqual({
      tag: ValueTag.Number,
      value: 2,
    });
    expect(evaluatePlan(lowerToPlan(parseFormula('SWITCH("b","a",1,"b",2,9)')), context)).toEqual({
      tag: ValueTag.Number,
      value: 2,
    });
    expect(evaluatePlan(lowerToPlan(parseFormula("XOR(TRUE,FALSE,TRUE)")), context)).toEqual({
      tag: ValueTag.Boolean,
      value: false,
    });
    expect(evaluatePlan(lowerToPlan(parseFormula("TRUE(1)")), context)).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });
    expect(evaluatePlan(lowerToPlan(parseFormula("LET(x,1)")), context)).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });
    expect(evaluatePlan(lowerToPlan(parseFormula("LET(1,2,3)")), context)).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });
    expect(
      evaluatePlan(
        [{ opcode: "push-range", start: "bad", end: "B2", refKind: "cells" }, { opcode: "return" }],
        {
          ...context,
          resolveRange: () => [num(7), num(8)],
        },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 7 });
    expect(
      evaluatePlan([{ opcode: "bind-name", name: "x" }, { opcode: "return" }], context),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });
    expect(evaluatePlan(lowerToPlan(parseFormula("SUM(LAMBDA(x,x))")), context)).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });
  });

  it("covers contextual reference builtins and array-lambda error paths", () => {
    const metadataContext = {
      ...context,
      sheetName: "Sheet2",
      currentAddress: "C4",
      listSheetNames: () => ["Sheet1", "Sheet2", "Summary"],
      resolveFormula: (sheetName: string, address: string): string | undefined =>
        sheetName === "Sheet2" && address === "B1"
          ? "SUM(A1:A2)"
          : sheetName === "Sheet2" && address === "C1"
            ? "A1*2"
            : undefined,
      resolveName: (name: string): CellValue =>
        name === "TaxRate"
          ? { tag: ValueTag.Number, value: 0.085 }
          : { tag: ValueTag.Error, code: ErrorCode.Name },
    };

    expect(evaluatePlan(lowerToPlan(parseFormula("ROW()")), metadataContext)).toEqual({
      tag: ValueTag.Number,
      value: 4,
    });
    expect(evaluatePlan(lowerToPlan(parseFormula("COLUMN(B:D)")), metadataContext)).toEqual({
      tag: ValueTag.Number,
      value: 2,
    });
    expect(
      evaluatePlan(lowerToPlan(parseFormula("FORMULATEXT(Sheet2!B1)")), metadataContext),
    ).toEqual({
      tag: ValueTag.String,
      value: "=SUM(A1:A2)",
      stringId: 0,
    });
    expect(evaluatePlan(lowerToPlan(parseFormula("FORMULA(Sheet2!C1)")), metadataContext)).toEqual({
      tag: ValueTag.String,
      value: "=A1*2",
      stringId: 0,
    });
    expect(evaluatePlan(lowerToPlan(parseFormula("FORMULATEXT(1)")), metadataContext)).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    });
    expect(
      evaluatePlan(lowerToPlan(parseFormula("FORMULATEXT(Sheet2!A9)")), metadataContext),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.NA,
    });
    expect(evaluatePlan(lowerToPlan(parseFormula("SHEET()")), metadataContext)).toEqual({
      tag: ValueTag.Number,
      value: 2,
    });
    expect(evaluatePlan(lowerToPlan(parseFormula('SHEET("Summary")')), metadataContext)).toEqual({
      tag: ValueTag.Number,
      value: 3,
    });
    expect(evaluatePlan(lowerToPlan(parseFormula('SHEET("Missing")')), metadataContext)).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.NA,
    });
    expect(evaluatePlan(lowerToPlan(parseFormula("SHEETS()")), metadataContext)).toEqual({
      tag: ValueTag.Number,
      value: 3,
    });
    expect(evaluatePlan(lowerToPlan(parseFormula('SHEETS("Summary")')), metadataContext)).toEqual({
      tag: ValueTag.Number,
      value: 1,
    });
    expect(evaluatePlan(lowerToPlan(parseFormula('CELL("filename")')), metadataContext)).toEqual({
      tag: ValueTag.String,
      value: "",
      stringId: 0,
    });
    expect(
      evaluatePlan(lowerToPlan(parseFormula('CELL("type")')), {
        ...metadataContext,
        currentAddress: undefined,
      }),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });
    expect(evaluatePlan(lowerToPlan(parseFormula('CELL("bogus",A1)')), metadataContext)).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });
    expect(
      evaluatePlan(lowerToPlan(parseFormula('CELL("address")')), {
        ...metadataContext,
        currentAddress: undefined,
      }),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });

    expect(evaluatePlan(lowerToPlan(parseFormula("LAMBDA(x,x)(1,2)")), metadataContext)).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });
    expect(
      evaluatePlan(lowerToPlan(parseFormula("MAKEARRAY(0,2,LAMBDA(r,c,r+c))")), metadataContext),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });
    expect(
      evaluatePlan(lowerToPlan(parseFormula("MAP(A1:B2,LAMBDA(x,SEQUENCE(2)))")), metadataContext),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });
    expect(
      evaluatePlan(
        lowerToPlan(parseFormula("BYROW(A1:B2,LAMBDA(r,SEQUENCE(2)))")),
        metadataContext,
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });
    expect(
      evaluatePlan(
        lowerToPlan(parseFormula("SCAN(A1:B2,LAMBDA(a,x,SEQUENCE(2)))")),
        metadataContext,
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });
    expect(evaluatePlan(lowerToPlan(parseFormula("LAMBDA()")), metadataContext)).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });
    expect(evaluatePlan(lowerToPlan(parseFormula("LAMBDA(1,1)")), metadataContext)).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });
    expect(evaluatePlan(lowerToPlan(parseFormula('INDIRECT("A1")')), metadataContext)).toEqual({
      tag: ValueTag.Number,
      value: 2,
    });
    expect(
      evaluatePlan(lowerToPlan(parseFormula('INDIRECT("TaxRate")+1')), metadataContext),
    ).toEqual({
      tag: ValueTag.Number,
      value: 1.085,
    });
    expect(
      evaluatePlan(lowerToPlan(parseFormula('INDIRECT("R1C1",FALSE())')), metadataContext),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });
    expect(evaluatePlan(lowerToPlan(parseFormula("INDIRECT()")), metadataContext)).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });
    expect(
      evaluatePlan(
        [
          { opcode: "push-error", code: ErrorCode.Ref },
          { opcode: "call", callee: "INDIRECT", argc: 1 },
          { opcode: "return" },
        ],
        metadataContext,
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    });
    expect(
      evaluatePlan(
        [
          { opcode: "push-string", value: "A1" },
          { opcode: "push-error", code: ErrorCode.Ref },
          { opcode: "call", callee: "INDIRECT", argc: 2 },
          { opcode: "return" },
        ],
        metadataContext,
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    });
    expect(evaluatePlan(lowerToPlan(parseFormula('INDIRECT("")')), metadataContext)).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    });
    expect(evaluatePlan(lowerToPlan(parseFormula('INDIRECT("A:A")')), metadataContext)).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    });
    expect(
      evaluatePlan(
        lowerToPlan(parseFormula('TEXTSPLIT("a,b,,c",",","",TRUE(),0,"-")')),
        metadataContext,
      ),
    ).toEqual({
      tag: ValueTag.String,
      value: "a",
      stringId: 0,
    });
    expect(
      evaluatePlan(lowerToPlan(parseFormula('TEXTSPLIT("alpha","")')), metadataContext),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });
    expect(evaluatePlan(lowerToPlan(parseFormula('TEXTSPLIT("alpha")')), metadataContext)).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });
    expect(
      evaluatePlan(
        [
          { opcode: "push-error", code: ErrorCode.Ref },
          { opcode: "push-string", value: "," },
          { opcode: "call", callee: "TEXTSPLIT", argc: 2 },
          { opcode: "return" },
        ],
        metadataContext,
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    });
    expect(
      evaluatePlan(
        [
          { opcode: "push-string", value: "alpha" },
          { opcode: "push-error", code: ErrorCode.Ref },
          { opcode: "call", callee: "TEXTSPLIT", argc: 2 },
          { opcode: "return" },
        ],
        metadataContext,
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    });
    expect(
      evaluatePlan(
        [
          { opcode: "push-string", value: "alpha" },
          { opcode: "push-string", value: "," },
          { opcode: "push-error", code: ErrorCode.Ref },
          { opcode: "call", callee: "TEXTSPLIT", argc: 3 },
          { opcode: "return" },
        ],
        metadataContext,
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    });
    expect(
      evaluatePlan(
        [
          { opcode: "push-string", value: "alpha" },
          { opcode: "push-string", value: "," },
          { opcode: "push-string", value: "" },
          { opcode: "push-error", code: ErrorCode.Ref },
          { opcode: "call", callee: "TEXTSPLIT", argc: 4 },
          { opcode: "return" },
        ],
        metadataContext,
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    });
    expect(
      evaluatePlan(
        lowerToPlan(parseFormula('TEXTSPLIT("alpha",",","",TRUE(),2)')),
        metadataContext,
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });
    expect(
      evaluatePlan(
        lowerToPlan(parseFormula('TEXTSPLIT("alpha",",","",TRUE(),0,SEQUENCE(2))')),
        metadataContext,
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });
    expect(evaluatePlan(lowerToPlan(parseFormula("EXPAND(A1:B2,3,3,0)")), metadataContext)).toEqual(
      {
        tag: ValueTag.Number,
        value: 2,
      },
    );
    expect(evaluatePlan(lowerToPlan(parseFormula("EXPAND(A1:B2,1,1)")), metadataContext)).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });
    expect(evaluatePlan(lowerToPlan(parseFormula("EXPAND(A1:B2)")), metadataContext)).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });
    expect(evaluatePlan(lowerToPlan(parseFormula("EXPAND(A1:B2,0,3)")), metadataContext)).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });
    expect(evaluatePlan(lowerToPlan(parseFormula("EXPAND(A1:B2,3,0)")), metadataContext)).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });
    expect(
      evaluatePlan(lowerToPlan(parseFormula("EXPAND(A1:B2,3,3,SEQUENCE(2))")), metadataContext),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });
  });

  it("optimizes unary and conditional expressions while preserving dynamic refs", () => {
    expect(optimizeFormula(parseFormula("+A1"))).toEqual({ kind: "CellRef", ref: "A1" });
    expect(optimizeFormula(parseFormula('-"text"'))).toEqual({
      kind: "ErrorLiteral",
      code: ErrorCode.Value,
    });
    expect(optimizeFormula(parseFormula("IF(FALSE, A1, 1+2)"))).toEqual({
      kind: "NumberLiteral",
      value: 3,
    });
    expect(optimizeFormula(parseFormula('IF("", 1, 2)'))).toEqual({
      kind: "NumberLiteral",
      value: 2,
    });
    expect(optimizeFormula(parseFormula('"a"&"b"'))).toEqual({
      kind: "StringLiteral",
      value: "ab",
    });
    expect(optimizeFormula(parseFormula("IF(A1, 1+2, B1)"))).toEqual({
      kind: "CallExpr",
      callee: "IF",
      args: [
        { kind: "CellRef", ref: "A1" },
        { kind: "NumberLiteral", value: 3 },
        { kind: "CellRef", ref: "B1" },
      ],
    });

    expect(optimizeFormula(parseFormula("LET(x,2,x+3)"))).toEqual({
      kind: "NumberLiteral",
      value: 5,
    });
    expect(optimizeFormula(parseFormula("LET(x,A1+1,x+3)"))).toEqual({
      kind: "BinaryExpr",
      operator: "+",
      left: {
        kind: "BinaryExpr",
        operator: "+",
        left: { kind: "CellRef", ref: "A1" },
        right: { kind: "NumberLiteral", value: 1 },
      },
      right: { kind: "NumberLiteral", value: 3 },
    });
    expect(optimizeFormula(parseFormula("LET(x,1,LET(x,2,x+3)+x)"))).toEqual({
      kind: "NumberLiteral",
      value: 6,
    });
    expect(optimizeFormula(parseFormula("LAMBDA(x,x+1)(A1)"))).toEqual({
      kind: "BinaryExpr",
      operator: "+",
      left: { kind: "CellRef", ref: "A1" },
      right: { kind: "NumberLiteral", value: 1 },
    });
    expect(optimizeFormula(parseFormula("LET(fn,LAMBDA(x,x+1),fn(A1))"))).toEqual({
      kind: "BinaryExpr",
      operator: "+",
      left: { kind: "CellRef", ref: "A1" },
      right: { kind: "NumberLiteral", value: 1 },
    });
    expect(optimizeFormula(parseFormula("LET(x,10,LAMBDA(x,x+1)(4)+x)"))).toEqual({
      kind: "NumberLiteral",
      value: 15,
    });
    expect(optimizeFormula(parseFormula("LET(1,2,3)"))).toEqual({
      kind: "ErrorLiteral",
      code: ErrorCode.Value,
    });
    expect(optimizeFormula(parseFormula("LAMBDA(x,x+1)(4,5)"))).toEqual({
      kind: "ErrorLiteral",
      code: ErrorCode.Value,
    });
  });
});

function num(value: number): CellValue {
  return { tag: ValueTag.Number, value };
}

function empty(): CellValue {
  return { tag: ValueTag.Empty };
}
