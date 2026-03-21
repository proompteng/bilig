import { describe, expect, it } from "vitest";
import { ErrorCode, ValueTag, type CellValue } from "@bilig/protocol";
import { evaluatePlan, lowerToPlan, optimizeFormula, parseFormula } from "../index.js";
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
        { tag: ValueTag.Empty }
      ];
    }
    return [];
  }
};

describe("js evaluator", () => {
  it("evaluates direct plans for ranges, jumps, and fallback stack handling", () => {
    expect(
      evaluatePlan(
        [
          { opcode: "push-range", start: "A1", end: "B2", refKind: "cells" },
          { opcode: "call", callee: "SUM", argc: 1 },
          { opcode: "return" }
        ],
        context
      )
    ).toEqual({ tag: ValueTag.Number, value: 6 });

    expect(
      evaluatePlan(
        [
          { opcode: "push-range", start: "A1", end: "B2", refKind: "cells" },
          { opcode: "return" }
        ],
        context
      )
    ).toEqual({ tag: ValueTag.Number, value: 2 });

    expect(
      evaluatePlan(
        [
          { opcode: "push-number", value: 0 },
          { opcode: "jump-if-false", target: 4 },
          { opcode: "push-number", value: 1 },
          { opcode: "jump", target: 5 },
          { opcode: "push-number", value: 2 },
          { opcode: "return" }
        ],
        context
      )
    ).toEqual({ tag: ValueTag.Number, value: 2 });

    expect(
      evaluatePlan(
        [
          { opcode: "call", callee: "SUM", argc: 1 },
          { opcode: "return" }
        ],
        context
      )
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });

    expect(
      evaluatePlan(
        [
          { opcode: "call", callee: "DOES_NOT_EXIST", argc: 0 },
          { opcode: "return" }
        ],
        context
      )
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Name });

    expect(
      evaluatePlan(
        [
          { opcode: "push-number", value: 9 }
        ],
        context
      )
    ).toEqual({ tag: ValueTag.Number, value: 9 });
  });

  it("keeps range shape for lookup/reference builtins", () => {
    expect(
      evaluatePlan(
        [
          { opcode: "push-number", value: 3 },
          { opcode: "push-range", start: "A1", end: "A4", refKind: "cells" },
          { opcode: "push-number", value: 0 },
          { opcode: "call", callee: "MATCH", argc: 3 },
          { opcode: "return" }
        ],
        {
          ...context,
          resolveRange: (_sheetName: string, start: string, end: string): CellValue[] => {
            if (start === "A1" && end === "A4") {
              return [num(2), num(3), num(4), num(5)];
            }
            return [];
          }
        }
      )
    ).toEqual({ tag: ValueTag.Number, value: 2 });

    expect(
      evaluatePlan(
        [
          { opcode: "push-range", start: "A1", end: "B2", refKind: "cells" },
          { opcode: "push-number", value: 2 },
          { opcode: "push-number", value: 2 },
          { opcode: "call", callee: "INDEX", argc: 3 },
          { opcode: "return" }
        ],
        context
      )
    ).toEqual({ tag: ValueTag.Empty });
  });

  it("lowers row and column refs into NaN sentinels for the JS path", () => {
    expect(lowerToPlan({ kind: "RowRef", ref: "3" } as FormulaNode)).toEqual([
      { opcode: "push-number", value: Number.NaN },
      { opcode: "return" }
    ]);
    expect(lowerToPlan({ kind: "ColumnRef", ref: "C" } as FormulaNode)).toEqual([
      { opcode: "push-number", value: Number.NaN },
      { opcode: "return" }
    ]);
  });

  it("resolves scalar defined names and returns #NAME? when missing", () => {
    expect(
      evaluatePlan(
        [
          { opcode: "push-name", name: "TaxRate" },
          { opcode: "push-number", value: 1 },
          { opcode: "binary", operator: "+" },
          { opcode: "return" }
        ],
        {
          ...context,
          resolveName: (name: string): CellValue =>
            name === "TaxRate" ? { tag: ValueTag.Number, value: 0.5 } : { tag: ValueTag.Error, code: ErrorCode.Name }
        }
      )
    ).toEqual({ tag: ValueTag.Number, value: 1.5 });

    expect(evaluatePlan([{ opcode: "push-name", name: "MissingName" }, { opcode: "return" }], context)).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Name
    });
  });

  it("optimizes unary and conditional expressions while preserving dynamic refs", () => {
    expect(optimizeFormula(parseFormula("+A1"))).toEqual({ kind: "CellRef", ref: "A1" });
    expect(optimizeFormula(parseFormula("-\"text\""))).toEqual({
      kind: "UnaryExpr",
      operator: "-",
      argument: { kind: "StringLiteral", value: "text" }
    });
    expect(optimizeFormula(parseFormula("IF(FALSE, A1, 1+2)"))).toEqual({
      kind: "NumberLiteral",
      value: 3
    });
    expect(optimizeFormula(parseFormula("IF(\"\", 1, 2)"))).toEqual({
      kind: "NumberLiteral",
      value: 2
    });
    expect(optimizeFormula(parseFormula("\"a\"&\"b\""))).toEqual({
      kind: "StringLiteral",
      value: "ab"
    });
    expect(optimizeFormula(parseFormula("IF(A1, 1+2, B1)"))).toEqual({
      kind: "CallExpr",
      callee: "IF",
      args: [
        { kind: "CellRef", ref: "A1" },
        { kind: "NumberLiteral", value: 3 },
        { kind: "CellRef", ref: "B1" }
      ]
    });
  });
});

function num(value: number): CellValue {
  return { tag: ValueTag.Number, value };
}
