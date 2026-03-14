import { describe, expect, it } from "vitest";
import { ErrorCode, FormulaMode, ValueTag, type CellValue } from "@bilig/protocol";
import {
  bindFormula,
  compileFormula,
  encodeBuiltin,
  evaluatePlan,
  isBuiltinAvailable,
  lexFormula,
  lowerToPlan,
  parseFormula
} from "../index.js";

const context = {
  sheetName: "Sheet1",
  resolveCell: (_sheetName: string, address: string): CellValue => {
    switch (address) {
      case "A1":
        return { tag: ValueTag.Number, value: 4 };
      case "B1":
        return { tag: ValueTag.Boolean, value: true };
      default:
        return { tag: ValueTag.Empty };
    }
  },
  resolveRange: (_sheetName: string, start: string, end: string, refKind: "cells" | "rows" | "cols"): CellValue[] => {
    if (refKind === "cells" && start === "A1" && end === "A2") {
      return [
        { tag: ValueTag.Number, value: 4 },
        { tag: ValueTag.Number, value: 6 }
      ];
    }
    return [];
  }
};

describe("formula parser/compiler edges", () => {
  it("lexes escaped quoted sheet names and rejects invalid tokens", () => {
    expect(lexFormula("'O''Brien'!A1").slice(0, 3)).toEqual([
      { kind: "quotedIdentifier", value: "O'Brien" },
      { kind: "bang", value: "!" },
      { kind: "identifier", value: "A1" }
    ]);
    expect(() => lexFormula("@oops")).toThrow("Unexpected token '@'");
  });

  it("rejects standalone axis refs and malformed ranges", () => {
    expect(() => parseFormula("'Sheet 1'!1")).toThrow("Row and column references must appear inside a range");
    expect(() => parseFormula("A1:B")).toThrow("Range endpoints must use the same reference type");
    expect(() => parseFormula("A1:2")).toThrow("Range endpoints must use the same reference type");
    expect(() => parseFormula("A1X")).toThrow("Unsupported reference 'A1X'");
  });

  it("binds quoted ranges and keeps unsupported/text formulas on the JS path", () => {
    const quotedRange = bindFormula(parseFormula("SUM('My Sheet'!A1:A2)"));
    expect(quotedRange.deps).toEqual(["'My Sheet'!A1:A2"]);
    expect(quotedRange.mode).toBe(FormulaMode.WasmFastPath);

    expect(bindFormula(parseFormula("\"hello\"")).mode).toBe(FormulaMode.JsOnly);
    expect(bindFormula(parseFormula("A1")).mode).toBe(FormulaMode.JsOnly);
    expect(bindFormula(parseFormula("LEN(A1)")).mode).toBe(FormulaMode.JsOnly);
  });

  it("throws on unsupported wasm builtin encodings and invalid axis compilation", () => {
    expect(isBuiltinAvailable("SUM")).toBe(true);
    expect(isBuiltinAvailable("DOES_NOT_EXIST")).toBe(false);
    expect(() => encodeBuiltin("LEN")).toThrow("Unsupported builtin for wasm: LEN");
    expect(() => compileFormula("A")).toThrow("Row and column references must appear inside a range");
  });

  it("compiles non-wasm IF and text formulas onto the JS plan only", () => {
    const textIf = compileFormula("IF(A1, CONCAT(\"x\", \"y\"), \"z\")");
    expect(textIf.mode).toBe(FormulaMode.JsOnly);
    expect(textIf.program).toEqual(Uint32Array.from([255 << 24]));

    const plainString = compileFormula("\"hello\"");
    expect(plainString.mode).toBe(FormulaMode.JsOnly);
    expect(plainString.jsPlan).toEqual([
      { opcode: "push-string", value: "hello" },
      { opcode: "return" }
    ]);
  });

  it("evaluates lowered plans across comparison, unary, jump, and builtin error paths", () => {
    expect(
      evaluatePlan(
        [
          { opcode: "push-cell", address: "A1" },
          { opcode: "push-number", value: 4 },
          { opcode: "binary", operator: "=" },
          { opcode: "return" }
        ],
        context
      )
    ).toEqual({ tag: ValueTag.Boolean, value: true });

    expect(
      evaluatePlan(
        [
          { opcode: "push-string", value: "x" },
          { opcode: "unary", operator: "-" },
          { opcode: "return" }
        ],
        context
      )
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });

    expect(
      evaluatePlan(
        [
          { opcode: "push-boolean", value: false },
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
          { opcode: "push-range", start: "A1", end: "A2", refKind: "cells" },
          { opcode: "call", callee: "SUM", argc: 1 },
          { opcode: "return" }
        ],
        context
      )
    ).toEqual({ tag: ValueTag.Number, value: 10 });

    expect(
      evaluatePlan(
        [
          { opcode: "push-number", value: 1 },
          { opcode: "call", callee: "UNKNOWN", argc: 1 },
          { opcode: "return" }
        ],
        context
      )
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Name });
  });

  it("lowers full IF expressions into explicit jump instructions", () => {
    expect(lowerToPlan(parseFormula("IF(A1, B1, 0)"))).toEqual([
      { opcode: "push-cell", address: "A1" },
      { opcode: "jump-if-false", target: 4 },
      { opcode: "push-cell", address: "B1" },
      { opcode: "jump", target: 5 },
      { opcode: "push-number", value: 0 },
      { opcode: "return" }
    ]);
  });
});
