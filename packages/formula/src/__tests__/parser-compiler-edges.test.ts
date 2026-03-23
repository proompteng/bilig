import { describe, expect, it } from "vitest";
import { ErrorCode, FormulaMode, Opcode, ValueTag, type CellValue } from "@bilig/protocol";
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
    expect(lexFormula("10%").slice(0, 2)).toEqual([
      { kind: "number", value: "10" },
      { kind: "percent", value: "%" }
    ]);
    expect(lexFormula("\"he said \"\"hi\"\"\"").slice(0, 1)).toEqual([
      { kind: "string", value: "he said \"hi\"" }
    ]);
    expect(() => lexFormula("@oops")).toThrow("Unexpected token '@'");
  });

  it("rejects standalone axis refs and malformed ranges", () => {
    expect(() => parseFormula("'Sheet 1'!1")).toThrow("Row and column references must appear inside a range");
    expect(() => parseFormula("'Sheet 1'!$1")).toThrow("Row and column references must appear inside a range");
    expect(() => parseFormula("A1:B")).toThrow("Range endpoints must use the same reference type");
    expect(() => parseFormula("A1:2")).toThrow("Range endpoints must use the same reference type");
  });

  it("binds quoted ranges and keeps unsupported/text formulas on the JS path", () => {
    const quotedRange = bindFormula(parseFormula("SUM('My Sheet'!A1:A2)"));
    expect(quotedRange.deps).toEqual(["'My Sheet'!A1:A2"]);
    expect(quotedRange.mode).toBe(FormulaMode.WasmFastPath);

    const anchoredRange = bindFormula(parseFormula("SUM('My Sheet'!$A:$B)"));
    expect(anchoredRange.deps).toEqual(["'My Sheet'!A:B"]);
    expect(anchoredRange.mode).toBe(FormulaMode.WasmFastPath);

    expect(bindFormula(parseFormula("\"hello\"")).mode).toBe(FormulaMode.WasmFastPath);
    expect(bindFormula(parseFormula("A1")).mode).toBe(FormulaMode.WasmFastPath);
    expect(bindFormula(parseFormula("LEN(A1)")).mode).toBe(FormulaMode.WasmFastPath);
    expect(bindFormula(parseFormula("LEN(A1:A2)")).mode).toBe(FormulaMode.JsOnly);
  });

  it("parses defined names, tracks them separately, and lowers them onto the JS plan", () => {
    const ast = parseFormula("TaxRate*A1");
    expect(ast).toEqual({
      kind: "BinaryExpr",
      operator: "*",
      left: { kind: "NameRef", name: "TaxRate" },
      right: { kind: "CellRef", ref: "A1" }
    });

    const bound = bindFormula(ast);
    expect(bound.deps).toEqual(["A1"]);
    expect(bound.symbolicNames).toEqual(["TaxRate"]);
    expect(bound.mode).toBe(FormulaMode.JsOnly);

    const compiled = compileFormula("TaxRate*A1");
    expect(compiled.symbolicNames).toEqual(["TaxRate"]);
    expect(compiled.jsPlan).toEqual([
      { opcode: "push-name", name: "TaxRate" },
      { opcode: "push-cell", address: "A1" },
      { opcode: "binary", operator: "*" },
      { opcode: "return" }
    ]);
  });

  it("parses structured references and spill refs as metadata-aware syntax", () => {
    const structured = parseFormula("SUM(Sales[Amount])");
    expect(structured).toEqual({
      kind: "CallExpr",
      callee: "SUM",
      args: [{ kind: "StructuredRef", tableName: "Sales", columnName: "Amount" }]
    });

    const structuredBound = bindFormula(structured);
    expect(structuredBound.symbolicTables).toEqual(["Sales"]);
    expect(structuredBound.mode).toBe(FormulaMode.JsOnly);

    const spill = parseFormula("A1#");
    expect(spill).toEqual({ kind: "SpillRef", ref: "A1" });

    const spillBound = bindFormula(spill);
    expect(spillBound.symbolicSpills).toEqual(["A1"]);
    expect(spillBound.mode).toBe(FormulaMode.JsOnly);
  });

  it("throws on unsupported wasm builtin encodings and invalid axis compilation", () => {
    expect(isBuiltinAvailable("SUM")).toBe(true);
    expect(isBuiltinAvailable("MATCH")).toBe(true);
    expect(isBuiltinAvailable("INDEX")).toBe(true);
    expect(isBuiltinAvailable("VLOOKUP")).toBe(true);
    expect(isBuiltinAvailable("DOES_NOT_EXIST")).toBe(false);
    expect(encodeBuiltin("LEN")).toBeDefined();
    expect(() => compileFormula("Sheet1!A")).toThrow("Row and column references must appear inside a range");
  });

  it("compiles literal text, CONCAT, and IF text branches onto the wasm path", () => {
    const textIf = compileFormula("IF(A1, CONCAT(\"x\", \"y\"), \"z\")");
    expect(textIf.mode).toBe(FormulaMode.WasmFastPath);
    expect(textIf.program[0] >>> 24).toBe(Opcode.PushCell);
    expect(textIf.symbolicStrings).toEqual(["xy", "z"]);

    const plainString = compileFormula("\"hello\"");
    expect(plainString.mode).toBe(FormulaMode.WasmFastPath);
    expect(plainString.symbolicStrings).toEqual(["hello"]);
    expect(plainString.program).toEqual(Uint32Array.from([(Opcode.PushString << 24) | 0, 255 << 24]));
    expect(plainString.jsPlan).toEqual([
      { opcode: "push-string", value: "hello" },
      { opcode: "return" }
    ]);

    const concat = compileFormula("CONCAT(\"x\", \"y\")");
    expect(concat.mode).toBe(FormulaMode.WasmFastPath);
    expect(concat.symbolicStrings).toEqual(["xy"]);

    const compared = compileFormula("A1=\"HELLO\"");
    expect(compared.mode).toBe(FormulaMode.WasmFastPath);
  });

  it("parses postfix percent as arithmetic scaling", () => {
    expect(parseFormula("10%")).toEqual({
      kind: "BinaryExpr",
      operator: "*",
      left: { kind: "NumberLiteral", value: 10 },
      right: { kind: "NumberLiteral", value: 0.01 }
    });

    expect(parseFormula("(A1+A2)%")).toEqual({
      kind: "BinaryExpr",
      operator: "*",
      left: {
        kind: "BinaryExpr",
        operator: "+",
        left: { kind: "CellRef", ref: "A1" },
        right: { kind: "CellRef", ref: "A2" }
      },
      right: { kind: "NumberLiteral", value: 0.01 }
    });
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

  it("parses and lowers lambda invocation syntax", () => {
    expect(parseFormula("LAMBDA(x,x+1)(4)")).toEqual({
      kind: "InvokeExpr",
      callee: {
        kind: "CallExpr",
        callee: "LAMBDA",
        args: [
          { kind: "NameRef", name: "x" },
          {
            kind: "BinaryExpr",
            operator: "+",
            left: { kind: "NameRef", name: "x" },
            right: { kind: "NumberLiteral", value: 1 }
          }
        ]
      },
      args: [{ kind: "NumberLiteral", value: 4 }]
    });

    expect(lowerToPlan(parseFormula("LAMBDA(x,x+1)(4)"))).toEqual([
      {
        opcode: "push-lambda",
        params: ["x"],
        body: [
          { opcode: "push-name", name: "x" },
          { opcode: "push-number", value: 1 },
          { opcode: "binary", operator: "+" },
          { opcode: "return" }
        ]
      },
      { opcode: "push-number", value: 4 },
      { opcode: "invoke", argc: 1 },
      { opcode: "return" }
    ]);
  });
});
