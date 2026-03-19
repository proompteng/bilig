import { describe, expect, it } from "vitest";
import { compileFormula, evaluateAst, evaluatePlan, parseCellAddress, parseFormula, parseRangeAddress } from "../index.js";
import { ValueTag, type CellValue } from "@bilig/protocol";

describe("formula", () => {
  it("parses A1 addresses", () => {
    expect(parseCellAddress("B12")).toMatchObject({ row: 11, col: 1, text: "B12" });
  });

  it("parses quoted sheet addresses", () => {
    expect(parseCellAddress("'My Sheet'!B12")).toMatchObject({ sheetName: "My Sheet", row: 11, col: 1, text: "B12" });
  });

  it("normalizes ranges", () => {
    expect(parseRangeAddress("B2:A1")).toMatchObject({
      kind: "cells",
      start: { text: "A1" },
      end: { text: "B2" }
    });
  });

  it("normalizes row and column ranges", () => {
    expect(parseRangeAddress("10:1")).toMatchObject({
      kind: "rows",
      start: { text: "1" },
      end: { text: "10" }
    });
    expect(parseRangeAddress("C:A")).toMatchObject({
      kind: "cols",
      start: { text: "A" },
      end: { text: "C" }
    });
  });

  it("compiles arithmetic formulas with wasm-safe mode", () => {
    const compiled = compileFormula("A1*2");
    expect(compiled.mode).toBe(1);
    expect([...compiled.symbolicRefs]).toEqual(["A1"]);
    expect(compiled.maxStackDepth).toBeGreaterThan(0);
    expect(compiled.id).toBe(0);
    expect(compiled.depsPtr).toBe(0);
    expect(compiled.depsLen).toBe(0);
    expect(compiled.programOffset).toBe(0);
    expect(compiled.constNumberOffset).toBe(0);
  });

  it("keeps pass-through cell refs on the JS path", () => {
    const compiled = compileFormula("A1");
    expect(compiled.mode).toBe(0);
  });

  it("compiles bounded aggregate formulas into the wasm-safe path", () => {
    const compiled = compileFormula("SUM(A1:B2)");
    expect(compiled.mode).toBe(1);
    expect([...compiled.symbolicRefs]).toEqual([]);
    expect([...compiled.symbolicRanges]).toEqual(["A1:B2"]);
  });

  it("keeps row and column aggregate formulas on the JS path", () => {
    expect(compileFormula("SUM(A:A)").mode).toBe(0);
    expect(compileFormula("SUM(1:10)").mode).toBe(0);
  });

  it("keeps numeric IF formulas on the JS path until wasm semantics catch up", () => {
    const compiled = compileFormula("IF(A1>0,A1*2,A2-1)");
    expect(compiled.mode).toBe(0);
    expect([...compiled.symbolicRefs]).toEqual([]);
    expect(compiled.deps).toEqual(["A1", "A2"]);
  });

  it("evaluates AST against a context", () => {
    const ast = parseFormula("A1+A2");
    const context = {
      sheetName: "Sheet1",
      resolveCell: (_sheet: string, address: string): CellValue => {
        if (address === "A1") return { tag: ValueTag.Number, value: 2 };
        return { tag: ValueTag.Number, value: 3 };
      },
      resolveRange: (): CellValue[] => []
    };
    const value = evaluateAst(ast, context);
    expect(value).toEqual({ tag: ValueTag.Number, value: 5 });
    expect(evaluatePlan(compileFormula("A1+A2").jsPlan, context)).toEqual({ tag: ValueTag.Number, value: 5 });
  });

  it("parses quoted sheet references inside formulas", () => {
    const ast = parseFormula("'My Sheet'!A1+1");
    expect(ast).toMatchObject({
      kind: "BinaryExpr",
      left: { kind: "CellRef", sheetName: "My Sheet", ref: "A1" }
    });
  });

  it("compiles quoted sheet ranges into symbolic refs", () => {
    const compiled = compileFormula("SUM('My Sheet'!A1:B2)");
    expect([...compiled.symbolicRefs]).toEqual([]);
    expect([...compiled.symbolicRanges]).toEqual(["'My Sheet'!A1:B2"]);
  });

  it("parses quoted sheet column ranges inside formulas", () => {
    const ast = parseFormula("SUM('My Sheet'!A:A)");
    expect(ast).toMatchObject({
      kind: "CallExpr",
      args: [{ kind: "RangeRef", refKind: "cols", sheetName: "My Sheet", start: "A", end: "A" }]
    });
  });

  it("preserves absolute and mixed references in formulas", () => {
    const ast = parseFormula("SUM($A1,B$2,$C$3,$4:5,$D:$F)");
    expect(ast).toMatchObject({
      kind: "CallExpr",
      callee: "SUM",
      args: [
        { kind: "CellRef", ref: "$A1" },
        { kind: "CellRef", ref: "B$2" },
        { kind: "CellRef", ref: "$C$3" },
        { kind: "RangeRef", refKind: "rows", start: "$4", end: "5" },
        { kind: "RangeRef", refKind: "cols", start: "$D", end: "$F" }
      ]
    });
  });

  it("constant folds numeric expressions and prunes IF branches before binding", () => {
    const compiled = compileFormula("IF(TRUE, 1+2*3, A1)");
    expect(compiled.optimizedAst).toEqual({ kind: "NumberLiteral", value: 7 });
    expect(compiled.deps).toEqual([]);
    expect(compiled.jsPlan).toEqual([
      { opcode: "push-number", value: 7 },
      { opcode: "return" }
    ]);
  });

  it("flattens concat calls in the optimized AST", () => {
    const compiled = compileFormula("CONCAT(A1, CONCAT(B1, C1))");
    expect(compiled.optimizedAst).toMatchObject({
      kind: "CallExpr",
      callee: "CONCAT",
      args: [
        { kind: "CellRef", ref: "A1" },
        { kind: "CellRef", ref: "B1" },
        { kind: "CellRef", ref: "C1" }
      ]
    });
  });
});
