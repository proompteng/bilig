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

  it("compiles postfix percent arithmetic through the existing numeric pipeline", () => {
    const compiled = compileFormula("A1*10%");
    expect(compiled.mode).toBe(1);
    expect([...compiled.symbolicRefs]).toEqual(["A1"]);
  });

  it("keeps pass-through cell refs on the wasm-safe path", () => {
    const compiled = compileFormula("A1");
    expect(compiled.mode).toBe(1);
  });

  it("compiles bounded aggregate formulas into the wasm-safe path", () => {
    const compiled = compileFormula("SUM(A1:B2)");
    expect(compiled.mode).toBe(1);
    expect([...compiled.symbolicRefs]).toEqual([]);
    expect([...compiled.symbolicRanges]).toEqual(["A1:B2"]);
  });

  it("compiles exact-parity logical and rounding builtins onto the wasm-safe path", () => {
    expect(compileFormula("AND(A1,TRUE)").mode).toBe(1);
    expect(compileFormula("OR(A1,FALSE)").mode).toBe(1);
    expect(compileFormula("NOT(A1)").mode).toBe(1);
    expect(compileFormula("ROUND(A1,1)").mode).toBe(1);
    expect(compileFormula("FLOOR(A1,2)").mode).toBe(1);
    expect(compileFormula("CEILING(A1,2)").mode).toBe(1);
  });

  it("compiles exact-parity info and date builtins onto the wasm-safe path", () => {
    expect(compileFormula("ISBLANK()").mode).toBe(1);
    expect(compileFormula("ISBLANK(A1)").mode).toBe(1);
    expect(compileFormula("ISNUMBER()").mode).toBe(1);
    expect(compileFormula("ISNUMBER(A1)").mode).toBe(1);
    expect(compileFormula("ISTEXT()").mode).toBe(1);
    expect(compileFormula("ISTEXT(A1)").mode).toBe(1);
    expect(compileFormula("LEN(A1)").mode).toBe(1);
    expect(compileFormula("DATE(2024,2,29)").mode).toBe(1);
    expect(compileFormula("TIME(12,30,0)").mode).toBe(1);
    expect(compileFormula("YEAR(A1)").mode).toBe(1);
    expect(compileFormula("MONTH(A1)").mode).toBe(1);
    expect(compileFormula("DAY(A1)").mode).toBe(1);
    expect(compileFormula("HOUR(A1)").mode).toBe(1);
    expect(compileFormula("MINUTE(A1)").mode).toBe(1);
    expect(compileFormula("SECOND(A1)").mode).toBe(1);
    expect(compileFormula("WEEKDAY(A1)").mode).toBe(1);
    expect(compileFormula("WEEKDAY(A1,2)").mode).toBe(1);
    expect(compileFormula("EDATE(A1,1)").mode).toBe(1);
    expect(compileFormula("EOMONTH(A1,1)").mode).toBe(1);
    expect(compileFormula("EXACT(A1,A2)").mode).toBe(1);
    expect(compileFormula("VALUE(\"42\")").mode).toBe(1);
    expect(compileFormula("INT(A1)").mode).toBe(1);
    expect(compileFormula("ROUNDUP(A1,2)").mode).toBe(1);
    expect(compileFormula("ROUNDDOWN(A1,2)").mode).toBe(1);
    expect(compileFormula("LEFT(A1,2)").mode).toBe(1);
    expect(compileFormula("RIGHT(A1,2)").mode).toBe(1);
    expect(compileFormula("MID(A1,2,3)").mode).toBe(1);
    expect(compileFormula("TRIM(A1)").mode).toBe(1);
    expect(compileFormula("UPPER(A1)").mode).toBe(1);
    expect(compileFormula("LOWER(A1)").mode).toBe(1);
    expect(compileFormula("FIND(\"a\",A1)").mode).toBe(1);
    expect(compileFormula("SEARCH(\"a\",A1)").mode).toBe(1);
  });

  it("keeps LEN on the JS path when it depends on a range until the range-string bridge lands", () => {
    expect(compileFormula("LEN(A1:B2)").mode).toBe(0);
    expect(compileFormula("VALUE(A1)").mode).toBe(1);
  });

  it("routes volatile scalar builtins through the wasm path", () => {
    expect(compileFormula("TODAY()").mode).toBe(1);
    expect(compileFormula("NOW()").mode).toBe(1);
    expect(compileFormula("RAND()").mode).toBe(1);
  });

  it("routes native sequence spills through the wasm path, including numeric aggregate consumers", () => {
    expect(compileFormula("SEQUENCE(3,1,1,1)")).toMatchObject({ mode: 1, producesSpill: true });
    expect(compileFormula("SUM(SEQUENCE(A1,1,1,1))").mode).toBe(1);
    expect(compileFormula("AVG(SEQUENCE(A1,1,1,1))").mode).toBe(1);
    expect(compileFormula("MIN(SEQUENCE(A1,1,1,1))").mode).toBe(1);
    expect(compileFormula("MAX(SEQUENCE(A1,1,1,1))").mode).toBe(1);
    expect(compileFormula("COUNT(SEQUENCE(A1,1,1,1))").mode).toBe(1);
    expect(compileFormula("COUNTA(SEQUENCE(A1,1,1,1))").mode).toBe(1);
  });

  it("keeps row and column aggregate formulas on the JS path", () => {
    expect(compileFormula("SUM(A:A)").mode).toBe(1);
    expect(compileFormula("SUM(1:10)").mode).toBe(1);
  });

  it("compiles IF, IFERROR, IFNA, and NA onto the wasm-safe path alongside exact-parity logical formulas", () => {
    const compiled = compileFormula("IF(A1>0,A1*2,A2-1)");
    expect(compiled.mode).toBe(1);
    expect(compileFormula("IFERROR(A1,\"missing\")").mode).toBe(1);
    expect(compileFormula("IFNA(NA(),\"missing\")").mode).toBe(1);
    expect(compileFormula("NA()").mode).toBe(1);
    expect(compileFormula("COUNTIF(A1:A4,\">0\")").mode).toBe(1);
    expect(compileFormula("COUNTIFS(A1:A4,\">0\",B1:B4,\"x\")").mode).toBe(1);
    expect(compileFormula("SUMIF(A1:A4,\">0\",B1:B4)").mode).toBe(1);
    expect(compileFormula("SUMIFS(C1:C4,A1:A4,\">0\",B1:B4,\"x\")").mode).toBe(1);
    expect(compileFormula("AVERAGEIF(A1:A4,\">0\")").mode).toBe(1);
    expect(compileFormula("AVERAGEIFS(C1:C4,A1:A4,\">0\",B1:B4,\"x\")").mode).toBe(1);
    expect(compileFormula("SUMPRODUCT(A1:A3,B1:B3)").mode).toBe(1);
    expect(compileFormula("MATCH(\"pear\",A1:A3,0)").mode).toBe(1);
    expect(compileFormula("XMATCH(\"pear\",A1:A3)").mode).toBe(1);
    expect(compileFormula("XLOOKUP(\"pear\",A1:A3,B1:B3)").mode).toBe(1);
    expect(compileFormula("INDEX(A1:B3,2,2)").mode).toBe(1);
    expect(compileFormula("VLOOKUP(\"pear\",A1:B3,2,FALSE)").mode).toBe(1);
    expect(compileFormula("HLOOKUP(\"pear\",A1:C2,2,FALSE)").mode).toBe(1);

    expect(compileFormula("AND(A1,TRUE)").mode).toBe(1);
    expect(compileFormula("OR(A1,FALSE)").mode).toBe(1);
    expect(compileFormula("NOT(A1)").mode).toBe(1);
    expect(compileFormula("ROUND(A1,-1)").mode).toBe(1);
    expect(compileFormula("FLOOR(A1,2)").mode).toBe(1);
    expect(compileFormula("CEILING(A1,2)").mode).toBe(1);
  });

  it("keeps unsupported candidate builtin arities on the JS path", () => {
    expect(compileFormula("IF(A1,1)").mode).toBe(0);
    expect(compileFormula("NOT(A1,A2)").mode).toBe(0);
    expect(compileFormula("COUNTIF(A:A,\">0\")").mode).toBe(0);
    expect(compileFormula("SUMIF(A1:A4,B1:B4,C1:C4,D1:D4)").mode).toBe(0);
    expect(compileFormula("SUMIFS(A1:A4,\">0\")").mode).toBe(0);
    expect(compileFormula("MATCH(\"pear\",A1:B3,0)").mode).toBe(0);
    expect(compileFormula("XLOOKUP(\"pear\",A1:B3,C1:D4)").mode).toBe(0);
    expect(compileFormula("VLOOKUP(\"pear\",A1:B3,2,FALSE,1)").mode).toBe(0);
    expect(compileFormula("ROUND(A1,A2,A3)").mode).toBe(0);
    expect(compileFormula("FLOOR(A1,A2,A3)").mode).toBe(0);
    expect(compileFormula("CEILING(A1,A2,A3)").mode).toBe(0);
    expect(compileFormula("TIME(A1,A2)").mode).toBe(0);
    expect(compileFormula("WEEKDAY(A1,A2,A3)").mode).toBe(0);
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
    expect(evaluateAst(parseFormula("10%"), context)).toEqual({ tag: ValueTag.Number, value: 0.1 });
    expect(evaluateAst(parseFormula("(A1+A2)%"), context)).toEqual({ tag: ValueTag.Number, value: 0.05 });
    expect(evaluatePlan(compileFormula("A1*10%").jsPlan, context)).toEqual({ tag: ValueTag.Number, value: 0.2 });
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
