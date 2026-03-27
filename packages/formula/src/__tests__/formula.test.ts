import { describe, expect, it } from "vitest";
import {
  compileFormula,
  evaluateAst,
  evaluatePlan,
  parseCellAddress,
  parseFormula,
  parseRangeAddress,
} from "../index.js";
import { ValueTag, type CellValue } from "@bilig/protocol";

describe("formula", () => {
  it("parses A1 addresses", () => {
    expect(parseCellAddress("B12")).toMatchObject({ row: 11, col: 1, text: "B12" });
  });

  it("parses quoted sheet addresses", () => {
    expect(parseCellAddress("'My Sheet'!B12")).toMatchObject({
      sheetName: "My Sheet",
      row: 11,
      col: 1,
      text: "B12",
    });
  });

  it("normalizes ranges", () => {
    expect(parseRangeAddress("B2:A1")).toMatchObject({
      kind: "cells",
      start: { text: "A1" },
      end: { text: "B2" },
    });
  });

  it("normalizes row and column ranges", () => {
    expect(parseRangeAddress("10:1")).toMatchObject({
      kind: "rows",
      start: { text: "1" },
      end: { text: "10" },
    });
    expect(parseRangeAddress("C:A")).toMatchObject({
      kind: "cols",
      start: { text: "A" },
      end: { text: "C" },
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
    expect(compileFormula("SIN(A1)").mode).toBe(1);
    expect(compileFormula("ATAN2(A1,A2)").mode).toBe(1);
    expect(compileFormula("LOG(A1,10)").mode).toBe(1);
    expect(compileFormula("PI()").mode).toBe(1);
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
    expect(compileFormula("DAYS(A1,A2)").mode).toBe(1);
    expect(compileFormula("WEEKNUM(A1)").mode).toBe(1);
    expect(compileFormula("WORKDAY(A1,1)").mode).toBe(1);
    expect(compileFormula("WORKDAY(A1,1,B1)").mode).toBe(1);
    expect(compileFormula("NETWORKDAYS(A1,A2)").mode).toBe(1);
    expect(compileFormula("NETWORKDAYS(A1,A2,B1)").mode).toBe(1);
    expect(compileFormula("EDATE(A1,1)").mode).toBe(1);
    expect(compileFormula("EOMONTH(A1,1)").mode).toBe(1);
    expect(compileFormula("EXACT(A1,A2)").mode).toBe(1);
    expect(compileFormula('VALUE("42")').mode).toBe(1);
    expect(compileFormula("INT(A1)").mode).toBe(1);
    expect(compileFormula("ROUNDUP(A1,2)").mode).toBe(1);
    expect(compileFormula("ROUNDDOWN(A1,2)").mode).toBe(1);
    expect(compileFormula("LEFT(A1,2)").mode).toBe(1);
    expect(compileFormula("RIGHT(A1,2)").mode).toBe(1);
    expect(compileFormula("MID(A1,2,3)").mode).toBe(1);
    expect(compileFormula("TRIM(A1)").mode).toBe(1);
    expect(compileFormula("UPPER(A1)").mode).toBe(1);
    expect(compileFormula("LOWER(A1)").mode).toBe(1);
    expect(compileFormula('FIND("a",A1)').mode).toBe(1);
    expect(compileFormula('SEARCH("a",A1)').mode).toBe(1);
    expect(compileFormula('REPLACE(A1,2,3,"x")').mode).toBe(1);
    expect(compileFormula('SUBSTITUTE(A1,"a","b")').mode).toBe(1);
    expect(compileFormula("REPT(A1,3)").mode).toBe(1);
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

  it("routes dynamic-array family builtins to the wasm path for numeric-compatible inputs", () => {
    expect(compileFormula("OFFSET(A1:B4,0,0,2,2)")).toMatchObject({ mode: 1, producesSpill: true });
    expect(compileFormula("TAKE(A1:B4,2)")).toMatchObject({ mode: 1, producesSpill: true });
    expect(compileFormula("DROP(A1:B4,1)")).toMatchObject({ mode: 1, producesSpill: true });
    expect(compileFormula("CHOOSECOLS(A1:B4,2)")).toMatchObject({ mode: 1, producesSpill: true });
    expect(compileFormula("CHOOSEROWS(A1:B4,2)")).toMatchObject({ mode: 1, producesSpill: true });
    expect(compileFormula("SORT(A1:B4)")).toMatchObject({ mode: 1, producesSpill: true });
    expect(compileFormula("SORTBY(A1:A4,A1:A4)")).toMatchObject({ mode: 1, producesSpill: true });
    expect(compileFormula("TOCOL(A1:B4)")).toMatchObject({ mode: 1, producesSpill: true });
    expect(compileFormula("TOROW(A1:B4)")).toMatchObject({ mode: 1, producesSpill: true });
    expect(compileFormula("WRAPROWS(A1:B4,2)")).toMatchObject({ mode: 1, producesSpill: true });
    expect(compileFormula("WRAPCOLS(A1:B4,2)")).toMatchObject({ mode: 1, producesSpill: true });
  });

  it("keeps JS-only text-splitting and indirection formulas off the wasm path while preserving spill metadata", () => {
    expect(compileFormula('TEXTSPLIT(A1,",")')).toMatchObject({ mode: 0, producesSpill: true });
    expect(compileFormula('INDIRECT("A1")').mode).toBe(0);
    expect(compileFormula("FORMULA(A1)").mode).toBe(0);
  });

  it("routes accelerated array-shape helpers to the wasm path", () => {
    expect(compileFormula("EXPAND(A1:B2,3,3)")).toMatchObject({ mode: 1, producesSpill: true });
    expect(compileFormula("TRIMRANGE(A1:C4)")).toMatchObject({ mode: 1, producesSpill: true });
    expect(compileFormula("TRIMRANGE(EXPAND(A1:B2,3,3,0))")).toMatchObject({
      mode: 1,
      producesSpill: true,
    });
  });

  it("routes accelerated array-shape and conditional aggregate builtins by public compile contract", () => {
    expect(compileFormula("TRANSPOSE(A1:B4)").mode).toBe(1);
    expect(compileFormula("HSTACK(A1:B2,C1:D2)").mode).toBe(1);
    expect(compileFormula("VSTACK(A1:B2,C1:D2)").mode).toBe(1);
    expect(compileFormula("AREAS(A1:B4)").mode).toBe(1);
    expect(compileFormula("ROWS(A1:B4)").mode).toBe(1);
    expect(compileFormula("COLUMNS(A1:B4)").mode).toBe(1);
    expect(compileFormula("ARRAYTOTEXT(A1:B4)").mode).toBe(1);
    expect(compileFormula("ARRAYTOTEXT(A1:B4,1)").mode).toBe(1);
    expect(compileFormula('MINIFS(A1:A4,B1:B4,">0")').mode).toBe(1);
    expect(compileFormula('MAXIFS(A1:A4,B1:B4,">0")').mode).toBe(1);

    expect(compileFormula("TAKE(A1,1)").mode).toBe(0);
    expect(compileFormula("DROP(A1,1)").mode).toBe(0);
    expect(compileFormula("CHOOSECOLS(A1,1)").mode).toBe(0);
    expect(compileFormula("CHOOSEROWS(A1,1)").mode).toBe(0);
    expect(compileFormula("SORT(A1)").mode).toBe(0);
    expect(compileFormula("TOCOL(A1)").mode).toBe(0);
    expect(compileFormula("TOROW(A1)").mode).toBe(0);
    expect(compileFormula("WRAPROWS(A1,2)").mode).toBe(0);
    expect(compileFormula("WRAPCOLS(A1,2)").mode).toBe(0);
    expect(compileFormula("LOOKUP(A1)").mode).toBe(0);
    expect(compileFormula("AREAS(A1)").mode).toBe(0);
    expect(compileFormula("ROWS(A1)").mode).toBe(0);
    expect(compileFormula("COLUMNS(A1)").mode).toBe(0);
    expect(compileFormula("ARRAYTOTEXT(A1:B4,1,2)").mode).toBe(0);
    expect(compileFormula("MINIFS(A1:A4,B1:B4)").mode).toBe(0);
    expect(compileFormula('MAXIFS(A1,B1:B4,">0")').mode).toBe(0);
    expect(compileFormula("SORTBY(A1:A4)").mode).toBe(0);
  });

  it("keeps row and column aggregate formulas on the JS path", () => {
    expect(compileFormula("SUM(A:A)").mode).toBe(1);
    expect(compileFormula("SUM(1:10)").mode).toBe(1);
  });

  it("keeps contextual metadata functions on the JS path without constant-folding them to errors", () => {
    const rowCompiled = compileFormula("ROW()");
    const sheetCompiled = compileFormula('SHEET("Sheet2")');

    expect(rowCompiled.mode).toBe(0);
    expect(sheetCompiled.mode).toBe(0);
    expect(
      evaluatePlan(rowCompiled.jsPlan, {
        sheetName: "Sheet1",
        currentAddress: "C7",
        resolveCell: (): CellValue => ({ tag: ValueTag.Empty }),
        resolveRange: (): CellValue[] => [],
      }),
    ).toEqual({ tag: ValueTag.Number, value: 7 });
    expect(
      evaluatePlan(sheetCompiled.jsPlan, {
        sheetName: "Sheet1",
        resolveCell: (): CellValue => ({ tag: ValueTag.Empty }),
        resolveRange: (): CellValue[] => [],
        listSheetNames: (): string[] => ["Sheet1", "Sheet2"],
      }),
    ).toEqual({ tag: ValueTag.Number, value: 2 });
  });

  it("compiles IF, IFERROR, IFNA, and NA onto the wasm-safe path alongside exact-parity logical formulas", () => {
    const compiled = compileFormula("IF(A1>0,A1*2,A2-1)");
    expect(compiled.mode).toBe(1);
    expect(compileFormula('IFERROR(A1,"missing")').mode).toBe(1);
    expect(compileFormula('IFNA(NA(),"missing")').mode).toBe(1);
    expect(compileFormula("NA()").mode).toBe(1);
    expect(compileFormula('COUNTIF(A1:A4,">0")').mode).toBe(1);
    expect(compileFormula('COUNTIFS(A1:A4,">0",B1:B4,"x")').mode).toBe(1);
    expect(compileFormula('SUMIF(A1:A4,">0",B1:B4)').mode).toBe(1);
    expect(compileFormula('SUMIFS(C1:C4,A1:A4,">0",B1:B4,"x")').mode).toBe(1);
    expect(compileFormula('AVERAGEIF(A1:A4,">0")').mode).toBe(1);
    expect(compileFormula('AVERAGEIFS(C1:C4,A1:A4,">0",B1:B4,"x")').mode).toBe(1);
    expect(compileFormula("SUMPRODUCT(A1:A3,B1:B3)").mode).toBe(1);
    expect(compileFormula('MATCH("pear",A1:A3,0)').mode).toBe(1);
    expect(compileFormula('XMATCH("pear",A1:A3)').mode).toBe(1);
    expect(compileFormula('XLOOKUP("pear",A1:A3,B1:B3)').mode).toBe(1);
    expect(compileFormula("INDEX(A1:B3,2,2)").mode).toBe(1);
    expect(compileFormula('VLOOKUP("pear",A1:B3,2,FALSE)').mode).toBe(1);
    expect(compileFormula('HLOOKUP("pear",A1:C2,2,FALSE)').mode).toBe(1);

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
    expect(compileFormula('COUNTIF(A:A,">0")').mode).toBe(0);
    expect(compileFormula("SUMIF(A1:A4,B1:B4,C1:C4,D1:D4)").mode).toBe(0);
    expect(compileFormula('SUMIFS(A1:A4,">0")').mode).toBe(0);
    expect(compileFormula('MATCH("pear",A1:B3,0)').mode).toBe(0);
    expect(compileFormula('XLOOKUP("pear",A1:B3,C1:D4)').mode).toBe(0);
    expect(compileFormula('VLOOKUP("pear",A1:B3,2,FALSE,1)').mode).toBe(0);
    expect(compileFormula("ROUND(A1,A2,A3)").mode).toBe(0);
    expect(compileFormula("FLOOR(A1,A2,A3)").mode).toBe(0);
    expect(compileFormula("CEILING(A1,A2,A3)").mode).toBe(0);
    expect(compileFormula("TIME(A1,A2)").mode).toBe(0);
    expect(compileFormula("WEEKDAY(A1,A2,A3)").mode).toBe(0);
    expect(compileFormula("SIN(A1,A2)").mode).toBe(0);
    expect(compileFormula('SWITCH(A1,1,"yes")').mode).toBe(1);
    expect(compileFormula("WORKDAY(A1,1,B1:B3)").mode).toBe(0);
    expect(compileFormula("NETWORKDAYS(A1,A2,B1:B3)").mode).toBe(0);
    expect(compileFormula("T.DIST(A1,2,TRUE)").mode).toBe(0);
    expect(compileFormula('TEXTJOIN(",",TRUE,A1,A2)').mode).toBe(0);
    expect(compileFormula("WORKDAY.INTL(A1,1)").mode).toBe(0);
    expect(compileFormula("LET(x,2,x+3)").mode).toBe(1);
    expect(compileFormula("LET(x,A1*2,x+3)").mode).toBe(1);
    expect(compileFormula('TEXTBEFORE(A1,"-")').mode).toBe(0);
    expect(compileFormula("FILTER(A1:A4,A1:A4>2)")).toMatchObject({ mode: 1, producesSpill: true });
    expect(compileFormula("UNIQUE(A1:A4)")).toMatchObject({ mode: 1, producesSpill: true });
    expect(compileFormula("FILTER(A1:A4,B1:B4)")).toMatchObject({ mode: 1, producesSpill: true });
    expect(compileFormula("A1:A4>2")).toMatchObject({ mode: 1, producesSpill: true });
  });

  it("accelerates rewritten logical calls while keeping higher-order lambda families on the JS path", () => {
    expect(compileFormula("TRUE()").mode).toBe(1);
    expect(compileFormula("FALSE()").mode).toBe(1);
    expect(compileFormula("IFS(A1>0,1,TRUE(),2)").mode).toBe(1);
    expect(compileFormula('SWITCH(A1,1,"yes","no")').mode).toBe(1);
    expect(compileFormula("XOR(A1>0,B1>0)").mode).toBe(1);

    expect(compileFormula("LAMBDA(x,x+1)(4)").mode).toBe(1);
    expect(compileFormula("LAMBDA(x,x+1)(A1)").mode).toBe(1);
    expect(compileFormula("LET(fn,LAMBDA(x,x+1),fn(4))").mode).toBe(1);
    expect(compileFormula("LAMBDA(x,y,IF(ISOMITTED(y),x,y))(4)").mode).toBe(0);
    expect(compileFormula("MAKEARRAY(2,2,LAMBDA(r,c,r+c))")).toMatchObject({
      mode: 1,
      producesSpill: true,
    });
    expect(compileFormula("MAP(A1:A3,LAMBDA(x,x*2))")).toMatchObject({
      mode: 1,
      producesSpill: true,
    });
    expect(compileFormula("MAP(A1:A3,B1:B3,LAMBDA(x,y,x+y))")).toMatchObject({
      mode: 1,
      producesSpill: true,
    });
    expect(compileFormula("SCAN(0,A1:A3,LAMBDA(acc,x,acc+x))")).toMatchObject({
      mode: 1,
      producesSpill: true,
    });
    expect(compileFormula("SCAN(1,A1:A3,LAMBDA(acc,x,acc*x))")).toMatchObject({
      mode: 1,
      producesSpill: true,
    });
    expect(compileFormula("BYROW(A1:B2,LAMBDA(r,SUM(r)))")).toMatchObject({
      mode: 1,
      producesSpill: true,
    });
    expect(compileFormula("BYROW(A1:B2,LAMBDA(r,AVERAGE(r)))")).toMatchObject({
      mode: 1,
      producesSpill: true,
    });
    expect(compileFormula("BYCOL(A1:B2,LAMBDA(c,SUM(c)))")).toMatchObject({
      mode: 1,
      producesSpill: true,
    });
    expect(compileFormula("BYCOL(A1:B2,LAMBDA(c,COUNTA(c)))")).toMatchObject({
      mode: 1,
      producesSpill: true,
    });
    expect(compileFormula("REDUCE(0,A1:A3,LAMBDA(acc,x,acc+x))").mode).toBe(1);
    expect(compileFormula("REDUCE(1,A1:A3,LAMBDA(acc,x,acc*x))").mode).toBe(1);
  });

  it("keeps newly added JS math families on the JS path while marking spill producers", () => {
    expect(compileFormula("TRUNC(A1,1)").mode).toBe(0);
    expect(compileFormula("ACOSH(A1)").mode).toBe(0);
    expect(compileFormula("PRODUCT(A1:A3)").mode).toBe(0);
    expect(compileFormula("SUMXMY2(A1:A3,B1:B3)").mode).toBe(0);
    expect(compileFormula("MDETERM(A1:B2)").mode).toBe(0);
    expect(compileFormula("MMULT(A1:B2,C1:D2)")).toMatchObject({ mode: 0, producesSpill: true });
    expect(compileFormula("MINVERSE(A1:B2)")).toMatchObject({ mode: 0, producesSpill: true });
    expect(compileFormula("MUNIT(3)")).toMatchObject({ mode: 0, producesSpill: true });
    expect(compileFormula("RANDARRAY(2,2)")).toMatchObject({
      mode: 0,
      producesSpill: true,
      volatile: true,
    });
    expect(compileFormula("RANDBETWEEN(1,10)").volatile).toBe(true);
  });

  it("evaluates AST against a context", () => {
    const ast = parseFormula("A1+A2");
    const context = {
      sheetName: "Sheet1",
      resolveCell: (_sheet: string, address: string): CellValue => {
        if (address === "A1") return { tag: ValueTag.Number, value: 2 };
        return { tag: ValueTag.Number, value: 3 };
      },
      resolveRange: (): CellValue[] => [],
    };
    const value = evaluateAst(ast, context);
    expect(value).toEqual({ tag: ValueTag.Number, value: 5 });
    expect(evaluatePlan(compileFormula("A1+A2").jsPlan, context)).toEqual({
      tag: ValueTag.Number,
      value: 5,
    });
    expect(evaluateAst(parseFormula("10%"), context)).toEqual({ tag: ValueTag.Number, value: 0.1 });
    expect(evaluateAst(parseFormula("(A1+A2)%"), context)).toEqual({
      tag: ValueTag.Number,
      value: 0.05,
    });
    expect(evaluatePlan(compileFormula("A1*10%").jsPlan, context)).toEqual({
      tag: ValueTag.Number,
      value: 0.2,
    });
  });

  it("parses quoted sheet references inside formulas", () => {
    const ast = parseFormula("'My Sheet'!A1+1");
    expect(ast).toMatchObject({
      kind: "BinaryExpr",
      left: { kind: "CellRef", sheetName: "My Sheet", ref: "A1" },
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
      args: [{ kind: "RangeRef", refKind: "cols", sheetName: "My Sheet", start: "A", end: "A" }],
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
        { kind: "RangeRef", refKind: "cols", start: "$D", end: "$F" },
      ],
    });
  });

  it("constant folds numeric expressions and prunes IF branches before binding", () => {
    const compiled = compileFormula("IF(TRUE, 1+2*3, A1)");
    expect(compiled.optimizedAst).toEqual({ kind: "NumberLiteral", value: 7 });
    expect(compiled.deps).toEqual([]);
    expect(compiled.jsPlan).toEqual([{ opcode: "push-number", value: 7 }, { opcode: "return" }]);
  });

  it("flattens concat calls in the optimized AST", () => {
    const compiled = compileFormula("CONCAT(A1, CONCAT(B1, C1))");
    expect(compiled.optimizedAst).toMatchObject({
      kind: "CallExpr",
      callee: "CONCAT",
      args: [
        { kind: "CellRef", ref: "A1" },
        { kind: "CellRef", ref: "B1" },
        { kind: "CellRef", ref: "C1" },
      ],
    });
  });
});
