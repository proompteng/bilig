import { describe, expect, it } from "vitest";
import { compileFormula, evaluateAst, parseCellAddress, parseFormula, parseRangeAddress } from "../index.js";
import { ValueTag } from "@bilig/protocol";

describe("formula", () => {
  it("parses A1 addresses", () => {
    expect(parseCellAddress("B12")).toMatchObject({ row: 11, col: 1, text: "B12" });
  });

  it("parses quoted sheet addresses", () => {
    expect(parseCellAddress("'My Sheet'!B12")).toMatchObject({ sheetName: "My Sheet", row: 11, col: 1, text: "B12" });
  });

  it("normalizes ranges", () => {
    expect(parseRangeAddress("B2:A1")).toMatchObject({
      start: { text: "A1" },
      end: { text: "B2" }
    });
  });

  it("compiles arithmetic formulas with wasm-safe mode", () => {
    const compiled = compileFormula("A1*2");
    expect(compiled.mode).toBe(1);
    expect([...compiled.symbolicRefs]).toEqual(["A1"]);
  });

  it("keeps pass-through cell refs on the JS path", () => {
    const compiled = compileFormula("A1");
    expect(compiled.mode).toBe(0);
  });

  it("compiles bounded aggregate formulas into the wasm-safe path", () => {
    const compiled = compileFormula("SUM(A1:B2)");
    expect(compiled.mode).toBe(1);
    expect([...compiled.symbolicRefs]).toEqual(["A1", "B1", "A2", "B2"]);
  });

  it("evaluates AST against a context", () => {
    const ast = parseFormula("A1+A2");
    const value = evaluateAst(ast, {
      sheetName: "Sheet1",
      resolveCell: (_sheet, address) => {
        if (address === "A1") return { tag: ValueTag.Number, value: 2 };
        return { tag: ValueTag.Number, value: 3 };
      },
      resolveRange: () => []
    });
    expect(value).toEqual({ tag: ValueTag.Number, value: 5 });
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
    expect([...compiled.symbolicRefs]).toEqual([
      "My Sheet!A1",
      "My Sheet!B1",
      "My Sheet!A2",
      "My Sheet!B2"
    ]);
  });
});
