import { describe, expect, it } from "vitest";
import { compileFormula, evaluateAst, parseCellAddress, parseFormula, parseRangeAddress } from "../index.js";
import { ValueTag } from "@bilig/protocol";

describe("formula", () => {
  it("parses A1 addresses", () => {
    expect(parseCellAddress("B12")).toMatchObject({ row: 11, col: 1, text: "B12" });
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
});
