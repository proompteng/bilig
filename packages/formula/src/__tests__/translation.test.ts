import { describe, expect, it } from "vitest";
import {
  rewriteAddressForStructuralTransform,
  rewriteFormulaForStructuralTransform,
  rewriteRangeForStructuralTransform,
  translateFormulaReferences,
} from "../translation.js";

describe("translateFormulaReferences", () => {
  it("shifts relative cell references", () => {
    expect(translateFormulaReferences("A1+B2", 2, 3)).toBe("D3+E4");
  });

  it("preserves absolute anchors while shifting relative axes", () => {
    expect(translateFormulaReferences("$A1+B$2+$C$3", 4, 5)).toBe("$A5+G$2+$C$3");
  });

  it("shifts ranges, row refs, and column refs", () => {
    expect(translateFormulaReferences("SUM(A1:B2)+SUM(C:C)+SUM(3:3)", 1, 2)).toBe(
      "SUM(C2:D3)+SUM(E:E)+SUM(4:4)",
    );
  });

  it("shifts sheet-qualified references without dropping the sheet name", () => {
    expect(translateFormulaReferences("'My Sheet'!A1+Sheet2!B$3", 2, 1)).toBe(
      "'My Sheet'!B3+Sheet2!C$3",
    );
  });

  it("preserves mixed anchors across mixed cell, column, and row ranges", () => {
    expect(translateFormulaReferences("SUM($A1:B$2,$C:$D,$5:6)", 2, 3)).toBe(
      "SUM($A3:E$2,$C:$D,$5:8)",
    );
  });

  it("keeps quoted sheet prefixes and nested precedence intact for mixed references", () => {
    expect(
      translateFormulaReferences(
        "('My Sheet'!$A1+Sheet2!B$2)*SUM('My Sheet'!$C:$D,Sheet2!3:$4)",
        4,
        2,
      ),
    ).toBe("('My Sheet'!$A5+Sheet2!D$2)*SUM('My Sheet'!$C:$D,Sheet2!7:$4)");
  });

  it("throws when a relative axis would move outside worksheet bounds even if the other axis is absolute", () => {
    expect(() => translateFormulaReferences("$A1+B$1", -1, -2)).toThrow(
      "Translated reference moved outside worksheet bounds: $A1",
    );
  });

  it("rewrites row references for structural inserts and deletes", () => {
    expect(
      rewriteFormulaForStructuralTransform("SUM(A1:A2)", "Sheet1", "Sheet1", {
        kind: "insert",
        axis: "row",
        start: 1,
        count: 1,
      }),
    ).toBe("SUM(A1:A3)");
    expect(
      rewriteFormulaForStructuralTransform("SUM(A1:A3)", "Sheet1", "Sheet1", {
        kind: "delete",
        axis: "row",
        start: 1,
        count: 1,
      }),
    ).toBe("SUM(A1:A2)");
  });

  it("rewrites column references for structural moves", () => {
    expect(
      rewriteFormulaForStructuralTransform("A1", "Sheet1", "Sheet1", {
        kind: "move",
        axis: "column",
        start: 0,
        count: 1,
        target: 2,
      }),
    ).toBe("C1");
    expect(
      rewriteFormulaForStructuralTransform("C1", "Sheet1", "Sheet1", {
        kind: "move",
        axis: "column",
        start: 2,
        count: 1,
        target: 0,
      }),
    ).toBe("A1");
  });

  it("collapses deleted references to surviving ranges or #REF", () => {
    expect(
      rewriteFormulaForStructuralTransform("SUM(A1:B1)", "Sheet1", "Sheet1", {
        kind: "delete",
        axis: "column",
        start: 0,
        count: 1,
      }),
    ).toBe("SUM(A1:A1)");
    expect(
      rewriteFormulaForStructuralTransform("B2", "Sheet1", "Sheet1", {
        kind: "delete",
        axis: "column",
        start: 1,
        count: 1,
      }),
    ).toBe("#REF!");
  });

  it("skips structural rewrites for formulas outside the target sheet", () => {
    expect(
      rewriteFormulaForStructuralTransform("A1+Sheet2!B2", "Sheet1", "OtherSheet", {
        kind: "insert",
        axis: "row",
        start: 1,
        count: 2,
      }),
    ).toBe("A1+Sheet2!B2");
  });

  it("rewrites single-cell addresses and throws for invalid address inputs", () => {
    expect(
      rewriteAddressForStructuralTransform("B2", {
        kind: "delete",
        axis: "row",
        start: 1,
        count: 1,
      }),
    ).toBeUndefined();
    expect(
      rewriteAddressForStructuralTransform("A4", {
        kind: "move",
        axis: "row",
        start: 1,
        count: 1,
        target: 3,
      }),
    ).toBe("A3");
    expect(() =>
      rewriteAddressForStructuralTransform("bad", {
        kind: "move",
        axis: "column",
        start: 1,
        count: 1,
        target: 2,
      }),
    ).toThrow("Invalid cell reference 'bad'");
  });

  it("rewrites ranges across structural inserts, deletes, and throws on bad references", () => {
    expect(
      rewriteRangeForStructuralTransform("A1", "A4", {
        kind: "insert",
        axis: "row",
        start: 1,
        count: 1,
      }),
    ).toEqual({ startAddress: "A1", endAddress: "A5" });
    expect(
      rewriteRangeForStructuralTransform("A1", "A4", {
        kind: "delete",
        axis: "row",
        start: 1,
        count: 1,
      }),
    ).toEqual({ startAddress: "A1", endAddress: "A3" });
    expect(() =>
      rewriteRangeForStructuralTransform("A1", "bad", {
        kind: "move",
        axis: "column",
        start: 0,
        count: 1,
        target: 1,
      }),
    ).toThrow("Invalid range reference");
    expect(
      rewriteRangeForStructuralTransform("A1", "A1", {
        kind: "delete",
        axis: "row",
        start: 0,
        count: 1,
      }),
    ).toBeUndefined();
  });
});
