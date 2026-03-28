import { describe, expect, it } from "vitest";
import { ErrorCode } from "@bilig/protocol";
import {
  rewriteAddressForStructuralTransform,
  rewriteFormulaForStructuralTransform,
  rewriteRangeForStructuralTransform,
  serializeFormula,
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

  it("translates spill refs, unary expressions, and invoke expressions through the public API", () => {
    expect(translateFormulaReferences("-A1+A1#", 2, 1)).toBe("-B3+B3#");
    expect(translateFormulaReferences("LAMBDA(x,x+1)(A1)", 1, 2)).toBe("LAMBDA(x,x+1)(C2)");
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

  it("serializes literals, structured refs, spill refs, invokes, and precedence-sensitive binaries", () => {
    expect(serializeFormula({ kind: "BooleanLiteral", value: false })).toBe("FALSE");
    expect(serializeFormula({ kind: "StringLiteral", value: 'a"b' })).toBe('"a""b"');
    expect(serializeFormula({ kind: "ErrorLiteral", code: ErrorCode.Spill })).toBe("#SPILL!");
    expect(
      serializeFormula({
        kind: "StructuredRef",
        tableName: "Sales",
        columnName: "Amount",
      }),
    ).toBe("Sales[Amount]");
    expect(
      serializeFormula({
        kind: "SpillRef",
        sheetName: "My Sheet",
        ref: "A1",
      }),
    ).toBe("'My Sheet'!A1#");
    expect(
      serializeFormula({
        kind: "InvokeExpr",
        callee: { kind: "NameRef", name: "fn" },
        args: [{ kind: "NumberLiteral", value: 1 }],
      }),
    ).toBe("(fn)(1)");
    expect(
      serializeFormula({
        kind: "BinaryExpr",
        operator: "^",
        left: {
          kind: "BinaryExpr",
          operator: "^",
          left: { kind: "NumberLiteral", value: 2 },
          right: { kind: "NumberLiteral", value: 3 },
        },
        right: { kind: "NumberLiteral", value: 4 },
      }),
    ).toBe("(2^3)^4");
    expect(
      serializeFormula({
        kind: "BinaryExpr",
        operator: "-",
        left: { kind: "NumberLiteral", value: 1 },
        right: {
          kind: "BinaryExpr",
          operator: "-",
          left: { kind: "NumberLiteral", value: 2 },
          right: { kind: "NumberLiteral", value: 3 },
        },
      }),
    ).toBe("1-(2-3)");
  });

  it("rewrites spill refs and axis ranges while collapsing deleted axis refs to #REF", () => {
    expect(
      rewriteFormulaForStructuralTransform("'My Sheet'!A1#+SUM(2:2)", "My Sheet", "My Sheet", {
        kind: "insert",
        axis: "row",
        start: 0,
        count: 1,
      }),
    ).toBe("'My Sheet'!A2#+SUM(3:3)");
    expect(
      rewriteFormulaForStructuralTransform("SUM(B:B)", "Sheet1", "Sheet1", {
        kind: "delete",
        axis: "column",
        start: 1,
        count: 1,
      }),
    ).toBe("SUM(#REF!)");
    expect(
      rewriteFormulaForStructuralTransform("SUM(2:2)", "Sheet1", "Sheet1", {
        kind: "delete",
        axis: "row",
        start: 1,
        count: 1,
      }),
    ).toBe("SUM(#REF!)");
  });

  it("rewrites invoke expressions while leaving unaffected axis refs on other sheets intact", () => {
    expect(
      rewriteFormulaForStructuralTransform("LAMBDA(x,x+A1)(B2)", "Sheet1", "Sheet1", {
        kind: "insert",
        axis: "row",
        start: 0,
        count: 1,
      }),
    ).toBe("LAMBDA(x,x+A2)(B3)");
    expect(
      rewriteFormulaForStructuralTransform(
        "SUM(Sheet2!C:C)+SUM(Sheet2!$4:$4)",
        "Sheet1",
        "Sheet1",
        {
          kind: "move",
          axis: "column",
          start: 0,
          count: 1,
          target: 2,
        },
      ),
    ).toBe("SUM(Sheet2!C:C)+SUM(Sheet2!$4:$4)");
  });

  it("throws when translated row or column refs move outside worksheet bounds", () => {
    expect(() => translateFormulaReferences("SUM(A:A)", 0, -1)).toThrow(
      "Translated reference moved outside worksheet bounds: A",
    );
    expect(() => translateFormulaReferences("SUM(1:1)", -1, 0)).toThrow(
      "Translated reference moved outside worksheet bounds: 1",
    );
  });

  it("serializes sheet-qualified refs, unary expressions, and call expressions", () => {
    expect(serializeFormula({ kind: "CellRef", sheetName: "Sheet 2", ref: "$B3" })).toBe(
      "'Sheet 2'!$B3",
    );
    expect(serializeFormula({ kind: "ColumnRef", sheetName: "Sheet2", ref: "C" })).toBe("Sheet2!C");
    expect(serializeFormula({ kind: "RowRef", sheetName: "Sheet2", ref: "$4" })).toBe("Sheet2!$4");
    expect(
      serializeFormula({
        kind: "RangeRef",
        sheetName: "Sheet2",
        refKind: "cells",
        start: "A1",
        end: "B2",
      }),
    ).toBe("Sheet2!A1:B2");
    expect(
      serializeFormula({
        kind: "UnaryExpr",
        operator: "-",
        argument: {
          kind: "BinaryExpr",
          operator: "+",
          left: { kind: "NumberLiteral", value: 1 },
          right: { kind: "NumberLiteral", value: 2 },
        },
      }),
    ).toBe("-(1+2)");
    expect(
      serializeFormula({
        kind: "CallExpr",
        callee: "SUM",
        args: [
          { kind: "CellRef", sheetName: null, ref: "A1" },
          { kind: "RangeRef", sheetName: null, refKind: "cols", start: "B", end: "C" },
        ],
      }),
    ).toBe("SUM(A1,B:C)");
  });

  it("rewrites move transforms across unaffected, shifted, and invalid intervals", () => {
    expect(
      rewriteFormulaForStructuralTransform("SUM(A:C)", "Sheet1", "Sheet1", {
        kind: "insert",
        axis: "column",
        start: 5,
        count: 2,
      }),
    ).toBe("SUM(A:C)");
    expect(
      rewriteFormulaForStructuralTransform("SUM(C:E)", "Sheet1", "Sheet1", {
        kind: "move",
        axis: "column",
        start: 1,
        count: 2,
        target: 4,
      }),
    ).toBe("SUM(B:F)");
    expect(
      rewriteFormulaForStructuralTransform("SUM(5:7)", "Sheet1", "Sheet1", {
        kind: "move",
        axis: "row",
        start: 4,
        count: 2,
        target: 1,
      }),
    ).toBe("SUM(2:7)");
    expect(
      rewriteFormulaForStructuralTransform("SUM(5:6)", "Sheet1", "Sheet1", {
        kind: "delete",
        axis: "row",
        start: 4,
        count: 2,
      }),
    ).toBe("SUM(#REF!)");
  });

  it("rewrites point references and intervals across move and delete edge segments", () => {
    expect(
      rewriteFormulaForStructuralTransform("3:3", "Sheet1", "Sheet1", {
        kind: "move",
        axis: "row",
        start: 4,
        count: 2,
        target: 1,
      }),
    ).toBe("5:5");
    expect(
      rewriteFormulaForStructuralTransform("10:10", "Sheet1", "Sheet1", {
        kind: "move",
        axis: "row",
        start: 4,
        count: 2,
        target: 1,
      }),
    ).toBe("10:10");
    expect(
      rewriteFormulaForStructuralTransform("SUM(5:7)", "Sheet1", "Sheet1", {
        kind: "delete",
        axis: "row",
        start: 1,
        count: 1,
      }),
    ).toBe("SUM(4:6)");
    expect(
      rewriteFormulaForStructuralTransform("SUM(5:7)", "Sheet1", "Sheet1", {
        kind: "delete",
        axis: "row",
        start: 10,
        count: 1,
      }),
    ).toBe("SUM(5:7)");
  });

  it("covers remaining structural transformation edge cases for intervals and sheet quoting", () => {
    expect(
      rewriteFormulaForStructuralTransform("SUM(A1:C3)", "Sheet1", "Sheet1", {
        kind: "move",
        axis: "row",
        start: 0,
        count: 1,
        target: 5,
      }),
    ).toBe("SUM(A1:C6)");

    expect(
      rewriteFormulaForStructuralTransform(
        "'Sheet With Spaces'!A1",
        "Sheet1",
        "Sheet With Spaces",
        {
          kind: "insert",
          axis: "row",
          start: 0,
          count: 1,
        },
      ),
    ).toBe("'Sheet With Spaces'!A2");

    expect(
      rewriteFormulaForStructuralTransform("'It''s a sheet'!A1", "Sheet1", "It's a sheet", {
        kind: "insert",
        axis: "row",
        start: 0,
        count: 1,
      }),
    ).toBe("'It''s a sheet'!A2");
  });

  it("serializes fallback errors, nested invoke callees, and quoted sheet names", () => {
    const unknownErrorNode: { kind: "ErrorLiteral"; code: ErrorCode } = {
      kind: "ErrorLiteral",
      code: ErrorCode.Value,
    };
    Reflect.set(unknownErrorNode, "code", 999);
    expect(serializeFormula(unknownErrorNode)).toBe("#ERROR!");
    expect(
      serializeFormula({
        kind: "InvokeExpr",
        callee: {
          kind: "CallExpr",
          callee: "LAMBDA",
          args: [
            { kind: "NameRef", name: "x" },
            { kind: "NameRef", name: "x" },
          ],
        },
        args: [{ kind: "NumberLiteral", value: 4 }],
      }),
    ).toBe("LAMBDA(x,x)(4)");
    expect(
      serializeFormula({
        kind: "CellRef",
        sheetName: "Sales.Q1",
        ref: "$C$5",
      }),
    ).toBe("Sales.Q1!$C$5");
  });
});
