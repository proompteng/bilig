import { describe, expect, it } from "vitest";
import { translateFormulaReferences } from "../translation.js";

describe("translateFormulaReferences", () => {
  it("shifts relative cell references", () => {
    expect(translateFormulaReferences("A1+B2", 2, 3)).toBe("D3+E4");
  });

  it("preserves absolute anchors while shifting relative axes", () => {
    expect(translateFormulaReferences("$A1+B$2+$C$3", 4, 5)).toBe("$A5+G$2+$C$3");
  });

  it("shifts ranges, row refs, and column refs", () => {
    expect(translateFormulaReferences("SUM(A1:B2)+SUM(C:C)+SUM(3:3)", 1, 2)).toBe("SUM(C2:D3)+SUM(E:E)+SUM(4:4)");
  });

  it("shifts sheet-qualified references without dropping the sheet name", () => {
    expect(translateFormulaReferences("'My Sheet'!A1+Sheet2!B$3", 2, 1)).toBe("'My Sheet'!B3+Sheet2!C$3");
  });

  it("preserves mixed anchors across mixed cell, column, and row ranges", () => {
    expect(translateFormulaReferences("SUM($A1:B$2,$C:$D,$5:6)", 2, 3)).toBe("SUM($A3:E$2,$C:$D,$5:8)");
  });

  it("keeps quoted sheet prefixes and nested precedence intact for mixed references", () => {
    expect(
      translateFormulaReferences("('My Sheet'!$A1+Sheet2!B$2)*SUM('My Sheet'!$C:$D,Sheet2!3:$4)", 4, 2)
    ).toBe("('My Sheet'!$A5+Sheet2!D$2)*SUM('My Sheet'!$C:$D,Sheet2!7:$4)");
  });

  it("throws when a relative axis would move outside worksheet bounds even if the other axis is absolute", () => {
    expect(() => translateFormulaReferences("$A1+B$1", -1, -2)).toThrow(
      "Translated reference moved outside worksheet bounds: $A1"
    );
  });
});
