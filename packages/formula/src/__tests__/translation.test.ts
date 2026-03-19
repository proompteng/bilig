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
});
