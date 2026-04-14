import { describe, expect, it } from "vitest";
import { buildRelativeFormulaTemplateTokenKey } from "../formula-template-key.js";
import { buildRelativeFormulaTemplateKey } from "../translation.js";

describe("buildRelativeFormulaTemplateTokenKey", () => {
  it("matches repeated row-shifted template families without parsing the full AST", () => {
    expect(buildRelativeFormulaTemplateTokenKey("A1+B1", 0, 2)).toBe(
      buildRelativeFormulaTemplateTokenKey("A2+B2", 1, 2),
    );
    expect(buildRelativeFormulaTemplateTokenKey("SUM(A1:B1)+C1#", 0, 3)).toBe(
      buildRelativeFormulaTemplateTokenKey("SUM(A2:B2)+C2#", 1, 3),
    );
  });

  it("keeps different formula families distinct", () => {
    expect(buildRelativeFormulaTemplateTokenKey("A1+B1", 0, 2)).not.toBe(
      buildRelativeFormulaTemplateTokenKey("A1*B1", 0, 2),
    );
    expect(buildRelativeFormulaTemplateTokenKey("SUM(A1:A1)", 0, 2)).not.toBe(
      buildRelativeFormulaTemplateTokenKey("SUM(A1:A2)", 1, 2),
    );
  });

  it("groups the same representative families that the AST-derived key groups", () => {
    expect(buildRelativeFormulaTemplateTokenKey("A3+B$1", 2, 2)).toBe(
      buildRelativeFormulaTemplateTokenKey("A4+B$1", 3, 2),
    );
    expect(buildRelativeFormulaTemplateKey("A3+B$1", 2, 2)).toBe(
      buildRelativeFormulaTemplateKey("A4+B$1", 3, 2),
    );
    expect(buildRelativeFormulaTemplateTokenKey("'My Sheet'!A3+SUM(B:B)", 2, 1)).toBe(
      buildRelativeFormulaTemplateTokenKey("'My Sheet'!A4+SUM(B:B)", 3, 1),
    );
    expect(buildRelativeFormulaTemplateKey("'My Sheet'!A3+SUM(B:B)", 2, 1)).toBe(
      buildRelativeFormulaTemplateKey("'My Sheet'!A4+SUM(B:B)", 3, 1),
    );
  });
});
