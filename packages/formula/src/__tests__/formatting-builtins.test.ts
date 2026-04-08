import { describe, expect, it } from "vitest";
import {
  countLeadingZeros,
  formatFixed,
  isValidDollarFraction,
  parseDollarDecimal,
  toColumnLabel,
} from "../builtins/formatting.js";

describe("formatting builtin helpers", () => {
  it("formats column labels for A1 references", () => {
    expect(toColumnLabel(1)).toBe("A");
    expect(toColumnLabel(28)).toBe("AB");
    expect(toColumnLabel(703)).toBe("AAA");
    expect(toColumnLabel(0)).toBeUndefined();
  });

  it("formats fixed strings with rounding and grouping", () => {
    expect(formatFixed(-1234.567, 2, true)).toBe("-1,234.57");
    expect(formatFixed(1234.567, -1, false)).toBe("1230");
    expect(formatFixed(Number.NaN, 2, true)).toBe("");
  });

  it("parses dollar fractions with Excel-compatible constraints", () => {
    expect(isValidDollarFraction(1)).toBe(true);
    expect(isValidDollarFraction(16)).toBe(true);
    expect(isValidDollarFraction(12)).toBe(false);
    expect(parseDollarDecimal(-5.08)).toEqual({ integerPart: 5, fractionalNumerator: 8 });
    expect(countLeadingZeros(32)).toBe(2);
    expect(countLeadingZeros(0)).toBe(1);
  });
});
