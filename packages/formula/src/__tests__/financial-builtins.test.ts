import { describe, expect, it } from "vitest";
import {
  cumulativePeriodicPayment,
  dbDepreciation,
  ddbDepreciation,
  futureValue,
  interestPayment,
  periodicPayment,
  presentValue,
  principalPayment,
  solveRate,
  totalPeriods,
  vdbDepreciation,
} from "../builtins/financial.js";

describe("financial helpers", () => {
  it("computes time-value-of-money helpers", () => {
    expect(futureValue(0.1, 2, -100, -1000, 0)).toBeCloseTo(1420, 12);
    expect(presentValue(0.1, 2, -100, 1420, 0)).toBeCloseTo(-1000, 12);
    expect(periodicPayment(0.1, 2, 1000, 0, 0)).toBeCloseTo(-576.1904761904761, 12);
    expect(totalPeriods(0.1, -576.1904761904761, 1000, 0, 0)).toBeCloseTo(2, 12);
    expect(solveRate(48, -200, 8000, 0, 0, 0.1)).toBeCloseTo(0.007701472488246008, 12);
  });

  it("computes depreciation helpers", () => {
    expect(dbDepreciation(10000, 1000, 5, 1, 12)).toBeCloseTo(3690, 12);
    expect(ddbDepreciation(2400, 300, 10, 2, 2)).toBeCloseTo(384, 12);
    expect(vdbDepreciation(2400, 300, 10, 1, 3, 2, false)).toBeCloseTo(691.2, 12);
  });

  it("computes interest and principal helpers", () => {
    expect(interestPayment(0.1, 1, 2, 1000, 0, 0)).toBeCloseTo(-100, 12);
    expect(principalPayment(0.1, 1, 2, 1000, 0, 0)).toBeCloseTo(-476.19047619047615, 12);
    expect(cumulativePeriodicPayment(0.09 / 12, 30 * 12, 125000, 13, 24, 0, false)).toBeCloseTo(
      -11135.232130750845,
      12,
    );
    expect(cumulativePeriodicPayment(0.09 / 12, 30 * 12, 125000, 13, 24, 0, true)).toBeCloseTo(
      -934.1071234208765,
      12,
    );
  });
});
