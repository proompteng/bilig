import { describe, expect, it } from "vitest";
import { utcDateToExcelSerial } from "../builtins/datetime.js";
import {
  couponMetrics,
  getAmordegrc,
  getAmorlinc,
  macaulayDuration,
  maturityAtIssueFractions,
  oddFirstPriceValue,
  oddLastCouponFractions,
  pricePeriodicSecurity,
  securityAnnualizedYearFraction,
  solveOddFirstYield,
  solvePeriodicSecurityYield,
  treasuryBillDays,
  yearFracByBasis,
} from "../builtins/fixed-income.js";

function serial(year: number, month: number, day: number): number {
  return Math.floor(utcDateToExcelSerial(new Date(Date.UTC(year, month - 1, day))));
}

describe("fixed-income builtin helpers", () => {
  it("computes year fractions, treasury bill days, and odd last coupon fractions", () => {
    expect(yearFracByBasis(serial(2020, 1, 1), serial(2020, 7, 1), 0)).toBeCloseTo(0.5, 12);
    expect(securityAnnualizedYearFraction(serial(2023, 1, 1), serial(2023, 4, 1), 2)).toBeCloseTo(
      0.25,
      12,
    );
    expect(treasuryBillDays(serial(2008, 3, 31), serial(2008, 6, 1))).toBe(62);
    expect(
      maturityAtIssueFractions(serial(2020, 4, 1), serial(2020, 7, 1), serial(2020, 1, 1), 0),
    ).toEqual({
      issueToMaturity: 0.5,
      settlementToMaturity: 0.25,
      issueToSettlement: 0.25,
    });

    const oddLastFractions = oddLastCouponFractions(
      serial(2020, 4, 1),
      serial(2020, 5, 1),
      serial(2020, 1, 1),
      2,
      0,
    );
    expect(oddLastFractions).toBeDefined();
    expect(oddLastFractions?.accruedFraction).toBeCloseTo(0.5, 12);
    expect(oddLastFractions?.remainingFraction).toBeCloseTo(1 / 6, 12);
    expect(oddLastFractions?.totalFraction).toBeCloseTo(2 / 3, 12);
  });

  it("computes coupon metrics with periodic bond price and yield", () => {
    const helperMetrics = couponMetrics(serial(2007, 1, 25), serial(2009, 11, 15), 2, 4);
    expect(helperMetrics).toBeDefined();
    expect(helperMetrics?.previousCoupon).toBe(39036);
    expect(helperMetrics?.nextCoupon).toBe(39217);
    expect(helperMetrics?.periodsRemaining).toBe(6);
    expect(helperMetrics?.accruedDays).toBe(70);
    expect(helperMetrics?.daysToNextCoupon).toBe(110);
    expect(helperMetrics?.daysInPeriod).toBe(180);

    const priceMetrics = couponMetrics(serial(2008, 2, 15), serial(2017, 11, 15), 2, 0);
    expect(priceMetrics).toBeDefined();
    expect(pricePeriodicSecurity(priceMetrics!, 0.0575, 0.065, 100, 2)).toBeCloseTo(
      94.63436162132213,
      12,
    );

    const yieldMetrics = couponMetrics(serial(2008, 2, 15), serial(2016, 11, 15), 2, 0);
    expect(yieldMetrics).toBeDefined();
    expect(solvePeriodicSecurityYield(yieldMetrics!, 0.0575, 95.04287, 100, 2)).toBeCloseTo(
      0.065,
      7,
    );
  });

  it("computes odd first coupon price and yield", () => {
    expect(
      oddFirstPriceValue(
        serial(2008, 11, 11),
        serial(2021, 3, 1),
        serial(2008, 10, 15),
        serial(2009, 3, 1),
        0.0785,
        0.0625,
        100,
        2,
        1,
      ),
    ).toBeCloseTo(113.597717474079, 12);

    expect(
      solveOddFirstYield(
        serial(2008, 11, 11),
        serial(2021, 3, 1),
        serial(2008, 10, 15),
        serial(2009, 3, 1),
        0.0575,
        84.5,
        100,
        2,
        0,
      ),
    ).toBeCloseTo(0.0772455415972989, 11);
  });

  it("computes duration and amortization helpers", () => {
    const durationMetrics = couponMetrics(serial(2018, 7, 1), serial(2048, 1, 1), 2, 1);
    expect(durationMetrics).toBeDefined();
    expect(macaulayDuration(durationMetrics!, 0.08, 0.09, 2)).toBeCloseTo(10.919145281591925, 12);

    expect(getAmorlinc(1000, serial(2020, 1, 1), serial(2021, 1, 1), 25, 0, 0.15, 0)).toBe(150);
    expect(getAmorlinc(1000, serial(2020, 1, 1), serial(2021, 1, 1), 25, 6, 0.15, 0)).toBe(75);
    expect(getAmordegrc(1000, serial(2020, 1, 1), serial(2021, 1, 1), 10, 1, 0.2, 0)).toBe(240);
    expect(getAmordegrc(1000, serial(2020, 1, 1), serial(2021, 1, 1), 10, 1, 0.3, 0)).toBe(247);
  });
});
