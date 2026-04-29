import { describe, expect, it } from 'vitest'
import { utcDateToExcelSerial } from '../builtins/datetime.js'
import {
  couponSchedule,
  couponMetrics,
  getAmordegrc,
  getAmorlinc,
  isValidBasis,
  isValidFrequency,
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
} from '../builtins/fixed-income.js'

function serial(year: number, month: number, day: number): number {
  return Math.floor(utcDateToExcelSerial(new Date(Date.UTC(year, month - 1, day))))
}

describe('fixed-income builtin helpers', () => {
  it('rejects invalid basis, frequency, date ordering, and non-economic inputs', () => {
    const jan1 = serial(2020, 1, 1)
    const feb1 = serial(2020, 2, 1)
    const jul1 = serial(2020, 7, 1)
    const jan1Next = serial(2021, 1, 1)

    expect(isValidBasis(-1)).toBe(false)
    expect(isValidBasis(0)).toBe(true)
    expect(isValidBasis(4)).toBe(true)
    expect(isValidBasis(5)).toBe(false)
    expect(isValidFrequency(0)).toBe(false)
    expect(isValidFrequency(1)).toBe(true)
    expect(isValidFrequency(2)).toBe(true)
    expect(isValidFrequency(4)).toBe(true)
    expect(isValidFrequency(12)).toBe(false)

    expect(yearFracByBasis(jan1, jul1, 1)).toBeCloseTo(182 / 366, 12)
    expect(yearFracByBasis(serial(2019, 1, 1), serial(2021, 7, 1), 1)).toBeCloseTo(912 / ((365 + 366 + 365) / 3), 12)
    expect(yearFracByBasis(jul1, jan1, 1)).toBeCloseTo(182 / 366, 12)
    expect(yearFracByBasis(serial(2020, 1, 31), serial(2020, 2, 29), 0)).toBeCloseTo(29 / 360, 12)
    expect(yearFracByBasis(serial(2021, 2, 28), serial(2021, 2, 28), 0)).toBeCloseTo(0, 12)
    expect(yearFracByBasis(serial(2020, 2, 29), serial(2020, 2, 29), 0)).toBeCloseTo(0, 12)
    expect(yearFracByBasis(serial(2020, 1, 31), serial(2020, 2, 29), 4)).toBeCloseTo(29 / 360, 12)
    expect(yearFracByBasis(serial(2020, 1, 31), serial(2020, 3, 31), 4)).toBeCloseTo(60 / 360, 12)
    expect(yearFracByBasis(jan1, jul1, 9)).toBeUndefined()
    expect(securityAnnualizedYearFraction(jul1, jan1, 0)).toBeUndefined()
    expect(securityAnnualizedYearFraction(jan1, jul1, 9)).toBeUndefined()
    expect(treasuryBillDays(jul1, jan1)).toBeUndefined()
    expect(treasuryBillDays(jan1, serial(2021, 2, 1))).toBeUndefined()
    expect(maturityAtIssueFractions(jan1, jul1, feb1, 0)).toBeUndefined()
    expect(maturityAtIssueFractions(feb1, jul1, jan1, 9)).toBeUndefined()

    expect(oddLastCouponFractions(jan1, jul1, jan1, 2, 0)).toBeUndefined()
    expect(oddLastCouponFractions(feb1, jul1, jan1, 3, 0)).toBeUndefined()
    expect(oddLastCouponFractions(feb1, jul1, jan1, 2, 9)).toBeUndefined()
    expect(oddFirstPriceValue(feb1, jan1Next, jan1, jul1, -0.01, 0.02, 100, 2, 0)).toBeUndefined()
    expect(oddFirstPriceValue(feb1, jan1Next, jan1, jul1, 0.01, -3, 100, 2, 0)).toBeUndefined()
    expect(oddFirstPriceValue(jul1, jan1Next, jan1, feb1, 0.01, 0.02, 100, 2, 0)).toBeUndefined()
    expect(solveOddFirstYield(feb1, jan1Next, jan1, jul1, -0.01, 100, 100, 2, 0)).toBeUndefined()
    expect(solveOddFirstYield(feb1, jan1Next, jan1, jul1, 0.01, 0, 100, 2, 0)).toBeUndefined()

    expect(couponSchedule(jul1, jan1, 2)).toBeUndefined()
    expect(couponSchedule(jan1, jul1, 3)).toBeUndefined()
    expect(couponMetrics(jan1, jul1, 2, 9)).toBeUndefined()
    expect(
      pricePeriodicSecurity(
        { previousCoupon: jan1, nextCoupon: jul1, periodsRemaining: 1, accruedDays: 0, daysToNextCoupon: 1, daysInPeriod: 180 },
        0.05,
        -400,
        100,
        2,
      ),
    ).toBeUndefined()
    expect(
      solvePeriodicSecurityYield(
        { previousCoupon: jan1, nextCoupon: jul1, periodsRemaining: 1, accruedDays: 200, daysToNextCoupon: 0, daysInPeriod: 180 },
        0.05,
        100,
        100,
        2,
      ),
    ).toBeUndefined()
    expect(
      pricePeriodicSecurity(
        { previousCoupon: jan1, nextCoupon: jul1, periodsRemaining: 2, accruedDays: 0, daysToNextCoupon: 90, daysInPeriod: 180 },
        0.05,
        -3,
        100,
        2,
      ),
    ).toBeUndefined()
    expect(
      macaulayDuration(
        { previousCoupon: jan1, nextCoupon: jul1, periodsRemaining: 2, accruedDays: 0, daysToNextCoupon: 90, daysInPeriod: 180 },
        0.05,
        -3,
        2,
      ),
    ).toBeUndefined()
    expect(oddFirstPriceValue(feb1, jan1Next, jan1, jul1, 0.01, -2, 100, 2, 0)).toBeUndefined()
    expect(
      solvePeriodicSecurityYield(
        { previousCoupon: jan1, nextCoupon: jul1, periodsRemaining: 2, accruedDays: 0, daysToNextCoupon: 90, daysInPeriod: 180 },
        0.05,
        10_000,
        100,
        2,
      ),
    ).toBeLessThan(-1)

    expect(getAmorlinc(0, jan1, jul1, 0, 0, 0.1, 0)).toBeUndefined()
    expect(getAmorlinc(100, jul1, jan1, 0, 0, 0.1, 0)).toBeUndefined()
    expect(getAmorlinc(100, jan1, jul1, 200, 0, 0.1, 0)).toBeUndefined()
    expect(getAmorlinc(100, jan1, jul1, 0, -1, 0.1, 0)).toBeUndefined()
    expect(getAmorlinc(100, jan1, jul1, 0, Number.POSITIVE_INFINITY, 0.1, 0)).toBeUndefined()
    expect(getAmordegrc(100, jan1, jul1, 0, 0, 0, 0)).toBeUndefined()
    expect(getAmordegrc(100, jan1, jul1, 0, -1, 0.1, 9)).toBeUndefined()
  })

  it('computes year fractions, treasury bill days, and odd last coupon fractions', () => {
    expect(yearFracByBasis(serial(2020, 1, 1), serial(2020, 7, 1), 0)).toBeCloseTo(0.5, 12)
    expect(securityAnnualizedYearFraction(serial(2023, 1, 1), serial(2023, 4, 1), 2)).toBeCloseTo(0.25, 12)
    expect(treasuryBillDays(serial(2008, 3, 31), serial(2008, 6, 1))).toBe(62)
    expect(maturityAtIssueFractions(serial(2020, 4, 1), serial(2020, 7, 1), serial(2020, 1, 1), 0)).toEqual({
      issueToMaturity: 0.5,
      settlementToMaturity: 0.25,
      issueToSettlement: 0.25,
    })

    const oddLastFractions = oddLastCouponFractions(serial(2020, 4, 1), serial(2020, 5, 1), serial(2020, 1, 1), 2, 0)
    expect(oddLastFractions).toBeDefined()
    expect(oddLastFractions?.accruedFraction).toBeCloseTo(0.5, 12)
    expect(oddLastFractions?.remainingFraction).toBeCloseTo(1 / 6, 12)
    expect(oddLastFractions?.totalFraction).toBeCloseTo(2 / 3, 12)
  })

  it('computes coupon metrics with periodic bond price and yield', () => {
    const helperMetrics = couponMetrics(serial(2007, 1, 25), serial(2009, 11, 15), 2, 4)
    expect(helperMetrics).toBeDefined()
    expect(helperMetrics?.previousCoupon).toBe(39036)
    expect(helperMetrics?.nextCoupon).toBe(39217)
    expect(helperMetrics?.periodsRemaining).toBe(6)
    expect(helperMetrics?.accruedDays).toBe(70)
    expect(helperMetrics?.daysToNextCoupon).toBe(110)
    expect(helperMetrics?.daysInPeriod).toBe(180)

    const priceMetrics = couponMetrics(serial(2008, 2, 15), serial(2017, 11, 15), 2, 0)
    expect(priceMetrics).toBeDefined()
    expect(pricePeriodicSecurity(priceMetrics!, 0.0575, 0.065, 100, 2)).toBeCloseTo(94.63436162132213, 12)
    expect(
      pricePeriodicSecurity(
        {
          previousCoupon: serial(2020, 1, 1),
          nextCoupon: serial(2020, 7, 1),
          periodsRemaining: 1,
          accruedDays: 30,
          daysToNextCoupon: 90,
          daysInPeriod: 180,
        },
        0.05,
        0.04,
        100,
        2,
      ),
    ).toBeCloseTo(101.06848184818482, 12)
    expect(
      solvePeriodicSecurityYield(
        {
          previousCoupon: serial(2020, 1, 1),
          nextCoupon: serial(2020, 7, 1),
          periodsRemaining: 1,
          accruedDays: 30,
          daysToNextCoupon: 90,
          daysInPeriod: 180,
        },
        0.05,
        101.06848184818482,
        100,
        2,
      ),
    ).toBeCloseTo(0.04, 12)

    const yieldMetrics = couponMetrics(serial(2008, 2, 15), serial(2016, 11, 15), 2, 0)
    expect(yieldMetrics).toBeDefined()
    expect(solvePeriodicSecurityYield(yieldMetrics!, 0.0575, 95.04287, 100, 2)).toBeCloseTo(0.065, 7)
  })

  it('computes odd first coupon price and yield', () => {
    expect(
      oddFirstPriceValue(serial(2008, 11, 11), serial(2021, 3, 1), serial(2008, 10, 15), serial(2009, 3, 1), 0.0785, 0.0625, 100, 2, 1),
    ).toBeCloseTo(113.597717474079, 12)

    expect(
      solveOddFirstYield(serial(2008, 11, 11), serial(2021, 3, 1), serial(2008, 10, 15), serial(2009, 3, 1), 0.0575, 84.5, 100, 2, 0),
    ).toBeCloseTo(0.0772455415972989, 11)
  })

  it('computes duration and amortization helpers', () => {
    const durationMetrics = couponMetrics(serial(2018, 7, 1), serial(2048, 1, 1), 2, 1)
    expect(durationMetrics).toBeDefined()
    expect(macaulayDuration(durationMetrics!, 0.08, 0.09, 2)).toBeCloseTo(10.919145281591925, 12)

    expect(getAmorlinc(1000, serial(2020, 1, 1), serial(2021, 1, 1), 25, 0, 0.15, 0)).toBe(150)
    expect(getAmorlinc(1000, serial(2020, 1, 1), serial(2021, 1, 1), 25, 6, 0.15, 0)).toBe(75)
    expect(getAmorlinc(1000, serial(2020, 1, 1), serial(2021, 1, 1), 25, 7, 0.15, 0)).toBe(0)
    expect(getAmordegrc(1000, serial(2020, 1, 1), serial(2021, 1, 1), 10, 1, 0.4, 0)).toBe(240)
    expect(getAmordegrc(1000, serial(2020, 1, 1), serial(2021, 1, 1), 10, 1, 0.2, 0)).toBe(240)
    expect(getAmordegrc(1000, serial(2020, 1, 1), serial(2021, 1, 1), 10, 1, 0.3, 0)).toBe(247)
    expect(getAmordegrc(1000, serial(2020, 1, 1), serial(2021, 1, 1), 10, 1, 0.1, 0)).toBe(188)
    expect(getAmordegrc(100, serial(2020, 1, 1), serial(2021, 1, 1), 99, 2, 0.9, 0)).toBe(0)
  })

  it('covers fixed-income boundary branches for schedules, solvers, and depreciation', () => {
    const jan1 = serial(2020, 1, 1)
    const feb1 = serial(2020, 2, 1)
    const jul1 = serial(2020, 7, 1)
    const jan1Next = serial(2021, 1, 1)

    expect(yearFracByBasis(Number.NaN, jul1, 0)).toBeUndefined()
    expect(yearFracByBasis(serial(2019, 3, 1), serial(2020, 3, 1), 1)).toBeCloseTo(366 / 366, 12)
    expect(yearFracByBasis(serial(2021, 3, 1), serial(2022, 3, 1), 1)).toBeCloseTo(365 / 365, 12)
    expect(yearFracByBasis(serial(2020, 2, 29), serial(2020, 3, 31), 0)).toBeCloseTo(31 / 360, 12)
    expect(yearFracByBasis(serial(2020, 1, 31), serial(2020, 3, 31), 0)).toBeCloseTo(60 / 360, 12)

    expect(oddLastCouponFractions(serial(2022, 1, 1), serial(2040, 1, 1), serial(2020, 1, 1), 2, 0)).toBeUndefined()
    expect(oddFirstPriceValue(feb1, jan1Next, jan1, jul1, 0.01, -2.1, 100, 2, 0)).toBeUndefined()
    expect(solveOddFirstYield(feb1, jan1Next, jan1, jul1, 0.01, 10_000, 100, 2, 0)).toBeLessThan(-1)

    const onePeriod = {
      previousCoupon: jan1,
      nextCoupon: jul1,
      periodsRemaining: 1,
      accruedDays: 30,
      daysToNextCoupon: 90,
      daysInPeriod: 180,
    }
    const twoPeriod = {
      previousCoupon: jan1,
      nextCoupon: jan1Next,
      periodsRemaining: 2,
      accruedDays: 30,
      daysToNextCoupon: 90,
      daysInPeriod: 180,
    }

    expect(pricePeriodicSecurity(twoPeriod, 0.05, -3, 100, 2)).toBeUndefined()
    expect(solvePeriodicSecurityYield(onePeriod, 0.05, -10, 100, 2)).toBeUndefined()
    expect(solvePeriodicSecurityYield(twoPeriod, 0.05, 10_000, 100, 2)).toBeLessThan(-1)
    expect(macaulayDuration(twoPeriod, 0.05, -3, 2)).toBeUndefined()

    expect(couponSchedule(serial(2020, 1, 30), serial(2020, 2, 29), 1)).toEqual({
      previousCoupon: serial(2019, 2, 28),
      nextCoupon: serial(2020, 2, 29),
      periodsRemaining: 1,
    })
    expect(couponMetrics(serial(2020, 1, 30), serial(2020, 2, 29), 1, 3)?.daysInPeriod).toBeCloseTo(365, 12)

    expect(getAmorlinc(100, jan1, jan1Next, -1, 0, 0.1, 0)).toBeUndefined()
    expect(getAmorlinc(100, jan1, jan1Next, 0, Number.NaN, 0.1, 0)).toBeUndefined()
    expect(getAmorlinc(100, jan1, jan1Next, 0, 1, 0.1, 9)).toBeUndefined()
    expect(getAmorlinc(100, jan1, jan1Next, 99, 2, 0.5, 0)).toBe(0)
    expect(getAmordegrc(100, jan1, jan1Next, -1, 0, 0.1, 0)).toBeUndefined()
    expect(getAmordegrc(100, jan1, jan1Next, 0, Number.NaN, 0.1, 0)).toBeUndefined()
    expect(getAmordegrc(100, jan1, jan1Next, 0, 1, 0.16, 0)).toBe(24)
  })
})
