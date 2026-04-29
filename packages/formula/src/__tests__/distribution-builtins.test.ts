import { describe, expect, it } from 'vitest'
import {
  betaDistributionCdf,
  betaDistributionDensity,
  betaDistributionInverse,
  besselIValue,
  besselJValue,
  besselKValue,
  besselYValue,
  binomialProbability,
  chiSquareDensity,
  chiSquareCdf,
  fDistributionDensity,
  fDistributionCdf,
  gammaDistributionDensity,
  gammaDistributionCdf,
  gammaFunction,
  hypergeometricProbability,
  inverseChiSquare,
  inverseFDistribution,
  inverseGammaDistribution,
  inverseNormal,
  inverseStandardNormal,
  inverseStudentT,
  kurtosis,
  logGamma,
  negativeBinomialProbability,
  percentileNormal,
  poissonProbability,
  skewPopulation,
  skewSample,
  standardNormalCdf,
  standardNormalPdf,
  studentTDensity,
  studentTCdf,
} from '../builtins/distributions.js'

describe('distribution helpers', () => {
  it('evaluates Bessel and gamma helpers', () => {
    expect(besselIValue(1.5, 1)).toBeCloseTo(0.981666428, 7)
    expect(besselKValue(1.5, 1)).toBeCloseTo(0.277387804, 7)
    expect(gammaFunction(5)).toBeCloseTo(24, 10)
  })

  it('computes normal-family helpers and inverses', () => {
    expect(standardNormalCdf(1)).toBeCloseTo(0.8413447460685429, 7)
    expect(inverseStandardNormal(0.001)).toBeCloseTo(-3.090232306167813, 8)
  })

  it('computes continuous distribution helpers', () => {
    const betaCdf = betaDistributionCdf(2, 8, 10, 1, 3)
    expect(betaCdf).toBeCloseTo(0.6854705810117458, 10)
    expect(betaDistributionInverse(betaCdf, 8, 10, 1, 3)).toBeCloseTo(2, 10)

    expect(fDistributionCdf(15.2068649, 6, 4)).toBeCloseTo(0.99, 9)
    expect(inverseFDistribution(0.01, 6, 4)).toBeCloseTo(0.10930991466299911, 8)

    expect(studentTCdf(1, 1)).toBeCloseTo(0.75, 12)
    expect(inverseStudentT(0.75, 1)).toBeCloseTo(1, 9)

    expect(inverseGammaDistribution(0.08030139707139418, 3, 2)).toBeCloseTo(2, 10)
    expect(chiSquareCdf(3, 4)).toBeCloseTo(0.4421745996289252, 12)
    expect(inverseChiSquare(0.93, 1)).toBeCloseTo(3.283020286759536, 8)
  })

  it('computes discrete probability helpers', () => {
    expect(poissonProbability(3, 2.5)).toBeCloseTo(0.21376301724973648, 12)
    expect(binomialProbability(2, 4, 0.5)).toBeCloseTo(0.375, 12)
    expect(hypergeometricProbability(1, 4, 3, 10)).toBeCloseTo(0.5, 12)
    expect(negativeBinomialProbability(2, 3, 0.5)).toBeCloseTo(0.1875, 12)
  })

  it('covers distribution helper boundaries and invalid inputs', () => {
    expect(Number.isNaN(logGamma(0))).toBe(true)
    expect(Number.isNaN(logGamma(Number.POSITIVE_INFINITY))).toBe(true)
    expect(Number.isNaN(gammaFunction(-2))).toBe(true)
    expect(Number.isNaN(gammaFunction(Number.NaN))).toBe(true)
    expect(gammaFunction(-0.5)).toBeCloseTo(-3.5449077018, 8)
    expect(gammaFunction(0.25)).toBeCloseTo(3.625609908, 8)

    expect(besselJValue(0, 0)).toBe(1)
    expect(besselJValue(0, 2)).toBe(0)
    expect(besselJValue(-1.5, 1)).toBeCloseTo(-0.5579365079, 8)
    expect(besselIValue(0, 0)).toBe(1)
    expect(besselIValue(0, 2)).toBe(0)
    expect(besselIValue(-1.5, 1)).toBeLessThan(0)
    expect(Number.isFinite(besselYValue(1.5, 1))).toBe(true)
    expect(Number.isFinite(besselKValue(1.5, 1))).toBe(true)

    expect(Number.isNaN(betaDistributionDensity(0.5, 0, 1))).toBe(true)
    expect(Number.isNaN(betaDistributionDensity(2, 2, 3))).toBe(true)
    expect(Number.isNaN(betaDistributionDensity(0.5, 2, 3, 1, 1))).toBe(true)
    expect(betaDistributionDensity(0.5, 2, 3)).toBeCloseTo(1.5, 12)
    expect(betaDistributionDensity(0, 1, 2)).toBeCloseTo(2, 12)
    expect(betaDistributionDensity(0, 0.5, 2)).toBe(Number.POSITIVE_INFINITY)
    expect(betaDistributionDensity(0, 2, 2)).toBe(0)
    expect(betaDistributionDensity(1, 2, 1)).toBeCloseTo(2, 12)
    expect(betaDistributionDensity(1, 2, 0.5)).toBe(Number.POSITIVE_INFINITY)
    expect(betaDistributionDensity(1, 2, 2)).toBe(0)
    expect(betaDistributionCdf(0, 2, 3)).toBe(0)
    expect(betaDistributionCdf(1, 2, 3)).toBe(1)
    expect(Number.isNaN(betaDistributionCdf(-0.1, 2, 3))).toBe(true)
    expect(betaDistributionInverse(-0.1, 2, 3)).toBeUndefined()
    expect(betaDistributionInverse(0.5, 0, 3)).toBeUndefined()

    expect(Number.isNaN(fDistributionDensity(-1, 2, 3))).toBe(true)
    expect(fDistributionDensity(0, 1, 3)).toBe(Number.POSITIVE_INFINITY)
    expect(fDistributionDensity(0, 2, 3)).toBe(1)
    expect(fDistributionDensity(0, 3, 3)).toBe(0)
    expect(Number.isNaN(fDistributionCdf(-1, 2, 3))).toBe(true)
    expect(inverseFDistribution(0, 2, 3)).toBeUndefined()
    expect(inverseFDistribution(1, 2, 3)).toBeUndefined()
    expect(inverseFDistribution(0.5, 0, 3)).toBeUndefined()

    expect(Number.isNaN(studentTDensity(0, 0))).toBe(true)
    expect(studentTDensity(0, 10)).toBeCloseTo(0.389108, 5)
    expect(studentTCdf(0, 10)).toBe(0.5)
    expect(Number.isNaN(studentTCdf(0, 0))).toBe(true)
    expect(studentTCdf(-1, 1)).toBeCloseTo(0.25, 12)
    expect(inverseStudentT(0, 1)).toBeUndefined()
    expect(inverseStudentT(1, 1)).toBeUndefined()
    expect(inverseStudentT(0.5, 10)).toBe(0)
    expect(inverseStudentT(0.25, 1)).toBeCloseTo(-1, 9)

    expect(Number.isNaN(gammaDistributionDensity(-1, 2, 3))).toBe(true)
    expect(Number.isNaN(gammaDistributionDensity(1, 0, 3))).toBe(true)
    expect(gammaDistributionDensity(0, 1, 3)).toBeCloseTo(1 / 3, 12)
    expect(gammaDistributionDensity(0, 0.5, 3)).toBe(Number.POSITIVE_INFINITY)
    expect(gammaDistributionDensity(0, 2, 3)).toBe(0)
    expect(gammaDistributionCdf(0, 2, 3)).toBe(0)
    expect(inverseGammaDistribution(0, 2, 3)).toBeUndefined()
    expect(inverseGammaDistribution(1, 2, 3)).toBeUndefined()
    expect(inverseGammaDistribution(0.5, 0, 3)).toBeUndefined()
    expect(chiSquareDensity(0, 1)).toBe(Number.POSITIVE_INFINITY)
    expect(inverseChiSquare(0, 2)).toBeUndefined()
    expect(inverseChiSquare(1, 2)).toBeUndefined()
    expect(inverseChiSquare(0.5, 0)).toBeUndefined()

    expect(standardNormalPdf(0)).toBeCloseTo(0.3989422804, 10)
    expect(inverseStandardNormal(0)).toBeUndefined()
    expect(inverseStandardNormal(1)).toBeUndefined()
    expect(inverseStandardNormal(0.001)).toBeLessThan(-3)
    expect(inverseStandardNormal(0.999)).toBeGreaterThan(3)
    expect(percentileNormal(10, 2, 10)).toBeCloseTo(0.5, 8)
    expect(inverseNormal(0.5, 10, 2)).toBeCloseTo(10, 12)
    expect(inverseNormal(0, 10, 2)).toBeUndefined()

    expect(skewSample([1, 2])).toBeUndefined()
    expect(skewSample([1, 1, 1])).toBeUndefined()
    expect(skewSample([1, 2, 3])).toBeCloseTo(0, 12)
    expect(skewPopulation([])).toBeUndefined()
    expect(skewPopulation([1, 1])).toBeUndefined()
    expect(skewPopulation([1, 2])).toBeCloseTo(0, 12)
    expect(kurtosis([1, 2, 3])).toBeUndefined()
    expect(kurtosis([1, 1, 1, 1])).toBeUndefined()
    expect(kurtosis([1, 2, 3, 4])).toBeCloseTo(-1.2, 12)

    expect(Number.isNaN(binomialProbability(5, 4, 0.5))).toBe(true)
    expect(binomialProbability(0, 4, 0)).toBe(1)
    expect(binomialProbability(1, 4, 0)).toBe(0)
    expect(binomialProbability(4, 4, 1)).toBe(1)
    expect(binomialProbability(3, 4, 1)).toBe(0)
    expect(Number.isNaN(negativeBinomialProbability(1, 0, 0.5))).toBe(true)
    expect(negativeBinomialProbability(1, 3, 0)).toBe(0)
    expect(negativeBinomialProbability(0, 3, 1)).toBe(1)
    expect(negativeBinomialProbability(2, 3, 1)).toBe(0)
    expect(Number.isNaN(hypergeometricProbability(1, 11, 3, 10))).toBe(true)
    expect(hypergeometricProbability(4, 4, 3, 10)).toBe(0)
    expect(Number.isNaN(poissonProbability(-1, 2))).toBe(true)
    expect(poissonProbability(0, 0)).toBe(1)
    expect(poissonProbability(2, 0)).toBe(0)
  })
})
