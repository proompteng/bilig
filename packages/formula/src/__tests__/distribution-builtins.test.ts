import { describe, expect, it } from "vitest";
import {
  betaDistributionCdf,
  betaDistributionInverse,
  besselIValue,
  besselKValue,
  binomialProbability,
  chiSquareCdf,
  fDistributionCdf,
  gammaFunction,
  hypergeometricProbability,
  inverseChiSquare,
  inverseFDistribution,
  inverseGammaDistribution,
  inverseStandardNormal,
  inverseStudentT,
  negativeBinomialProbability,
  poissonProbability,
  standardNormalCdf,
  studentTCdf,
} from "../builtins/distributions.js";

describe("distribution helpers", () => {
  it("evaluates Bessel and gamma helpers", () => {
    expect(besselIValue(1.5, 1)).toBeCloseTo(0.981666428, 7);
    expect(besselKValue(1.5, 1)).toBeCloseTo(0.277387804, 7);
    expect(gammaFunction(5)).toBeCloseTo(24, 10);
  });

  it("computes normal-family helpers and inverses", () => {
    expect(standardNormalCdf(1)).toBeCloseTo(0.8413447460685429, 7);
    expect(inverseStandardNormal(0.001)).toBeCloseTo(-3.090232306167813, 8);
  });

  it("computes continuous distribution helpers", () => {
    const betaCdf = betaDistributionCdf(2, 8, 10, 1, 3);
    expect(betaCdf).toBeCloseTo(0.6854705810117458, 10);
    expect(betaDistributionInverse(betaCdf, 8, 10, 1, 3)).toBeCloseTo(2, 10);

    expect(fDistributionCdf(15.2068649, 6, 4)).toBeCloseTo(0.99, 9);
    expect(inverseFDistribution(0.01, 6, 4)).toBeCloseTo(0.10930991466299911, 8);

    expect(studentTCdf(1, 1)).toBeCloseTo(0.75, 12);
    expect(inverseStudentT(0.75, 1)).toBeCloseTo(1, 9);

    expect(inverseGammaDistribution(0.08030139707139418, 3, 2)).toBeCloseTo(2, 10);
    expect(chiSquareCdf(3, 4)).toBeCloseTo(0.4421745996289252, 12);
    expect(inverseChiSquare(0.93, 1)).toBeCloseTo(3.283020286759536, 8);
  });

  it("computes discrete probability helpers", () => {
    expect(poissonProbability(3, 2.5)).toBeCloseTo(0.21376301724973648, 12);
    expect(binomialProbability(2, 4, 0.5)).toBeCloseTo(0.375, 12);
    expect(hypergeometricProbability(1, 4, 3, 10)).toBeCloseTo(0.5, 12);
    expect(negativeBinomialProbability(2, 3, 0.5)).toBeCloseTo(0.1875, 12);
  });
});
