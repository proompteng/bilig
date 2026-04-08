import { ErrorCode, ValueTag } from "@bilig/protocol";
import type { CellValue } from "@bilig/protocol";
import {
  betaDistributionCdf,
  betaDistributionDensity,
  betaDistributionInverse,
  binomialProbability,
  chiSquareCdf,
  chiSquareDensity,
  fDistributionCdf,
  fDistributionDensity,
  gammaDistributionCdf,
  gammaDistributionDensity,
  gammaFunction,
  hypergeometricProbability,
  inverseChiSquare,
  inverseFDistribution,
  inverseGammaDistribution,
  inverseStandardNormal,
  inverseStudentT,
  logGamma,
  negativeBinomialProbability,
  poissonProbability,
  regularizedUpperGamma,
  studentTCdf,
  studentTDensity,
} from "./distributions.js";
import { erfApprox } from "./statistics.js";
import type { EvaluationResult } from "../runtime-values.js";

type Builtin = (...args: CellValue[]) => EvaluationResult;

interface DistributionBuiltinDeps {
  toNumber: (value: CellValue) => number | undefined;
  coerceBoolean: (value: CellValue | undefined, fallback: boolean) => boolean | undefined;
  coerceNumber: (value: CellValue | undefined, fallback: number) => number | undefined;
  integerValue: (value: CellValue | undefined, fallback?: number) => number | undefined;
  nonNegativeIntegerValue: (value: CellValue | undefined, fallback?: number) => number | undefined;
  positiveIntegerValue: (value: CellValue | undefined, fallback?: number) => number | undefined;
  numberResult: (value: number) => EvaluationResult;
  numericResultOrError: (value: number) => EvaluationResult;
  valueError: () => EvaluationResult;
}

function naError(): EvaluationResult {
  return { tag: ValueTag.Error, code: ErrorCode.NA };
}

export function createDistributionBuiltins({
  toNumber,
  coerceBoolean,
  coerceNumber,
  integerValue,
  nonNegativeIntegerValue,
  positiveIntegerValue,
  numberResult,
  numericResultOrError,
  valueError,
}: DistributionBuiltinDeps): Record<string, Builtin> {
  const builtins: Record<string, Builtin> = {
    "CONFIDENCE.NORM": (alphaArg, standardDeviationArg, sizeArg) => {
      const alpha = toNumber(alphaArg);
      const standardDeviation = toNumber(standardDeviationArg);
      const size = toNumber(sizeArg);
      if (
        alpha === undefined ||
        standardDeviation === undefined ||
        size === undefined ||
        !(alpha > 0 && alpha < 1) ||
        standardDeviation <= 0 ||
        size < 1
      ) {
        return valueError();
      }
      const criticalValue = inverseStandardNormal(1 - alpha / 2);
      return criticalValue === undefined
        ? valueError()
        : numberResult((criticalValue * standardDeviation) / Math.sqrt(size));
    },
    ERF: (lowerArg, upperArg) => {
      const lower = toNumber(lowerArg);
      if (lower === undefined) {
        return valueError();
      }
      if (upperArg === undefined) {
        return numberResult(erfApprox(lower));
      }
      const upper = toNumber(upperArg);
      return upper === undefined ? valueError() : numberResult(erfApprox(upper) - erfApprox(lower));
    },
    "ERF.PRECISE": (valueArg) => {
      const value = toNumber(valueArg);
      return value === undefined ? valueError() : numberResult(erfApprox(value));
    },
    ERFC: (valueArg) => {
      const value = toNumber(valueArg);
      return value === undefined ? valueError() : numberResult(1 - erfApprox(value));
    },
    "ERFC.PRECISE": (valueArg) => {
      const value = toNumber(valueArg);
      return value === undefined ? valueError() : numberResult(1 - erfApprox(value));
    },
    FISHER: (valueArg) => {
      const value = toNumber(valueArg);
      if (value === undefined || value <= -1 || value >= 1) {
        return valueError();
      }
      return numberResult(0.5 * Math.log((1 + value) / (1 - value)));
    },
    FISHERINV: (valueArg) => {
      const value = toNumber(valueArg);
      if (value === undefined) {
        return valueError();
      }
      const exponent = Math.exp(2 * value);
      return numberResult((exponent - 1) / (exponent + 1));
    },
    GAMMALN: (valueArg) => {
      const value = toNumber(valueArg);
      return value === undefined ? valueError() : numericResultOrError(logGamma(value));
    },
    "GAMMALN.PRECISE": (valueArg) => builtins["GAMMALN"]!(valueArg),
    GAMMA: (valueArg) => {
      const value = toNumber(valueArg);
      return value === undefined ? valueError() : numericResultOrError(gammaFunction(value));
    },
    CONFIDENCE: (alphaArg, standardDeviationArg, sizeArg) =>
      builtins["CONFIDENCE.NORM"]!(alphaArg, standardDeviationArg, sizeArg),
    "CONFIDENCE.T": (alphaArg, standardDeviationArg, sizeArg) => {
      const alpha = toNumber(alphaArg);
      const standardDeviation = toNumber(standardDeviationArg);
      const size = integerValue(sizeArg);
      if (
        alpha === undefined ||
        standardDeviation === undefined ||
        size === undefined ||
        !(alpha > 0 && alpha < 1) ||
        !(standardDeviation > 0) ||
        size < 2
      ) {
        return valueError();
      }
      const critical = inverseStudentT(1 - alpha / 2, size - 1);
      return critical === undefined
        ? valueError()
        : numericResultOrError((critical * standardDeviation) / Math.sqrt(size));
    },
    "BETA.DIST": (xArg, alphaArg, betaArg, cumulativeArg, lowerBoundArg, upperBoundArg) => {
      const x = toNumber(xArg);
      const alpha = toNumber(alphaArg);
      const beta = toNumber(betaArg);
      const cumulative = coerceBoolean(cumulativeArg, false);
      const lowerBound = coerceNumber(lowerBoundArg, 0);
      const upperBound = coerceNumber(upperBoundArg, 1);
      if (
        x === undefined ||
        alpha === undefined ||
        beta === undefined ||
        cumulative === undefined ||
        lowerBound === undefined ||
        upperBound === undefined
      ) {
        return valueError();
      }
      return numericResultOrError(
        cumulative
          ? betaDistributionCdf(x, alpha, beta, lowerBound, upperBound)
          : betaDistributionDensity(x, alpha, beta, lowerBound, upperBound),
      );
    },
    BETADIST: (xArg, alphaArg, betaArg, lowerBoundArg, upperBoundArg) =>
      builtins["BETA.DIST"]!(
        xArg,
        alphaArg,
        betaArg,
        { tag: ValueTag.Boolean, value: true },
        lowerBoundArg,
        upperBoundArg,
      ),
    "BETA.INV": (probabilityArg, alphaArg, betaArg, lowerBoundArg, upperBoundArg) => {
      const probability = toNumber(probabilityArg);
      const alpha = toNumber(alphaArg);
      const beta = toNumber(betaArg);
      const lowerBound = coerceNumber(lowerBoundArg, 0);
      const upperBound = coerceNumber(upperBoundArg, 1);
      if (
        probability === undefined ||
        alpha === undefined ||
        beta === undefined ||
        lowerBound === undefined ||
        upperBound === undefined
      ) {
        return valueError();
      }
      const result = betaDistributionInverse(probability, alpha, beta, lowerBound, upperBound);
      return result === undefined ? valueError() : numericResultOrError(result);
    },
    BETAINV: (probabilityArg, alphaArg, betaArg, lowerBoundArg, upperBoundArg) =>
      builtins["BETA.INV"]!(probabilityArg, alphaArg, betaArg, lowerBoundArg, upperBoundArg),
    EXPONDIST: (xArg, lambdaArg, cumulativeArg) => {
      const x = toNumber(xArg);
      const lambda = toNumber(lambdaArg);
      const cumulative = coerceBoolean(cumulativeArg, false);
      if (
        x === undefined ||
        lambda === undefined ||
        cumulative === undefined ||
        x < 0 ||
        lambda <= 0
      ) {
        return valueError();
      }
      return numberResult(cumulative ? 1 - Math.exp(-lambda * x) : lambda * Math.exp(-lambda * x));
    },
    "EXPON.DIST": (xArg, lambdaArg, cumulativeArg) =>
      builtins["EXPONDIST"]!(xArg, lambdaArg, cumulativeArg),
    POISSON: (eventsArg, meanArg, cumulativeArg) => {
      const events = nonNegativeIntegerValue(eventsArg);
      const mean = toNumber(meanArg);
      const cumulative = coerceBoolean(cumulativeArg, false);
      if (events === undefined || mean === undefined || cumulative === undefined || mean < 0) {
        return valueError();
      }
      if (!cumulative) {
        return numericResultOrError(poissonProbability(events, mean));
      }
      let total = 0;
      for (let index = 0; index <= events; index += 1) {
        total += poissonProbability(index, mean);
      }
      return numericResultOrError(total);
    },
    "POISSON.DIST": (eventsArg, meanArg, cumulativeArg) =>
      builtins["POISSON"]!(eventsArg, meanArg, cumulativeArg),
    WEIBULL: (xArg, alphaArg, betaArg, cumulativeArg) => {
      const x = toNumber(xArg);
      const alpha = toNumber(alphaArg);
      const beta = toNumber(betaArg);
      const cumulative = coerceBoolean(cumulativeArg, false);
      if (
        x === undefined ||
        alpha === undefined ||
        beta === undefined ||
        cumulative === undefined ||
        x < 0 ||
        alpha <= 0 ||
        beta <= 0
      ) {
        return valueError();
      }
      if (cumulative) {
        return numberResult(1 - Math.exp(-((x / beta) ** alpha)));
      }
      if (x === 0) {
        return numberResult(alpha === 1 ? 1 / beta : alpha < 1 ? Number.POSITIVE_INFINITY : 0);
      }
      return numberResult(
        (alpha / beta ** alpha) * x ** (alpha - 1) * Math.exp(-((x / beta) ** alpha)),
      );
    },
    "WEIBULL.DIST": (xArg, alphaArg, betaArg, cumulativeArg) =>
      builtins["WEIBULL"]!(xArg, alphaArg, betaArg, cumulativeArg),
    GAMMADIST: (xArg, alphaArg, betaArg, cumulativeArg) => {
      const x = toNumber(xArg);
      const alpha = toNumber(alphaArg);
      const beta = toNumber(betaArg);
      const cumulative = coerceBoolean(cumulativeArg, false);
      if (
        x === undefined ||
        alpha === undefined ||
        beta === undefined ||
        cumulative === undefined ||
        x < 0 ||
        alpha <= 0 ||
        beta <= 0
      ) {
        return valueError();
      }
      return numberResult(
        cumulative
          ? gammaDistributionCdf(x, alpha, beta)
          : gammaDistributionDensity(x, alpha, beta),
      );
    },
    "GAMMA.DIST": (xArg, alphaArg, betaArg, cumulativeArg) =>
      builtins["GAMMADIST"]!(xArg, alphaArg, betaArg, cumulativeArg),
    "GAMMA.INV": (probabilityArg, alphaArg, betaArg) => {
      const probability = toNumber(probabilityArg);
      const alpha = toNumber(alphaArg);
      const beta = toNumber(betaArg);
      if (
        probability === undefined ||
        alpha === undefined ||
        beta === undefined ||
        !(probability > 0 && probability < 1) ||
        !(alpha > 0) ||
        !(beta > 0)
      ) {
        return valueError();
      }
      const result = inverseGammaDistribution(probability, alpha, beta);
      return result === undefined ? valueError() : numericResultOrError(result);
    },
    GAMMAINV: (probabilityArg, alphaArg, betaArg) =>
      builtins["GAMMA.INV"]!(probabilityArg, alphaArg, betaArg),
    CHIDIST: (xArg, degreesArg) => {
      const x = toNumber(xArg);
      const degrees = toNumber(degreesArg);
      if (x === undefined || degrees === undefined || x < 0 || degrees < 1) {
        return valueError();
      }
      return numericResultOrError(regularizedUpperGamma(degrees / 2, x / 2));
    },
    "LEGACY.CHIDIST": (xArg, degreesArg) => builtins["CHIDIST"]!(xArg, degreesArg),
    CHIINV: (probabilityArg, degreesArg) => builtins["CHISQ.INV.RT"]!(probabilityArg, degreesArg),
    "CHISQ.DIST.RT": (xArg, degreesArg) => builtins["CHIDIST"]!(xArg, degreesArg),
    CHISQDIST: (xArg, degreesArg) => builtins["CHISQ.DIST.RT"]!(xArg, degreesArg),
    "CHISQ.DIST": (xArg, degreesArg, cumulativeArg) => {
      const x = toNumber(xArg);
      const degrees = toNumber(degreesArg);
      const cumulative = coerceBoolean(cumulativeArg, false);
      if (
        x === undefined ||
        degrees === undefined ||
        cumulative === undefined ||
        x < 0 ||
        degrees < 1
      ) {
        return valueError();
      }
      return numberResult(cumulative ? chiSquareCdf(x, degrees) : chiSquareDensity(x, degrees));
    },
    "CHISQ.INV.RT": (probabilityArg, degreesArg) => {
      const probability = toNumber(probabilityArg);
      const degrees = toNumber(degreesArg);
      if (
        probability === undefined ||
        degrees === undefined ||
        !(probability > 0 && probability < 1) ||
        !(degrees >= 1)
      ) {
        return valueError();
      }
      const result = inverseChiSquare(1 - probability, degrees);
      return result === undefined ? naError() : numberResult(result);
    },
    CHISQINV: (probabilityArg, degreesArg) => builtins["CHISQ.INV.RT"]!(probabilityArg, degreesArg),
    "LEGACY.CHIINV": (probabilityArg, degreesArg) =>
      builtins["CHISQ.INV.RT"]!(probabilityArg, degreesArg),
    "CHISQ.INV": (probabilityArg, degreesArg) => {
      const probability = toNumber(probabilityArg);
      const degrees = toNumber(degreesArg);
      if (
        probability === undefined ||
        degrees === undefined ||
        !(probability > 0 && probability < 1) ||
        !(degrees >= 1)
      ) {
        return valueError();
      }
      const result = inverseChiSquare(probability, degrees);
      return result === undefined ? naError() : numberResult(result);
    },
    "F.DIST": (xArg, degrees1Arg, degrees2Arg, cumulativeArg) => {
      const x = toNumber(xArg);
      const degrees1 = integerValue(degrees1Arg);
      const degrees2 = integerValue(degrees2Arg);
      const cumulative = coerceBoolean(cumulativeArg, false);
      if (
        x === undefined ||
        degrees1 === undefined ||
        degrees2 === undefined ||
        cumulative === undefined
      ) {
        return valueError();
      }
      return numericResultOrError(
        cumulative
          ? fDistributionCdf(x, degrees1, degrees2)
          : fDistributionDensity(x, degrees1, degrees2),
      );
    },
    "F.DIST.RT": (xArg, degrees1Arg, degrees2Arg) => {
      const x = toNumber(xArg);
      const degrees1 = integerValue(degrees1Arg);
      const degrees2 = integerValue(degrees2Arg);
      if (x === undefined || degrees1 === undefined || degrees2 === undefined) {
        return valueError();
      }
      return numericResultOrError(1 - fDistributionCdf(x, degrees1, degrees2));
    },
    FDIST: (xArg, degrees1Arg, degrees2Arg) =>
      builtins["F.DIST.RT"]!(xArg, degrees1Arg, degrees2Arg),
    "LEGACY.FDIST": (xArg, degrees1Arg, degrees2Arg) =>
      builtins["F.DIST.RT"]!(xArg, degrees1Arg, degrees2Arg),
    "F.INV": (probabilityArg, degrees1Arg, degrees2Arg) => {
      const probability = toNumber(probabilityArg);
      const degrees1 = integerValue(degrees1Arg);
      const degrees2 = integerValue(degrees2Arg);
      if (probability === undefined || degrees1 === undefined || degrees2 === undefined) {
        return valueError();
      }
      const result = inverseFDistribution(probability, degrees1, degrees2);
      return result === undefined ? valueError() : numericResultOrError(result);
    },
    "F.INV.RT": (probabilityArg, degrees1Arg, degrees2Arg) => {
      const probability = toNumber(probabilityArg);
      const degrees1 = integerValue(degrees1Arg);
      const degrees2 = integerValue(degrees2Arg);
      if (probability === undefined || degrees1 === undefined || degrees2 === undefined) {
        return valueError();
      }
      const result = inverseFDistribution(1 - probability, degrees1, degrees2);
      return result === undefined ? valueError() : numericResultOrError(result);
    },
    FINV: (probabilityArg, degrees1Arg, degrees2Arg) =>
      builtins["F.INV.RT"]!(probabilityArg, degrees1Arg, degrees2Arg),
    "LEGACY.FINV": (probabilityArg, degrees1Arg, degrees2Arg) =>
      builtins["F.INV.RT"]!(probabilityArg, degrees1Arg, degrees2Arg),
    "T.DIST": (xArg, degreesArg, cumulativeArg) => {
      const x = toNumber(xArg);
      const degrees = integerValue(degreesArg);
      const cumulative = coerceBoolean(cumulativeArg, false);
      if (x === undefined || degrees === undefined || cumulative === undefined) {
        return valueError();
      }
      return numericResultOrError(
        cumulative ? studentTCdf(x, degrees) : studentTDensity(x, degrees),
      );
    },
    "T.DIST.RT": (xArg, degreesArg) => {
      const x = toNumber(xArg);
      const degrees = integerValue(degreesArg);
      if (x === undefined || degrees === undefined) {
        return valueError();
      }
      return numericResultOrError(1 - studentTCdf(x, degrees));
    },
    "T.DIST.2T": (xArg, degreesArg) => {
      const x = toNumber(xArg);
      const degrees = integerValue(degreesArg);
      if (x === undefined || degrees === undefined || x < 0) {
        return valueError();
      }
      return numericResultOrError(Math.min(1, 2 * (1 - studentTCdf(x, degrees))));
    },
    TDIST: (xArg, degreesArg, tailsArg) => {
      const x = toNumber(xArg);
      const degrees = integerValue(degreesArg);
      const tails = integerValue(tailsArg);
      if (x === undefined || degrees === undefined || tails === undefined || x < 0) {
        return valueError();
      }
      if (tails !== 1 && tails !== 2) {
        return valueError();
      }
      const upperTail = 1 - studentTCdf(x, degrees);
      return numericResultOrError(tails === 1 ? upperTail : Math.min(1, upperTail * 2));
    },
    "T.INV": (probabilityArg, degreesArg) => {
      const probability = toNumber(probabilityArg);
      const degrees = integerValue(degreesArg);
      if (probability === undefined || degrees === undefined) {
        return valueError();
      }
      const result = inverseStudentT(probability, degrees);
      return result === undefined ? valueError() : numericResultOrError(result);
    },
    "T.INV.2T": (probabilityArg, degreesArg) => {
      const probability = toNumber(probabilityArg);
      const degrees = integerValue(degreesArg);
      if (
        probability === undefined ||
        degrees === undefined ||
        !(probability > 0 && probability < 1)
      ) {
        return valueError();
      }
      const result = inverseStudentT(1 - probability / 2, degrees);
      return result === undefined ? valueError() : numericResultOrError(result);
    },
    TINV: (probabilityArg, degreesArg) => builtins["T.INV.2T"]!(probabilityArg, degreesArg),
    BINOMDIST: (successesArg, trialsArg, probabilityArg, cumulativeArg) => {
      const successes = nonNegativeIntegerValue(successesArg);
      const trials = nonNegativeIntegerValue(trialsArg);
      const probability = toNumber(probabilityArg);
      const cumulative = coerceBoolean(cumulativeArg, false);
      if (
        successes === undefined ||
        trials === undefined ||
        probability === undefined ||
        cumulative === undefined ||
        successes > trials ||
        probability < 0 ||
        probability > 1
      ) {
        return valueError();
      }
      if (!cumulative) {
        return numericResultOrError(binomialProbability(successes, trials, probability));
      }
      let total = 0;
      for (let index = 0; index <= successes; index += 1) {
        total += binomialProbability(index, trials, probability);
      }
      return numericResultOrError(total);
    },
    "BINOM.DIST": (successesArg, trialsArg, probabilityArg, cumulativeArg) =>
      builtins["BINOMDIST"]!(successesArg, trialsArg, probabilityArg, cumulativeArg),
    "BINOM.DIST.RANGE": (trialsArg, probabilityArg, successesArg, upperSuccessesArg) => {
      const trials = nonNegativeIntegerValue(trialsArg);
      const probability = toNumber(probabilityArg);
      const lower = nonNegativeIntegerValue(successesArg);
      const upper = nonNegativeIntegerValue(upperSuccessesArg, lower);
      if (
        trials === undefined ||
        probability === undefined ||
        lower === undefined ||
        upper === undefined ||
        lower > upper ||
        upper > trials ||
        probability < 0 ||
        probability > 1
      ) {
        return valueError();
      }
      let total = 0;
      for (let index = lower; index <= upper; index += 1) {
        total += binomialProbability(index, trials, probability);
      }
      return numericResultOrError(total);
    },
    CRITBINOM: (trialsArg, probabilityArg, alphaArg) => {
      const trials = nonNegativeIntegerValue(trialsArg);
      const probability = toNumber(probabilityArg);
      const alpha = toNumber(alphaArg);
      if (
        trials === undefined ||
        probability === undefined ||
        alpha === undefined ||
        probability < 0 ||
        probability > 1 ||
        alpha <= 0 ||
        alpha >= 1
      ) {
        return valueError();
      }
      let cumulative = 0;
      for (let index = 0; index <= trials; index += 1) {
        cumulative += binomialProbability(index, trials, probability);
        if (cumulative >= alpha) {
          return numberResult(index);
        }
      }
      return numberResult(trials);
    },
    "BINOM.INV": (trialsArg, probabilityArg, alphaArg) =>
      builtins["CRITBINOM"]!(trialsArg, probabilityArg, alphaArg),
    HYPGEOMDIST: (sampleSuccessesArg, sampleSizeArg, populationSuccessesArg, populationSizeArg) => {
      const sampleSuccesses = nonNegativeIntegerValue(sampleSuccessesArg);
      const sampleSize = nonNegativeIntegerValue(sampleSizeArg);
      const populationSuccesses = nonNegativeIntegerValue(populationSuccessesArg);
      const populationSize = positiveIntegerValue(populationSizeArg);
      if (
        sampleSuccesses === undefined ||
        sampleSize === undefined ||
        populationSuccesses === undefined ||
        populationSize === undefined
      ) {
        return valueError();
      }
      return numericResultOrError(
        hypergeometricProbability(sampleSuccesses, sampleSize, populationSuccesses, populationSize),
      );
    },
    "HYPGEOM.DIST": (
      sampleSuccessesArg,
      sampleSizeArg,
      populationSuccessesArg,
      populationSizeArg,
      cumulativeArg,
    ) => {
      const sampleSuccesses = nonNegativeIntegerValue(sampleSuccessesArg);
      const sampleSize = nonNegativeIntegerValue(sampleSizeArg);
      const populationSuccesses = nonNegativeIntegerValue(populationSuccessesArg);
      const populationSize = positiveIntegerValue(populationSizeArg);
      const cumulative = coerceBoolean(cumulativeArg, false);
      if (
        sampleSuccesses === undefined ||
        sampleSize === undefined ||
        populationSuccesses === undefined ||
        populationSize === undefined ||
        cumulative === undefined
      ) {
        return valueError();
      }
      if (!cumulative) {
        return numericResultOrError(
          hypergeometricProbability(
            sampleSuccesses,
            sampleSize,
            populationSuccesses,
            populationSize,
          ),
        );
      }
      const minimum = Math.max(0, sampleSize - (populationSize - populationSuccesses));
      let total = 0;
      for (let index = minimum; index <= sampleSuccesses; index += 1) {
        total += hypergeometricProbability(index, sampleSize, populationSuccesses, populationSize);
      }
      return numericResultOrError(total);
    },
    NEGBINOMDIST: (failuresArg, successesArg, probabilityArg) => {
      const failures = nonNegativeIntegerValue(failuresArg);
      const successes = positiveIntegerValue(successesArg);
      const probability = toNumber(probabilityArg);
      if (
        failures === undefined ||
        successes === undefined ||
        probability === undefined ||
        probability < 0 ||
        probability > 1
      ) {
        return valueError();
      }
      return numericResultOrError(negativeBinomialProbability(failures, successes, probability));
    },
    "NEGBINOM.DIST": (failuresArg, successesArg, probabilityArg, cumulativeArg) => {
      const failures = nonNegativeIntegerValue(failuresArg);
      const successes = positiveIntegerValue(successesArg);
      const probability = toNumber(probabilityArg);
      const cumulative = coerceBoolean(cumulativeArg, false);
      if (
        failures === undefined ||
        successes === undefined ||
        probability === undefined ||
        cumulative === undefined ||
        probability < 0 ||
        probability > 1
      ) {
        return valueError();
      }
      if (!cumulative) {
        return numericResultOrError(negativeBinomialProbability(failures, successes, probability));
      }
      let total = 0;
      for (let index = 0; index <= failures; index += 1) {
        total += negativeBinomialProbability(index, successes, probability);
      }
      return numericResultOrError(total);
    },
  };

  return builtins;
}
