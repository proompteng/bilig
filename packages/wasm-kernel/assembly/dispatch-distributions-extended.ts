import { BuiltinId, ErrorCode, ValueTag } from "./protocol";
import { toNumberExact } from "./operands";
import { isNumericResult, rangeSupportedScalarOnly, scalarErrorAt } from "./builtin-args";
import { STACK_KIND_SCALAR, writeResult } from "./result-io";
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
  hypergeometricProbability,
  inverseChiSquare,
  inverseFDistribution,
  inverseStudentT,
  negativeBinomialProbability,
  poissonProbability,
  regularizedUpperGamma,
  studentTCdf,
  studentTDensity,
} from "./distributions";

function coerceBoolean(tag: u8, value: f64): i32 {
  if (tag == ValueTag.Boolean || tag == ValueTag.Number) {
    return value != 0 ? 1 : 0;
  }
  if (tag == ValueTag.Empty) {
    return 0;
  }
  return -1;
}

export function tryApplyExtendedDistributionBuiltin(
  builtinId: i32,
  argc: i32,
  base: i32,
  rangeIndexStack: Uint32Array,
  valueStack: Float64Array,
  tagStack: Uint8Array,
  kindStack: Uint8Array,
): i32 {
  if (
    (builtinId == BuiltinId.Expondist ||
      builtinId == BuiltinId.ExponDist ||
      builtinId == BuiltinId.Poisson ||
      builtinId == BuiltinId.PoissonDist ||
      builtinId == BuiltinId.Negbinomdist) &&
    argc == 3
  ) {
    if (!rangeSupportedScalarOnly(base, argc, kindStack)) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack);
    if (scalarError >= 0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        scalarError,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    let result = NaN;
    if (builtinId == BuiltinId.Expondist || builtinId == BuiltinId.ExponDist) {
      const x = toNumberExact(tagStack[base], valueStack[base]);
      const lambda = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
      const cumulative = coerceBoolean(tagStack[base + 2], valueStack[base + 2]);
      result =
        isNaN(x) || isNaN(lambda) || cumulative < 0 || x < 0.0 || lambda <= 0.0
          ? NaN
          : cumulative == 1
            ? 1.0 - Math.exp(-lambda * x)
            : lambda * Math.exp(-lambda * x);
    } else if (builtinId == BuiltinId.Poisson || builtinId == BuiltinId.PoissonDist) {
      const eventsRaw = toNumberExact(tagStack[base], valueStack[base]);
      const mean = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
      const cumulative = coerceBoolean(tagStack[base + 2], valueStack[base + 2]);
      const events = <i32>eventsRaw;
      if (!isNaN(eventsRaw) && mean >= 0.0 && cumulative >= 0 && events >= 0) {
        if (cumulative == 1) {
          result = 0.0;
          for (let index = 0; index <= events; index += 1) {
            result += poissonProbability(index, mean);
          }
        } else {
          result = poissonProbability(events, mean);
        }
      }
    } else {
      const failuresRaw = toNumberExact(tagStack[base], valueStack[base]);
      const successesRaw = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
      const probability = toNumberExact(tagStack[base + 2], valueStack[base + 2]);
      const failures = <i32>failuresRaw;
      const successes = <i32>successesRaw;
      if (!isNaN(failuresRaw) && !isNaN(successesRaw)) {
        result = negativeBinomialProbability(failures, successes, probability);
      }
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      isNumericResult(result) ? <u8>ValueTag.Number : <u8>ValueTag.Error,
      isNumericResult(result) ? result : ErrorCode.Value,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (
    (builtinId == BuiltinId.Weibull ||
      builtinId == BuiltinId.WeibullDist ||
      builtinId == BuiltinId.Gammadist ||
      builtinId == BuiltinId.GammaDist ||
      builtinId == BuiltinId.Binomdist ||
      builtinId == BuiltinId.BinomDist ||
      builtinId == BuiltinId.NegbinomDist) &&
    argc == 4
  ) {
    if (!rangeSupportedScalarOnly(base, argc, kindStack)) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack);
    if (scalarError >= 0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        scalarError,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    let result = NaN;
    if (builtinId == BuiltinId.Weibull || builtinId == BuiltinId.WeibullDist) {
      const x = toNumberExact(tagStack[base], valueStack[base]);
      const alpha = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
      const beta = toNumberExact(tagStack[base + 2], valueStack[base + 2]);
      const cumulative = coerceBoolean(tagStack[base + 3], valueStack[base + 3]);
      if (
        !isNaN(x) &&
        !isNaN(alpha) &&
        !isNaN(beta) &&
        cumulative >= 0 &&
        x >= 0.0 &&
        alpha > 0.0 &&
        beta > 0.0
      ) {
        if (cumulative == 1) {
          result = 1.0 - Math.exp(-Math.pow(x / beta, alpha));
        } else if (x == 0.0) {
          result = alpha == 1.0 ? 1.0 / beta : alpha < 1.0 ? Infinity : 0.0;
        } else {
          result =
            (alpha / Math.pow(beta, alpha)) *
            Math.pow(x, alpha - 1.0) *
            Math.exp(-Math.pow(x / beta, alpha));
        }
      }
    } else if (builtinId == BuiltinId.Gammadist || builtinId == BuiltinId.GammaDist) {
      const x = toNumberExact(tagStack[base], valueStack[base]);
      const alpha = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
      const beta = toNumberExact(tagStack[base + 2], valueStack[base + 2]);
      const cumulative = coerceBoolean(tagStack[base + 3], valueStack[base + 3]);
      if (
        !isNaN(x) &&
        !isNaN(alpha) &&
        !isNaN(beta) &&
        cumulative >= 0 &&
        x >= 0.0 &&
        alpha > 0.0 &&
        beta > 0.0
      ) {
        result =
          cumulative == 1
            ? gammaDistributionCdf(x, alpha, beta)
            : gammaDistributionDensity(x, alpha, beta);
      }
    } else if (builtinId == BuiltinId.Binomdist || builtinId == BuiltinId.BinomDist) {
      const successesRaw = toNumberExact(tagStack[base], valueStack[base]);
      const trialsRaw = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
      const probability = toNumberExact(tagStack[base + 2], valueStack[base + 2]);
      const cumulative = coerceBoolean(tagStack[base + 3], valueStack[base + 3]);
      const successes = <i32>successesRaw;
      const trials = <i32>trialsRaw;
      if (!isNaN(successesRaw) && !isNaN(trialsRaw) && cumulative >= 0) {
        if (cumulative == 1) {
          result = 0.0;
          for (let index = 0; index <= successes; index += 1) {
            result += binomialProbability(index, trials, probability);
          }
        } else {
          result = binomialProbability(successes, trials, probability);
        }
      }
    } else {
      const failuresRaw = toNumberExact(tagStack[base], valueStack[base]);
      const successesRaw = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
      const probability = toNumberExact(tagStack[base + 2], valueStack[base + 2]);
      const cumulative = coerceBoolean(tagStack[base + 3], valueStack[base + 3]);
      const failures = <i32>failuresRaw;
      const successes = <i32>successesRaw;
      if (!isNaN(failuresRaw) && !isNaN(successesRaw) && cumulative >= 0) {
        if (cumulative == 1) {
          result = 0.0;
          for (let index = 0; index <= failures; index += 1) {
            result += negativeBinomialProbability(index, successes, probability);
          }
        } else {
          result = negativeBinomialProbability(failures, successes, probability);
        }
      }
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      isNumericResult(result) ? <u8>ValueTag.Number : <u8>ValueTag.Error,
      isNumericResult(result) ? result : ErrorCode.Value,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (
    (builtinId == BuiltinId.Chidist ||
      builtinId == BuiltinId.LegacyChidist ||
      builtinId == BuiltinId.ChisqDistRt ||
      builtinId == BuiltinId.Chisqdist) &&
    argc == 2
  ) {
    if (!rangeSupportedScalarOnly(base, argc, kindStack)) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack);
    if (scalarError >= 0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        scalarError,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const x = toNumberExact(tagStack[base], valueStack[base]);
    const degrees = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
    const result =
      !isNaN(x) && !isNaN(degrees) && x >= 0.0 && degrees >= 1.0
        ? regularizedUpperGamma(degrees / 2.0, x / 2.0)
        : NaN;
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      isNumericResult(result) ? <u8>ValueTag.Number : <u8>ValueTag.Error,
      isNumericResult(result) ? result : ErrorCode.Value,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.ChisqDist && argc == 3) {
    if (!rangeSupportedScalarOnly(base, argc, kindStack)) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack);
    if (scalarError >= 0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        scalarError,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const x = toNumberExact(tagStack[base], valueStack[base]);
    const degrees = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
    const cumulative = coerceBoolean(tagStack[base + 2], valueStack[base + 2]);
    const result =
      !isNaN(x) && !isNaN(degrees) && cumulative >= 0 && x >= 0.0 && degrees >= 1.0
        ? cumulative == 1
          ? chiSquareCdf(x, degrees)
          : chiSquareDensity(x, degrees)
        : NaN;
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      isNumericResult(result) ? <u8>ValueTag.Number : <u8>ValueTag.Error,
      isNumericResult(result) ? result : ErrorCode.Value,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (
    (builtinId == BuiltinId.Chiinv ||
      builtinId == BuiltinId.ChisqInvRt ||
      builtinId == BuiltinId.Chisqinv ||
      builtinId == BuiltinId.LegacyChiinv) &&
    argc == 2
  ) {
    if (!rangeSupportedScalarOnly(base, argc, kindStack)) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack);
    if (scalarError >= 0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        scalarError,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const probability = toNumberExact(tagStack[base], valueStack[base]);
    const degrees = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
    const valid =
      !isNaN(probability) &&
      !isNaN(degrees) &&
      probability > 0.0 &&
      probability < 1.0 &&
      degrees >= 1.0;
    let result = NaN;
    if (valid) {
      result = inverseChiSquare(1.0 - probability, degrees);
    }
    const resultTag = isNumericResult(result) ? <u8>ValueTag.Number : <u8>ValueTag.Error;
    const resultValue = isNumericResult(result) ? result : valid ? ErrorCode.NA : ErrorCode.Value;
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      resultTag,
      resultValue,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.ChisqInv && argc == 2) {
    if (!rangeSupportedScalarOnly(base, argc, kindStack)) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack);
    if (scalarError >= 0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        scalarError,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const probability = toNumberExact(tagStack[base], valueStack[base]);
    const degrees = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
    const valid =
      !isNaN(probability) &&
      !isNaN(degrees) &&
      probability > 0.0 &&
      probability < 1.0 &&
      degrees >= 1.0;
    let result = NaN;
    if (valid) {
      result = inverseChiSquare(probability, degrees);
    }
    const resultTag = isNumericResult(result) ? <u8>ValueTag.Number : <u8>ValueTag.Error;
    const resultValue = isNumericResult(result) ? result : valid ? ErrorCode.NA : ErrorCode.Value;
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      resultTag,
      resultValue,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (
    (builtinId == BuiltinId.BetaDist && argc >= 4 && argc <= 6) ||
    (builtinId == BuiltinId.Betadist && argc >= 3 && argc <= 5)
  ) {
    if (!rangeSupportedScalarOnly(base, argc, kindStack)) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack);
    if (scalarError >= 0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        scalarError,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const x = toNumberExact(tagStack[base], valueStack[base]);
    const alpha = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
    const beta = toNumberExact(tagStack[base + 2], valueStack[base + 2]);
    const modern = builtinId == BuiltinId.BetaDist;
    const cumulative = modern ? coerceBoolean(tagStack[base + 3], valueStack[base + 3]) : 1;
    const lowerBound = modern
      ? argc >= 5
        ? toNumberExact(tagStack[base + 4], valueStack[base + 4])
        : 0.0
      : argc >= 4
        ? toNumberExact(tagStack[base + 3], valueStack[base + 3])
        : 0.0;
    const upperBound = modern
      ? argc >= 6
        ? toNumberExact(tagStack[base + 5], valueStack[base + 5])
        : 1.0
      : argc >= 5
        ? toNumberExact(tagStack[base + 4], valueStack[base + 4])
        : 1.0;
    const result =
      !isNaN(x) &&
      !isNaN(alpha) &&
      !isNaN(beta) &&
      cumulative >= 0 &&
      !isNaN(lowerBound) &&
      !isNaN(upperBound)
        ? cumulative == 1 || !modern
          ? betaDistributionCdf(x, alpha, beta, lowerBound, upperBound)
          : betaDistributionDensity(x, alpha, beta, lowerBound, upperBound)
        : NaN;
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      isNumericResult(result) ? <u8>ValueTag.Number : <u8>ValueTag.Error,
      isNumericResult(result) ? result : ErrorCode.Value,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (
    (builtinId == BuiltinId.BetaInv || builtinId == BuiltinId.Betainv) &&
    argc >= 3 &&
    argc <= 5
  ) {
    if (!rangeSupportedScalarOnly(base, argc, kindStack)) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack);
    if (scalarError >= 0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        scalarError,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const probability = toNumberExact(tagStack[base], valueStack[base]);
    const alpha = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
    const beta = toNumberExact(tagStack[base + 2], valueStack[base + 2]);
    const lowerBound = argc >= 4 ? toNumberExact(tagStack[base + 3], valueStack[base + 3]) : 0.0;
    const upperBound = argc >= 5 ? toNumberExact(tagStack[base + 4], valueStack[base + 4]) : 1.0;
    const result = betaDistributionInverse(probability, alpha, beta, lowerBound, upperBound);
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      isNumericResult(result) ? <u8>ValueTag.Number : <u8>ValueTag.Error,
      isNumericResult(result) ? result : ErrorCode.Value,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.FDist && argc == 4) {
    if (!rangeSupportedScalarOnly(base, argc, kindStack)) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack);
    if (scalarError >= 0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        scalarError,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const x = toNumberExact(tagStack[base], valueStack[base]);
    const degrees1Raw = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
    const degrees2Raw = toNumberExact(tagStack[base + 2], valueStack[base + 2]);
    const cumulative = coerceBoolean(tagStack[base + 3], valueStack[base + 3]);
    const degrees1 = Math.floor(degrees1Raw);
    const degrees2 = Math.floor(degrees2Raw);
    const result =
      !isNaN(x) && !isNaN(degrees1Raw) && !isNaN(degrees2Raw) && cumulative >= 0
        ? cumulative == 1
          ? fDistributionCdf(x, degrees1, degrees2)
          : fDistributionDensity(x, degrees1, degrees2)
        : NaN;
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      isNumericResult(result) ? <u8>ValueTag.Number : <u8>ValueTag.Error,
      isNumericResult(result) ? result : ErrorCode.Value,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (
    (builtinId == BuiltinId.FDistRt ||
      builtinId == BuiltinId.Fdist ||
      builtinId == BuiltinId.LegacyFdist) &&
    argc == 3
  ) {
    if (!rangeSupportedScalarOnly(base, argc, kindStack)) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack);
    if (scalarError >= 0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        scalarError,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const x = toNumberExact(tagStack[base], valueStack[base]);
    const degrees1Raw = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
    const degrees2Raw = toNumberExact(tagStack[base + 2], valueStack[base + 2]);
    const degrees1 = Math.floor(degrees1Raw);
    const degrees2 = Math.floor(degrees2Raw);
    const cdf = fDistributionCdf(x, degrees1, degrees2);
    const result =
      !isNaN(x) && !isNaN(degrees1Raw) && !isNaN(degrees2Raw) && isFinite(cdf) ? 1.0 - cdf : NaN;
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      isNumericResult(result) ? <u8>ValueTag.Number : <u8>ValueTag.Error,
      isNumericResult(result) ? result : ErrorCode.Value,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (
    (builtinId == BuiltinId.FInv ||
      builtinId == BuiltinId.FInvRt ||
      builtinId == BuiltinId.Finv ||
      builtinId == BuiltinId.LegacyFinv) &&
    argc == 3
  ) {
    if (!rangeSupportedScalarOnly(base, argc, kindStack)) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack);
    if (scalarError >= 0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        scalarError,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const probabilityRaw = toNumberExact(tagStack[base], valueStack[base]);
    const degrees1Raw = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
    const degrees2Raw = toNumberExact(tagStack[base + 2], valueStack[base + 2]);
    const degrees1 = Math.floor(degrees1Raw);
    const degrees2 = Math.floor(degrees2Raw);
    const probability = builtinId == BuiltinId.FInv ? probabilityRaw : 1.0 - probabilityRaw;
    const result = inverseFDistribution(probability, degrees1, degrees2);
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      isNumericResult(result) ? <u8>ValueTag.Number : <u8>ValueTag.Error,
      isNumericResult(result) ? result : ErrorCode.Value,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.TDist && argc == 3) {
    if (!rangeSupportedScalarOnly(base, argc, kindStack)) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack);
    if (scalarError >= 0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        scalarError,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const x = toNumberExact(tagStack[base], valueStack[base]);
    const degreesRaw = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
    const cumulative = coerceBoolean(tagStack[base + 2], valueStack[base + 2]);
    const degrees = Math.floor(degreesRaw);
    const result =
      !isNaN(x) && !isNaN(degreesRaw) && cumulative >= 0
        ? cumulative == 1
          ? studentTCdf(x, degrees)
          : studentTDensity(x, degrees)
        : NaN;
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      isNumericResult(result) ? <u8>ValueTag.Number : <u8>ValueTag.Error,
      isNumericResult(result) ? result : ErrorCode.Value,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if ((builtinId == BuiltinId.TDistRt || builtinId == BuiltinId.TDist2T) && argc == 2) {
    if (!rangeSupportedScalarOnly(base, argc, kindStack)) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack);
    if (scalarError >= 0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        scalarError,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const x = toNumberExact(tagStack[base], valueStack[base]);
    const degreesRaw = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
    const degrees = Math.floor(degreesRaw);
    const upperTail = 1.0 - studentTCdf(x, degrees);
    const result =
      !isNaN(x) &&
      !isNaN(degreesRaw) &&
      (builtinId != BuiltinId.TDist2T || x >= 0.0) &&
      isFinite(upperTail)
        ? builtinId == BuiltinId.TDistRt
          ? upperTail
          : min(1.0, upperTail * 2.0)
        : NaN;
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      isNumericResult(result) ? <u8>ValueTag.Number : <u8>ValueTag.Error,
      isNumericResult(result) ? result : ErrorCode.Value,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Tdist && argc == 3) {
    if (!rangeSupportedScalarOnly(base, argc, kindStack)) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack);
    if (scalarError >= 0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        scalarError,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const x = toNumberExact(tagStack[base], valueStack[base]);
    const degreesRaw = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
    const tailsRaw = toNumberExact(tagStack[base + 2], valueStack[base + 2]);
    const degrees = Math.floor(degreesRaw);
    const tails = <i32>tailsRaw;
    const upperTail = 1.0 - studentTCdf(x, degrees);
    const result =
      !isNaN(x) &&
      !isNaN(degreesRaw) &&
      !isNaN(tailsRaw) &&
      tailsRaw == <f64>tails &&
      x >= 0.0 &&
      (tails == 1 || tails == 2) &&
      isFinite(upperTail)
        ? tails == 1
          ? upperTail
          : min(1.0, upperTail * 2.0)
        : NaN;
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      isNumericResult(result) ? <u8>ValueTag.Number : <u8>ValueTag.Error,
      isNumericResult(result) ? result : ErrorCode.Value,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (
    (builtinId == BuiltinId.TInv || builtinId == BuiltinId.TInv2T || builtinId == BuiltinId.Tinv) &&
    argc == 2
  ) {
    if (!rangeSupportedScalarOnly(base, argc, kindStack)) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack);
    if (scalarError >= 0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        scalarError,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const probabilityRaw = toNumberExact(tagStack[base], valueStack[base]);
    const degreesRaw = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
    const degrees = Math.floor(degreesRaw);
    const probability = builtinId == BuiltinId.TInv ? probabilityRaw : 1.0 - probabilityRaw / 2.0;
    const result = inverseStudentT(probability, degrees);
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      isNumericResult(result) ? <u8>ValueTag.Number : <u8>ValueTag.Error,
      isNumericResult(result) ? result : ErrorCode.Value,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (
    (builtinId == BuiltinId.BinomDistRange && (argc == 3 || argc == 4)) ||
    (builtinId == BuiltinId.Critbinom && argc == 3) ||
    (builtinId == BuiltinId.BinomInv && argc == 3) ||
    (builtinId == BuiltinId.Hypgeomdist && argc == 4) ||
    (builtinId == BuiltinId.HypgeomDist && argc == 5)
  ) {
    if (!rangeSupportedScalarOnly(base, argc, kindStack)) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack);
    if (scalarError >= 0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        scalarError,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    let result = NaN;
    if (builtinId == BuiltinId.BinomDistRange) {
      const trialsRaw = toNumberExact(tagStack[base], valueStack[base]);
      const probability = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
      const lowerRaw = toNumberExact(tagStack[base + 2], valueStack[base + 2]);
      const upperRaw =
        argc == 4 ? toNumberExact(tagStack[base + 3], valueStack[base + 3]) : lowerRaw;
      const trials = <i32>trialsRaw;
      const lower = <i32>lowerRaw;
      const upper = <i32>upperRaw;
      if (!isNaN(trialsRaw) && !isNaN(lowerRaw) && !isNaN(upperRaw) && lower <= upper) {
        result = 0.0;
        for (let index = lower; index <= upper; index += 1) {
          result += binomialProbability(index, trials, probability);
        }
      }
    } else if (builtinId == BuiltinId.Critbinom || builtinId == BuiltinId.BinomInv) {
      const trialsRaw = toNumberExact(tagStack[base], valueStack[base]);
      const probability = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
      const alpha = toNumberExact(tagStack[base + 2], valueStack[base + 2]);
      const trials = <i32>trialsRaw;
      if (
        !isNaN(trialsRaw) &&
        !isNaN(probability) &&
        !isNaN(alpha) &&
        trials >= 0 &&
        probability >= 0.0 &&
        probability <= 1.0 &&
        alpha > 0.0 &&
        alpha < 1.0
      ) {
        let cumulative = 0.0;
        for (let index = 0; index <= trials; index += 1) {
          cumulative += binomialProbability(index, trials, probability);
          if (cumulative >= alpha) {
            result = <f64>index;
            break;
          }
        }
        if (isNaN(result)) {
          result = <f64>trials;
        }
      }
    } else if (builtinId == BuiltinId.Hypgeomdist) {
      const sampleSuccessesRaw = toNumberExact(tagStack[base], valueStack[base]);
      const sampleSizeRaw = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
      const populationSuccessesRaw = toNumberExact(tagStack[base + 2], valueStack[base + 2]);
      const populationSizeRaw = toNumberExact(tagStack[base + 3], valueStack[base + 3]);
      if (
        !isNaN(sampleSuccessesRaw) &&
        !isNaN(sampleSizeRaw) &&
        !isNaN(populationSuccessesRaw) &&
        !isNaN(populationSizeRaw)
      ) {
        result = hypergeometricProbability(
          <i32>sampleSuccessesRaw,
          <i32>sampleSizeRaw,
          <i32>populationSuccessesRaw,
          <i32>populationSizeRaw,
        );
      }
    } else {
      const sampleSuccessesRaw = toNumberExact(tagStack[base], valueStack[base]);
      const sampleSizeRaw = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
      const populationSuccessesRaw = toNumberExact(tagStack[base + 2], valueStack[base + 2]);
      const populationSizeRaw = toNumberExact(tagStack[base + 3], valueStack[base + 3]);
      const cumulative = coerceBoolean(tagStack[base + 4], valueStack[base + 4]);
      const sampleSuccesses = <i32>sampleSuccessesRaw;
      const sampleSize = <i32>sampleSizeRaw;
      const populationSuccesses = <i32>populationSuccessesRaw;
      const populationSize = <i32>populationSizeRaw;
      if (
        !isNaN(sampleSuccessesRaw) &&
        !isNaN(sampleSizeRaw) &&
        !isNaN(populationSuccessesRaw) &&
        !isNaN(populationSizeRaw) &&
        cumulative >= 0
      ) {
        if (cumulative == 1) {
          result = 0.0;
          const minimum = max<i32>(0, sampleSize - (populationSize - populationSuccesses));
          for (let index = minimum; index <= sampleSuccesses; index += 1) {
            result += hypergeometricProbability(
              index,
              sampleSize,
              populationSuccesses,
              populationSize,
            );
          }
        } else {
          result = hypergeometricProbability(
            sampleSuccesses,
            sampleSize,
            populationSuccesses,
            populationSize,
          );
        }
      }
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      isNumericResult(result) ? <u8>ValueTag.Number : <u8>ValueTag.Error,
      isNumericResult(result) ? result : ErrorCode.Value,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  return -1;
}
