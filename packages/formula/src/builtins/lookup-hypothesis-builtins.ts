import { ErrorCode, ValueTag, type CellValue } from "@bilig/protocol";
import type { LookupBuiltin, LookupBuiltinArgument, RangeBuiltinArgument } from "./lookup.js";

interface LookupHypothesisBuiltinDeps {
  errorValue: (code: ErrorCode) => CellValue;
  isRangeArg: (value: LookupBuiltinArgument | undefined) => value is RangeBuiltinArgument;
  toNumber: (value: CellValue) => number | undefined;
  toNumericMatrix: (arg: LookupBuiltinArgument) => number[][] | CellValue;
}

const LANCZOS_G = 7;
const LANCZOS_COEFFICIENTS = [
  676.5203681218851, -1259.1392167224028, 771.3234287776531, -176.6150291621406, 12.507343278686905,
  -0.13857109526572012, 9.984369578019572e-6, 1.5056327351493116e-7,
] as const;

function logGamma(value: number): number {
  if (!Number.isFinite(value) || value <= 0) {
    return Number.NaN;
  }
  let sum = 0.9999999999998099;
  const shifted = value - 1;
  LANCZOS_COEFFICIENTS.forEach((coefficient, index) => {
    sum += coefficient / (shifted + index + 1);
  });
  const t = shifted + LANCZOS_G + 0.5;
  return 0.5 * Math.log(2 * Math.PI) + (shifted + 0.5) * Math.log(t) - t + Math.log(sum);
}

function regularizedLowerGamma(shape: number, x: number): number {
  if (!Number.isFinite(shape) || !Number.isFinite(x) || shape <= 0 || x < 0) {
    return Number.NaN;
  }
  if (x === 0) {
    return 0;
  }
  const logGammaShape = logGamma(shape);
  if (!Number.isFinite(logGammaShape)) {
    return Number.NaN;
  }
  if (x < shape + 1) {
    let term = 1 / shape;
    let sum = term;
    for (let iteration = 1; iteration < 1000; iteration += 1) {
      term *= x / (shape + iteration);
      sum += term;
      if (Math.abs(term) <= Math.abs(sum) * 1e-14) {
        break;
      }
    }
    return sum * Math.exp(-x + shape * Math.log(x) - logGammaShape);
  }

  let b = x + 1 - shape;
  let c = 1 / 1e-300;
  let d = 1 / b;
  let h = d;
  for (let iteration = 1; iteration < 1000; iteration += 1) {
    const factor = -iteration * (iteration - shape);
    b += 2;
    d = factor * d + b;
    if (Math.abs(d) < 1e-300) {
      d = 1e-300;
    }
    c = b + factor / c;
    if (Math.abs(c) < 1e-300) {
      c = 1e-300;
    }
    d = 1 / d;
    const delta = d * c;
    h *= delta;
    if (Math.abs(delta - 1) <= 1e-14) {
      break;
    }
  }
  return 1 - Math.exp(-x + shape * Math.log(x) - logGammaShape) * h;
}

function regularizedUpperGamma(shape: number, x: number): number {
  const lower = regularizedLowerGamma(shape, x);
  return Number.isFinite(lower) ? 1 - lower : Number.NaN;
}

function erfApprox(value: number): number {
  const sign = value < 0 ? -1 : 1;
  const absolute = Math.abs(value);
  const t = 1 / (1 + 0.3275911 * absolute);
  const y =
    1 -
    ((((1.061405429 * t - 1.453152027) * t + 1.421413741) * t - 0.284496736) * t + 0.254829592) *
      t *
      Math.exp(-absolute * absolute);
  return sign * y;
}

function standardNormalCdf(value: number): number {
  return 0.5 * (1 + erfApprox(value / Math.sqrt(2)));
}

function logBeta(alpha: number, beta: number): number {
  return logGamma(alpha) + logGamma(beta) - logGamma(alpha + beta);
}

function betaContinuedFraction(x: number, alpha: number, beta: number): number {
  const maxIterations = 200;
  const epsilon = 1e-14;
  const tiny = 1e-300;
  const qab = alpha + beta;
  const qap = alpha + 1;
  const qam = alpha - 1;
  let c = 1;
  let d = 1 - (qab * x) / qap;
  if (Math.abs(d) < tiny) {
    d = tiny;
  }
  d = 1 / d;
  let h = d;
  for (let iteration = 1; iteration <= maxIterations; iteration += 1) {
    const step = iteration * 2;
    let factor = (iteration * (beta - iteration) * x) / ((qam + step) * (alpha + step));
    d = 1 + factor * d;
    if (Math.abs(d) < tiny) {
      d = tiny;
    }
    c = 1 + factor / c;
    if (Math.abs(c) < tiny) {
      c = tiny;
    }
    d = 1 / d;
    h *= d * c;

    factor = (-(alpha + iteration) * (qab + iteration) * x) / ((alpha + step) * (qap + step));
    d = 1 + factor * d;
    if (Math.abs(d) < tiny) {
      d = tiny;
    }
    c = 1 + factor / c;
    if (Math.abs(c) < tiny) {
      c = tiny;
    }
    d = 1 / d;
    const delta = d * c;
    h *= delta;
    if (Math.abs(delta - 1) <= epsilon) {
      break;
    }
  }
  return h;
}

function regularizedBeta(x: number, alpha: number, beta: number): number {
  if (
    !Number.isFinite(x) ||
    !Number.isFinite(alpha) ||
    !Number.isFinite(beta) ||
    x < 0 ||
    x > 1 ||
    alpha <= 0 ||
    beta <= 0
  ) {
    return Number.NaN;
  }
  if (x === 0) {
    return 0;
  }
  if (x === 1) {
    return 1;
  }

  const logTerm = alpha * Math.log(x) + beta * Math.log(1 - x) - logBeta(alpha, beta);
  if (!Number.isFinite(logTerm)) {
    return Number.NaN;
  }
  const front = Math.exp(logTerm);
  if (x < (alpha + 1) / (alpha + beta + 2)) {
    return (front * betaContinuedFraction(x, alpha, beta)) / alpha;
  }
  return 1 - (front * betaContinuedFraction(1 - x, beta, alpha)) / beta;
}

function fDistributionCdf(x: number, degreesFreedom1: number, degreesFreedom2: number): number {
  if (
    !Number.isFinite(x) ||
    !Number.isFinite(degreesFreedom1) ||
    !Number.isFinite(degreesFreedom2) ||
    x < 0 ||
    degreesFreedom1 < 1 ||
    degreesFreedom2 < 1
  ) {
    return Number.NaN;
  }
  const alpha = degreesFreedom1 / 2;
  const beta = degreesFreedom2 / 2;
  const transformed = (degreesFreedom1 * x) / (degreesFreedom1 * x + degreesFreedom2);
  return regularizedBeta(transformed, alpha, beta);
}

function studentTCdf(x: number, degreesFreedom: number): number {
  if (!Number.isFinite(x) || !Number.isFinite(degreesFreedom) || degreesFreedom < 1) {
    return Number.NaN;
  }
  if (x === 0) {
    return 0.5;
  }
  const transformed = degreesFreedom / (degreesFreedom + x * x);
  const tail = regularizedBeta(transformed, degreesFreedom / 2, 0.5);
  if (!Number.isFinite(tail)) {
    return Number.NaN;
  }
  return x > 0 ? 1 - tail / 2 : tail / 2;
}

function collectSampleNumbers(
  arg: LookupBuiltinArgument,
  deps: LookupHypothesisBuiltinDeps,
): number[] | CellValue {
  if (!deps.isRangeArg(arg)) {
    if (arg.tag === ValueTag.Error) {
      return arg;
    }
    return arg.tag === ValueTag.Number ? [arg.value] : deps.errorValue(ErrorCode.Value);
  }

  const values: number[] = [];
  for (const value of arg.values) {
    if (value.tag === ValueTag.Error) {
      return value;
    }
    if (value.tag === ValueTag.Number) {
      values.push(value.value);
    }
  }
  return values;
}

function chiSquareTestResult(
  actualArg: LookupBuiltinArgument,
  expectedArg: LookupBuiltinArgument,
  deps: LookupHypothesisBuiltinDeps,
): CellValue {
  const actual = deps.toNumericMatrix(actualArg);
  if (!Array.isArray(actual)) {
    return actual;
  }
  const expected = deps.toNumericMatrix(expectedArg);
  if (!Array.isArray(expected)) {
    return expected;
  }

  const rows = actual.length;
  const cols = actual[0]?.length ?? 0;
  if (rows !== expected.length || cols !== (expected[0]?.length ?? 0)) {
    return deps.errorValue(ErrorCode.NA);
  }
  if ((rows === 1 && cols === 1) || rows === 0 || cols === 0) {
    return deps.errorValue(ErrorCode.NA);
  }

  let statistic = 0;
  for (let row = 0; row < rows; row += 1) {
    const actualRow = actual[row]!;
    const expectedRow = expected[row]!;
    if (actualRow.length !== cols || expectedRow.length !== cols) {
      return deps.errorValue(ErrorCode.NA);
    }
    for (let col = 0; col < cols; col += 1) {
      const actualValue = actualRow[col]!;
      const expectedValue = expectedRow[col]!;
      if (actualValue < 0 || expectedValue < 0) {
        return deps.errorValue(ErrorCode.Value);
      }
      if (expectedValue === 0) {
        return deps.errorValue(ErrorCode.Div0);
      }
      const delta = actualValue - expectedValue;
      statistic += (delta * delta) / expectedValue;
    }
  }

  const degrees = rows > 1 && cols > 1 ? (rows - 1) * (cols - 1) : rows > 1 ? rows - 1 : cols - 1;
  if (degrees <= 0) {
    return deps.errorValue(ErrorCode.NA);
  }
  const probability = regularizedUpperGamma(degrees / 2, statistic / 2);
  return Number.isFinite(probability)
    ? { tag: ValueTag.Number, value: probability }
    : deps.errorValue(ErrorCode.Value);
}

function fTestResult(
  firstArg: LookupBuiltinArgument,
  secondArg: LookupBuiltinArgument,
  deps: LookupHypothesisBuiltinDeps,
): CellValue {
  const first = collectSampleNumbers(firstArg, deps);
  if (!Array.isArray(first)) {
    return first;
  }
  const second = collectSampleNumbers(secondArg, deps);
  if (!Array.isArray(second)) {
    return second;
  }
  if (first.length < 2 || second.length < 2) {
    return deps.errorValue(ErrorCode.Div0);
  }

  const firstMean = first.reduce((sum, value) => sum + value, 0) / first.length;
  const secondMean = second.reduce((sum, value) => sum + value, 0) / second.length;
  const firstVariance =
    first.reduce((sum, value) => sum + (value - firstMean) ** 2, 0) / (first.length - 1);
  const secondVariance =
    second.reduce((sum, value) => sum + (value - secondMean) ** 2, 0) / (second.length - 1);
  if (!(firstVariance > 0) || !(secondVariance > 0)) {
    return deps.errorValue(ErrorCode.Div0);
  }

  const firstLeads = firstVariance >= secondVariance;
  const numeratorVariance = firstLeads ? firstVariance : secondVariance;
  const denominatorVariance = firstLeads ? secondVariance : firstVariance;
  const numeratorDf = firstLeads ? first.length - 1 : second.length - 1;
  const denominatorDf = firstLeads ? second.length - 1 : first.length - 1;
  const upperTail =
    1 - fDistributionCdf(numeratorVariance / denominatorVariance, numeratorDf, denominatorDf);
  const probability = Math.min(1, upperTail * 2);
  return Number.isFinite(probability)
    ? { tag: ValueTag.Number, value: probability }
    : deps.errorValue(ErrorCode.Value);
}

function zTestResult(
  arrayArg: LookupBuiltinArgument,
  xArg: LookupBuiltinArgument,
  deps: LookupHypothesisBuiltinDeps,
  sigmaArg?: LookupBuiltinArgument,
): CellValue {
  const sample = collectSampleNumbers(arrayArg, deps);
  if (!Array.isArray(sample)) {
    return sample;
  }
  const x = !deps.isRangeArg(xArg) ? deps.toNumber(xArg) : undefined;
  if (x === undefined || sample.length === 0) {
    return deps.errorValue(ErrorCode.Value);
  }

  let sigma: number | undefined;
  if (sigmaArg !== undefined) {
    sigma = !deps.isRangeArg(sigmaArg) ? deps.toNumber(sigmaArg) : undefined;
  } else if (sample.length >= 2) {
    const mean = sample.reduce((sum, value) => sum + value, 0) / sample.length;
    const variance =
      sample.reduce((sum, value) => sum + (value - mean) ** 2, 0) / (sample.length - 1);
    sigma = variance > 0 ? Math.sqrt(variance) : undefined;
  }

  if (sigma === undefined || !(sigma > 0)) {
    return deps.errorValue(ErrorCode.Div0);
  }
  const mean = sample.reduce((sum, value) => sum + value, 0) / sample.length;
  const zScore = (mean - x) / (sigma / Math.sqrt(sample.length));
  const probability = 1 - standardNormalCdf(zScore);
  return Number.isFinite(probability)
    ? { tag: ValueTag.Number, value: probability }
    : deps.errorValue(ErrorCode.Value);
}

function tTestResult(
  firstArg: LookupBuiltinArgument,
  secondArg: LookupBuiltinArgument,
  tailsArg: LookupBuiltinArgument,
  typeArg: LookupBuiltinArgument,
  deps: LookupHypothesisBuiltinDeps,
): CellValue {
  const first = collectSampleNumbers(firstArg, deps);
  if (!Array.isArray(first)) {
    return first;
  }
  const second = collectSampleNumbers(secondArg, deps);
  if (!Array.isArray(second)) {
    return second;
  }
  const tails = !deps.isRangeArg(tailsArg) ? deps.toNumber(tailsArg) : undefined;
  const type = !deps.isRangeArg(typeArg) ? deps.toNumber(typeArg) : undefined;
  if (
    tails === undefined ||
    type === undefined ||
    !Number.isInteger(tails) ||
    !Number.isInteger(type) ||
    ![1, 2].includes(tails) ||
    ![1, 2, 3].includes(type)
  ) {
    return deps.errorValue(ErrorCode.Value);
  }

  let statistic: number;
  let degreesFreedom: number;
  if (type === 1) {
    if (first.length !== second.length) {
      return deps.errorValue(ErrorCode.NA);
    }
    if (first.length < 2) {
      return deps.errorValue(ErrorCode.Div0);
    }
    const deltas = first.map((value, index) => value - second[index]!);
    const mean = deltas.reduce((sum, value) => sum + value, 0) / deltas.length;
    const variance =
      deltas.reduce((sum, value) => sum + (value - mean) ** 2, 0) / (deltas.length - 1);
    if (!(variance > 0)) {
      return deps.errorValue(ErrorCode.Div0);
    }
    statistic = mean / Math.sqrt(variance / deltas.length);
    degreesFreedom = deltas.length - 1;
  } else {
    if (first.length < 2 || second.length < 2) {
      return deps.errorValue(ErrorCode.Div0);
    }
    const firstMean = first.reduce((sum, value) => sum + value, 0) / first.length;
    const secondMean = second.reduce((sum, value) => sum + value, 0) / second.length;
    const firstVariance =
      first.reduce((sum, value) => sum + (value - firstMean) ** 2, 0) / (first.length - 1);
    const secondVariance =
      second.reduce((sum, value) => sum + (value - secondMean) ** 2, 0) / (second.length - 1);
    if (!(firstVariance > 0) || !(secondVariance > 0)) {
      return deps.errorValue(ErrorCode.Div0);
    }

    if (type === 2) {
      const pooledVariance =
        ((first.length - 1) * firstVariance + (second.length - 1) * secondVariance) /
        (first.length + second.length - 2);
      if (!(pooledVariance > 0)) {
        return deps.errorValue(ErrorCode.Div0);
      }
      statistic =
        (firstMean - secondMean) /
        Math.sqrt(pooledVariance * (1 / first.length + 1 / second.length));
      degreesFreedom = first.length + second.length - 2;
    } else {
      const firstTerm = firstVariance / first.length;
      const secondTerm = secondVariance / second.length;
      const denominator = Math.sqrt(firstTerm + secondTerm);
      const welchDenominator =
        (firstTerm * firstTerm) / (first.length - 1) +
        (secondTerm * secondTerm) / (second.length - 1);
      if (!(denominator > 0) || !(welchDenominator > 0)) {
        return deps.errorValue(ErrorCode.Div0);
      }
      statistic = (firstMean - secondMean) / denominator;
      degreesFreedom = (firstTerm + secondTerm) ** 2 / welchDenominator;
    }
  }

  const upperTail = 1 - studentTCdf(Math.abs(statistic), degreesFreedom);
  const probability = tails === 1 ? upperTail : Math.min(1, upperTail * 2);
  return Number.isFinite(probability)
    ? { tag: ValueTag.Number, value: probability }
    : deps.errorValue(ErrorCode.Value);
}

export function createLookupHypothesisBuiltins(
  deps: LookupHypothesisBuiltinDeps,
): Record<string, LookupBuiltin> {
  return {
    "CHISQ.TEST": (actualArg, expectedArg) => {
      return actualArg === undefined || expectedArg === undefined
        ? deps.errorValue(ErrorCode.Value)
        : chiSquareTestResult(actualArg, expectedArg, deps);
    },
    CHITEST: (actualArg, expectedArg) => {
      return actualArg === undefined || expectedArg === undefined
        ? deps.errorValue(ErrorCode.Value)
        : chiSquareTestResult(actualArg, expectedArg, deps);
    },
    "LEGACY.CHITEST": (actualArg, expectedArg) => {
      return actualArg === undefined || expectedArg === undefined
        ? deps.errorValue(ErrorCode.Value)
        : chiSquareTestResult(actualArg, expectedArg, deps);
    },
    "F.TEST": (firstArg, secondArg) => {
      return firstArg === undefined || secondArg === undefined
        ? deps.errorValue(ErrorCode.Value)
        : fTestResult(firstArg, secondArg, deps);
    },
    FTEST: (firstArg, secondArg) => {
      return firstArg === undefined || secondArg === undefined
        ? deps.errorValue(ErrorCode.Value)
        : fTestResult(firstArg, secondArg, deps);
    },
    "Z.TEST": (arrayArg, xArg, sigmaArg) => {
      return arrayArg === undefined || xArg === undefined
        ? deps.errorValue(ErrorCode.Value)
        : zTestResult(arrayArg, xArg, deps, sigmaArg);
    },
    ZTEST: (arrayArg, xArg, sigmaArg) => {
      return arrayArg === undefined || xArg === undefined
        ? deps.errorValue(ErrorCode.Value)
        : zTestResult(arrayArg, xArg, deps, sigmaArg);
    },
    "T.TEST": (firstArg, secondArg, tailsArg, typeArg) => {
      return firstArg === undefined ||
        secondArg === undefined ||
        tailsArg === undefined ||
        typeArg === undefined
        ? deps.errorValue(ErrorCode.Value)
        : tTestResult(firstArg, secondArg, tailsArg, typeArg, deps);
    },
    TTEST: (firstArg, secondArg, tailsArg, typeArg) => {
      return firstArg === undefined ||
        secondArg === undefined ||
        tailsArg === undefined ||
        typeArg === undefined
        ? deps.errorValue(ErrorCode.Value)
        : tTestResult(firstArg, secondArg, tailsArg, typeArg, deps);
    },
  };
}
