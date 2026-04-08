import { ErrorCode, ValueTag, type CellValue } from "@bilig/protocol";
import type { EvaluationResult } from "../runtime-values.js";
import type { LookupBuiltin, LookupBuiltinArgument, RangeBuiltinArgument } from "./lookup.js";

interface LookupRegressionBuiltinDeps {
  errorValue: (code: ErrorCode) => CellValue;
  numberResult: (value: number) => CellValue;
  isRangeArg: (value: LookupBuiltinArgument | undefined) => value is RangeBuiltinArgument;
  toNumber: (value: CellValue) => number | undefined;
  toBoolean: (value: CellValue) => boolean | undefined;
  flattenNumbers: (arg: LookupBuiltinArgument) => number[] | CellValue;
}

function flattenNumbersOrValueError(
  arg: LookupBuiltinArgument | undefined,
  { errorValue, flattenNumbers }: LookupRegressionBuiltinDeps,
): number[] | CellValue {
  return arg === undefined ? errorValue(ErrorCode.Value) : flattenNumbers(arg);
}

function parseCorrelationOperands(
  firstArg: LookupBuiltinArgument | undefined,
  secondArg: LookupBuiltinArgument | undefined,
  deps: LookupRegressionBuiltinDeps,
): { first: number[]; second: number[] } | CellValue {
  const first = flattenNumbersOrValueError(firstArg, deps);
  if (!Array.isArray(first)) {
    return first;
  }
  const second = flattenNumbersOrValueError(secondArg, deps);
  if (!Array.isArray(second)) {
    return second;
  }
  if (first.length === 0 || first.length !== second.length) {
    return deps.errorValue(ErrorCode.Value);
  }
  return { first, second };
}

function covarianceFromPairs(
  first: readonly number[],
  second: readonly number[],
  useSample: boolean,
  { errorValue }: LookupRegressionBuiltinDeps,
): number | CellValue {
  const count = first.length;
  const firstMean = first.reduce((sum, value) => sum + value, 0) / count;
  const secondMean = second.reduce((sum, value) => sum + value, 0) / count;

  let covarianceSum = 0;
  for (let index = 0; index < count; index += 1) {
    covarianceSum += (first[index]! - firstMean) * (second[index]! - secondMean);
  }

  const denominator = useSample ? count - 1 : count;
  if (denominator <= 0) {
    return errorValue(ErrorCode.Div0);
  }
  return covarianceSum / denominator;
}

function correlationFromPairs(
  first: readonly number[],
  second: readonly number[],
  { errorValue }: LookupRegressionBuiltinDeps,
): number | CellValue {
  if (first.length < 2) {
    return errorValue(ErrorCode.Div0);
  }
  const count = first.length;
  const firstMean = first.reduce((sum, value) => sum + value, 0) / count;
  const secondMean = second.reduce((sum, value) => sum + value, 0) / count;

  let crossProducts = 0;
  let firstVariance = 0;
  let secondVariance = 0;
  for (let index = 0; index < count; index += 1) {
    const firstDeviation = first[index]! - firstMean;
    const secondDeviation = second[index]! - secondMean;
    crossProducts += firstDeviation * secondDeviation;
    firstVariance += firstDeviation ** 2;
    secondVariance += secondDeviation ** 2;
  }
  const denominator = Math.sqrt(firstVariance * secondVariance);
  if (denominator === 0) {
    return errorValue(ErrorCode.Div0);
  }
  return crossProducts / denominator;
}

function linearRegressionFromPairs(
  knownY: readonly number[],
  knownX: readonly number[],
  deps: LookupRegressionBuiltinDeps,
  includeIntercept = true,
):
  | {
      slope: number;
      intercept: number;
      sumSquaresX: number;
      sumSquaresY: number;
      sumCrossProducts: number;
      residualSumSquares: number;
    }
  | CellValue {
  if (knownY.length !== knownX.length || knownY.length === 0) {
    return deps.errorValue(ErrorCode.Value);
  }

  const count = knownY.length;
  const meanY = knownY.reduce((sum, value) => sum + value, 0) / count;
  const meanX = knownX.reduce((sum, value) => sum + value, 0) / count;

  let sumSquaresX = 0;
  let sumSquaresY = 0;
  let sumCrossProducts = 0;
  let slope: number;
  let intercept: number;

  if (includeIntercept) {
    for (let index = 0; index < count; index += 1) {
      const xDeviation = knownX[index]! - meanX;
      const yDeviation = knownY[index]! - meanY;
      sumSquaresX += xDeviation ** 2;
      sumSquaresY += yDeviation ** 2;
      sumCrossProducts += xDeviation * yDeviation;
    }

    if (sumSquaresX === 0) {
      return deps.errorValue(ErrorCode.Div0);
    }

    slope = sumCrossProducts / sumSquaresX;
    intercept = meanY - slope * meanX;
  } else {
    for (let index = 0; index < count; index += 1) {
      const xValue = knownX[index]!;
      const yValue = knownY[index]!;
      sumSquaresX += xValue ** 2;
      sumSquaresY += (yValue - meanY) ** 2;
      sumCrossProducts += xValue * yValue;
    }
    if (sumSquaresX === 0) {
      return deps.errorValue(ErrorCode.Div0);
    }
    slope = sumCrossProducts / sumSquaresX;
    intercept = 0;
  }

  let residualSumSquares = 0;
  for (let index = 0; index < count; index += 1) {
    const residual = knownY[index]! - (intercept + slope * knownX[index]!);
    residualSumSquares += residual ** 2;
  }

  return {
    slope,
    intercept,
    sumSquaresX,
    sumSquaresY,
    sumCrossProducts,
    residualSumSquares,
  };
}

interface RegressionMatrix {
  values: number[];
  rows: number;
  cols: number;
}

function regressionMatrixFromArg(
  arg: LookupBuiltinArgument | undefined,
  { errorValue, isRangeArg, toNumber }: LookupRegressionBuiltinDeps,
): RegressionMatrix | CellValue {
  if (arg === undefined) {
    return errorValue(ErrorCode.Value);
  }
  if (!isRangeArg(arg)) {
    if (arg.tag === ValueTag.Error) {
      return arg;
    }
    const numeric = toNumber(arg);
    return numeric === undefined
      ? errorValue(ErrorCode.Value)
      : { values: [numeric], rows: 1, cols: 1 };
  }
  if (arg.refKind !== "cells") {
    return errorValue(ErrorCode.Value);
  }
  const values: number[] = [];
  for (const value of arg.values) {
    if (value.tag === ValueTag.Error) {
      return value;
    }
    const numeric = toNumber(value);
    if (numeric === undefined) {
      return errorValue(ErrorCode.Value);
    }
    values.push(numeric);
  }
  return { values, rows: arg.rows, cols: arg.cols };
}

function defaultRegressionSequence(rows: number, cols: number): RegressionMatrix {
  const values: number[] = [];
  for (let index = 0; index < rows * cols; index += 1) {
    values.push(index + 1);
  }
  return { values, rows, cols };
}

function coerceRegressionConstant(
  arg: LookupBuiltinArgument | undefined,
  { errorValue, isRangeArg, toBoolean }: LookupRegressionBuiltinDeps,
): boolean | CellValue {
  if (arg === undefined) {
    return true;
  }
  if (isRangeArg(arg)) {
    return errorValue(ErrorCode.Value);
  }
  if (arg.tag === ValueTag.Error) {
    return arg;
  }
  const value = toBoolean(arg);
  return value === undefined ? errorValue(ErrorCode.Value) : value;
}

function coerceRegressionFlag(
  arg: LookupBuiltinArgument | undefined,
  defaultValue: boolean,
  { errorValue, isRangeArg, toBoolean }: LookupRegressionBuiltinDeps,
): boolean | CellValue {
  if (arg === undefined) {
    return defaultValue;
  }
  if (isRangeArg(arg)) {
    return errorValue(ErrorCode.Value);
  }
  if (arg.tag === ValueTag.Error) {
    return arg;
  }
  const value = toBoolean(arg);
  return value === undefined ? errorValue(ErrorCode.Value) : value;
}

function predictionResultFromMatrix(
  values: readonly number[],
  rows: number,
  cols: number,
  { errorValue, numberResult }: LookupRegressionBuiltinDeps,
): EvaluationResult {
  if (values.some((value) => !Number.isFinite(value))) {
    return errorValue(ErrorCode.Value);
  }
  const resultValues = values.map((value) => numberResult(value));
  return rows === 1 && cols === 1
    ? (resultValues[0] ?? errorValue(ErrorCode.Value))
    : {
        kind: "array",
        rows,
        cols,
        values: resultValues,
      };
}

function forecastResult(
  xArg: LookupBuiltinArgument | undefined,
  knownYArg: LookupBuiltinArgument | undefined,
  knownXArg: LookupBuiltinArgument | undefined,
  deps: LookupRegressionBuiltinDeps,
): CellValue {
  if (xArg === undefined || deps.isRangeArg(xArg)) {
    return deps.errorValue(ErrorCode.Value);
  }
  if (xArg.tag === ValueTag.Error) {
    return xArg;
  }
  const x = deps.toNumber(xArg);
  if (x === undefined) {
    return deps.errorValue(ErrorCode.Value);
  }
  const values = parseCorrelationOperands(knownYArg, knownXArg, deps);
  if (!("first" in values)) {
    return values;
  }
  const regression = linearRegressionFromPairs(values.first, values.second, deps);
  return "intercept" in regression
    ? deps.numberResult(regression.intercept + regression.slope * x)
    : regression;
}

function trendLikeResult(
  mode: "trend" | "growth",
  knownYArg: LookupBuiltinArgument | undefined,
  knownXArg: LookupBuiltinArgument | undefined,
  newXArg: LookupBuiltinArgument | undefined,
  constArg: LookupBuiltinArgument | undefined,
  deps: LookupRegressionBuiltinDeps,
): EvaluationResult {
  const knownYMatrix = regressionMatrixFromArg(knownYArg, deps);
  if (!("values" in knownYMatrix)) {
    return knownYMatrix;
  }

  const knownXMatrix =
    knownXArg === undefined
      ? defaultRegressionSequence(knownYMatrix.rows, knownYMatrix.cols)
      : regressionMatrixFromArg(knownXArg, deps);
  if (!("values" in knownXMatrix)) {
    return knownXMatrix;
  }
  if (
    knownXMatrix.values.length !== knownYMatrix.values.length ||
    knownXMatrix.values.length === 0
  ) {
    return deps.errorValue(ErrorCode.Value);
  }

  const includeIntercept = coerceRegressionConstant(constArg, deps);
  if (typeof includeIntercept !== "boolean") {
    return includeIntercept;
  }

  const regressionY =
    mode === "growth"
      ? knownYMatrix.values.map((value) => {
          return value > 0 ? Math.log(value) : Number.NaN;
        })
      : [...knownYMatrix.values];
  if (regressionY.some((value) => !Number.isFinite(value))) {
    return deps.errorValue(ErrorCode.Value);
  }

  const regression = linearRegressionFromPairs(
    regressionY,
    knownXMatrix.values,
    deps,
    includeIntercept,
  );
  if (!("slope" in regression)) {
    return regression;
  }

  const newXMatrix =
    newXArg === undefined
      ? knownXArg === undefined
        ? defaultRegressionSequence(knownYMatrix.rows, knownYMatrix.cols)
        : knownXMatrix
      : regressionMatrixFromArg(newXArg, deps);
  if (!("values" in newXMatrix)) {
    return newXMatrix;
  }

  const predicted = newXMatrix.values.map((xValue) => {
    const linear = regression.intercept + regression.slope * xValue;
    return mode === "growth" ? Math.exp(linear) : linear;
  });
  return predictionResultFromMatrix(predicted, newXMatrix.rows, newXMatrix.cols, deps);
}

interface UnivariateRegressionDataset {
  knownY: number[];
  knownX: number[];
}

interface UnivariateRegressionStats {
  slope: number;
  intercept: number;
  slopeStandardError: number | undefined;
  interceptStandardError: number | undefined;
  rSquared: number | undefined;
  standardErrorY: number | undefined;
  fStatistic: number | undefined;
  degreesFreedom: number | undefined;
  regressionSumSquares: number;
  residualSumSquares: number;
}

function parseUnivariateRegressionDataset(
  mode: "linest" | "logest",
  knownYArg: LookupBuiltinArgument | undefined,
  knownXArg: LookupBuiltinArgument | undefined,
  deps: LookupRegressionBuiltinDeps,
): UnivariateRegressionDataset | CellValue {
  const knownYMatrix = regressionMatrixFromArg(knownYArg, deps);
  if (!("values" in knownYMatrix)) {
    return knownYMatrix;
  }
  if (knownYMatrix.values.length === 0) {
    return deps.errorValue(ErrorCode.Value);
  }

  const knownY =
    mode === "logest"
      ? knownYMatrix.values.map((value) => (value > 0 ? Math.log(value) : Number.NaN))
      : [...knownYMatrix.values];
  if (knownY.some((value) => !Number.isFinite(value))) {
    return deps.errorValue(ErrorCode.Value);
  }

  if (knownXArg === undefined) {
    return {
      knownY,
      knownX: knownY.map((_, index) => index + 1),
    };
  }

  const knownXMatrix = regressionMatrixFromArg(knownXArg, deps);
  if (!("values" in knownXMatrix)) {
    return knownXMatrix;
  }
  if (knownXMatrix.values.length !== knownY.length || knownXMatrix.values.length === 0) {
    return deps.errorValue(ErrorCode.Value);
  }
  if (knownXMatrix.values.some((value) => !Number.isFinite(value))) {
    return deps.errorValue(ErrorCode.Value);
  }

  return {
    knownY,
    knownX: [...knownXMatrix.values],
  };
}

function analyzeUnivariateRegression(
  knownY: readonly number[],
  knownX: readonly number[],
  includeIntercept: boolean,
  deps: LookupRegressionBuiltinDeps,
): UnivariateRegressionStats | CellValue {
  const regression = linearRegressionFromPairs(knownY, knownX, deps, includeIntercept);
  if (!("slope" in regression)) {
    return regression;
  }

  const count = knownY.length;
  const meanY = knownY.reduce((sum, value) => sum + value, 0) / count;
  const totalSumSquares = includeIntercept
    ? knownY.reduce((sum, value) => sum + (value - meanY) ** 2, 0)
    : knownY.reduce((sum, value) => sum + value ** 2, 0);
  const residualSumSquares = Math.max(0, regression.residualSumSquares);
  const regressionSumSquares = Math.max(0, totalSumSquares - residualSumSquares);
  const parameterCount = includeIntercept ? 2 : 1;
  const degreesFreedom = count - parameterCount;

  let slopeStandardError: number | undefined;
  let interceptStandardError: number | undefined;
  let standardErrorY: number | undefined;
  let fStatistic: number | undefined;

  if (degreesFreedom > 0) {
    const meanSquaredError = residualSumSquares / degreesFreedom;
    standardErrorY = Math.sqrt(meanSquaredError);
    if (includeIntercept) {
      const meanX = knownX.reduce((sum, value) => sum + value, 0) / count;
      const sumSquaresX = knownX.reduce((sum, value) => {
        const deviation = value - meanX;
        return sum + deviation * deviation;
      }, 0);
      if (sumSquaresX > 0) {
        slopeStandardError = Math.sqrt(meanSquaredError / sumSquaresX);
        interceptStandardError = Math.sqrt(
          meanSquaredError * (1 / count + (meanX * meanX) / sumSquaresX),
        );
      }
    } else {
      const sumSquaresX = knownX.reduce((sum, value) => sum + value * value, 0);
      if (sumSquaresX > 0) {
        slopeStandardError = Math.sqrt(meanSquaredError / sumSquaresX);
        interceptStandardError = 0;
      }
    }
    if (residualSumSquares === 0) {
      fStatistic = Number.POSITIVE_INFINITY;
    } else {
      fStatistic = regressionSumSquares / (residualSumSquares / degreesFreedom);
    }
  }

  let rSquared: number | undefined;
  if (totalSumSquares === 0) {
    rSquared = residualSumSquares === 0 ? 1 : undefined;
  } else {
    rSquared = 1 - residualSumSquares / totalSumSquares;
  }

  return {
    slope: regression.slope,
    intercept: regression.intercept,
    slopeStandardError,
    interceptStandardError,
    rSquared,
    standardErrorY,
    fStatistic,
    degreesFreedom: degreesFreedom > 0 ? degreesFreedom : undefined,
    regressionSumSquares,
    residualSumSquares,
  };
}

function regressionStatCell(
  value: number | undefined,
  { errorValue, numberResult }: LookupRegressionBuiltinDeps,
): CellValue {
  if (value === undefined || Number.isNaN(value)) {
    return errorValue(ErrorCode.Div0);
  }
  return Number.isFinite(value) ? numberResult(value) : errorValue(ErrorCode.Div0);
}

function linearEstimationResult(
  mode: "linest" | "logest",
  knownYArg: LookupBuiltinArgument | undefined,
  knownXArg: LookupBuiltinArgument | undefined,
  constArg: LookupBuiltinArgument | undefined,
  statsArg: LookupBuiltinArgument | undefined,
  deps: LookupRegressionBuiltinDeps,
): EvaluationResult {
  const dataset = parseUnivariateRegressionDataset(mode, knownYArg, knownXArg, deps);
  if (!("knownY" in dataset)) {
    return dataset;
  }

  const includeIntercept = coerceRegressionFlag(constArg, true, deps);
  if (typeof includeIntercept !== "boolean") {
    return includeIntercept;
  }
  const includeStats = coerceRegressionFlag(statsArg, false, deps);
  if (typeof includeStats !== "boolean") {
    return includeStats;
  }

  const regression = analyzeUnivariateRegression(
    dataset.knownY,
    dataset.knownX,
    includeIntercept,
    deps,
  );
  if (!("slope" in regression)) {
    return regression;
  }

  const leading = mode === "logest" ? Math.exp(regression.slope) : regression.slope;
  const trailing =
    mode === "logest"
      ? includeIntercept
        ? Math.exp(regression.intercept)
        : 1
      : includeIntercept
        ? regression.intercept
        : 0;

  if (!includeStats) {
    return {
      kind: "array",
      rows: 1,
      cols: 2,
      values: [deps.numberResult(leading), deps.numberResult(trailing)],
    };
  }

  return {
    kind: "array",
    rows: 5,
    cols: 2,
    values: [
      deps.numberResult(leading),
      deps.numberResult(trailing),
      regressionStatCell(regression.slopeStandardError, deps),
      regressionStatCell(regression.interceptStandardError, deps),
      regressionStatCell(regression.rSquared, deps),
      regressionStatCell(regression.standardErrorY, deps),
      regressionStatCell(regression.fStatistic, deps),
      regressionStatCell(regression.degreesFreedom, deps),
      deps.numberResult(regression.regressionSumSquares),
      deps.numberResult(regression.residualSumSquares),
    ],
  };
}

export function createLookupRegressionBuiltins(
  deps: LookupRegressionBuiltinDeps,
): Record<string, LookupBuiltin> {
  return {
    CORREL: (firstArg, secondArg) => {
      const values = parseCorrelationOperands(firstArg, secondArg, deps);
      if (!("first" in values)) {
        return values;
      }
      const correlation = correlationFromPairs(values.first, values.second, deps);
      return typeof correlation === "number" ? deps.numberResult(correlation) : correlation;
    },
    COVAR: (firstArg, secondArg) => {
      const values = parseCorrelationOperands(firstArg, secondArg, deps);
      if (!("first" in values)) {
        return values;
      }
      const covariance = covarianceFromPairs(values.first, values.second, false, deps);
      return typeof covariance === "number" ? deps.numberResult(covariance) : covariance;
    },
    PEARSON: (firstArg, secondArg) => {
      const values = parseCorrelationOperands(firstArg, secondArg, deps);
      if (!("first" in values)) {
        return values;
      }
      const correlation = correlationFromPairs(values.first, values.second, deps);
      return typeof correlation === "number" ? deps.numberResult(correlation) : correlation;
    },
    "COVARIANCE.P": (firstArg, secondArg) => {
      const values = parseCorrelationOperands(firstArg, secondArg, deps);
      if (!("first" in values)) {
        return values;
      }
      const covariance = covarianceFromPairs(values.first, values.second, false, deps);
      return typeof covariance === "number" ? deps.numberResult(covariance) : covariance;
    },
    "COVARIANCE.S": (firstArg, secondArg) => {
      const values = parseCorrelationOperands(firstArg, secondArg, deps);
      if (!("first" in values)) {
        return values;
      }
      const covariance = covarianceFromPairs(values.first, values.second, true, deps);
      return typeof covariance === "number" ? deps.numberResult(covariance) : covariance;
    },
    INTERCEPT: (knownYArg, knownXArg) => {
      const values = parseCorrelationOperands(knownYArg, knownXArg, deps);
      if (!("first" in values)) {
        return values;
      }
      const regression = linearRegressionFromPairs(values.first, values.second, deps);
      return "intercept" in regression ? deps.numberResult(regression.intercept) : regression;
    },
    SLOPE: (knownYArg, knownXArg) => {
      const values = parseCorrelationOperands(knownYArg, knownXArg, deps);
      if (!("first" in values)) {
        return values;
      }
      const regression = linearRegressionFromPairs(values.first, values.second, deps);
      return "slope" in regression ? deps.numberResult(regression.slope) : regression;
    },
    RSQ: (knownYArg, knownXArg) => {
      const values = parseCorrelationOperands(knownYArg, knownXArg, deps);
      if (!("first" in values)) {
        return values;
      }
      const correlation = correlationFromPairs(values.first, values.second, deps);
      return typeof correlation === "number"
        ? deps.numberResult(correlation * correlation)
        : correlation;
    },
    STEYX: (knownYArg, knownXArg) => {
      const values = parseCorrelationOperands(knownYArg, knownXArg, deps);
      if (!("first" in values)) {
        return values;
      }
      if (values.first.length <= 2) {
        return deps.errorValue(ErrorCode.Div0);
      }
      const regression = linearRegressionFromPairs(values.first, values.second, deps);
      if (!("residualSumSquares" in regression)) {
        return regression;
      }
      return deps.numberResult(
        Math.sqrt(Math.max(0, regression.residualSumSquares) / (values.first.length - 2)),
      );
    },
    FORECAST: (xArg, knownYArg, knownXArg) => {
      return forecastResult(xArg, knownYArg, knownXArg, deps);
    },
    "FORECAST.LINEAR": (xArg, knownYArg, knownXArg) => {
      return forecastResult(xArg, knownYArg, knownXArg, deps);
    },
    TREND: (knownYArg, knownXArg, newXArg, constArg) => {
      return trendLikeResult("trend", knownYArg, knownXArg, newXArg, constArg, deps);
    },
    GROWTH: (knownYArg, knownXArg, newXArg, constArg) => {
      return trendLikeResult("growth", knownYArg, knownXArg, newXArg, constArg, deps);
    },
    LINEST: (knownYArg, knownXArg, constArg, statsArg) => {
      return linearEstimationResult("linest", knownYArg, knownXArg, constArg, statsArg, deps);
    },
    LOGEST: (knownYArg, knownXArg, constArg, statsArg) => {
      return linearEstimationResult("logest", knownYArg, knownXArg, constArg, statsArg, deps);
    },
  };
}
