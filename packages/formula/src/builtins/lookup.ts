import { ErrorCode, ValueTag, type CellValue } from "@bilig/protocol";
import { getExternalLookupFunction } from "../external-function-adapter.js";
import type { ArrayValue, EvaluationResult } from "../runtime-values.js";
import { createLookupDatabaseBuiltins } from "./lookup-database-builtins.js";
import { createLookupFinancialBuiltins } from "./lookup-financial-builtins.js";
import { createLookupOrderStatisticsBuiltins } from "./lookup-order-statistics-builtins.js";
import { createLookupRegressionBuiltins } from "./lookup-regression-builtins.js";

export interface RangeBuiltinArgument {
  kind: "range";
  values: CellValue[];
  refKind: "cells" | "rows" | "cols";
  rows: number;
  cols: number;
}

export type LookupBuiltinArgument = CellValue | RangeBuiltinArgument;
export type LookupBuiltin = (...args: LookupBuiltinArgument[]) => EvaluationResult;

function errorValue(code: ErrorCode): CellValue {
  return { tag: ValueTag.Error, code };
}

function numberResult(value: number): CellValue {
  return { tag: ValueTag.Number, value };
}

function isError(
  value: LookupBuiltinArgument | undefined,
): value is Extract<CellValue, { tag: ValueTag.Error }> {
  return value !== undefined && !isRangeArg(value) && value.tag === ValueTag.Error;
}

function isRangeArg(value: LookupBuiltinArgument | undefined): value is RangeBuiltinArgument {
  return typeof value === "object" && value !== null && "kind" in value && value.kind === "range";
}

function isCriteriaOperator(value: string): value is CriteriaOperator {
  return (
    value === "=" ||
    value === "<>" ||
    value === ">" ||
    value === ">=" ||
    value === "<" ||
    value === "<="
  );
}

function findFirstNonRange(
  values: readonly (RangeBuiltinArgument | CellValue)[],
): CellValue | undefined {
  for (const value of values) {
    if (!isRangeArg(value)) {
      return value;
    }
  }
  return undefined;
}

function areRangeArgs(
  values: readonly (RangeBuiltinArgument | CellValue)[],
): values is RangeBuiltinArgument[] {
  return values.every((value) => isRangeArg(value));
}

function toNumber(value: CellValue): number | undefined {
  switch (value.tag) {
    case ValueTag.Number:
      return value.value;
    case ValueTag.Boolean:
      return value.value ? 1 : 0;
    case ValueTag.Empty:
      return 0;
    case ValueTag.String:
    case ValueTag.Error:
      return undefined;
    default:
      return undefined;
  }
}

function toInteger(value: CellValue): number | undefined {
  const numeric = toNumber(value);
  if (numeric === undefined || !Number.isFinite(numeric)) {
    return undefined;
  }
  return Math.trunc(numeric);
}

function toBoolean(value: CellValue): boolean | undefined {
  switch (value.tag) {
    case ValueTag.Boolean:
      return value.value;
    case ValueTag.Number:
      return value.value !== 0;
    case ValueTag.Empty:
      return false;
    case ValueTag.String:
    case ValueTag.Error:
      return undefined;
    default:
      return undefined;
  }
}

function toStringValue(value: CellValue): string {
  switch (value.tag) {
    case ValueTag.Empty:
      return "";
    case ValueTag.Number:
      return String(value.value);
    case ValueTag.Boolean:
      return value.value ? "TRUE" : "FALSE";
    case ValueTag.String:
      return value.value;
    case ValueTag.Error:
      return "";
  }
}

function compareScalars(left: CellValue, right: CellValue): number | undefined {
  if (
    (left.tag === ValueTag.String || left.tag === ValueTag.Empty) &&
    (right.tag === ValueTag.String || right.tag === ValueTag.Empty)
  ) {
    const normalizedLeft = toStringValue(left).toUpperCase();
    const normalizedRight = toStringValue(right).toUpperCase();
    if (normalizedLeft === normalizedRight) {
      return 0;
    }
    return normalizedLeft < normalizedRight ? -1 : 1;
  }

  const leftNum = toNumber(left);
  const rightNum = toNumber(right);
  if (leftNum === undefined || rightNum === undefined) {
    return undefined;
  }
  if (leftNum === rightNum) {
    return 0;
  }
  return leftNum < rightNum ? -1 : 1;
}

function requireCellVector(arg: LookupBuiltinArgument): RangeBuiltinArgument | CellValue {
  if (!isRangeArg(arg)) {
    return errorValue(ErrorCode.Value);
  }
  if (arg.refKind !== "cells") {
    return errorValue(ErrorCode.Value);
  }
  if (arg.rows !== 1 && arg.cols !== 1) {
    return errorValue(ErrorCode.NA);
  }
  return arg;
}

function requireCellRange(arg: LookupBuiltinArgument): RangeBuiltinArgument | CellValue {
  if (!isRangeArg(arg) || arg.refKind !== "cells") {
    return errorValue(ErrorCode.Value);
  }
  return arg;
}

function getRangeValue(range: RangeBuiltinArgument, row: number, col: number): CellValue {
  const index = row * range.cols + col;
  return range.values[index] ?? { tag: ValueTag.Empty };
}

function validateCriteriaPairs(
  criteriaArgs: readonly LookupBuiltinArgument[],
): { range: RangeBuiltinArgument; criteria: CellValue }[] | CellValue {
  if (criteriaArgs.length === 0 || criteriaArgs.length % 2 !== 0) {
    return errorValue(ErrorCode.Value);
  }
  const rangeCriteriaPairs: { range: RangeBuiltinArgument; criteria: CellValue }[] = [];
  for (let index = 0; index < criteriaArgs.length; index += 2) {
    const range = requireCellRange(criteriaArgs[index]!);
    if (!isRangeArg(range)) {
      return range;
    }
    const criteria = criteriaArgs[index + 1]!;
    if (isRangeArg(criteria)) {
      return errorValue(ErrorCode.Value);
    }
    if (isError(criteria)) {
      return criteria;
    }
    rangeCriteriaPairs.push({ range, criteria });
  }
  return rangeCriteriaPairs;
}

function findMatchingRowIndexes(
  targetRange: RangeBuiltinArgument,
  criteriaArgs: readonly LookupBuiltinArgument[],
): number[] | CellValue {
  const rangeCriteriaPairs = validateCriteriaPairs(criteriaArgs);
  if (!Array.isArray(rangeCriteriaPairs)) {
    return rangeCriteriaPairs;
  }
  if (rangeCriteriaPairs.some((pair) => pair.range.values.length !== targetRange.values.length)) {
    return errorValue(ErrorCode.Value);
  }

  const matchingRows: number[] = [];
  for (let row = 0; row < targetRange.values.length; row += 1) {
    if (
      rangeCriteriaPairs.every((pair) => matchesCriteria(pair.range.values[row]!, pair.criteria))
    ) {
      matchingRows.push(row);
    }
  }
  return matchingRows;
}

function arrayResult(values: CellValue[], rows: number, cols: number): ArrayValue {
  return { kind: "array", values, rows, cols };
}

function collectNumericSeries(
  arg: LookupBuiltinArgument,
  mode: "lenient" | "strict",
): number[] | CellValue {
  const values: number[] = [];
  const cells = isRangeArg(arg) ? arg.values : [arg];
  if (isRangeArg(arg) && arg.refKind !== "cells") {
    return errorValue(ErrorCode.Value);
  }
  for (const cell of cells) {
    if (cell.tag === ValueTag.Error) {
      return cell;
    }
    if (cell.tag === ValueTag.Number) {
      values.push(cell.value);
      continue;
    }
    if (mode === "strict") {
      return errorValue(ErrorCode.Value);
    }
  }
  return values;
}

function numericAggregateCandidate(value: CellValue): number | undefined {
  return value.tag === ValueTag.Number ? value.value : undefined;
}

function toCellRange(arg: LookupBuiltinArgument): RangeBuiltinArgument | CellValue {
  if (!isRangeArg(arg)) {
    return { kind: "range", values: [arg], refKind: "cells", rows: 1, cols: 1 };
  }
  if (arg.refKind !== "cells") {
    return errorValue(ErrorCode.Value);
  }
  return arg;
}

function toNumericMatrix(arg: LookupBuiltinArgument): number[][] | CellValue {
  const range = toCellRange(arg);
  if (!isRangeArg(range)) {
    return range;
  }
  const matrix: number[][] = [];
  for (let row = 0; row < range.rows; row += 1) {
    const rowValues: number[] = [];
    for (let col = 0; col < range.cols; col += 1) {
      const numeric = toNumber(getRangeValue(range, row, col));
      if (numeric === undefined) {
        return errorValue(ErrorCode.Value);
      }
      rowValues.push(numeric);
    }
    matrix.push(rowValues);
  }
  return matrix;
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

function collectSampleNumbers(arg: LookupBuiltinArgument): number[] | CellValue {
  if (!isRangeArg(arg)) {
    if (isError(arg)) {
      return arg;
    }
    return arg.tag === ValueTag.Number ? [arg.value] : errorValue(ErrorCode.Value);
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
): CellValue {
  const actual = toNumericMatrix(actualArg);
  if (!Array.isArray(actual)) {
    return actual;
  }
  const expected = toNumericMatrix(expectedArg);
  if (!Array.isArray(expected)) {
    return expected;
  }

  const rows = actual.length;
  const cols = actual[0]?.length ?? 0;
  if (rows !== expected.length || cols !== (expected[0]?.length ?? 0)) {
    return errorValue(ErrorCode.NA);
  }
  if ((rows === 1 && cols === 1) || rows === 0 || cols === 0) {
    return errorValue(ErrorCode.NA);
  }

  let statistic = 0;
  for (let row = 0; row < rows; row += 1) {
    const actualRow = actual[row]!;
    const expectedRow = expected[row]!;
    if (actualRow.length !== cols || expectedRow.length !== cols) {
      return errorValue(ErrorCode.NA);
    }
    for (let col = 0; col < cols; col += 1) {
      const actualValue = actualRow[col]!;
      const expectedValue = expectedRow[col]!;
      if (actualValue < 0 || expectedValue < 0) {
        return errorValue(ErrorCode.Value);
      }
      if (expectedValue === 0) {
        return errorValue(ErrorCode.Div0);
      }
      const delta = actualValue - expectedValue;
      statistic += (delta * delta) / expectedValue;
    }
  }

  const degrees = rows > 1 && cols > 1 ? (rows - 1) * (cols - 1) : rows > 1 ? rows - 1 : cols - 1;
  if (degrees <= 0) {
    return errorValue(ErrorCode.NA);
  }
  const probability = regularizedUpperGamma(degrees / 2, statistic / 2);
  return Number.isFinite(probability)
    ? { tag: ValueTag.Number, value: probability }
    : errorValue(ErrorCode.Value);
}

function fTestResult(firstArg: LookupBuiltinArgument, secondArg: LookupBuiltinArgument): CellValue {
  const first = collectSampleNumbers(firstArg);
  if (!Array.isArray(first)) {
    return first;
  }
  const second = collectSampleNumbers(secondArg);
  if (!Array.isArray(second)) {
    return second;
  }
  if (first.length < 2 || second.length < 2) {
    return errorValue(ErrorCode.Div0);
  }

  const firstMean = first.reduce((sum, value) => sum + value, 0) / first.length;
  const secondMean = second.reduce((sum, value) => sum + value, 0) / second.length;
  const firstVariance =
    first.reduce((sum, value) => sum + (value - firstMean) ** 2, 0) / (first.length - 1);
  const secondVariance =
    second.reduce((sum, value) => sum + (value - secondMean) ** 2, 0) / (second.length - 1);
  if (!(firstVariance > 0) || !(secondVariance > 0)) {
    return errorValue(ErrorCode.Div0);
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
    : errorValue(ErrorCode.Value);
}

function zTestResult(
  arrayArg: LookupBuiltinArgument,
  xArg: LookupBuiltinArgument,
  sigmaArg?: LookupBuiltinArgument,
): CellValue {
  const sample = collectSampleNumbers(arrayArg);
  if (!Array.isArray(sample)) {
    return sample;
  }
  const x = !isRangeArg(xArg) ? toNumber(xArg) : undefined;
  if (x === undefined || sample.length === 0) {
    return errorValue(ErrorCode.Value);
  }

  let sigma: number | undefined;
  if (sigmaArg !== undefined) {
    sigma = !isRangeArg(sigmaArg) ? toNumber(sigmaArg) : undefined;
  } else if (sample.length >= 2) {
    const mean = sample.reduce((sum, value) => sum + value, 0) / sample.length;
    const variance =
      sample.reduce((sum, value) => sum + (value - mean) ** 2, 0) / (sample.length - 1);
    sigma = variance > 0 ? Math.sqrt(variance) : undefined;
  }

  if (sigma === undefined || !(sigma > 0)) {
    return errorValue(ErrorCode.Div0);
  }
  const mean = sample.reduce((sum, value) => sum + value, 0) / sample.length;
  const zScore = (mean - x) / (sigma / Math.sqrt(sample.length));
  const probability = 1 - standardNormalCdf(zScore);
  return Number.isFinite(probability)
    ? { tag: ValueTag.Number, value: probability }
    : errorValue(ErrorCode.Value);
}

function tTestResult(
  firstArg: LookupBuiltinArgument,
  secondArg: LookupBuiltinArgument,
  tailsArg: LookupBuiltinArgument,
  typeArg: LookupBuiltinArgument,
): CellValue {
  const first = collectSampleNumbers(firstArg);
  if (!Array.isArray(first)) {
    return first;
  }
  const second = collectSampleNumbers(secondArg);
  if (!Array.isArray(second)) {
    return second;
  }
  const tails = !isRangeArg(tailsArg) ? toNumber(tailsArg) : undefined;
  const type = !isRangeArg(typeArg) ? toNumber(typeArg) : undefined;
  if (
    tails === undefined ||
    type === undefined ||
    !Number.isInteger(tails) ||
    !Number.isInteger(type) ||
    ![1, 2].includes(tails) ||
    ![1, 2, 3].includes(type)
  ) {
    return errorValue(ErrorCode.Value);
  }

  let statistic: number;
  let degreesFreedom: number;
  if (type === 1) {
    if (first.length !== second.length) {
      return errorValue(ErrorCode.NA);
    }
    if (first.length < 2) {
      return errorValue(ErrorCode.Div0);
    }
    const deltas = first.map((value, index) => value - second[index]!);
    const mean = deltas.reduce((sum, value) => sum + value, 0) / deltas.length;
    const variance =
      deltas.reduce((sum, value) => sum + (value - mean) ** 2, 0) / (deltas.length - 1);
    if (!(variance > 0)) {
      return errorValue(ErrorCode.Div0);
    }
    statistic = mean / Math.sqrt(variance / deltas.length);
    degreesFreedom = deltas.length - 1;
  } else {
    if (first.length < 2 || second.length < 2) {
      return errorValue(ErrorCode.Div0);
    }
    const firstMean = first.reduce((sum, value) => sum + value, 0) / first.length;
    const secondMean = second.reduce((sum, value) => sum + value, 0) / second.length;
    const firstVariance =
      first.reduce((sum, value) => sum + (value - firstMean) ** 2, 0) / (first.length - 1);
    const secondVariance =
      second.reduce((sum, value) => sum + (value - secondMean) ** 2, 0) / (second.length - 1);
    if (!(firstVariance > 0) || !(secondVariance > 0)) {
      return errorValue(ErrorCode.Div0);
    }

    if (type === 2) {
      const pooledVariance =
        ((first.length - 1) * firstVariance + (second.length - 1) * secondVariance) /
        (first.length + second.length - 2);
      if (!(pooledVariance > 0)) {
        return errorValue(ErrorCode.Div0);
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
        return errorValue(ErrorCode.Div0);
      }
      statistic = (firstMean - secondMean) / denominator;
      degreesFreedom = (firstTerm + secondTerm) ** 2 / welchDenominator;
    }
  }

  const upperTail = 1 - studentTCdf(Math.abs(statistic), degreesFreedom);
  const probability = tails === 1 ? upperTail : Math.min(1, upperTail * 2);
  return Number.isFinite(probability)
    ? { tag: ValueTag.Number, value: probability }
    : errorValue(ErrorCode.Value);
}

function determinantOf(matrix: number[][]): number {
  const size = matrix.length;
  const working = matrix.map((row) => [...row]);
  let determinant = 1;
  let sign = 1;
  for (let pivot = 0; pivot < size; pivot += 1) {
    let pivotRow = pivot;
    while (pivotRow < size && working[pivotRow]![pivot] === 0) {
      pivotRow += 1;
    }
    if (pivotRow === size) {
      return 0;
    }
    if (pivotRow !== pivot) {
      const pivotValues = working[pivot];
      const swapValues = working[pivotRow];
      if (!pivotValues || !swapValues) {
        return 0;
      }
      [working[pivot], working[pivotRow]] = [swapValues, pivotValues];
      sign *= -1;
    }
    const pivotValue = working[pivot]![pivot]!;
    determinant *= pivotValue;
    for (let row = pivot + 1; row < size; row += 1) {
      const factor = working[row]![pivot]! / pivotValue;
      for (let col = pivot; col < size; col += 1) {
        working[row]![col] = working[row]![col]! - factor * working[pivot]![col]!;
      }
    }
  }
  return determinant * sign;
}

function inverseOf(matrix: number[][]): number[][] | undefined {
  const size = matrix.length;
  const augmented = matrix.map((row, rowIndex) => [
    ...row,
    ...Array.from({ length: size }, (_, colIndex) => (rowIndex === colIndex ? 1 : 0)),
  ]);
  for (let pivot = 0; pivot < size; pivot += 1) {
    let pivotRow = pivot;
    while (pivotRow < size && augmented[pivotRow]![pivot] === 0) {
      pivotRow += 1;
    }
    if (pivotRow === size) {
      return undefined;
    }
    if (pivotRow !== pivot) {
      const pivotValues = augmented[pivot];
      const swapValues = augmented[pivotRow];
      if (!pivotValues || !swapValues) {
        return undefined;
      }
      [augmented[pivot], augmented[pivotRow]] = [swapValues, pivotValues];
    }
    const pivotValue = augmented[pivot]![pivot]!;
    if (pivotValue === 0) {
      return undefined;
    }
    for (let col = 0; col < size * 2; col += 1) {
      augmented[pivot]![col] = augmented[pivot]![col]! / pivotValue;
    }
    for (let row = 0; row < size; row += 1) {
      if (row === pivot) {
        continue;
      }
      const factor = augmented[row]![pivot]!;
      for (let col = 0; col < size * 2; col += 1) {
        augmented[row]![col] = augmented[row]![col]! - factor * augmented[pivot]![col]!;
      }
    }
  }
  return augmented.map((row) => row.slice(size));
}

function flattenNumbers(arg: LookupBuiltinArgument): number[] | CellValue {
  if (!isRangeArg(arg)) {
    const numeric = toNumber(arg);
    return numeric === undefined ? errorValue(ErrorCode.Value) : [numeric];
  }
  const values: number[] = [];
  for (const value of arg.values) {
    const numeric = toNumber(value);
    if (numeric === undefined) {
      return errorValue(ErrorCode.Value);
    }
    values.push(numeric);
  }
  return values;
}

function arrayTextCell(value: CellValue, strict: boolean): string | undefined {
  switch (value.tag) {
    case ValueTag.Empty:
      return "";
    case ValueTag.Number:
      return String(value.value);
    case ValueTag.Boolean:
      return value.value ? "TRUE" : "FALSE";
    case ValueTag.String:
      return strict ? `"${value.value.replace(/"/g, '""')}"` : value.value;
    case ValueTag.Error:
      return undefined;
  }
}

function sumOfNumbers(arg: LookupBuiltinArgument): number | CellValue {
  const values = flattenNumbers(arg);
  return Array.isArray(values) ? values.reduce((sum, value) => sum + value, 0) : values;
}

function pickRangeRow(range: RangeBuiltinArgument, row: number): CellValue[] {
  const values: CellValue[] = [];
  for (let col = 0; col < range.cols; col += 1) {
    values.push(getRangeValue(range, row, col));
  }
  return values;
}

function pickRangeCol(range: RangeBuiltinArgument, col: number): CellValue[] {
  const values: CellValue[] = [];
  for (let row = 0; row < range.rows; row += 1) {
    values.push(getRangeValue(range, row, col));
  }
  return values;
}

function normalizeKeyValue(value: CellValue): CellValue {
  if (value.tag !== ValueTag.String) {
    return value;
  }
  return {
    tag: ValueTag.String,
    value: value.value.toUpperCase(),
    stringId: value.stringId,
  };
}

function rowKey(range: RangeBuiltinArgument, row: number): string | undefined {
  const values = pickRangeRow(range, row);
  if (values.some(isError)) {
    return undefined;
  }
  return JSON.stringify(values.map(normalizeKeyValue));
}

function clipIndex(value: number, length: number): number | undefined {
  if (!Number.isFinite(value) || length <= 0) {
    return undefined;
  }
  const index = Math.trunc(value);
  if (index === 0) {
    return undefined;
  }
  return index < 0 ? Math.max(index, -length) : Math.min(index, length);
}

function flattenValues(
  range: RangeBuiltinArgument,
  scanByCol: boolean,
  ignoreEmpty = false,
): CellValue[] {
  const values: CellValue[] = [];
  if (scanByCol) {
    for (let col = 0; col < range.cols; col += 1) {
      for (let row = 0; row < range.rows; row += 1) {
        const value = getRangeValue(range, row, col);
        if (ignoreEmpty && value.tag === ValueTag.Empty) {
          continue;
        }
        values.push(value);
      }
    }
    return values;
  }
  for (let row = 0; row < range.rows; row += 1) {
    for (let col = 0; col < range.cols; col += 1) {
      const value = getRangeValue(range, row, col);
      if (ignoreEmpty && value.tag === ValueTag.Empty) {
        continue;
      }
      values.push(value);
    }
  }
  return values;
}

function getRangeWindowValues(
  range: RangeBuiltinArgument,
  rowStart: number,
  colStart: number,
  rowCount: number,
  colCount: number,
): CellValue[] {
  const values: CellValue[] = [];
  for (let row = 0; row < rowCount; row += 1) {
    for (let col = 0; col < colCount; col += 1) {
      values.push(getRangeValue(range, rowStart + row, colStart + col));
    }
  }
  return values;
}

function colKey(range: RangeBuiltinArgument, col: number): string | undefined {
  const values = pickRangeCol(range, col);
  if (values.some(isError)) {
    return undefined;
  }
  return JSON.stringify(values.map(normalizeKeyValue));
}

function exactMatch(lookupValue: CellValue, range: RangeBuiltinArgument): number {
  for (let index = 0; index < range.values.length; index += 1) {
    const comparison = compareScalars(range.values[index]!, lookupValue);
    if (comparison === 0) {
      return index + 1;
    }
  }
  return -1;
}

function approximateMatchAscending(lookupValue: CellValue, range: RangeBuiltinArgument): number {
  let best = -1;
  for (let index = 0; index < range.values.length; index += 1) {
    const comparison = compareScalars(range.values[index]!, lookupValue);
    if (comparison === undefined) {
      return -1;
    }
    if (comparison <= 0) {
      best = index + 1;
    } else {
      break;
    }
  }
  return best;
}

function approximateMatchDescending(lookupValue: CellValue, range: RangeBuiltinArgument): number {
  let best = -1;
  for (let index = 0; index < range.values.length; index += 1) {
    const comparison = compareScalars(range.values[index]!, lookupValue);
    if (comparison === undefined) {
      return -1;
    }
    if (comparison >= 0) {
      best = index + 1;
      continue;
    }
    break;
  }
  return best;
}

function firstLookupError(args: readonly LookupBuiltinArgument[]): CellValue | undefined {
  return args.find((arg) => isError(arg));
}

const externalLookupBuiltinNames = ["FILTERXML", "STOCKHISTORY"] as const;

function createExternalLookupBuiltin(name: string): LookupBuiltin {
  return (...args) => {
    const existingError = firstLookupError(args);
    if (existingError) {
      return existingError;
    }
    const external = getExternalLookupFunction(name);
    return external ? external(...args) : errorValue(ErrorCode.Blocked);
  };
}

const externalLookupBuiltins = Object.fromEntries(
  externalLookupBuiltinNames.map((name) => [name, createExternalLookupBuiltin(name)]),
) as Record<string, LookupBuiltin>;

const lookupRegressionBuiltins = createLookupRegressionBuiltins({
  errorValue,
  numberResult,
  isRangeArg,
  toNumber,
  toBoolean,
  flattenNumbers,
});

const lookupOrderStatisticsBuiltins = createLookupOrderStatisticsBuiltins({
  errorValue,
  numberResult,
  arrayResult,
  requireCellRange,
  isError,
  isRangeArg,
  toNumber,
  toInteger,
  flattenNumbers,
});

const lookupFinancialBuiltins = createLookupFinancialBuiltins({
  errorValue,
  numberResult,
  isRangeArg,
  toNumber,
  collectNumericSeries,
});

const lookupDatabaseBuiltins = createLookupDatabaseBuiltins({
  errorValue,
  numberResult,
  isError,
  isRangeArg,
  toNumber,
  toStringValue,
  requireCellRange,
  getRangeValue,
  matchesCriteria,
});

export const lookupBuiltins: Record<string, LookupBuiltin> = {
  ...lookupDatabaseBuiltins,
  ...lookupFinancialBuiltins,
  MATCH: (lookupValue, lookupArray, matchTypeValue = { tag: ValueTag.Number, value: 1 }) => {
    if (isRangeArg(lookupValue)) {
      return errorValue(ErrorCode.Value);
    }
    if (isRangeArg(matchTypeValue)) {
      return errorValue(ErrorCode.Value);
    }
    if (isError(lookupValue)) {
      return lookupValue;
    }
    if (isError(matchTypeValue)) {
      return matchTypeValue;
    }

    const rangeOrError = requireCellVector(lookupArray);
    if (!isRangeArg(rangeOrError)) {
      return rangeOrError;
    }

    const matchType = toInteger(matchTypeValue);
    if (matchType === undefined || ![-1, 0, 1].includes(matchType)) {
      return errorValue(ErrorCode.Value);
    }

    const position =
      matchType === 0
        ? exactMatch(lookupValue, rangeOrError)
        : matchType === 1
          ? approximateMatchAscending(lookupValue, rangeOrError)
          : approximateMatchDescending(lookupValue, rangeOrError);

    return position === -1 ? errorValue(ErrorCode.NA) : { tag: ValueTag.Number, value: position };
  },
  LOOKUP: (lookupValue, lookupVectorArg, resultVectorArg = lookupVectorArg) => {
    if (isRangeArg(lookupValue) || lookupValue === undefined || resultVectorArg === undefined) {
      return errorValue(ErrorCode.Value);
    }

    const existingError = isError(lookupValue)
      ? lookupValue
      : isError(lookupVectorArg)
        ? lookupVectorArg
        : isError(resultVectorArg)
          ? resultVectorArg
          : undefined;
    if (existingError) {
      return existingError;
    }

    const lookupRangeOrError = toCellRange(lookupVectorArg);
    const resultRangeOrError = toCellRange(resultVectorArg);
    if (!isRangeArg(lookupRangeOrError)) {
      return lookupRangeOrError;
    }
    if (!isRangeArg(resultRangeOrError)) {
      return resultRangeOrError;
    }

    if (lookupRangeOrError.rows !== 1 && lookupRangeOrError.cols !== 1) {
      return errorValue(ErrorCode.Value);
    }
    if (resultRangeOrError.rows !== 1 && resultRangeOrError.cols !== 1) {
      return errorValue(ErrorCode.Value);
    }
    if (lookupRangeOrError.values.length !== resultRangeOrError.values.length) {
      return errorValue(ErrorCode.Value);
    }

    const exactPosition = exactMatch(lookupValue, lookupRangeOrError);
    const shouldApproximate = exactPosition === -1 && lookupValue.tag === ValueTag.Number;
    const position = shouldApproximate
      ? approximateMatchAscending(lookupValue, lookupRangeOrError)
      : exactPosition;

    if (position === -1) {
      return errorValue(ErrorCode.NA);
    }

    const resultIndex = position - 1;
    return resultRangeOrError.values[resultIndex] ?? errorValue(ErrorCode.NA);
  },
  AREAS: (arrayArg) => {
    const range = requireCellRange(arrayArg);
    if (!isRangeArg(range)) {
      return range;
    }
    return { tag: ValueTag.Number, value: 1 };
  },
  ARRAYTOTEXT: (arrayArg, formatArg = { tag: ValueTag.Number, value: 0 }) => {
    const array = toCellRange(arrayArg);
    if (!isRangeArg(array)) {
      return array;
    }
    if (isRangeArg(formatArg)) {
      return errorValue(ErrorCode.Value);
    }
    const format = toInteger(formatArg);
    if (format === undefined || (format !== 0 && format !== 1)) {
      return errorValue(ErrorCode.Value);
    }
    const strict = format === 1;
    const lines: string[] = [];
    for (let row = 0; row < array.rows; row += 1) {
      const lineValues: string[] = [];
      for (let col = 0; col < array.cols; col += 1) {
        const value = arrayTextCell(getRangeValue(array, row, col), strict);
        if (value === undefined) {
          return errorValue(ErrorCode.Value);
        }
        lineValues.push(value);
      }
      lines.push(strict ? lineValues.join(", ") : lineValues.join("\t"));
    }
    const body = lines.join(";");
    return {
      tag: ValueTag.String,
      value: strict ? `{${body}}` : body,
      stringId: 0,
    };
  },
  COLUMNS: (arrayArg) => {
    const range = requireCellRange(arrayArg);
    if (!isRangeArg(range)) {
      return range;
    }
    return { tag: ValueTag.Number, value: range.cols };
  },
  ROWS: (arrayArg) => {
    const range = requireCellRange(arrayArg);
    if (!isRangeArg(range)) {
      return range;
    }
    return { tag: ValueTag.Number, value: range.rows };
  },
  ...lookupRegressionBuiltins,
  "CHISQ.TEST": (actualArg, expectedArg) => {
    return chiSquareTestResult(actualArg, expectedArg);
  },
  CHITEST: (actualArg, expectedArg) => {
    return chiSquareTestResult(actualArg, expectedArg);
  },
  "LEGACY.CHITEST": (actualArg, expectedArg) => {
    return chiSquareTestResult(actualArg, expectedArg);
  },
  "F.TEST": (firstArg, secondArg) => {
    return fTestResult(firstArg, secondArg);
  },
  FTEST: (firstArg, secondArg) => {
    return fTestResult(firstArg, secondArg);
  },
  "Z.TEST": (arrayArg, xArg, sigmaArg) => {
    return zTestResult(arrayArg, xArg, sigmaArg);
  },
  ZTEST: (arrayArg, xArg, sigmaArg) => {
    return zTestResult(arrayArg, xArg, sigmaArg);
  },
  "T.TEST": (firstArg, secondArg, tailsArg, typeArg) => {
    return tTestResult(firstArg, secondArg, tailsArg, typeArg);
  },
  TTEST: (firstArg, secondArg, tailsArg, typeArg) => {
    return tTestResult(firstArg, secondArg, tailsArg, typeArg);
  },
  ...lookupOrderStatisticsBuiltins,
  INDEX: (array, rowNumValue, colNumValue = { tag: ValueTag.Number, value: 1 }) => {
    if (!isRangeArg(array) || array.refKind !== "cells") {
      return errorValue(ErrorCode.Value);
    }
    if (isRangeArg(rowNumValue) || isRangeArg(colNumValue)) {
      return errorValue(ErrorCode.Value);
    }
    if (isError(rowNumValue)) {
      return rowNumValue;
    }
    if (isError(colNumValue)) {
      return colNumValue;
    }

    const rawRowNum = toInteger(rowNumValue);
    const rawColNum = toInteger(colNumValue);
    if (rawRowNum === undefined || rawColNum === undefined) {
      return errorValue(ErrorCode.Value);
    }

    let rowNum = rawRowNum;
    let colNum = rawColNum;
    if (array.rows === 1 && rawColNum === 1) {
      rowNum = 1;
      colNum = rawRowNum;
    }

    if (rowNum < 1 || colNum < 1 || rowNum > array.rows || colNum > array.cols) {
      return errorValue(ErrorCode.Ref);
    }

    return getRangeValue(array, rowNum - 1, colNum - 1);
  },
  VLOOKUP: (
    lookupValue,
    tableArray,
    colIndexValue,
    rangeLookupValue = { tag: ValueTag.Boolean, value: true },
  ) => {
    if (isRangeArg(lookupValue)) {
      return errorValue(ErrorCode.Value);
    }
    if (!isRangeArg(tableArray) || tableArray.refKind !== "cells") {
      return errorValue(ErrorCode.Value);
    }
    if (isRangeArg(colIndexValue) || isRangeArg(rangeLookupValue)) {
      return errorValue(ErrorCode.Value);
    }
    if (isError(lookupValue)) {
      return lookupValue;
    }
    if (isError(colIndexValue)) {
      return colIndexValue;
    }
    if (isError(rangeLookupValue)) {
      return rangeLookupValue;
    }

    const colIndex = toInteger(colIndexValue);
    const rangeLookup = toBoolean(rangeLookupValue);
    if (
      colIndex === undefined ||
      colIndex < 1 ||
      colIndex > tableArray.cols ||
      rangeLookup === undefined
    ) {
      return errorValue(ErrorCode.Value);
    }

    let matchedRow = -1;
    for (let row = 0; row < tableArray.rows; row += 1) {
      const comparison = compareScalars(getRangeValue(tableArray, row, 0), lookupValue);
      if (comparison === undefined) {
        return errorValue(ErrorCode.Value);
      }
      if (comparison === 0) {
        matchedRow = row;
        break;
      }
      if (rangeLookup && comparison < 0) {
        matchedRow = row;
        continue;
      }
      if (rangeLookup && comparison > 0) {
        break;
      }
    }

    if (matchedRow === -1) {
      return errorValue(ErrorCode.NA);
    }
    return getRangeValue(tableArray, matchedRow, colIndex - 1);
  },
  HLOOKUP: (
    lookupValue,
    tableArray,
    rowIndexValue,
    rangeLookupValue = { tag: ValueTag.Boolean, value: true },
  ) => {
    if (isRangeArg(lookupValue)) {
      return errorValue(ErrorCode.Value);
    }
    if (!isRangeArg(tableArray) || tableArray.refKind !== "cells") {
      return errorValue(ErrorCode.Value);
    }
    if (isRangeArg(rowIndexValue) || isRangeArg(rangeLookupValue)) {
      return errorValue(ErrorCode.Value);
    }
    if (isError(lookupValue)) {
      return lookupValue;
    }
    if (isError(rowIndexValue)) {
      return rowIndexValue;
    }
    if (isError(rangeLookupValue)) {
      return rangeLookupValue;
    }

    const rowIndex = toInteger(rowIndexValue);
    const rangeLookup = toBoolean(rangeLookupValue);
    if (
      rowIndex === undefined ||
      rowIndex < 1 ||
      rowIndex > tableArray.rows ||
      rangeLookup === undefined
    ) {
      return errorValue(ErrorCode.Value);
    }

    let matchedCol = -1;
    for (let col = 0; col < tableArray.cols; col += 1) {
      const comparison = compareScalars(getRangeValue(tableArray, 0, col), lookupValue);
      if (comparison === undefined) {
        return errorValue(ErrorCode.Value);
      }
      if (comparison === 0) {
        matchedCol = col;
        break;
      }
      if (rangeLookup && comparison < 0) {
        matchedCol = col;
        continue;
      }
      if (rangeLookup && comparison > 0) {
        break;
      }
    }

    if (matchedCol === -1) {
      return errorValue(ErrorCode.NA);
    }
    return getRangeValue(tableArray, rowIndex - 1, matchedCol);
  },
  XLOOKUP: (
    lookupValue,
    lookupArray,
    returnArray,
    ifNotFound = { tag: ValueTag.Error, code: ErrorCode.NA },
    matchMode = { tag: ValueTag.Number, value: 0 },
    searchMode = { tag: ValueTag.Number, value: 1 },
  ) => {
    if (
      isRangeArg(lookupValue) ||
      isRangeArg(ifNotFound) ||
      isRangeArg(matchMode) ||
      isRangeArg(searchMode)
    ) {
      return errorValue(ErrorCode.Value);
    }
    const lookupRange = requireCellVector(lookupArray);
    const returnRange = requireCellVector(returnArray);
    if (!isRangeArg(lookupRange)) {
      return lookupRange;
    }
    if (!isRangeArg(returnRange)) {
      return returnRange;
    }
    if (lookupRange.values.length !== returnRange.values.length) {
      return errorValue(ErrorCode.Value);
    }
    if (isError(lookupValue)) {
      return lookupValue;
    }
    if (isError(matchMode)) {
      return matchMode;
    }
    if (isError(searchMode)) {
      return searchMode;
    }

    const matchModeNumber = toInteger(matchMode);
    const searchModeNumber = toInteger(searchMode);
    if ((matchModeNumber ?? 0) !== 0 || (searchModeNumber !== 1 && searchModeNumber !== -1)) {
      return errorValue(ErrorCode.Value);
    }

    if (searchModeNumber === -1) {
      for (let index = lookupRange.values.length - 1; index >= 0; index -= 1) {
        if (compareScalars(lookupRange.values[index]!, lookupValue) === 0) {
          return returnRange.values[index] ?? errorValue(ErrorCode.NA);
        }
      }
      return ifNotFound;
    }

    for (let index = 0; index < lookupRange.values.length; index += 1) {
      if (compareScalars(lookupRange.values[index]!, lookupValue) === 0) {
        return returnRange.values[index] ?? errorValue(ErrorCode.NA);
      }
    }
    return ifNotFound;
  },
  XMATCH: (
    lookupValue,
    lookupArray,
    matchModeValue = { tag: ValueTag.Number, value: 0 },
    searchModeValue = { tag: ValueTag.Number, value: 1 },
  ) => {
    if (isRangeArg(lookupValue) || isRangeArg(matchModeValue) || isRangeArg(searchModeValue)) {
      return errorValue(ErrorCode.Value);
    }
    if (isError(lookupValue)) {
      return lookupValue;
    }
    if (isError(matchModeValue)) {
      return matchModeValue;
    }
    if (isError(searchModeValue)) {
      return searchModeValue;
    }
    const rangeOrError = requireCellVector(lookupArray);
    if (!isRangeArg(rangeOrError)) {
      return rangeOrError;
    }
    const matchMode = toInteger(matchModeValue);
    const searchMode = toInteger(searchModeValue);
    if (matchMode === undefined || searchMode === undefined) {
      return errorValue(ErrorCode.Value);
    }
    if (![0, -1, 1].includes(matchMode) || ![1, -1].includes(searchMode)) {
      return errorValue(ErrorCode.Value);
    }

    const values = searchMode === -1 ? rangeOrError.values.toReversed() : rangeOrError.values;
    const probe = searchMode === -1 ? { ...rangeOrError, values } : rangeOrError;
    const position =
      matchMode === 0
        ? exactMatch(lookupValue, probe)
        : matchMode === 1
          ? approximateMatchAscending(lookupValue, probe)
          : approximateMatchDescending(lookupValue, probe);
    if (position === -1) {
      return errorValue(ErrorCode.NA);
    }
    const normalizedPosition =
      searchMode === -1 ? rangeOrError.values.length - position + 1 : position;
    return { tag: ValueTag.Number, value: normalizedPosition };
  },
  OFFSET: (referenceArg, rowsArg, colsArg, heightArg, widthArg, areaNumberArg) => {
    if (
      isRangeArg(rowsArg) ||
      isRangeArg(colsArg) ||
      isRangeArg(heightArg) ||
      isRangeArg(widthArg) ||
      isRangeArg(areaNumberArg)
    ) {
      return errorValue(ErrorCode.Value);
    }
    if (
      isError(rowsArg) ||
      isError(colsArg) ||
      isError(heightArg) ||
      isError(widthArg) ||
      isError(areaNumberArg)
    ) {
      return isError(rowsArg)
        ? rowsArg
        : isError(colsArg)
          ? colsArg
          : isError(heightArg)
            ? heightArg
            : isError(widthArg)
              ? widthArg
              : areaNumberArg;
    }
    const reference = toCellRange(referenceArg);
    if (!isRangeArg(reference)) {
      return reference;
    }
    const rows = toInteger(rowsArg);
    const cols = toInteger(colsArg);
    const height = heightArg === undefined ? reference.rows : toInteger(heightArg);
    const width = widthArg === undefined ? reference.cols : toInteger(widthArg);
    const areaNumber = areaNumberArg === undefined ? 1 : toInteger(areaNumberArg);
    if (
      rows === undefined ||
      cols === undefined ||
      height === undefined ||
      width === undefined ||
      areaNumber === undefined
    ) {
      return errorValue(ErrorCode.Value);
    }
    if (areaNumber !== 1) {
      return errorValue(ErrorCode.Value);
    }
    if (height < 1 || width < 1) {
      return errorValue(ErrorCode.Value);
    }

    const rowStart = rows < 0 ? reference.rows + rows : rows;
    const colStart = cols < 0 ? reference.cols + cols : cols;
    if (
      rowStart < 0 ||
      colStart < 0 ||
      rowStart + height > reference.rows ||
      colStart + width > reference.cols
    ) {
      return errorValue(ErrorCode.Ref);
    }
    if (height === 1 && width === 1) {
      return getRangeValue(reference, rowStart, colStart);
    }
    return arrayResult(
      getRangeWindowValues(reference, rowStart, colStart, height, width),
      height,
      width,
    );
  },
  TAKE: (arrayArg, rowsArg, colsArg) => {
    const array = toCellRange(arrayArg);
    if (!isRangeArg(array)) {
      return array;
    }
    if (isRangeArg(rowsArg) || isRangeArg(colsArg)) {
      return errorValue(ErrorCode.Value);
    }
    if (isError(rowsArg) || isError(colsArg)) {
      return rowsArg.tag === ValueTag.Error ? rowsArg : colsArg;
    }

    const requestedRows = rowsArg === undefined ? array.rows : toInteger(rowsArg);
    const requestedCols = colsArg === undefined ? array.cols : toInteger(colsArg);
    if (requestedRows === undefined || requestedCols === undefined) {
      return errorValue(ErrorCode.Value);
    }

    const clippedRows = clipIndex(requestedRows, array.rows);
    const clippedCols = clipIndex(requestedCols, array.cols);
    if (clippedRows === undefined || clippedCols === undefined) {
      return errorValue(ErrorCode.Value);
    }

    const rowCount =
      clippedRows > 0 ? Math.min(clippedRows, array.rows) : Math.min(-clippedRows, array.rows);
    const colCount =
      clippedCols > 0 ? Math.min(clippedCols, array.cols) : Math.min(-clippedCols, array.cols);
    const rowOffset = clippedRows > 0 ? 0 : Math.max(array.rows - rowCount, 0);
    const colOffset = clippedCols > 0 ? 0 : Math.max(array.cols - colCount, 0);
    if (rowCount === 0 || colCount === 0) {
      return errorValue(ErrorCode.Value);
    }

    const values: CellValue[] = [];
    for (let row = 0; row < rowCount; row += 1) {
      for (let col = 0; col < colCount; col += 1) {
        values.push(getRangeValue(array, row + rowOffset, col + colOffset));
      }
    }
    return arrayResult(values, rowCount, colCount);
  },
  DROP: (arrayArg, rowsArg, colsArg) => {
    const array = toCellRange(arrayArg);
    if (!isRangeArg(array)) {
      return array;
    }
    if (isRangeArg(rowsArg) || isRangeArg(colsArg)) {
      return errorValue(ErrorCode.Value);
    }
    if (isError(rowsArg) || isError(colsArg)) {
      return rowsArg.tag === ValueTag.Error ? rowsArg : colsArg;
    }

    const requestedRows = rowsArg === undefined ? 0 : toInteger(rowsArg);
    const requestedCols = colsArg === undefined ? 0 : toInteger(colsArg);
    if (requestedRows === undefined || requestedCols === undefined) {
      return errorValue(ErrorCode.Value);
    }

    const clippedRows = requestedRows === 0 ? 0 : clipIndex(requestedRows, array.rows);
    const clippedCols = requestedCols === 0 ? 0 : clipIndex(requestedCols, array.cols);
    if (clippedRows === undefined || clippedCols === undefined) {
      return errorValue(ErrorCode.Value);
    }

    const dropRows =
      clippedRows >= 0 ? Math.min(clippedRows, array.rows) : Math.min(-clippedRows, array.rows);
    const dropCols =
      clippedCols >= 0 ? Math.min(clippedCols, array.cols) : Math.min(-clippedCols, array.cols);
    const rowCount = array.rows - dropRows;
    const colCount = array.cols - dropCols;
    const rowOffset = clippedRows > 0 ? dropRows : 0;
    const colOffset = clippedCols > 0 ? dropCols : 0;
    if (rowCount <= 0 || colCount <= 0) {
      return errorValue(ErrorCode.Value);
    }

    const values: CellValue[] = [];
    for (let row = 0; row < rowCount; row += 1) {
      for (let col = 0; col < colCount; col += 1) {
        values.push(getRangeValue(array, row + rowOffset, col + colOffset));
      }
    }
    return arrayResult(values, rowCount, colCount);
  },
  CHOOSECOLS: (arrayArg, ...columnArgs) => {
    const array = toCellRange(arrayArg);
    if (!isRangeArg(array)) {
      return array;
    }
    if (columnArgs.length === 0) {
      return errorValue(ErrorCode.Value);
    }

    const selectedCols: number[] = [];
    for (const arg of columnArgs) {
      if (isRangeArg(arg)) {
        return errorValue(ErrorCode.Value);
      }
      if (isError(arg)) {
        return arg;
      }
      const selected = toInteger(arg);
      if (selected === undefined || selected < 1 || selected > array.cols) {
        return errorValue(ErrorCode.Value);
      }
      selectedCols.push(selected - 1);
    }

    const values: CellValue[] = [];
    for (let row = 0; row < array.rows; row += 1) {
      for (const col of selectedCols) {
        values.push(getRangeValue(array, row, col));
      }
    }
    return arrayResult(values, array.rows, selectedCols.length);
  },
  CHOOSEROWS: (arrayArg, ...rowArgs) => {
    const array = toCellRange(arrayArg);
    if (!isRangeArg(array)) {
      return array;
    }
    if (rowArgs.length === 0) {
      return errorValue(ErrorCode.Value);
    }

    const selectedRows: number[] = [];
    for (const arg of rowArgs) {
      if (isRangeArg(arg)) {
        return errorValue(ErrorCode.Value);
      }
      if (isError(arg)) {
        return arg;
      }
      const selected = toInteger(arg);
      if (selected === undefined || selected < 1 || selected > array.rows) {
        return errorValue(ErrorCode.Value);
      }
      selectedRows.push(selected - 1);
    }

    const values: CellValue[] = [];
    for (const row of selectedRows) {
      values.push(...pickRangeRow(array, row));
    }
    return arrayResult(values, selectedRows.length, array.cols);
  },
  SORT: (arrayArg, sortIndexArg, sortOrderArg = { tag: ValueTag.Number, value: 1 }, byColArg) => {
    const array = toCellRange(arrayArg);
    if (!isRangeArg(array)) {
      return array;
    }
    if (isRangeArg(sortIndexArg) || isRangeArg(sortOrderArg) || isRangeArg(byColArg)) {
      return errorValue(ErrorCode.Value);
    }
    if (isError(sortOrderArg)) {
      return sortOrderArg;
    }
    if (isError(byColArg)) {
      return byColArg;
    }

    const sortByCol = byColArg === undefined ? false : toBoolean(byColArg);
    const sortOrder = sortOrderArg ? toInteger(sortOrderArg) : 1;
    const sortIndex = sortIndexArg === undefined ? 1 : toInteger(sortIndexArg);
    if (sortOrder === undefined || ![1, -1].includes(sortOrder)) {
      return errorValue(ErrorCode.Value);
    }
    if (sortIndex === undefined || sortIndex < 1) {
      return errorValue(ErrorCode.Value);
    }

    let sortError: CellValue | undefined;
    if (array.rows === 1 || array.cols === 1) {
      const values = [...array.values];
      const order: number[] = Array.from({ length: values.length }, (_, index) => index);
      order.sort((left, right) => {
        const cmp = compareScalars(values[left]!, values[right]!);
        if (cmp === undefined) {
          sortError = errorValue(ErrorCode.Value);
          return 0;
        }
        return cmp * sortOrder || left - right;
      });
      if (sortError) {
        return sortError;
      }
      const sortedValues = order.map((index) => values[index]!);
      return arrayResult(sortedValues, array.rows, array.cols);
    }
    if (sortByCol) {
      if (sortIndex > array.rows) {
        return errorValue(ErrorCode.Value);
      }
      const rowIndex = sortIndex - 1;
      const colOrder = Array.from({ length: array.cols }, (_, col) => col);
      colOrder.sort((left, right) => {
        const cmp = compareScalars(
          getRangeValue(array, rowIndex, left),
          getRangeValue(array, rowIndex, right),
        );
        if (cmp === undefined) {
          sortError = errorValue(ErrorCode.Value);
          return 0;
        }
        return cmp * sortOrder || left - right;
      });
      if (sortError) {
        return sortError;
      }
      const values: CellValue[] = [];
      for (let row = 0; row < array.rows; row += 1) {
        for (const col of colOrder) {
          values.push(getRangeValue(array, row, col));
        }
      }
      return arrayResult(values, array.rows, array.cols);
    }
    if (sortIndex > array.cols) {
      return errorValue(ErrorCode.Value);
    }
    const columnIndex = sortIndex - 1;
    const rowOrder = Array.from({ length: array.rows }, (_, row) => row);
    rowOrder.sort((left, right) => {
      const cmp = compareScalars(
        getRangeValue(array, left, columnIndex),
        getRangeValue(array, right, columnIndex),
      );
      if (cmp === undefined) {
        sortError = errorValue(ErrorCode.Value);
        return 0;
      }
      return cmp * sortOrder || left - right;
    });
    if (sortError) {
      return sortError;
    }
    const values: CellValue[] = [];
    for (const row of rowOrder) {
      values.push(...pickRangeRow(array, row));
    }
    return arrayResult(values, array.rows, array.cols);
  },
  SORTBY: (arrayArg, ...criteriaArgs) => {
    const array = toCellRange(arrayArg);
    if (!isRangeArg(array)) {
      return array;
    }
    if (array.rows > 1 && array.cols > 1) {
      return errorValue(ErrorCode.Value);
    }
    if (criteriaArgs.length === 0) {
      return errorValue(ErrorCode.Value);
    }

    const source = array.values;
    const indexes = Array.from({ length: source.length }, (_, index) => index);
    const criteria: { values: CellValue[]; order: number }[] = [];
    for (let index = 0; index < criteriaArgs.length; index += 1) {
      const criteriaArg = criteriaArgs[index]!;
      const byRange = toCellRange(criteriaArg);
      if (!isRangeArg(byRange)) {
        return byRange;
      }
      const nextArg = criteriaArgs[index + 1];
      if (nextArg !== undefined && !isRangeArg(nextArg) && !isError(nextArg)) {
        const orderValue = toInteger(nextArg);
        if (orderValue === undefined || ![1, -1].includes(orderValue)) {
          return errorValue(ErrorCode.Value);
        }
        criteria.push({ values: byRange.values, order: orderValue });
        index += 1;
        continue;
      }
      criteria.push({ values: byRange.values, order: 1 });
    }
    const expectedLength = source.length;
    if (
      criteria.some(
        (criterion) => criterion.values.length !== 1 && criterion.values.length !== expectedLength,
      )
    ) {
      return errorValue(ErrorCode.Value);
    }

    let sortError: CellValue | undefined;
    indexes.sort((left, right) => {
      if (left === right) {
        return 0;
      }
      for (const criterion of criteria) {
        const leftValue =
          criterion.values.length === 1
            ? (criterion.values[0] ?? array.values[0]!)
            : criterion.values[left]!;
        const rightValue =
          criterion.values.length === 1
            ? (criterion.values[0] ?? array.values[0]!)
            : criterion.values[right]!;
        const cmp = compareScalars(leftValue, rightValue);
        if (cmp === undefined) {
          sortError = errorValue(ErrorCode.Value);
          return 0;
        }
        if (cmp !== 0) {
          return cmp * criterion.order;
        }
      }
      return left - right;
    });
    if (sortError) {
      return sortError;
    }
    return arrayResult(
      indexes.map((index) => array.values[index] ?? { tag: ValueTag.Empty }),
      array.rows,
      array.cols,
    );
  },
  TRANSPOSE: (arrayArg) => {
    const array = toCellRange(arrayArg);
    if (!isRangeArg(array)) {
      return array;
    }
    if (array.rows === 1 && array.cols === 1) {
      return array.values[0] ?? { tag: ValueTag.Empty };
    }
    const values: CellValue[] = [];
    for (let col = 0; col < array.cols; col += 1) {
      for (let row = 0; row < array.rows; row += 1) {
        values.push(getRangeValue(array, row, col));
      }
    }
    return arrayResult(values, array.cols, array.rows);
  },
  HSTACK: (...arrayArgs) => {
    if (arrayArgs.length === 0) {
      return errorValue(ErrorCode.Value);
    }
    const arrays = arrayArgs.map(toCellRange);
    const rangeError = findFirstNonRange(arrays);
    if (rangeError) {
      return rangeError;
    }
    if (!areRangeArgs(arrays)) {
      return errorValue(ErrorCode.Value);
    }

    const rowCount = Math.max(...arrays.map((array) => array.rows));
    for (const array of arrays) {
      if (array.rows !== 1 && array.rows !== rowCount) {
        return errorValue(ErrorCode.Value);
      }
    }

    const values: CellValue[] = [];
    const totalCols = arrays.reduce((acc, array) => acc + array.cols, 0);
    for (let row = 0; row < rowCount; row += 1) {
      for (const array of arrays) {
        for (let col = 0; col < array.cols; col += 1) {
          const sourceRow = array.rows === 1 ? 0 : row;
          values.push(getRangeValue(array, sourceRow, col));
        }
      }
    }
    return arrayResult(values, rowCount, totalCols);
  },
  VSTACK: (...arrayArgs) => {
    if (arrayArgs.length === 0) {
      return errorValue(ErrorCode.Value);
    }
    const arrays = arrayArgs.map(toCellRange);
    const rangeError = findFirstNonRange(arrays);
    if (rangeError) {
      return rangeError;
    }
    if (!areRangeArgs(arrays)) {
      return errorValue(ErrorCode.Value);
    }

    const colCount = Math.max(...arrays.map((array) => array.cols));
    for (const array of arrays) {
      if (array.cols !== 1 && array.cols !== colCount) {
        return errorValue(ErrorCode.Value);
      }
    }

    const values: CellValue[] = [];
    const totalRows = arrays.reduce((acc, array) => acc + array.rows, 0);
    for (const array of arrays) {
      for (let row = 0; row < array.rows; row += 1) {
        for (let col = 0; col < colCount; col += 1) {
          const sourceCol = array.cols === 1 ? 0 : col;
          values.push(getRangeValue(array, row, sourceCol));
        }
      }
    }
    return arrayResult(values, totalRows, colCount);
  },
  TOCOL: (arrayArg, ignoreArg = { tag: ValueTag.Number, value: 0 }, scanByColArg) => {
    const array = toCellRange(arrayArg);
    if (!isRangeArg(array)) {
      return array;
    }
    if (isRangeArg(ignoreArg) || isRangeArg(scanByColArg)) {
      return errorValue(ErrorCode.Value);
    }
    if (isError(ignoreArg) || isError(scanByColArg)) {
      return isError(ignoreArg) ? ignoreArg : scanByColArg;
    }
    const ignoreValue = ignoreArg === undefined ? 0 : toInteger(ignoreArg);
    if (ignoreValue === undefined || ![0, 1].includes(ignoreValue)) {
      return errorValue(ErrorCode.Value);
    }
    const scanByCol = scanByColArg === undefined ? true : toBoolean(scanByColArg);
    if (scanByCol === undefined) {
      return errorValue(ErrorCode.Value);
    }
    const values = flattenValues(array, scanByCol, ignoreValue === 1);
    return arrayResult(values, values.length, 1);
  },
  TOROW: (arrayArg, ignoreArg = { tag: ValueTag.Number, value: 0 }, scanByColArg) => {
    const array = toCellRange(arrayArg);
    if (!isRangeArg(array)) {
      return array;
    }
    if (isRangeArg(ignoreArg) || isRangeArg(scanByColArg)) {
      return errorValue(ErrorCode.Value);
    }
    if (isError(ignoreArg) || isError(scanByColArg)) {
      return isError(ignoreArg) ? ignoreArg : scanByColArg;
    }
    const ignoreValue = ignoreArg === undefined ? 0 : toInteger(ignoreArg);
    if (ignoreValue === undefined || ![0, 1].includes(ignoreValue)) {
      return errorValue(ErrorCode.Value);
    }
    const scanByCol = scanByColArg === undefined ? false : toBoolean(scanByColArg);
    if (scanByCol === undefined) {
      return errorValue(ErrorCode.Value);
    }
    const values = flattenValues(array, scanByCol, ignoreValue === 1);
    return arrayResult(values, 1, values.length);
  },
  WRAPROWS: (arrayArg, wrapCountArg, padWithArg, padByColArg) => {
    const array = toCellRange(arrayArg);
    if (!isRangeArg(array)) {
      return array;
    }
    if (isRangeArg(wrapCountArg) || isRangeArg(padWithArg) || isRangeArg(padByColArg)) {
      return errorValue(ErrorCode.Value);
    }
    if (isError(wrapCountArg) || isError(padByColArg)) {
      return isError(wrapCountArg) ? wrapCountArg : padByColArg;
    }
    if (padWithArg !== undefined && isError(padWithArg)) {
      return padWithArg;
    }
    const wrapCount = toInteger(wrapCountArg);
    if (wrapCount === undefined || wrapCount < 1) {
      return errorValue(ErrorCode.Value);
    }
    if (padByColArg !== undefined && toBoolean(padByColArg) === undefined) {
      return errorValue(ErrorCode.Value);
    }

    const values = array.values.slice();
    const rows = Math.ceil(values.length / wrapCount);
    const cols = wrapCount;
    const padValue: CellValue = padWithArg === undefined ? errorValue(ErrorCode.NA) : padWithArg;
    while (values.length < rows * cols) {
      values.push(padValue);
    }
    return arrayResult(values, rows, cols);
  },
  WRAPCOLS: (arrayArg, wrapCountArg, padWithArg, padByColArg) => {
    const array = toCellRange(arrayArg);
    if (!isRangeArg(array)) {
      return array;
    }
    if (isRangeArg(wrapCountArg) || isRangeArg(padWithArg) || isRangeArg(padByColArg)) {
      return errorValue(ErrorCode.Value);
    }
    if (isError(wrapCountArg) || isError(padByColArg)) {
      return isError(wrapCountArg) ? wrapCountArg : padByColArg;
    }
    if (padWithArg !== undefined && isError(padWithArg)) {
      return padWithArg;
    }
    const wrapCount = toInteger(wrapCountArg);
    if (wrapCount === undefined || wrapCount < 1) {
      return errorValue(ErrorCode.Value);
    }
    if (padByColArg !== undefined && toBoolean(padByColArg) === undefined) {
      return errorValue(ErrorCode.Value);
    }

    const values = array.values.slice();
    const rows = wrapCount;
    const cols = Math.ceil(values.length / rows);
    const padValue: CellValue = padWithArg === undefined ? errorValue(ErrorCode.NA) : padWithArg;
    const paddedValues = Array.from(
      { length: rows * cols },
      (_, index) => values[index] ?? padValue,
    );
    const wrappedValues: CellValue[] = [];
    for (let row = 0; row < rows; row += 1) {
      for (let col = 0; col < cols; col += 1) {
        wrappedValues.push(paddedValues[col * rows + row] ?? padValue);
      }
    }
    return arrayResult(wrappedValues, rows, cols);
  },
  COUNTIF: (rangeArg, criteriaArg) => {
    const range = requireCellRange(rangeArg);
    if (!isRangeArg(range)) {
      return range;
    }
    if (isRangeArg(criteriaArg)) {
      return errorValue(ErrorCode.Value);
    }
    if (isError(criteriaArg)) {
      return criteriaArg;
    }
    let count = 0;
    for (const value of range.values) {
      if (matchesCriteria(value, criteriaArg)) {
        count += 1;
      }
    }
    return { tag: ValueTag.Number, value: count };
  },
  COUNTIFS: (...args) => {
    if (args.length === 0 || args.length % 2 !== 0) {
      return errorValue(ErrorCode.Value);
    }
    const rangeCriteriaPairs: { range: RangeBuiltinArgument; criteria: CellValue }[] = [];
    for (let index = 0; index < args.length; index += 2) {
      const range = requireCellRange(args[index]!);
      if (!isRangeArg(range)) {
        return range;
      }
      const criteria = args[index + 1]!;
      if (isRangeArg(criteria)) {
        return errorValue(ErrorCode.Value);
      }
      if (isError(criteria)) {
        return criteria;
      }
      rangeCriteriaPairs.push({ range, criteria });
    }
    const expectedLength = rangeCriteriaPairs[0]!.range.values.length;
    if (rangeCriteriaPairs.some((pair) => pair.range.values.length !== expectedLength)) {
      return errorValue(ErrorCode.Value);
    }

    let count = 0;
    for (let row = 0; row < expectedLength; row += 1) {
      if (
        rangeCriteriaPairs.every((pair) => matchesCriteria(pair.range.values[row]!, pair.criteria))
      ) {
        count += 1;
      }
    }
    return { tag: ValueTag.Number, value: count };
  },
  SUMIF: (rangeArg, criteriaArg, sumRangeArg = rangeArg) => {
    const range = requireCellRange(rangeArg);
    const sumRange = requireCellRange(sumRangeArg);
    if (!isRangeArg(range)) {
      return range;
    }
    if (!isRangeArg(sumRange)) {
      return sumRange;
    }
    if (range.values.length !== sumRange.values.length) {
      return errorValue(ErrorCode.Value);
    }
    if (isRangeArg(criteriaArg)) {
      return errorValue(ErrorCode.Value);
    }
    if (isError(criteriaArg)) {
      return criteriaArg;
    }

    let sum = 0;
    for (let index = 0; index < range.values.length; index += 1) {
      if (!matchesCriteria(range.values[index]!, criteriaArg)) {
        continue;
      }
      sum += toNumber(sumRange.values[index]!) ?? 0;
    }
    return { tag: ValueTag.Number, value: sum };
  },
  SUMIFS: (sumRangeArg, ...criteriaArgs) => {
    const sumRange = requireCellRange(sumRangeArg);
    if (!isRangeArg(sumRange)) {
      return sumRange;
    }
    const matchingRows = findMatchingRowIndexes(sumRange, criteriaArgs);
    if (!Array.isArray(matchingRows)) {
      return matchingRows;
    }

    let sum = 0;
    for (const row of matchingRows) {
      sum += toNumber(sumRange.values[row]!) ?? 0;
    }
    return { tag: ValueTag.Number, value: sum };
  },
  AVERAGEIF: (rangeArg, criteriaArg, averageRangeArg = rangeArg) => {
    const range = requireCellRange(rangeArg);
    const averageRange = requireCellRange(averageRangeArg);
    if (!isRangeArg(range)) {
      return range;
    }
    if (!isRangeArg(averageRange)) {
      return averageRange;
    }
    if (range.values.length !== averageRange.values.length) {
      return errorValue(ErrorCode.Value);
    }
    if (isRangeArg(criteriaArg)) {
      return errorValue(ErrorCode.Value);
    }
    if (isError(criteriaArg)) {
      return criteriaArg;
    }

    let count = 0;
    let sum = 0;
    for (let index = 0; index < range.values.length; index += 1) {
      if (!matchesCriteria(range.values[index]!, criteriaArg)) {
        continue;
      }
      const numeric = toNumber(averageRange.values[index]!);
      if (numeric === undefined) {
        continue;
      }
      count += 1;
      sum += numeric;
    }

    if (count === 0) {
      return errorValue(ErrorCode.Div0);
    }
    return { tag: ValueTag.Number, value: sum / count };
  },
  AVERAGEIFS: (averageRangeArg, ...criteriaArgs) => {
    const averageRange = requireCellRange(averageRangeArg);
    if (!isRangeArg(averageRange)) {
      return averageRange;
    }
    const matchingRows = findMatchingRowIndexes(averageRange, criteriaArgs);
    if (!Array.isArray(matchingRows)) {
      return matchingRows;
    }

    let count = 0;
    let sum = 0;
    for (const row of matchingRows) {
      const numeric = toNumber(averageRange.values[row]!);
      if (numeric === undefined) {
        continue;
      }
      count += 1;
      sum += numeric;
    }
    if (count === 0) {
      return errorValue(ErrorCode.Div0);
    }
    return { tag: ValueTag.Number, value: sum / count };
  },
  MINIFS: (minRangeArg, ...criteriaArgs) => {
    const minRange = requireCellRange(minRangeArg);
    if (!isRangeArg(minRange)) {
      return minRange;
    }
    const matchingRows = findMatchingRowIndexes(minRange, criteriaArgs);
    if (!Array.isArray(matchingRows)) {
      return matchingRows;
    }

    let minimum = Number.POSITIVE_INFINITY;
    for (const row of matchingRows) {
      const numeric = numericAggregateCandidate(minRange.values[row]!);
      if (numeric === undefined) {
        continue;
      }
      minimum = Math.min(minimum, numeric);
    }
    return { tag: ValueTag.Number, value: minimum === Number.POSITIVE_INFINITY ? 0 : minimum };
  },
  MAXIFS: (maxRangeArg, ...criteriaArgs) => {
    const maxRange = requireCellRange(maxRangeArg);
    if (!isRangeArg(maxRange)) {
      return maxRange;
    }
    const matchingRows = findMatchingRowIndexes(maxRange, criteriaArgs);
    if (!Array.isArray(matchingRows)) {
      return matchingRows;
    }

    let maximum = Number.NEGATIVE_INFINITY;
    for (const row of matchingRows) {
      const numeric = numericAggregateCandidate(maxRange.values[row]!);
      if (numeric === undefined) {
        continue;
      }
      maximum = Math.max(maximum, numeric);
    }
    return { tag: ValueTag.Number, value: maximum === Number.NEGATIVE_INFINITY ? 0 : maximum };
  },
  SUMPRODUCT: (...args) => {
    if (args.length === 0) {
      return errorValue(ErrorCode.Value);
    }
    const ranges = args.map((arg) => requireCellRange(arg));
    const rangeError = findFirstNonRange(ranges);
    if (rangeError) {
      return rangeError;
    }
    if (!areRangeArgs(ranges)) {
      return errorValue(ErrorCode.Value);
    }
    const typedRanges = ranges;
    const expectedLength = typedRanges[0]!.values.length;
    if (typedRanges.some((range) => range.values.length !== expectedLength)) {
      return errorValue(ErrorCode.Value);
    }
    let sum = 0;
    for (let index = 0; index < expectedLength; index += 1) {
      let product = 1;
      for (const range of typedRanges) {
        product *= toNumber(range.values[index]!) ?? 0;
      }
      sum += product;
    }
    return { tag: ValueTag.Number, value: sum };
  },
  SUMX2MY2: (xArg, yArg) => {
    const xValues = flattenNumbers(xArg);
    const yValues = flattenNumbers(yArg);
    if (!Array.isArray(xValues)) {
      return xValues;
    }
    if (!Array.isArray(yValues)) {
      return yValues;
    }
    if (xValues.length !== yValues.length) {
      return errorValue(ErrorCode.Value);
    }
    let sum = 0;
    for (let index = 0; index < xValues.length; index += 1) {
      sum += xValues[index]! ** 2 - yValues[index]! ** 2;
    }
    return { tag: ValueTag.Number, value: sum };
  },
  SUMX2PY2: (xArg, yArg) => {
    const xValues = flattenNumbers(xArg);
    const yValues = flattenNumbers(yArg);
    if (!Array.isArray(xValues)) {
      return xValues;
    }
    if (!Array.isArray(yValues)) {
      return yValues;
    }
    if (xValues.length !== yValues.length) {
      return errorValue(ErrorCode.Value);
    }
    let sum = 0;
    for (let index = 0; index < xValues.length; index += 1) {
      sum += xValues[index]! ** 2 + yValues[index]! ** 2;
    }
    return { tag: ValueTag.Number, value: sum };
  },
  SUMXMY2: (xArg, yArg) => {
    const xValues = flattenNumbers(xArg);
    const yValues = flattenNumbers(yArg);
    if (!Array.isArray(xValues)) {
      return xValues;
    }
    if (!Array.isArray(yValues)) {
      return yValues;
    }
    if (xValues.length !== yValues.length) {
      return errorValue(ErrorCode.Value);
    }
    let sum = 0;
    for (let index = 0; index < xValues.length; index += 1) {
      sum += (xValues[index]! - yValues[index]!) ** 2;
    }
    return { tag: ValueTag.Number, value: sum };
  },
  MDETERM: (matrixArg) => {
    const matrix = toNumericMatrix(matrixArg);
    if (!Array.isArray(matrix)) {
      return matrix;
    }
    if (matrix.length === 0 || matrix.some((row) => row.length !== matrix.length)) {
      return errorValue(ErrorCode.Value);
    }
    return { tag: ValueTag.Number, value: determinantOf(matrix) };
  },
  MINVERSE: (matrixArg) => {
    const matrix = toNumericMatrix(matrixArg);
    if (!Array.isArray(matrix)) {
      return matrix;
    }
    if (matrix.length === 0 || matrix.some((row) => row.length !== matrix.length)) {
      return errorValue(ErrorCode.Value);
    }
    const inverse = inverseOf(matrix);
    if (!inverse) {
      return errorValue(ErrorCode.Value);
    }
    return arrayResult(
      inverse.flat().map((value) => ({ tag: ValueTag.Number, value })),
      matrix.length,
      matrix.length,
    );
  },
  MMULT: (leftArg, rightArg) => {
    const left = toNumericMatrix(leftArg);
    const right = toNumericMatrix(rightArg);
    if (!Array.isArray(left)) {
      return left;
    }
    if (!Array.isArray(right)) {
      return right;
    }
    if (left.length === 0 || right.length === 0 || left[0]!.length !== right.length) {
      return errorValue(ErrorCode.Value);
    }
    const rows = left.length;
    const cols = right[0]!.length;
    const inner = right.length;
    const values: CellValue[] = [];
    for (let row = 0; row < rows; row += 1) {
      for (let col = 0; col < cols; col += 1) {
        let sum = 0;
        for (let index = 0; index < inner; index += 1) {
          sum += left[row]![index]! * right[index]![col]!;
        }
        values.push({ tag: ValueTag.Number, value: sum });
      }
    }
    return arrayResult(values, rows, cols);
  },
  PERCENTOF: (subsetArg, totalArg) => {
    const subset = sumOfNumbers(subsetArg);
    const total = sumOfNumbers(totalArg);
    if (typeof subset !== "number") {
      return subset;
    }
    if (typeof total !== "number") {
      return total;
    }
    if (total === 0) {
      return errorValue(ErrorCode.Div0);
    }
    return { tag: ValueTag.Number, value: subset / total };
  },
  FILTER: (arrayArg, includeArg, ifEmptyArg = { tag: ValueTag.Error, code: ErrorCode.Value }) => {
    const array = requireCellRange(arrayArg);
    const include = requireCellRange(includeArg);
    if (!isRangeArg(array)) {
      return array;
    }
    if (!isRangeArg(include)) {
      return include;
    }
    if (include.rows === array.rows && include.cols === 1) {
      const values: CellValue[] = [];
      let keptRows = 0;
      for (let row = 0; row < array.rows; row += 1) {
        const includeValue = getRangeValue(include, row, 0);
        if (isError(includeValue)) {
          return includeValue;
        }
        const keep = toBoolean(includeValue);
        if (keep === undefined) {
          return errorValue(ErrorCode.Value);
        }
        if (!keep) {
          continue;
        }
        values.push(...pickRangeRow(array, row));
        keptRows += 1;
      }
      if (keptRows === 0) {
        return isRangeArg(ifEmptyArg) ? errorValue(ErrorCode.Value) : ifEmptyArg;
      }
      return arrayResult(values, keptRows, array.cols);
    }

    if (include.cols === array.cols && include.rows === 1) {
      const keptCols: number[] = [];
      for (let col = 0; col < array.cols; col += 1) {
        const includeValue = getRangeValue(include, 0, col);
        if (isError(includeValue)) {
          return includeValue;
        }
        const keep = toBoolean(includeValue);
        if (keep === undefined) {
          return errorValue(ErrorCode.Value);
        }
        if (keep) {
          keptCols.push(col);
        }
      }
      if (keptCols.length === 0) {
        return isRangeArg(ifEmptyArg) ? errorValue(ErrorCode.Value) : ifEmptyArg;
      }
      const values: CellValue[] = [];
      for (let row = 0; row < array.rows; row += 1) {
        for (const col of keptCols) {
          values.push(getRangeValue(array, row, col));
        }
      }
      return arrayResult(values, array.rows, keptCols.length);
    }

    return errorValue(ErrorCode.Value);
  },
  UNIQUE: (
    arrayArg,
    byColArg = { tag: ValueTag.Boolean, value: false },
    exactlyOnceArg = { tag: ValueTag.Boolean, value: false },
  ) => {
    const array = requireCellRange(arrayArg);
    if (!isRangeArg(array)) {
      return array;
    }
    if (isRangeArg(byColArg) || isRangeArg(exactlyOnceArg)) {
      return errorValue(ErrorCode.Value);
    }
    if (isError(byColArg)) {
      return byColArg;
    }
    if (isError(exactlyOnceArg)) {
      return exactlyOnceArg;
    }
    const byCol = toBoolean(byColArg);
    const exactlyOnce = toBoolean(exactlyOnceArg);
    if (byCol === undefined || exactlyOnce === undefined) {
      return errorValue(ErrorCode.Value);
    }

    if (array.rows === 1 || array.cols === 1) {
      const counts = new Map<string, number>();
      const keys: string[] = [];
      for (const value of array.values) {
        if (isError(value)) {
          return value;
        }
        const key = JSON.stringify(
          value.tag === ValueTag.String ? { ...value, value: value.value.toUpperCase() } : value,
        );
        keys.push(key);
        counts.set(key, (counts.get(key) ?? 0) + 1);
      }
      const values: CellValue[] = [];
      const seen = new Set<string>();
      for (let index = 0; index < array.values.length; index += 1) {
        const key = keys[index]!;
        if (seen.has(key)) {
          continue;
        }
        seen.add(key);
        if (exactlyOnce && counts.get(key) !== 1) {
          continue;
        }
        values.push(array.values[index]!);
      }
      return array.rows === 1
        ? arrayResult(values, 1, values.length)
        : arrayResult(values, values.length, 1);
    }

    if (byCol) {
      const counts = new Map<string, number>();
      const keys: string[] = [];
      for (let col = 0; col < array.cols; col += 1) {
        const key = colKey(array, col);
        if (key === undefined) {
          return errorValue(ErrorCode.Value);
        }
        keys.push(key);
        counts.set(key, (counts.get(key) ?? 0) + 1);
      }
      const keptCols: number[] = [];
      const seen = new Set<string>();
      for (let col = 0; col < array.cols; col += 1) {
        const key = keys[col]!;
        if (seen.has(key)) {
          continue;
        }
        seen.add(key);
        if (exactlyOnce && counts.get(key) !== 1) {
          continue;
        }
        keptCols.push(col);
      }
      const values: CellValue[] = [];
      for (let row = 0; row < array.rows; row += 1) {
        for (const col of keptCols) {
          values.push(getRangeValue(array, row, col));
        }
      }
      return arrayResult(values, array.rows, keptCols.length);
    }

    const counts = new Map<string, number>();
    const keys: string[] = [];
    for (let row = 0; row < array.rows; row += 1) {
      const key = rowKey(array, row);
      if (key === undefined) {
        return errorValue(ErrorCode.Value);
      }
      keys.push(key);
      counts.set(key, (counts.get(key) ?? 0) + 1);
    }
    const keptRows: number[] = [];
    const seen = new Set<string>();
    for (let row = 0; row < array.rows; row += 1) {
      const key = keys[row]!;
      if (seen.has(key)) {
        continue;
      }
      seen.add(key);
      if (exactlyOnce && counts.get(key) !== 1) {
        continue;
      }
      keptRows.push(row);
    }
    const values: CellValue[] = [];
    for (const row of keptRows) {
      values.push(...pickRangeRow(array, row));
    }
    return arrayResult(values, keptRows.length, array.cols);
  },
  ...externalLookupBuiltins,
};

type CriteriaOperator = "=" | "<>" | ">" | ">=" | "<" | "<=";

function matchesCriteria(value: CellValue, criteria: CellValue): boolean {
  if (isError(value)) {
    return false;
  }
  const { operator, operand } = parseCriteria(criteria);
  const comparison = compareScalars(value, operand);
  if (comparison === undefined) {
    return false;
  }
  switch (operator) {
    case "=":
      return comparison === 0;
    case "<>":
      return comparison !== 0;
    case ">":
      return comparison > 0;
    case ">=":
      return comparison >= 0;
    case "<":
      return comparison < 0;
    case "<=":
      return comparison <= 0;
  }
}

function parseCriteria(criteria: CellValue): { operator: CriteriaOperator; operand: CellValue } {
  if (criteria.tag !== ValueTag.String) {
    return { operator: "=", operand: criteria };
  }

  const match = /^(<=|>=|<>|=|<|>)(.*)$/.exec(criteria.value);
  if (!match) {
    return { operator: "=", operand: criteria };
  }

  const operator = match[1] ?? "=";
  return {
    operator: isCriteriaOperator(operator) ? operator : "=",
    operand: parseCriteriaOperand(match[2] ?? ""),
  };
}

function parseCriteriaOperand(raw: string): CellValue {
  const trimmed = raw.trim();
  if (trimmed === "") {
    return { tag: ValueTag.String, value: "", stringId: 0 };
  }
  const upper = trimmed.toUpperCase();
  if (upper === "TRUE" || upper === "FALSE") {
    return { tag: ValueTag.Boolean, value: upper === "TRUE" };
  }
  const numeric = Number(trimmed);
  if (Number.isFinite(numeric)) {
    return { tag: ValueTag.Number, value: numeric };
  }
  return { tag: ValueTag.String, value: trimmed, stringId: 0 };
}

export function getLookupBuiltin(name: string): LookupBuiltin | undefined {
  const upper = name.toUpperCase();
  if (upper === "USE.THE.COUNTIF") {
    return lookupBuiltins["COUNTIF"];
  }
  return lookupBuiltins[upper] ?? getExternalLookupFunction(name);
}
