import { ErrorCode, ValueTag, type CellValue } from "@bilig/protocol";
import { getExternalLookupFunction } from "../external-function-adapter.js";
import type { ArrayValue, EvaluationResult } from "../runtime-values.js";

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

function arrayResult(values: CellValue[], rows: number, cols: number): ArrayValue {
  return { kind: "array", values, rows, cols };
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

function flattenNumericArguments(args: readonly LookupBuiltinArgument[]): number[] | CellValue {
  const values: number[] = [];
  for (const arg of args) {
    const flattened = flattenNumbers(arg);
    if (!Array.isArray(flattened)) {
      return flattened;
    }
    values.push(...flattened);
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

function parseCorrelationOperands(
  firstArg: LookupBuiltinArgument,
  secondArg: LookupBuiltinArgument,
): { first: number[]; second: number[] } | CellValue {
  const first = flattenNumbers(firstArg);
  if (!Array.isArray(first)) {
    return first;
  }
  const second = flattenNumbers(secondArg);
  if (!Array.isArray(second)) {
    return second;
  }
  if (first.length === 0 || first.length !== second.length) {
    return errorValue(ErrorCode.Value);
  }
  return { first, second };
}

function covarianceFromPairs(
  first: readonly number[],
  second: readonly number[],
  useSample: boolean,
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

function rankFromValues(
  numberArg: LookupBuiltinArgument,
  arrayArg: LookupBuiltinArgument,
  orderArg: LookupBuiltinArgument | undefined,
  useAverage: boolean,
): CellValue {
  if (isRangeArg(numberArg) || isRangeArg(orderArg)) {
    return errorValue(ErrorCode.Value);
  }
  if (isError(numberArg)) {
    return numberArg;
  }
  if (isError(arrayArg)) {
    return arrayArg;
  }
  if (isError(orderArg)) {
    return orderArg;
  }

  const target = toNumber(numberArg);
  if (target === undefined) {
    return errorValue(ErrorCode.Value);
  }
  const order = orderArg === undefined ? 0 : toInteger(orderArg);
  if (order === undefined || ![0, 1].includes(order)) {
    return errorValue(ErrorCode.Value);
  }

  const values = flattenNumbers(arrayArg);
  if (!Array.isArray(values)) {
    return values;
  }
  if (values.length === 0) {
    return errorValue(ErrorCode.NA);
  }

  let preceding = 0;
  let ties = 0;
  for (const value of values) {
    if (value === target) {
      ties += 1;
      continue;
    }
    if (order === 0 ? value > target : value < target) {
      preceding += 1;
    }
  }

  if (ties === 0) {
    return errorValue(ErrorCode.NA);
  }

  return {
    tag: ValueTag.Number,
    value: useAverage ? preceding + (ties + 1) / 2 : preceding + 1,
  };
}

function nthValue(
  arg: LookupBuiltinArgument,
  positionArg: LookupBuiltinArgument,
  ascending: boolean,
): CellValue {
  if (isRangeArg(positionArg)) {
    return errorValue(ErrorCode.Value);
  }
  if (isError(positionArg)) {
    return positionArg;
  }
  const values = flattenNumbers(arg);
  if (!Array.isArray(values)) {
    return values;
  }
  if (values.length === 0) {
    return errorValue(ErrorCode.Value);
  }

  const position = toInteger(positionArg);
  if (position === undefined || position < 1 || position > values.length) {
    return errorValue(ErrorCode.Value);
  }

  const sortedValues = values.toSorted(ascending ? (a, b) => a - b : (a, b) => b - a);
  return { tag: ValueTag.Number, value: sortedValues[position - 1] ?? 0 };
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

export const lookupBuiltins: Record<string, LookupBuiltin> = {
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
  CORREL: (firstArg, secondArg) => {
    const values = parseCorrelationOperands(firstArg, secondArg);
    if (!("first" in values)) {
      return values;
    }
    const correlation = correlationFromPairs(values.first, values.second);
    return typeof correlation === "number"
      ? { tag: ValueTag.Number, value: correlation }
      : correlation;
  },
  COVAR: (firstArg, secondArg) => {
    const values = parseCorrelationOperands(firstArg, secondArg);
    if (!("first" in values)) {
      return values;
    }
    const covariance = covarianceFromPairs(values.first, values.second, false);
    return typeof covariance === "number"
      ? { tag: ValueTag.Number, value: covariance }
      : covariance;
  },
  PEARSON: (firstArg, secondArg) => {
    const values = parseCorrelationOperands(firstArg, secondArg);
    if (!("first" in values)) {
      return values;
    }
    const correlation = correlationFromPairs(values.first, values.second);
    return typeof correlation === "number"
      ? { tag: ValueTag.Number, value: correlation }
      : correlation;
  },
  "COVARIANCE.P": (firstArg, secondArg) => {
    const values = parseCorrelationOperands(firstArg, secondArg);
    if (!("first" in values)) {
      return values;
    }
    const covariance = covarianceFromPairs(values.first, values.second, false);
    return typeof covariance === "number"
      ? { tag: ValueTag.Number, value: covariance }
      : covariance;
  },
  "COVARIANCE.S": (firstArg, secondArg) => {
    const values = parseCorrelationOperands(firstArg, secondArg);
    if (!("first" in values)) {
      return values;
    }
    const covariance = covarianceFromPairs(values.first, values.second, true);
    return typeof covariance === "number"
      ? { tag: ValueTag.Number, value: covariance }
      : covariance;
  },
  AVEDEV: (...args) => {
    const values = flattenNumericArguments(args);
    if (!Array.isArray(values)) {
      return values;
    }
    if (values.length === 0) {
      return errorValue(ErrorCode.Value);
    }
    const mean = values.reduce((sum, value) => sum + value, 0) / values.length;
    const totalAbsoluteDeviation = values.reduce((sum, value) => sum + Math.abs(value - mean), 0);
    return { tag: ValueTag.Number, value: totalAbsoluteDeviation / values.length };
  },
  DEVSQ: (...args) => {
    const values = flattenNumericArguments(args);
    if (!Array.isArray(values)) {
      return values;
    }
    if (values.length === 0) {
      return errorValue(ErrorCode.Value);
    }
    const mean = values.reduce((sum, value) => sum + value, 0) / values.length;
    const total = values.reduce((sum, value) => {
      const deviation = value - mean;
      return sum + deviation * deviation;
    }, 0);
    return { tag: ValueTag.Number, value: total };
  },
  MEDIAN: (...args) => {
    const values = flattenNumericArguments(args);
    if (!Array.isArray(values)) {
      return values;
    }
    if (values.length === 0) {
      return errorValue(ErrorCode.Value);
    }

    const sortedValues = values.toSorted((left, right) => left - right);
    const center = Math.floor(sortedValues.length / 2);
    const isEven = sortedValues.length % 2 === 0;
    if (!isEven) {
      return { tag: ValueTag.Number, value: sortedValues[center] ?? 0 };
    }

    const lower = sortedValues[center - 1];
    const upper = sortedValues[center];
    return { tag: ValueTag.Number, value: ((lower ?? 0) + (upper ?? 0)) / 2 };
  },
  SMALL: (arg, positionArg) => nthValue(arg, positionArg, true),
  LARGE: (arg, positionArg) => nthValue(arg, positionArg, false),
  RANK: (numberArg, arrayArg, orderArg = { tag: ValueTag.Number, value: 0 }) =>
    rankFromValues(numberArg, arrayArg, orderArg, false),
  "RANK.EQ": (numberArg, arrayArg, orderArg = { tag: ValueTag.Number, value: 0 }) =>
    rankFromValues(numberArg, arrayArg, orderArg, false),
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
      const colIndex = sortIndex - 1;
      const rowOrder = Array.from({ length: array.rows }, (_, row) => row);
      rowOrder.sort((left, right) => {
        const cmp = compareScalars(
          getRangeValue(array, left, colIndex),
          getRangeValue(array, right, colIndex),
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
    }
    if (sortIndex > array.cols) {
      return errorValue(ErrorCode.Value);
    }
    const columnIndex = sortIndex - 1;
    const colOrder = Array.from({ length: array.cols }, (_, col) => col);
    colOrder.sort((left, right) => {
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
    for (let row = 0; row < array.rows; row += 1) {
      for (const col of colOrder) {
        values.push(getRangeValue(array, row, col));
      }
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
    if (rangeCriteriaPairs.some((pair) => pair.range.values.length !== sumRange.values.length)) {
      return errorValue(ErrorCode.Value);
    }

    let sum = 0;
    for (let row = 0; row < sumRange.values.length; row += 1) {
      if (
        !rangeCriteriaPairs.every((pair) => matchesCriteria(pair.range.values[row]!, pair.criteria))
      ) {
        continue;
      }
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
    if (
      rangeCriteriaPairs.some((pair) => pair.range.values.length !== averageRange.values.length)
    ) {
      return errorValue(ErrorCode.Value);
    }

    let count = 0;
    let sum = 0;
    for (let row = 0; row < averageRange.values.length; row += 1) {
      if (
        !rangeCriteriaPairs.every((pair) => matchesCriteria(pair.range.values[row]!, pair.criteria))
      ) {
        continue;
      }
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
  return lookupBuiltins[name.toUpperCase()] ?? getExternalLookupFunction(name);
}
