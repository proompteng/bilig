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

function isError(value: CellValue): value is Extract<CellValue, { tag: ValueTag.Error }> {
  return value.tag === ValueTag.Error;
}

function isRangeArg(value: LookupBuiltinArgument): value is RangeBuiltinArgument {
  return typeof value === "object" && value !== null && "kind" in value && value.kind === "range";
}

function isCriteriaOperator(value: string): value is CriteriaOperator {
  return value === "="
    || value === "<>"
    || value === ">"
    || value === ">="
    || value === "<"
    || value === "<=";
}

function findFirstNonRange(values: readonly (RangeBuiltinArgument | CellValue)[]): CellValue | undefined {
  for (const value of values) {
    if (!isRangeArg(value)) {
      return value;
    }
  }
  return undefined;
}

function areRangeArgs(values: readonly (RangeBuiltinArgument | CellValue)[]): values is RangeBuiltinArgument[] {
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
  if ((left.tag === ValueTag.String || left.tag === ValueTag.Empty) && (right.tag === ValueTag.String || right.tag === ValueTag.Empty)) {
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
    ...Array.from({ length: size }, (_, colIndex) => (rowIndex === colIndex ? 1 : 0))
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
    stringId: value.stringId
  };
}

function rowKey(range: RangeBuiltinArgument, row: number): string | undefined {
  const values = pickRangeRow(range, row);
  if (values.some(isError)) {
    return undefined;
  }
  return JSON.stringify(values.map(normalizeKeyValue));
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
  VLOOKUP: (lookupValue, tableArray, colIndexValue, rangeLookupValue = { tag: ValueTag.Boolean, value: true }) => {
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
    if (colIndex === undefined || colIndex < 1 || colIndex > tableArray.cols || rangeLookup === undefined) {
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
  HLOOKUP: (lookupValue, tableArray, rowIndexValue, rangeLookupValue = { tag: ValueTag.Boolean, value: true }) => {
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
    if (rowIndex === undefined || rowIndex < 1 || rowIndex > tableArray.rows || rangeLookup === undefined) {
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
    searchMode = { tag: ValueTag.Number, value: 1 }
  ) => {
    if (isRangeArg(lookupValue) || isRangeArg(ifNotFound) || isRangeArg(matchMode) || isRangeArg(searchMode)) {
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
    searchModeValue = { tag: ValueTag.Number, value: 1 }
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

    const values = searchMode === -1 ? [...rangeOrError.values].reverse() : rangeOrError.values;
    const probe =
      searchMode === -1
        ? { ...rangeOrError, values }
        : rangeOrError;
    const position =
      matchMode === 0
        ? exactMatch(lookupValue, probe)
        : matchMode === 1
          ? approximateMatchAscending(lookupValue, probe)
          : approximateMatchDescending(lookupValue, probe);
    if (position === -1) {
      return errorValue(ErrorCode.NA);
    }
    const normalizedPosition = searchMode === -1 ? rangeOrError.values.length - position + 1 : position;
    return { tag: ValueTag.Number, value: normalizedPosition };
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
      if (rangeCriteriaPairs.every((pair) => matchesCriteria(pair.range.values[row]!, pair.criteria))) {
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
    if (
      rangeCriteriaPairs.some((pair) => pair.range.values.length !== sumRange.values.length)
    ) {
      return errorValue(ErrorCode.Value);
    }

    let sum = 0;
    for (let row = 0; row < sumRange.values.length; row += 1) {
      if (!rangeCriteriaPairs.every((pair) => matchesCriteria(pair.range.values[row]!, pair.criteria))) {
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
      if (!rangeCriteriaPairs.every((pair) => matchesCriteria(pair.range.values[row]!, pair.criteria))) {
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
    return arrayResult(inverse.flat().map((value) => ({ tag: ValueTag.Number, value })), matrix.length, matrix.length);
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
    exactlyOnceArg = { tag: ValueTag.Boolean, value: false }
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
        const key = JSON.stringify(value.tag === ValueTag.String ? { ...value, value: value.value.toUpperCase() } : value);
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
  }
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
    operand: parseCriteriaOperand(match[2] ?? "")
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
