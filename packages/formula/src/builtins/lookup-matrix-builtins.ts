import { ErrorCode, ValueTag, type CellValue } from "@bilig/protocol";
import type { ArrayValue } from "../runtime-values.js";
import type { LookupBuiltin, LookupBuiltinArgument, RangeBuiltinArgument } from "./lookup.js";

interface LookupMatrixBuiltinDeps {
  errorValue: (code: ErrorCode) => CellValue;
  numberResult: (value: number) => CellValue;
  arrayResult: (values: CellValue[], rows: number, cols: number) => ArrayValue;
  isRangeArg: (value: LookupBuiltinArgument | undefined) => value is RangeBuiltinArgument;
  requireCellRange: (arg: LookupBuiltinArgument) => RangeBuiltinArgument | CellValue;
  findFirstNonRange: (
    values: readonly (RangeBuiltinArgument | CellValue)[],
  ) => CellValue | undefined;
  areRangeArgs: (
    values: readonly (RangeBuiltinArgument | CellValue)[],
  ) => values is RangeBuiltinArgument[];
  toNumber: (value: CellValue) => number | undefined;
  toNumericMatrix: (arg: LookupBuiltinArgument) => number[][] | CellValue;
  flattenNumbers: (arg: LookupBuiltinArgument) => number[] | CellValue;
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

function sumOfNumbers(
  arg: LookupBuiltinArgument | undefined,
  deps: LookupMatrixBuiltinDeps,
): number | CellValue {
  if (arg === undefined) {
    return deps.errorValue(ErrorCode.Value);
  }
  const values = deps.flattenNumbers(arg);
  return Array.isArray(values) ? values.reduce((sum, value) => sum + value, 0) : values;
}

function pairwiseMatrixSum(
  xArg: LookupBuiltinArgument | undefined,
  yArg: LookupBuiltinArgument | undefined,
  deps: LookupMatrixBuiltinDeps,
  combine: (x: number, y: number) => number,
): CellValue {
  if (xArg === undefined || yArg === undefined) {
    return deps.errorValue(ErrorCode.Value);
  }
  const xValues = deps.flattenNumbers(xArg);
  const yValues = deps.flattenNumbers(yArg);
  if (!Array.isArray(xValues)) {
    return xValues;
  }
  if (!Array.isArray(yValues)) {
    return yValues;
  }
  if (xValues.length !== yValues.length) {
    return deps.errorValue(ErrorCode.Value);
  }
  let sum = 0;
  for (let index = 0; index < xValues.length; index += 1) {
    sum += combine(xValues[index]!, yValues[index]!);
  }
  return deps.numberResult(sum);
}

export function createLookupMatrixBuiltins(
  deps: LookupMatrixBuiltinDeps,
): Record<string, LookupBuiltin> {
  return {
    SUMPRODUCT: (...args) => {
      if (args.length === 0) {
        return deps.errorValue(ErrorCode.Value);
      }
      const ranges = args.map((arg) => deps.requireCellRange(arg));
      const rangeError = deps.findFirstNonRange(ranges);
      if (rangeError) {
        return rangeError;
      }
      if (!deps.areRangeArgs(ranges)) {
        return deps.errorValue(ErrorCode.Value);
      }
      const typedRanges = ranges;
      const expectedLength = typedRanges[0]!.values.length;
      if (typedRanges.some((range) => range.values.length !== expectedLength)) {
        return deps.errorValue(ErrorCode.Value);
      }
      let sum = 0;
      for (let index = 0; index < expectedLength; index += 1) {
        let product = 1;
        for (const range of typedRanges) {
          product *= deps.toNumber(range.values[index]!) ?? 0;
        }
        sum += product;
      }
      return deps.numberResult(sum);
    },
    SUMX2MY2: (xArg, yArg) => {
      return pairwiseMatrixSum(xArg, yArg, deps, (x, y) => x ** 2 - y ** 2);
    },
    SUMX2PY2: (xArg, yArg) => {
      return pairwiseMatrixSum(xArg, yArg, deps, (x, y) => x ** 2 + y ** 2);
    },
    SUMXMY2: (xArg, yArg) => {
      return pairwiseMatrixSum(xArg, yArg, deps, (x, y) => (x - y) ** 2);
    },
    MDETERM: (matrixArg) => {
      if (matrixArg === undefined) {
        return deps.errorValue(ErrorCode.Value);
      }
      const matrix = deps.toNumericMatrix(matrixArg);
      if (!Array.isArray(matrix)) {
        return matrix;
      }
      if (matrix.length === 0 || matrix.some((row) => row.length !== matrix.length)) {
        return deps.errorValue(ErrorCode.Value);
      }
      return deps.numberResult(determinantOf(matrix));
    },
    MINVERSE: (matrixArg) => {
      if (matrixArg === undefined) {
        return deps.errorValue(ErrorCode.Value);
      }
      const matrix = deps.toNumericMatrix(matrixArg);
      if (!Array.isArray(matrix)) {
        return matrix;
      }
      if (matrix.length === 0 || matrix.some((row) => row.length !== matrix.length)) {
        return deps.errorValue(ErrorCode.Value);
      }
      const inverse = inverseOf(matrix);
      if (!inverse) {
        return deps.errorValue(ErrorCode.Value);
      }
      return deps.arrayResult(
        inverse.flat().map((value) => ({ tag: ValueTag.Number, value })),
        matrix.length,
        matrix.length,
      );
    },
    MMULT: (leftArg, rightArg) => {
      if (leftArg === undefined || rightArg === undefined) {
        return deps.errorValue(ErrorCode.Value);
      }
      const left = deps.toNumericMatrix(leftArg);
      const right = deps.toNumericMatrix(rightArg);
      if (!Array.isArray(left)) {
        return left;
      }
      if (!Array.isArray(right)) {
        return right;
      }
      if (left.length === 0 || right.length === 0 || left[0]!.length !== right.length) {
        return deps.errorValue(ErrorCode.Value);
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
      return deps.arrayResult(values, rows, cols);
    },
    PERCENTOF: (subsetArg, totalArg) => {
      const subset = sumOfNumbers(subsetArg, deps);
      const total = sumOfNumbers(totalArg, deps);
      if (typeof subset !== "number") {
        return subset;
      }
      if (typeof total !== "number") {
        return total;
      }
      if (total === 0) {
        return deps.errorValue(ErrorCode.Div0);
      }
      return deps.numberResult(subset / total);
    },
  };
}
