import { ErrorCode, ValueTag, type CellValue } from "@bilig/protocol";
import type { ArrayValue } from "../runtime-values.js";
import type { LookupBuiltin, LookupBuiltinArgument, RangeBuiltinArgument } from "./lookup.js";

interface LookupSortFilterBuiltinDeps {
  errorValue: (code: ErrorCode) => CellValue;
  arrayResult: (values: CellValue[], rows: number, cols: number) => ArrayValue;
  isError: (
    value: LookupBuiltinArgument | undefined,
  ) => value is Extract<CellValue, { tag: ValueTag.Error }>;
  isRangeArg: (value: LookupBuiltinArgument | undefined) => value is RangeBuiltinArgument;
  toBoolean: (value: CellValue) => boolean | undefined;
  toInteger: (value: CellValue) => number | undefined;
  requireCellRange: (arg: LookupBuiltinArgument) => RangeBuiltinArgument | CellValue;
  toCellRange: (arg: LookupBuiltinArgument) => RangeBuiltinArgument | CellValue;
  compareScalars: (left: CellValue, right: CellValue) => number | undefined;
  getRangeValue: (range: RangeBuiltinArgument, row: number, col: number) => CellValue;
  pickRangeRow: (range: RangeBuiltinArgument, row: number) => CellValue[];
}

function pickRangeCol(
  range: RangeBuiltinArgument,
  col: number,
  deps: LookupSortFilterBuiltinDeps,
): CellValue[] {
  const values: CellValue[] = [];
  for (let row = 0; row < range.rows; row += 1) {
    values.push(deps.getRangeValue(range, row, col));
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

function rowKey(
  range: RangeBuiltinArgument,
  row: number,
  deps: LookupSortFilterBuiltinDeps,
): string | undefined {
  const values = deps.pickRangeRow(range, row);
  if (values.some(deps.isError)) {
    return undefined;
  }
  return JSON.stringify(values.map(normalizeKeyValue));
}

function colKey(
  range: RangeBuiltinArgument,
  col: number,
  deps: LookupSortFilterBuiltinDeps,
): string | undefined {
  const values = pickRangeCol(range, col, deps);
  if (values.some(deps.isError)) {
    return undefined;
  }
  return JSON.stringify(values.map(normalizeKeyValue));
}

export function createLookupSortFilterBuiltins(
  deps: LookupSortFilterBuiltinDeps,
): Record<string, LookupBuiltin> {
  return {
    SORT: (
      arrayArg,
      sortIndexArg,
      sortOrderArg = { tag: ValueTag.Number, value: 1 },
      byColArg,
    ) => {
      if (arrayArg === undefined) {
        return deps.errorValue(ErrorCode.Value);
      }
      const array = deps.toCellRange(arrayArg);
      if (!deps.isRangeArg(array)) {
        return array;
      }
      if (
        deps.isRangeArg(sortIndexArg) ||
        deps.isRangeArg(sortOrderArg) ||
        deps.isRangeArg(byColArg)
      ) {
        return deps.errorValue(ErrorCode.Value);
      }
      if (deps.isError(sortIndexArg)) {
        return sortIndexArg;
      }
      if (deps.isError(sortOrderArg)) {
        return sortOrderArg;
      }
      if (deps.isError(byColArg)) {
        return byColArg;
      }

      const sortByCol = byColArg === undefined ? false : deps.toBoolean(byColArg);
      const sortOrder = sortOrderArg ? deps.toInteger(sortOrderArg) : 1;
      const sortIndex = sortIndexArg === undefined ? 1 : deps.toInteger(sortIndexArg);
      if (sortOrder === undefined || ![1, -1].includes(sortOrder)) {
        return deps.errorValue(ErrorCode.Value);
      }
      if (sortIndex === undefined || sortIndex < 1) {
        return deps.errorValue(ErrorCode.Value);
      }

      let sortError: CellValue | undefined;
      if (array.rows === 1 || array.cols === 1) {
        const values = [...array.values];
        const order: number[] = Array.from({ length: values.length }, (_, index) => index);
        order.sort((left, right) => {
          const cmp = deps.compareScalars(values[left]!, values[right]!);
          if (cmp === undefined) {
            sortError = deps.errorValue(ErrorCode.Value);
            return 0;
          }
          return cmp * sortOrder || left - right;
        });
        if (sortError) {
          return sortError;
        }
        return deps.arrayResult(
          order.map((index) => values[index]!),
          array.rows,
          array.cols,
        );
      }
      if (sortByCol) {
        if (sortIndex > array.rows) {
          return deps.errorValue(ErrorCode.Value);
        }
        const rowIndex = sortIndex - 1;
        const colOrder = Array.from({ length: array.cols }, (_, col) => col);
        colOrder.sort((left, right) => {
          const cmp = deps.compareScalars(
            deps.getRangeValue(array, rowIndex, left),
            deps.getRangeValue(array, rowIndex, right),
          );
          if (cmp === undefined) {
            sortError = deps.errorValue(ErrorCode.Value);
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
            values.push(deps.getRangeValue(array, row, col));
          }
        }
        return deps.arrayResult(values, array.rows, array.cols);
      }
      if (sortIndex > array.cols) {
        return deps.errorValue(ErrorCode.Value);
      }
      const columnIndex = sortIndex - 1;
      const rowOrder = Array.from({ length: array.rows }, (_, row) => row);
      rowOrder.sort((left, right) => {
        const cmp = deps.compareScalars(
          deps.getRangeValue(array, left, columnIndex),
          deps.getRangeValue(array, right, columnIndex),
        );
        if (cmp === undefined) {
          sortError = deps.errorValue(ErrorCode.Value);
          return 0;
        }
        return cmp * sortOrder || left - right;
      });
      if (sortError) {
        return sortError;
      }
      const values: CellValue[] = [];
      for (const row of rowOrder) {
        values.push(...deps.pickRangeRow(array, row));
      }
      return deps.arrayResult(values, array.rows, array.cols);
    },
    SORTBY: (arrayArg, ...criteriaArgs) => {
      if (arrayArg === undefined) {
        return deps.errorValue(ErrorCode.Value);
      }
      const array = deps.toCellRange(arrayArg);
      if (!deps.isRangeArg(array)) {
        return array;
      }
      if (array.rows > 1 && array.cols > 1) {
        return deps.errorValue(ErrorCode.Value);
      }
      if (criteriaArgs.length === 0) {
        return deps.errorValue(ErrorCode.Value);
      }

      const source = array.values;
      const indexes = Array.from({ length: source.length }, (_, index) => index);
      const criteria: { values: CellValue[]; order: number }[] = [];
      for (let index = 0; index < criteriaArgs.length; index += 1) {
        const criteriaArg = criteriaArgs[index];
        if (criteriaArg === undefined) {
          return deps.errorValue(ErrorCode.Value);
        }
        if (deps.isError(criteriaArg)) {
          return criteriaArg;
        }
        const byRange = deps.toCellRange(criteriaArg);
        if (!deps.isRangeArg(byRange)) {
          return byRange;
        }
        const nextArg = criteriaArgs[index + 1];
        if (deps.isError(nextArg)) {
          return nextArg;
        }
        if (nextArg !== undefined && !deps.isRangeArg(nextArg) && !deps.isError(nextArg)) {
          const orderValue = deps.toInteger(nextArg);
          if (orderValue === undefined || ![1, -1].includes(orderValue)) {
            return deps.errorValue(ErrorCode.Value);
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
        return deps.errorValue(ErrorCode.Value);
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
          const cmp = deps.compareScalars(leftValue, rightValue);
          if (cmp === undefined) {
            sortError = deps.errorValue(ErrorCode.Value);
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
      return deps.arrayResult(
        indexes.map((index) => array.values[index] ?? { tag: ValueTag.Empty }),
        array.rows,
        array.cols,
      );
    },
    FILTER: (
      arrayArg,
      includeArg,
      ifEmptyArg = { tag: ValueTag.Error, code: ErrorCode.Value },
    ) => {
      if (arrayArg === undefined || includeArg === undefined) {
        return deps.errorValue(ErrorCode.Value);
      }
      const array = deps.requireCellRange(arrayArg);
      const include = deps.requireCellRange(includeArg);
      if (!deps.isRangeArg(array)) {
        return array;
      }
      if (!deps.isRangeArg(include)) {
        return include;
      }
      if (include.rows === array.rows && include.cols === 1) {
        const values: CellValue[] = [];
        let keptRows = 0;
        for (let row = 0; row < array.rows; row += 1) {
          const includeValue = deps.getRangeValue(include, row, 0);
          if (deps.isError(includeValue)) {
            return includeValue;
          }
          const keep = deps.toBoolean(includeValue);
          if (keep === undefined) {
            return deps.errorValue(ErrorCode.Value);
          }
          if (!keep) {
            continue;
          }
          values.push(...deps.pickRangeRow(array, row));
          keptRows += 1;
        }
        if (keptRows === 0) {
          return deps.isRangeArg(ifEmptyArg) ? deps.errorValue(ErrorCode.Value) : ifEmptyArg;
        }
        return deps.arrayResult(values, keptRows, array.cols);
      }

      if (include.cols === array.cols && include.rows === 1) {
        const keptCols: number[] = [];
        for (let col = 0; col < array.cols; col += 1) {
          const includeValue = deps.getRangeValue(include, 0, col);
          if (deps.isError(includeValue)) {
            return includeValue;
          }
          const keep = deps.toBoolean(includeValue);
          if (keep === undefined) {
            return deps.errorValue(ErrorCode.Value);
          }
          if (keep) {
            keptCols.push(col);
          }
        }
        if (keptCols.length === 0) {
          return deps.isRangeArg(ifEmptyArg) ? deps.errorValue(ErrorCode.Value) : ifEmptyArg;
        }
        const values: CellValue[] = [];
        for (let row = 0; row < array.rows; row += 1) {
          for (const col of keptCols) {
            values.push(deps.getRangeValue(array, row, col));
          }
        }
        return deps.arrayResult(values, array.rows, keptCols.length);
      }

      return deps.errorValue(ErrorCode.Value);
    },
    UNIQUE: (
      arrayArg,
      byColArg = { tag: ValueTag.Boolean, value: false },
      exactlyOnceArg = { tag: ValueTag.Boolean, value: false },
    ) => {
      if (arrayArg === undefined) {
        return deps.errorValue(ErrorCode.Value);
      }
      const array = deps.requireCellRange(arrayArg);
      if (!deps.isRangeArg(array)) {
        return array;
      }
      if (deps.isRangeArg(byColArg) || deps.isRangeArg(exactlyOnceArg)) {
        return deps.errorValue(ErrorCode.Value);
      }
      if (deps.isError(byColArg)) {
        return byColArg;
      }
      if (deps.isError(exactlyOnceArg)) {
        return exactlyOnceArg;
      }
      const byCol = deps.toBoolean(byColArg);
      const exactlyOnce = deps.toBoolean(exactlyOnceArg);
      if (byCol === undefined || exactlyOnce === undefined) {
        return deps.errorValue(ErrorCode.Value);
      }

      if (array.rows === 1 || array.cols === 1) {
        const counts = new Map<string, number>();
        const keys: string[] = [];
        for (const value of array.values) {
          if (deps.isError(value)) {
            return value;
          }
          const key = JSON.stringify(normalizeKeyValue(value));
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
          ? deps.arrayResult(values, 1, values.length)
          : deps.arrayResult(values, values.length, 1);
      }

      if (byCol) {
        const counts = new Map<string, number>();
        const keys: string[] = [];
        for (let col = 0; col < array.cols; col += 1) {
          const key = colKey(array, col, deps);
          if (key === undefined) {
            return deps.errorValue(ErrorCode.Value);
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
            values.push(deps.getRangeValue(array, row, col));
          }
        }
        return deps.arrayResult(values, array.rows, keptCols.length);
      }

      const counts = new Map<string, number>();
      const keys: string[] = [];
      for (let row = 0; row < array.rows; row += 1) {
        const key = rowKey(array, row, deps);
        if (key === undefined) {
          return deps.errorValue(ErrorCode.Value);
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
        values.push(...deps.pickRangeRow(array, row));
      }
      return deps.arrayResult(values, keptRows.length, array.cols);
    },
  };
}
