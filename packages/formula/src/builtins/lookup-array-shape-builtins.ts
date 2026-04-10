import { ErrorCode, ValueTag, type CellValue } from "@bilig/protocol";
import type { ArrayValue } from "../runtime-values.js";
import type { LookupBuiltin, LookupBuiltinArgument, RangeBuiltinArgument } from "./lookup.js";

interface LookupArrayShapeBuiltinDeps {
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
  getRangeValue: (range: RangeBuiltinArgument, row: number, col: number) => CellValue;
  findFirstNonRange: (
    values: readonly (RangeBuiltinArgument | CellValue)[],
  ) => CellValue | undefined;
  areRangeArgs: (
    values: readonly (RangeBuiltinArgument | CellValue)[],
  ) => values is RangeBuiltinArgument[];
  pickRangeRow: (range: RangeBuiltinArgument, row: number) => CellValue[];
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
  getRangeValue: LookupArrayShapeBuiltinDeps["getRangeValue"],
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
  getRangeValue: LookupArrayShapeBuiltinDeps["getRangeValue"],
): CellValue[] {
  const values: CellValue[] = [];
  for (let row = 0; row < rowCount; row += 1) {
    for (let col = 0; col < colCount; col += 1) {
      values.push(getRangeValue(range, rowStart + row, colStart + col));
    }
  }
  return values;
}

export function createLookupArrayShapeBuiltins(
  deps: LookupArrayShapeBuiltinDeps,
): Record<string, LookupBuiltin> {
  return {
    AREAS: (arrayArg) => {
      const range = deps.requireCellRange(arrayArg);
      if (!deps.isRangeArg(range)) {
        return range;
      }
      return { tag: ValueTag.Number, value: 1 };
    },
    ARRAYTOTEXT: (arrayArg, formatArg = { tag: ValueTag.Number, value: 0 }) => {
      const array = deps.toCellRange(arrayArg);
      if (!deps.isRangeArg(array)) {
        return array;
      }
      if (deps.isRangeArg(formatArg)) {
        return deps.errorValue(ErrorCode.Value);
      }
      const format = deps.toInteger(formatArg);
      if (format === undefined || (format !== 0 && format !== 1)) {
        return deps.errorValue(ErrorCode.Value);
      }
      const strict = format === 1;
      const lines: string[] = [];
      for (let row = 0; row < array.rows; row += 1) {
        const lineValues: string[] = [];
        for (let col = 0; col < array.cols; col += 1) {
          const value = arrayTextCell(deps.getRangeValue(array, row, col), strict);
          if (value === undefined) {
            return deps.errorValue(ErrorCode.Value);
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
      const range = deps.requireCellRange(arrayArg);
      if (!deps.isRangeArg(range)) {
        return range;
      }
      return { tag: ValueTag.Number, value: range.cols };
    },
    ROWS: (arrayArg) => {
      const range = deps.requireCellRange(arrayArg);
      if (!deps.isRangeArg(range)) {
        return range;
      }
      return { tag: ValueTag.Number, value: range.rows };
    },
    INDEX: (array, rowNumValue, colNumValue = { tag: ValueTag.Number, value: 1 }) => {
      if (!deps.isRangeArg(array) || array.refKind !== "cells") {
        return deps.errorValue(ErrorCode.Value);
      }
      if (deps.isRangeArg(rowNumValue) || deps.isRangeArg(colNumValue)) {
        return deps.errorValue(ErrorCode.Value);
      }
      if (deps.isError(rowNumValue)) {
        return rowNumValue;
      }
      if (deps.isError(colNumValue)) {
        return colNumValue;
      }

      const rawRowNum = deps.toInteger(rowNumValue);
      const rawColNum = deps.toInteger(colNumValue);
      if (rawRowNum === undefined || rawColNum === undefined) {
        return deps.errorValue(ErrorCode.Value);
      }

      let rowNum = rawRowNum;
      let colNum = rawColNum;
      if (array.rows === 1 && rawColNum === 1) {
        rowNum = 1;
        colNum = rawRowNum;
      }

      if (rowNum < 1 || colNum < 1 || rowNum > array.rows || colNum > array.cols) {
        return deps.errorValue(ErrorCode.Ref);
      }

      return deps.getRangeValue(array, rowNum - 1, colNum - 1);
    },
    OFFSET: (referenceArg, rowsArg, colsArg, heightArg, widthArg, areaNumberArg) => {
      if (
        deps.isRangeArg(rowsArg) ||
        deps.isRangeArg(colsArg) ||
        deps.isRangeArg(heightArg) ||
        deps.isRangeArg(widthArg) ||
        deps.isRangeArg(areaNumberArg)
      ) {
        return deps.errorValue(ErrorCode.Value);
      }
      if (
        deps.isError(rowsArg) ||
        deps.isError(colsArg) ||
        deps.isError(heightArg) ||
        deps.isError(widthArg) ||
        deps.isError(areaNumberArg)
      ) {
        return deps.isError(rowsArg)
          ? rowsArg
          : deps.isError(colsArg)
            ? colsArg
            : deps.isError(heightArg)
              ? heightArg
              : deps.isError(widthArg)
                ? widthArg
                : areaNumberArg;
      }
      const reference = deps.toCellRange(referenceArg);
      if (!deps.isRangeArg(reference)) {
        return reference;
      }
      const rows = deps.toInteger(rowsArg);
      const cols = deps.toInteger(colsArg);
      const height = heightArg === undefined ? reference.rows : deps.toInteger(heightArg);
      const width = widthArg === undefined ? reference.cols : deps.toInteger(widthArg);
      const areaNumber = areaNumberArg === undefined ? 1 : deps.toInteger(areaNumberArg);
      if (
        rows === undefined ||
        cols === undefined ||
        height === undefined ||
        width === undefined ||
        areaNumber === undefined
      ) {
        return deps.errorValue(ErrorCode.Value);
      }
      if (areaNumber !== 1 || height < 1 || width < 1) {
        return deps.errorValue(ErrorCode.Value);
      }

      const rowStart = rows < 0 ? reference.rows + rows : rows;
      const colStart = cols < 0 ? reference.cols + cols : cols;
      if (
        rowStart < 0 ||
        colStart < 0 ||
        rowStart + height > reference.rows ||
        colStart + width > reference.cols
      ) {
        return deps.errorValue(ErrorCode.Ref);
      }
      if (height === 1 && width === 1) {
        return deps.getRangeValue(reference, rowStart, colStart);
      }
      return deps.arrayResult(
        getRangeWindowValues(reference, rowStart, colStart, height, width, deps.getRangeValue),
        height,
        width,
      );
    },
    TAKE: (arrayArg, rowsArg, colsArg) => {
      const array = deps.toCellRange(arrayArg);
      if (!deps.isRangeArg(array)) {
        return array;
      }
      if (deps.isRangeArg(rowsArg) || deps.isRangeArg(colsArg)) {
        return deps.errorValue(ErrorCode.Value);
      }
      if (deps.isError(rowsArg) || deps.isError(colsArg)) {
        return deps.isError(rowsArg) ? rowsArg : colsArg;
      }

      const requestedRows = rowsArg === undefined ? array.rows : deps.toInteger(rowsArg);
      const requestedCols = colsArg === undefined ? array.cols : deps.toInteger(colsArg);
      if (requestedRows === undefined || requestedCols === undefined) {
        return deps.errorValue(ErrorCode.Value);
      }

      const clippedRows = clipIndex(requestedRows, array.rows);
      const clippedCols = clipIndex(requestedCols, array.cols);
      if (clippedRows === undefined || clippedCols === undefined) {
        return deps.errorValue(ErrorCode.Value);
      }

      const rowCount =
        clippedRows > 0 ? Math.min(clippedRows, array.rows) : Math.min(-clippedRows, array.rows);
      const colCount =
        clippedCols > 0 ? Math.min(clippedCols, array.cols) : Math.min(-clippedCols, array.cols);
      const rowOffset = clippedRows > 0 ? 0 : Math.max(array.rows - rowCount, 0);
      const colOffset = clippedCols > 0 ? 0 : Math.max(array.cols - colCount, 0);
      if (rowCount === 0 || colCount === 0) {
        return deps.errorValue(ErrorCode.Value);
      }

      const values: CellValue[] = [];
      for (let row = 0; row < rowCount; row += 1) {
        for (let col = 0; col < colCount; col += 1) {
          values.push(deps.getRangeValue(array, row + rowOffset, col + colOffset));
        }
      }
      return deps.arrayResult(values, rowCount, colCount);
    },
    DROP: (arrayArg, rowsArg, colsArg) => {
      const array = deps.toCellRange(arrayArg);
      if (!deps.isRangeArg(array)) {
        return array;
      }
      if (deps.isRangeArg(rowsArg) || deps.isRangeArg(colsArg)) {
        return deps.errorValue(ErrorCode.Value);
      }
      if (deps.isError(rowsArg) || deps.isError(colsArg)) {
        return deps.isError(rowsArg) ? rowsArg : colsArg;
      }

      const requestedRows = rowsArg === undefined ? 0 : deps.toInteger(rowsArg);
      const requestedCols = colsArg === undefined ? 0 : deps.toInteger(colsArg);
      if (requestedRows === undefined || requestedCols === undefined) {
        return deps.errorValue(ErrorCode.Value);
      }

      const clippedRows = requestedRows === 0 ? 0 : clipIndex(requestedRows, array.rows);
      const clippedCols = requestedCols === 0 ? 0 : clipIndex(requestedCols, array.cols);
      if (clippedRows === undefined || clippedCols === undefined) {
        return deps.errorValue(ErrorCode.Value);
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
        return deps.errorValue(ErrorCode.Value);
      }

      const values: CellValue[] = [];
      for (let row = 0; row < rowCount; row += 1) {
        for (let col = 0; col < colCount; col += 1) {
          values.push(deps.getRangeValue(array, row + rowOffset, col + colOffset));
        }
      }
      return deps.arrayResult(values, rowCount, colCount);
    },
    CHOOSECOLS: (arrayArg, ...columnArgs) => {
      const array = deps.toCellRange(arrayArg);
      if (!deps.isRangeArg(array)) {
        return array;
      }
      if (columnArgs.length === 0) {
        return deps.errorValue(ErrorCode.Value);
      }

      const selectedCols: number[] = [];
      for (const arg of columnArgs) {
        if (deps.isRangeArg(arg)) {
          return deps.errorValue(ErrorCode.Value);
        }
        if (deps.isError(arg)) {
          return arg;
        }
        const selected = deps.toInteger(arg);
        if (selected === undefined || selected < 1 || selected > array.cols) {
          return deps.errorValue(ErrorCode.Value);
        }
        selectedCols.push(selected - 1);
      }

      const values: CellValue[] = [];
      for (let row = 0; row < array.rows; row += 1) {
        for (const col of selectedCols) {
          values.push(deps.getRangeValue(array, row, col));
        }
      }
      return deps.arrayResult(values, array.rows, selectedCols.length);
    },
    CHOOSEROWS: (arrayArg, ...rowArgs) => {
      const array = deps.toCellRange(arrayArg);
      if (!deps.isRangeArg(array)) {
        return array;
      }
      if (rowArgs.length === 0) {
        return deps.errorValue(ErrorCode.Value);
      }

      const selectedRows: number[] = [];
      for (const arg of rowArgs) {
        if (deps.isRangeArg(arg)) {
          return deps.errorValue(ErrorCode.Value);
        }
        if (deps.isError(arg)) {
          return arg;
        }
        const selected = deps.toInteger(arg);
        if (selected === undefined || selected < 1 || selected > array.rows) {
          return deps.errorValue(ErrorCode.Value);
        }
        selectedRows.push(selected - 1);
      }

      const values: CellValue[] = [];
      for (const row of selectedRows) {
        values.push(...deps.pickRangeRow(array, row));
      }
      return deps.arrayResult(values, selectedRows.length, array.cols);
    },
    TRANSPOSE: (arrayArg) => {
      const array = deps.toCellRange(arrayArg);
      if (!deps.isRangeArg(array)) {
        return array;
      }
      if (array.rows === 1 && array.cols === 1) {
        return array.values[0] ?? { tag: ValueTag.Empty };
      }
      const values: CellValue[] = [];
      for (let col = 0; col < array.cols; col += 1) {
        for (let row = 0; row < array.rows; row += 1) {
          values.push(deps.getRangeValue(array, row, col));
        }
      }
      return deps.arrayResult(values, array.cols, array.rows);
    },
    HSTACK: (...arrayArgs) => {
      if (arrayArgs.length === 0) {
        return deps.errorValue(ErrorCode.Value);
      }
      const arrays = arrayArgs.map(deps.toCellRange);
      const rangeError = deps.findFirstNonRange(arrays);
      if (rangeError) {
        return rangeError;
      }
      if (!deps.areRangeArgs(arrays)) {
        return deps.errorValue(ErrorCode.Value);
      }

      const rowCount = Math.max(...arrays.map((array) => array.rows));
      for (const array of arrays) {
        if (array.rows !== 1 && array.rows !== rowCount) {
          return deps.errorValue(ErrorCode.Value);
        }
      }

      const values: CellValue[] = [];
      const totalCols = arrays.reduce((acc, array) => acc + array.cols, 0);
      for (let row = 0; row < rowCount; row += 1) {
        for (const array of arrays) {
          for (let col = 0; col < array.cols; col += 1) {
            const sourceRow = array.rows === 1 ? 0 : row;
            values.push(deps.getRangeValue(array, sourceRow, col));
          }
        }
      }
      return deps.arrayResult(values, rowCount, totalCols);
    },
    VSTACK: (...arrayArgs) => {
      if (arrayArgs.length === 0) {
        return deps.errorValue(ErrorCode.Value);
      }
      const arrays = arrayArgs.map(deps.toCellRange);
      const rangeError = deps.findFirstNonRange(arrays);
      if (rangeError) {
        return rangeError;
      }
      if (!deps.areRangeArgs(arrays)) {
        return deps.errorValue(ErrorCode.Value);
      }

      const colCount = Math.max(...arrays.map((array) => array.cols));
      for (const array of arrays) {
        if (array.cols !== 1 && array.cols !== colCount) {
          return deps.errorValue(ErrorCode.Value);
        }
      }

      const values: CellValue[] = [];
      const totalRows = arrays.reduce((acc, array) => acc + array.rows, 0);
      for (const array of arrays) {
        for (let row = 0; row < array.rows; row += 1) {
          for (let col = 0; col < colCount; col += 1) {
            const sourceCol = array.cols === 1 ? 0 : col;
            values.push(deps.getRangeValue(array, row, sourceCol));
          }
        }
      }
      return deps.arrayResult(values, totalRows, colCount);
    },
    TOCOL: (arrayArg, ignoreArg = { tag: ValueTag.Number, value: 0 }, scanByColArg) => {
      const array = deps.toCellRange(arrayArg);
      if (!deps.isRangeArg(array)) {
        return array;
      }
      if (deps.isRangeArg(ignoreArg) || deps.isRangeArg(scanByColArg)) {
        return deps.errorValue(ErrorCode.Value);
      }
      if (deps.isError(ignoreArg) || deps.isError(scanByColArg)) {
        return deps.isError(ignoreArg) ? ignoreArg : scanByColArg;
      }
      const ignoreValue = ignoreArg === undefined ? 0 : deps.toInteger(ignoreArg);
      if (ignoreValue === undefined || ![0, 1].includes(ignoreValue)) {
        return deps.errorValue(ErrorCode.Value);
      }
      const scanByCol = scanByColArg === undefined ? true : deps.toBoolean(scanByColArg);
      if (scanByCol === undefined) {
        return deps.errorValue(ErrorCode.Value);
      }
      const values = flattenValues(array, scanByCol, deps.getRangeValue, ignoreValue === 1);
      return deps.arrayResult(values, values.length, 1);
    },
    TOROW: (arrayArg, ignoreArg = { tag: ValueTag.Number, value: 0 }, scanByColArg) => {
      const array = deps.toCellRange(arrayArg);
      if (!deps.isRangeArg(array)) {
        return array;
      }
      if (deps.isRangeArg(ignoreArg) || deps.isRangeArg(scanByColArg)) {
        return deps.errorValue(ErrorCode.Value);
      }
      if (deps.isError(ignoreArg) || deps.isError(scanByColArg)) {
        return deps.isError(ignoreArg) ? ignoreArg : scanByColArg;
      }
      const ignoreValue = ignoreArg === undefined ? 0 : deps.toInteger(ignoreArg);
      if (ignoreValue === undefined || ![0, 1].includes(ignoreValue)) {
        return deps.errorValue(ErrorCode.Value);
      }
      const scanByCol = scanByColArg === undefined ? false : deps.toBoolean(scanByColArg);
      if (scanByCol === undefined) {
        return deps.errorValue(ErrorCode.Value);
      }
      const values = flattenValues(array, scanByCol, deps.getRangeValue, ignoreValue === 1);
      return deps.arrayResult(values, 1, values.length);
    },
    WRAPROWS: (arrayArg, wrapCountArg, padWithArg, padByColArg) => {
      const array = deps.toCellRange(arrayArg);
      if (!deps.isRangeArg(array)) {
        return array;
      }
      if (deps.isRangeArg(wrapCountArg) || deps.isRangeArg(padWithArg) || deps.isRangeArg(padByColArg)) {
        return deps.errorValue(ErrorCode.Value);
      }
      if (deps.isError(wrapCountArg) || deps.isError(padByColArg)) {
        return deps.isError(wrapCountArg) ? wrapCountArg : padByColArg;
      }
      if (padWithArg !== undefined && deps.isError(padWithArg)) {
        return padWithArg;
      }
      const wrapCount = deps.toInteger(wrapCountArg);
      if (wrapCount === undefined || wrapCount < 1) {
        return deps.errorValue(ErrorCode.Value);
      }
      if (padByColArg !== undefined && deps.toBoolean(padByColArg) === undefined) {
        return deps.errorValue(ErrorCode.Value);
      }

      const values = array.values.slice();
      const rows = Math.ceil(values.length / wrapCount);
      const cols = wrapCount;
      const padValue: CellValue =
        padWithArg === undefined ? deps.errorValue(ErrorCode.NA) : padWithArg;
      while (values.length < rows * cols) {
        values.push(padValue);
      }
      return deps.arrayResult(values, rows, cols);
    },
    WRAPCOLS: (arrayArg, wrapCountArg, padWithArg, padByColArg) => {
      const array = deps.toCellRange(arrayArg);
      if (!deps.isRangeArg(array)) {
        return array;
      }
      if (deps.isRangeArg(wrapCountArg) || deps.isRangeArg(padWithArg) || deps.isRangeArg(padByColArg)) {
        return deps.errorValue(ErrorCode.Value);
      }
      if (deps.isError(wrapCountArg) || deps.isError(padByColArg)) {
        return deps.isError(wrapCountArg) ? wrapCountArg : padByColArg;
      }
      if (padWithArg !== undefined && deps.isError(padWithArg)) {
        return padWithArg;
      }
      const wrapCount = deps.toInteger(wrapCountArg);
      if (wrapCount === undefined || wrapCount < 1) {
        return deps.errorValue(ErrorCode.Value);
      }
      if (padByColArg !== undefined && deps.toBoolean(padByColArg) === undefined) {
        return deps.errorValue(ErrorCode.Value);
      }

      const values = array.values.slice();
      const rows = wrapCount;
      const cols = Math.ceil(values.length / rows);
      const padValue: CellValue =
        padWithArg === undefined ? deps.errorValue(ErrorCode.NA) : padWithArg;
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
      return deps.arrayResult(wrappedValues, rows, cols);
    },
  };
}
