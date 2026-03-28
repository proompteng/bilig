import { ErrorCode, ValueTag, formatErrorCode, type CellValue } from "@bilig/protocol";
import type { ArrayValue } from "./runtime-values.js";

export interface MatrixValue {
  rows: number;
  cols: number;
  values: readonly CellValue[];
}

export type AggregateEvaluator = (
  subset: readonly CellValue[],
  totalSet?: readonly CellValue[],
) => CellValue;

export interface GroupByOptions {
  fieldHeadersMode?: number;
  totalDepth?: number;
  sortOrder?: readonly number[];
  filterArray?: MatrixValue;
  fieldRelationship?: number;
  aggregate: AggregateEvaluator;
}

export interface PivotByOptions {
  fieldHeadersMode?: number;
  rowTotalDepth?: number;
  rowSortOrder?: readonly number[];
  colTotalDepth?: number;
  colSortOrder?: readonly number[];
  filterArray?: MatrixValue;
  relativeTo?: number;
  aggregate: AggregateEvaluator;
}

interface HeaderOptions {
  consumeInputHeaders: boolean;
  showHeaderRow: boolean;
  showFieldLabels: boolean;
}

interface GroupRowBucket {
  keyValues: CellValue[];
  subsets: CellValue[][];
  aggregates: CellValue[];
}

interface PivotKeyBucket {
  keyValues: CellValue[];
  subsets: CellValue[][];
}

function emptyValue(): CellValue {
  return { tag: ValueTag.Empty };
}

function numberValue(value: number): CellValue {
  return { tag: ValueTag.Number, value };
}

function stringValue(value: string): CellValue {
  return { tag: ValueTag.String, value, stringId: 0 };
}

function errorValue(code: ErrorCode): CellValue {
  return { tag: ValueTag.Error, code };
}

function cloneCellValue(value: CellValue): CellValue {
  switch (value.tag) {
    case ValueTag.Empty:
      return emptyValue();
    case ValueTag.Number:
      return numberValue(value.value);
    case ValueTag.Boolean:
      return { tag: ValueTag.Boolean, value: value.value };
    case ValueTag.String:
      return stringValue(value.value);
    case ValueTag.Error:
      return errorValue(value.code);
  }
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
  }
}

function toText(value: CellValue): string {
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
      return formatErrorCode(value.code);
  }
}

function matrixCell(matrix: MatrixValue, row: number, col: number): CellValue {
  return matrix.values[row * matrix.cols + col] ?? emptyValue();
}

function autoDetectHeaders(values: MatrixValue): boolean {
  if (values.rows < 2 || values.cols < 1) {
    return false;
  }
  const first = matrixCell(values, 0, 0);
  const second = matrixCell(values, 1, 0);
  return first.tag === ValueTag.String && second.tag === ValueTag.Number;
}

function normalizeHeaderOptions(
  fieldHeadersMode: number | undefined,
  values: MatrixValue,
  alwaysShowHeaders: boolean,
): HeaderOptions | undefined {
  const mode = fieldHeadersMode;
  if (mode !== undefined && ![-1, 0, 1, 2, 3].includes(Math.trunc(mode))) {
    return undefined;
  }
  const consumeInputHeaders =
    mode === 1 || mode === 3 || (mode === undefined && autoDetectHeaders(values));
  return {
    consumeInputHeaders,
    showHeaderRow: alwaysShowHeaders || mode === 2 || mode === 3 || mode === undefined,
    showFieldLabels: mode !== 1,
  };
}

function headerLabels(matrix: MatrixValue, consumeInputHeaders: boolean, prefix: string): string[] {
  return Array.from({ length: matrix.cols }, (_, index) => {
    if (consumeInputHeaders) {
      const text = toText(matrixCell(matrix, 0, index)).trim();
      if (text.length > 0) {
        return text;
      }
    }
    return `${prefix} ${index + 1}`;
  });
}

function normalizeFilterMask(
  filterArray: MatrixValue | undefined,
  totalRows: number,
  dataStartRow: number,
): boolean[] | undefined {
  const dataRows = Math.max(0, totalRows - dataStartRow);
  if (!filterArray) {
    return Array.from({ length: dataRows }, () => true);
  }
  if (!(filterArray.cols === 1 || filterArray.rows === 1)) {
    return undefined;
  }
  const source =
    filterArray.rows === totalRows
      ? Array.from({ length: dataRows }, (_, index) =>
          truthy(matrixCell(filterArray, dataStartRow + index, 0)),
        )
      : filterArray.rows === dataRows
        ? Array.from({ length: dataRows }, (_, index) => truthy(matrixCell(filterArray, index, 0)))
        : filterArray.cols === totalRows
          ? Array.from({ length: dataRows }, (_, index) =>
              truthy(matrixCell(filterArray, 0, dataStartRow + index)),
            )
          : filterArray.cols === dataRows
            ? Array.from({ length: dataRows }, (_, index) =>
                truthy(matrixCell(filterArray, 0, index)),
              )
            : undefined;
  return source;
}

function truthy(value: CellValue): boolean {
  return (toNumber(value) ?? 0) !== 0;
}

function cellKey(value: CellValue): string {
  switch (value.tag) {
    case ValueTag.Empty:
      return "E";
    case ValueTag.Number:
      return `N:${Object.is(value.value, -0) ? "-0" : String(value.value)}`;
    case ValueTag.Boolean:
      return value.value ? "B:1" : "B:0";
    case ValueTag.String:
      return `S:${value.value}`;
    case ValueTag.Error:
      return `R:${value.code}`;
  }
}

function compareCellValues(left: CellValue, right: CellValue): number {
  if (left.tag === ValueTag.String || right.tag === ValueTag.String) {
    const leftText = toText(left).toUpperCase();
    const rightText = toText(right).toUpperCase();
    if (leftText === rightText) {
      return 0;
    }
    return leftText < rightText ? -1 : 1;
  }
  const leftNumber = toNumber(left);
  const rightNumber = toNumber(right);
  if (leftNumber !== undefined && rightNumber !== undefined) {
    if (leftNumber === rightNumber) {
      return 0;
    }
    return leftNumber < rightNumber ? -1 : 1;
  }
  const leftText = toText(left).toUpperCase();
  const rightText = toText(right).toUpperCase();
  if (leftText === rightText) {
    return 0;
  }
  return leftText < rightText ? -1 : 1;
}

function normalizeSortOrders(
  sortOrder: readonly number[] | undefined,
): Array<{ index: number; descending: boolean }> {
  if (!sortOrder || sortOrder.length === 0) {
    return [];
  }
  return sortOrder
    .map((value) => Math.trunc(value))
    .filter((value) => Number.isFinite(value) && value !== 0)
    .map((value) => ({ index: Math.abs(value) - 1, descending: value < 0 }));
}

function orderBuckets<T extends { keyValues: CellValue[]; aggregates: CellValue[] }>(
  buckets: T[],
  sortOrder: readonly number[] | undefined,
): T[] {
  const orders = normalizeSortOrders(sortOrder);
  if (orders.length === 0) {
    return buckets;
  }
  return buckets.toSorted((left, right) => {
    for (const order of orders) {
      const leftValue =
        order.index < left.keyValues.length
          ? left.keyValues[order.index]!
          : (left.aggregates[order.index - left.keyValues.length] ?? emptyValue());
      const rightValue =
        order.index < right.keyValues.length
          ? right.keyValues[order.index]!
          : (right.aggregates[order.index - right.keyValues.length] ?? emptyValue());
      const comparison = compareCellValues(leftValue, rightValue);
      if (comparison !== 0) {
        return order.descending ? -comparison : comparison;
      }
    }
    return 0;
  });
}

function totalRowLabel(width: number): CellValue[] {
  return Array.from({ length: width }, (_, index) =>
    index === 0 ? stringValue("Total") : emptyValue(),
  );
}

export function evaluateGroupBy(
  rowFields: MatrixValue,
  values: MatrixValue,
  options: GroupByOptions,
): ArrayValue | CellValue {
  if (
    rowFields.rows !== values.rows ||
    rowFields.rows === 0 ||
    rowFields.cols === 0 ||
    values.cols === 0
  ) {
    return errorValue(ErrorCode.Value);
  }
  if (options.fieldRelationship !== undefined) {
    const relationship = Math.trunc(options.fieldRelationship);
    if (relationship !== 0 && relationship !== 1) {
      return errorValue(ErrorCode.Value);
    }
  }

  const headers = normalizeHeaderOptions(options.fieldHeadersMode, values, false);
  if (!headers) {
    return errorValue(ErrorCode.Value);
  }
  const dataStartRow = headers.consumeInputHeaders ? 1 : 0;
  const filterMask = normalizeFilterMask(options.filterArray, rowFields.rows, dataStartRow);
  if (!filterMask) {
    return errorValue(ErrorCode.Value);
  }

  const rowLabels = headerLabels(rowFields, headers.consumeInputHeaders, "Row Field");
  const valueLabels = headerLabels(values, headers.consumeInputHeaders, "Value");

  const buckets = new Map<string, GroupRowBucket>();
  const order: string[] = [];
  const totalSubsets = Array.from({ length: values.cols }, () => new Array<CellValue>());

  for (let row = dataStartRow; row < rowFields.rows; row += 1) {
    if (!filterMask[row - dataStartRow]) {
      continue;
    }
    const keyValues = Array.from({ length: rowFields.cols }, (_, col) =>
      cloneCellValue(matrixCell(rowFields, row, col)),
    );
    const key = keyValues.map(cellKey).join("\u001f");
    let bucket = buckets.get(key);
    if (!bucket) {
      bucket = {
        keyValues,
        subsets: Array.from({ length: values.cols }, () => new Array<CellValue>()),
        aggregates: [],
      };
      buckets.set(key, bucket);
      order.push(key);
    }
    for (let col = 0; col < values.cols; col += 1) {
      const value = cloneCellValue(matrixCell(values, row, col));
      bucket.subsets[col]!.push(value);
      totalSubsets[col]!.push(cloneCellValue(value));
    }
  }

  const rows = order.map((key) => {
    const bucket = buckets.get(key)!;
    const aggregates = bucket.subsets.map((subset, valueIndex) =>
      cloneCellValue(options.aggregate(subset, totalSubsets[valueIndex])),
    );
    return {
      keyValues: bucket.keyValues,
      subsets: bucket.subsets,
      aggregates,
    };
  });
  const orderedRows = orderBuckets(rows, options.sortOrder);

  const totalDepth = Math.trunc(options.totalDepth ?? 1);
  const includeTotals = totalDepth !== 0;
  const totalsAtTop = totalDepth < 0;
  const totalAggregates = totalSubsets.map((subset) => cloneCellValue(options.aggregate(subset)));

  const outputRows: CellValue[][] = [];
  if (headers.showHeaderRow) {
    outputRows.push([
      ...(headers.showFieldLabels ? rowLabels.map(stringValue) : totalRowLabel(rowFields.cols)),
      ...valueLabels.map(stringValue),
    ]);
  }
  if (includeTotals && totalsAtTop) {
    outputRows.push([...totalRowLabel(rowFields.cols), ...totalAggregates]);
  }
  orderedRows.forEach((row) => {
    outputRows.push([...row.keyValues.map(cloneCellValue), ...row.aggregates.map(cloneCellValue)]);
  });
  if (includeTotals && !totalsAtTop) {
    outputRows.push([...totalRowLabel(rowFields.cols), ...totalAggregates]);
  }

  if (outputRows.length === 0) {
    return { kind: "array", rows: 1, cols: 1, values: [emptyValue()] };
  }

  return {
    kind: "array",
    rows: outputRows.length,
    cols: rowFields.cols + values.cols,
    values: outputRows.flat(),
  };
}

export function evaluatePivotBy(
  rowFields: MatrixValue,
  colFields: MatrixValue,
  values: MatrixValue,
  options: PivotByOptions,
): ArrayValue | CellValue {
  if (
    rowFields.rows !== colFields.rows ||
    rowFields.rows !== values.rows ||
    rowFields.rows === 0 ||
    rowFields.cols === 0 ||
    colFields.cols === 0 ||
    values.cols === 0
  ) {
    return errorValue(ErrorCode.Value);
  }

  const headers = normalizeHeaderOptions(options.fieldHeadersMode, values, true);
  if (!headers) {
    return errorValue(ErrorCode.Value);
  }
  const dataStartRow = headers.consumeInputHeaders ? 1 : 0;
  const filterMask = normalizeFilterMask(options.filterArray, rowFields.rows, dataStartRow);
  if (!filterMask) {
    return errorValue(ErrorCode.Value);
  }

  const rowLabels = headerLabels(rowFields, headers.consumeInputHeaders, "Row Field");
  const valueLabels = headerLabels(values, headers.consumeInputHeaders, "Value");

  const rowBuckets = new Map<string, PivotKeyBucket>();
  const colBuckets = new Map<string, PivotKeyBucket>();
  const cellBuckets = new Map<string, CellValue[][]>();
  const rowOrder: string[] = [];
  const colOrder: string[] = [];
  const grandSubsets = Array.from({ length: values.cols }, () => new Array<CellValue>());

  for (let row = dataStartRow; row < rowFields.rows; row += 1) {
    if (!filterMask[row - dataStartRow]) {
      continue;
    }
    const rowKeyValues = Array.from({ length: rowFields.cols }, (_, col) =>
      cloneCellValue(matrixCell(rowFields, row, col)),
    );
    const colKeyValues = Array.from({ length: colFields.cols }, (_, col) =>
      cloneCellValue(matrixCell(colFields, row, col)),
    );
    const rowKey = rowKeyValues.map(cellKey).join("\u001f");
    const colKey = colKeyValues.map(cellKey).join("\u001f");
    if (!rowBuckets.has(rowKey)) {
      rowBuckets.set(rowKey, {
        keyValues: rowKeyValues,
        subsets: Array.from({ length: values.cols }, () => new Array<CellValue>()),
      });
      rowOrder.push(rowKey);
    }
    if (!colBuckets.has(colKey)) {
      colBuckets.set(colKey, {
        keyValues: colKeyValues,
        subsets: Array.from({ length: values.cols }, () => new Array<CellValue>()),
      });
      colOrder.push(colKey);
    }
    const pairKey = `${rowKey}\u0000${colKey}`;
    let pairBucket = cellBuckets.get(pairKey);
    if (!pairBucket) {
      pairBucket = Array.from({ length: values.cols }, () => new Array<CellValue>());
      cellBuckets.set(pairKey, pairBucket);
    }
    const rowBucket = rowBuckets.get(rowKey)!;
    const colBucket = colBuckets.get(colKey)!;
    for (let valueIndex = 0; valueIndex < values.cols; valueIndex += 1) {
      const value = cloneCellValue(matrixCell(values, row, valueIndex));
      pairBucket[valueIndex]!.push(cloneCellValue(value));
      rowBucket.subsets[valueIndex]!.push(cloneCellValue(value));
      colBucket.subsets[valueIndex]!.push(cloneCellValue(value));
      grandSubsets[valueIndex]!.push(cloneCellValue(value));
    }
  }

  const rowResults = rowOrder.map((key) => {
    const bucket = rowBuckets.get(key)!;
    return {
      keyValues: bucket.keyValues,
      subsets: bucket.subsets,
      aggregates: bucket.subsets.map((subset, index) =>
        cloneCellValue(options.aggregate(subset, grandSubsets[index])),
      ),
    };
  });
  const colResults = colOrder.map((key) => {
    const bucket = colBuckets.get(key)!;
    return {
      keyValues: bucket.keyValues,
      subsets: bucket.subsets,
      aggregates: bucket.subsets.map((subset, index) =>
        cloneCellValue(options.aggregate(subset, grandSubsets[index])),
      ),
    };
  });
  const orderedRows = orderBuckets(rowResults, options.rowSortOrder);
  const orderedCols = orderBuckets(colResults, options.colSortOrder);

  const colTotalDepth = Math.trunc(options.colTotalDepth ?? 1);
  const rowTotalDepth = Math.trunc(options.rowTotalDepth ?? 1);
  const includeColTotals = colTotalDepth !== 0;
  const includeRowTotals = rowTotalDepth !== 0;
  const colTotalsFirst = colTotalDepth < 0;
  const rowTotalsFirst = rowTotalDepth < 0;

  const pivotColumnLabels: string[] = [];
  const materializedColumns: Array<{ colKey: string | null; valueIndex: number; label: string }> =
    [];
  if (includeColTotals && colTotalsFirst) {
    for (let valueIndex = 0; valueIndex < values.cols; valueIndex += 1) {
      const label = values.cols === 1 ? "Total" : `Total | ${valueLabels[valueIndex]}`;
      pivotColumnLabels.push(label);
      materializedColumns.push({ colKey: null, valueIndex, label });
    }
  }
  orderedCols.forEach((column) => {
    const keyLabel = column.keyValues.map(toText).join(" | ");
    for (let valueIndex = 0; valueIndex < values.cols; valueIndex += 1) {
      const label = values.cols === 1 ? keyLabel : `${keyLabel} | ${valueLabels[valueIndex]}`;
      pivotColumnLabels.push(label);
      materializedColumns.push({
        colKey: column.keyValues.map(cellKey).join("\u001f"),
        valueIndex,
        label,
      });
    }
  });
  if (includeColTotals && !colTotalsFirst) {
    for (let valueIndex = 0; valueIndex < values.cols; valueIndex += 1) {
      const label = values.cols === 1 ? "Total" : `Total | ${valueLabels[valueIndex]}`;
      pivotColumnLabels.push(label);
      materializedColumns.push({ colKey: null, valueIndex, label });
    }
  }

  const headerRow: CellValue[] = [
    ...(headers.showFieldLabels
      ? rowLabels.map(stringValue)
      : Array.from({ length: rowFields.cols }, () => emptyValue())),
    ...pivotColumnLabels.map(stringValue),
  ];

  const rows: CellValue[][] = [headerRow];
  const totalRowPrefix = totalRowLabel(rowFields.cols);

  const buildDataRow = (
    rowKey: string,
    rowValues: CellValue[],
    rowSubsets: CellValue[][],
  ): CellValue[] => {
    const output: CellValue[] = rowValues.map(cloneCellValue);
    materializedColumns.forEach((column) => {
      if (column.colKey === null) {
        output.push(
          cloneCellValue(
            options.aggregate(
              rowSubsets[column.valueIndex] ?? [],
              grandSubsets[column.valueIndex] ?? [],
            ),
          ),
        );
        return;
      }
      if (rowKey === "__TOTAL__") {
        const columnSubset = colBuckets.get(column.colKey)?.subsets[column.valueIndex] ?? [];
        output.push(
          cloneCellValue(options.aggregate(columnSubset, grandSubsets[column.valueIndex] ?? [])),
        );
        return;
      }
      const pairSubsets = cellBuckets.get(`${rowKey}\u0000${column.colKey}`);
      const subset = pairSubsets?.[column.valueIndex] ?? [];
      const totalSet =
        options.relativeTo === 1 || options.relativeTo === 4
          ? (rowSubsets[column.valueIndex] ?? [])
          : options.relativeTo === 2
            ? (grandSubsets[column.valueIndex] ?? [])
            : (colBuckets.get(column.colKey)?.subsets[column.valueIndex] ?? []);
      output.push(cloneCellValue(options.aggregate(subset, totalSet)));
    });
    return output;
  };

  if (includeRowTotals && rowTotalsFirst) {
    rows.push(
      buildDataRow(
        "__TOTAL__",
        totalRowPrefix,
        grandSubsets.map((subset) => subset.slice()),
      ),
    );
  }
  orderedRows.forEach((row) => {
    const rowKey = row.keyValues.map(cellKey).join("\u001f");
    rows.push(buildDataRow(rowKey, row.keyValues, row.subsets));
  });
  if (includeRowTotals && !rowTotalsFirst) {
    rows.push(
      buildDataRow(
        "__TOTAL__",
        totalRowPrefix,
        grandSubsets.map((subset) => subset.slice()),
      ),
    );
  }

  return {
    kind: "array",
    rows: rows.length,
    cols: rowFields.cols + materializedColumns.length,
    values: rows.flat(),
  };
}
