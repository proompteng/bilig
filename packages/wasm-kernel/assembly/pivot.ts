import { ErrorCode, ValueTag } from "./protocol";
import { tags, numbers, stringIds, errors, rangeOffsets, rangeLengths, rangeMembers } from "./vm";

export const PIVOT_AGG_SUM: u8 = 1;
export const PIVOT_AGG_COUNT: u8 = 2;

class PivotValueField {
  sourceColumnIndex: i32;
  summarizeBy: u8;
}

class PivotDefinition {
  sourceRangeIndex: i32;
  groupByColumnIndices: i32[];
  valueFields: PivotValueField[];
}

function getCellValue(cellIndex: i32): f64 {
  const tag = tags[cellIndex];
  if (tag == ValueTag.Number || tag == ValueTag.Boolean) return numbers[cellIndex];
  if (tag == ValueTag.Empty) return 0;
  return NaN;
}

// Result buffer management
export let pivotResultTags = new Uint8Array(64);
export let pivotResultNumbers = new Float64Array(64);
export let pivotResultStringIds = new Uint32Array(64);
export let pivotResultErrors = new Uint16Array(64);
export let pivotResultRows: i32 = 0;
export let pivotResultCols: i32 = 0;

function ensureResultCapacity(size: i32): void {
  if (pivotResultTags.length >= size) return;
  let nextLength = pivotResultTags.length;
  while (nextLength < size) nextLength *= 2;

  const nextTags = new Uint8Array(nextLength);
  nextTags.set(pivotResultTags);
  pivotResultTags = nextTags;

  const nextNumbers = new Float64Array(nextLength);
  nextNumbers.set(pivotResultNumbers);
  pivotResultNumbers = nextNumbers;

  const nextStringIds = new Uint32Array(nextLength);
  nextStringIds.set(pivotResultStringIds);
  pivotResultStringIds = nextStringIds;

  const nextErrors = new Uint16Array(nextLength);
  nextErrors.set(pivotResultErrors);
  pivotResultErrors = nextErrors;
}

export function materializePivotTable(
  sourceRangeIndex: i32,
  sourceWidth: i32,
  groupByCount: i32,
  groupByColumnIndices: Uint32Array,
  valueCount: i32,
  valueColumnIndices: Uint32Array,
  valueAggregations: Uint8Array,
): void {
  const rangeStart = rangeOffsets[sourceRangeIndex];
  const rangeLength = <i32>rangeLengths[sourceRangeIndex];

  if (
    sourceWidth <= 0 ||
    valueCount <= 0 ||
    rangeLength < sourceWidth ||
    rangeLength % sourceWidth != 0 ||
    groupByCount != groupByColumnIndices.length ||
    valueCount != valueColumnIndices.length ||
    valueCount != valueAggregations.length
  ) {
    writePivotError(<u16>ErrorCode.Value);
    return;
  }

  const rowCount = rangeLength / sourceWidth;
  if (rowCount <= 0) {
    writePivotError(<u16>ErrorCode.Value);
    return;
  }

  for (let i = 0; i < groupByCount; i++) {
    const columnIndex = groupByColumnIndices[i];
    if (columnIndex >= <u32>sourceWidth) {
      writePivotError(<u16>ErrorCode.Value);
      return;
    }
  }
  for (let i = 0; i < valueCount; i++) {
    const columnIndex = valueColumnIndices[i];
    if (columnIndex >= <u32>sourceWidth) {
      writePivotError(<u16>ErrorCode.Value);
      return;
    }
  }

  const buckets = new Map<string, Float64Array>();
  const bucketKeys = new Map<string, i32[]>();
  const keys: string[] = [];

  for (let rowIndex = 1; rowIndex < rowCount; rowIndex++) {
    const rowBase = rangeStart + rowIndex * sourceWidth;
    if (
      !hasMeaningfulRowValue(
        rowBase,
        groupByCount,
        groupByColumnIndices,
        valueCount,
        valueColumnIndices,
      )
    ) {
      continue;
    }

    let key = "";
    const rowKeyIndices: i32[] = [];
    for (let i = 0; i < groupByCount; i++) {
      const colIndex = groupByColumnIndices[i];
      const cellIndex = <i32>rangeMembers[rowBase + colIndex];
      key += cellKeyPart(cellIndex) + "|";
      rowKeyIndices.push(cellIndex);
    }

    if (!buckets.has(key)) {
      buckets.set(key, new Float64Array(valueCount));
      bucketKeys.set(key, rowKeyIndices);
      keys.push(key);
    }

    const aggregates = buckets.get(key);
    if (aggregates == null) {
      continue;
    }
    for (let i = 0; i < valueCount; i++) {
      const colIndex = valueColumnIndices[i];
      const aggType = valueAggregations[i];
      const cellIndex = <i32>rangeMembers[rowBase + colIndex];
      const tag = tags[cellIndex];
      if (aggType == PIVOT_AGG_SUM) {
        if (tag == ValueTag.Number) {
          aggregates[i] += numbers[cellIndex];
        }
      } else if (aggType == PIVOT_AGG_COUNT) {
        if (tag != ValueTag.Empty) {
          aggregates[i] += 1;
        }
      } else {
        writePivotError(<u16>ErrorCode.Value);
        return;
      }
    }
  }

  pivotResultCols = groupByCount + valueCount;
  pivotResultRows = keys.length + 1; // +1 for header
  ensureResultCapacity(pivotResultRows * pivotResultCols);

  for (let i = 0; i < groupByCount; i++) {
    const colIndex = groupByColumnIndices[i];
    copyCellToResult(i, <i32>rangeMembers[rangeStart + colIndex]);
  }
  for (let i = 0; i < valueCount; i++) {
    const colIndex = valueColumnIndices[i];
    const destIndex = groupByCount + i;
    copyCellToResult(destIndex, <i32>rangeMembers[rangeStart + colIndex]);
  }

  for (let rowIndex = 0; rowIndex < keys.length; rowIndex++) {
    const key = keys[rowIndex];
    const rowKeyIndices = bucketKeys.get(key);
    const aggregates = buckets.get(key);
    if (rowKeyIndices == null || aggregates == null) {
      continue;
    }
    const destRowBase = (rowIndex + 1) * pivotResultCols;

    for (let col = 0; col < groupByCount; col++) {
      copyCellToResult(destRowBase + col, rowKeyIndices[col]);
    }

    for (let col = 0; col < valueCount; col++) {
      const destIndex = destRowBase + groupByCount + col;
      pivotResultTags[destIndex] = ValueTag.Number;
      pivotResultNumbers[destIndex] = aggregates[col];
      pivotResultStringIds[destIndex] = 0;
      pivotResultErrors[destIndex] = ErrorCode.None;
    }
  }
}

function hasMeaningfulRowValue(
  rowBase: i32,
  groupByCount: i32,
  groupByColumnIndices: Uint32Array,
  valueCount: i32,
  valueColumnIndices: Uint32Array,
): bool {
  for (let i = 0; i < groupByCount; i++) {
    if (tags[rangeMembers[rowBase + groupByColumnIndices[i]]] != ValueTag.Empty) {
      return true;
    }
  }
  for (let i = 0; i < valueCount; i++) {
    if (tags[rangeMembers[rowBase + valueColumnIndices[i]]] != ValueTag.Empty) {
      return true;
    }
  }
  return false;
}

function cellKeyPart(cellIndex: i32): string {
  const tag = tags[cellIndex];
  if (tag == ValueTag.Empty) return "E";
  if (tag == ValueTag.Number) return "N:" + numbers[cellIndex].toString();
  if (tag == ValueTag.Boolean) return "B:" + (numbers[cellIndex] != 0 ? "1" : "0");
  if (tag == ValueTag.String) return "S:" + stringIds[cellIndex].toString();
  if (tag == ValueTag.Error) return "R:" + errors[cellIndex].toString();
  return "U";
}

function copyCellToResult(destIndex: i32, sourceCellIndex: i32): void {
  pivotResultTags[destIndex] = tags[sourceCellIndex];
  pivotResultNumbers[destIndex] = numbers[sourceCellIndex];
  pivotResultStringIds[destIndex] = stringIds[sourceCellIndex];
  pivotResultErrors[destIndex] = errors[sourceCellIndex];
}

function writePivotError(code: u16): void {
  ensureResultCapacity(1);
  pivotResultRows = 1;
  pivotResultCols = 1;
  pivotResultTags[0] = ValueTag.Error;
  pivotResultNumbers[0] = 0;
  pivotResultStringIds[0] = 0;
  pivotResultErrors[0] = code;
}

export function getPivotResultTagsPtr(): usize {
  return changetype<usize>(pivotResultTags.dataStart);
}

export function getPivotResultNumbersPtr(): usize {
  return changetype<usize>(pivotResultNumbers.dataStart);
}

export function getPivotResultStringIdsPtr(): usize {
  return changetype<usize>(pivotResultStringIds.dataStart);
}

export function getPivotResultErrorsPtr(): usize {
  return changetype<usize>(pivotResultErrors.dataStart);
}
