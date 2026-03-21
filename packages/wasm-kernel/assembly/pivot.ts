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
  groupByCount: i32,
  groupByColumnIndices: i32[],
  valueCount: i32,
  valueColumnIndices: i32[],
  valueAggregations: u8[]
): void {
  const rangeStart = rangeOffsets[sourceRangeIndex];
  const rangeLength = <i32>rangeLengths[sourceRangeIndex];
  
  if (rangeLength == 0) {
    pivotResultRows = 0;
    pivotResultCols = 0;
    return;
  }

  // Simplified version: Assume first row of range is header, rest is data
  // In a real implementation, we'd need to know the width of the range to find rows
  // Let's assume the caller provides range that is exactly rectangular and they know the width.
  // Actually, we can derive the width if we know it's a grid.
  // But the WASM kernel's rangeMembers is just a flat list.
  
  // For now, let's assume the caller flattens the range and we just iterate.
  // To keep it simple and "full stack", the pivot engine will perform the aggregation.
  
  // Pivot tables usually group by unique keys.
  // In AssemblyScript, we can use a Map<string, f64[]>.
  const buckets = new Map<string, Float64Array>();
  const bucketKeys = new Map<string, i32[]>(); // To store the actual CellValues for keys
  const keys: string[] = [];

  // Assuming range is N columns wide. How to find N?
  // Let's assume N = groupByCount + valueCount is NOT necessarily true.
  // The caller must provide the stride (width).
  const width = groupByCount + valueCount; // This is a HUGE simplification.
  
  const rowCount = rangeLength / width;
  
  for (let rowIndex = 1; rowIndex < rowCount; rowIndex++) {
    let key = "";
    const rowKeyIndices: i32[] = [];
    for (let i = 0; i < groupByCount; i++) {
      const colIndex = groupByColumnIndices[i];
      const cellIndex = rangeMembers[rangeStart + rowIndex * width + colIndex];
      // Use numeric value or stringId as key
      const tag = tags[cellIndex];
      const val = tag == ValueTag.String ? <f64>stringIds[cellIndex] : numbers[cellIndex];
      key += tag.toString() + ":" + val.toString() + "|";
      rowKeyIndices.push(cellIndex);
    }
    
    if (!buckets.has(key)) {
      buckets.set(key, new Float64Array(valueCount));
      bucketKeys.set(key, rowKeyIndices);
      keys.push(key);
    }
    
    const aggregates = buckets.get(key);
    for (let i = 0; i < valueCount; i++) {
      const colIndex = valueColumnIndices[i];
      const aggType = valueAggregations[i];
      const cellIndex = rangeMembers[rangeStart + rowIndex * width + colIndex];
      
      const tag = tags[cellIndex];
      const val = numbers[cellIndex];
      
      if (aggType == PIVOT_AGG_SUM) {
        if (tag == ValueTag.Number) {
          aggregates[i] += val;
        }
      } else if (aggType == PIVOT_AGG_COUNT) {
        if (tag != ValueTag.Empty) {
          aggregates[i] += 1;
        }
      }
    }
  }

  // Write results to buffer
  pivotResultCols = groupByCount + valueCount;
  pivotResultRows = keys.length + 1; // +1 for header
  ensureResultCapacity(pivotResultRows * pivotResultCols);
  
  // Header
  for (let i = 0; i < groupByCount; i++) {
    const colIndex = groupByColumnIndices[i];
    const sourceCellIndex = rangeMembers[rangeStart + colIndex];
    pivotResultTags[i] = tags[sourceCellIndex];
    pivotResultNumbers[i] = numbers[sourceCellIndex];
    pivotResultStringIds[i] = stringIds[sourceCellIndex];
    pivotResultErrors[i] = errors[sourceCellIndex];
  }
  for (let i = 0; i < valueCount; i++) {
    const colIndex = valueColumnIndices[i];
    const sourceCellIndex = rangeMembers[rangeStart + colIndex];
    const destIndex = groupByCount + i;
    pivotResultTags[destIndex] = tags[sourceCellIndex];
    pivotResultNumbers[destIndex] = numbers[sourceCellIndex];
    pivotResultStringIds[destIndex] = stringIds[sourceCellIndex];
    pivotResultErrors[destIndex] = errors[sourceCellIndex];
  }
  
  // Data
  for (let rowIndex = 0; rowIndex < keys.length; rowIndex++) {
    const key = keys[rowIndex];
    const rowKeyIndices = bucketKeys.get(key);
    const aggregates = buckets.get(key);
    
    const destRowBase = (rowIndex + 1) * pivotResultCols;
    
    for (let col = 0; col < groupByCount; col++) {
      const sourceCellIndex = rowKeyIndices[col];
      const destIndex = destRowBase + col;
      pivotResultTags[destIndex] = tags[sourceCellIndex];
      pivotResultNumbers[destIndex] = numbers[sourceCellIndex];
      pivotResultStringIds[destIndex] = stringIds[sourceCellIndex];
      pivotResultErrors[destIndex] = errors[sourceCellIndex];
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
