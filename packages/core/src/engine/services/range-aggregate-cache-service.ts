import { ErrorCode, ValueTag } from "@bilig/protocol";
import type { EngineRuntimeState } from "../runtime-state.js";
import type { EngineRuntimeColumnStoreService } from "./runtime-column-store-service.js";

interface AggregatePrefixEntry {
  sheetName: string;
  rowStart: number;
  rowEnd: number;
  col: number;
  columnVersion: number;
  structureVersion: number;
  prefixSums: Float64Array;
  prefixCount: Uint32Array;
  prefixAverageCount: Uint32Array;
  prefixErrorCodes: Uint16Array;
  prefixErrorCounts: Uint32Array;
  prefixMinimums: Float64Array;
  prefixMaximums: Float64Array;
}

export interface RangeAggregateCacheService {
  readonly getOrBuildPrefix: (request: {
    sheetName: string;
    rowStart: number;
    rowEnd: number;
    col: number;
  }) => AggregatePrefixEntry;
}

function cacheKey(sheetName: string, col: number, rowStart: number): string {
  return `${sheetName}\t${col}\t${rowStart}`;
}

function decodeValueTag(rawTag: number | undefined): ValueTag {
  if (rawTag === undefined) {
    return ValueTag.Empty;
  }
  switch (rawTag) {
    case 1:
      return ValueTag.Number;
    case 2:
      return ValueTag.Boolean;
    case 3:
      return ValueTag.String;
    case 4:
      return ValueTag.Error;
    case 0:
    default:
      return ValueTag.Empty;
  }
}

function ensurePrefixCapacity(existing: AggregatePrefixEntry, totalLength: number): void {
  if (existing.prefixSums.length >= totalLength) {
    return;
  }
  const nextCapacity = Math.max(existing.prefixSums.length * 2, totalLength);
  const nextPrefixSums = new Float64Array(nextCapacity);
  const nextPrefixCount = new Uint32Array(nextCapacity);
  const nextPrefixAverageCount = new Uint32Array(nextCapacity);
  const nextPrefixErrorCodes = new Uint16Array(nextCapacity);
  const nextPrefixErrorCounts = new Uint32Array(nextCapacity);
  const nextPrefixMinimums = new Float64Array(nextCapacity);
  const nextPrefixMaximums = new Float64Array(nextCapacity);
  const currentLength = existing.rowEnd - existing.rowStart + 1;
  nextPrefixSums.set(existing.prefixSums.subarray(0, currentLength), 0);
  nextPrefixCount.set(existing.prefixCount.subarray(0, currentLength), 0);
  nextPrefixAverageCount.set(existing.prefixAverageCount.subarray(0, currentLength), 0);
  nextPrefixErrorCodes.set(existing.prefixErrorCodes.subarray(0, currentLength), 0);
  nextPrefixErrorCounts.set(existing.prefixErrorCounts.subarray(0, currentLength), 0);
  nextPrefixMinimums.set(existing.prefixMinimums.subarray(0, currentLength), 0);
  nextPrefixMaximums.set(existing.prefixMaximums.subarray(0, currentLength), 0);
  existing.prefixSums = nextPrefixSums;
  existing.prefixCount = nextPrefixCount;
  existing.prefixAverageCount = nextPrefixAverageCount;
  existing.prefixErrorCodes = nextPrefixErrorCodes;
  existing.prefixErrorCounts = nextPrefixErrorCounts;
  existing.prefixMinimums = nextPrefixMinimums;
  existing.prefixMaximums = nextPrefixMaximums;
}

export function createRangeAggregateCacheService(args: {
  readonly state: Pick<EngineRuntimeState, "workbook">;
  readonly runtimeColumnStore: EngineRuntimeColumnStoreService;
}): RangeAggregateCacheService {
  const emptyColumnVersions = new Uint32Array(0);
  const cache = new Map<string, AggregatePrefixEntry>();

  const getCurrentVersions = (sheetName: string, col: number) => {
    const sheet = args.state.workbook.getSheet(sheetName);
    const columnVersions = sheet?.columnVersions ?? emptyColumnVersions;
    return {
      columnVersion: columnVersions[col] ?? 0,
      structureVersion: sheet?.structureVersion ?? 0,
    };
  };

  const buildPrefix = (request: {
    sheetName: string;
    rowStart: number;
    rowEnd: number;
    col: number;
  }): AggregatePrefixEntry => {
    const slice = args.runtimeColumnStore.getColumnSlice(request);
    const prefixSums = new Float64Array(slice.length);
    const prefixCount = new Uint32Array(slice.length);
    const prefixAverageCount = new Uint32Array(slice.length);
    const prefixErrorCodes = new Uint16Array(slice.length);
    const prefixErrorCounts = new Uint32Array(slice.length);
    const prefixMinimums = new Float64Array(slice.length);
    const prefixMaximums = new Float64Array(slice.length);
    let runningSum = 0;
    let runningCount = 0;
    let runningAverageCount = 0;
    let runningErrorCode = ErrorCode.None;
    let runningErrorCount = 0;
    let runningMinimum = Number.POSITIVE_INFINITY;
    let runningMaximum = Number.NEGATIVE_INFINITY;
    for (let offset = 0; offset < slice.length; offset += 1) {
      const tag = decodeValueTag(slice.tags[offset]);
      switch (tag) {
        case ValueTag.Number:
          runningSum += slice.numbers[offset] ?? 0;
          runningCount += 1;
          runningAverageCount += 1;
          runningMinimum = Math.min(runningMinimum, slice.numbers[offset] ?? 0);
          runningMaximum = Math.max(runningMaximum, slice.numbers[offset] ?? 0);
          break;
        case ValueTag.Boolean:
          const booleanNumber = (slice.numbers[offset] ?? 0) !== 0 ? 1 : 0;
          runningSum += booleanNumber;
          runningCount += 1;
          runningAverageCount += 1;
          runningMinimum = Math.min(runningMinimum, booleanNumber);
          runningMaximum = Math.max(runningMaximum, booleanNumber);
          break;
        case ValueTag.Empty:
          runningAverageCount += 1;
          runningMinimum = Math.min(runningMinimum, 0);
          runningMaximum = Math.max(runningMaximum, 0);
          break;
        case ValueTag.Error:
          runningErrorCode ||= slice.errors[offset] ?? ErrorCode.None;
          runningErrorCount += 1;
          break;
        case ValueTag.String:
        default:
          break;
      }
      prefixSums[offset] = runningSum;
      prefixCount[offset] = runningCount;
      prefixAverageCount[offset] = runningAverageCount;
      prefixErrorCodes[offset] = runningErrorCode;
      prefixErrorCounts[offset] = runningErrorCount;
      prefixMinimums[offset] = runningMinimum;
      prefixMaximums[offset] = runningMaximum;
    }
    return {
      sheetName: request.sheetName,
      rowStart: request.rowStart,
      rowEnd: request.rowEnd,
      col: request.col,
      columnVersion: slice.columnVersion,
      structureVersion: slice.structureVersion,
      prefixSums,
      prefixCount,
      prefixAverageCount,
      prefixErrorCodes,
      prefixErrorCounts,
      prefixMinimums,
      prefixMaximums,
    };
  };

  const extendPrefix = (
    existing: AggregatePrefixEntry,
    request: {
      sheetName: string;
      rowStart: number;
      rowEnd: number;
      col: number;
    },
  ): AggregatePrefixEntry => {
    const deltaSlice = args.runtimeColumnStore.getColumnSlice({
      sheetName: request.sheetName,
      rowStart: existing.rowEnd + 1,
      rowEnd: request.rowEnd,
      col: request.col,
    });
    const totalLength = request.rowEnd - request.rowStart + 1;
    ensurePrefixCapacity(existing, totalLength);
    const currentLength = existing.rowEnd - existing.rowStart + 1;
    let runningSum = existing.prefixSums[currentLength - 1] ?? 0;
    let runningCount = existing.prefixCount[currentLength - 1] ?? 0;
    let runningAverageCount = existing.prefixAverageCount[currentLength - 1] ?? 0;
    let runningErrorCode =
      (existing.prefixErrorCodes[currentLength - 1] as ErrorCode | undefined) ?? ErrorCode.None;
    let runningErrorCount = existing.prefixErrorCounts[currentLength - 1] ?? 0;
    let runningMinimum = existing.prefixMinimums[currentLength - 1] ?? Number.POSITIVE_INFINITY;
    let runningMaximum = existing.prefixMaximums[currentLength - 1] ?? Number.NEGATIVE_INFINITY;
    for (let offset = 0; offset < deltaSlice.length; offset += 1) {
      const tag = decodeValueTag(deltaSlice.tags[offset]);
      switch (tag) {
        case ValueTag.Number:
          runningSum += deltaSlice.numbers[offset] ?? 0;
          runningCount += 1;
          runningAverageCount += 1;
          runningMinimum = Math.min(runningMinimum, deltaSlice.numbers[offset] ?? 0);
          runningMaximum = Math.max(runningMaximum, deltaSlice.numbers[offset] ?? 0);
          break;
        case ValueTag.Boolean:
          const booleanNumber = (deltaSlice.numbers[offset] ?? 0) !== 0 ? 1 : 0;
          runningSum += booleanNumber;
          runningCount += 1;
          runningAverageCount += 1;
          runningMinimum = Math.min(runningMinimum, booleanNumber);
          runningMaximum = Math.max(runningMaximum, booleanNumber);
          break;
        case ValueTag.Empty:
          runningAverageCount += 1;
          runningMinimum = Math.min(runningMinimum, 0);
          runningMaximum = Math.max(runningMaximum, 0);
          break;
        case ValueTag.Error:
          runningErrorCode ||= deltaSlice.errors[offset] ?? ErrorCode.None;
          runningErrorCount += 1;
          break;
        case ValueTag.String:
        default:
          break;
      }
      const targetOffset = currentLength + offset;
      existing.prefixSums[targetOffset] = runningSum;
      existing.prefixCount[targetOffset] = runningCount;
      existing.prefixAverageCount[targetOffset] = runningAverageCount;
      existing.prefixErrorCodes[targetOffset] = runningErrorCode;
      existing.prefixErrorCounts[targetOffset] = runningErrorCount;
      existing.prefixMinimums[targetOffset] = runningMinimum;
      existing.prefixMaximums[targetOffset] = runningMaximum;
    }
    existing.rowEnd = request.rowEnd;
    existing.columnVersion = deltaSlice.columnVersion;
    existing.structureVersion = deltaSlice.structureVersion;
    return existing;
  };

  return {
    getOrBuildPrefix(request) {
      const key = cacheKey(request.sheetName, request.col, request.rowStart);
      const currentVersions = getCurrentVersions(request.sheetName, request.col);
      const existing = cache.get(key);
      if (
        existing &&
        existing.rowEnd >= request.rowEnd &&
        existing.columnVersion === currentVersions.columnVersion &&
        existing.structureVersion === currentVersions.structureVersion
      ) {
        return existing;
      }
      if (
        existing &&
        existing.rowEnd < request.rowEnd &&
        existing.columnVersion === currentVersions.columnVersion &&
        existing.structureVersion === currentVersions.structureVersion
      ) {
        const extended = extendPrefix(existing, request);
        cache.set(key, extended);
        return extended;
      }
      const built = buildPrefix(request);
      cache.set(key, built);
      return built;
    },
  };
}
