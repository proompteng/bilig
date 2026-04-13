import { ValueTag, type CellValue } from "@bilig/protocol";
import { parseRangeAddress } from "@bilig/formula";
import type { EngineRuntimeState, PreparedExactVectorLookup } from "../runtime-state.js";
import type {
  EngineRuntimeColumnStoreService,
  RuntimeColumnSlice,
} from "./runtime-column-store-service.js";

export interface ExactVectorMatchRequest {
  lookupValue: CellValue;
  sheetName: string;
  start: string;
  end: string;
  startRow?: number;
  endRow?: number;
  startCol?: number;
  endCol?: number;
  searchMode: 1 | -1;
}

export type ExactVectorMatchResult =
  | { handled: false }
  | { handled: true; position: number | undefined };

export interface ExactColumnIndexService {
  readonly primeColumnIndex: (request: {
    sheetName: string;
    rowStart: number;
    rowEnd: number;
    col: number;
  }) => void;
  readonly prepareVectorLookup: (request: {
    sheetName: string;
    rowStart: number;
    rowEnd: number;
    col: number;
  }) => PreparedExactVectorLookup;
  readonly findPreparedVectorMatch: (request: {
    lookupValue: CellValue;
    prepared: PreparedExactVectorLookup;
    searchMode: 1 | -1;
  }) => ExactVectorMatchResult;
  readonly findVectorMatch: (request: ExactVectorMatchRequest) => ExactVectorMatchResult;
}

interface ExactColumnIndexEntry {
  columnVersion: number;
  structureVersion: number;
  comparableKind: "numeric" | "text" | "mixed";
  uniformStart: number | undefined;
  uniformStep: number | undefined;
  firstPositions: Map<string, number>;
  lastPositions: Map<string, number>;
  firstNumericPositions: Map<number, number> | undefined;
  lastNumericPositions: Map<number, number> | undefined;
  firstTextPositions: Map<string, number> | undefined;
  lastTextPositions: Map<string, number> | undefined;
}

interface ExactColumnBounds {
  rowStart: number;
  rowEnd: number;
  col: number;
}

interface VectorLookupBoundsRequest {
  sheetName: string;
  start: string;
  end: string;
  startRow?: number;
  endRow?: number;
  startCol?: number;
  endCol?: number;
}

function getExactColumnCacheKey(
  sheetName: string,
  col: number,
  rowStart: number,
  rowEnd: number,
): string {
  return `${sheetName}\t${col}\t${rowStart}\t${rowEnd}`;
}

function normalizeExactLookupKey(
  value: CellValue,
  lookupString: (id: number) => string,
  stringId = 0,
): string | undefined {
  switch (value.tag) {
    case ValueTag.Empty:
      return "e:";
    case ValueTag.Number:
      return `n:${Object.is(value.value, -0) ? 0 : value.value}`;
    case ValueTag.Boolean:
      return value.value ? "b:1" : "b:0";
    case ValueTag.String:
      return `s:${(stringId !== 0 ? lookupString(stringId) : value.value).toUpperCase()}`;
    case ValueTag.Error:
      return undefined;
  }
}

function detectUniformNumericStep(
  values: Float64Array,
): { start: number; step: number } | undefined {
  if (values.length < 2) {
    return undefined;
  }
  const start = values[0]!;
  const step = values[1]! - start;
  if (!Number.isFinite(step) || step === 0) {
    return undefined;
  }
  for (let index = 2; index < values.length; index += 1) {
    if (values[index]! - values[index - 1]! !== step) {
      return undefined;
    }
  }
  return { start, step };
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

function resolveExactColumnBounds(
  request: VectorLookupBoundsRequest,
): ExactColumnBounds | undefined {
  if (
    request.startRow !== undefined &&
    request.endRow !== undefined &&
    request.startCol !== undefined &&
    request.endCol !== undefined
  ) {
    if (request.startCol !== request.endCol) {
      return undefined;
    }
    return {
      rowStart: request.startRow,
      rowEnd: request.endRow,
      col: request.startCol,
    };
  }

  const parsedRange = parseRangeAddress(`${request.start}:${request.end}`, request.sheetName);
  if (parsedRange.kind !== "cells" || parsedRange.start.col !== parsedRange.end.col) {
    return undefined;
  }
  return {
    rowStart: parsedRange.start.row,
    rowEnd: parsedRange.end.row,
    col: parsedRange.start.col,
  };
}

export function createExactColumnIndexService(args: {
  readonly state: Pick<EngineRuntimeState, "workbook" | "strings">;
  readonly runtimeColumnStore: EngineRuntimeColumnStoreService;
}): ExactColumnIndexService {
  const exactColumnIndices = new Map<string, ExactColumnIndexEntry>();

  const keyAtOffset = (slice: RuntimeColumnSlice, offset: number): string | undefined => {
    const tag = decodeValueTag(slice.tags[offset]);
    switch (tag) {
      case ValueTag.Empty:
        return "e:";
      case ValueTag.Number:
        return `n:${slice.numbers[offset] ?? 0}`;
      case ValueTag.Boolean:
        return (slice.numbers[offset] ?? 0) !== 0 ? "b:1" : "b:0";
      case ValueTag.String: {
        const stringId = slice.stringIds[offset] ?? 0;
        return `s:${args.runtimeColumnStore.normalizeStringId(stringId)}`;
      }
      case ValueTag.Error:
      default:
        return undefined;
    }
  };

  const buildExactColumnIndex = (
    sheetName: string,
    col: number,
    rowStart: number,
    rowEnd: number,
  ): ExactColumnIndexEntry => {
    const slice = args.runtimeColumnStore.getColumnSlice({
      sheetName,
      rowStart,
      rowEnd,
      col,
    });
    const firstPositions = new Map<string, number>();
    const lastPositions = new Map<string, number>();
    const firstNumericPositions = new Map<number, number>();
    const lastNumericPositions = new Map<number, number>();
    const firstTextPositions = new Map<string, number>();
    const lastTextPositions = new Map<string, number>();
    const numericSequence: number[] = [];
    let sawNumeric = false;
    let sawText = false;
    let sawOther = false;
    for (let offset = 0; offset < slice.length; offset += 1) {
      const row = rowStart + offset;
      const key = keyAtOffset(slice, offset);
      if (key === undefined) {
        continue;
      }
      if (!firstPositions.has(key)) {
        firstPositions.set(key, row);
      }
      lastPositions.set(key, row);
      if (key.startsWith("n:")) {
        const numericValue = Number(key.slice(2));
        if (!firstNumericPositions.has(numericValue)) {
          firstNumericPositions.set(numericValue, row);
        }
        lastNumericPositions.set(numericValue, row);
        numericSequence.push(numericValue);
        sawNumeric = true;
        continue;
      }
      if (key.startsWith("s:")) {
        const textValue = key.slice(2);
        if (!firstTextPositions.has(textValue)) {
          firstTextPositions.set(textValue, row);
        }
        lastTextPositions.set(textValue, row);
        sawText = true;
        continue;
      }
      sawOther = true;
    }
    const comparableKind =
      sawOther || (sawNumeric && sawText)
        ? "mixed"
        : sawNumeric
          ? "numeric"
          : sawText
            ? "text"
            : "mixed";
    const uniformNumericStep =
      comparableKind === "numeric"
        ? detectUniformNumericStep(Float64Array.from(numericSequence))
        : undefined;
    return {
      columnVersion: slice.columnVersion,
      structureVersion: slice.structureVersion,
      comparableKind,
      uniformStart: uniformNumericStep?.start,
      uniformStep: uniformNumericStep?.step,
      firstPositions,
      lastPositions,
      firstNumericPositions: comparableKind === "numeric" ? firstNumericPositions : undefined,
      lastNumericPositions: comparableKind === "numeric" ? lastNumericPositions : undefined,
      firstTextPositions: comparableKind === "text" ? firstTextPositions : undefined,
      lastTextPositions: comparableKind === "text" ? lastTextPositions : undefined,
    };
  };

  const ensureExactColumnIndex = (
    sheetName: string,
    col: number,
    rowStart: number,
    rowEnd: number,
  ): ExactColumnIndexEntry => {
    const cacheKey = getExactColumnCacheKey(sheetName, col, rowStart, rowEnd);
    const columnVersion = args.runtimeColumnStore.getColumnSlice({
      sheetName,
      rowStart,
      rowEnd,
      col,
    });
    let entry = exactColumnIndices.get(cacheKey);
    if (
      !entry ||
      entry.columnVersion !== columnVersion.columnVersion ||
      entry.structureVersion !== columnVersion.structureVersion
    ) {
      entry = buildExactColumnIndex(sheetName, col, rowStart, rowEnd);
      exactColumnIndices.set(cacheKey, entry);
    }
    return entry;
  };

  const prepareVectorLookup = (request: {
    sheetName: string;
    rowStart: number;
    rowEnd: number;
    col: number;
  }): PreparedExactVectorLookup => {
    const entry = ensureExactColumnIndex(
      request.sheetName,
      request.col,
      request.rowStart,
      request.rowEnd,
    );
    const columnSlice = args.runtimeColumnStore.getColumnSlice(request);
    return {
      sheetName: request.sheetName,
      rowStart: request.rowStart,
      rowEnd: request.rowEnd,
      col: request.col,
      length: request.rowEnd - request.rowStart + 1,
      columnVersion: entry.columnVersion,
      structureVersion: entry.structureVersion,
      sheetColumnVersions: columnSlice.sheetColumnVersions,
      comparableKind: entry.comparableKind,
      uniformStart: entry.uniformStart,
      uniformStep: entry.uniformStep,
      firstPositions: entry.firstPositions,
      lastPositions: entry.lastPositions,
      firstNumericPositions: entry.firstNumericPositions,
      lastNumericPositions: entry.lastNumericPositions,
      firstTextPositions: entry.firstTextPositions,
      lastTextPositions: entry.lastTextPositions,
    };
  };

  const refreshPreparedVectorLookup = (
    prepared: PreparedExactVectorLookup,
  ): PreparedExactVectorLookup => {
    const currentSlice = args.runtimeColumnStore.getColumnSlice({
      sheetName: prepared.sheetName,
      rowStart: prepared.rowStart,
      rowEnd: prepared.rowEnd,
      col: prepared.col,
    });
    if (
      currentSlice.columnVersion === prepared.columnVersion &&
      currentSlice.structureVersion === prepared.structureVersion
    ) {
      return prepared;
    }
    const refreshed = prepareVectorLookup(prepared);
    prepared.length = refreshed.length;
    prepared.columnVersion = refreshed.columnVersion;
    prepared.structureVersion = refreshed.structureVersion;
    prepared.sheetColumnVersions = refreshed.sheetColumnVersions;
    prepared.comparableKind = refreshed.comparableKind;
    prepared.uniformStart = refreshed.uniformStart;
    prepared.uniformStep = refreshed.uniformStep;
    prepared.firstPositions = refreshed.firstPositions;
    prepared.lastPositions = refreshed.lastPositions;
    prepared.firstNumericPositions = refreshed.firstNumericPositions;
    prepared.lastNumericPositions = refreshed.lastNumericPositions;
    prepared.firstTextPositions = refreshed.firstTextPositions;
    prepared.lastTextPositions = refreshed.lastTextPositions;
    return prepared;
  };

  const findPreparedVectorMatch = (request: {
    lookupValue: CellValue;
    prepared: PreparedExactVectorLookup;
    searchMode: 1 | -1;
  }): ExactVectorMatchResult => {
    const prepared = refreshPreparedVectorLookup(request.prepared);
    if (prepared.comparableKind === "numeric") {
      if (request.lookupValue.tag === ValueTag.Error) {
        return { handled: false };
      }
      if (request.lookupValue.tag !== ValueTag.Number) {
        return { handled: true, position: undefined };
      }
      const numericValue = Object.is(request.lookupValue.value, -0) ? 0 : request.lookupValue.value;
      if (prepared.uniformStart !== undefined && prepared.uniformStep !== undefined) {
        const relative = (numericValue - prepared.uniformStart) / prepared.uniformStep;
        const position = Number.isInteger(relative) ? relative + 1 : undefined;
        return {
          handled: true,
          position:
            position !== undefined && position >= 1 && position <= prepared.length
              ? position
              : undefined,
        };
      }
      const numericMap =
        request.searchMode === -1 ? prepared.lastNumericPositions : prepared.firstNumericPositions;
      const row = numericMap?.get(numericValue);
      return {
        handled: true,
        position: row === undefined ? undefined : row - prepared.rowStart + 1,
      };
    }
    if (prepared.comparableKind === "text") {
      if (request.lookupValue.tag === ValueTag.Error) {
        return { handled: false };
      }
      if (request.lookupValue.tag !== ValueTag.String) {
        return { handled: true, position: undefined };
      }
      const textValue = args.runtimeColumnStore.normalizeLookupText(request.lookupValue);
      const textMap =
        request.searchMode === -1 ? prepared.lastTextPositions : prepared.firstTextPositions;
      const row = textMap?.get(textValue);
      return {
        handled: true,
        position: row === undefined ? undefined : row - prepared.rowStart + 1,
      };
    }
    const normalizedLookupKey = normalizeExactLookupKey(request.lookupValue, (id) =>
      args.state.strings.get(id),
    );
    if (normalizedLookupKey === undefined) {
      return { handled: false };
    }
    const row =
      request.searchMode === -1
        ? prepared.lastPositions.get(normalizedLookupKey)
        : prepared.firstPositions.get(normalizedLookupKey);
    return {
      handled: true,
      position: row === undefined ? undefined : row - prepared.rowStart + 1,
    };
  };

  return {
    primeColumnIndex(request) {
      ensureExactColumnIndex(request.sheetName, request.col, request.rowStart, request.rowEnd);
    },
    prepareVectorLookup(request) {
      return prepareVectorLookup(request);
    },
    findPreparedVectorMatch(request) {
      return findPreparedVectorMatch(request);
    },
    findVectorMatch(request) {
      const normalizedLookupKey = normalizeExactLookupKey(request.lookupValue, (id) =>
        args.state.strings.get(id),
      );
      if (normalizedLookupKey === undefined) {
        return { handled: false };
      }

      const bounds = resolveExactColumnBounds(request);
      if (!bounds) {
        return { handled: false };
      }

      const entry = ensureExactColumnIndex(
        request.sheetName,
        bounds.col,
        bounds.rowStart,
        bounds.rowEnd,
      );
      if (entry.comparableKind === "numeric") {
        if (request.lookupValue.tag === ValueTag.Error) {
          return { handled: false };
        }
        if (request.lookupValue.tag !== ValueTag.Number) {
          return { handled: true, position: undefined };
        }
        const numericValue = Object.is(request.lookupValue.value, -0)
          ? 0
          : request.lookupValue.value;
        if (entry.uniformStart !== undefined && entry.uniformStep !== undefined) {
          const relative = (numericValue - entry.uniformStart) / entry.uniformStep;
          const position = Number.isInteger(relative) ? relative + 1 : undefined;
          return {
            handled: true,
            position:
              position !== undefined &&
              position >= 1 &&
              position <= bounds.rowEnd - bounds.rowStart + 1
                ? position
                : undefined,
          };
        }
        const numericMap =
          request.searchMode === -1 ? entry.lastNumericPositions : entry.firstNumericPositions;
        const row = numericMap?.get(numericValue);
        return {
          handled: true,
          position: row === undefined ? undefined : row - bounds.rowStart + 1,
        };
      }
      if (entry.comparableKind === "text") {
        if (request.lookupValue.tag === ValueTag.Error) {
          return { handled: false };
        }
        if (request.lookupValue.tag !== ValueTag.String) {
          return { handled: true, position: undefined };
        }
        const textValue = args.runtimeColumnStore.normalizeLookupText(request.lookupValue);
        const textMap =
          request.searchMode === -1 ? entry.lastTextPositions : entry.firstTextPositions;
        const row = textMap?.get(textValue);
        return {
          handled: true,
          position: row === undefined ? undefined : row - bounds.rowStart + 1,
        };
      }

      const row =
        request.searchMode === -1
          ? entry.lastPositions.get(normalizedLookupKey)
          : entry.firstPositions.get(normalizedLookupKey);
      return {
        handled: true,
        position: row === undefined ? undefined : row - bounds.rowStart + 1,
      };
    },
  };
}
