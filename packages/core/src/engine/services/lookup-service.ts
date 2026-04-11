import { ValueTag, type CellValue } from "@bilig/protocol";
import { parseRangeAddress } from "@bilig/formula";
import type { EngineRuntimeState } from "../runtime-state.js";

interface ExactVectorMatchRequest {
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

interface ApproximateVectorMatchRequest {
  lookupValue: CellValue;
  sheetName: string;
  start: string;
  end: string;
  startRow?: number;
  endRow?: number;
  startCol?: number;
  endCol?: number;
  matchMode: 1 | -1;
}

export type ExactVectorMatchResult =
  | { handled: false }
  | { handled: true; position: number | undefined };

export type ApproximateVectorMatchResult = ExactVectorMatchResult;

export interface EngineLookupService {
  readonly findExactVectorMatch: (request: ExactVectorMatchRequest) => ExactVectorMatchResult;
  readonly findApproximateVectorMatch: (
    request: ApproximateVectorMatchRequest,
  ) => ApproximateVectorMatchResult;
  readonly primeExactColumnIndex: (request: {
    sheetName: string;
    rowStart: number;
    rowEnd: number;
    col: number;
  }) => void;
  readonly primeApproximateColumnIndex: (request: {
    sheetName: string;
    rowStart: number;
    rowEnd: number;
    col: number;
  }) => void;
}

interface ExactColumnIndexEntry {
  columnVersion: number;
  firstPositions: Map<string, number>;
  lastPositions: Map<string, number>;
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

type ApproximateComparableValue =
  | { kind: "empty" }
  | { kind: "numeric"; value: number }
  | { kind: "text"; value: string }
  | { kind: "invalid" };

interface ApproximateColumnIndexEntry {
  columnVersion: number;
  comparableKind: "numeric" | "text" | undefined;
  sortedAscending: boolean;
  sortedDescending: boolean;
  numericValues?: Float64Array;
  textValues?: string[];
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

function normalizeApproximateComparableValue(
  value: CellValue,
  lookupString: (id: number) => string,
  stringId = 0,
): ApproximateComparableValue {
  switch (value.tag) {
    case ValueTag.Empty:
      return { kind: "empty" };
    case ValueTag.Number:
      return { kind: "numeric", value: Object.is(value.value, -0) ? 0 : value.value };
    case ValueTag.Boolean:
      return { kind: "numeric", value: value.value ? 1 : 0 };
    case ValueTag.String:
      return {
        kind: "text",
        value: (stringId !== 0 ? lookupString(stringId) : value.value).toUpperCase(),
      };
    case ValueTag.Error:
      return { kind: "invalid" };
  }
}

function compareApproximateNumeric(left: number, right: number): number {
  if (left === right) {
    return 0;
  }
  return left < right ? -1 : 1;
}

function compareApproximateText(left: string, right: string): number {
  if (left === right) {
    return 0;
  }
  return left < right ? -1 : 1;
}

export function createEngineLookupService(args: {
  readonly state: Pick<EngineRuntimeState, "workbook" | "strings">;
}): EngineLookupService {
  const exactColumnIndices = new Map<string, ExactColumnIndexEntry>();
  const approximateColumnIndices = new Map<string, ApproximateColumnIndexEntry>();
  const normalizedStrings = new Map<number, string>();

  const readNormalizedKeyAt = (sheetName: string, row: number, col: number): string | undefined => {
    const sheet = args.state.workbook.getSheet(sheetName);
    if (!sheet) {
      return "e:";
    }
    const cellIndex = sheet.grid.get(row, col);
    if (cellIndex === -1) {
      return "e:";
    }
    const tag = args.state.workbook.cellStore.tags[cellIndex];
    switch (tag) {
      case undefined:
        return "e:";
      case ValueTag.Empty:
        return "e:";
      case ValueTag.Number: {
        const numeric = args.state.workbook.cellStore.numbers[cellIndex]!;
        return `n:${Object.is(numeric, -0) ? 0 : numeric}`;
      }
      case ValueTag.Boolean:
        return args.state.workbook.cellStore.numbers[cellIndex]! !== 0 ? "b:1" : "b:0";
      case ValueTag.String: {
        const stringId = args.state.workbook.cellStore.stringIds[cellIndex]!;
        let normalized = normalizedStrings.get(stringId);
        if (normalized === undefined) {
          normalized = args.state.strings.get(stringId).toUpperCase();
          normalizedStrings.set(stringId, normalized);
        }
        return `s:${normalized}`;
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
    const firstPositions = new Map<string, number>();
    const lastPositions = new Map<string, number>();
    for (let row = rowStart; row <= rowEnd; row += 1) {
      const key = readNormalizedKeyAt(sheetName, row, col);
      if (key === undefined) {
        continue;
      }
      if (!firstPositions.has(key)) {
        firstPositions.set(key, row);
      }
      lastPositions.set(key, row);
    }
    return {
      columnVersion: args.state.workbook.getSheetColumnVersion(sheetName, col),
      firstPositions,
      lastPositions,
    };
  };

  const buildApproximateColumnIndex = (
    sheetName: string,
    col: number,
    rowStart: number,
    rowEnd: number,
  ): ApproximateColumnIndexEntry => {
    const rawValues: ApproximateComparableValue[] = [];
    let hasNumeric = false;
    let hasText = false;
    let hasInvalid = false;
    for (let row = rowStart; row <= rowEnd; row += 1) {
      const sheet = args.state.workbook.getSheet(sheetName);
      if (!sheet) {
        rawValues.push({ kind: "empty" });
        continue;
      }
      const cellIndex = sheet.grid.get(row, col);
      if (cellIndex === -1) {
        rawValues.push({ kind: "empty" });
        continue;
      }
      const tag = args.state.workbook.cellStore.tags[cellIndex];
      const comparable = normalizeApproximateComparableValue(
        tag === undefined
          ? { tag: ValueTag.Empty }
          : args.state.workbook.cellStore.getValue(cellIndex, (id) => args.state.strings.get(id)),
        (id) => args.state.strings.get(id),
        tag === ValueTag.String ? (args.state.workbook.cellStore.stringIds[cellIndex] ?? 0) : 0,
      );
      rawValues.push(comparable);
      hasNumeric ||= comparable.kind === "numeric";
      hasText ||= comparable.kind === "text";
      hasInvalid ||= comparable.kind === "invalid";
    }

    if (hasInvalid || (hasNumeric && hasText)) {
      return {
        columnVersion: args.state.workbook.getSheetColumnVersion(sheetName, col),
        comparableKind: undefined,
        sortedAscending: false,
        sortedDescending: false,
      };
    }

    if (hasText) {
      const textValues = rawValues.map((value) => (value.kind === "text" ? value.value : ""));
      let sortedAscending = true;
      let sortedDescending = true;
      for (let index = 1; index < textValues.length; index += 1) {
        const comparison = compareApproximateText(textValues[index - 1]!, textValues[index]!);
        if (comparison > 0) {
          sortedAscending = false;
        }
        if (comparison < 0) {
          sortedDescending = false;
        }
      }
      return {
        columnVersion: args.state.workbook.getSheetColumnVersion(sheetName, col),
        comparableKind: "text",
        sortedAscending,
        sortedDescending,
        textValues,
      };
    }

    const numericValues = Float64Array.from(rawValues, (value) =>
      value.kind === "numeric" ? value.value : 0,
    );
    let sortedAscending = true;
    let sortedDescending = true;
    for (let index = 1; index < numericValues.length; index += 1) {
      const comparison = compareApproximateNumeric(
        numericValues[index - 1]!,
        numericValues[index]!,
      );
      if (comparison > 0) {
        sortedAscending = false;
      }
      if (comparison < 0) {
        sortedDescending = false;
      }
    }
    return {
      columnVersion: args.state.workbook.getSheetColumnVersion(sheetName, col),
      comparableKind: "numeric",
      sortedAscending,
      sortedDescending,
      numericValues,
    };
  };

  const resolveExactColumnBounds = (
    request: VectorLookupBoundsRequest,
  ): ExactColumnBounds | undefined => {
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
  };

  const ensureExactColumnIndex = (
    sheetName: string,
    col: number,
    rowStart: number,
    rowEnd: number,
  ): ExactColumnIndexEntry => {
    const cacheKey = getExactColumnCacheKey(sheetName, col, rowStart, rowEnd);
    const columnVersion = args.state.workbook.getSheetColumnVersion(sheetName, col);
    let entry = exactColumnIndices.get(cacheKey);
    if (!entry || entry.columnVersion !== columnVersion) {
      entry = buildExactColumnIndex(sheetName, col, rowStart, rowEnd);
      exactColumnIndices.set(cacheKey, entry);
    }
    return entry;
  };

  const ensureApproximateColumnIndex = (
    sheetName: string,
    col: number,
    rowStart: number,
    rowEnd: number,
  ): ApproximateColumnIndexEntry => {
    const cacheKey = getExactColumnCacheKey(sheetName, col, rowStart, rowEnd);
    const columnVersion = args.state.workbook.getSheetColumnVersion(sheetName, col);
    let entry = approximateColumnIndices.get(cacheKey);
    if (!entry || entry.columnVersion !== columnVersion) {
      entry = buildApproximateColumnIndex(sheetName, col, rowStart, rowEnd);
      approximateColumnIndices.set(cacheKey, entry);
    }
    return entry;
  };

  return {
    primeExactColumnIndex(request) {
      ensureExactColumnIndex(request.sheetName, request.col, request.rowStart, request.rowEnd);
    },
    primeApproximateColumnIndex(request) {
      ensureApproximateColumnIndex(
        request.sheetName,
        request.col,
        request.rowStart,
        request.rowEnd,
      );
    },
    findExactVectorMatch(request) {
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

      const row =
        request.searchMode === -1
          ? entry.lastPositions.get(normalizedLookupKey)
          : entry.firstPositions.get(normalizedLookupKey);
      return {
        handled: true,
        position: row === undefined ? undefined : row - bounds.rowStart + 1,
      };
    },
    findApproximateVectorMatch(request) {
      const bounds = resolveExactColumnBounds(request);
      if (!bounds) {
        return { handled: false };
      }

      const entry = ensureApproximateColumnIndex(
        request.sheetName,
        bounds.col,
        bounds.rowStart,
        bounds.rowEnd,
      );
      if (entry.comparableKind === undefined) {
        return { handled: false };
      }

      if (request.matchMode === 1 && !entry.sortedAscending) {
        return { handled: false };
      }
      if (request.matchMode === -1 && !entry.sortedDescending) {
        return { handled: false };
      }

      if (entry.comparableKind === "numeric") {
        const lookup = normalizeApproximateComparableValue(request.lookupValue, (id) =>
          args.state.strings.get(id),
        );
        if (lookup.kind !== "numeric" && lookup.kind !== "empty") {
          return { handled: false };
        }
        const values = entry.numericValues;
        if (!values) {
          return { handled: false };
        }
        const lookupValue = lookup.kind === "numeric" ? lookup.value : 0;
        let low = 0;
        let high = values.length - 1;
        let best = -1;
        while (low <= high) {
          const mid = (low + high) >> 1;
          const comparison = compareApproximateNumeric(values[mid]!, lookupValue);
          if (request.matchMode === 1) {
            if (comparison <= 0) {
              best = mid;
              low = mid + 1;
            } else {
              high = mid - 1;
            }
          } else if (comparison >= 0) {
            best = mid;
            low = mid + 1;
          } else {
            high = mid - 1;
          }
        }
        return {
          handled: true,
          position: best === -1 ? undefined : best + 1,
        };
      }

      const lookup = normalizeApproximateComparableValue(request.lookupValue, (id) =>
        args.state.strings.get(id),
      );
      if (lookup.kind !== "text" && lookup.kind !== "empty") {
        return { handled: false };
      }
      const values = entry.textValues;
      if (!values) {
        return { handled: false };
      }
      const lookupValue = lookup.kind === "text" ? lookup.value : "";
      let low = 0;
      let high = values.length - 1;
      let best = -1;
      while (low <= high) {
        const mid = (low + high) >> 1;
        const comparison = compareApproximateText(values[mid]!, lookupValue);
        if (request.matchMode === 1) {
          if (comparison <= 0) {
            best = mid;
            low = mid + 1;
          } else {
            high = mid - 1;
          }
        } else if (comparison >= 0) {
          best = mid;
          low = mid + 1;
        } else {
          high = mid - 1;
        }
      }
      return {
        handled: true,
        position: best === -1 ? undefined : best + 1,
      };
    },
  };
}
