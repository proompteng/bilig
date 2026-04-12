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

export interface PreparedExactVectorLookup {
  sheetName: string;
  rowStart: number;
  rowEnd: number;
  col: number;
  columnVersion: number;
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

export interface PreparedApproximateVectorLookup {
  sheetName: string;
  rowStart: number;
  rowEnd: number;
  col: number;
  columnVersion: number;
  comparableKind: "numeric" | "text" | undefined;
  uniformStart: number | undefined;
  uniformStep: number | undefined;
  sortedAscending: boolean;
  sortedDescending: boolean;
  numericValues: Float64Array | undefined;
  textValues: string[] | undefined;
}

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
  readonly prepareExactVectorLookup: (request: {
    sheetName: string;
    rowStart: number;
    rowEnd: number;
    col: number;
  }) => PreparedExactVectorLookup;
  readonly findPreparedExactVectorMatch: (request: {
    lookupValue: CellValue;
    prepared: PreparedExactVectorLookup;
    searchMode: 1 | -1;
  }) => ExactVectorMatchResult;
  readonly prepareApproximateVectorLookup: (request: {
    sheetName: string;
    rowStart: number;
    rowEnd: number;
    col: number;
  }) => PreparedApproximateVectorLookup;
  readonly findPreparedApproximateVectorMatch: (request: {
    lookupValue: CellValue;
    prepared: PreparedApproximateVectorLookup;
    matchMode: 1 | -1;
  }) => ApproximateVectorMatchResult;
}

interface ExactColumnIndexEntry {
  columnVersion: number;
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

type ApproximateComparableValue =
  | { kind: "empty" }
  | { kind: "numeric"; value: number }
  | { kind: "text"; value: string }
  | { kind: "invalid" };

interface ApproximateColumnIndexEntry {
  columnVersion: number;
  comparableKind: "numeric" | "text" | undefined;
  uniformStart: number | undefined;
  uniformStep: number | undefined;
  sortedAscending: boolean;
  sortedDescending: boolean;
  numericValues: Float64Array | undefined;
  textValues: string[] | undefined;
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

function normalizeLookupText(
  value: Extract<CellValue, { tag: ValueTag.String }>,
  lookupString: (id: number) => string,
): string {
  return (
    value.stringId !== undefined && value.stringId !== 0
      ? lookupString(value.stringId)
      : value.value
  ).toUpperCase();
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
    const firstNumericPositions = new Map<number, number>();
    const lastNumericPositions = new Map<number, number>();
    const firstTextPositions = new Map<string, number>();
    const lastTextPositions = new Map<string, number>();
    const numericSequence: number[] = [];
    let sawNumeric = false;
    let sawText = false;
    let sawOther = false;
    for (let row = rowStart; row <= rowEnd; row += 1) {
      const key = readNormalizedKeyAt(sheetName, row, col);
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
      columnVersion: args.state.workbook.getSheetColumnVersion(sheetName, col),
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

  const buildApproximateColumnIndex = (
    sheetName: string,
    col: number,
    rowStart: number,
    rowEnd: number,
  ): ApproximateColumnIndexEntry => {
    const sheet = args.state.workbook.getSheet(sheetName);
    const cellStore = args.state.workbook.cellStore;
    const rawValues: ApproximateComparableValue[] = [];
    let hasNumeric = false;
    let hasText = false;
    let hasInvalid = false;
    for (let row = rowStart; row <= rowEnd; row += 1) {
      if (!sheet) {
        rawValues.push({ kind: "empty" });
        continue;
      }
      const cellIndex = sheet.grid.get(row, col);
      if (cellIndex === -1) {
        rawValues.push({ kind: "empty" });
        continue;
      }
      const tag = cellStore.tags[cellIndex];
      let comparable: ApproximateComparableValue;
      switch (tag) {
        case undefined:
        case ValueTag.Empty:
          comparable = { kind: "empty" };
          break;
        case ValueTag.Number:
          comparable = {
            kind: "numeric",
            value: Object.is(cellStore.numbers[cellIndex]!, -0) ? 0 : cellStore.numbers[cellIndex]!,
          };
          break;
        case ValueTag.Boolean:
          comparable = { kind: "numeric", value: cellStore.numbers[cellIndex]! !== 0 ? 1 : 0 };
          break;
        case ValueTag.String: {
          const stringId = cellStore.stringIds[cellIndex]!;
          let normalized = normalizedStrings.get(stringId);
          if (normalized === undefined) {
            normalized = args.state.strings.get(stringId).toUpperCase();
            normalizedStrings.set(stringId, normalized);
          }
          comparable = { kind: "text", value: normalized };
          break;
        }
        case ValueTag.Error:
        default:
          comparable = { kind: "invalid" };
          break;
      }
      rawValues.push(comparable);
      hasNumeric ||= comparable.kind === "numeric";
      hasText ||= comparable.kind === "text";
      hasInvalid ||= comparable.kind === "invalid";
    }

    if (hasInvalid || (hasNumeric && hasText)) {
      return {
        columnVersion: args.state.workbook.getSheetColumnVersion(sheetName, col),
        comparableKind: undefined,
        uniformStart: undefined,
        uniformStep: undefined,
        sortedAscending: false,
        sortedDescending: false,
        numericValues: undefined,
        textValues: undefined,
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
        uniformStart: undefined,
        uniformStep: undefined,
        sortedAscending,
        sortedDescending,
        numericValues: undefined,
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
    const uniformNumericStep = detectUniformNumericStep(numericValues);
    return {
      columnVersion: args.state.workbook.getSheetColumnVersion(sheetName, col),
      comparableKind: "numeric",
      uniformStart: uniformNumericStep?.start,
      uniformStep: uniformNumericStep?.step,
      sortedAscending,
      sortedDescending,
      numericValues,
      textValues: undefined,
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

  const prepareExactVectorLookup = (request: {
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
    return {
      sheetName: request.sheetName,
      rowStart: request.rowStart,
      rowEnd: request.rowEnd,
      col: request.col,
      columnVersion: entry.columnVersion,
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

  const prepareApproximateVectorLookup = (request: {
    sheetName: string;
    rowStart: number;
    rowEnd: number;
    col: number;
  }): PreparedApproximateVectorLookup => {
    const entry = ensureApproximateColumnIndex(
      request.sheetName,
      request.col,
      request.rowStart,
      request.rowEnd,
    );
    return {
      sheetName: request.sheetName,
      rowStart: request.rowStart,
      rowEnd: request.rowEnd,
      col: request.col,
      columnVersion: entry.columnVersion,
      comparableKind: entry.comparableKind,
      uniformStart: entry.uniformStart,
      uniformStep: entry.uniformStep,
      sortedAscending: entry.sortedAscending,
      sortedDescending: entry.sortedDescending,
      numericValues: entry.numericValues,
      textValues: entry.textValues,
    };
  };

  const refreshPreparedExactVectorLookup = (
    prepared: PreparedExactVectorLookup,
  ): PreparedExactVectorLookup => {
    const columnVersion = args.state.workbook.getSheetColumnVersion(
      prepared.sheetName,
      prepared.col,
    );
    if (columnVersion === prepared.columnVersion) {
      return prepared;
    }
    const refreshed = prepareExactVectorLookup(prepared);
    prepared.columnVersion = refreshed.columnVersion;
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

  const refreshPreparedApproximateVectorLookup = (
    prepared: PreparedApproximateVectorLookup,
  ): PreparedApproximateVectorLookup => {
    const columnVersion = args.state.workbook.getSheetColumnVersion(
      prepared.sheetName,
      prepared.col,
    );
    if (columnVersion === prepared.columnVersion) {
      return prepared;
    }
    const refreshed = prepareApproximateVectorLookup(prepared);
    prepared.columnVersion = refreshed.columnVersion;
    prepared.comparableKind = refreshed.comparableKind;
    prepared.uniformStart = refreshed.uniformStart;
    prepared.uniformStep = refreshed.uniformStep;
    prepared.sortedAscending = refreshed.sortedAscending;
    prepared.sortedDescending = refreshed.sortedDescending;
    prepared.numericValues = refreshed.numericValues;
    prepared.textValues = refreshed.textValues;
    return prepared;
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
    prepareExactVectorLookup(request) {
      return prepareExactVectorLookup(request);
    },
    findPreparedExactVectorMatch(request) {
      const prepared = refreshPreparedExactVectorLookup(request.prepared);
      if (prepared.comparableKind === "numeric") {
        if (request.lookupValue.tag === ValueTag.Error) {
          return { handled: false };
        }
        if (request.lookupValue.tag !== ValueTag.Number) {
          return { handled: true, position: undefined };
        }
        const numericValue = Object.is(request.lookupValue.value, -0)
          ? 0
          : request.lookupValue.value;
        if (prepared.uniformStart !== undefined && prepared.uniformStep !== undefined) {
          const relative = (numericValue - prepared.uniformStart) / prepared.uniformStep;
          const position = Number.isInteger(relative) ? relative + 1 : undefined;
          return {
            handled: true,
            position:
              position !== undefined &&
              position >= 1 &&
              position <= prepared.rowEnd - prepared.rowStart + 1
                ? position
                : undefined,
          };
        }
        const numericMap =
          request.searchMode === -1
            ? prepared.lastNumericPositions
            : prepared.firstNumericPositions;
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
        const textValue = normalizeLookupText(request.lookupValue, (id) =>
          args.state.strings.get(id),
        );
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
    },
    prepareApproximateVectorLookup(request) {
      return prepareApproximateVectorLookup(request);
    },
    findPreparedApproximateVectorMatch(request) {
      const prepared = refreshPreparedApproximateVectorLookup(request.prepared);
      if (prepared.comparableKind === undefined) {
        return { handled: false };
      }

      if (request.matchMode === 1 && !prepared.sortedAscending) {
        return { handled: false };
      }
      if (request.matchMode === -1 && !prepared.sortedDescending) {
        return { handled: false };
      }

      if (prepared.comparableKind === "numeric") {
        let lookupValue: number;
        switch (request.lookupValue.tag) {
          case ValueTag.Empty:
            lookupValue = 0;
            break;
          case ValueTag.Number:
            lookupValue = Object.is(request.lookupValue.value, -0) ? 0 : request.lookupValue.value;
            break;
          case ValueTag.Boolean:
            lookupValue = request.lookupValue.value ? 1 : 0;
            break;
          case ValueTag.Error:
            return { handled: false };
          case ValueTag.String:
            return { handled: false };
        }
        const values = prepared.numericValues;
        if (!values) {
          return { handled: false };
        }
        if (prepared.uniformStart !== undefined && prepared.uniformStep !== undefined) {
          const { uniformStart, uniformStep } = prepared;
          const lastValue = uniformStart + uniformStep * (values.length - 1);
          if (request.matchMode === 1 && uniformStep > 0) {
            if (lookupValue < uniformStart) {
              return { handled: true, position: undefined };
            }
            if (lookupValue >= lastValue) {
              return { handled: true, position: values.length };
            }
            const position = Math.floor((lookupValue - uniformStart) / uniformStep) + 1;
            return { handled: true, position: Math.min(values.length, Math.max(1, position)) };
          }
          if (request.matchMode === -1 && uniformStep < 0) {
            if (lookupValue > uniformStart) {
              return { handled: true, position: undefined };
            }
            if (lookupValue <= lastValue) {
              return { handled: true, position: values.length };
            }
            const position = Math.floor((uniformStart - lookupValue) / -uniformStep) + 1;
            return { handled: true, position: Math.min(values.length, Math.max(1, position)) };
          }
        }
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

      let lookupValue: string;
      switch (request.lookupValue.tag) {
        case ValueTag.Empty:
          lookupValue = "";
          break;
        case ValueTag.String:
          lookupValue = normalizeLookupText(request.lookupValue, (id) =>
            args.state.strings.get(id),
          );
          break;
        case ValueTag.Error:
          return { handled: false };
        case ValueTag.Number:
        case ValueTag.Boolean:
          return { handled: false };
      }
      const values = prepared.textValues;
      if (!values) {
        return { handled: false };
      }
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
        const textValue = normalizeLookupText(request.lookupValue, (id) =>
          args.state.strings.get(id),
        );
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
        let lookupValue: number;
        switch (request.lookupValue.tag) {
          case ValueTag.Empty:
            lookupValue = 0;
            break;
          case ValueTag.Number:
            lookupValue = Object.is(request.lookupValue.value, -0) ? 0 : request.lookupValue.value;
            break;
          case ValueTag.Boolean:
            lookupValue = request.lookupValue.value ? 1 : 0;
            break;
          case ValueTag.Error:
            return { handled: false };
          case ValueTag.String:
            return { handled: false };
        }
        const values = entry.numericValues;
        if (!values) {
          return { handled: false };
        }
        if (entry.uniformStart !== undefined && entry.uniformStep !== undefined) {
          const { uniformStart, uniformStep } = entry;
          const lastValue = uniformStart + uniformStep * (values.length - 1);
          if (request.matchMode === 1 && uniformStep > 0) {
            if (lookupValue < uniformStart) {
              return { handled: true, position: undefined };
            }
            if (lookupValue >= lastValue) {
              return { handled: true, position: values.length };
            }
            const position = Math.floor((lookupValue - uniformStart) / uniformStep) + 1;
            return { handled: true, position: Math.min(values.length, Math.max(1, position)) };
          }
          if (request.matchMode === -1 && uniformStep < 0) {
            if (lookupValue > uniformStart) {
              return { handled: true, position: undefined };
            }
            if (lookupValue <= lastValue) {
              return { handled: true, position: values.length };
            }
            const position = Math.floor((uniformStart - lookupValue) / -uniformStep) + 1;
            return { handled: true, position: Math.min(values.length, Math.max(1, position)) };
          }
        }
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
