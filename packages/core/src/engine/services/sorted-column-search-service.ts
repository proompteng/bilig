import { ValueTag, type CellValue } from "@bilig/protocol";
import { parseRangeAddress } from "@bilig/formula";
import type { EngineRuntimeState, PreparedApproximateVectorLookup } from "../runtime-state.js";
import type { ExactVectorMatchResult } from "./exact-column-index-service.js";
import type {
  EngineRuntimeColumnStoreService,
  RuntimeColumnSlice,
} from "./runtime-column-store-service.js";

export interface ApproximateVectorMatchRequest {
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

export type ApproximateVectorMatchResult = ExactVectorMatchResult;

export interface SortedColumnSearchService {
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
  }) => PreparedApproximateVectorLookup;
  readonly findPreparedVectorMatch: (request: {
    lookupValue: CellValue;
    prepared: PreparedApproximateVectorLookup;
    matchMode: 1 | -1;
  }) => ApproximateVectorMatchResult;
  readonly findVectorMatch: (
    request: ApproximateVectorMatchRequest,
  ) => ApproximateVectorMatchResult;
  readonly invalidateColumn: (request: { sheetName: string; col: number }) => void;
  readonly recordLiteralWrite: (request: {
    sheetName: string;
    row: number;
    col: number;
    oldValue: CellValue;
    newValue: CellValue;
    oldStringId?: number;
    newStringId?: number;
  }) => void;
}

interface ApproximateColumnIndexEntry {
  sheetName: string;
  rowStart: number;
  rowEnd: number;
  col: number;
  columnVersion: number;
  structureVersion: number;
  comparableKind: "numeric" | "text" | undefined;
  uniformStart: number | undefined;
  uniformStep: number | undefined;
  sortedAscending: boolean;
  sortedDescending: boolean;
  numericValues: Float64Array | undefined;
  textValues: string[] | undefined;
}

type ApproximateComparable =
  | { kind: "empty" }
  | { kind: "numeric"; value: number }
  | { kind: "text"; value: string }
  | { kind: "invalid" };

interface SingleColumnBounds {
  rowStart: number;
  rowEnd: number;
  col: number;
}

function getColumnCacheKey(
  sheetName: string,
  col: number,
  rowStart: number,
  rowEnd: number,
): string {
  return `${sheetName}\t${col}\t${rowStart}\t${rowEnd}`;
}

function normalizeApproximateComparableValue(
  value: CellValue,
  lookupString: (id: number) => string,
  stringId = 0,
): ApproximateComparable {
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

function resolveSingleColumnBounds(
  request: Pick<
    ApproximateVectorMatchRequest,
    "sheetName" | "start" | "end" | "startRow" | "endRow" | "startCol" | "endCol"
  >,
): SingleColumnBounds | undefined {
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

function columnRegistryKey(sheetName: string, col: number): string {
  return `${sheetName}\t${col}`;
}

export function createSortedColumnSearchService(args: {
  readonly state: Pick<EngineRuntimeState, "workbook" | "strings">;
  readonly runtimeColumnStore: EngineRuntimeColumnStoreService;
}): SortedColumnSearchService {
  const emptyColumnVersions = new Uint32Array(0);
  const approximateColumnIndices = new Map<string, ApproximateColumnIndexEntry>();
  const cacheKeysByColumn = new Map<string, Set<string>>();

  const getCurrentColumnVersions = (
    sheetName: string,
    col: number,
  ): {
    columnVersion: number;
    structureVersion: number;
    sheetColumnVersions: Uint32Array;
  } => {
    const sheet = args.state.workbook.getSheet(sheetName);
    const sheetColumnVersions = sheet?.columnVersions ?? emptyColumnVersions;
    return {
      columnVersion: sheetColumnVersions[col] ?? 0,
      structureVersion: sheet?.structureVersion ?? 0,
      sheetColumnVersions,
    };
  };

  const trackCacheKey = (sheetName: string, col: number, cacheKey: string): void => {
    const registryKey = columnRegistryKey(sheetName, col);
    const existing = cacheKeysByColumn.get(registryKey);
    if (existing) {
      existing.add(cacheKey);
      return;
    }
    cacheKeysByColumn.set(registryKey, new Set([cacheKey]));
  };

  const untrackCacheKey = (sheetName: string, col: number, cacheKey: string): void => {
    const registryKey = columnRegistryKey(sheetName, col);
    const existing = cacheKeysByColumn.get(registryKey);
    if (!existing) {
      return;
    }
    existing.delete(cacheKey);
    if (existing.size === 0) {
      cacheKeysByColumn.delete(registryKey);
    }
  };

  const replaceColumnIndex = (
    cacheKey: string,
    entry: ApproximateColumnIndexEntry | undefined,
  ): void => {
    const existing = approximateColumnIndices.get(cacheKey);
    if (existing) {
      untrackCacheKey(existing.sheetName, existing.col, cacheKey);
      approximateColumnIndices.delete(cacheKey);
    }
    if (!entry) {
      return;
    }
    approximateColumnIndices.set(cacheKey, entry);
    trackCacheKey(entry.sheetName, entry.col, cacheKey);
  };

  const comparableAtOffset = (slice: RuntimeColumnSlice, offset: number): ApproximateComparable => {
    const tag = decodeValueTag(slice.tags[offset]);
    switch (tag) {
      case ValueTag.Empty:
        return { kind: "empty" };
      case ValueTag.Number:
        return { kind: "numeric", value: slice.numbers[offset] ?? 0 };
      case ValueTag.Boolean:
        return { kind: "numeric", value: (slice.numbers[offset] ?? 0) !== 0 ? 1 : 0 };
      case ValueTag.String:
        return {
          kind: "text",
          value: args.runtimeColumnStore.normalizeStringId(slice.stringIds[offset] ?? 0),
        };
      case ValueTag.Error:
      default:
        return { kind: "invalid" };
    }
  };

  const buildApproximateColumnIndex = (
    sheetName: string,
    col: number,
    rowStart: number,
    rowEnd: number,
  ): ApproximateColumnIndexEntry => {
    const slice = args.runtimeColumnStore.getColumnSlice({
      sheetName,
      rowStart,
      rowEnd,
      col,
    });
    const rawValues: ApproximateComparable[] = [];
    let hasNumeric = false;
    let hasText = false;
    let hasInvalid = false;
    for (let offset = 0; offset < slice.length; offset += 1) {
      const comparable = comparableAtOffset(slice, offset);
      rawValues.push(comparable);
      hasNumeric ||= comparable.kind === "numeric";
      hasText ||= comparable.kind === "text";
      hasInvalid ||= comparable.kind === "invalid";
    }

    if (hasInvalid || (hasNumeric && hasText)) {
      return {
        sheetName,
        rowStart,
        rowEnd,
        col,
        columnVersion: slice.columnVersion,
        structureVersion: slice.structureVersion,
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
        sheetName,
        rowStart,
        rowEnd,
        col,
        columnVersion: slice.columnVersion,
        structureVersion: slice.structureVersion,
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
      sheetName,
      rowStart,
      rowEnd,
      col,
      columnVersion: slice.columnVersion,
      structureVersion: slice.structureVersion,
      comparableKind: "numeric",
      uniformStart: uniformNumericStep?.start,
      uniformStep: uniformNumericStep?.step,
      sortedAscending,
      sortedDescending,
      numericValues,
      textValues: undefined,
    };
  };

  const ensureColumnIndex = (
    sheetName: string,
    col: number,
    rowStart: number,
    rowEnd: number,
  ): ApproximateColumnIndexEntry => {
    const cacheKey = getColumnCacheKey(sheetName, col, rowStart, rowEnd);
    const currentVersions = getCurrentColumnVersions(sheetName, col);
    let entry = approximateColumnIndices.get(cacheKey);
    if (
      !entry ||
      entry.columnVersion !== currentVersions.columnVersion ||
      entry.structureVersion !== currentVersions.structureVersion
    ) {
      entry = buildApproximateColumnIndex(sheetName, col, rowStart, rowEnd);
      replaceColumnIndex(cacheKey, entry);
    }
    return entry;
  };

  const updateEntryLiteralWrite = (
    entry: ApproximateColumnIndexEntry,
    row: number,
    oldValue: ApproximateComparable,
    newValue: ApproximateComparable,
    currentColumnVersion: number,
    currentStructureVersion: number,
  ): boolean => {
    entry.columnVersion = currentColumnVersion;
    entry.structureVersion = currentStructureVersion;
    if (row < entry.rowStart || row > entry.rowEnd) {
      return true;
    }
    if (oldValue.kind !== newValue.kind) {
      return false;
    }
    if (oldValue.kind === "invalid" || newValue.kind === "invalid") {
      return false;
    }
    if (entry.comparableKind === undefined) {
      return false;
    }
    if (
      (entry.comparableKind === "numeric" &&
        newValue.kind !== "numeric" &&
        newValue.kind !== "empty") ||
      (entry.comparableKind === "text" && newValue.kind !== "text" && newValue.kind !== "empty")
    ) {
      return false;
    }

    const offset = row - entry.rowStart;
    if (entry.comparableKind === "numeric") {
      if (!entry.numericValues) {
        return false;
      }
      const nextNumeric = newValue.kind === "numeric" ? newValue.value : 0;
      entry.numericValues[offset] = nextNumeric;
      entry.uniformStart = undefined;
      entry.uniformStep = undefined;
      const previous = offset > 0 ? entry.numericValues[offset - 1]! : undefined;
      const next =
        offset + 1 < entry.numericValues.length ? entry.numericValues[offset + 1]! : undefined;
      if (
        (previous !== undefined && compareApproximateNumeric(previous, nextNumeric) > 0) ||
        (next !== undefined && compareApproximateNumeric(nextNumeric, next) > 0)
      ) {
        entry.sortedAscending = false;
      }
      if (
        (previous !== undefined && compareApproximateNumeric(previous, nextNumeric) < 0) ||
        (next !== undefined && compareApproximateNumeric(nextNumeric, next) < 0)
      ) {
        entry.sortedDescending = false;
      }
      return true;
    }

    if (!entry.textValues) {
      return false;
    }
    const nextText = newValue.kind === "text" ? newValue.value : "";
    entry.textValues[offset] = nextText;
    const previous = offset > 0 ? entry.textValues[offset - 1]! : undefined;
    const next = offset + 1 < entry.textValues.length ? entry.textValues[offset + 1]! : undefined;
    if (
      (previous !== undefined && compareApproximateText(previous, nextText) > 0) ||
      (next !== undefined && compareApproximateText(nextText, next) > 0)
    ) {
      entry.sortedAscending = false;
    }
    if (
      (previous !== undefined && compareApproximateText(previous, nextText) < 0) ||
      (next !== undefined && compareApproximateText(nextText, next) < 0)
    ) {
      entry.sortedDescending = false;
    }
    return true;
  };

  const prepareVectorLookup = (request: {
    sheetName: string;
    rowStart: number;
    rowEnd: number;
    col: number;
  }): PreparedApproximateVectorLookup => {
    const entry = ensureColumnIndex(
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
      length: request.rowEnd - request.rowStart + 1,
      columnVersion: entry.columnVersion,
      structureVersion: entry.structureVersion,
      sheetColumnVersions: getCurrentColumnVersions(request.sheetName, request.col)
        .sheetColumnVersions,
      comparableKind: entry.comparableKind,
      uniformStart: entry.uniformStart,
      uniformStep: entry.uniformStep,
      sortedAscending: entry.sortedAscending,
      sortedDescending: entry.sortedDescending,
      numericValues: entry.numericValues,
      textValues: entry.textValues,
    };
  };

  const refreshPreparedVectorLookup = (
    prepared: PreparedApproximateVectorLookup,
  ): PreparedApproximateVectorLookup => {
    const currentVersions = getCurrentColumnVersions(prepared.sheetName, prepared.col);
    if (
      currentVersions.columnVersion === prepared.columnVersion &&
      currentVersions.structureVersion === prepared.structureVersion
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
    prepared.sortedAscending = refreshed.sortedAscending;
    prepared.sortedDescending = refreshed.sortedDescending;
    prepared.numericValues = refreshed.numericValues;
    prepared.textValues = refreshed.textValues;
    return prepared;
  };

  const findPreparedVectorMatch = (request: {
    lookupValue: CellValue;
    prepared: PreparedApproximateVectorLookup;
    matchMode: 1 | -1;
  }): ApproximateVectorMatchResult => {
    const prepared = refreshPreparedVectorLookup(request.prepared);
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
      return { handled: true, position: best === -1 ? undefined : best + 1 };
    }

    let lookupValue: string;
    switch (request.lookupValue.tag) {
      case ValueTag.Empty:
        lookupValue = "";
        break;
      case ValueTag.String:
        lookupValue = args.runtimeColumnStore.normalizeLookupText(request.lookupValue);
        break;
      case ValueTag.Error:
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
    return { handled: true, position: best === -1 ? undefined : best + 1 };
  };

  return {
    primeColumnIndex(request) {
      ensureColumnIndex(request.sheetName, request.col, request.rowStart, request.rowEnd);
    },
    prepareVectorLookup(request) {
      return prepareVectorLookup(request);
    },
    findPreparedVectorMatch(request) {
      return findPreparedVectorMatch(request);
    },
    findVectorMatch(request) {
      const bounds = resolveSingleColumnBounds(request);
      if (!bounds) {
        return { handled: false };
      }

      const entry = ensureColumnIndex(
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
        return { handled: true, position: best === -1 ? undefined : best + 1 };
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
      return { handled: true, position: best === -1 ? undefined : best + 1 };
    },
    invalidateColumn(request) {
      const cacheKeys = cacheKeysByColumn.get(columnRegistryKey(request.sheetName, request.col));
      if (!cacheKeys) {
        return;
      }
      for (const cacheKey of cacheKeys) {
        replaceColumnIndex(cacheKey, undefined);
      }
    },
    recordLiteralWrite(request) {
      const cacheKeys = cacheKeysByColumn.get(columnRegistryKey(request.sheetName, request.col));
      if (!cacheKeys || cacheKeys.size === 0) {
        return;
      }
      const currentVersions = getCurrentColumnVersions(request.sheetName, request.col);
      const oldComparable = normalizeApproximateComparableValue(
        request.oldValue,
        (id) => args.state.strings.get(id),
        request.oldStringId,
      );
      const newComparable = normalizeApproximateComparableValue(
        request.newValue,
        (id) => args.state.strings.get(id),
        request.newStringId,
      );
      for (const cacheKey of cacheKeys) {
        const entry = approximateColumnIndices.get(cacheKey);
        if (!entry) {
          continue;
        }
        if (entry.structureVersion !== currentVersions.structureVersion) {
          replaceColumnIndex(cacheKey, undefined);
          continue;
        }
        if (
          !updateEntryLiteralWrite(
            entry,
            request.row,
            oldComparable,
            newComparable,
            currentVersions.columnVersion,
            currentVersions.structureVersion,
          )
        ) {
          replaceColumnIndex(cacheKey, undefined);
        }
      }
    },
  };
}
