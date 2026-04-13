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
}

interface ApproximateColumnIndexEntry {
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

export function createSortedColumnSearchService(args: {
  readonly state: Pick<EngineRuntimeState, "workbook" | "strings">;
  readonly runtimeColumnStore: EngineRuntimeColumnStoreService;
}): SortedColumnSearchService {
  const approximateColumnIndices = new Map<string, ApproximateColumnIndexEntry>();

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
    const currentSlice = args.runtimeColumnStore.getColumnSlice({
      sheetName,
      rowStart,
      rowEnd,
      col,
    });
    let entry = approximateColumnIndices.get(cacheKey);
    if (
      !entry ||
      entry.columnVersion !== currentSlice.columnVersion ||
      entry.structureVersion !== currentSlice.structureVersion
    ) {
      entry = buildApproximateColumnIndex(sheetName, col, rowStart, rowEnd);
      approximateColumnIndices.set(cacheKey, entry);
    }
    return entry;
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
      sortedAscending: entry.sortedAscending,
      sortedDescending: entry.sortedDescending,
      numericValues: entry.numericValues,
      textValues: entry.textValues,
    };
  };

  const refreshPreparedVectorLookup = (
    prepared: PreparedApproximateVectorLookup,
  ): PreparedApproximateVectorLookup => {
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
  };
}
