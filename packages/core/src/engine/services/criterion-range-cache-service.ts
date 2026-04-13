import { ErrorCode, ValueTag, type CellValue } from "@bilig/protocol";
import {
  compileCriteriaMatcher,
  matchesCompiledCriteria,
  type CompiledCriteriaMatcher,
  type CriteriaOperator,
} from "@bilig/formula";
import type {
  EngineRuntimeColumnStoreService,
  RuntimeColumnSlice,
} from "./runtime-column-store-service.js";

export interface CriterionRangeDescriptor {
  readonly sheetName: string;
  readonly rowStart: number;
  readonly rowEnd: number;
  readonly col: number;
  readonly length: number;
}

export interface CriterionRangePair {
  readonly range: CriterionRangeDescriptor;
  readonly criteria: CellValue;
}

export interface CriterionRangeMatch {
  readonly rows: Uint32Array;
  readonly length: number;
}

export interface CriterionRangeCacheService {
  readonly getOrBuildMatchingRows: (request: {
    criteriaPairs: readonly CriterionRangePair[];
  }) => CriterionRangeMatch | CellValue;
}

interface CriterionCacheEntry {
  readonly rows: Uint32Array;
  readonly length: number;
  readonly pairVersions: ReadonlyArray<{
    columnVersion: number;
    structureVersion: number;
  }>;
}

type SliceFastPredicate =
  | {
      kind: "eq-empty";
      negate: boolean;
    }
  | {
      kind: "eq-bool";
      negate: boolean;
      value: boolean;
    }
  | {
      kind: "eq-number";
      negate: boolean;
      value: number;
    }
  | {
      kind: "eq-string";
      negate: boolean;
      value: string;
    }
  | {
      kind: "cmp-number";
      operator: Exclude<CriteriaOperator, "=" | "<>">;
      value: number;
    }
  | {
      kind: "generic";
      compiled: CompiledCriteriaMatcher;
    };

function errorValue(code: ErrorCode): CellValue {
  return { tag: ValueTag.Error, code };
}

function criteriaCacheKey(value: CellValue): string {
  switch (value.tag) {
    case ValueTag.Empty:
      return "e:";
    case ValueTag.Number:
      return `n:${Object.is(value.value, -0) ? 0 : value.value}`;
    case ValueTag.Boolean:
      return value.value ? "b:1" : "b:0";
    case ValueTag.String:
      return `s:${value.value}`;
    case ValueTag.Error:
      return `r:${value.code}`;
  }
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

function normalizeSliceString(
  runtimeColumnStore: EngineRuntimeColumnStoreService,
  stringId: number,
): string {
  return stringId === 0 ? "" : runtimeColumnStore.normalizeStringId(stringId);
}

function materializeSliceValue(
  slice: RuntimeColumnSlice,
  offset: number,
  runtimeColumnStore: EngineRuntimeColumnStoreService,
): CellValue {
  const tag = decodeValueTag(slice.tags[offset]);
  switch (tag) {
    case ValueTag.Empty:
      return { tag: ValueTag.Empty };
    case ValueTag.Number:
      return { tag: ValueTag.Number, value: slice.numbers[offset] ?? 0 };
    case ValueTag.Boolean:
      return { tag: ValueTag.Boolean, value: (slice.numbers[offset] ?? 0) !== 0 };
    case ValueTag.String: {
      const stringId = slice.stringIds[offset] ?? 0;
      return {
        tag: ValueTag.String,
        value: stringId === 0 ? "" : runtimeColumnStore.normalizeStringId(stringId),
        stringId,
      };
    }
    case ValueTag.Error:
      return { tag: ValueTag.Error, code: slice.errors[offset] ?? ErrorCode.None };
  }
}

function buildSlicePredicate(compiled: CompiledCriteriaMatcher): SliceFastPredicate {
  const { operator, operand, wildcardPattern } = compiled;
  if (wildcardPattern) {
    return { kind: "generic", compiled };
  }
  if (operator === "=" || operator === "<>") {
    const negate = operator === "<>";
    switch (operand.tag) {
      case ValueTag.Empty:
        return { kind: "eq-empty", negate };
      case ValueTag.Boolean:
        return { kind: "eq-bool", negate, value: operand.value };
      case ValueTag.Number:
        return {
          kind: "eq-number",
          negate,
          value: Object.is(operand.value, -0) ? 0 : operand.value,
        };
      case ValueTag.String:
        return { kind: "eq-string", negate, value: operand.value.toUpperCase() };
      case ValueTag.Error:
        return { kind: "generic", compiled };
    }
  }
  if (operand.tag === ValueTag.Number) {
    return {
      kind: "cmp-number",
      operator,
      value: Object.is(operand.value, -0) ? 0 : operand.value,
    };
  }
  return { kind: "generic", compiled };
}

function slicePredicateMatches(
  predicate: SliceFastPredicate,
  slice: RuntimeColumnSlice,
  offset: number,
  runtimeColumnStore: EngineRuntimeColumnStoreService,
): boolean {
  switch (predicate.kind) {
    case "eq-empty": {
      const tag = decodeValueTag(slice.tags[offset]);
      const matches =
        tag === ValueTag.Empty ||
        (tag === ValueTag.String &&
          normalizeSliceString(runtimeColumnStore, slice.stringIds[offset] ?? 0) === "");
      return predicate.negate ? !matches : matches;
    }
    case "eq-bool": {
      const tag = decodeValueTag(slice.tags[offset]);
      const numeric =
        tag === ValueTag.Number || tag === ValueTag.Boolean || tag === ValueTag.Empty
          ? (slice.numbers[offset] ?? 0)
          : undefined;
      const matches =
        numeric !== undefined &&
        (Object.is(numeric, -0) ? 0 : numeric) === (predicate.value ? 1 : 0);
      return predicate.negate ? !matches : matches;
    }
    case "eq-number": {
      const tag = decodeValueTag(slice.tags[offset]);
      const numeric = Object.is(slice.numbers[offset] ?? 0, -0) ? 0 : (slice.numbers[offset] ?? 0);
      const matches =
        (tag === ValueTag.Number || tag === ValueTag.Boolean || tag === ValueTag.Empty) &&
        numeric === predicate.value;
      return predicate.negate ? !matches : matches;
    }
    case "eq-string": {
      const tag = decodeValueTag(slice.tags[offset]);
      const matches =
        (tag === ValueTag.String || tag === ValueTag.Empty) &&
        (tag === ValueTag.Empty
          ? ""
          : normalizeSliceString(runtimeColumnStore, slice.stringIds[offset] ?? 0)) ===
          predicate.value;
      return predicate.negate ? !matches : matches;
    }
    case "cmp-number": {
      const tag = decodeValueTag(slice.tags[offset]);
      if (tag !== ValueTag.Number && tag !== ValueTag.Boolean && tag !== ValueTag.Empty) {
        return false;
      }
      const numeric = Object.is(slice.numbers[offset] ?? 0, -0) ? 0 : (slice.numbers[offset] ?? 0);
      switch (predicate.operator) {
        case ">":
          return numeric > predicate.value;
        case ">=":
          return numeric >= predicate.value;
        case "<":
          return numeric < predicate.value;
        case "<=":
          return numeric <= predicate.value;
        default:
          return false;
      }
    }
    case "generic":
      return matchesCompiledCriteria(
        materializeSliceValue(slice, offset, runtimeColumnStore),
        predicate.compiled,
      );
  }
}

function requestKey(criteriaPairs: readonly CriterionRangePair[]): string {
  return criteriaPairs
    .map(
      ({ range, criteria }) =>
        `${range.sheetName}\t${range.col}\t${range.rowStart}\t${range.rowEnd}\t${criteriaCacheKey(criteria)}`,
    )
    .join("\u0001");
}

export function createCriterionRangeCacheService(args: {
  readonly runtimeColumnStore: EngineRuntimeColumnStoreService;
}): CriterionRangeCacheService {
  const cache = new Map<string, CriterionCacheEntry>();

  const getOrBuildMatchingRows = (request: {
    criteriaPairs: readonly CriterionRangePair[];
  }): CriterionRangeMatch | CellValue => {
    const { criteriaPairs } = request;
    if (criteriaPairs.length === 0) {
      return errorValue(ErrorCode.Value);
    }
    const expectedLength = criteriaPairs[0]!.range.length;
    if (criteriaPairs.some((pair) => pair.range.length !== expectedLength)) {
      return errorValue(ErrorCode.Value);
    }

    const slices = criteriaPairs.map((pair) =>
      args.runtimeColumnStore.getColumnSlice({
        sheetName: pair.range.sheetName,
        rowStart: pair.range.rowStart,
        rowEnd: pair.range.rowEnd,
        col: pair.range.col,
      }),
    );
    const key = requestKey(criteriaPairs);
    const existing = cache.get(key);
    if (
      existing &&
      existing.pairVersions.length === slices.length &&
      existing.pairVersions.every(
        (version, index) =>
          version.columnVersion === slices[index]!.columnVersion &&
          version.structureVersion === slices[index]!.structureVersion,
      )
    ) {
      return existing;
    }

    const predicates = criteriaPairs.map((pair) =>
      buildSlicePredicate(compileCriteriaMatcher(pair.criteria)),
    );
    const matchingRows: number[] = [];
    for (let rowOffset = 0; rowOffset < expectedLength; rowOffset += 1) {
      let matches = true;
      for (let pairIndex = 0; pairIndex < predicates.length; pairIndex += 1) {
        if (
          !slicePredicateMatches(
            predicates[pairIndex]!,
            slices[pairIndex]!,
            rowOffset,
            args.runtimeColumnStore,
          )
        ) {
          matches = false;
          break;
        }
      }
      if (matches) {
        matchingRows.push(rowOffset);
      }
    }

    const entry: CriterionCacheEntry = {
      rows: Uint32Array.from(matchingRows),
      length: matchingRows.length,
      pairVersions: slices.map((slice) => ({
        columnVersion: slice.columnVersion,
        structureVersion: slice.structureVersion,
      })),
    };
    cache.set(key, entry);
    return entry;
  };

  return {
    getOrBuildMatchingRows,
  };
}
