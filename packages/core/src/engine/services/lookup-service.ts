import { ValueTag, type CellValue } from "@bilig/protocol";
import { parseRangeAddress } from "@bilig/formula";
import type { EngineRuntimeState } from "../runtime-state.js";

interface ExactVectorMatchRequest {
  lookupValue: CellValue;
  sheetName: string;
  start: string;
  end: string;
  searchMode: 1 | -1;
}

export type ExactVectorMatchResult =
  | { handled: false }
  | { handled: true; position: number | undefined };

export interface EngineLookupService {
  readonly findExactVectorMatch: (request: ExactVectorMatchRequest) => ExactVectorMatchResult;
}

interface ExactColumnIndexEntry {
  columnVersion: number;
  firstPositions: Map<string, number>;
  lastPositions: Map<string, number>;
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

export function createEngineLookupService(args: {
  readonly state: Pick<EngineRuntimeState, "workbook" | "strings">;
}): EngineLookupService {
  const exactColumnIndices = new Map<string, ExactColumnIndexEntry>();
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

  return {
    findExactVectorMatch(request) {
      const normalizedLookupKey = normalizeExactLookupKey(request.lookupValue, (id) =>
        args.state.strings.get(id),
      );
      if (normalizedLookupKey === undefined) {
        return { handled: false };
      }

      const parsedRange = parseRangeAddress(`${request.start}:${request.end}`, request.sheetName);
      if (parsedRange.kind !== "cells" || parsedRange.start.col !== parsedRange.end.col) {
        return { handled: false };
      }

      const cacheKey = `${request.sheetName}\t${parsedRange.start.col}\t${parsedRange.start.row}\t${parsedRange.end.row}`;
      const columnVersion = args.state.workbook.getSheetColumnVersion(
        request.sheetName,
        parsedRange.start.col,
      );
      let entry = exactColumnIndices.get(cacheKey);
      if (!entry || entry.columnVersion !== columnVersion) {
        entry = buildExactColumnIndex(
          request.sheetName,
          parsedRange.start.col,
          parsedRange.start.row,
          parsedRange.end.row,
        );
        exactColumnIndices.set(cacheKey, entry);
      }

      const row =
        request.searchMode === -1
          ? entry.lastPositions.get(normalizedLookupKey)
          : entry.firstPositions.get(normalizedLookupKey);
      return {
        handled: true,
        position: row === undefined ? undefined : row - parsedRange.start.row + 1,
      };
    },
  };
}
