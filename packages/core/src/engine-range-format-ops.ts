import {
  createCellNumberFormatRecord,
  type CellNumberFormatInput,
  type CellNumberFormatRecord,
  type CellRangeRef,
  type CellStyleField,
  type CellStylePatch,
  type CellStyleRecord,
} from "@bilig/protocol";
import { formatAddress } from "@bilig/formula";
import type { EngineOp } from "@bilig/workbook-domain";
import { intersectRangeBounds, normalizeRange } from "./engine-range-utils.js";
import {
  applyStylePatch,
  clearStyleFields,
  cloneCellStyleRecord,
  normalizeCellStylePatch,
} from "./engine-style-utils.js";
import { WorkbookStore } from "./workbook-store.js";

interface StyleTile {
  range: CellRangeRef;
  styleId: string;
}

interface FormatTile {
  range: CellRangeRef;
  formatId: string;
}

export function buildStylePatchOps(
  workbook: WorkbookStore,
  range: CellRangeRef,
  patch: CellStylePatch,
): EngineOp[] {
  const normalizedPatch = normalizeCellStylePatch(patch);
  if (
    !normalizedPatch.fill &&
    !normalizedPatch.font &&
    !normalizedPatch.alignment &&
    !normalizedPatch.borders
  ) {
    return [];
  }
  return materializeStyleRangeOps(workbook, range, (baseStyle) =>
    workbook.internCellStyle(applyStylePatch(baseStyle, normalizedPatch)),
  );
}

export function buildStyleClearOps(
  workbook: WorkbookStore,
  range: CellRangeRef,
  fields?: readonly CellStyleField[],
): EngineOp[] {
  return materializeStyleRangeOps(workbook, range, (baseStyle) =>
    workbook.internCellStyle(clearStyleFields(baseStyle, fields)),
  );
}

export function restoreStyleRangeOps(workbook: WorkbookStore, range: CellRangeRef): EngineOp[] {
  return materializeStyleRangeOps(workbook, range, (baseStyle, currentStyleId) => ({
    id: currentStyleId,
    ...baseStyle,
  }));
}

export function buildFormatPatchOps(
  workbook: WorkbookStore,
  range: CellRangeRef,
  format: CellNumberFormatInput,
): EngineOp[] {
  const normalized = workbook.internCellNumberFormat(
    typeof format === "string"
      ? format
      : createCellNumberFormatRecord(WorkbookStore.defaultFormatId, format),
  );
  return materializeFormatRangeOps(workbook, range, () => normalized.id, normalized);
}

export function buildFormatClearOps(workbook: WorkbookStore, range: CellRangeRef): EngineOp[] {
  return materializeFormatRangeOps(workbook, range, () => WorkbookStore.defaultFormatId);
}

export function restoreFormatRangeOps(workbook: WorkbookStore, range: CellRangeRef): EngineOp[] {
  return materializeFormatRangeOps(workbook, range, (_currentFormatId, tile) => tile.formatId);
}

function materializeStyleRangeOps(
  workbook: WorkbookStore,
  range: CellRangeRef,
  resolveStyle: (baseStyle: Omit<CellStyleRecord, "id">, currentStyleId: string) => CellStyleRecord,
): EngineOp[] {
  const tiles = resolveStyleTiles(workbook, range);
  const ops: EngineOp[] = [];
  tiles.forEach((tile) => {
    const current = workbook.getCellStyle(tile.styleId) ?? {
      id: WorkbookStore.defaultStyleId,
    };
    const next = resolveStyle(current, tile.styleId);
    const normalizedId = next.id || WorkbookStore.defaultStyleId;
    if (normalizedId === tile.styleId) {
      return;
    }
    if (normalizedId !== WorkbookStore.defaultStyleId) {
      ops.push({
        kind: "upsertCellStyle",
        style: cloneCellStyleRecord(next),
      });
    }
    ops.push({
      kind: "setStyleRange",
      range: tile.range,
      styleId: normalizedId,
    });
  });
  return ops;
}

function materializeFormatRangeOps(
  workbook: WorkbookStore,
  range: CellRangeRef,
  resolveFormatId: (currentFormatId: string, tile: FormatTile) => string,
  upsertFormat?: CellNumberFormatRecord,
): EngineOp[] {
  const tiles = resolveFormatTiles(workbook, range);
  const ops: EngineOp[] = [];
  if (upsertFormat && upsertFormat.id !== WorkbookStore.defaultFormatId) {
    ops.push({ kind: "upsertCellNumberFormat", format: { ...upsertFormat } });
  }
  tiles.forEach((tile) => {
    const nextFormatId = resolveFormatId(tile.formatId, tile);
    if (nextFormatId === tile.formatId) {
      return;
    }
    ops.push({
      kind: "setFormatRange",
      range: tile.range,
      formatId: nextFormatId,
    });
  });
  return ops;
}

function resolveStyleTiles(workbook: WorkbookStore, range: CellRangeRef): StyleTile[] {
  const bounds = normalizeRange(range);
  const sheetRanges = workbook.listStyleRanges(range.sheetName);
  const rowBoundaries = new Set<number>([bounds.startRow, bounds.endRow + 1]);
  const colBoundaries = new Set<number>([bounds.startCol, bounds.endCol + 1]);

  sheetRanges.forEach((record) => {
    const clipped = intersectRangeBounds(record.range, bounds);
    if (!clipped) {
      return;
    }
    rowBoundaries.add(clipped.startRow);
    rowBoundaries.add(clipped.endRow + 1);
    colBoundaries.add(clipped.startCol);
    colBoundaries.add(clipped.endCol + 1);
  });

  const rows = [...rowBoundaries].toSorted((left, right) => left - right);
  const cols = [...colBoundaries].toSorted((left, right) => left - right);
  const tiles: StyleTile[] = [];

  for (let rowIndex = 0; rowIndex < rows.length - 1; rowIndex += 1) {
    const startRow = rows[rowIndex]!;
    const endRow = rows[rowIndex + 1]! - 1;
    for (let colIndex = 0; colIndex < cols.length - 1; colIndex += 1) {
      const startCol = cols[colIndex]!;
      const endCol = cols[colIndex + 1]! - 1;
      tiles.push({
        range: {
          sheetName: range.sheetName,
          startAddress: formatAddress(startRow, startCol),
          endAddress: formatAddress(endRow, endCol),
        },
        styleId: workbook.getStyleId(range.sheetName, startRow, startCol),
      });
    }
  }

  return tiles;
}

function resolveFormatTiles(workbook: WorkbookStore, range: CellRangeRef): FormatTile[] {
  const bounds = normalizeRange(range);
  const sheetRanges = workbook.listFormatRanges(range.sheetName);
  const rowBoundaries = new Set<number>([bounds.startRow, bounds.endRow + 1]);
  const colBoundaries = new Set<number>([bounds.startCol, bounds.endCol + 1]);

  sheetRanges.forEach((record) => {
    const clipped = intersectRangeBounds(record.range, bounds);
    if (!clipped) {
      return;
    }
    rowBoundaries.add(clipped.startRow);
    rowBoundaries.add(clipped.endRow + 1);
    colBoundaries.add(clipped.startCol);
    colBoundaries.add(clipped.endCol + 1);
  });

  const rows = [...rowBoundaries].toSorted((left, right) => left - right);
  const cols = [...colBoundaries].toSorted((left, right) => left - right);
  const tiles: FormatTile[] = [];

  for (let rowIndex = 0; rowIndex < rows.length - 1; rowIndex += 1) {
    const startRow = rows[rowIndex]!;
    const endRow = rows[rowIndex + 1]! - 1;
    for (let colIndex = 0; colIndex < cols.length - 1; colIndex += 1) {
      const startCol = cols[colIndex]!;
      const endCol = cols[colIndex + 1]! - 1;
      tiles.push({
        range: {
          sheetName: range.sheetName,
          startAddress: formatAddress(startRow, startCol),
          endAddress: formatAddress(endRow, endCol),
        },
        formatId: workbook.getRangeFormatId(range.sheetName, startRow, startCol),
      });
    }
  }

  return tiles;
}
