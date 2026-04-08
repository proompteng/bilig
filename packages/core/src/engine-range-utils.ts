import type { CellRangeRef } from "@bilig/protocol";
import { parseCellAddress } from "@bilig/formula";

export function intersectRangeBounds(
  range: CellRangeRef,
  bounds: { startRow: number; endRow: number; startCol: number; endCol: number },
): { startRow: number; endRow: number; startCol: number; endCol: number } | undefined {
  const normalized = normalizeRange(range);
  const startRow = Math.max(bounds.startRow, normalized.startRow);
  const endRow = Math.min(bounds.endRow, normalized.endRow);
  const startCol = Math.max(bounds.startCol, normalized.startCol);
  const endCol = Math.min(bounds.endCol, normalized.endCol);
  if (startRow > endRow || startCol > endCol) {
    return undefined;
  }
  return { startRow, endRow, startCol, endCol };
}

export function normalizeRange(range: CellRangeRef): {
  startRow: number;
  endRow: number;
  startCol: number;
  endCol: number;
} {
  const start = parseCellAddress(range.startAddress, range.sheetName);
  const end = parseCellAddress(range.endAddress, range.sheetName);
  return {
    startRow: Math.min(start.row, end.row),
    endRow: Math.max(start.row, end.row),
    startCol: Math.min(start.col, end.col),
    endCol: Math.max(start.col, end.col),
  };
}
