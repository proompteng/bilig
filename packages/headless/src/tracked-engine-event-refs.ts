import { makeCellKey, type SpreadsheetEngine } from "@bilig/core";
import { formatAddress } from "@bilig/formula";
import type { EngineEvent } from "@bilig/protocol";

export interface TrackedCellRef {
  sheetId: number;
  sheetName: string;
  address: string;
  row: number;
  col: number;
}

export function collectTrackedCellRefsFromEvents(
  engine: SpreadsheetEngine,
  events: readonly EngineEvent[],
): TrackedCellRef[] | null {
  if (events.length === 0) {
    return [];
  }

  const refs = new Map<number, TrackedCellRef>();
  for (const event of events) {
    if (
      event.invalidation === "full" ||
      event.invalidatedRanges.length > 0 ||
      event.invalidatedRows.length > 0 ||
      event.invalidatedColumns.length > 0
    ) {
      return null;
    }
    for (let index = 0; index < event.changedCellIndices.length; index += 1) {
      const cellIndex = event.changedCellIndices[index]!;
      const sheetId = engine.workbook.cellStore.sheetIds[cellIndex];
      const row = engine.workbook.cellStore.rows[cellIndex];
      const col = engine.workbook.cellStore.cols[cellIndex];
      if (sheetId === undefined || row === undefined || col === undefined) {
        return null;
      }
      const sheet = engine.workbook.getSheetById(sheetId);
      if (!sheet) {
        return null;
      }
      refs.set(makeCellKey(sheetId, row, col), {
        sheetId,
        sheetName: sheet.name,
        address: formatAddress(row, col),
        row,
        col,
      });
    }
  }
  return [...refs.values()];
}
