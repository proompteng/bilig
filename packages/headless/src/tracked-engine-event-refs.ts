import { makeCellKey, type SpreadsheetEngine } from "@bilig/core";
import { formatAddress } from "@bilig/formula";
import type { CellValue, EngineEvent } from "@bilig/protocol";

export interface TrackedCellRef {
  cellIndex: number;
  sheetId: number;
  sheetName: string;
  address: string;
  row: number;
  col: number;
  value: CellValue;
}

export function collectTrackedCellRefsFromEvents(
  engine: SpreadsheetEngine,
  events: readonly EngineEvent[],
): TrackedCellRef[] | null {
  if (events.length === 0) {
    return [];
  }

  const toTrackedCellRef = (cellIndex: number): TrackedCellRef | null => {
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
    return {
      cellIndex,
      sheetId,
      sheetName: sheet.name,
      address: formatAddress(row, col),
      row,
      col,
      value: engine.workbook.cellStore.getValue(cellIndex, (id) => engine.strings.get(id)),
    };
  };

  if (events.length === 1) {
    const event = events[0]!;
    if (
      event.invalidation === "full" ||
      event.invalidatedRanges.length > 0 ||
      event.invalidatedRows.length > 0 ||
      event.invalidatedColumns.length > 0
    ) {
      return null;
    }

    const refs: TrackedCellRef[] = [];
    const seen = new Set<number>();
    for (let index = 0; index < event.changedCellIndices.length; index += 1) {
      const ref = toTrackedCellRef(event.changedCellIndices[index]!);
      if (!ref) {
        return null;
      }
      const key = makeCellKey(ref.sheetId, ref.row, ref.col);
      if (seen.has(key)) {
        continue;
      }
      seen.add(key);
      refs.push(ref);
    }
    return refs;
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
      const ref = toTrackedCellRef(event.changedCellIndices[index]!);
      if (!ref) {
        return null;
      }
      refs.set(makeCellKey(ref.sheetId, ref.row, ref.col), ref);
    }
  }
  return [...refs.values()];
}
