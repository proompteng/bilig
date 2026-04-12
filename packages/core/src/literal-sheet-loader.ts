import { ErrorCode, ValueTag, type LiteralInput } from "@bilig/protocol";
import type { StringPool } from "./string-pool.js";
import { WorkbookStore, makeCellKey } from "./workbook-store.js";
import { CellFlags } from "./cell-store.js";

export function loadLiteralSheetIntoEmptySheet(
  workbook: WorkbookStore,
  strings: StringPool,
  sheetId: number,
  content: readonly (readonly LiteralInput[])[],
  shouldMaterialize: (raw: LiteralInput, rowIndex: number, colIndex: number) => boolean = (raw) =>
    raw !== null,
): number {
  const sheet = workbook.getSheetById(sheetId);
  if (!sheet) {
    throw new Error(`Unknown sheet id: ${sheetId}`);
  }

  const literalCount = countMaterializedLiteralCells(content, shouldMaterialize);
  if (literalCount === 0) {
    return 0;
  }

  const cellStore = workbook.cellStore;
  cellStore.ensureCapacity(cellStore.size + literalCount);
  const previousOnSetValue = cellStore.onSetValue;
  cellStore.onSetValue = null;
  try {
    content.forEach((row, rowIndex) => {
      row.forEach((raw, colIndex) => {
        if (!shouldMaterialize(raw, rowIndex, colIndex)) {
          return;
        }
        if (raw === null) {
          return;
        }

        const cellIndex = cellStore.allocate(sheetId, rowIndex, colIndex);
        workbook.cellKeyToIndex.set(makeCellKey(sheetId, rowIndex, colIndex), cellIndex);
        sheet.grid.set(rowIndex, colIndex, cellIndex);
        writeLiteralCell(cellStore, strings, cellIndex, raw);
        sheet.columnVersions[colIndex] = (sheet.columnVersions[colIndex] ?? 0) + 1;
      });
    });
  } finally {
    cellStore.onSetValue = previousOnSetValue;
  }

  return literalCount;
}

function countMaterializedLiteralCells(
  content: readonly (readonly LiteralInput[])[],
  shouldMaterialize: (raw: LiteralInput, rowIndex: number, colIndex: number) => boolean,
): number {
  let count = 0;
  content.forEach((row, rowIndex) => {
    row.forEach((value, colIndex) => {
      if (shouldMaterialize(value, rowIndex, colIndex)) {
        count += 1;
      }
    });
  });
  return count;
}

function writeLiteralCell(
  cellStore: WorkbookStore["cellStore"],
  strings: StringPool,
  cellIndex: number,
  raw: Exclude<LiteralInput, null>,
): void {
  cellStore.flags[cellIndex] = CellFlags.Materialized;
  cellStore.formulaIds[cellIndex] = 0;
  cellStore.errors[cellIndex] = ErrorCode.None;
  cellStore.versions[cellIndex] = 1;
  cellStore.topoRanks[cellIndex] = 0;
  cellStore.cycleGroupIds[cellIndex] = -1;

  if (typeof raw === "number") {
    cellStore.tags[cellIndex] = ValueTag.Number;
    cellStore.numbers[cellIndex] = raw;
    cellStore.stringIds[cellIndex] = 0;
    return;
  }

  if (typeof raw === "boolean") {
    cellStore.tags[cellIndex] = ValueTag.Boolean;
    cellStore.numbers[cellIndex] = raw ? 1 : 0;
    cellStore.stringIds[cellIndex] = 0;
    return;
  }

  cellStore.tags[cellIndex] = ValueTag.String;
  cellStore.numbers[cellIndex] = 0;
  cellStore.stringIds[cellIndex] = strings.intern(raw);
}
