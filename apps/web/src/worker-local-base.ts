import type { SpreadsheetEngine } from "@bilig/core";
import type { WorkbookLocalAuthoritativeBase } from "@bilig/storage-browser";
import { formatAddress, parseCellAddress } from "@bilig/formula";
import type { CellRangeRef } from "@bilig/protocol";

function collectRangeAddresses(range: CellRangeRef, addresses: Set<string>): void {
  const start = parseCellAddress(range.startAddress, range.sheetName);
  const end = parseCellAddress(range.endAddress, range.sheetName);
  const rowStart = Math.min(start.row, end.row);
  const rowEnd = Math.max(start.row, end.row);
  const colStart = Math.min(start.col, end.col);
  const colEnd = Math.max(start.col, end.col);
  for (let row = rowStart; row <= rowEnd; row += 1) {
    for (let col = colStart; col <= colEnd; col += 1) {
      addresses.add(formatAddress(row, col));
    }
  }
}

function collectMaterializedSheetAddresses(
  engine: SpreadsheetEngine,
  sheetName: string,
): readonly string[] {
  const addresses = new Set<string>();
  const sheet = engine.workbook.getSheet(sheetName);
  sheet?.grid.forEachCellEntry((_cellIndex, row, col) => {
    addresses.add(formatAddress(row, col));
  });
  engine.workbook.listStyleRanges(sheetName).forEach((entry) => {
    collectRangeAddresses(entry.range, addresses);
  });
  engine.workbook.listFormatRanges(sheetName).forEach((entry) => {
    collectRangeAddresses(entry.range, addresses);
  });
  return [...addresses].toSorted((left, right) => {
    const leftParsed = parseCellAddress(left, sheetName);
    const rightParsed = parseCellAddress(right, sheetName);
    return leftParsed.row - rightParsed.row || leftParsed.col - rightParsed.col;
  });
}

export function buildWorkbookLocalAuthoritativeBase(
  engine: SpreadsheetEngine,
): WorkbookLocalAuthoritativeBase {
  const sheets = [...engine.workbook.sheetsByName.values()]
    .toSorted((left, right) => left.order - right.order)
    .map((sheet) => {
      const freezePane = engine.getFreezePane(sheet.name);
      return {
        name: sheet.name,
        sortOrder: sheet.order,
        freezeRows: freezePane?.rows ?? 0,
        freezeCols: freezePane?.cols ?? 0,
      };
    });

  const cellInputs: Array<WorkbookLocalAuthoritativeBase["cellInputs"][number]> = [];
  const cellRenders: Array<WorkbookLocalAuthoritativeBase["cellRenders"][number]> = [];
  const rowAxisEntries: Array<WorkbookLocalAuthoritativeBase["rowAxisEntries"][number]> = [];
  const columnAxisEntries: Array<WorkbookLocalAuthoritativeBase["columnAxisEntries"][number]> = [];

  for (const sheet of sheets) {
    for (const address of collectMaterializedSheetAddresses(engine, sheet.name)) {
      const parsed = parseCellAddress(address, sheet.name);
      const snapshot = engine.getCell(sheet.name, address);
      cellRenders.push({
        sheetName: sheet.name,
        address,
        rowNum: parsed.row,
        colNum: parsed.col,
        value: snapshot.value,
        flags: snapshot.flags,
        version: snapshot.version,
        styleId: snapshot.styleId,
        numberFormatId: snapshot.numberFormatId,
      });
      if (
        snapshot.input !== undefined ||
        snapshot.formula !== undefined ||
        snapshot.format !== undefined
      ) {
        cellInputs.push({
          sheetName: sheet.name,
          address,
          rowNum: parsed.row,
          colNum: parsed.col,
          input: snapshot.input,
          formula: snapshot.formula,
          format: snapshot.format,
        });
      }
    }

    engine.getRowAxisEntries(sheet.name).forEach((entry) => {
      rowAxisEntries.push({ sheetName: sheet.name, entry });
    });
    engine.getColumnAxisEntries(sheet.name).forEach((entry) => {
      columnAxisEntries.push({ sheetName: sheet.name, entry });
    });
  }

  return {
    sheets,
    cellInputs,
    cellRenders,
    rowAxisEntries,
    columnAxisEntries,
    styles: engine.workbook.listCellStyles(),
  };
}
