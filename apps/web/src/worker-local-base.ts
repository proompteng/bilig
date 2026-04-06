import type { SpreadsheetEngine } from "@bilig/core";
import type { WorkbookLocalAuthoritativeBase } from "@bilig/storage-browser";
import { parseCellAddress } from "@bilig/formula";
import { collectMaterializedSheetAddresses } from "./worker-local-materialization.js";

function buildWorkbookLocalSheetRecords(
  engine: SpreadsheetEngine,
  sheetNames: readonly string[],
): WorkbookLocalAuthoritativeBase["sheets"] {
  return sheetNames
    .flatMap((sheetName) => {
      const sheet = engine.workbook.getSheet(sheetName);
      if (!sheet) {
        return [];
      }
      const freezePane = engine.getFreezePane(sheet.name);
      return [
        {
          sheetId: sheet.id,
          name: sheet.name,
          sortOrder: sheet.order,
          freezeRows: freezePane?.rows ?? 0,
          freezeCols: freezePane?.cols ?? 0,
        },
      ];
    })
    .toSorted((left, right) => left.sortOrder - right.sortOrder);
}

export function buildWorkbookLocalAuthoritativeBaseForSheets(
  engine: SpreadsheetEngine,
  sheetNames: readonly string[],
): WorkbookLocalAuthoritativeBase {
  const sheets = buildWorkbookLocalSheetRecords(engine, sheetNames);
  const cellInputs: Array<WorkbookLocalAuthoritativeBase["cellInputs"][number]> = [];
  const cellRenders: Array<WorkbookLocalAuthoritativeBase["cellRenders"][number]> = [];
  const rowAxisEntries: Array<WorkbookLocalAuthoritativeBase["rowAxisEntries"][number]> = [];
  const columnAxisEntries: Array<WorkbookLocalAuthoritativeBase["columnAxisEntries"][number]> = [];

  for (const sheet of sheets) {
    for (const address of collectMaterializedSheetAddresses(engine, sheet.name)) {
      const parsed = parseCellAddress(address, sheet.name);
      const snapshot = engine.getCell(sheet.name, address);
      cellRenders.push({
        sheetId: sheet.sheetId,
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
          sheetId: sheet.sheetId,
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
      rowAxisEntries.push({ sheetId: sheet.sheetId, sheetName: sheet.name, entry });
    });
    engine.getColumnAxisEntries(sheet.name).forEach((entry) => {
      columnAxisEntries.push({ sheetId: sheet.sheetId, sheetName: sheet.name, entry });
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

export function buildWorkbookLocalAuthoritativeBase(
  engine: SpreadsheetEngine,
): WorkbookLocalAuthoritativeBase {
  return buildWorkbookLocalAuthoritativeBaseForSheets(
    engine,
    [...engine.workbook.sheetsByName.values()].map((sheet) => sheet.name),
  );
}
