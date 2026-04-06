import type { SpreadsheetEngine } from "@bilig/core";
import type { WorkbookLocalAuthoritativeBase } from "@bilig/storage-browser";
import { parseCellAddress } from "@bilig/formula";
import { collectMaterializedSheetAddresses } from "./worker-local-materialization.js";

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
