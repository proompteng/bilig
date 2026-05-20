import { ImportedWorkbookArena, ImportedWorksheetStyleIndexArena, type ImportedWorksheetCellScan } from './xlsx-large-simple-arena.js'
import type { HeadlessLargeSimpleWorksheetScan } from './xlsx-large-simple-headless-worksheet-scanner.js'

export function importedWorksheetCellScanFromHeadless(scan: HeadlessLargeSimpleWorksheetScan): ImportedWorksheetCellScan {
  return {
    arena: new ImportedWorkbookArena(),
    sheetIndex: scan.sheetIndex,
    richTextCells: [],
    styleIndexes: new ImportedWorksheetStyleIndexArena(),
    blankStyleCellCount: 0,
    cellCount: scan.cellCount,
    valueCellCount: scan.valueCellCount,
    formulaCellCount: scan.formulaCellCount,
    mergeCount: scan.mergeCount,
    conditionalFormatCount: scan.conditionalFormatCount,
    dataValidationCount: scan.dataValidationCount,
    tableCount: scan.tableCount,
    rowCount: scan.rowCount,
    columnCount: scan.columnCount,
    usedRange: scan.usedRange,
  }
}
