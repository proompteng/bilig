import type { CellStyleRecord, SheetStyleRangeSnapshot } from '@bilig/protocol'
import { internImportedStyle } from './xlsx-import-cell-styles.js'
import type { ImportedWorksheetCellScan } from './xlsx-large-simple-arena.js'

export function buildLargeSimpleStyleRanges(
  sheetName: string,
  cellScan: ImportedWorksheetCellScan,
  stylesByIndex: ReadonlyMap<number, Omit<CellStyleRecord, 'id'>>,
  styleCatalog: Map<string, CellStyleRecord>,
): SheetStyleRangeSnapshot[] {
  return cellScan.styleIndexes.flatMap((entry) => {
    const style = stylesByIndex.get(entry.styleIndex)
    if (!style) {
      return []
    }
    const address = encodeCellAddress(entry.row, entry.column)
    return [
      {
        range: { sheetName, startAddress: address, endAddress: address },
        styleId: internImportedStyle(style, styleCatalog),
      },
    ]
  })
}

function encodeCellAddress(row: number, column: number): string {
  let value = column + 1
  let columnName = ''
  while (value > 0) {
    value -= 1
    columnName = String.fromCharCode(65 + (value % 26)) + columnName
    value = Math.floor(value / 26)
  }
  return `${columnName}${String(row + 1)}`
}
