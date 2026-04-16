import type { SpreadsheetEngine } from '@bilig/core'
import { formatAddress, parseCellAddress } from '@bilig/formula'
import type { CellRangeRef } from '@bilig/protocol'

function collectRangeAddresses(range: CellRangeRef, addresses: Set<string>): void {
  const start = parseCellAddress(range.startAddress, range.sheetName)
  const end = parseCellAddress(range.endAddress, range.sheetName)
  const rowStart = Math.min(start.row, end.row)
  const rowEnd = Math.max(start.row, end.row)
  const colStart = Math.min(start.col, end.col)
  const colEnd = Math.max(start.col, end.col)
  for (let row = rowStart; row <= rowEnd; row += 1) {
    for (let col = colStart; col <= colEnd; col += 1) {
      addresses.add(formatAddress(row, col))
    }
  }
}

export function collectMaterializedSheetAddresses(engine: SpreadsheetEngine, sheetName: string): readonly string[] {
  const addresses = new Set<string>()
  const sheet = engine.workbook.getSheet(sheetName)
  sheet?.grid.forEachCellEntry((_cellIndex, row, col) => {
    addresses.add(formatAddress(row, col))
  })
  engine.workbook.listStyleRanges(sheetName).forEach((entry) => {
    collectRangeAddresses(entry.range, addresses)
  })
  engine.workbook.listFormatRanges(sheetName).forEach((entry) => {
    collectRangeAddresses(entry.range, addresses)
  })
  return [...addresses].toSorted((left, right) => {
    const leftParsed = parseCellAddress(left, sheetName)
    const rightParsed = parseCellAddress(right, sheetName)
    return leftParsed.row - rightParsed.row || leftParsed.col - rightParsed.col
  })
}
