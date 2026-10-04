import type { WorkbookStore } from '../../workbook-store.js'
import type { DirtyRegion } from './recalc-service-types.js'

export function forEachDirtyRegionCell(workbook: WorkbookStore, region: DirtyRegion, listener: (cellIndex: number) => void): void {
  const sheet = workbook.getSheet(region.sheetName)
  if (!sheet) return
  const area = (region.rowEnd - region.rowStart + 1) * (region.colEnd - region.colStart + 1)
  if (area <= workbook.cellStore.size) {
    sheet.grid.forEachInRange(region.rowStart, region.colStart, region.rowEnd, region.colEnd, listener)
    return
  }
  sheet.grid.forEachCellEntry((cellIndex, row, col) => {
    if (row < region.rowStart || row > region.rowEnd || col < region.colStart || col > region.colEnd) return
    listener(cellIndex)
  })
}
