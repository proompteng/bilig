import type { CellSnapshot, RecalcMetrics, SelectionState, Viewport } from '@bilig/protocol'
import type { SpreadsheetEngine } from './engine.js'

export function selectCellSnapshot(engine: SpreadsheetEngine, sheetName: string, address: string): CellSnapshot {
  return engine.getCell(sheetName, address)
}

export function selectMetrics(engine: SpreadsheetEngine): RecalcMetrics {
  return engine.getLastMetrics()
}

export function selectSelectionState(engine: SpreadsheetEngine): SelectionState {
  return engine.getSelectionState()
}

export function selectViewportCells(engine: SpreadsheetEngine, sheetName: string, viewport: Viewport): CellSnapshot[] {
  const result: CellSnapshot[] = []
  const sheet = engine.workbook.getSheet(sheetName)
  if (!sheet) return result
  sheet.grid.forEachInRange(viewport.rowStart, viewport.colStart, viewport.rowEnd, viewport.colEnd, (cellIndex) => {
    result.push(engine.getCellByIndex(cellIndex))
  })
  return result
}
