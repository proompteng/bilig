import { formatAddress, parseCellAddress } from '@bilig/formula'
import type { WorkbookMergeRangeSnapshot } from '@bilig/protocol'
import type { GridEngineLike } from './grid-engine.js'
import type { Rectangle } from './gridTypes.js'

export interface ResolvedMergedCell {
  readonly range: WorkbookMergeRangeSnapshot
  readonly startRow: number
  readonly endRow: number
  readonly startCol: number
  readonly endCol: number
  readonly isAnchor: boolean
}

export function resolveMergedCell(engine: GridEngineLike, sheetName: string, row: number, col: number): ResolvedMergedCell | undefined {
  const range = engine.getMergeRange?.(sheetName, formatAddress(row, col))
  if (!range) {
    return undefined
  }
  const start = parseCellAddress(range.startAddress, range.sheetName)
  const end = parseCellAddress(range.endAddress, range.sheetName)
  const startRow = Math.min(start.row, end.row)
  const endRow = Math.max(start.row, end.row)
  const startCol = Math.min(start.col, end.col)
  const endCol = Math.max(start.col, end.col)
  return {
    range: {
      sheetName: range.sheetName,
      startAddress: formatAddress(startRow, startCol),
      endAddress: formatAddress(endRow, endCol),
    },
    startRow,
    endRow,
    startCol,
    endCol,
    isAnchor: row === startRow && col === startCol,
  }
}

export function resolveMergedCellBounds(input: {
  readonly merged: ResolvedMergedCell
  readonly fallback: Rectangle
  readonly getCellBounds: (col: number, row: number) => Rectangle | undefined
}): Rectangle {
  const { merged, fallback, getCellBounds } = input
  const startBounds = getCellBounds(merged.startCol, merged.startRow)
  const endBounds = getCellBounds(merged.endCol, merged.endRow)
  if (!startBounds || !endBounds) {
    return fallback
  }
  return {
    x: startBounds.x,
    y: startBounds.y,
    width: endBounds.x + endBounds.width - startBounds.x,
    height: endBounds.y + endBounds.height - startBounds.y,
  }
}
