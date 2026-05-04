import { formatAddress, parseCellAddress } from '@bilig/formula'
import type { CellRangeRef } from '@bilig/protocol'

export interface NormalizedWorkbookAgentRange extends CellRangeRef {
  readonly startRow: number
  readonly endRow: number
  readonly startCol: number
  readonly endCol: number
}

export interface WorkbookAgentRangeChunk {
  readonly sheetName: string
  readonly startAddress: string
  readonly endAddress: string
  readonly rowStart: number
  readonly rowEnd: number
  readonly colStart: number
  readonly colEnd: number
  readonly rowCount: number
  readonly columnCount: number
  readonly cellCount: number
  readonly index: number
}

export interface WorkbookAgentRangeChunkPlan {
  readonly requestedRange: CellRangeRef
  readonly totalCells: number
  readonly maxCellsPerChunk: number
  readonly chunkCount: number
  readonly chunks: readonly WorkbookAgentRangeChunk[]
}

export function normalizeWorkbookAgentRange(range: CellRangeRef): NormalizedWorkbookAgentRange {
  const start = parseCellAddress(range.startAddress, range.sheetName)
  const end = parseCellAddress(range.endAddress, range.sheetName)
  const startRow = Math.min(start.row, end.row)
  const endRow = Math.max(start.row, end.row)
  const startCol = Math.min(start.col, end.col)
  const endCol = Math.max(start.col, end.col)
  return {
    sheetName: range.sheetName,
    startAddress: formatAddress(startRow, startCol),
    endAddress: formatAddress(endRow, endCol),
    startRow,
    endRow,
    startCol,
    endCol,
  }
}

export function countWorkbookAgentRangeCells(range: CellRangeRef): number {
  const normalized = normalizeWorkbookAgentRange(range)
  return (normalized.endRow - normalized.startRow + 1) * (normalized.endCol - normalized.startCol + 1)
}

export function countWorkbookAgentRangesCells(ranges: readonly CellRangeRef[]): number {
  return ranges.reduce((sum, range) => sum + countWorkbookAgentRangeCells(range), 0)
}

export function countWorkbookAgentRangeRows(range: CellRangeRef): number {
  const normalized = normalizeWorkbookAgentRange(range)
  return normalized.endRow - normalized.startRow + 1
}

export function countWorkbookAgentRangeColumns(range: CellRangeRef): number {
  const normalized = normalizeWorkbookAgentRange(range)
  return normalized.endCol - normalized.startCol + 1
}

export function enumerateWorkbookAgentRangeAddresses(range: CellRangeRef, limit = 64): string[] {
  const normalized = normalizeWorkbookAgentRange(range)
  const addresses: string[] = []
  for (let row = normalized.startRow; row <= normalized.endRow && addresses.length < limit; row += 1) {
    for (let col = normalized.startCol; col <= normalized.endCol && addresses.length < limit; col += 1) {
      addresses.push(formatAddress(row, col))
    }
  }
  return addresses
}

export function toWorkbookAgentRangeRef(range: CellRangeRef): CellRangeRef {
  const normalized = normalizeWorkbookAgentRange(range)
  return {
    sheetName: normalized.sheetName,
    startAddress: normalized.startAddress,
    endAddress: normalized.endAddress,
  }
}

export function workbookAgentRangesIntersect(left: CellRangeRef, right: CellRangeRef): boolean {
  const leftBounds = normalizeWorkbookAgentRange(left)
  const rightBounds = normalizeWorkbookAgentRange(right)
  return !(
    leftBounds.sheetName !== rightBounds.sheetName ||
    leftBounds.endRow < rightBounds.startRow ||
    rightBounds.endRow < leftBounds.startRow ||
    leftBounds.endCol < rightBounds.startCol ||
    rightBounds.endCol < leftBounds.startCol
  )
}

export function createWorkbookAgentRangeChunkPlan(range: CellRangeRef, maxCellsPerChunk: number): WorkbookAgentRangeChunkPlan {
  if (!Number.isInteger(maxCellsPerChunk) || maxCellsPerChunk <= 0) {
    throw new Error(`maxCellsPerChunk must be a positive integer, received ${String(maxCellsPerChunk)}`)
  }
  const normalized = normalizeWorkbookAgentRange(range)
  const columnCount = normalized.endCol - normalized.startCol + 1
  const totalCells = countWorkbookAgentRangeCells(normalized)
  const rowsPerChunk = Math.max(1, Math.floor(maxCellsPerChunk / columnCount))
  const chunks: WorkbookAgentRangeChunk[] = []
  for (let rowStart = normalized.startRow; rowStart <= normalized.endRow; rowStart += rowsPerChunk) {
    const rowEnd = Math.min(normalized.endRow, rowStart + rowsPerChunk - 1)
    chunks.push({
      sheetName: normalized.sheetName,
      startAddress: formatAddress(rowStart, normalized.startCol),
      endAddress: formatAddress(rowEnd, normalized.endCol),
      rowStart,
      rowEnd,
      colStart: normalized.startCol,
      colEnd: normalized.endCol,
      rowCount: rowEnd - rowStart + 1,
      columnCount,
      cellCount: (rowEnd - rowStart + 1) * columnCount,
      index: chunks.length,
    })
  }
  return {
    requestedRange: toWorkbookAgentRangeRef(normalized),
    totalCells,
    maxCellsPerChunk,
    chunkCount: chunks.length,
    chunks,
  }
}

export function ensureWorkbookAgentRangeCellLimit(range: CellRangeRef, limit: number): void {
  const count = countWorkbookAgentRangeCells(range)
  if (count > limit) {
    throw new Error(
      `Range ${range.sheetName}!${range.startAddress}:${range.endAddress} has ${String(count)} cells; tool limit is ${String(limit)} cells per call`,
    )
  }
}
