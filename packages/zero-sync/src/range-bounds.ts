import { parseCellAddress } from '@bilig/formula'
import type { CellRangeRef } from '@bilig/protocol'

export interface RangeBounds {
  readonly sheetName: string
  readonly rowStart: number
  readonly rowEnd: number
  readonly colStart: number
  readonly colEnd: number
}

export function normalizeRangeBounds(range: CellRangeRef): RangeBounds {
  return normalizeAddressBounds(range.sheetName, range.startAddress, range.endAddress)
}

export function normalizeAddressBounds(sheetName: string, startAddress: string, endAddress: string): RangeBounds {
  const start = parseCellAddress(startAddress, sheetName)
  const end = parseCellAddress(endAddress, sheetName)
  return {
    sheetName,
    rowStart: Math.min(start.row, end.row),
    rowEnd: Math.max(start.row, end.row),
    colStart: Math.min(start.col, end.col),
    colEnd: Math.max(start.col, end.col),
  }
}

export function rangeBoundsForSheet(sheetName: string, range?: CellRangeRef): RangeBounds | null | undefined {
  if (!range) {
    return undefined
  }
  if (range.sheetName !== sheetName) {
    return null
  }
  return normalizeRangeBounds(range)
}

export function intersectRangeBounds(left: RangeBounds, right: RangeBounds): RangeBounds | null {
  if (left.sheetName !== right.sheetName) {
    return null
  }
  const rowStart = Math.max(left.rowStart, right.rowStart)
  const rowEnd = Math.min(left.rowEnd, right.rowEnd)
  const colStart = Math.max(left.colStart, right.colStart)
  const colEnd = Math.min(left.colEnd, right.colEnd)
  return rowStart <= rowEnd && colStart <= colEnd
    ? {
        sheetName: left.sheetName,
        rowStart,
        rowEnd,
        colStart,
        colEnd,
      }
    : null
}

export function cellCoordinatesWithinBounds(row: number, col: number, bounds?: RangeBounds | null): boolean {
  if (bounds === null) {
    return false
  }
  if (bounds === undefined) {
    return true
  }
  return row >= bounds.rowStart && row <= bounds.rowEnd && col >= bounds.colStart && col <= bounds.colEnd
}
