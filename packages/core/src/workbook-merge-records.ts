import { formatAddress, parseCellAddress } from '@bilig/formula'
import type { CellRangeRef } from '@bilig/protocol'
import type { WorkbookMergeRangeRecord } from './workbook-metadata-types.js'

interface NormalizedRange {
  readonly sheetName: string
  readonly startAddress: string
  readonly endAddress: string
  readonly startRow: number
  readonly endRow: number
  readonly startCol: number
  readonly endCol: number
}

export function canonicalMergeRangeRef(range: CellRangeRef): WorkbookMergeRangeRecord {
  const normalized = normalizeRange(range)
  return {
    sheetName: normalized.sheetName,
    startAddress: normalized.startAddress,
    endAddress: normalized.endAddress,
  }
}

export function cloneMergeRangeRecord(record: WorkbookMergeRangeRecord): WorkbookMergeRangeRecord {
  return {
    sheetName: record.sheetName,
    startAddress: record.startAddress,
    endAddress: record.endAddress,
  }
}

export function mergeRangeKey(range: CellRangeRef): string {
  const normalized = normalizeRange(range)
  return `${normalized.sheetName}:${normalized.startAddress}:${normalized.endAddress}`
}

export function isSingleCellMergeRange(range: CellRangeRef): boolean {
  const normalized = normalizeRange(range)
  return normalized.startRow === normalized.endRow && normalized.startCol === normalized.endCol
}

export function rangesIntersect(left: CellRangeRef, right: CellRangeRef): boolean {
  if (left.sheetName !== right.sheetName) {
    return false
  }
  const normalizedLeft = normalizeRange(left)
  const normalizedRight = normalizeRange(right)
  return !(
    normalizedLeft.endRow < normalizedRight.startRow ||
    normalizedRight.endRow < normalizedLeft.startRow ||
    normalizedLeft.endCol < normalizedRight.startCol ||
    normalizedRight.endCol < normalizedLeft.startCol
  )
}

export function rangeContainsAddress(range: CellRangeRef, sheetName: string, address: string): boolean {
  if (range.sheetName !== sheetName) {
    return false
  }
  const normalized = normalizeRange(range)
  const parsed = parseCellAddress(address, sheetName)
  return (
    parsed.row >= normalized.startRow &&
    parsed.row <= normalized.endRow &&
    parsed.col >= normalized.startCol &&
    parsed.col <= normalized.endCol
  )
}

export function normalizeMergeRangeBounds(range: CellRangeRef): NormalizedRange {
  return normalizeRange(range)
}

function normalizeRange(range: CellRangeRef): NormalizedRange {
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
