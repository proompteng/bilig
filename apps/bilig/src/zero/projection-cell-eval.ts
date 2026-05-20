import type { SpreadsheetEngine } from '@bilig/core'
import { formatAddress, parseCellAddress } from '@bilig/formula'
import { ValueTag, type CellRangeRef } from '@bilig/protocol'
import {
  cellCoordinatesWithinBounds,
  intersectRangeBounds,
  normalizeRangeBounds,
  rangeBoundsForSheet,
  type RangeBounds,
} from '@bilig/zero-sync'
import type { CellEvalRow } from './projection.js'

function addAddressesForBounds(addresses: Set<string>, bounds: RangeBounds): void {
  for (let row = bounds.rowStart; row <= bounds.rowEnd; row += 1) {
    for (let col = bounds.colStart; col <= bounds.colEnd; col += 1) {
      addresses.add(formatAddress(row, col))
    }
  }
}

function collectCellEvalAddresses(engine: SpreadsheetEngine, sheetName: string, bounds?: RangeBounds): Set<string> {
  const sheet = engine.workbook.sheetsByName.get(sheetName)
  const addresses = new Set<string>()
  if (!sheet) {
    return addresses
  }

  sheet.grid.forEachCellEntry((_cellIndex, row, col) => {
    if (cellCoordinatesWithinBounds(row, col, bounds)) {
      addresses.add(formatAddress(row, col))
    }
  })

  for (const range of engine.workbook.listStyleRanges(sheetName)) {
    const rangeBounds = rangeBoundsForSheet(sheetName, range.range)
    if (!rangeBounds) {
      continue
    }
    const clippedBounds = bounds ? intersectRangeBounds(rangeBounds, bounds) : rangeBounds
    if (clippedBounds) {
      addAddressesForBounds(addresses, clippedBounds)
    }
  }

  for (const range of engine.workbook.listFormatRanges(sheetName)) {
    const rangeBounds = rangeBoundsForSheet(sheetName, range.range)
    if (!rangeBounds) {
      continue
    }
    const clippedBounds = bounds ? intersectRangeBounds(rangeBounds, bounds) : rangeBounds
    if (clippedBounds) {
      addAddressesForBounds(addresses, clippedBounds)
    }
  }

  return addresses
}

function materializeCellEvalRow(
  engine: SpreadsheetEngine,
  documentId: string,
  revision: number,
  updatedAt: string,
  sheetName: string,
  address: string,
  includeEmpty = false,
): CellEvalRow | null {
  const { row, col } = parseCellAddress(address, sheetName)
  const cell = engine.getCell(sheetName, address)
  if (
    !includeEmpty &&
    cell.value.tag === ValueTag.Empty &&
    cell.flags === 0 &&
    cell.styleId === undefined &&
    cell.numberFormatId === undefined &&
    cell.format === undefined
  ) {
    return null
  }
  return {
    workbookId: documentId,
    sheetName,
    address,
    rowNum: row,
    colNum: col,
    value: cell.value,
    flags: cell.flags,
    version: cell.version,
    styleId: cell.styleId ?? null,
    styleJson: engine.getCellStyle(cell.styleId) ?? null,
    formatId: cell.numberFormatId ?? null,
    formatCode: cell.format ?? null,
    calcRevision: revision,
    updatedAt,
  }
}

export function materializeCellEvalProjection(
  engine: SpreadsheetEngine,
  documentId: string,
  revision: number,
  updatedAt: string,
  changedCellIndices?: readonly number[],
): CellEvalRow[] {
  const entries: CellEvalRow[] = []

  if (changedCellIndices) {
    for (let i = 0; i < changedCellIndices.length; i += 1) {
      const cellIndex = changedCellIndices[i]!
      const qualifiedAddress = engine.workbook.getQualifiedAddress(cellIndex)
      const separatorIndex = qualifiedAddress.lastIndexOf('!')
      if (separatorIndex <= 0 || separatorIndex >= qualifiedAddress.length - 1) {
        continue
      }
      const sheetName = qualifiedAddress.slice(0, separatorIndex)
      const address = qualifiedAddress.slice(separatorIndex + 1)
      const row = materializeCellEvalRow(engine, documentId, revision, updatedAt, sheetName, address, true)
      if (row) {
        entries.push(row)
      }
    }
    return entries
  }

  for (const sheetName of engine.workbook.sheetsByName.keys()) {
    for (const address of collectCellEvalAddresses(engine, sheetName)) {
      const row = materializeCellEvalRow(engine, documentId, revision, updatedAt, sheetName, address)
      if (row) {
        entries.push(row)
      }
    }
  }

  return entries
}

export function materializeCellEvalRangeProjection(
  engine: SpreadsheetEngine,
  documentId: string,
  revision: number,
  updatedAt: string,
  range: CellRangeRef,
): CellEvalRow[] {
  const bounds = normalizeRangeBounds(range)
  const entries: CellEvalRow[] = []
  for (const address of collectCellEvalAddresses(engine, range.sheetName, bounds)) {
    const row = materializeCellEvalRow(engine, documentId, revision, updatedAt, range.sheetName, address)
    if (row) {
      entries.push(row)
    }
  }
  return entries
}
