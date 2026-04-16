import { SpreadsheetEngine } from '@bilig/core'
import type { CellRangeRef, EngineEvent } from '@bilig/protocol'

export interface RangeBounds {
  startRow: number
  endRow: number
  startCol: number
  endCol: number
}

export function getRangeBounds(range: CellRangeRef): RangeBounds {
  const [startColPart, startRowPart] = splitAddress(range.startAddress)
  const [endColPart, endRowPart] = splitAddress(range.endAddress)
  return {
    startCol: decodeColumn(startColPart),
    endCol: decodeColumn(endColPart),
    startRow: Number.parseInt(startRowPart, 10),
    endRow: Number.parseInt(endRowPart, 10),
  }
}

export function iterateRangeBounds(bounds: RangeBounds): string[] {
  const width = bounds.endCol - bounds.startCol + 1
  const height = bounds.endRow - bounds.startRow + 1
  const addresses = Array.from<string>({ length: width * height })
  let index = 0
  for (let row = bounds.startRow; row <= bounds.endRow; row += 1) {
    for (let col = bounds.startCol; col <= bounds.endCol; col += 1) {
      addresses[index] = `${encodeColumn(col)}${row}`
      index += 1
    }
  }
  return addresses
}

export function iterateRange(range: CellRangeRef): string[] {
  return iterateRangeBounds(getRangeBounds(range))
}

function splitAddress(address: string): [string, string] {
  const match = /^([A-Z]+)(\d+)$/i.exec(address.trim())
  if (!match) {
    throw new Error(`Invalid cell address: ${address}`)
  }
  return [match[1]!.toUpperCase(), match[2]!]
}

function decodeColumn(column: string): number {
  let value = 0
  for (let index = 0; index < column.length; index += 1) {
    value = value * 26 + (column.charCodeAt(index) - 64)
  }
  return value
}

function encodeColumn(value: number): string {
  let next = value
  let output = ''
  while (next > 0) {
    const remainder = (next - 1) % 26
    output = String.fromCharCode(65 + remainder) + output
    next = Math.floor((next - 1) / 26)
  }
  return output
}

export function cellCountForRange(range: CellRangeRef): number {
  const bounds = getRangeBounds(range)
  const width = bounds.endCol - bounds.startCol + 1
  const height = bounds.endRow - bounds.startRow + 1
  return width * height
}

function collectChangedAddressesInRange(range: CellRangeRef, bounds: RangeBounds, changedCells: EngineEvent['changedCells']): string[] {
  const changedAddresses = Array.from<string>({ length: changedCells.length })
  let changedAddressCount = 0

  for (let index = 0; index < changedCells.length; index += 1) {
    const change = changedCells[index]!
    if (change.sheetName !== range.sheetName) {
      continue
    }
    const row = change.address.row
    const col = change.address.col
    if (col < bounds.startCol || col > bounds.endCol || row < bounds.startRow || row > bounds.endRow) {
      continue
    }
    changedAddresses[changedAddressCount] = change.a1
    changedAddressCount += 1
  }

  return changedAddresses.slice(0, changedAddressCount)
}

function collectAddressesForIntersection(rangeBounds: RangeBounds, invalidatedRange: CellRangeRef): string[] {
  const invalidatedBounds = getRangeBounds(invalidatedRange)
  const startCol = Math.max(rangeBounds.startCol, invalidatedBounds.startCol)
  const endCol = Math.min(rangeBounds.endCol, invalidatedBounds.endCol)
  const startRow = Math.max(rangeBounds.startRow, invalidatedBounds.startRow)
  const endRow = Math.min(rangeBounds.endRow, invalidatedBounds.endRow)

  if (startCol > endCol || startRow > endRow) {
    return []
  }

  return iterateRangeBounds({ startCol, endCol, startRow, endRow })
}

export function collectChangedAddressesForEvent(
  _engine: SpreadsheetEngine,
  range: CellRangeRef,
  bounds: RangeBounds,
  event: EngineEvent,
): string[] {
  if (event.invalidation === 'full') {
    return iterateRangeBounds(bounds)
  }

  const changedAddresses = new Set(collectChangedAddressesInRange(range, bounds, event.changedCells))
  for (let index = 0; index < event.invalidatedRanges.length; index += 1) {
    const invalidatedRange = event.invalidatedRanges[index]!
    if (invalidatedRange.sheetName !== range.sheetName) {
      continue
    }
    collectAddressesForIntersection(bounds, invalidatedRange).forEach((address) => {
      changedAddresses.add(address)
    })
  }
  return [...changedAddresses]
}
