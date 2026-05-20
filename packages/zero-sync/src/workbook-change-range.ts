import { formatAddress, parseCellAddress } from '@bilig/formula'
import type { CellRangeRef } from '@bilig/protocol'
import { normalizeRangeBounds, type RangeBounds } from './range-bounds.js'

export type WorkbookChangeRangeScope = 'cells' | 'rows' | 'columns' | 'sheet'

export interface WorkbookChangeRange {
  readonly sheetName: string
  readonly startAddress: string
  readonly endAddress: string
  readonly scope?: WorkbookChangeRangeScope
}

export interface WorkbookChangeRangeBounds extends RangeBounds {
  readonly scope: WorkbookChangeRangeScope
}

export function isWorkbookChangeRangeScope(value: unknown): value is WorkbookChangeRangeScope {
  return value === 'cells' || value === 'rows' || value === 'columns' || value === 'sheet'
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function isNonEmptyString(value: unknown): value is string {
  return typeof value === 'string' && value.length > 0
}

export function normalizeWorkbookChangeRange(value: unknown): WorkbookChangeRange | null {
  if (!isRecord(value)) {
    return null
  }
  return canonicalizeWorkbookChangeRange(
    {
      sheetName: value['sheetName'],
      startAddress: value['startAddress'],
      endAddress: value['endAddress'],
    },
    value['scope'],
  )
}

export function canonicalizeWorkbookChangeRange(
  range: {
    readonly sheetName: unknown
    readonly startAddress: unknown
    readonly endAddress: unknown
  },
  scope?: unknown,
): WorkbookChangeRange | null {
  if (!isNonEmptyString(range.sheetName) || !isNonEmptyString(range.startAddress) || !isNonEmptyString(range.endAddress)) {
    return null
  }
  if (scope !== undefined && !isWorkbookChangeRangeScope(scope)) {
    return null
  }
  const cellRange: CellRangeRef = {
    sheetName: range.sheetName,
    startAddress: range.startAddress,
    endAddress: range.endAddress,
  }
  let bounds: RangeBounds
  try {
    bounds = normalizeRangeBounds(cellRange)
  } catch {
    return null
  }
  return {
    sheetName: bounds.sheetName,
    startAddress: formatAddress(bounds.rowStart, bounds.colStart),
    endAddress: formatAddress(bounds.rowEnd, bounds.colEnd),
    ...(isWorkbookChangeRangeScope(scope) && scope !== 'cells' ? { scope } : {}),
  }
}

export function normalizeWorkbookChangeRangeBounds(value: unknown): WorkbookChangeRangeBounds | null {
  const range = normalizeWorkbookChangeRange(value)
  if (!range) {
    return null
  }
  const bounds = normalizeRangeBounds(range)
  return {
    ...bounds,
    scope: range.scope ?? 'cells',
  }
}

export function workbookChangeRangeFromAddresses(
  sheetName: string,
  addresses: readonly string[],
  scope?: WorkbookChangeRangeScope,
): WorkbookChangeRange | null {
  if (addresses.length === 0) {
    return null
  }
  let rowStart = Number.POSITIVE_INFINITY
  let rowEnd = Number.NEGATIVE_INFINITY
  let colStart = Number.POSITIVE_INFINITY
  let colEnd = Number.NEGATIVE_INFINITY
  for (const address of addresses) {
    const parsed = parseCellAddress(address, sheetName)
    rowStart = Math.min(rowStart, parsed.row)
    rowEnd = Math.max(rowEnd, parsed.row)
    colStart = Math.min(colStart, parsed.col)
    colEnd = Math.max(colEnd, parsed.col)
  }
  return canonicalizeWorkbookChangeRange(
    {
      sheetName,
      startAddress: formatAddress(rowStart, colStart),
      endAddress: formatAddress(rowEnd, colEnd),
    },
    scope,
  )
}

export function isWorkbookChangeRange(value: unknown): value is WorkbookChangeRange {
  return normalizeWorkbookChangeRange(value) !== null
}
