export type WorkbookChangeRangeScope = 'cells' | 'rows' | 'columns' | 'sheet'

export interface WorkbookChangeRange {
  readonly sheetName: string
  readonly startAddress: string
  readonly endAddress: string
  readonly scope?: WorkbookChangeRangeScope
}

export function isWorkbookChangeRangeScope(value: unknown): value is WorkbookChangeRangeScope {
  return value === 'cells' || value === 'rows' || value === 'columns' || value === 'sheet'
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

export function normalizeWorkbookChangeRange(value: unknown): WorkbookChangeRange | null {
  if (!isRecord(value)) {
    return null
  }
  const sheetName = value['sheetName']
  const startAddress = value['startAddress']
  const endAddress = value['endAddress']
  if (typeof sheetName !== 'string' || typeof startAddress !== 'string' || typeof endAddress !== 'string') {
    return null
  }
  const scope = value['scope']
  return {
    sheetName,
    startAddress,
    endAddress,
    ...(isWorkbookChangeRangeScope(scope) && scope !== 'cells' ? { scope } : {}),
  }
}

export function isWorkbookChangeRange(value: unknown): value is WorkbookChangeRange {
  if (!isRecord(value)) {
    return false
  }
  const scope = value['scope']
  return normalizeWorkbookChangeRange(value) !== null && (scope === undefined || isWorkbookChangeRangeScope(scope))
}
