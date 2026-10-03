import { formatAddress, formatRangeAddress, indexToColumn, parseCellAddress, parseRangeAddress } from '@bilig/formula'
import type { CellRangeRef } from '@bilig/protocol'
import type { WorkPaperAddressFormatOptions, WorkPaperCellAddress, WorkPaperCellRange } from './work-paper-types.js'

const RUNTIME_COLUMN_LABEL_CACHE: string[] = []
const RUNTIME_A1_CACHE_COLUMN_LIMIT = 64
const RUNTIME_A1_CACHE_ROW_LIMIT = 8192
const RUNTIME_A1_CACHE: string[] = []

export function quoteSheetNameIfNeeded(sheetName: string): string {
  return /^[A-Za-z0-9_.$]+$/.test(sheetName) ? sheetName : `'${sheetName.replaceAll("'", "''")}'`
}

export function formatQualifiedCellAddress(sheetName: string | undefined, row: number, col: number): string {
  const base = formatAddress(row, col)
  return sheetName ? `${quoteSheetNameIfNeeded(sheetName)}!${base}` : base
}

export function formatTrackedA1(row: number, col: number): string {
  if (row >= 0 && row < RUNTIME_A1_CACHE_ROW_LIMIT && col >= 0 && col < RUNTIME_A1_CACHE_COLUMN_LIMIT) {
    const cacheKey = row * RUNTIME_A1_CACHE_COLUMN_LIMIT + col
    let cached = RUNTIME_A1_CACHE[cacheKey]
    if (cached === undefined) {
      let column = RUNTIME_COLUMN_LABEL_CACHE[col]
      if (column === undefined) {
        column = indexToColumn(col)
        RUNTIME_COLUMN_LABEL_CACHE[col] = column
      }
      cached = `${column}${row + 1}`
      RUNTIME_A1_CACHE[cacheKey] = cached
    }
    return cached
  }
  let column = RUNTIME_COLUMN_LABEL_CACHE[col]
  if (column === undefined) {
    column = indexToColumn(col)
    RUNTIME_COLUMN_LABEL_CACHE[col] = column
  }
  return `${column}${row + 1}`
}

export function sourceRangeRef(sheetName: string, range: WorkPaperCellRange): CellRangeRef {
  return {
    sheetName,
    startAddress: formatAddress(range.start.row, range.start.col),
    endAddress: formatAddress(range.end.row, range.end.col),
  }
}

export function resolveDefaultWorkPaperSheetName(args: {
  readonly defaultSheetId?: number
  readonly sheets: readonly { readonly name: string }[]
  readonly sheetName: (sheetId: number) => string
}): string | undefined {
  return args.defaultSheetId !== undefined
    ? args.sheetName(args.defaultSheetId)
    : args.sheets.length === 1
      ? args.sheets[0]!.name
      : undefined
}

export function parseWorkPaperCellAddressText(args: {
  readonly value: string
  readonly defaultSheetName?: string
  readonly requireSheetId: (sheetName: string) => number
}): WorkPaperCellAddress | undefined {
  try {
    const parsed = parseCellAddress(args.value, args.defaultSheetName)
    const sheetName = parsed.sheetName ?? args.defaultSheetName
    if (!sheetName) {
      return undefined
    }
    return {
      sheet: args.requireSheetId(sheetName),
      row: parsed.row,
      col: parsed.col,
    }
  } catch {
    return undefined
  }
}

export function parseWorkPaperCellRangeText(args: {
  readonly value: string
  readonly defaultSheetName?: string
  readonly requireSheetId: (sheetName: string) => number
}): WorkPaperCellRange | undefined {
  try {
    const parsed = parseRangeAddress(args.value, args.defaultSheetName)
    if (parsed.kind !== 'cells') {
      return undefined
    }
    const sheetName = parsed.sheetName ?? args.defaultSheetName
    if (!sheetName) {
      return undefined
    }
    const sheetId = args.requireSheetId(sheetName)
    return {
      start: { sheet: sheetId, row: parsed.start.row, col: parsed.start.col },
      end: { sheet: sheetId, row: parsed.end.row, col: parsed.end.col },
    }
  } catch {
    return undefined
  }
}

export function formatWorkPaperCellAddressText(args: {
  readonly address: WorkPaperCellAddress
  readonly optionsOrContextSheetId?: WorkPaperAddressFormatOptions | number
  readonly sheetName: (sheetId: number) => string
}): string {
  const optionsOrContextSheetId = args.optionsOrContextSheetId ?? {}
  const includeSheetName =
    typeof optionsOrContextSheetId === 'number'
      ? optionsOrContextSheetId !== args.address.sheet
      : optionsOrContextSheetId.includeSheetName === true
  return formatQualifiedCellAddress(includeSheetName ? args.sheetName(args.address.sheet) : undefined, args.address.row, args.address.col)
}

export function formatWorkPaperCellRangeText(args: {
  readonly range: WorkPaperCellRange
  readonly optionsOrContextSheetId?: WorkPaperAddressFormatOptions | number
  readonly sheetName: (sheetId: number) => string
}): string {
  const optionsOrContextSheetId = args.optionsOrContextSheetId ?? {}
  const includeSheetName =
    typeof optionsOrContextSheetId === 'number'
      ? optionsOrContextSheetId !== args.range.start.sheet
      : optionsOrContextSheetId.includeSheetName === true
  const rangeAddress: Parameters<typeof formatRangeAddress>[0] = {
    kind: 'cells',
    start: {
      row: args.range.start.row,
      col: args.range.start.col,
      text: formatAddress(args.range.start.row, args.range.start.col),
    },
    end: {
      row: args.range.end.row,
      col: args.range.end.col,
      text: formatAddress(args.range.end.row, args.range.end.col),
    },
  }
  if (includeSheetName) {
    rangeAddress.sheetName = args.sheetName(args.range.start.sheet)
  }
  return formatRangeAddress(rangeAddress)
}
