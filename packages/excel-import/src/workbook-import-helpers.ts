import type { LiteralInput } from '@bilig/protocol'

export interface ImportedWorkbookSheetPreview {
  name: string
  rowCount: number
  columnCount: number
  nonEmptyCellCount: number
  previewRows: readonly (readonly string[])[]
}

export function normalizeWorkbookName(fileName: string): string {
  const trimmed = fileName.trim()
  if (trimmed.length === 0) {
    return 'Imported workbook'
  }
  return trimmed.replace(/\.(xlsx|xlsm|xlsb|csv)$/i, '') || 'Imported workbook'
}

export function normalizeCsvSheetName(workbookName: string): string {
  const trimmed = workbookName.trim()
  return trimmed.length > 0 ? trimmed : 'Sheet1'
}

export function toLiteralInput(value: unknown): LiteralInput | undefined {
  if (value === null || value === undefined) {
    return undefined
  }
  if (typeof value === 'string') {
    return value.replace(/\r\n?/gu, '\n')
  }
  if (typeof value === 'number' || typeof value === 'boolean') {
    return value
  }
  if (value instanceof Date) {
    return value.getTime()
  }
  return undefined
}

export function toDisplayText(value: LiteralInput | undefined): string {
  if (value === null || value === undefined) {
    return ''
  }
  if (typeof value === 'boolean') {
    return value ? 'TRUE' : 'FALSE'
  }
  return String(value)
}

export function readImportedAlignmentBoolean(value: unknown): boolean | undefined {
  if (typeof value === 'boolean') {
    return value
  }
  if (value === '1' || value === 'true') {
    return true
  }
  if (value === '0' || value === 'false') {
    return false
  }
  return undefined
}

export function readImportedAlignmentNumber(value: unknown): number | null {
  if (typeof value === 'number' && Number.isFinite(value)) {
    return value
  }
  if (typeof value === 'string' && value.trim().length > 0) {
    const number = Number(value)
    return Number.isFinite(number) ? number : null
  }
  return null
}

export function createSheetPreview(input: {
  name: string
  rowCount: number
  columnCount: number
  nonEmptyCellCount: number
  readCellText: (row: number, col: number) => string
}): ImportedWorkbookSheetPreview {
  const previewRows: string[][] = []
  const rowLimit = Math.min(input.rowCount, 8)
  const columnLimit = Math.min(input.columnCount, 6)
  for (let row = 0; row < rowLimit; row += 1) {
    const previewRow: string[] = []
    for (let col = 0; col < columnLimit; col += 1) {
      previewRow.push(input.readCellText(row, col))
    }
    previewRows.push(previewRow)
  }
  return {
    name: input.name,
    rowCount: input.rowCount,
    columnCount: input.columnCount,
    nonEmptyCellCount: input.nonEmptyCellCount,
    previewRows,
  }
}
