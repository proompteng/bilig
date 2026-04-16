const CELL_RE = /^\$?([A-Z]+)\$?([1-9][0-9]*)$/
const COLUMN_RE = /^\$?([A-Z]+)$/
const ROW_RE = /^\$?([1-9][0-9]*)$/
const QUALIFIED_RE = /^(?:(?:'((?:[^']|'')+)'|([^!]+))!)?(.+)$/

export interface CellAddress {
  sheetName?: string
  row: number
  col: number
  text: string
}

export interface RowAddress {
  sheetName?: string
  row: number
  text: string
}

export interface ColumnAddress {
  sheetName?: string
  col: number
  text: string
}

export type ParsedReference = CellReference | RowReference | ColumnReference

interface CellReference {
  kind: 'cell'
  sheetName?: string
  row: number
  col: number
  text: string
}

interface RowReference {
  kind: 'row'
  sheetName?: string
  row: number
  text: string
}

interface ColumnReference {
  kind: 'col'
  sheetName?: string
  col: number
  text: string
}

export type RangeAddress = CellRangeAddress | RowRangeAddress | ColumnRangeAddress

export interface CellRangeAddress {
  kind: 'cells'
  sheetName?: string
  start: CellAddress
  end: CellAddress
}

export interface RowRangeAddress {
  kind: 'rows'
  sheetName?: string
  start: RowAddress
  end: RowAddress
}

export interface ColumnRangeAddress {
  kind: 'cols'
  sheetName?: string
  start: ColumnAddress
  end: ColumnAddress
}

export function columnToIndex(column: string): number {
  let value = 0
  for (const char of column) {
    value = value * 26 + (char.charCodeAt(0) - 64)
  }
  return value - 1
}

export function indexToColumn(index: number): string {
  let current = index + 1
  let output = ''
  while (current > 0) {
    const rem = (current - 1) % 26
    output = String.fromCharCode(65 + rem) + output
    current = Math.floor((current - 1) / 26)
  }
  return output
}

export function formatAddress(row: number, col: number): string {
  return `${indexToColumn(col)}${row + 1}`
}

export function isCellReferenceText(value: string): boolean {
  return CELL_RE.test(value.toUpperCase())
}

export function isColumnReferenceText(value: string): boolean {
  return COLUMN_RE.test(value.toUpperCase())
}

export function isRowReferenceText(value: string): boolean {
  return ROW_RE.test(value)
}

export function parseCellAddress(raw: string, defaultSheetName?: string): CellAddress {
  const parsed = parseReference(raw, defaultSheetName)
  if (parsed.kind !== 'cell') {
    throw new Error(`Invalid cell address: ${raw}`)
  }
  const result: CellAddress = {
    col: parsed.col,
    row: parsed.row,
    text: parsed.text,
  }
  if (parsed.sheetName !== undefined) {
    result.sheetName = parsed.sheetName
  }
  return result
}

export function parseRangeAddress(raw: string, defaultSheetName?: string): RangeAddress {
  const separator = raw.indexOf(':')
  if (separator <= 0 || separator >= raw.length - 1) {
    throw new Error(`Invalid range address: ${raw}`)
  }

  const left = raw.slice(0, separator).trim()
  const right = raw.slice(separator + 1).trim()
  const start = parseReference(left, defaultSheetName)
  const end = parseReference(right, start.sheetName ?? defaultSheetName)
  const sheetName = start.sheetName ?? end.sheetName ?? defaultSheetName

  if (start.kind !== end.kind) {
    throw new Error(`Range endpoints must use the same reference type: ${raw}`)
  }
  if (start.sheetName && end.sheetName && start.sheetName !== end.sheetName) {
    throw new Error(`Range endpoints must target the same sheet: ${raw}`)
  }

  switch (start.kind) {
    case 'cell': {
      if (end.kind !== 'cell') {
        throw new Error(`Range endpoints must use the same reference type: ${raw}`)
      }
      const endCell = end
      const row1 = Math.min(start.row, endCell.row)
      const row2 = Math.max(start.row, endCell.row)
      const col1 = Math.min(start.col, endCell.col)
      const col2 = Math.max(start.col, endCell.col)
      const result: CellRangeAddress = {
        kind: 'cells',
        start: { ...start, row: row1, col: col1, text: formatAddress(row1, col1) },
        end: { ...endCell, row: row2, col: col2, text: formatAddress(row2, col2) },
      }
      if (sheetName !== undefined) {
        result.sheetName = sheetName
      }
      return result
    }
    case 'row': {
      if (end.kind !== 'row') {
        throw new Error(`Range endpoints must use the same reference type: ${raw}`)
      }
      const endRow = end
      const row1 = Math.min(start.row, endRow.row)
      const row2 = Math.max(start.row, endRow.row)
      const result: RowRangeAddress = {
        kind: 'rows',
        start: { ...start, row: row1, text: `${row1 + 1}` },
        end: { ...endRow, row: row2, text: `${row2 + 1}` },
      }
      if (sheetName !== undefined) {
        result.sheetName = sheetName
      }
      return result
    }
    case 'col': {
      if (end.kind !== 'col') {
        throw new Error(`Range endpoints must use the same reference type: ${raw}`)
      }
      const endColumn = end
      const col1 = Math.min(start.col, endColumn.col)
      const col2 = Math.max(start.col, endColumn.col)
      const result: ColumnRangeAddress = {
        kind: 'cols',
        start: { ...start, col: col1, text: indexToColumn(col1) },
        end: { ...endColumn, col: col2, text: indexToColumn(col2) },
      }
      if (sheetName !== undefined) {
        result.sheetName = sheetName
      }
      return result
    }
  }
}

export function toQualifiedAddress(sheetName: string, addr: string): string {
  return `${sheetName}!${parseCellAddress(addr, sheetName).text}`
}

export function formatRangeAddress(range: RangeAddress): string {
  const prefix = range.sheetName ? `${quoteSheetNameIfNeeded(range.sheetName)}!` : ''
  return `${prefix}${range.start.text}:${range.end.text}`
}

function parseReference(raw: string, defaultSheetName?: string): ParsedReference {
  const trimmed = raw.trim()
  const qualified = QUALIFIED_RE.exec(trimmed)
  if (!qualified) {
    throw new Error(`Invalid reference: ${raw}`)
  }

  const [, quotedSheet, plainSheet, refPart] = qualified
  const sheetName = quotedSheet?.replaceAll("''", "'") ?? plainSheet ?? defaultSheetName
  const normalizedRefPart = refPart!.toUpperCase()

  const cellMatch = CELL_RE.exec(normalizedRefPart)
  if (cellMatch) {
    const result: CellReference = {
      kind: 'cell',
      col: columnToIndex(cellMatch[1]!),
      row: Number.parseInt(cellMatch[2]!, 10) - 1,
      text: `${cellMatch[1]!}${cellMatch[2]!}`,
    }
    if (sheetName !== undefined) {
      result.sheetName = sheetName
    }
    return result
  }

  const columnMatch = COLUMN_RE.exec(normalizedRefPart)
  if (columnMatch) {
    const result: ColumnReference = {
      kind: 'col',
      col: columnToIndex(columnMatch[1]!),
      text: columnMatch[1]!,
    }
    if (sheetName !== undefined) {
      result.sheetName = sheetName
    }
    return result
  }

  const rowMatch = ROW_RE.exec(refPart!)
  if (rowMatch) {
    const result: RowReference = {
      kind: 'row',
      row: Number.parseInt(rowMatch[1]!, 10) - 1,
      text: rowMatch[1]!,
    }
    if (sheetName !== undefined) {
      result.sheetName = sheetName
    }
    return result
  }

  throw new Error(`Invalid reference: ${raw}`)
}

function quoteSheetNameIfNeeded(sheetName: string): string {
  return /^[A-Za-z0-9_.$]+$/.test(sheetName) ? sheetName : `'${sheetName.replaceAll("'", "''")}'`
}
