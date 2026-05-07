import { ErrorCode, ValueTag } from '@bilig/protocol'
import type { CellSnapshot, LiteralInput } from '@bilig/protocol'

interface CsvCellInput {
  formula?: string
  value?: LiteralInput
}

export type CsvDelimiter = ',' | ';' | '\t'
export type CsvDecimalSeparator = '.' | ','

export interface CsvParseOptions {
  delimiter?: CsvDelimiter
  decimalSeparator?: CsvDecimalSeparator
}

export interface ResolvedCsvParseOptions {
  delimiter: CsvDelimiter
  decimalSeparator: CsvDecimalSeparator
}

function escapeCsvValue(value: string): string {
  if (!/[",\n\r]/.test(value)) {
    return value
  }
  return `"${value.replaceAll('"', '""')}"`
}

export function cellToCsvValue(cell: CellSnapshot): string {
  if (cell.formula !== undefined) {
    return `=${cell.formula}`
  }

  switch (cell.value.tag) {
    case ValueTag.Empty:
      return ''
    case ValueTag.Number:
      return String(cell.value.value)
    case ValueTag.Boolean:
      return cell.value.value ? 'TRUE' : 'FALSE'
    case ValueTag.String:
      return cell.value.value
    case ValueTag.Error:
      return `#${ErrorCode[cell.value.code] ?? cell.value.code}`
  }
}

export function serializeCsv(rows: string[][]): string {
  return rows.map((row) => row.map((value) => escapeCsvValue(value)).join(',')).join('\n')
}

export function resolveCsvParseOptions(csv: string, options: CsvParseOptions = {}): ResolvedCsvParseOptions {
  const delimiter = options.delimiter ?? detectCsvDelimiter(csv)
  return {
    delimiter,
    decimalSeparator: options.decimalSeparator ?? (delimiter !== ',' && hasDecimalCommaCell(csv, delimiter) ? ',' : '.'),
  }
}

export function parseCsv(csv: string, options: CsvParseOptions = {}): string[][] {
  const { delimiter } = resolveCsvParseOptions(csv, options)
  const rows: string[][] = []
  let currentRow: string[] = []
  let currentValue = ''
  let index = 0
  let inQuotes = false

  while (index < csv.length) {
    const char = csv[index]!
    const nextChar = csv[index + 1]

    if (inQuotes) {
      if (char === '"' && nextChar === '"') {
        currentValue += '"'
        index += 2
        continue
      }
      if (char === '"') {
        inQuotes = false
        index += 1
        continue
      }
      currentValue += char
      index += 1
      continue
    }

    if (char === '"') {
      inQuotes = true
      index += 1
      continue
    }

    if (char === delimiter) {
      currentRow.push(currentValue)
      currentValue = ''
      index += 1
      continue
    }

    if (char === '\r' || char === '\n') {
      currentRow.push(currentValue)
      currentValue = ''
      rows.push(currentRow)
      currentRow = []
      if (char === '\r' && nextChar === '\n') {
        index += 2
      } else {
        index += 1
      }
      continue
    }

    currentValue += char
    index += 1
  }

  currentRow.push(currentValue)
  if (currentRow.length > 1 || currentRow[0] !== '' || rows.length > 0) {
    rows.push(currentRow)
  }

  return rows
}

export function parseCsvCellInput(raw: string, options: CsvParseOptions = {}): CsvCellInput | undefined {
  const normalized = raw.trim()
  if (normalized === '') {
    return undefined
  }
  if (normalized.startsWith('=')) {
    return { formula: normalized.slice(1) }
  }
  if (normalized === 'TRUE' || normalized === 'FALSE') {
    return { value: normalized === 'TRUE' }
  }
  const accountingNumber = parseAccountingNumberInput(normalized, options.decimalSeparator ?? '.')
  if (accountingNumber !== null) {
    return { value: accountingNumber }
  }
  return { value: raw }
}

function detectCsvDelimiter(csv: string): CsvDelimiter {
  const commaScore = countDelimiterOutsideQuotes(csv, ',')
  const semicolonScore = countDelimiterOutsideQuotes(csv, ';')
  const tabScore = countDelimiterOutsideQuotes(csv, '\t')
  if (tabScore > commaScore && tabScore > semicolonScore) {
    return '\t'
  }
  return semicolonScore > commaScore ? ';' : ','
}

function countDelimiterOutsideQuotes(csv: string, delimiter: CsvDelimiter): number {
  let count = 0
  let index = 0
  let inQuotes = false
  while (index < csv.length) {
    const char = csv[index]!
    const nextChar = csv[index + 1]
    if (inQuotes) {
      if (char === '"' && nextChar === '"') {
        index += 2
        continue
      }
      if (char === '"') {
        inQuotes = false
      }
      index += 1
      continue
    }
    if (char === '"') {
      inQuotes = true
      index += 1
      continue
    }
    if (char === delimiter) {
      count += 1
    }
    index += 1
  }
  return count
}

function hasDecimalCommaCell(csv: string, delimiter: CsvDelimiter): boolean {
  const rows = parseCsv(csv, { delimiter, decimalSeparator: '.' })
  return rows.some((row) => row.some((value) => /^-?\d+,\d+(%?)$/u.test(value.trim())))
}

function parseAccountingNumberInput(normalized: string, decimalSeparator: CsvDecimalSeparator): number | null {
  let text = normalized
  let sign = 1

  if (text.startsWith('(') && text.endsWith(')')) {
    sign = -1
    text = text.slice(1, -1).trim()
  }

  if (text.startsWith('-')) {
    sign *= -1
    text = text.slice(1).trim()
  }

  if (text.startsWith('$')) {
    text = text.slice(1).trim()
  }

  if (text.startsWith('-')) {
    sign *= -1
    text = text.slice(1).trim()
  }

  const isPercent = text.endsWith('%')
  if (isPercent) {
    text = text.slice(0, -1).trim()
  }

  const groupSeparator = decimalSeparator === ',' ? '.' : ','
  const decimal = escapeRegExp(decimalSeparator)
  const group = escapeRegExp(groupSeparator)
  const plainPositiveNumericRe = new RegExp(`^\\d+(?:${decimal}\\d+)?$`, 'u')
  const groupedNumericRe = new RegExp(`^\\d{1,3}(?:${group}\\d{3})+(?:${decimal}\\d+)?$`, 'u')
  if (!plainPositiveNumericRe.test(text) && !groupedNumericRe.test(text)) {
    return null
  }

  const parsed = Number(text.replaceAll(groupSeparator, '').replace(decimalSeparator, '.'))
  if (!Number.isFinite(parsed)) {
    return null
  }

  return (sign * parsed) / (isPercent ? 100 : 1)
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/gu, '\\$&')
}
