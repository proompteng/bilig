import { ValueTag } from './enums.js'
import { formatErrorCode } from './types.js'
import type {
  CellDateStyle,
  CellNumberFormatInput,
  CellNumberFormatKind,
  CellNumberFormatPreset,
  CellNumberFormatRecord,
  CellValue,
} from './types.js'

const EXCEL_EPOCH_OFFSET = 25569
const DAY_MS = 86_400_000

export function buildCellNumberFormatCode(input: CellNumberFormatInput): string {
  if (typeof input === 'string') {
    const trimmed = input.trim()
    return trimmed.length === 0 ? 'general' : trimmed
  }
  const preset = normalizeCellNumberFormatPreset(input)
  switch (preset.kind) {
    case 'general':
    case 'text':
      return preset.kind
    case 'number':
      return `number:${preset.decimals}:${preset.useGrouping ? 1 : 0}`
    case 'percent':
      return `percent:${preset.decimals}`
    case 'currency':
    case 'accounting':
      return [preset.kind, preset.currency, preset.decimals, preset.useGrouping ? 1 : 0, preset.negativeStyle, preset.zeroStyle].join(':')
    case 'date':
    case 'time':
    case 'datetime':
      return `${preset.kind}:${preset.dateStyle}`
  }
}

export function parseCellNumberFormatCode(code: string | undefined): CellNumberFormatPreset {
  const normalized = (code ?? 'general').trim() || 'general'
  if (normalized === 'currency-usd') {
    return {
      kind: 'currency',
      currency: 'USD',
      decimals: 2,
      useGrouping: true,
      negativeStyle: 'minus',
      zeroStyle: 'zero',
    }
  }
  const [kindToken, ...rest] = normalized.split(':')
  const kind = toNumberFormatKind(kindToken)
  switch (kind) {
    case 'general':
    case 'text':
      return { kind }
    case 'number':
      return normalizeCellNumberFormatPreset({
        kind,
        decimals: toInteger(rest[0], 2),
        useGrouping: rest[1] !== '0',
      })
    case 'percent':
      return normalizeCellNumberFormatPreset({
        kind,
        decimals: toInteger(rest[0], 2),
      })
    case 'currency':
    case 'accounting':
      return normalizeCellNumberFormatPreset({
        kind,
        currency: rest[0] || 'USD',
        decimals: toInteger(rest[1], 2),
        useGrouping: rest[2] !== '0',
        negativeStyle: rest[3] === 'parentheses' ? 'parentheses' : 'minus',
        zeroStyle: rest[4] === 'dash' ? 'dash' : 'zero',
      })
    case 'date':
    case 'time':
    case 'datetime':
      return normalizeCellNumberFormatPreset({
        kind,
        dateStyle: rest[0] === 'iso' ? 'iso' : 'short',
      })
  }
}

export function normalizeCellNumberFormatPreset(input: CellNumberFormatPreset): CellNumberFormatPreset {
  switch (input.kind) {
    case 'general':
    case 'text':
      return { kind: input.kind }
    case 'number':
      return {
        kind: 'number',
        decimals: clampDecimals(input.decimals ?? 2),
        useGrouping: input.useGrouping ?? true,
      }
    case 'percent':
      return {
        kind: 'percent',
        decimals: clampDecimals(input.decimals ?? 2),
      }
    case 'currency':
      return {
        kind: 'currency',
        currency: normalizeCurrency(input.currency),
        decimals: clampDecimals(input.decimals ?? 2),
        useGrouping: input.useGrouping ?? true,
        negativeStyle: input.negativeStyle ?? 'minus',
        zeroStyle: input.zeroStyle ?? 'zero',
      }
    case 'accounting':
      return {
        kind: 'accounting',
        currency: normalizeCurrency(input.currency),
        decimals: clampDecimals(input.decimals ?? 2),
        useGrouping: input.useGrouping ?? true,
        negativeStyle: input.negativeStyle ?? 'parentheses',
        zeroStyle: input.zeroStyle ?? 'dash',
      }
    case 'date':
    case 'time':
    case 'datetime':
      return {
        kind: input.kind,
        dateStyle: normalizeDateStyle(input.dateStyle),
      }
  }
}

export function createCellNumberFormatRecord(id: string, input: CellNumberFormatInput): CellNumberFormatRecord {
  const code = buildCellNumberFormatCode(input)
  return {
    id,
    code,
    kind: parseCellNumberFormatCode(code).kind,
  }
}

export function getCellNumberFormatKind(code: string | undefined): CellNumberFormatKind {
  return parseCellNumberFormatCode(code).kind
}

export function shouldRightAlignCell(value: CellValue, formatCode: string | undefined): boolean {
  const format = parseCellNumberFormatCode(formatCode)
  if (format.kind === 'number' || format.kind === 'currency' || format.kind === 'accounting' || format.kind === 'percent') {
    return true
  }
  return value.tag === ValueTag.Number
}

export function formatCellDisplayValue(value: CellValue, formatCode: string | undefined): string {
  const format = parseCellNumberFormatCode(formatCode)
  switch (value.tag) {
    case ValueTag.Empty:
      return ''
    case ValueTag.Boolean:
      return value.value ? 'TRUE' : 'FALSE'
    case ValueTag.String:
      return value.value
    case ValueTag.Error:
      return formatErrorCode(value.code)
    case ValueTag.Number:
      return formatNumberValue(value.value, format)
  }
}

function formatNumberValue(value: number, format: CellNumberFormatPreset): string {
  const normalized = normalizeCellNumberFormatPreset(format)
  switch (format.kind) {
    case 'general':
      return String(value)
    case 'text':
      return String(value)
    case 'number':
      return new Intl.NumberFormat('en-US', {
        minimumFractionDigits: normalized.decimals,
        maximumFractionDigits: normalized.decimals,
        useGrouping: normalized.useGrouping,
      }).format(value)
    case 'percent':
      return new Intl.NumberFormat('en-US', {
        style: 'percent',
        minimumFractionDigits: normalized.decimals,
        maximumFractionDigits: normalized.decimals,
      }).format(value)
    case 'currency':
      return formatSignedCurrency(value, normalized, false)
    case 'accounting':
      return formatSignedCurrency(value, normalized, true)
    case 'date':
      return formatDateSerial(value, normalized.dateStyle, 'date')
    case 'time':
      return formatDateSerial(value, normalized.dateStyle, 'time')
    case 'datetime':
      return formatDateSerial(value, normalized.dateStyle, 'datetime')
  }
}

function formatSignedCurrency(value: number, format: CellNumberFormatPreset, accounting: boolean): string {
  if (format.kind !== 'currency' && format.kind !== 'accounting') {
    return String(value)
  }
  if (value === 0 && format.zeroStyle === 'dash') {
    return '—'
  }
  const absolute = Math.abs(value)
  const formatter = new Intl.NumberFormat('en-US', {
    style: 'currency',
    currency: format.currency,
    currencyDisplay: 'symbol',
    currencySign: accounting || format.negativeStyle === 'parentheses' ? 'accounting' : 'standard',
    minimumFractionDigits: format.decimals,
    maximumFractionDigits: format.decimals,
    useGrouping: format.useGrouping,
  })
  const rendered = formatter.format(value < 0 && format.negativeStyle !== 'parentheses' ? absolute : value)
  if (value < 0 && format.negativeStyle === 'minus' && !rendered.startsWith('-')) {
    return `-${rendered}`
  }
  return rendered
}

function formatDateSerial(serial: number, style: CellDateStyle | undefined, kind: 'date' | 'time' | 'datetime'): string {
  const date = excelSerialToDate(serial)
  if (Number.isNaN(date.getTime())) {
    return String(serial)
  }
  if (style === 'iso') {
    switch (kind) {
      case 'date':
        return date.toISOString().slice(0, 10)
      case 'time':
        return date.toISOString().slice(11, 19)
      case 'datetime':
        return date.toISOString().slice(0, 19).replace('T', ' ')
    }
  }
  switch (kind) {
    case 'date':
      return new Intl.DateTimeFormat('en-US', {
        timeZone: 'UTC',
        year: 'numeric',
        month: '2-digit',
        day: '2-digit',
      }).format(date)
    case 'time':
      return new Intl.DateTimeFormat('en-US', {
        timeZone: 'UTC',
        hour: 'numeric',
        minute: '2-digit',
      }).format(date)
    case 'datetime':
      return new Intl.DateTimeFormat('en-US', {
        timeZone: 'UTC',
        year: 'numeric',
        month: '2-digit',
        day: '2-digit',
        hour: 'numeric',
        minute: '2-digit',
      }).format(date)
  }
}

function excelSerialToDate(serial: number): Date {
  return new Date((serial - EXCEL_EPOCH_OFFSET) * DAY_MS)
}

function normalizeCurrency(currency: string | undefined): string {
  const normalized = (currency ?? 'USD').trim().toUpperCase()
  return /^[A-Z]{3}$/.test(normalized) ? normalized : 'USD'
}

function normalizeDateStyle(style: CellDateStyle | undefined): CellDateStyle {
  return style === 'iso' ? 'iso' : 'short'
}

function clampDecimals(value: number): number {
  return Math.max(0, Math.min(8, Math.trunc(value)))
}

function toInteger(value: string | undefined, fallback: number): number {
  const parsed = Number.parseInt(value ?? '', 10)
  return Number.isFinite(parsed) ? parsed : fallback
}

function toNumberFormatKind(value: string | undefined): CellNumberFormatKind {
  switch (value) {
    case undefined:
      return 'general'
    case 'number':
    case 'currency':
    case 'accounting':
    case 'percent':
    case 'date':
    case 'time':
    case 'datetime':
    case 'text':
      return value
    default:
      return 'general'
  }
}
