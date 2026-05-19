import type { WorkbookAgentWriteCellInput } from '@bilig/agent-api'
import { type CellNumberFormatInput, type CellNumberFormatPreset, normalizeCellNumberFormatPreset } from '@bilig/protocol'

type WorkbookAgentToolNumberFormatInput =
  | string
  | {
      readonly kind: CellNumberFormatPreset['kind']
      readonly currency?: string | undefined
      readonly decimals?: number | undefined
      readonly useGrouping?: boolean | undefined
      readonly negativeStyle?: 'minus' | 'parentheses' | undefined
      readonly zeroStyle?: 'zero' | 'dash' | undefined
      readonly dateStyle?: 'short' | 'iso' | undefined
    }

function normalizeFormula(formula: string): string {
  return formula.startsWith('=') ? formula.slice(1) : formula
}

function isNumericText(value: string): boolean {
  const trimmed = value.trim()
  if (trimmed.length === 0 || trimmed !== value) {
    return false
  }
  if (!/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/.test(trimmed)) {
    return false
  }
  const unsigned = trimmed.replace(/^[+-]/, '')
  if (/^0\d/.test(unsigned)) {
    return false
  }
  const parsed = Number(trimmed)
  return Number.isFinite(parsed)
}

function coerceNumericText(value: string): number | string {
  return isNumericText(value) ? Number(value) : value
}

function excelDateSerialFromIsoDate(value: string): number {
  const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(value)
  if (!match) {
    throw new Error(`Date value must use YYYY-MM-DD format, received ${value}`)
  }
  const year = Number(match[1])
  const month = Number(match[2])
  const day = Number(match[3])
  const timestamp = Date.UTC(year, month - 1, day)
  const date = new Date(timestamp)
  if (date.getUTCFullYear() !== year || date.getUTCMonth() !== month - 1 || date.getUTCDate() !== day) {
    throw new Error(`Invalid date value ${value}`)
  }
  const excelEpoch = Date.UTC(1899, 11, 30)
  return Math.trunc((timestamp - excelEpoch) / 86_400_000)
}

function normalizeBooleanInput(value: string | boolean): boolean {
  if (typeof value === 'boolean') {
    return value
  }
  const normalized = value.trim().toLowerCase()
  if (normalized === 'true') {
    return true
  }
  if (normalized === 'false') {
    return false
  }
  throw new Error(`Boolean value must be true or false, received ${value}`)
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

export function normalizeWorkbookAgentToolNumberFormatInput(input: WorkbookAgentToolNumberFormatInput): CellNumberFormatInput {
  if (typeof input === 'string') {
    return input
  }

  const preset: CellNumberFormatPreset = {
    kind: input.kind,
  }
  if (typeof input.currency === 'string') {
    preset.currency = input.currency
  }
  if (typeof input.decimals === 'number') {
    preset.decimals = input.decimals
  }
  if (typeof input.useGrouping === 'boolean') {
    preset.useGrouping = input.useGrouping
  }
  if (input.negativeStyle === 'minus' || input.negativeStyle === 'parentheses') {
    preset.negativeStyle = input.negativeStyle
  }
  if (input.zeroStyle === 'zero' || input.zeroStyle === 'dash') {
    preset.zeroStyle = input.zeroStyle
  }
  if (input.dateStyle === 'short' || input.dateStyle === 'iso') {
    preset.dateStyle = input.dateStyle
  }

  return normalizeCellNumberFormatPreset(preset)
}

export function normalizeWorkbookAgentWriteCellInput(cellInput: unknown): WorkbookAgentWriteCellInput {
  if (cellInput === null || typeof cellInput === 'boolean') {
    return cellInput
  }
  if (typeof cellInput === 'number') {
    if (!Number.isFinite(cellInput)) {
      throw new Error(`Number cell input must be finite, received ${String(cellInput)}`)
    }
    return cellInput
  }
  if (typeof cellInput === 'string') {
    return cellInput.startsWith('=')
      ? {
          formula: `=${normalizeFormula(cellInput)}`,
        }
      : coerceNumericText(cellInput)
  }
  if (!isRecord(cellInput)) {
    throw new Error('Unsupported write_range cell input')
  }
  if (cellInput['type'] === 'blank') {
    return null
  }
  if (cellInput['type'] === 'formula') {
    if (typeof cellInput['formula'] !== 'string') {
      throw new Error('Typed formula cell requires a formula string')
    }
    return {
      formula: `=${normalizeFormula(cellInput['formula'])}`,
    }
  }
  if (cellInput['type'] === 'text') {
    if (typeof cellInput['value'] !== 'string') {
      throw new Error('Typed text cell requires a string value')
    }
    return cellInput['value']
  }
  if (cellInput['type'] === 'number') {
    const value = cellInput['value']
    if (typeof value === 'number') {
      if (!Number.isFinite(value)) {
        throw new Error(`Typed number cell requires a finite number or numeric string, received ${String(value)}`)
      }
      return value
    }
    if (typeof value === 'string' && isNumericText(value)) {
      return Number(value)
    }
    throw new Error(`Typed number cell requires a finite number or numeric string, received ${String(value)}`)
  }
  if (cellInput['type'] === 'date') {
    const value = cellInput['value']
    if (typeof value === 'number' && Number.isFinite(value)) {
      return value
    }
    if (typeof value === 'string') {
      return excelDateSerialFromIsoDate(value)
    }
    throw new Error('Typed date cell requires an Excel serial number or YYYY-MM-DD string')
  }
  if (cellInput['type'] === 'boolean') {
    const value = cellInput['value']
    if (typeof value === 'string' || typeof value === 'boolean') {
      return normalizeBooleanInput(value)
    }
    throw new Error('Typed boolean cell requires a boolean or true/false string')
  }
  if (typeof cellInput['formula'] === 'string') {
    return {
      formula: `=${normalizeFormula(cellInput['formula'])}`,
    }
  }
  if ('value' in cellInput) {
    const value = cellInput['value']
    if (value === null || typeof value === 'boolean') {
      return value
    }
    if (typeof value === 'number') {
      if (!Number.isFinite(value)) {
        throw new Error(`Number cell input must be finite, received ${String(value)}`)
      }
      return value
    }
    if (typeof value === 'string') {
      return value.startsWith('=')
        ? {
            formula: `=${normalizeFormula(value)}`,
          }
        : coerceNumericText(value)
    }
  }
  throw new Error('Unsupported write_range cell input')
}
