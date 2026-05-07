import { ErrorCode, ValueTag, formatGeneralNumberValue, type CellValue } from '@bilig/protocol'

export function valueError(): CellValue {
  return { tag: ValueTag.Error, code: ErrorCode.Value }
}

export function numberResult(value: number): CellValue {
  return { tag: ValueTag.Number, value }
}

export function firstError(args: readonly CellValue[]): CellValue | undefined {
  return args.find((arg) => arg.tag === ValueTag.Error)
}

export function coerceNumber(value: CellValue): number | undefined {
  switch (value.tag) {
    case ValueTag.Number:
      return Number.isFinite(value.value) ? value.value : undefined
    case ValueTag.Boolean:
      return value.value ? 1 : 0
    case ValueTag.Empty:
      return 0
    case ValueTag.String:
    case ValueTag.Error:
      return undefined
    default:
      return undefined
  }
}

export function coerceText(value: CellValue): string | undefined {
  switch (value.tag) {
    case ValueTag.String:
      return value.value
    case ValueTag.Number:
      return formatGeneralNumberValue(value.value)
    case ValueTag.Boolean:
      return value.value ? 'TRUE' : 'FALSE'
    case ValueTag.Empty:
      return ''
    case ValueTag.Error:
      return undefined
  }
}

export function integerValue(value: CellValue | undefined, fallback?: number): number | undefined {
  if (value === undefined) {
    return fallback
  }
  const numeric = coerceNumber(value)
  if (numeric === undefined || !Number.isFinite(numeric)) {
    return undefined
  }
  return Math.trunc(numeric)
}

export function truncArg(value: CellValue): number | CellValue {
  if (value.tag === ValueTag.Error) {
    return value
  }
  const coerced = coerceNumber(value)
  if (coerced === undefined) {
    return valueError()
  }
  return Math.trunc(coerced)
}

export function isErrorValue(value: Set<number> | CellValue): value is CellValue {
  return !('size' in value)
}
