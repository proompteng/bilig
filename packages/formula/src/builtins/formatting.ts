import { roundToDigits } from './numeric.js'

export function toColumnLabel(column: number): string | undefined {
  if (!Number.isInteger(column) || column < 1) {
    return undefined
  }
  let current = column
  let label = ''
  while (current > 0) {
    const offset = (current - 1) % 26
    label = String.fromCharCode(65 + offset) + label
    current = Math.floor((current - 1) / 26)
  }
  return label
}

function formatThousands(value: string): string {
  return value.replace(/\B(?=(\d{3})+(?!\d))/g, ',')
}

export function formatFixed(value: number, decimals: number, includeThousands: boolean): string {
  if (!Number.isFinite(value) || !Number.isInteger(decimals)) {
    return ''
  }
  const rounded = roundToDigits(value, decimals)
  const sign = rounded < 0 ? '-' : ''
  const unsigned = Math.abs(rounded)
  const fixedDecimals = decimals >= 0 ? decimals : 0
  const fixed = unsigned.toFixed(fixedDecimals)
  const [integerPart = '0', fractionPart] = fixed.split('.')
  const normalizedInteger = includeThousands ? formatThousands(integerPart) : integerPart
  return `${sign}${normalizedInteger}${fractionPart === undefined ? '' : `.${fractionPart}`}`
}

export function countLeadingZeros(value: number): number {
  if (value <= 0) {
    return 1
  }
  return Math.max(1, Math.ceil(Math.log10(value)))
}

export function isValidDollarFraction(fraction: number): boolean {
  if (!Number.isInteger(fraction) || fraction <= 0) {
    return false
  }
  if (fraction === 1) {
    return true
  }
  return Number.isInteger(Math.log2(fraction))
}

export function parseDollarDecimal(value: number): {
  integerPart: number
  fractionalNumerator: number
} {
  const absolute = Math.abs(value)
  const parts = absolute.toString().split('.')
  const integerPart = Number(parts[0] ?? 0)
  const fractionalText = parts[1] ?? '0'
  return { integerPart, fractionalNumerator: Number(fractionalText) }
}
