import { ValueTag, type CellValue } from '@bilig/protocol'
import type { EvaluationResult } from '../runtime-values.js'

type Builtin = (...args: CellValue[]) => EvaluationResult

interface RadixBuiltinHelpers {
  toNumber(this: void, value: CellValue): number | undefined
  integerValue(this: void, value: CellValue | undefined, fallback?: number): number | undefined
  nonNegativeIntegerValue(this: void, value: CellValue | undefined, fallback?: number): number | undefined
  valueError(this: void): CellValue
  numberResult(this: void, value: number): CellValue
}

export function createRadixBuiltins({
  toNumber,
  integerValue,
  nonNegativeIntegerValue,
  valueError,
  numberResult,
}: RadixBuiltinHelpers): Record<string, Builtin> {
  return {
    BASE: (numberArg, radixArg, minLengthArg) => baseString(numberArg, radixArg, minLengthArg),
    DECIMAL: (textArg, radixArg) => decimalValue(textArg, radixArg),
    BIN2DEC: (valueArg) => {
      const numeric = parseSignedRadixText(valueArg, 2, 10, toNumber)
      return numeric === undefined ? valueError() : numberResult(numeric)
    },
    BIN2HEX: (valueArg, placesArg) => convertSignedRadixToRadix(valueArg, 2, 16, 10, 10, -549755813888, 549755813887, placesArg),
    BIN2OCT: (valueArg, placesArg) => convertSignedRadixToRadix(valueArg, 2, 8, 10, 10, -536870912, 536870911, placesArg),
    DEC2BIN: (valueArg, placesArg) => {
      const numeric = integerValue(valueArg)
      const places = toPositivePlaces(placesArg)
      if (numeric === undefined || places === undefined) {
        return valueError()
      }
      return formatSignedRadixValue(numeric, 2, places, 10, -512, 511)
    },
    DEC2HEX: (valueArg, placesArg) => {
      const numeric = integerValue(valueArg)
      const places = toPositivePlaces(placesArg)
      if (numeric === undefined || places === undefined) {
        return valueError()
      }
      return formatSignedRadixValue(numeric, 16, places, 10, -549755813888, 549755813887)
    },
    DEC2OCT: (valueArg, placesArg) => {
      const numeric = integerValue(valueArg)
      const places = toPositivePlaces(placesArg)
      if (numeric === undefined || places === undefined) {
        return valueError()
      }
      return formatSignedRadixValue(numeric, 8, places, 10, -536870912, 536870911)
    },
    HEX2BIN: (valueArg, placesArg) => convertSignedRadixToRadix(valueArg, 16, 2, 10, 10, -512, 511, placesArg),
    HEX2DEC: (valueArg) => {
      const numeric = parseSignedRadixText(valueArg, 16, 10, toNumber)
      return numeric === undefined ? valueError() : numberResult(numeric)
    },
    HEX2OCT: (valueArg, placesArg) => convertSignedRadixToRadix(valueArg, 16, 8, 10, 10, -536870912, 536870911, placesArg),
    OCT2BIN: (valueArg, placesArg) => convertSignedRadixToRadix(valueArg, 8, 2, 10, 10, -512, 511, placesArg),
    OCT2DEC: (valueArg) => {
      const numeric = parseSignedRadixText(valueArg, 8, 10, toNumber)
      return numeric === undefined ? valueError() : numberResult(numeric)
    },
    OCT2HEX: (valueArg, placesArg) => convertSignedRadixToRadix(valueArg, 8, 16, 10, 10, -549755813888, 549755813887, placesArg),
    ROMAN: (value) => {
      const roman = romanValue(toNumber(value) ?? Number.NaN)
      return roman === undefined ? valueError() : { tag: ValueTag.String, value: roman, stringId: 0 }
    },
    ARABIC: (value) => {
      if (value.tag !== ValueTag.String) {
        return valueError()
      }
      const numeric = arabicValue(value.value)
      return numeric === undefined ? valueError() : numberResult(numeric)
    },
  }

  function baseString(numberArg: CellValue, radixArg: CellValue, minLengthArg?: CellValue): CellValue {
    const numberValue = integerValue(numberArg)
    const radixValue = integerValue(radixArg)
    const minLengthValue = nonNegativeIntegerValue(minLengthArg, 0)
    if (
      numberValue === undefined ||
      numberValue < 0 ||
      radixValue === undefined ||
      radixValue < 2 ||
      radixValue > 36 ||
      minLengthValue === undefined
    ) {
      return valueError()
    }
    return {
      tag: ValueTag.String,
      value: numberValue.toString(radixValue).toUpperCase().padStart(minLengthValue, '0'),
      stringId: 0,
    }
  }

  function decimalValue(textArg: CellValue, radixArg: CellValue): CellValue {
    if (textArg.tag === ValueTag.Error) {
      return textArg
    }
    const radixValue = integerValue(radixArg)
    if (radixValue === undefined || radixValue < 2 || radixValue > 36) {
      return valueError()
    }
    const raw = textArg.tag === ValueTag.String ? textArg.value.trim() : String(Math.trunc(toNumber(textArg) ?? Number.NaN))
    if (raw === '' || raw === 'NaN' || !isValidBaseDigits(raw, radixValue)) {
      return valueError()
    }
    return numberResult(Number.parseInt(raw, radixValue))
  }

  function toPositivePlaces(value: CellValue | undefined): number | undefined {
    return nonNegativeIntegerValue(value, 0)
  }

  function convertSignedRadixToRadix(
    textArg: CellValue,
    fromRadix: 2 | 8 | 16,
    toRadix: 2 | 8 | 16,
    inputWidth: number,
    negativeWidth: number,
    min: number,
    max: number,
    placesArg?: CellValue,
  ): CellValue {
    const numeric = parseSignedRadixText(textArg, fromRadix, inputWidth, toNumber)
    const places = toPositivePlaces(placesArg)
    if (numeric === undefined || places === undefined) {
      return valueError()
    }
    return formatSignedRadixValue(numeric, toRadix, places, negativeWidth, min, max)
  }

  function formatSignedRadixValue(
    numeric: number,
    radix: 2 | 8 | 16,
    minLength: number,
    negativeWidth: number,
    min: number,
    max: number,
  ): CellValue {
    if (!Number.isSafeInteger(numeric) || numeric < min || numeric > max) {
      return valueError()
    }
    if (numeric < 0) {
      const encoded = numeric + radix ** negativeWidth
      return {
        tag: ValueTag.String,
        value: encoded.toString(radix).toUpperCase().padStart(negativeWidth, '0'),
        stringId: 0,
      }
    }
    const raw = numeric.toString(radix).toUpperCase()
    if (minLength < raw.length) {
      return valueError()
    }
    return {
      tag: ValueTag.String,
      value: raw.padStart(minLength, '0'),
      stringId: 0,
    }
  }
}

function romanValue(numberValue: number): string | undefined {
  const number = Math.trunc(numberValue)
  if (!Number.isFinite(numberValue) || number < 1 || number > 3999) {
    return undefined
  }
  const numerals: Array<[number, string]> = [
    [1000, 'M'],
    [900, 'CM'],
    [500, 'D'],
    [400, 'CD'],
    [100, 'C'],
    [90, 'XC'],
    [50, 'L'],
    [40, 'XL'],
    [10, 'X'],
    [9, 'IX'],
    [5, 'V'],
    [4, 'IV'],
    [1, 'I'],
  ]
  let remaining = number
  let result = ''
  for (const [value, numeral] of numerals) {
    while (remaining >= value) {
      result += numeral
      remaining -= value
    }
  }
  return result
}

function arabicValue(text: string): number | undefined {
  const numerals = new Map<string, number>([
    ['I', 1],
    ['V', 5],
    ['X', 10],
    ['L', 50],
    ['C', 100],
    ['D', 500],
    ['M', 1000],
  ])
  const upper = text.trim().toUpperCase()
  if (!/^[IVXLCDM]+$/.test(upper)) {
    return undefined
  }
  let total = 0
  let index = 0
  while (index < upper.length) {
    const current = numerals.get(upper[index] ?? '')
    const next = numerals.get(upper[index + 1] ?? '')
    if (current === undefined) {
      return undefined
    }
    if (next !== undefined && current < next) {
      total += next - current
      index += 2
      continue
    }
    total += current
    index += 1
  }
  return romanValue(total) === upper ? total : undefined
}

function isValidBaseDigits(raw: string, radix: number): boolean {
  const upper = raw.toUpperCase()
  const digits = new Set('0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ'.slice(0, radix))
  for (const char of upper) {
    if (!digits.has(char)) {
      return false
    }
  }
  return true
}

function parseSignedRadixText(
  textArg: CellValue,
  radix: 2 | 8 | 16,
  width: number,
  toNumber: RadixBuiltinHelpers['toNumber'],
): number | undefined {
  if (textArg.tag === ValueTag.Error) {
    return undefined
  }
  const raw =
    textArg.tag === ValueTag.String ? textArg.value.trim().toUpperCase() : String(Math.trunc(toNumber(textArg) ?? Number.NaN)).toUpperCase()
  if (raw === '' || raw === 'NAN' || raw.length > width || !isValidBaseDigits(raw, radix)) {
    return undefined
  }
  const parsed = Number.parseInt(raw, radix)
  if (!Number.isFinite(parsed)) {
    return undefined
  }
  const fullRange = radix ** width
  const negativeThreshold = fullRange / 2
  return raw.length === width && parsed >= negativeThreshold ? parsed - fullRange : parsed
}
