import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import { createBlockedBuiltinMap, textPlaceholderBuiltinNames } from './placeholder.js'
import { createTextCoreBuiltins } from './text-core-builtins.js'
import { createTextFormatBuiltins } from './text-format-builtins.js'
import { createTextSearchBuiltins } from './text-search-builtins.js'
import type { EvaluationResult } from '../runtime-values.js'

export type TextBuiltin = (...args: CellValue[]) => EvaluationResult

function error(code: ErrorCode): CellValue {
  return { tag: ValueTag.Error, code }
}

function stringResult(value: string): CellValue {
  return { tag: ValueTag.String, value, stringId: 0 }
}

function numberResult(value: number): CellValue {
  return { tag: ValueTag.Number, value }
}

function booleanResult(value: boolean): CellValue {
  return { tag: ValueTag.Boolean, value }
}

function firstError(args: readonly (CellValue | undefined)[]): CellValue | undefined {
  return args.find((arg) => arg?.tag === ValueTag.Error)
}

function coerceText(value: CellValue): string {
  switch (value.tag) {
    case ValueTag.Empty:
      return ''
    case ValueTag.Number:
      return String(value.value)
    case ValueTag.Boolean:
      return value.value ? 'TRUE' : 'FALSE'
    case ValueTag.String:
      return value.value
    case ValueTag.Error:
      return ''
  }
}

const utf8Encoder = new TextEncoder()
const utf8Decoder = new TextDecoder()

function utf8Bytes(value: string): Uint8Array {
  return utf8Encoder.encode(value)
}

function utf8Text(bytes: Uint8Array): string {
  return utf8Decoder.decode(bytes)
}

function findSubBytes(haystack: Uint8Array, needle: Uint8Array, start: number): number {
  if (needle.length === 0) {
    return Math.max(0, Math.min(start, haystack.length))
  }

  for (let index = start; index + needle.length <= haystack.length; index += 1) {
    let match = true
    for (let offset = 0; offset < needle.length; offset += 1) {
      if (haystack[index + offset] !== needle[offset]) {
        match = false
        break
      }
    }
    if (match) {
      return index
    }
  }
  return -1
}

function leftBytes(text: string, byteCount: number): string {
  const bytes = utf8Bytes(text)
  const normalizedCount = Math.max(0, Math.min(byteCount, bytes.length))
  return utf8Text(bytes.slice(0, normalizedCount))
}

function rightBytes(text: string, byteCount: number): string {
  const bytes = utf8Bytes(text)
  const normalizedCount = Math.max(0, Math.min(byteCount, bytes.length))
  return utf8Text(bytes.slice(bytes.length - normalizedCount))
}

function midBytes(text: string, start: number, byteCount: number): string {
  const bytes = utf8Bytes(text)
  if (byteCount <= 0) {
    return ''
  }

  const zeroBasedStart = Math.max(0, start - 1)
  const zeroBasedEnd = Math.min(bytes.length, zeroBasedStart + byteCount)
  if (zeroBasedStart >= bytes.length) {
    return ''
  }
  return utf8Text(bytes.slice(zeroBasedStart, zeroBasedEnd))
}

function replaceBytes(text: string, start: number, byteCount: number, replacement: string): string {
  const bytes = utf8Bytes(text)
  const replacementBytes = utf8Bytes(replacement)
  const zeroBasedStart = Math.max(0, start - 1)
  if (zeroBasedStart >= bytes.length) {
    return text
  }
  const zeroBasedEnd = Math.min(bytes.length, zeroBasedStart + Math.max(0, byteCount))
  return utf8Text(new Uint8Array([...bytes.slice(0, zeroBasedStart), ...replacementBytes, ...bytes.slice(zeroBasedEnd)]))
}

function bytePositionToCharPosition(text: string, startByte: number): number {
  if (startByte <= 1) {
    return 1
  }
  return utf8Text(utf8Bytes(text).slice(0, startByte - 1)).length + 1
}

function charPositionToBytePosition(text: string, charPosition: number): number {
  return utf8Bytes(text.slice(0, Math.max(0, charPosition - 1))).length + 1
}

function coerceNumber(value: CellValue): number | undefined {
  switch (value.tag) {
    case ValueTag.Number:
      return value.value
    case ValueTag.Boolean:
      return value.value ? 1 : 0
    case ValueTag.Empty:
      return 0
    case ValueTag.String: {
      const trimmed = value.value.trim()
      if (trimmed === '') {
        return 0
      }
      const parsed = Number(trimmed)
      return Number.isFinite(parsed) ? parsed : undefined
    }
    case ValueTag.Error:
      return undefined
  }
}

function coerceBoolean(value: CellValue, fallback: boolean): boolean | CellValue {
  if (value.tag === ValueTag.Boolean) {
    return value.value
  }
  if (value.tag === ValueTag.Empty) {
    return fallback
  }
  const numeric = coerceNumber(value)
  return numeric === undefined ? error(ErrorCode.Value) : numeric !== 0
}

function coerceInteger(value: CellValue | undefined, defaultValue: number): number | CellValue {
  if (value === undefined) {
    return defaultValue
  }
  const numeric = coerceNumber(value)
  if (numeric === undefined || !Number.isInteger(numeric)) {
    return error(ErrorCode.Value)
  }
  return numeric
}

function coercePositiveStart(value: CellValue | undefined, defaultValue: number): number | CellValue {
  if (value === undefined) {
    return defaultValue
  }
  const numeric = coerceNumber(value)
  if (numeric === undefined) {
    return error(ErrorCode.Value)
  }
  const truncated = Math.trunc(numeric)
  return truncated >= 1 ? truncated : error(ErrorCode.Value)
}

function coerceLength(value: CellValue | undefined, defaultValue: number): number | CellValue {
  if (value === undefined) {
    return defaultValue
  }
  const numeric = coerceNumber(value)
  if (numeric === undefined) {
    return error(ErrorCode.Value)
  }
  const truncated = Math.trunc(numeric)
  return truncated >= 0 ? truncated : error(ErrorCode.Value)
}

function isErrorValue(value: number | CellValue): value is CellValue {
  return typeof value !== 'number'
}

function coerceNonNegativeInt(value: CellValue | undefined, defaultValue: number): number | CellValue {
  if (value === undefined) {
    return defaultValue
  }
  const numeric = coerceNumber(value)
  if (numeric === undefined) {
    return error(ErrorCode.Value)
  }
  const truncated = Math.trunc(numeric)
  return truncated >= 0 ? truncated : error(ErrorCode.Value)
}

function replaceSingle(text: string, start: number, count: number, replacement: string): string {
  const index = start - 1
  if (index >= text.length) {
    return text
  }
  return text.slice(0, index) + replacement + text.slice(index + count)
}

function substituteText(text: string, oldText: string, newText: string, instance?: number): string {
  if (oldText === '') {
    return text
  }
  if (instance === undefined) {
    if (!text.includes(oldText)) {
      return text
    }
    return text.split(oldText).join(newText)
  }

  let occurrence = 0
  let searchIndex = 0
  while (searchIndex <= text.length) {
    const foundAt = text.indexOf(oldText, searchIndex)
    if (foundAt === -1) {
      return text
    }
    occurrence += 1
    if (occurrence === instance) {
      return text.slice(0, foundAt) + newText + text.slice(foundAt + oldText.length)
    }
    searchIndex = foundAt + oldText.length
  }
  return text
}

function createReplaceBuiltin(): TextBuiltin {
  return (...args) => {
    const existingError = firstError(args)
    if (existingError) {
      return existingError
    }
    const [textValue, startValue, countValue, replacementValue] = args
    if (textValue === undefined || startValue === undefined || countValue === undefined || replacementValue === undefined) {
      return error(ErrorCode.Value)
    }
    const start = coercePositiveStart(startValue, 1)
    if (isErrorValue(start)) {
      return start
    }
    const count = coerceLength(countValue, 0)
    if (isErrorValue(count)) {
      return count
    }
    const replacement = coerceText(replacementValue)
    const replaced = replaceSingle(coerceText(textValue), start, count, replacement)
    return stringResult(replaced)
  }
}

function createSubstituteBuiltin(): TextBuiltin {
  return (...args) => {
    const existingError = firstError(args)
    if (existingError) {
      return existingError
    }
    const [textValue, oldValue, newValue, instanceValue] = args
    if (textValue === undefined || oldValue === undefined || newValue === undefined) {
      return error(ErrorCode.Value)
    }
    const text = coerceText(textValue)
    const oldText = coerceText(oldValue)
    if (oldText === '') {
      return error(ErrorCode.Value)
    }
    const newText = coerceText(newValue)
    if (instanceValue === undefined) {
      return stringResult(substituteText(text, oldText, newText))
    }

    const instance = coercePositiveStart(instanceValue, 1)
    if (isErrorValue(instance)) {
      return instance
    }
    return stringResult(substituteText(text, oldText, newText, instance))
  }
}

function createReptBuiltin(): TextBuiltin {
  return (...args) => {
    const existingError = firstError(args)
    if (existingError) {
      return existingError
    }
    const [textValue, countValue] = args
    if (textValue === undefined || countValue === undefined) {
      return error(ErrorCode.Value)
    }
    const count = coerceNonNegativeInt(countValue, 0)
    if (isErrorValue(count)) {
      return count
    }
    const text = coerceText(textValue)
    let repeated = ''
    for (let index = 0; index < count; index += 1) {
      repeated += text
    }
    return stringResult(repeated)
  }
}

function charCodeFromArgument(value: CellValue | undefined): number | CellValue {
  if (value === undefined) {
    return error(ErrorCode.Value)
  }
  const code = coerceNumber(value)
  if (code === undefined) {
    return error(ErrorCode.Value)
  }
  const integerCode = Math.trunc(code)
  if (!Number.isFinite(integerCode) || integerCode < 1 || integerCode > 255) {
    return error(ErrorCode.Value)
  }
  return integerCode
}

const textPlaceholderBuiltins = createBlockedBuiltinMap(textPlaceholderBuiltinNames)
const textCoreBuiltins = createTextCoreBuiltins({
  error,
  stringResult,
  booleanResult,
  firstError,
  coerceText,
  coerceNumber,
})
const textFormatBuiltins = createTextFormatBuiltins({
  error,
  stringResult,
  numberResult,
  firstError,
  coerceText,
  coerceNumber,
  coerceInteger,
  isErrorValue,
})
const textSearchBuiltins = createTextSearchBuiltins({
  error,
  stringResult,
  numberResult,
  booleanResult,
  firstError,
  coerceText,
  coerceNumber,
  coerceBoolean,
  coerceInteger,
  coercePositiveStart,
  isErrorValue,
  utf8Bytes,
  findSubBytes,
  bytePositionToCharPosition,
  charPositionToBytePosition,
})

export const textBuiltins: Record<string, TextBuiltin> = {
  LEN: (...args) => {
    const existingError = firstError(args)
    if (existingError) {
      return existingError
    }
    const [value] = args
    if (value === undefined) {
      return error(ErrorCode.Value)
    }
    return numberResult(coerceText(value).length)
  },
  LENB: (...args) => {
    const existingError = firstError(args)
    if (existingError) {
      return existingError
    }
    const [value] = args
    if (value === undefined) {
      return error(ErrorCode.Value)
    }
    return numberResult(utf8Bytes(coerceText(value)).length)
  },
  CHAR: (...args) => {
    const [codeValue] = args
    const codePoint = charCodeFromArgument(codeValue)
    if (isErrorValue(codePoint)) {
      return codePoint
    }
    return stringResult(String.fromCodePoint(codePoint))
  },
  CODE: (...args) => {
    const [textValue] = args
    if (textValue === undefined) {
      return error(ErrorCode.Value)
    }
    const text = coerceText(textValue)
    if (text.length === 0) {
      return error(ErrorCode.Value)
    }
    const codePoint = text.codePointAt(0)
    return codePoint === undefined ? error(ErrorCode.Value) : numberResult(codePoint)
  },
  UNICODE: (...args) => {
    const [textValue] = args
    if (textValue === undefined) {
      return error(ErrorCode.Value)
    }
    const text = coerceText(textValue)
    if (text.length === 0) {
      return error(ErrorCode.Value)
    }
    const codePoint = text.codePointAt(0)
    return codePoint === undefined ? error(ErrorCode.Value) : numberResult(codePoint)
  },
  UNICHAR: (...args) => {
    const [codeValue] = args
    if (codeValue === undefined) {
      return error(ErrorCode.Value)
    }
    const code = coerceNumber(codeValue)
    if (code === undefined) {
      return error(ErrorCode.Value)
    }
    const integerCode = Math.trunc(code)
    if (!Number.isFinite(integerCode) || integerCode < 0 || integerCode > 0x10ffff) {
      return error(ErrorCode.Value)
    }
    return stringResult(String.fromCodePoint(integerCode))
  },
  ...textCoreBuiltins,
  ...textFormatBuiltins,
  ...textSearchBuiltins,
  LEFT: (...args) => {
    const existingError = firstError(args)
    if (existingError) {
      return existingError
    }
    const [textValue, countValue] = args
    if (textValue === undefined) {
      return error(ErrorCode.Value)
    }
    const count = coerceLength(countValue, 1)
    if (isErrorValue(count)) {
      return count
    }
    return stringResult(coerceText(textValue).slice(0, count))
  },
  RIGHT: (...args) => {
    const existingError = firstError(args)
    if (existingError) {
      return existingError
    }
    const [textValue, countValue] = args
    if (textValue === undefined) {
      return error(ErrorCode.Value)
    }
    const count = coerceLength(countValue, 1)
    if (isErrorValue(count)) {
      return count
    }
    const text = coerceText(textValue)
    return stringResult(count === 0 ? '' : text.slice(-count))
  },
  MID: (...args) => {
    const existingError = firstError(args)
    if (existingError) {
      return existingError
    }
    const [textValue, startValue, countValue] = args
    if (textValue === undefined || startValue === undefined || countValue === undefined) {
      return error(ErrorCode.Value)
    }
    const start = coercePositiveStart(startValue, 1)
    if (isErrorValue(start)) {
      return start
    }
    const count = coerceLength(countValue, 0)
    if (isErrorValue(count)) {
      return count
    }
    const text = coerceText(textValue)
    return stringResult(text.slice(start - 1, start - 1 + count))
  },
  ENCODEURL: (...args) => {
    const existingError = firstError(args)
    if (existingError) {
      return existingError
    }
    const [value] = args
    if (value === undefined) {
      return error(ErrorCode.Value)
    }
    return stringResult(encodeURI(coerceText(value)))
  },
  LEFTB: (...args) => {
    const existingError = firstError(args)
    if (existingError) {
      return existingError
    }
    const [textValue, countValue] = args
    if (textValue === undefined) {
      return error(ErrorCode.Value)
    }
    const count = coerceLength(countValue, 1)
    if (isErrorValue(count)) {
      return count
    }
    return stringResult(leftBytes(coerceText(textValue), count))
  },
  MIDB: (...args) => {
    const existingError = firstError(args)
    if (existingError) {
      return existingError
    }
    const [textValue, startValue, countValue] = args
    if (textValue === undefined || startValue === undefined || countValue === undefined) {
      return error(ErrorCode.Value)
    }
    const start = coercePositiveStart(startValue, 1)
    if (isErrorValue(start)) {
      return start
    }
    const count = coerceLength(countValue, 0)
    if (isErrorValue(count)) {
      return count
    }
    return stringResult(midBytes(coerceText(textValue), start, count))
  },
  RIGHTB: (...args) => {
    const existingError = firstError(args)
    if (existingError) {
      return existingError
    }
    const [textValue, countValue] = args
    if (textValue === undefined) {
      return error(ErrorCode.Value)
    }
    const count = coerceLength(countValue, 1)
    if (isErrorValue(count)) {
      return count
    }
    return stringResult(rightBytes(coerceText(textValue), count))
  },
  REPLACE: createReplaceBuiltin(),
  REPLACEB: (...args) => {
    const existingError = firstError(args)
    if (existingError) {
      return existingError
    }
    const [textValue, startValue, countValue, replacementValue] = args
    if (textValue === undefined || startValue === undefined || countValue === undefined || replacementValue === undefined) {
      return error(ErrorCode.Value)
    }
    const start = coercePositiveStart(startValue, 1)
    if (isErrorValue(start)) {
      return start
    }
    const count = coerceLength(countValue, 0)
    if (isErrorValue(count)) {
      return count
    }
    return stringResult(replaceBytes(coerceText(textValue), start, count, coerceText(replacementValue)))
  },
  SUBSTITUTE: createSubstituteBuiltin(),
  REPT: createReptBuiltin(),
  ...textPlaceholderBuiltins,
}

export function getTextBuiltin(name: string): TextBuiltin | undefined {
  return textBuiltins[name.toUpperCase()]
}
