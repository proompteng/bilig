import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import type { TextBuiltin } from './text.js'

interface TextSearchBuiltinDeps {
  error: (code: ErrorCode) => CellValue
  stringResult: (value: string) => CellValue
  numberResult: (value: number) => CellValue
  booleanResult: (value: boolean) => CellValue
  firstError: (args: readonly (CellValue | undefined)[]) => CellValue | undefined
  coerceText: (value: CellValue) => string
  coerceNumber: (value: CellValue) => number | undefined
  coerceBoolean: (value: CellValue, fallback: boolean) => boolean | CellValue
  coerceInteger: (value: CellValue | undefined, defaultValue: number) => number | CellValue
  coercePositiveStart: (value: CellValue | undefined, defaultValue: number) => number | CellValue
  isErrorValue: (value: number | CellValue) => value is CellValue
  utf8Bytes: (value: string) => Uint8Array
  findSubBytes: (haystack: Uint8Array, needle: Uint8Array, start: number) => number
  bytePositionToCharPosition: (text: string, startByte: number) => number
  charPositionToBytePosition: (text: string, charPosition: number) => number
}

function compileRegex(pattern: string, caseSensitivity: number, global = false): RegExp | CellValue {
  try {
    const flags = `${caseSensitivity === 1 ? 'i' : ''}${global ? 'g' : ''}`
    return new RegExp(pattern, flags)
  } catch {
    return { tag: ValueTag.Error, code: ErrorCode.Value }
  }
}

function isRegexError(value: RegExp | CellValue): value is CellValue {
  return !(value instanceof RegExp)
}

function applyReplacementTemplate(template: string, match: string, captures: readonly (string | undefined)[]): string {
  let output = ''

  for (let index = 0; index < template.length; index += 1) {
    const char = template[index]!
    if (char !== '$') {
      output += char
      continue
    }

    const next = template[index + 1]
    if (next === undefined) {
      output += '$'
      continue
    }
    if (next === '$') {
      output += '$'
      index += 1
      continue
    }
    if (next === '&') {
      output += match
      index += 1
      continue
    }
    if (!/\d/.test(next)) {
      output += '$'
      continue
    }

    const secondDigit = template[index + 2]
    const twoDigitIndex = secondDigit !== undefined && /\d/.test(secondDigit) ? Number(`${next}${secondDigit}`) : -1
    if (twoDigitIndex >= 1 && twoDigitIndex <= captures.length) {
      output += captures[twoDigitIndex - 1] ?? ''
      index += 2
      continue
    }

    const captureIndex = Number(next)
    if (captureIndex >= 1 && captureIndex <= captures.length) {
      output += captures[captureIndex - 1] ?? ''
      index += 1
      continue
    }

    output += `$${next}`
    index += 1
  }

  return output
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
}

function indexOfWithMode(text: string, delimiter: string, start: number, matchMode: number): number {
  if (matchMode === 1) {
    return text.toLowerCase().indexOf(delimiter.toLowerCase(), start)
  }
  return text.indexOf(delimiter, start)
}

function lastIndexOfWithMode(text: string, delimiter: string, start: number, matchMode: number): number {
  if (matchMode === 1) {
    return text.toLowerCase().lastIndexOf(delimiter.toLowerCase(), start)
  }
  return text.lastIndexOf(delimiter, start)
}

function hasSearchSyntax(pattern: string): boolean {
  for (let index = 0; index < pattern.length; index += 1) {
    const char = pattern[index]!
    if (char === '~' || char === '*' || char === '?') {
      return true
    }
  }
  return false
}

function buildSearchRegex(pattern: string): RegExp {
  let source = '^'

  for (let index = 0; index < pattern.length; index += 1) {
    const char = pattern[index]!
    if (char === '~') {
      const next = pattern[index + 1]
      if (next === undefined) {
        source += escapeRegExp(char)
      } else {
        source += escapeRegExp(next)
        index += 1
      }
      continue
    }
    if (char === '*') {
      source += '[\\s\\S]*'
      continue
    }
    if (char === '?') {
      source += '[\\s\\S]'
      continue
    }
    source += escapeRegExp(char)
  }

  return new RegExp(source, 'i')
}

function findPosition(
  deps: TextSearchBuiltinDeps,
  needle: string,
  haystack: string,
  start: number,
  caseSensitive: boolean,
  wildcardAware: boolean,
): number | CellValue {
  const startIndex = start - 1

  if (needle === '') {
    return start
  }
  if (startIndex > haystack.length) {
    return deps.error(ErrorCode.Value)
  }

  if (!wildcardAware || !hasSearchSyntax(needle)) {
    const normalizedHaystack = caseSensitive ? haystack : haystack.toLowerCase()
    const normalizedNeedle = caseSensitive ? needle : needle.toLowerCase()
    const found = normalizedHaystack.indexOf(normalizedNeedle, startIndex)
    return found === -1 ? deps.error(ErrorCode.Value) : found + 1
  }

  const regex = buildSearchRegex(needle)
  for (let index = startIndex; index <= haystack.length; index += 1) {
    if (regex.test(haystack.slice(index))) {
      return index + 1
    }
  }
  return deps.error(ErrorCode.Value)
}

export function createTextSearchBuiltins(deps: TextSearchBuiltinDeps): Record<string, TextBuiltin> {
  return {
    FIND: (...args) => {
      const existingError = deps.firstError(args)
      if (existingError) {
        return existingError
      }
      const [findTextValue, withinTextValue, startValue] = args
      if (findTextValue === undefined || withinTextValue === undefined) {
        return deps.error(ErrorCode.Value)
      }
      const start = deps.coercePositiveStart(startValue, 1)
      if (deps.isErrorValue(start)) {
        return start
      }
      const found = findPosition(deps, deps.coerceText(findTextValue), deps.coerceText(withinTextValue), start, true, false)
      return deps.isErrorValue(found) ? found : deps.numberResult(found)
    },
    SEARCH: (...args) => {
      const existingError = deps.firstError(args)
      if (existingError) {
        return existingError
      }
      const [findTextValue, withinTextValue, startValue] = args
      if (findTextValue === undefined || withinTextValue === undefined) {
        return deps.error(ErrorCode.Value)
      }
      const start = deps.coercePositiveStart(startValue, 1)
      if (deps.isErrorValue(start)) {
        return start
      }
      const found = findPosition(deps, deps.coerceText(findTextValue), deps.coerceText(withinTextValue), start, false, true)
      return deps.isErrorValue(found) ? found : deps.numberResult(found)
    },
    SEARCHB: (...args) => {
      const existingError = deps.firstError(args)
      if (existingError) {
        return existingError
      }
      const [findTextValue, withinTextValue, startValue] = args
      if (findTextValue === undefined || withinTextValue === undefined) {
        return deps.error(ErrorCode.Value)
      }
      const text = deps.coerceText(withinTextValue)
      const start = deps.coercePositiveStart(startValue, 1)
      if (deps.isErrorValue(start)) {
        return start
      }
      if (start > deps.utf8Bytes(text).length + 1) {
        return deps.error(ErrorCode.Value)
      }
      const found = findPosition(deps, deps.coerceText(findTextValue), text, deps.bytePositionToCharPosition(text, start), false, true)
      return deps.isErrorValue(found) ? found : deps.numberResult(deps.charPositionToBytePosition(text, found))
    },
    FINDB: (...args) => {
      const existingError = deps.firstError(args)
      if (existingError) {
        return existingError
      }
      const [findTextValue, withinTextValue, startValue] = args
      if (findTextValue === undefined || withinTextValue === undefined) {
        return deps.error(ErrorCode.Value)
      }
      const start = deps.coercePositiveStart(startValue, 1)
      if (deps.isErrorValue(start)) {
        return start
      }
      const findBytes = deps.utf8Bytes(deps.coerceText(findTextValue))
      const withinBytes = deps.utf8Bytes(deps.coerceText(withinTextValue))
      if (start > withinBytes.length + 1) {
        return deps.error(ErrorCode.Value)
      }
      const found = deps.findSubBytes(withinBytes, findBytes, start - 1)
      return found === -1 ? deps.error(ErrorCode.Value) : deps.numberResult(found + 1)
    },
    REGEXTEST: (...args) => {
      const existingError = deps.firstError(args)
      if (existingError) {
        return existingError
      }
      const [textValue, patternValue, caseSensitivityValue] = args
      if (textValue === undefined || patternValue === undefined) {
        return deps.error(ErrorCode.Value)
      }
      const caseSensitivity = deps.coerceInteger(caseSensitivityValue, 0)
      if (deps.isErrorValue(caseSensitivity) || (caseSensitivity !== 0 && caseSensitivity !== 1)) {
        return deps.error(ErrorCode.Value)
      }
      const regex = compileRegex(deps.coerceText(patternValue), caseSensitivity)
      if (isRegexError(regex)) {
        return regex
      }
      return deps.booleanResult(regex.test(deps.coerceText(textValue)))
    },
    REGEXREPLACE: (...args) => {
      const existingError = deps.firstError(args)
      if (existingError) {
        return existingError
      }
      const [textValue, patternValue, replacementValue, occurrenceValue, caseSensitivityValue] = args
      if (textValue === undefined || patternValue === undefined || replacementValue === undefined) {
        return deps.error(ErrorCode.Value)
      }
      const occurrence = deps.coerceInteger(occurrenceValue, 0)
      const caseSensitivity = deps.coerceInteger(caseSensitivityValue, 0)
      if (deps.isErrorValue(occurrence) || deps.isErrorValue(caseSensitivity) || (caseSensitivity !== 0 && caseSensitivity !== 1)) {
        return deps.error(ErrorCode.Value)
      }
      const text = deps.coerceText(textValue)
      const replacement = deps.coerceText(replacementValue)
      const regex = compileRegex(deps.coerceText(patternValue), caseSensitivity, true)
      if (isRegexError(regex)) {
        return regex
      }
      if (occurrence === 0) {
        return deps.stringResult(text.replace(regex, replacement))
      }
      const matches = [...text.matchAll(regex)]
      if (matches.length === 0) {
        return deps.stringResult(text)
      }
      const targetIndex = occurrence > 0 ? occurrence - 1 : matches.length + occurrence
      if (targetIndex < 0 || targetIndex >= matches.length) {
        return deps.stringResult(text)
      }
      let currentIndex = -1
      return deps.stringResult(
        text.replace(regex, (match, ...rest) => {
          currentIndex += 1
          if (currentIndex !== targetIndex) {
            return match
          }
          const captures = rest.slice(0, -2).map((value) => (typeof value === 'string' ? value : undefined))
          return applyReplacementTemplate(replacement, match, captures)
        }),
      )
    },
    REGEXEXTRACT: (...args) => {
      const existingError = deps.firstError(args)
      if (existingError) {
        return existingError
      }
      const [textValue, patternValue, returnModeValue, caseSensitivityValue] = args
      if (textValue === undefined || patternValue === undefined) {
        return deps.error(ErrorCode.Value)
      }
      const returnMode = deps.coerceInteger(returnModeValue, 0)
      const caseSensitivity = deps.coerceInteger(caseSensitivityValue, 0)
      if (
        deps.isErrorValue(returnMode) ||
        deps.isErrorValue(caseSensitivity) ||
        ![0, 1, 2].includes(returnMode) ||
        (caseSensitivity !== 0 && caseSensitivity !== 1)
      ) {
        return deps.error(ErrorCode.Value)
      }

      const text = deps.coerceText(textValue)
      const pattern = deps.coerceText(patternValue)
      if (returnMode === 1) {
        const regex = compileRegex(pattern, caseSensitivity, true)
        if (isRegexError(regex)) {
          return regex
        }
        const matches = [...text.matchAll(regex)].map((entry) => entry[0])
        if (matches.length === 0) {
          return deps.error(ErrorCode.NA)
        }
        return {
          kind: 'array',
          rows: matches.length,
          cols: 1,
          values: matches.map((match) => deps.stringResult(match)),
        }
      }

      const regex = compileRegex(pattern, caseSensitivity, false)
      if (isRegexError(regex)) {
        return regex
      }
      const match = text.match(regex)
      if (!match) {
        return deps.error(ErrorCode.NA)
      }
      if (returnMode === 0) {
        return deps.stringResult(match[0])
      }
      const groups = match.slice(1)
      if (groups.length === 0) {
        return deps.error(ErrorCode.NA)
      }
      return {
        kind: 'array',
        rows: 1,
        cols: groups.length,
        values: groups.map((group) => deps.stringResult(group ?? '')),
      }
    },
    TEXTBEFORE: (...args) => {
      const existingError = deps.firstError(args)
      if (existingError) {
        return existingError
      }
      const [textValue, delimiterValue, instanceValue, matchModeValue, matchEndValue, ifNotFoundValue] = args
      if (textValue === undefined || delimiterValue === undefined) {
        return deps.error(ErrorCode.Value)
      }

      const text = deps.coerceText(textValue)
      const delimiter = deps.coerceText(delimiterValue)
      if (delimiter === '') {
        return deps.error(ErrorCode.Value)
      }

      const instanceNumber = instanceValue === undefined ? 1 : deps.coerceNumber(instanceValue)
      const matchMode = matchModeValue === undefined ? 0 : deps.coerceNumber(matchModeValue)
      const matchEndNumber = matchEndValue === undefined ? 0 : deps.coerceNumber(matchEndValue)
      if (
        instanceNumber === undefined ||
        matchMode === undefined ||
        matchEndNumber === undefined ||
        !Number.isInteger(instanceNumber) ||
        instanceNumber === 0 ||
        !Number.isInteger(matchMode) ||
        (matchMode !== 0 && matchMode !== 1)
      ) {
        return deps.error(ErrorCode.Value)
      }

      const matchEnd = matchEndNumber !== 0
      if (instanceNumber > 0) {
        let searchFrom = 0
        let found = -1
        for (let count = 0; count < instanceNumber; count += 1) {
          found = indexOfWithMode(text, delimiter, searchFrom, matchMode)
          if (found === -1) {
            return ifNotFoundValue ?? deps.error(ErrorCode.NA)
          }
          searchFrom = found + delimiter.length
        }
        return deps.stringResult(text.slice(0, found))
      }

      let searchFrom = text.length
      let found = matchEnd ? text.length : -1
      for (let count = 0; count < Math.abs(instanceNumber); count += 1) {
        found = lastIndexOfWithMode(text, delimiter, searchFrom, matchMode)
        if (found === -1) {
          return ifNotFoundValue ?? deps.error(ErrorCode.NA)
        }
        searchFrom = found - 1
      }
      return deps.stringResult(text.slice(0, found))
    },
    TEXTAFTER: (...args) => {
      const existingError = deps.firstError(args)
      if (existingError) {
        return existingError
      }
      const [textValue, delimiterValue, instanceValue, matchModeValue, matchEndValue, ifNotFoundValue] = args
      if (textValue === undefined || delimiterValue === undefined) {
        return deps.error(ErrorCode.Value)
      }

      const text = deps.coerceText(textValue)
      const delimiter = deps.coerceText(delimiterValue)
      if (delimiter === '') {
        return deps.error(ErrorCode.Value)
      }

      const instanceNumber = instanceValue === undefined ? 1 : deps.coerceNumber(instanceValue)
      const matchMode = matchModeValue === undefined ? 0 : deps.coerceNumber(matchModeValue)
      const matchEndNumber = matchEndValue === undefined ? 0 : deps.coerceNumber(matchEndValue)
      if (
        instanceNumber === undefined ||
        matchMode === undefined ||
        matchEndNumber === undefined ||
        !Number.isInteger(instanceNumber) ||
        instanceNumber === 0 ||
        !Number.isInteger(matchMode) ||
        (matchMode !== 0 && matchMode !== 1)
      ) {
        return deps.error(ErrorCode.Value)
      }

      const matchEnd = matchEndNumber !== 0
      if (instanceNumber > 0) {
        let searchFrom = 0
        let found = -1
        for (let count = 0; count < instanceNumber; count += 1) {
          found = indexOfWithMode(text, delimiter, searchFrom, matchMode)
          if (found === -1) {
            return ifNotFoundValue ?? deps.error(ErrorCode.NA)
          }
          searchFrom = found + delimiter.length
        }
        return deps.stringResult(text.slice(found + delimiter.length))
      }

      let searchFrom = text.length
      let found = matchEnd ? text.length : -1
      for (let count = 0; count < Math.abs(instanceNumber); count += 1) {
        found = lastIndexOfWithMode(text, delimiter, searchFrom, matchMode)
        if (found === -1) {
          return ifNotFoundValue ?? deps.error(ErrorCode.NA)
        }
        searchFrom = found - 1
      }
      return deps.stringResult(text.slice(found + delimiter.length))
    },
    TEXTJOIN: (...args) => {
      const existingError = deps.firstError(args)
      if (existingError) {
        return existingError
      }
      const [delimiterValue, ignoreEmptyValue, ...values] = args
      if (delimiterValue === undefined || ignoreEmptyValue === undefined || values.length === 0) {
        return deps.error(ErrorCode.Value)
      }

      const delimiter = deps.coerceText(delimiterValue)
      const ignoreEmpty = deps.coerceBoolean(ignoreEmptyValue, false)
      if (typeof ignoreEmpty !== 'boolean') {
        return ignoreEmpty
      }

      const valuesJoined: string[] = []
      for (const value of values) {
        if (value === undefined) {
          continue
        }
        if (value.tag === ValueTag.Empty) {
          if (!ignoreEmpty) {
            valuesJoined.push('')
          }
          continue
        }
        if (value.tag === ValueTag.String && value.value === '' && ignoreEmpty) {
          continue
        }
        if (value.tag === ValueTag.String && value.value === '' && !ignoreEmpty) {
          valuesJoined.push('')
          continue
        }
        valuesJoined.push(deps.coerceText(value))
      }

      return deps.stringResult(valuesJoined.join(delimiter))
    },
  }
}
