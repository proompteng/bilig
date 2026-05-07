import { ValueTag } from './protocol'
import { compareScalarValues } from './comparison'
import { scalarText, trimAsciiWhitespace } from './text-codec'
import { parseNumericText } from './text-special'

const CRITERIA_OP_EQ: i32 = 0
const CRITERIA_OP_NE: i32 = 1
const CRITERIA_OP_GT: i32 = 2
const CRITERIA_OP_GTE: i32 = 3
const CRITERIA_OP_LT: i32 = 4
const CRITERIA_OP_LTE: i32 = 5

function hasCriteriaWildcardPattern(pattern: string): bool {
  for (let index = 0; index < pattern.length; index += 1) {
    const char = pattern.charCodeAt(index)
    if (char == 126) {
      index += 1
      continue
    }
    if (char == 42 || char == 63) {
      return true
    }
  }
  return false
}

function unescapeCriteriaPattern(pattern: string): string {
  let unescaped = ''
  for (let index = 0; index < pattern.length; index += 1) {
    const char = pattern.charCodeAt(index)
    if (char == 126) {
      const escapedIndex = index + 1
      if (escapedIndex < pattern.length) {
        unescaped += String.fromCharCode(pattern.charCodeAt(escapedIndex))
        index = escapedIndex
        continue
      }
    }
    unescaped += String.fromCharCode(char)
  }
  return unescaped
}

function wildcardCriteriaMatches(pattern: string, value: string): bool {
  let p = 0
  let v = 0
  let starPatternIndex = -1
  let starValueIndex = -1

  while (v < value.length) {
    if (p < pattern.length) {
      const char = pattern.charCodeAt(p)
      if (char == 126) {
        const escapedIndex = p + 1
        const expected = escapedIndex < pattern.length ? pattern.charCodeAt(escapedIndex) : 126
        if (value.charCodeAt(v) == expected) {
          p = escapedIndex < pattern.length ? escapedIndex + 1 : escapedIndex
          v += 1
          continue
        }
      }
      if (char == 42) {
        p += 1
        while (p < pattern.length && pattern.charCodeAt(p) == 42) {
          p += 1
        }
        starPatternIndex = p
        starValueIndex = v
        continue
      }
      if (char == 63 || value.charCodeAt(v) == char) {
        p += 1
        v += 1
        continue
      }
    }

    if (starPatternIndex >= 0) {
      starValueIndex += 1
      p = starPatternIndex
      v = starValueIndex
      continue
    }

    return false
  }

  while (p < pattern.length && pattern.charCodeAt(p) == 42) {
    p += 1
  }
  return p == pattern.length
}

export function matchesCriteriaValue(
  valueTag: u8,
  valueValue: f64,
  criteriaTag: u8,
  criteriaValue: f64,
  stringOffsets: Uint32Array,
  stringLengths: Uint32Array,
  stringData: Uint16Array,
  outputStringOffsets: Uint32Array,
  outputStringLengths: Uint32Array,
  outputStringData: Uint16Array,
): bool {
  if (valueTag == ValueTag.Error) {
    return false
  }

  let operator = CRITERIA_OP_EQ
  let operandTag = criteriaTag
  let operandValue = criteriaValue
  let operandText: string | null = null

  if (criteriaTag == ValueTag.String) {
    const criteriaText = scalarText(
      criteriaTag,
      criteriaValue,
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    )
    if (criteriaText == null) {
      return false
    }

    let rawOperand = criteriaText
    let parsedOperator = false
    if (criteriaText.length >= 2) {
      const prefix = criteriaText.slice(0, 2)
      if (prefix == '<=') {
        operator = CRITERIA_OP_LTE
        rawOperand = criteriaText.slice(2)
        parsedOperator = true
      } else if (prefix == '>=') {
        operator = CRITERIA_OP_GTE
        rawOperand = criteriaText.slice(2)
        parsedOperator = true
      } else if (prefix == '<>') {
        operator = CRITERIA_OP_NE
        rawOperand = criteriaText.slice(2)
        parsedOperator = true
      }
    }
    if (!parsedOperator && criteriaText.length >= 1) {
      const first = criteriaText.charCodeAt(0)
      if (first == 61) {
        operator = CRITERIA_OP_EQ
        rawOperand = criteriaText.slice(1)
        parsedOperator = true
      } else if (first == 62) {
        operator = CRITERIA_OP_GT
        rawOperand = criteriaText.slice(1)
        parsedOperator = true
      } else if (first == 60) {
        operator = CRITERIA_OP_LT
        rawOperand = criteriaText.slice(1)
        parsedOperator = true
      }
    }

    if (!parsedOperator) {
      operandText = criteriaText
    } else {
      const trimmed = trimAsciiWhitespace(rawOperand)
      if (trimmed.length == 0) {
        operandTag = <u8>ValueTag.String
        operandValue = 0
        operandText = ''
      } else {
        const upper = trimmed.toUpperCase()
        if (upper == 'TRUE' || upper == 'FALSE') {
          operandTag = <u8>ValueTag.Boolean
          operandValue = upper == 'TRUE' ? 1 : 0
        } else {
          const numeric = parseNumericText(trimmed)
          if (!isNaN(numeric)) {
            operandTag = <u8>ValueTag.Number
            operandValue = numeric
          } else {
            operandTag = <u8>ValueTag.String
            operandValue = 0
            operandText = trimmed
          }
        }
      }
    }
  }

  if ((operator == CRITERIA_OP_EQ || operator == CRITERIA_OP_NE) && operandText != null) {
    const textPattern = operandText
    if (hasCriteriaWildcardPattern(textPattern)) {
      const valueText = scalarText(
        valueTag,
        valueValue,
        stringOffsets,
        stringLengths,
        stringData,
        outputStringOffsets,
        outputStringLengths,
        outputStringData,
      )
      if (valueText == null) {
        return false
      }
      const matches = wildcardCriteriaMatches(textPattern.toUpperCase(), valueText.toUpperCase())
      return operator == CRITERIA_OP_EQ ? matches : !matches
    }
    if (textPattern.indexOf('~') >= 0) {
      operandText = unescapeCriteriaPattern(textPattern)
    }
  }

  const comparison = compareScalarValues(
    valueTag,
    valueValue,
    operandTag,
    operandValue,
    operandText,
    stringOffsets,
    stringLengths,
    stringData,
    outputStringOffsets,
    outputStringLengths,
    outputStringData,
  )
  if (comparison == i32.MIN_VALUE) {
    return false
  }
  if (operator == CRITERIA_OP_EQ) {
    return comparison == 0
  }
  if (operator == CRITERIA_OP_NE) {
    return comparison != 0
  }
  if (operator == CRITERIA_OP_GT) {
    return comparison > 0
  }
  if (operator == CRITERIA_OP_GTE) {
    return comparison >= 0
  }
  if (operator == CRITERIA_OP_LT) {
    return comparison < 0
  }
  return comparison <= 0
}
