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
