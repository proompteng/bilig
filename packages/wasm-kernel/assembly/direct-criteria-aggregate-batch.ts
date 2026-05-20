import { ErrorCode, ValueTag } from './protocol'

const DIRECT_AGGREGATE_OP_SUM: u8 = 1
const DIRECT_AGGREGATE_OP_AVERAGE: u8 = 2
const DIRECT_AGGREGATE_OP_COUNT: u8 = 3
const DIRECT_AGGREGATE_OP_MIN: u8 = 4
const DIRECT_AGGREGATE_OP_MAX: u8 = 5
const CRITERIA_OP_EQ: u8 = 0
const CRITERIA_OP_NE: u8 = 1
const CRITERIA_OP_GT: u8 = 2
const CRITERIA_OP_GTE: u8 = 3
const CRITERIA_OP_LT: u8 = 4
const CRITERIA_OP_LTE: u8 = 5
const CRITERIA_KIND_NUMBER: u8 = 0
const CRITERIA_KIND_STRING_ID: u8 = 1

function writeNumber(index: i32, value: f64, outTags: Uint8Array, outNumbers: Float64Array, outErrors: Uint16Array): void {
  outTags[index] = <u8>ValueTag.Number
  outNumbers[index] = value
  outErrors[index] = ErrorCode.None
}

function writeError(index: i32, code: u16, outTags: Uint8Array, outNumbers: Float64Array, outErrors: Uint16Array): void {
  outTags[index] = <u8>ValueTag.Error
  outNumbers[index] = 0
  outErrors[index] = code
}

function matchesNumericCriteria(valueTag: u8, value: f64, operator: u8, operand: f64): bool {
  if (valueTag == ValueTag.Error || valueTag == ValueTag.String) {
    return false
  }
  if (valueTag == ValueTag.Empty && operator != CRITERIA_OP_EQ && operator != CRITERIA_OP_NE) {
    return false
  }

  const numeric = valueTag == ValueTag.Empty ? 0 : value
  if (operator == CRITERIA_OP_EQ) return numeric == operand
  if (operator == CRITERIA_OP_NE) return numeric != operand
  if (operator == CRITERIA_OP_GT) return numeric > operand
  if (operator == CRITERIA_OP_GTE) return numeric >= operand
  if (operator == CRITERIA_OP_LT) return numeric < operand
  if (operator == CRITERIA_OP_LTE) return numeric <= operand
  return false
}

function matchesStringIdCriteria(valueTag: u8, stringId: u32, operator: u8, operandStringId: u32): bool {
  return operator == CRITERIA_OP_EQ && valueTag == ValueTag.String && stringId == operandStringId
}

function rowMatchesAllCriteria(
  row: i32,
  rowCount: i32,
  criteriaOps: Uint8Array,
  criteriaKinds: Uint8Array,
  criteriaValues: Float64Array,
  criteriaStringIds: Uint32Array,
  criteriaTags: Uint8Array,
  criteriaNumbers: Float64Array,
  criteriaStringIdsByRow: Uint32Array,
): bool {
  const pairCount = <i32>criteriaOps.length
  if (criteriaKinds.length < pairCount || criteriaValues.length < pairCount || criteriaStringIds.length < pairCount) {
    return false
  }
  for (let pairIndex = 0; pairIndex < pairCount; pairIndex += 1) {
    const offset = pairIndex * rowCount + row
    if (offset >= <i32>criteriaTags.length || offset >= <i32>criteriaNumbers.length || offset >= <i32>criteriaStringIdsByRow.length) {
      return false
    }
    const criterionKind = criteriaKinds[pairIndex]
    if (criterionKind == CRITERIA_KIND_STRING_ID) {
      if (
        !matchesStringIdCriteria(criteriaTags[offset], criteriaStringIdsByRow[offset], criteriaOps[pairIndex], criteriaStringIds[pairIndex])
      ) {
        return false
      }
      continue
    }
    if (criterionKind != CRITERIA_KIND_NUMBER) {
      return false
    }
    if (!matchesNumericCriteria(criteriaTags[offset], criteriaNumbers[offset], criteriaOps[pairIndex], criteriaValues[pairIndex])) {
      return false
    }
  }
  return true
}

export function evalDirectCriteriaPredicateAggregateBatch(
  aggregateKind: u8,
  rowCount: i32,
  criteriaOps: Uint8Array,
  criteriaKinds: Uint8Array,
  criteriaValues: Float64Array,
  criteriaStringIds: Uint32Array,
  criteriaTags: Uint8Array,
  criteriaNumbers: Float64Array,
  criteriaStringIdsByRow: Uint32Array,
  aggregateTags: Uint8Array,
  aggregateNumbers: Float64Array,
  aggregateErrors: Uint16Array,
  outTags: Uint8Array,
  outNumbers: Float64Array,
  outErrors: Uint16Array,
): void {
  if (outTags.length == 0 || outNumbers.length == 0 || outErrors.length == 0) {
    return
  }
  if (
    rowCount < 0 ||
    criteriaOps.length == 0 ||
    criteriaKinds.length < criteriaOps.length ||
    criteriaValues.length < criteriaOps.length ||
    criteriaStringIds.length < criteriaOps.length
  ) {
    writeError(0, <u16>ErrorCode.Value, outTags, outNumbers, outErrors)
    return
  }
  const expectedCriteriaLength = rowCount * <i32>criteriaOps.length
  if (
    criteriaTags.length < expectedCriteriaLength ||
    criteriaNumbers.length < expectedCriteriaLength ||
    criteriaStringIdsByRow.length < expectedCriteriaLength
  ) {
    writeError(0, <u16>ErrorCode.Value, outTags, outNumbers, outErrors)
    return
  }
  if (aggregateKind != DIRECT_AGGREGATE_OP_COUNT) {
    if (aggregateTags.length < rowCount || aggregateNumbers.length < rowCount || aggregateErrors.length < rowCount) {
      writeError(0, <u16>ErrorCode.Value, outTags, outNumbers, outErrors)
      return
    }
  }

  let matchedCount: u32 = 0
  let sum: f64 = 0
  let numericCount: u32 = 0
  let minimum: f64 = Infinity
  let maximum: f64 = -Infinity

  for (let row = 0; row < rowCount; row += 1) {
    if (
      !rowMatchesAllCriteria(
        row,
        rowCount,
        criteriaOps,
        criteriaKinds,
        criteriaValues,
        criteriaStringIds,
        criteriaTags,
        criteriaNumbers,
        criteriaStringIdsByRow,
      )
    ) {
      continue
    }
    matchedCount += 1
    if (aggregateKind == DIRECT_AGGREGATE_OP_COUNT) {
      continue
    }

    const tag = aggregateTags[row]
    if (tag == ValueTag.Error) {
      writeError(0, aggregateErrors[row], outTags, outNumbers, outErrors)
      return
    }
    if (aggregateKind == DIRECT_AGGREGATE_OP_SUM) {
      if (tag == ValueTag.Number) {
        sum += aggregateNumbers[row]
      } else if (tag == ValueTag.Boolean) {
        sum += aggregateNumbers[row] != 0 ? 1 : 0
      }
      continue
    }
    if (aggregateKind == DIRECT_AGGREGATE_OP_AVERAGE) {
      if (tag == ValueTag.Number) {
        sum += aggregateNumbers[row]
        numericCount += 1
      } else if (tag == ValueTag.Boolean) {
        sum += aggregateNumbers[row] != 0 ? 1 : 0
        numericCount += 1
      } else if (tag == ValueTag.Empty) {
        numericCount += 1
      }
      continue
    }
    if (aggregateKind == DIRECT_AGGREGATE_OP_MIN) {
      if (tag == ValueTag.Number && aggregateNumbers[row] < minimum) {
        minimum = aggregateNumbers[row]
      }
      continue
    }
    if (aggregateKind == DIRECT_AGGREGATE_OP_MAX) {
      if (tag == ValueTag.Number && aggregateNumbers[row] > maximum) {
        maximum = aggregateNumbers[row]
      }
      continue
    }
    writeError(0, <u16>ErrorCode.Value, outTags, outNumbers, outErrors)
    return
  }

  if (aggregateKind == DIRECT_AGGREGATE_OP_COUNT) {
    writeNumber(0, <f64>matchedCount, outTags, outNumbers, outErrors)
  } else if (aggregateKind == DIRECT_AGGREGATE_OP_SUM) {
    writeNumber(0, sum, outTags, outNumbers, outErrors)
  } else if (aggregateKind == DIRECT_AGGREGATE_OP_AVERAGE) {
    if (numericCount == 0) {
      writeError(0, <u16>ErrorCode.Div0, outTags, outNumbers, outErrors)
    } else {
      writeNumber(0, sum / <f64>numericCount, outTags, outNumbers, outErrors)
    }
  } else if (aggregateKind == DIRECT_AGGREGATE_OP_MIN) {
    writeNumber(0, minimum == Infinity ? 0 : minimum, outTags, outNumbers, outErrors)
  } else if (aggregateKind == DIRECT_AGGREGATE_OP_MAX) {
    writeNumber(0, maximum == -Infinity ? 0 : maximum, outTags, outNumbers, outErrors)
  } else {
    writeError(0, <u16>ErrorCode.Value, outTags, outNumbers, outErrors)
  }
}

export function evalDirectCriteriaMatchedAggregateBatch(
  aggregateKinds: Uint8Array,
  matchStarts: Uint32Array,
  matchLengths: Uint32Array,
  matchedRows: Uint32Array,
  aggregateTags: Uint8Array,
  aggregateNumbers: Float64Array,
  aggregateErrors: Uint16Array,
  outTags: Uint8Array,
  outNumbers: Float64Array,
  outErrors: Uint16Array,
): void {
  const aggregateLength = <u32>aggregateTags.length
  for (let resultIndex = 0; resultIndex < aggregateKinds.length; resultIndex++) {
    const aggregateKind = aggregateKinds[resultIndex]
    const matchStart = matchStarts[resultIndex]
    const matchLength = matchLengths[resultIndex]
    const matchEnd = matchStart + matchLength

    if (aggregateKind == DIRECT_AGGREGATE_OP_COUNT) {
      writeNumber(resultIndex, <f64>matchLength, outTags, outNumbers, outErrors)
      continue
    }

    let sum: f64 = 0
    let count: u32 = 0
    let minimum: f64 = Infinity
    let maximum: f64 = -Infinity
    let hasResult = true

    for (let matchCursor = matchStart; matchCursor < matchEnd; matchCursor++) {
      if (matchCursor >= <u32>matchedRows.length) {
        writeError(resultIndex, <u16>ErrorCode.Value, outTags, outNumbers, outErrors)
        hasResult = false
        break
      }

      const rowOffset = matchedRows[matchCursor]
      if (rowOffset >= aggregateLength) {
        writeError(resultIndex, <u16>ErrorCode.Value, outTags, outNumbers, outErrors)
        hasResult = false
        break
      }

      const tag = aggregateTags[rowOffset]
      if (tag == ValueTag.Error) {
        writeError(resultIndex, aggregateErrors[rowOffset], outTags, outNumbers, outErrors)
        hasResult = false
        break
      }

      if (aggregateKind == DIRECT_AGGREGATE_OP_SUM) {
        if (tag == ValueTag.Number) {
          sum += aggregateNumbers[rowOffset]
        } else if (tag == ValueTag.Boolean) {
          sum += aggregateNumbers[rowOffset] != 0 ? 1 : 0
        }
        continue
      }

      if (aggregateKind == DIRECT_AGGREGATE_OP_AVERAGE) {
        if (tag == ValueTag.Number) {
          sum += aggregateNumbers[rowOffset]
          count += 1
        } else if (tag == ValueTag.Boolean) {
          sum += aggregateNumbers[rowOffset] != 0 ? 1 : 0
          count += 1
        } else if (tag == ValueTag.Empty) {
          count += 1
        }
        continue
      }

      if (aggregateKind == DIRECT_AGGREGATE_OP_MIN) {
        if (tag == ValueTag.Number && aggregateNumbers[rowOffset] < minimum) {
          minimum = aggregateNumbers[rowOffset]
        }
        continue
      }

      if (aggregateKind == DIRECT_AGGREGATE_OP_MAX) {
        if (tag == ValueTag.Number && aggregateNumbers[rowOffset] > maximum) {
          maximum = aggregateNumbers[rowOffset]
        }
        continue
      }

      writeError(resultIndex, <u16>ErrorCode.Value, outTags, outNumbers, outErrors)
      hasResult = false
      break
    }

    if (!hasResult) {
      continue
    }
    if (aggregateKind == DIRECT_AGGREGATE_OP_SUM) {
      writeNumber(resultIndex, sum, outTags, outNumbers, outErrors)
    } else if (aggregateKind == DIRECT_AGGREGATE_OP_AVERAGE) {
      if (count == 0) {
        writeError(resultIndex, <u16>ErrorCode.Div0, outTags, outNumbers, outErrors)
      } else {
        writeNumber(resultIndex, sum / <f64>count, outTags, outNumbers, outErrors)
      }
    } else if (aggregateKind == DIRECT_AGGREGATE_OP_MIN) {
      writeNumber(resultIndex, minimum == Infinity ? 0 : minimum, outTags, outNumbers, outErrors)
    } else if (aggregateKind == DIRECT_AGGREGATE_OP_MAX) {
      writeNumber(resultIndex, maximum == -Infinity ? 0 : maximum, outTags, outNumbers, outErrors)
    } else {
      writeError(resultIndex, <u16>ErrorCode.Value, outTags, outNumbers, outErrors)
    }
  }
}
