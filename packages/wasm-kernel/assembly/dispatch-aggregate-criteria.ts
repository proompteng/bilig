import { BuiltinId, ErrorCode, ValueTag } from './protocol'
import { scalarErrorAt, rangeErrorAt } from './builtin-args'
import { matchesCriteriaValue } from './criteria'
import { gcdPairCalc, lcmPairCalc, truncAbs } from './numeric-core'
import { toNumberOrNaN, toNumberOrZero } from './operands'
import { STACK_KIND_ARRAY, STACK_KIND_RANGE, STACK_KIND_SCALAR, writeResult } from './result-io'
import {
  readCachedRangeSumTag,
  readCachedRangeSumValue,
  readSpillArrayLength,
  readSpillArrayNumber,
  readSpillArrayTag,
  writeCachedRangeSum,
} from './vm'

function writeAggregateError(
  base: i32,
  error: i32,
  rangeIndexStack: Uint32Array,
  valueStack: Float64Array,
  tagStack: Uint8Array,
  kindStack: Uint8Array,
): i32 {
  return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, error, rangeIndexStack, valueStack, tagStack, kindStack)
}

function criteriaScalarValue(memberIndex: i32, cellTags: Uint8Array, cellNumbers: Float64Array, cellStringIds: Uint32Array): f64 {
  const memberTag = cellTags[memberIndex]
  return memberTag == ValueTag.String ? <f64>cellStringIds[memberIndex] : cellNumbers[memberIndex]
}

export function tryApplyAggregateCriteriaBuiltin(
  builtinId: i32,
  argc: i32,
  base: i32,
  rangeIndexStack: Uint32Array,
  valueStack: Float64Array,
  tagStack: Uint8Array,
  kindStack: Uint8Array,
  cellTags: Uint8Array,
  cellNumbers: Float64Array,
  cellStringIds: Uint32Array,
  cellErrors: Uint16Array,
  stringOffsets: Uint32Array,
  stringLengths: Uint32Array,
  stringData: Uint16Array,
  rangeOffsets: Uint32Array,
  rangeLengths: Uint32Array,
  rangeRowCounts: Uint32Array,
  rangeColCounts: Uint32Array,
  rangeMembers: Uint32Array,
  outputStringOffsets: Uint32Array,
  outputStringLengths: Uint32Array,
  outputStringData: Uint16Array,
): i32 {
  if (builtinId == BuiltinId.Sum) {
    if (argc == 1 && kindStack[base] == STACK_KIND_RANGE) {
      const rangeIndex = rangeIndexStack[base]
      const cachedTag = readCachedRangeSumTag(rangeIndex)
      if (cachedTag != 0xff) {
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          cachedTag,
          readCachedRangeSumValue(rangeIndex),
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        )
      }

      const start = rangeOffsets[rangeIndex]
      const length = <i32>rangeLengths[rangeIndex]
      let sum = 0.0
      for (let cursor = 0; cursor < length; cursor += 1) {
        const memberIndex = rangeMembers[start + cursor]
        const memberTag = cellTags[memberIndex]
        if (memberTag == ValueTag.Error) {
          const errorCode = cellErrors[memberIndex]
          writeCachedRangeSum(rangeIndex, <u8>ValueTag.Error, errorCode)
          return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, errorCode, rangeIndexStack, valueStack, tagStack, kindStack)
        }
        const numeric = toNumberOrNaN(memberTag, cellNumbers[memberIndex])
        if (!isNaN(numeric)) {
          sum += numeric
        }
      }

      writeCachedRangeSum(rangeIndex, <u8>ValueTag.Number, sum)
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, sum, rangeIndexStack, valueStack, tagStack, kindStack)
    }

    const scalarError = <i32>scalarErrorAt(base, argc, kindStack, tagStack, valueStack)
    if (scalarError >= 0) {
      return writeAggregateError(base, scalarError, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const rangeError = <i32>(
      rangeErrorAt(base, argc, kindStack, rangeIndexStack, rangeOffsets, rangeLengths, rangeMembers, cellTags, cellErrors)
    )
    if (rangeError >= 0) {
      return writeAggregateError(base, rangeError, rangeIndexStack, valueStack, tagStack, kindStack)
    }

    let sum = 0.0
    for (let index = 0; index < argc; index += 1) {
      const slot = base + index
      if (kindStack[slot] == STACK_KIND_RANGE) {
        const rangeIndex = rangeIndexStack[slot]
        const start = rangeOffsets[rangeIndex]
        const length = <i32>rangeLengths[rangeIndex]
        for (let cursor = 0; cursor < length; cursor += 1) {
          const memberIndex = rangeMembers[start + cursor]
          const numeric = toNumberOrNaN(cellTags[memberIndex], cellNumbers[memberIndex])
          if (!isNaN(numeric)) {
            sum += numeric
          }
        }
        continue
      }
      if (kindStack[slot] == STACK_KIND_ARRAY) {
        const arrayIndex = rangeIndexStack[slot]
        const length = readSpillArrayLength(arrayIndex)
        for (let cursor = 0; cursor < length; cursor += 1) {
          const numeric = readSpillArrayNumber(arrayIndex, cursor)
          if (!isNaN(numeric)) {
            sum += numeric
          }
        }
        continue
      }
      const numeric = toNumberOrNaN(tagStack[slot], valueStack[slot])
      if (!isNaN(numeric)) {
        sum += numeric
      }
    }
    return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, sum, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.Avg) {
    const scalarError = <i32>scalarErrorAt(base, argc, kindStack, tagStack, valueStack)
    if (scalarError >= 0) {
      return writeAggregateError(base, scalarError, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const rangeError = <i32>(
      rangeErrorAt(base, argc, kindStack, rangeIndexStack, rangeOffsets, rangeLengths, rangeMembers, cellTags, cellErrors)
    )
    if (rangeError >= 0) {
      return writeAggregateError(base, rangeError, rangeIndexStack, valueStack, tagStack, kindStack)
    }

    let sum = 0.0
    let count = 0
    for (let index = 0; index < argc; index += 1) {
      const slot = base + index
      if (kindStack[slot] == STACK_KIND_RANGE) {
        const rangeIndex = rangeIndexStack[slot]
        const start = rangeOffsets[rangeIndex]
        const length = <i32>rangeLengths[rangeIndex]
        for (let cursor = 0; cursor < length; cursor += 1) {
          const memberIndex = rangeMembers[start + cursor]
          const numeric = toNumberOrNaN(cellTags[memberIndex], cellNumbers[memberIndex])
          if (!isNaN(numeric)) {
            sum += numeric
            count += 1
          }
        }
        continue
      }
      if (kindStack[slot] == STACK_KIND_ARRAY) {
        const arrayIndex = rangeIndexStack[slot]
        const length = readSpillArrayLength(arrayIndex)
        for (let cursor = 0; cursor < length; cursor += 1) {
          const numeric = readSpillArrayNumber(arrayIndex, cursor)
          if (!isNaN(numeric)) {
            sum += numeric
            count += 1
          }
        }
        continue
      }
      const numeric = toNumberOrNaN(tagStack[slot], valueStack[slot])
      if (!isNaN(numeric)) {
        sum += numeric
        count += 1
      }
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      count == 0 ? 0.0 : sum / <f64>count,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    )
  }

  if (builtinId == BuiltinId.Min) {
    let minValue = Infinity
    for (let index = 0; index < argc; index += 1) {
      const slot = base + index
      if (kindStack[slot] == STACK_KIND_RANGE) {
        const rangeIndex = rangeIndexStack[slot]
        const start = rangeOffsets[rangeIndex]
        const length = <i32>rangeLengths[rangeIndex]
        for (let cursor = 0; cursor < length; cursor += 1) {
          const memberIndex = rangeMembers[start + cursor]
          const numeric = toNumberOrNaN(cellTags[memberIndex], cellNumbers[memberIndex])
          if (!isNaN(numeric) && numeric < minValue) {
            minValue = numeric
          }
        }
        continue
      }
      if (kindStack[slot] == STACK_KIND_ARRAY) {
        const arrayIndex = rangeIndexStack[slot]
        const length = readSpillArrayLength(arrayIndex)
        for (let cursor = 0; cursor < length; cursor += 1) {
          const numeric = readSpillArrayNumber(arrayIndex, cursor)
          if (!isNaN(numeric) && numeric < minValue) {
            minValue = numeric
          }
        }
        continue
      }
      const numeric = toNumberOrNaN(tagStack[slot], valueStack[slot])
      if (!isNaN(numeric) && numeric < minValue) {
        minValue = numeric
      }
    }
    return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, minValue, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.Max) {
    let maxValue = -Infinity
    for (let index = 0; index < argc; index += 1) {
      const slot = base + index
      if (kindStack[slot] == STACK_KIND_RANGE) {
        const rangeIndex = rangeIndexStack[slot]
        const start = rangeOffsets[rangeIndex]
        const length = <i32>rangeLengths[rangeIndex]
        for (let cursor = 0; cursor < length; cursor += 1) {
          const memberIndex = rangeMembers[start + cursor]
          const numeric = toNumberOrNaN(cellTags[memberIndex], cellNumbers[memberIndex])
          if (!isNaN(numeric) && numeric > maxValue) {
            maxValue = numeric
          }
        }
        continue
      }
      if (kindStack[slot] == STACK_KIND_ARRAY) {
        const arrayIndex = rangeIndexStack[slot]
        const length = readSpillArrayLength(arrayIndex)
        for (let cursor = 0; cursor < length; cursor += 1) {
          const numeric = readSpillArrayNumber(arrayIndex, cursor)
          if (!isNaN(numeric) && numeric > maxValue) {
            maxValue = numeric
          }
        }
        continue
      }
      const numeric = toNumberOrNaN(tagStack[slot], valueStack[slot])
      if (!isNaN(numeric) && numeric > maxValue) {
        maxValue = numeric
      }
    }
    return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, maxValue, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.Count) {
    let count = 0
    for (let index = 0; index < argc; index += 1) {
      const slot = base + index
      if (kindStack[slot] == STACK_KIND_RANGE) {
        const rangeIndex = rangeIndexStack[slot]
        const start = rangeOffsets[rangeIndex]
        const length = <i32>rangeLengths[rangeIndex]
        for (let cursor = 0; cursor < length; cursor += 1) {
          const memberIndex = rangeMembers[start + cursor]
          if (!isNaN(toNumberOrNaN(cellTags[memberIndex], cellNumbers[memberIndex]))) {
            count += 1
          }
        }
        continue
      }
      if (kindStack[slot] == STACK_KIND_ARRAY) {
        count += readSpillArrayLength(rangeIndexStack[slot])
        continue
      }
      if (!isNaN(toNumberOrNaN(tagStack[slot], valueStack[slot]))) {
        count += 1
      }
    }
    return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, count, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.CountA) {
    let count = 0
    for (let index = 0; index < argc; index += 1) {
      const slot = base + index
      if (kindStack[slot] == STACK_KIND_RANGE) {
        const rangeIndex = rangeIndexStack[slot]
        const start = rangeOffsets[rangeIndex]
        const length = <i32>rangeLengths[rangeIndex]
        for (let cursor = 0; cursor < length; cursor += 1) {
          const memberIndex = rangeMembers[start + cursor]
          if (cellTags[memberIndex] != ValueTag.Empty) {
            count += 1
          }
        }
        continue
      }
      if (kindStack[slot] == STACK_KIND_ARRAY) {
        count += readSpillArrayLength(rangeIndexStack[slot])
        continue
      }
      if (tagStack[slot] != ValueTag.Empty) {
        count += 1
      }
    }
    return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, count, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.Countblank) {
    let count = 0
    for (let index = 0; index < argc; index += 1) {
      const slot = base + index
      if (kindStack[slot] == STACK_KIND_RANGE) {
        const rangeIndex = rangeIndexStack[slot]
        const start = rangeOffsets[rangeIndex]
        const length = <i32>rangeLengths[rangeIndex]
        for (let cursor = 0; cursor < length; cursor += 1) {
          const memberIndex = rangeMembers[start + cursor]
          if (cellTags[memberIndex] == ValueTag.Empty) {
            count += 1
          }
        }
        continue
      }
      if (kindStack[slot] == STACK_KIND_ARRAY) {
        const arrayIndex = rangeIndexStack[slot]
        const length = readSpillArrayLength(arrayIndex)
        for (let cursor = 0; cursor < length; cursor += 1) {
          if (readSpillArrayTag(arrayIndex, cursor) == ValueTag.Empty) {
            count += 1
          }
        }
        continue
      }
      if (tagStack[slot] == ValueTag.Empty) {
        count += 1
      }
    }
    return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, count, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (
    builtinId == BuiltinId.Gcd ||
    builtinId == BuiltinId.Lcm ||
    builtinId == BuiltinId.Product ||
    builtinId == BuiltinId.Geomean ||
    builtinId == BuiltinId.Harmean ||
    builtinId == BuiltinId.Sumsq
  ) {
    const scalarError = <i32>scalarErrorAt(base, argc, kindStack, tagStack, valueStack)
    if (scalarError >= 0) {
      return writeAggregateError(base, scalarError, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const rangeError = <i32>(
      rangeErrorAt(base, argc, kindStack, rangeIndexStack, rangeOffsets, rangeLengths, rangeMembers, cellTags, cellErrors)
    )
    if (rangeError >= 0) {
      return writeAggregateError(base, rangeError, rangeIndexStack, valueStack, tagStack, kindStack)
    }

    let count = 0
    let product = 1.0
    let sumSquares = 0.0
    let gcdValue = 0.0
    let lcmValue = 0.0
    let logSum = 0.0
    let reciprocalSum = 0.0

    for (let index = 0; index < argc; index += 1) {
      const slot = base + index
      if (kindStack[slot] == STACK_KIND_RANGE) {
        const rangeIndex = rangeIndexStack[slot]
        const start = rangeOffsets[rangeIndex]
        const length = <i32>rangeLengths[rangeIndex]
        for (let cursor = 0; cursor < length; cursor += 1) {
          const memberIndex = rangeMembers[start + cursor]
          const numeric = toNumberOrNaN(cellTags[memberIndex], cellNumbers[memberIndex])
          if (isNaN(numeric)) {
            continue
          }
          if (builtinId == BuiltinId.Geomean) {
            if (numeric < 0.0) {
              return writeAggregateError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
            }
            if (numeric == 0.0) {
              return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, 0.0, rangeIndexStack, valueStack, tagStack, kindStack)
            }
            logSum += Math.log(numeric)
          } else if (builtinId == BuiltinId.Harmean) {
            if (numeric <= 0.0) {
              return writeAggregateError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
            }
            reciprocalSum += 1.0 / numeric
          } else if (builtinId == BuiltinId.Gcd) {
            gcdValue = count == 0 ? truncAbs(numeric) : gcdPairCalc(gcdValue, numeric)
          } else if (builtinId == BuiltinId.Lcm) {
            lcmValue = count == 0 ? truncAbs(numeric) : lcmPairCalc(lcmValue, numeric)
          } else if (builtinId == BuiltinId.Product) {
            product *= numeric
          } else if (builtinId == BuiltinId.Sumsq) {
            sumSquares += numeric * numeric
          }
          count += 1
        }
        continue
      }
      if (kindStack[slot] == STACK_KIND_ARRAY) {
        const arrayIndex = rangeIndexStack[slot]
        const length = readSpillArrayLength(arrayIndex)
        for (let cursor = 0; cursor < length; cursor += 1) {
          const numeric = readSpillArrayNumber(arrayIndex, cursor)
          if (isNaN(numeric)) {
            continue
          }
          if (builtinId == BuiltinId.Geomean) {
            if (numeric < 0.0) {
              return writeAggregateError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
            }
            if (numeric == 0.0) {
              return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, 0.0, rangeIndexStack, valueStack, tagStack, kindStack)
            }
            logSum += Math.log(numeric)
          } else if (builtinId == BuiltinId.Harmean) {
            if (numeric <= 0.0) {
              return writeAggregateError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
            }
            reciprocalSum += 1.0 / numeric
          } else if (builtinId == BuiltinId.Gcd) {
            gcdValue = count == 0 ? truncAbs(numeric) : gcdPairCalc(gcdValue, numeric)
          } else if (builtinId == BuiltinId.Lcm) {
            lcmValue = count == 0 ? truncAbs(numeric) : lcmPairCalc(lcmValue, numeric)
          } else if (builtinId == BuiltinId.Product) {
            product *= numeric
          } else if (builtinId == BuiltinId.Sumsq) {
            sumSquares += numeric * numeric
          }
          count += 1
        }
        continue
      }

      const numeric = toNumberOrNaN(tagStack[slot], valueStack[slot])
      if (isNaN(numeric)) {
        continue
      }
      if (builtinId == BuiltinId.Geomean) {
        if (numeric < 0.0) {
          return writeAggregateError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
        }
        if (numeric == 0.0) {
          return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, 0.0, rangeIndexStack, valueStack, tagStack, kindStack)
        }
        logSum += Math.log(numeric)
      } else if (builtinId == BuiltinId.Harmean) {
        if (numeric <= 0.0) {
          return writeAggregateError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
        }
        reciprocalSum += 1.0 / numeric
      } else if (builtinId == BuiltinId.Gcd) {
        gcdValue = count == 0 ? truncAbs(numeric) : gcdPairCalc(gcdValue, numeric)
      } else if (builtinId == BuiltinId.Lcm) {
        lcmValue = count == 0 ? truncAbs(numeric) : lcmPairCalc(lcmValue, numeric)
      } else if (builtinId == BuiltinId.Product) {
        product *= numeric
      } else if (builtinId == BuiltinId.Sumsq) {
        sumSquares += numeric * numeric
      }
      count += 1
    }

    if (
      (builtinId == BuiltinId.Gcd || builtinId == BuiltinId.Lcm || builtinId == BuiltinId.Geomean || builtinId == BuiltinId.Harmean) &&
      count == 0
    ) {
      return writeAggregateError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }

    let result = 0.0
    if (builtinId == BuiltinId.Gcd) {
      result = gcdValue
    } else if (builtinId == BuiltinId.Lcm) {
      result = lcmValue
    } else if (builtinId == BuiltinId.Product) {
      result = count == 0 ? 0.0 : product
    } else if (builtinId == BuiltinId.Geomean) {
      result = Math.exp(logSum / <f64>count)
    } else if (builtinId == BuiltinId.Harmean) {
      result = <f64>count / reciprocalSum
    } else {
      result = sumSquares
    }

    return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, result, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.Countif && argc == 2) {
    if (kindStack[base] != STACK_KIND_RANGE || kindStack[base + 1] != STACK_KIND_SCALAR) {
      return writeAggregateError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    if (tagStack[base + 1] == ValueTag.Error) {
      return writeAggregateError(base, <i32>valueStack[base + 1], rangeIndexStack, valueStack, tagStack, kindStack)
    }

    const rangeIndex = rangeIndexStack[base]
    const start = rangeOffsets[rangeIndex]
    const length = <i32>rangeLengths[rangeIndex]
    let count = 0
    for (let cursor = 0; cursor < length; cursor += 1) {
      const memberIndex = rangeMembers[start + cursor]
      if (
        matchesCriteriaValue(
          cellTags[memberIndex],
          criteriaScalarValue(memberIndex, cellTags, cellNumbers, cellStringIds),
          tagStack[base + 1],
          valueStack[base + 1],
          stringOffsets,
          stringLengths,
          stringData,
          outputStringOffsets,
          outputStringLengths,
          outputStringData,
        )
      ) {
        count += 1
      }
    }
    return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, count, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.Countifs) {
    if (argc == 0 || argc % 2 != 0) {
      return writeAggregateError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }

    const firstRangeIndex = rangeIndexStack[base]
    if (kindStack[base] != STACK_KIND_RANGE) {
      return writeAggregateError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const expectedLength = <i32>rangeLengths[firstRangeIndex]
    for (let index = 0; index < argc; index += 2) {
      const rangeSlot = base + index
      const criteriaSlot = rangeSlot + 1
      if (kindStack[rangeSlot] != STACK_KIND_RANGE || kindStack[criteriaSlot] != STACK_KIND_SCALAR) {
        return writeAggregateError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      if (tagStack[criteriaSlot] == ValueTag.Error) {
        return writeAggregateError(base, <i32>valueStack[criteriaSlot], rangeIndexStack, valueStack, tagStack, kindStack)
      }
      if (<i32>rangeLengths[rangeIndexStack[rangeSlot]] != expectedLength) {
        return writeAggregateError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      }
    }

    let count = 0
    for (let row = 0; row < expectedLength; row += 1) {
      let matchesAll = true
      for (let index = 0; index < argc; index += 2) {
        const rangeSlot = base + index
        const criteriaSlot = rangeSlot + 1
        const rangeIndex = rangeIndexStack[rangeSlot]
        const memberIndex = rangeMembers[rangeOffsets[rangeIndex] + row]
        if (
          !matchesCriteriaValue(
            cellTags[memberIndex],
            criteriaScalarValue(memberIndex, cellTags, cellNumbers, cellStringIds),
            tagStack[criteriaSlot],
            valueStack[criteriaSlot],
            stringOffsets,
            stringLengths,
            stringData,
            outputStringOffsets,
            outputStringLengths,
            outputStringData,
          )
        ) {
          matchesAll = false
          break
        }
      }
      if (matchesAll) {
        count += 1
      }
    }
    return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, count, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.Sumif && (argc == 2 || argc == 3)) {
    const rangeSlot = base
    const criteriaSlot = base + 1
    const sumRangeSlot = argc == 3 ? base + 2 : base
    if (
      kindStack[rangeSlot] != STACK_KIND_RANGE ||
      kindStack[criteriaSlot] != STACK_KIND_SCALAR ||
      kindStack[sumRangeSlot] != STACK_KIND_RANGE
    ) {
      return writeAggregateError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    if (tagStack[criteriaSlot] == ValueTag.Error) {
      return writeAggregateError(base, <i32>valueStack[criteriaSlot], rangeIndexStack, valueStack, tagStack, kindStack)
    }

    const rangeIndex = rangeIndexStack[rangeSlot]
    const sumRangeIndex = rangeIndexStack[sumRangeSlot]
    const length = <i32>rangeLengths[rangeIndex]
    if (<i32>rangeLengths[sumRangeIndex] != length) {
      return writeAggregateError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }

    let sum = 0.0
    for (let cursor = 0; cursor < length; cursor += 1) {
      const criteriaMemberIndex = rangeMembers[rangeOffsets[rangeIndex] + cursor]
      if (
        !matchesCriteriaValue(
          cellTags[criteriaMemberIndex],
          criteriaScalarValue(criteriaMemberIndex, cellTags, cellNumbers, cellStringIds),
          tagStack[criteriaSlot],
          valueStack[criteriaSlot],
          stringOffsets,
          stringLengths,
          stringData,
          outputStringOffsets,
          outputStringLengths,
          outputStringData,
        )
      ) {
        continue
      }
      const sumMemberIndex = rangeMembers[rangeOffsets[sumRangeIndex] + cursor]
      sum += toNumberOrZero(cellTags[sumMemberIndex], cellNumbers[sumMemberIndex])
    }
    return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, sum, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.Sumifs) {
    if (argc < 3 || argc % 2 == 0 || kindStack[base] != STACK_KIND_RANGE) {
      return writeAggregateError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }

    const sumRangeIndex = rangeIndexStack[base]
    const expectedLength = <i32>rangeLengths[sumRangeIndex]
    for (let index = 1; index < argc; index += 2) {
      const rangeSlot = base + index
      const criteriaSlot = rangeSlot + 1
      if (kindStack[rangeSlot] != STACK_KIND_RANGE || kindStack[criteriaSlot] != STACK_KIND_SCALAR) {
        return writeAggregateError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      if (tagStack[criteriaSlot] == ValueTag.Error) {
        return writeAggregateError(base, <i32>valueStack[criteriaSlot], rangeIndexStack, valueStack, tagStack, kindStack)
      }
      if (<i32>rangeLengths[rangeIndexStack[rangeSlot]] != expectedLength) {
        return writeAggregateError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      }
    }

    let sum = 0.0
    for (let row = 0; row < expectedLength; row += 1) {
      let matchesAll = true
      for (let index = 1; index < argc; index += 2) {
        const rangeSlot = base + index
        const criteriaSlot = rangeSlot + 1
        const rangeIndex = rangeIndexStack[rangeSlot]
        const memberIndex = rangeMembers[rangeOffsets[rangeIndex] + row]
        if (
          !matchesCriteriaValue(
            cellTags[memberIndex],
            criteriaScalarValue(memberIndex, cellTags, cellNumbers, cellStringIds),
            tagStack[criteriaSlot],
            valueStack[criteriaSlot],
            stringOffsets,
            stringLengths,
            stringData,
            outputStringOffsets,
            outputStringLengths,
            outputStringData,
          )
        ) {
          matchesAll = false
          break
        }
      }
      if (!matchesAll) {
        continue
      }
      const sumMemberIndex = rangeMembers[rangeOffsets[sumRangeIndex] + row]
      sum += toNumberOrZero(cellTags[sumMemberIndex], cellNumbers[sumMemberIndex])
    }
    return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, sum, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.Averageif && (argc == 2 || argc == 3)) {
    const rangeSlot = base
    const criteriaSlot = base + 1
    const averageRangeSlot = argc == 3 ? base + 2 : base
    if (
      kindStack[rangeSlot] != STACK_KIND_RANGE ||
      kindStack[criteriaSlot] != STACK_KIND_SCALAR ||
      kindStack[averageRangeSlot] != STACK_KIND_RANGE
    ) {
      return writeAggregateError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    if (tagStack[criteriaSlot] == ValueTag.Error) {
      return writeAggregateError(base, <i32>valueStack[criteriaSlot], rangeIndexStack, valueStack, tagStack, kindStack)
    }

    const rangeIndex = rangeIndexStack[rangeSlot]
    const averageRangeIndex = rangeIndexStack[averageRangeSlot]
    const length = <i32>rangeLengths[rangeIndex]
    if (<i32>rangeLengths[averageRangeIndex] != length) {
      return writeAggregateError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }

    let count = 0
    let sum = 0.0
    for (let cursor = 0; cursor < length; cursor += 1) {
      const criteriaMemberIndex = rangeMembers[rangeOffsets[rangeIndex] + cursor]
      if (
        !matchesCriteriaValue(
          cellTags[criteriaMemberIndex],
          criteriaScalarValue(criteriaMemberIndex, cellTags, cellNumbers, cellStringIds),
          tagStack[criteriaSlot],
          valueStack[criteriaSlot],
          stringOffsets,
          stringLengths,
          stringData,
          outputStringOffsets,
          outputStringLengths,
          outputStringData,
        )
      ) {
        continue
      }
      const averageMemberIndex = rangeMembers[rangeOffsets[averageRangeIndex] + cursor]
      const numeric = toNumberOrNaN(cellTags[averageMemberIndex], cellNumbers[averageMemberIndex])
      if (isNaN(numeric)) {
        continue
      }
      count += 1
      sum += numeric
    }
    if (count == 0) {
      return writeAggregateError(base, ErrorCode.Div0, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, sum / <f64>count, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.Averageifs) {
    if (argc < 3 || argc % 2 == 0 || kindStack[base] != STACK_KIND_RANGE) {
      return writeAggregateError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }

    const averageRangeIndex = rangeIndexStack[base]
    const expectedLength = <i32>rangeLengths[averageRangeIndex]
    for (let index = 1; index < argc; index += 2) {
      const rangeSlot = base + index
      const criteriaSlot = rangeSlot + 1
      if (kindStack[rangeSlot] != STACK_KIND_RANGE || kindStack[criteriaSlot] != STACK_KIND_SCALAR) {
        return writeAggregateError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      if (tagStack[criteriaSlot] == ValueTag.Error) {
        return writeAggregateError(base, <i32>valueStack[criteriaSlot], rangeIndexStack, valueStack, tagStack, kindStack)
      }
      if (<i32>rangeLengths[rangeIndexStack[rangeSlot]] != expectedLength) {
        return writeAggregateError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      }
    }

    let count = 0
    let sum = 0.0
    for (let row = 0; row < expectedLength; row += 1) {
      let matchesAll = true
      for (let index = 1; index < argc; index += 2) {
        const rangeSlot = base + index
        const criteriaSlot = rangeSlot + 1
        const rangeIndex = rangeIndexStack[rangeSlot]
        const memberIndex = rangeMembers[rangeOffsets[rangeIndex] + row]
        if (
          !matchesCriteriaValue(
            cellTags[memberIndex],
            criteriaScalarValue(memberIndex, cellTags, cellNumbers, cellStringIds),
            tagStack[criteriaSlot],
            valueStack[criteriaSlot],
            stringOffsets,
            stringLengths,
            stringData,
            outputStringOffsets,
            outputStringLengths,
            outputStringData,
          )
        ) {
          matchesAll = false
          break
        }
      }
      if (!matchesAll) {
        continue
      }
      const averageMemberIndex = rangeMembers[rangeOffsets[averageRangeIndex] + row]
      const numeric = toNumberOrNaN(cellTags[averageMemberIndex], cellNumbers[averageMemberIndex])
      if (isNaN(numeric)) {
        continue
      }
      count += 1
      sum += numeric
    }
    if (count == 0) {
      return writeAggregateError(base, ErrorCode.Div0, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, sum / <f64>count, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.Minifs || builtinId == BuiltinId.Maxifs) {
    if (argc < 3 || argc % 2 == 0 || kindStack[base] != STACK_KIND_RANGE) {
      return writeAggregateError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }

    const targetRangeIndex = rangeIndexStack[base]
    const expectedLength = <i32>rangeLengths[targetRangeIndex]
    for (let index = 1; index < argc; index += 2) {
      const rangeSlot = base + index
      const criteriaSlot = rangeSlot + 1
      if (kindStack[rangeSlot] != STACK_KIND_RANGE || kindStack[criteriaSlot] != STACK_KIND_SCALAR) {
        return writeAggregateError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      if (tagStack[criteriaSlot] == ValueTag.Error) {
        return writeAggregateError(base, <i32>valueStack[criteriaSlot], rangeIndexStack, valueStack, tagStack, kindStack)
      }
      if (<i32>rangeLengths[rangeIndexStack[rangeSlot]] != expectedLength) {
        return writeAggregateError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      }
    }

    let found = false
    let result = builtinId == BuiltinId.Minifs ? Infinity : -Infinity
    for (let row = 0; row < expectedLength; row += 1) {
      let matchesAll = true
      for (let index = 1; index < argc; index += 2) {
        const rangeSlot = base + index
        const criteriaSlot = rangeSlot + 1
        const rangeIndex = rangeIndexStack[rangeSlot]
        const memberIndex = rangeMembers[rangeOffsets[rangeIndex] + row]
        if (
          !matchesCriteriaValue(
            cellTags[memberIndex],
            criteriaScalarValue(memberIndex, cellTags, cellNumbers, cellStringIds),
            tagStack[criteriaSlot],
            valueStack[criteriaSlot],
            stringOffsets,
            stringLengths,
            stringData,
            outputStringOffsets,
            outputStringLengths,
            outputStringData,
          )
        ) {
          matchesAll = false
          break
        }
      }
      if (!matchesAll) {
        continue
      }
      const targetMemberIndex = rangeMembers[rangeOffsets[targetRangeIndex] + row]
      if (cellTags[targetMemberIndex] != ValueTag.Number) {
        continue
      }
      const numeric = cellNumbers[targetMemberIndex]
      result = builtinId == BuiltinId.Minifs ? min(result, numeric) : max(result, numeric)
      found = true
    }
    return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, found ? result : 0.0, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  return -1
}
