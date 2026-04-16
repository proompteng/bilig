import { BuiltinId, ErrorCode, ValueTag } from './protocol'
import { compareScalarValues } from './comparison'
import { memberScalarValue } from './operands'
import { truncToInt } from './numeric-core'
import { STACK_KIND_RANGE, STACK_KIND_SCALAR, vectorSlotLength, writeMemberResult, writeResult } from './result-io'

export function tryApplyLookupMatchBuiltin(
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
  if (builtinId == BuiltinId.Match && (argc == 2 || argc == 3)) {
    if (kindStack[base] != STACK_KIND_SCALAR || kindStack[base + 1] != STACK_KIND_RANGE) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    if (tagStack[base] == ValueTag.Error) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, valueStack[base], rangeIndexStack, valueStack, tagStack, kindStack)
    }
    if (argc == 3 && (kindStack[base + 2] != STACK_KIND_SCALAR || tagStack[base + 2] == ValueTag.Error)) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        argc == 3 && tagStack[base + 2] == ValueTag.Error ? valueStack[base + 2] : ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      )
    }

    const matchType = argc == 3 ? truncToInt(tagStack[base + 2], valueStack[base + 2]) : 1
    if (!(matchType == -1 || matchType == 0 || matchType == 1)) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }

    const rangeIndex = rangeIndexStack[base + 1]
    const start = rangeOffsets[rangeIndex]
    const length = <i32>rangeLengths[rangeIndex]
    let best = -1
    for (let index = 0; index < length; index++) {
      const memberIndex = rangeMembers[start + index]
      const comparison = compareScalarValues(
        cellTags[memberIndex],
        memberScalarValue(memberIndex, cellTags, cellNumbers, cellStringIds, cellErrors),
        tagStack[base],
        valueStack[base],
        null,
        stringOffsets,
        stringLengths,
        stringData,
        outputStringOffsets,
        outputStringLengths,
        outputStringData,
      )
      if (matchType == 0) {
        if (comparison == 0) {
          return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, index + 1, rangeIndexStack, valueStack, tagStack, kindStack)
        }
        continue
      }
      if (comparison == i32.MIN_VALUE) {
        return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.NA, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      if (matchType == 1) {
        if (comparison <= 0) {
          best = index + 1
        } else {
          break
        }
      } else if (comparison >= 0) {
        best = index + 1
      } else {
        break
      }
    }
    return best < 0
      ? writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.NA, rangeIndexStack, valueStack, tagStack, kindStack)
      : writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, best, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.Xmatch && argc >= 2 && argc <= 4) {
    if (kindStack[base] != STACK_KIND_SCALAR || kindStack[base + 1] != STACK_KIND_RANGE) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    if (tagStack[base] == ValueTag.Error) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, valueStack[base], rangeIndexStack, valueStack, tagStack, kindStack)
    }
    if (argc >= 3 && (kindStack[base + 2] != STACK_KIND_SCALAR || tagStack[base + 2] == ValueTag.Error)) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        argc >= 3 && tagStack[base + 2] == ValueTag.Error ? valueStack[base + 2] : ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      )
    }
    if (argc == 4 && (kindStack[base + 3] != STACK_KIND_SCALAR || tagStack[base + 3] == ValueTag.Error)) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        argc == 4 && tagStack[base + 3] == ValueTag.Error ? valueStack[base + 3] : ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      )
    }

    const matchMode = argc >= 3 ? truncToInt(tagStack[base + 2], valueStack[base + 2]) : 0
    const searchMode = argc == 4 ? truncToInt(tagStack[base + 3], valueStack[base + 3]) : 1
    if (!(matchMode == -1 || matchMode == 0 || matchMode == 1) || !(searchMode == -1 || searchMode == 1)) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }

    const rangeIndex = rangeIndexStack[base + 1]
    const start = rangeOffsets[rangeIndex]
    const length = <i32>rangeLengths[rangeIndex]
    if (searchMode == -1) {
      if (matchMode == 0) {
        for (let index = length - 1; index >= 0; index--) {
          const memberIndex = rangeMembers[start + index]
          const comparison = compareScalarValues(
            cellTags[memberIndex],
            memberScalarValue(memberIndex, cellTags, cellNumbers, cellStringIds, cellErrors),
            tagStack[base],
            valueStack[base],
            null,
            stringOffsets,
            stringLengths,
            stringData,
            outputStringOffsets,
            outputStringLengths,
            outputStringData,
          )
          if (comparison == 0) {
            return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, index + 1, rangeIndexStack, valueStack, tagStack, kindStack)
          }
        }
        return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.NA, rangeIndexStack, valueStack, tagStack, kindStack)
      }

      let bestReversed = -1
      let reversedPosition = 0
      for (let index = length - 1; index >= 0; index--) {
        reversedPosition += 1
        const memberIndex = rangeMembers[start + index]
        const comparison = compareScalarValues(
          cellTags[memberIndex],
          memberScalarValue(memberIndex, cellTags, cellNumbers, cellStringIds, cellErrors),
          tagStack[base],
          valueStack[base],
          null,
          stringOffsets,
          stringLengths,
          stringData,
          outputStringOffsets,
          outputStringLengths,
          outputStringData,
        )
        if (comparison == i32.MIN_VALUE) {
          return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.NA, rangeIndexStack, valueStack, tagStack, kindStack)
        }
        if (matchMode == 1) {
          if (comparison <= 0) {
            bestReversed = reversedPosition
          } else {
            break
          }
        } else if (comparison >= 0) {
          bestReversed = reversedPosition
        } else {
          break
        }
      }
      return bestReversed < 0
        ? writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.NA, rangeIndexStack, valueStack, tagStack, kindStack)
        : writeResult(
            base,
            STACK_KIND_SCALAR,
            <u8>ValueTag.Number,
            length - bestReversed + 1,
            rangeIndexStack,
            valueStack,
            tagStack,
            kindStack,
          )
    }

    if (matchMode == 0) {
      for (let index = 0; index < length; index++) {
        const memberIndex = rangeMembers[start + index]
        const comparison = compareScalarValues(
          cellTags[memberIndex],
          memberScalarValue(memberIndex, cellTags, cellNumbers, cellStringIds, cellErrors),
          tagStack[base],
          valueStack[base],
          null,
          stringOffsets,
          stringLengths,
          stringData,
          outputStringOffsets,
          outputStringLengths,
          outputStringData,
        )
        if (comparison == 0) {
          return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, index + 1, rangeIndexStack, valueStack, tagStack, kindStack)
        }
      }
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.NA, rangeIndexStack, valueStack, tagStack, kindStack)
    }

    let best = -1
    for (let index = 0; index < length; index++) {
      const memberIndex = rangeMembers[start + index]
      const comparison = compareScalarValues(
        cellTags[memberIndex],
        memberScalarValue(memberIndex, cellTags, cellNumbers, cellStringIds, cellErrors),
        tagStack[base],
        valueStack[base],
        null,
        stringOffsets,
        stringLengths,
        stringData,
        outputStringOffsets,
        outputStringLengths,
        outputStringData,
      )
      if (comparison == i32.MIN_VALUE) {
        return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.NA, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      if (matchMode == 1) {
        if (comparison <= 0) {
          best = index + 1
        } else {
          break
        }
      } else if (comparison >= 0) {
        best = index + 1
      } else {
        break
      }
    }
    return best < 0
      ? writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.NA, rangeIndexStack, valueStack, tagStack, kindStack)
      : writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, best, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.Xlookup && argc >= 3 && argc <= 6) {
    if (kindStack[base] != STACK_KIND_SCALAR || kindStack[base + 1] != STACK_KIND_RANGE || kindStack[base + 2] != STACK_KIND_RANGE) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    if (tagStack[base] == ValueTag.Error) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, valueStack[base], rangeIndexStack, valueStack, tagStack, kindStack)
    }
    if (argc >= 4 && (kindStack[base + 3] != STACK_KIND_SCALAR || tagStack[base + 3] == ValueTag.Error)) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        argc >= 4 && tagStack[base + 3] == ValueTag.Error ? valueStack[base + 3] : ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      )
    }
    if (argc >= 5 && (kindStack[base + 4] != STACK_KIND_SCALAR || tagStack[base + 4] == ValueTag.Error)) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        argc >= 5 && tagStack[base + 4] == ValueTag.Error ? valueStack[base + 4] : ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      )
    }
    if (argc == 6 && (kindStack[base + 5] != STACK_KIND_SCALAR || tagStack[base + 5] == ValueTag.Error)) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        argc == 6 && tagStack[base + 5] == ValueTag.Error ? valueStack[base + 5] : ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      )
    }

    const lookupRangeIndex = rangeIndexStack[base + 1]
    const returnRangeIndex = rangeIndexStack[base + 2]
    const length = <i32>rangeLengths[lookupRangeIndex]
    if (<i32>rangeLengths[returnRangeIndex] != length) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }

    const matchMode = argc >= 5 ? truncToInt(tagStack[base + 4], valueStack[base + 4]) : 0
    const searchMode = argc == 6 ? truncToInt(tagStack[base + 5], valueStack[base + 5]) : 1
    if (matchMode != 0 || !(searchMode == -1 || searchMode == 1)) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }

    const lookupStart = rangeOffsets[lookupRangeIndex]
    const returnStart = rangeOffsets[returnRangeIndex]
    if (searchMode == -1) {
      for (let index = length - 1; index >= 0; index--) {
        const lookupMemberIndex = rangeMembers[lookupStart + index]
        const comparison = compareScalarValues(
          cellTags[lookupMemberIndex],
          memberScalarValue(lookupMemberIndex, cellTags, cellNumbers, cellStringIds, cellErrors),
          tagStack[base],
          valueStack[base],
          null,
          stringOffsets,
          stringLengths,
          stringData,
          outputStringOffsets,
          outputStringLengths,
          outputStringData,
        )
        if (comparison == 0) {
          const returnMemberIndex = rangeMembers[returnStart + index]
          return writeMemberResult(
            base,
            returnMemberIndex,
            rangeIndexStack,
            valueStack,
            tagStack,
            kindStack,
            cellTags,
            cellNumbers,
            cellStringIds,
            cellErrors,
          )
        }
      }
    } else {
      for (let index = 0; index < length; index++) {
        const lookupMemberIndex = rangeMembers[lookupStart + index]
        const comparison = compareScalarValues(
          cellTags[lookupMemberIndex],
          memberScalarValue(lookupMemberIndex, cellTags, cellNumbers, cellStringIds, cellErrors),
          tagStack[base],
          valueStack[base],
          null,
          stringOffsets,
          stringLengths,
          stringData,
          outputStringOffsets,
          outputStringLengths,
          outputStringData,
        )
        if (comparison == 0) {
          const returnMemberIndex = rangeMembers[returnStart + index]
          return writeMemberResult(
            base,
            returnMemberIndex,
            rangeIndexStack,
            valueStack,
            tagStack,
            kindStack,
            cellTags,
            cellNumbers,
            cellStringIds,
            cellErrors,
          )
        }
      }
    }

    if (argc >= 4) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        tagStack[base + 3],
        valueStack[base + 3],
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      )
    }
    return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.NA, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.Lookup && (argc == 2 || argc == 3)) {
    if (kindStack[base] != STACK_KIND_SCALAR) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    if (
      (kindStack[base + 1] != STACK_KIND_SCALAR && kindStack[base + 1] != STACK_KIND_RANGE) ||
      (argc == 3 && kindStack[base + 2] != STACK_KIND_SCALAR && kindStack[base + 2] != STACK_KIND_RANGE)
    ) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    if (tagStack[base] == ValueTag.Error) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, valueStack[base], rangeIndexStack, valueStack, tagStack, kindStack)
    }
    if (kindStack[base + 1] == STACK_KIND_SCALAR && tagStack[base + 1] == ValueTag.Error) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        valueStack[base + 1],
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      )
    }
    if (argc == 3 && kindStack[base + 2] == STACK_KIND_SCALAR && tagStack[base + 2] == ValueTag.Error) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        valueStack[base + 2],
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      )
    }

    const lookupLength = vectorSlotLength(base + 1, kindStack, rangeIndexStack, rangeLengths, rangeRowCounts, rangeColCounts)
    const resultSlot = argc == 3 ? base + 2 : base + 1
    const resultLength = vectorSlotLength(resultSlot, kindStack, rangeIndexStack, rangeLengths, rangeRowCounts, rangeColCounts)
    if (lookupLength == i32.MIN_VALUE || resultLength == i32.MIN_VALUE || lookupLength != resultLength) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }

    let position = -1
    if (kindStack[base + 1] == STACK_KIND_SCALAR) {
      const exactComparison = compareScalarValues(
        tagStack[base + 1],
        valueStack[base + 1],
        tagStack[base],
        valueStack[base],
        null,
        stringOffsets,
        stringLengths,
        stringData,
        outputStringOffsets,
        outputStringLengths,
        outputStringData,
      )
      if (exactComparison == 0) {
        position = 1
      } else if (position < 0 && tagStack[base] == ValueTag.Number && exactComparison != i32.MIN_VALUE && exactComparison <= 0) {
        position = 1
      }
    } else {
      const lookupRangeIndex = rangeIndexStack[base + 1]
      const lookupStart = rangeOffsets[lookupRangeIndex]
      for (let index = 0; index < lookupLength; index++) {
        const memberIndex = rangeMembers[lookupStart + index]
        const comparison = compareScalarValues(
          cellTags[memberIndex],
          memberScalarValue(memberIndex, cellTags, cellNumbers, cellStringIds, cellErrors),
          tagStack[base],
          valueStack[base],
          null,
          stringOffsets,
          stringLengths,
          stringData,
          outputStringOffsets,
          outputStringLengths,
          outputStringData,
        )
        if (comparison == 0) {
          position = index + 1
          break
        }
      }
      if (position < 0 && tagStack[base] == ValueTag.Number) {
        let best = -1
        for (let index = 0; index < lookupLength; index++) {
          const memberIndex = rangeMembers[lookupStart + index]
          const comparison = compareScalarValues(
            cellTags[memberIndex],
            memberScalarValue(memberIndex, cellTags, cellNumbers, cellStringIds, cellErrors),
            tagStack[base],
            valueStack[base],
            null,
            stringOffsets,
            stringLengths,
            stringData,
            outputStringOffsets,
            outputStringLengths,
            outputStringData,
          )
          if (comparison == i32.MIN_VALUE) {
            best = -1
            break
          }
          if (comparison <= 0) {
            best = index + 1
            continue
          }
          break
        }
        position = best
      }
    }

    if (position < 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.NA, rangeIndexStack, valueStack, tagStack, kindStack)
    }

    if (kindStack[resultSlot] == STACK_KIND_SCALAR) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        tagStack[resultSlot],
        valueStack[resultSlot],
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      )
    }

    const resultRangeIndex = rangeIndexStack[resultSlot]
    const resultMemberIndex = rangeMembers[rangeOffsets[resultRangeIndex] + position - 1]
    return writeMemberResult(
      base,
      resultMemberIndex,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
      cellTags,
      cellNumbers,
      cellStringIds,
      cellErrors,
    )
  }

  return -1
}
