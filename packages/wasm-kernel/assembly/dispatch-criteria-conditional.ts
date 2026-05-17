import {
  criteriaSlotLength,
  criteriaSlotMatches,
  criteriaSlotNumberOnlyOrNaN,
  criteriaSlotNumberOrNaN,
  criteriaSlotNumberOrZero,
  isCriteriaRangeSlot,
} from './aggregate-criteria-slot'
import { BuiltinId, ErrorCode, ValueTag } from './protocol'
import { STACK_KIND_SCALAR, writeResult } from './result-io'

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

export function tryApplyCriteriaConditionalBuiltin(
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
  if (builtinId == BuiltinId.Countif && argc == 2) {
    if (!isCriteriaRangeSlot(kindStack[base]) || kindStack[base + 1] != STACK_KIND_SCALAR) {
      return writeAggregateError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    if (tagStack[base + 1] == ValueTag.Error) {
      return writeAggregateError(base, <i32>valueStack[base + 1], rangeIndexStack, valueStack, tagStack, kindStack)
    }

    const length = criteriaSlotLength(base, kindStack, rangeIndexStack, rangeRowCounts, rangeColCounts)
    if (length == i32.MIN_VALUE) {
      return writeAggregateError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    let count = 0
    for (let cursor = 0; cursor < length; cursor += 1) {
      if (
        criteriaSlotMatches(
          base,
          cursor,
          tagStack[base + 1],
          valueStack[base + 1],
          kindStack,
          valueStack,
          tagStack,
          rangeIndexStack,
          rangeOffsets,
          rangeLengths,
          rangeRowCounts,
          rangeColCounts,
          rangeMembers,
          cellTags,
          cellNumbers,
          cellStringIds,
          cellErrors,
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

    if (!isCriteriaRangeSlot(kindStack[base])) {
      return writeAggregateError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const expectedLength = criteriaSlotLength(base, kindStack, rangeIndexStack, rangeRowCounts, rangeColCounts)
    if (expectedLength == i32.MIN_VALUE) {
      return writeAggregateError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    for (let index = 0; index < argc; index += 2) {
      const rangeSlot = base + index
      const criteriaSlot = rangeSlot + 1
      if (!isCriteriaRangeSlot(kindStack[rangeSlot]) || kindStack[criteriaSlot] != STACK_KIND_SCALAR) {
        return writeAggregateError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      if (tagStack[criteriaSlot] == ValueTag.Error) {
        return writeAggregateError(base, <i32>valueStack[criteriaSlot], rangeIndexStack, valueStack, tagStack, kindStack)
      }
      if (criteriaSlotLength(rangeSlot, kindStack, rangeIndexStack, rangeRowCounts, rangeColCounts) != expectedLength) {
        return writeAggregateError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      }
    }

    let count = 0
    for (let row = 0; row < expectedLength; row += 1) {
      let matchesAll = true
      for (let index = 0; index < argc; index += 2) {
        const rangeSlot = base + index
        const criteriaSlot = rangeSlot + 1
        if (
          !criteriaSlotMatches(
            rangeSlot,
            row,
            tagStack[criteriaSlot],
            valueStack[criteriaSlot],
            kindStack,
            valueStack,
            tagStack,
            rangeIndexStack,
            rangeOffsets,
            rangeLengths,
            rangeRowCounts,
            rangeColCounts,
            rangeMembers,
            cellTags,
            cellNumbers,
            cellStringIds,
            cellErrors,
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
      !isCriteriaRangeSlot(kindStack[rangeSlot]) ||
      kindStack[criteriaSlot] != STACK_KIND_SCALAR ||
      !isCriteriaRangeSlot(kindStack[sumRangeSlot])
    ) {
      return writeAggregateError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    if (tagStack[criteriaSlot] == ValueTag.Error) {
      return writeAggregateError(base, <i32>valueStack[criteriaSlot], rangeIndexStack, valueStack, tagStack, kindStack)
    }

    const length = criteriaSlotLength(rangeSlot, kindStack, rangeIndexStack, rangeRowCounts, rangeColCounts)
    if (length == i32.MIN_VALUE || criteriaSlotLength(sumRangeSlot, kindStack, rangeIndexStack, rangeRowCounts, rangeColCounts) != length) {
      return writeAggregateError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }

    let sum = 0.0
    for (let cursor = 0; cursor < length; cursor += 1) {
      if (
        !criteriaSlotMatches(
          rangeSlot,
          cursor,
          tagStack[criteriaSlot],
          valueStack[criteriaSlot],
          kindStack,
          valueStack,
          tagStack,
          rangeIndexStack,
          rangeOffsets,
          rangeLengths,
          rangeRowCounts,
          rangeColCounts,
          rangeMembers,
          cellTags,
          cellNumbers,
          cellStringIds,
          cellErrors,
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
      sum += criteriaSlotNumberOrZero(
        sumRangeSlot,
        cursor,
        kindStack,
        valueStack,
        tagStack,
        rangeIndexStack,
        rangeOffsets,
        rangeLengths,
        rangeRowCounts,
        rangeColCounts,
        rangeMembers,
        cellTags,
        cellNumbers,
        cellStringIds,
        cellErrors,
      )
    }
    return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, sum, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.Sumifs) {
    if (argc < 3 || argc % 2 == 0 || !isCriteriaRangeSlot(kindStack[base])) {
      return writeAggregateError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }

    const expectedLength = criteriaSlotLength(base, kindStack, rangeIndexStack, rangeRowCounts, rangeColCounts)
    if (expectedLength == i32.MIN_VALUE) {
      return writeAggregateError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    for (let index = 1; index < argc; index += 2) {
      const rangeSlot = base + index
      const criteriaSlot = rangeSlot + 1
      if (!isCriteriaRangeSlot(kindStack[rangeSlot]) || kindStack[criteriaSlot] != STACK_KIND_SCALAR) {
        return writeAggregateError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      if (tagStack[criteriaSlot] == ValueTag.Error) {
        return writeAggregateError(base, <i32>valueStack[criteriaSlot], rangeIndexStack, valueStack, tagStack, kindStack)
      }
      if (criteriaSlotLength(rangeSlot, kindStack, rangeIndexStack, rangeRowCounts, rangeColCounts) != expectedLength) {
        return writeAggregateError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      }
    }

    let sum = 0.0
    for (let row = 0; row < expectedLength; row += 1) {
      let matchesAll = true
      for (let index = 1; index < argc; index += 2) {
        const rangeSlot = base + index
        const criteriaSlot = rangeSlot + 1
        if (
          !criteriaSlotMatches(
            rangeSlot,
            row,
            tagStack[criteriaSlot],
            valueStack[criteriaSlot],
            kindStack,
            valueStack,
            tagStack,
            rangeIndexStack,
            rangeOffsets,
            rangeLengths,
            rangeRowCounts,
            rangeColCounts,
            rangeMembers,
            cellTags,
            cellNumbers,
            cellStringIds,
            cellErrors,
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
      sum += criteriaSlotNumberOrZero(
        base,
        row,
        kindStack,
        valueStack,
        tagStack,
        rangeIndexStack,
        rangeOffsets,
        rangeLengths,
        rangeRowCounts,
        rangeColCounts,
        rangeMembers,
        cellTags,
        cellNumbers,
        cellStringIds,
        cellErrors,
      )
    }
    return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, sum, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.Averageif && (argc == 2 || argc == 3)) {
    const rangeSlot = base
    const criteriaSlot = base + 1
    const averageRangeSlot = argc == 3 ? base + 2 : base
    if (
      !isCriteriaRangeSlot(kindStack[rangeSlot]) ||
      kindStack[criteriaSlot] != STACK_KIND_SCALAR ||
      !isCriteriaRangeSlot(kindStack[averageRangeSlot])
    ) {
      return writeAggregateError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    if (tagStack[criteriaSlot] == ValueTag.Error) {
      return writeAggregateError(base, <i32>valueStack[criteriaSlot], rangeIndexStack, valueStack, tagStack, kindStack)
    }

    const length = criteriaSlotLength(rangeSlot, kindStack, rangeIndexStack, rangeRowCounts, rangeColCounts)
    if (
      length == i32.MIN_VALUE ||
      criteriaSlotLength(averageRangeSlot, kindStack, rangeIndexStack, rangeRowCounts, rangeColCounts) != length
    ) {
      return writeAggregateError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }

    let count = 0
    let sum = 0.0
    for (let cursor = 0; cursor < length; cursor += 1) {
      if (
        !criteriaSlotMatches(
          rangeSlot,
          cursor,
          tagStack[criteriaSlot],
          valueStack[criteriaSlot],
          kindStack,
          valueStack,
          tagStack,
          rangeIndexStack,
          rangeOffsets,
          rangeLengths,
          rangeRowCounts,
          rangeColCounts,
          rangeMembers,
          cellTags,
          cellNumbers,
          cellStringIds,
          cellErrors,
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
      const numeric = criteriaSlotNumberOrNaN(
        averageRangeSlot,
        cursor,
        kindStack,
        valueStack,
        tagStack,
        rangeIndexStack,
        rangeOffsets,
        rangeLengths,
        rangeRowCounts,
        rangeColCounts,
        rangeMembers,
        cellTags,
        cellNumbers,
        cellStringIds,
        cellErrors,
      )
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
    if (argc < 3 || argc % 2 == 0 || !isCriteriaRangeSlot(kindStack[base])) {
      return writeAggregateError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }

    const expectedLength = criteriaSlotLength(base, kindStack, rangeIndexStack, rangeRowCounts, rangeColCounts)
    if (expectedLength == i32.MIN_VALUE) {
      return writeAggregateError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    for (let index = 1; index < argc; index += 2) {
      const rangeSlot = base + index
      const criteriaSlot = rangeSlot + 1
      if (!isCriteriaRangeSlot(kindStack[rangeSlot]) || kindStack[criteriaSlot] != STACK_KIND_SCALAR) {
        return writeAggregateError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      if (tagStack[criteriaSlot] == ValueTag.Error) {
        return writeAggregateError(base, <i32>valueStack[criteriaSlot], rangeIndexStack, valueStack, tagStack, kindStack)
      }
      if (criteriaSlotLength(rangeSlot, kindStack, rangeIndexStack, rangeRowCounts, rangeColCounts) != expectedLength) {
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
        if (
          !criteriaSlotMatches(
            rangeSlot,
            row,
            tagStack[criteriaSlot],
            valueStack[criteriaSlot],
            kindStack,
            valueStack,
            tagStack,
            rangeIndexStack,
            rangeOffsets,
            rangeLengths,
            rangeRowCounts,
            rangeColCounts,
            rangeMembers,
            cellTags,
            cellNumbers,
            cellStringIds,
            cellErrors,
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
      const numeric = criteriaSlotNumberOrNaN(
        base,
        row,
        kindStack,
        valueStack,
        tagStack,
        rangeIndexStack,
        rangeOffsets,
        rangeLengths,
        rangeRowCounts,
        rangeColCounts,
        rangeMembers,
        cellTags,
        cellNumbers,
        cellStringIds,
        cellErrors,
      )
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
    if (argc < 3 || argc % 2 == 0 || !isCriteriaRangeSlot(kindStack[base])) {
      return writeAggregateError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }

    const expectedLength = criteriaSlotLength(base, kindStack, rangeIndexStack, rangeRowCounts, rangeColCounts)
    if (expectedLength == i32.MIN_VALUE) {
      return writeAggregateError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    for (let index = 1; index < argc; index += 2) {
      const rangeSlot = base + index
      const criteriaSlot = rangeSlot + 1
      if (!isCriteriaRangeSlot(kindStack[rangeSlot]) || kindStack[criteriaSlot] != STACK_KIND_SCALAR) {
        return writeAggregateError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      if (tagStack[criteriaSlot] == ValueTag.Error) {
        return writeAggregateError(base, <i32>valueStack[criteriaSlot], rangeIndexStack, valueStack, tagStack, kindStack)
      }
      if (criteriaSlotLength(rangeSlot, kindStack, rangeIndexStack, rangeRowCounts, rangeColCounts) != expectedLength) {
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
        if (
          !criteriaSlotMatches(
            rangeSlot,
            row,
            tagStack[criteriaSlot],
            valueStack[criteriaSlot],
            kindStack,
            valueStack,
            tagStack,
            rangeIndexStack,
            rangeOffsets,
            rangeLengths,
            rangeRowCounts,
            rangeColCounts,
            rangeMembers,
            cellTags,
            cellNumbers,
            cellStringIds,
            cellErrors,
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
      const numeric = criteriaSlotNumberOnlyOrNaN(
        base,
        row,
        kindStack,
        valueStack,
        tagStack,
        rangeIndexStack,
        rangeOffsets,
        rangeLengths,
        rangeRowCounts,
        rangeColCounts,
        rangeMembers,
        cellTags,
        cellNumbers,
        cellStringIds,
        cellErrors,
      )
      if (isNaN(numeric)) {
        continue
      }
      result = builtinId == BuiltinId.Minifs ? min(result, numeric) : max(result, numeric)
      found = true
    }
    return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, found ? result : 0.0, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  return -1
}
