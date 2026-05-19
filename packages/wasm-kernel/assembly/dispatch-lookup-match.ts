import { BuiltinId, ErrorCode, ValueTag } from './protocol'
import { compareScalarValues } from './comparison'
import { inputCellScalarValue, inputCellTag, inputColsFromSlot, inputRowsFromSlot } from './operands'
import {
  lookupVectorSlotLength,
  slotVectorCol,
  slotVectorLength,
  slotVectorRow,
  writeLookupInputCellResult,
  writeLookupInputCellToSpill,
} from './lookup-slot'
import { compareLookupCell, compareLookupVectorCandidate } from './lookup-candidate-comparison'
import { truncToInt } from './numeric-core'
import { STACK_KIND_ARRAY, STACK_KIND_RANGE, STACK_KIND_SCALAR, writeArrayResult, writeResult } from './result-io'
import { allocateSpillArrayResult } from './vm'

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
    if (kindStack[base] != STACK_KIND_SCALAR) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const length = slotVectorLength(base + 1, kindStack, rangeIndexStack, rangeLengths, rangeRowCounts, rangeColCounts)
    if (length == i32.MIN_VALUE) {
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

    let best = -1
    for (let index = 0; index < length; index++) {
      const comparison = compareLookupVectorCandidate(
        base + 1,
        index,
        tagStack[base],
        valueStack[base],
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
    if (kindStack[base] != STACK_KIND_SCALAR) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const length = slotVectorLength(base + 1, kindStack, rangeIndexStack, rangeLengths, rangeRowCounts, rangeColCounts)
    if (length == i32.MIN_VALUE) {
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

    if (searchMode == -1) {
      if (matchMode == 0) {
        for (let index = length - 1; index >= 0; index--) {
          const comparison = compareLookupVectorCandidate(
            base + 1,
            index,
            tagStack[base],
            valueStack[base],
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
        const comparison = compareLookupVectorCandidate(
          base + 1,
          index,
          tagStack[base],
          valueStack[base],
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
        const comparison = compareLookupVectorCandidate(
          base + 1,
          index,
          tagStack[base],
          valueStack[base],
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
        if (comparison == 0) {
          return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, index + 1, rangeIndexStack, valueStack, tagStack, kindStack)
        }
      }
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.NA, rangeIndexStack, valueStack, tagStack, kindStack)
    }

    let best = -1
    for (let index = 0; index < length; index++) {
      const comparison = compareLookupVectorCandidate(
        base + 1,
        index,
        tagStack[base],
        valueStack[base],
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
    if (kindStack[base] != STACK_KIND_SCALAR) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const length = slotVectorLength(base + 1, kindStack, rangeIndexStack, rangeLengths, rangeRowCounts, rangeColCounts)
    if (length == i32.MIN_VALUE) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const returnRows = inputRowsFromSlot(base + 2, kindStack, rangeIndexStack, rangeRowCounts)
    const returnCols = inputColsFromSlot(base + 2, kindStack, rangeIndexStack, rangeColCounts)
    const lookupCols = inputColsFromSlot(base + 1, kindStack, rangeIndexStack, rangeColCounts)
    if (returnRows <= 0 || returnCols <= 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    let outputRows = 1
    let outputCols = 1
    if (lookupCols == 1) {
      if (returnRows != length) {
        return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      outputCols = returnCols
    } else {
      if (returnCols != length) {
        return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      outputRows = returnRows
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

    const matchMode = argc >= 5 ? truncToInt(tagStack[base + 4], valueStack[base + 4]) : 0
    const searchMode = argc == 6 ? truncToInt(tagStack[base + 5], valueStack[base + 5]) : 1
    if (!(matchMode == -1 || matchMode == 0 || matchMode == 1) || !(searchMode == -1 || searchMode == 1)) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }

    let matchedIndex = -1
    const first = searchMode == -1 ? length - 1 : 0
    const last = searchMode == -1 ? -1 : length
    const step = searchMode == -1 ? -1 : 1
    for (let index = first; index != last; index += step) {
      const comparison = compareLookupVectorCandidate(
        base + 1,
        index,
        tagStack[base],
        valueStack[base],
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
      if (comparison == 0) {
        matchedIndex = index
        break
      }
    }

    if (matchedIndex < 0 && matchMode != 0) {
      let bestTag: u8 = <u8>ValueTag.Empty
      let bestValue = 0.0
      for (let index = first; index != last; index += step) {
        const comparison = compareLookupVectorCandidate(
          base + 1,
          index,
          tagStack[base],
          valueStack[base],
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
        if (comparison == i32.MIN_VALUE) {
          continue
        }
        if ((matchMode == -1 && comparison >= 0) || (matchMode == 1 && comparison <= 0)) {
          continue
        }

        const row = slotVectorRow(base + 1, index, kindStack, rangeIndexStack, rangeRowCounts)
        const col = slotVectorCol(base + 1, index, kindStack, rangeIndexStack, rangeRowCounts)
        const candidateTag = inputCellTag(
          base + 1,
          row,
          col,
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
        )
        const candidateValue = inputCellScalarValue(
          base + 1,
          row,
          col,
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
        if (matchedIndex < 0) {
          matchedIndex = index
          bestTag = candidateTag
          bestValue = candidateValue
          continue
        }

        const bestComparison = compareScalarValues(
          candidateTag,
          candidateValue,
          bestTag,
          bestValue,
          null,
          stringOffsets,
          stringLengths,
          stringData,
          outputStringOffsets,
          outputStringLengths,
          outputStringData,
        )
        if (bestComparison == i32.MIN_VALUE) {
          continue
        }
        if ((matchMode == -1 && bestComparison > 0) || (matchMode == 1 && bestComparison < 0)) {
          matchedIndex = index
          bestTag = candidateTag
          bestValue = candidateValue
        }
      }
    }

    if (matchedIndex >= 0) {
      if (outputRows == 1 && outputCols == 1) {
        const resultRow = lookupCols == 1 ? matchedIndex : 0
        const resultCol = lookupCols == 1 ? 0 : matchedIndex
        return writeLookupInputCellResult(
          base,
          base + 2,
          resultRow,
          resultCol,
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

      const arrayIndex = allocateSpillArrayResult(outputRows, outputCols)
      let outputOffset = 0
      for (let row = 0; row < outputRows; row++) {
        for (let col = 0; col < outputCols; col++) {
          const sourceRow = lookupCols == 1 ? matchedIndex : row
          const sourceCol = lookupCols == 1 ? col : matchedIndex
          const copyError = writeLookupInputCellToSpill(
            arrayIndex,
            outputOffset,
            base + 2,
            sourceRow,
            sourceCol,
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
          if (copyError != ErrorCode.None) {
            return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, copyError, rangeIndexStack, valueStack, tagStack, kindStack)
          }
          outputOffset += 1
        }
      }
      return writeArrayResult(base, arrayIndex, outputRows, outputCols, rangeIndexStack, valueStack, tagStack, kindStack)
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
      (kindStack[base + 1] != STACK_KIND_SCALAR && kindStack[base + 1] != STACK_KIND_RANGE && kindStack[base + 1] != STACK_KIND_ARRAY) ||
      (argc == 3 &&
        kindStack[base + 2] != STACK_KIND_SCALAR &&
        kindStack[base + 2] != STACK_KIND_RANGE &&
        kindStack[base + 2] != STACK_KIND_ARRAY)
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

    const lookupLength = lookupVectorSlotLength(base + 1, kindStack, rangeIndexStack, rangeLengths, rangeRowCounts, rangeColCounts)
    const resultSlot = argc == 3 ? base + 2 : base + 1
    const resultLength = lookupVectorSlotLength(resultSlot, kindStack, rangeIndexStack, rangeLengths, rangeRowCounts, rangeColCounts)
    if (lookupLength == i32.MIN_VALUE || resultLength == i32.MIN_VALUE || lookupLength != resultLength) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }

    let position = -1
    const lookupRows = inputRowsFromSlot(base + 1, kindStack, rangeIndexStack, rangeRowCounts)
    for (let index = 0; index < lookupLength; index++) {
      const row = lookupRows == 1 ? 0 : index
      const col = lookupRows == 1 ? index : 0
      const comparison = compareLookupCell(
        base + 1,
        row,
        col,
        tagStack[base],
        valueStack[base],
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
      if (comparison == 0) {
        position = index + 1
        break
      }
    }
    if (position < 0 && tagStack[base] == ValueTag.Number) {
      let best = -1
      for (let index = 0; index < lookupLength; index++) {
        const row = lookupRows == 1 ? 0 : index
        const col = lookupRows == 1 ? index : 0
        const lookupTag = inputCellTag(
          base + 1,
          row,
          col,
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
        )
        const comparison = compareLookupCell(
          base + 1,
          row,
          col,
          tagStack[base],
          valueStack[base],
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
        if (comparison == i32.MIN_VALUE) {
          if (lookupTag == ValueTag.Error) {
            continue
          }
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

    const resultRows = inputRowsFromSlot(resultSlot, kindStack, rangeIndexStack, rangeRowCounts)
    const resultIndex = position - 1
    const resultRow = resultRows == 1 ? 0 : resultIndex
    const resultCol = resultRows == 1 ? resultIndex : 0
    const resultTag = inputCellTag(
      resultSlot,
      resultRow,
      resultCol,
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
    )
    const resultValue = inputCellScalarValue(
      resultSlot,
      resultRow,
      resultCol,
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
    return writeResult(base, STACK_KIND_SCALAR, resultTag, resultValue, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  return -1
}
