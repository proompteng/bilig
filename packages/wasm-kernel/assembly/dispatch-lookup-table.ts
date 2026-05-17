import { BuiltinId, ErrorCode, ValueTag } from './protocol'
import { copyInputCellToResult, copyInputCellToSpill } from './array-materialize'
import { compareScalarValues } from './comparison'
import { inputCellScalarValue, inputCellTag, inputColsFromSlot, inputRowsFromSlot } from './operands'
import { writeLookupInputCellResult } from './lookup-slot'
import { truncToInt } from './numeric-core'
import { STACK_KIND_ARRAY, STACK_KIND_RANGE, STACK_KIND_SCALAR, writeArrayResult, writeResult } from './result-io'
import { allocateSpillArrayResult } from './vm'

function coerceLookupBoolean(tag: u8, value: f64): i32 {
  if (tag == ValueTag.Boolean || tag == ValueTag.Number) {
    return value != 0 ? 1 : 0
  }
  if (tag == ValueTag.Empty) {
    return 0
  }
  return -1
}

function writeLookupError(
  base: i32,
  error: i32,
  rangeIndexStack: Uint32Array,
  valueStack: Float64Array,
  tagStack: Uint8Array,
  kindStack: Uint8Array,
): i32 {
  return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, error, rangeIndexStack, valueStack, tagStack, kindStack)
}

function isTableSlot(kind: u8): bool {
  return kind == STACK_KIND_RANGE || kind == STACK_KIND_ARRAY
}

export function tryApplyLookupTableBuiltin(
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
  if (builtinId == BuiltinId.Index && (argc == 2 || argc == 3)) {
    if (!isTableSlot(kindStack[base]) || kindStack[base + 1] != STACK_KIND_SCALAR) {
      return writeLookupError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    if (tagStack[base + 1] == ValueTag.Error) {
      return writeLookupError(base, <i32>valueStack[base + 1], rangeIndexStack, valueStack, tagStack, kindStack)
    }
    if (argc == 3 && (kindStack[base + 2] != STACK_KIND_SCALAR || tagStack[base + 2] == ValueTag.Error)) {
      return writeLookupError(
        base,
        argc == 3 && tagStack[base + 2] == ValueTag.Error ? <i32>valueStack[base + 2] : ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      )
    }

    const rowCount = inputRowsFromSlot(base, kindStack, rangeIndexStack, rangeRowCounts)
    const colCount = inputColsFromSlot(base, kindStack, rangeIndexStack, rangeColCounts)
    const rawRowNum = truncToInt(tagStack[base + 1], valueStack[base + 1])
    const colNumOmitted = argc != 3
    const rawColNum = colNumOmitted ? 0 : truncToInt(tagStack[base + 2], valueStack[base + 2])
    if (rowCount <= 0 || colCount <= 0 || rawRowNum == i32.MIN_VALUE || rawColNum == i32.MIN_VALUE) {
      return writeLookupError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }

    let rowNum = rawRowNum
    let colNum = rawColNum
    if (rowCount == 1 && colNumOmitted && rawRowNum != 0) {
      rowNum = 1
      colNum = rawRowNum
    }
    if (colCount == 1 && colNumOmitted && rawRowNum != 0) {
      colNum = 1
    }
    if (rowNum < 0 || colNum < 0 || rowNum > rowCount || colNum > colCount) {
      return writeLookupError(base, ErrorCode.Ref, rangeIndexStack, valueStack, tagStack, kindStack)
    }

    if (rowNum == 0 || colNum == 0) {
      const outputRows = rowNum == 0 ? rowCount : 1
      const outputCols = colNum == 0 ? colCount : 1
      const arrayIndex = allocateSpillArrayResult(outputRows, outputCols)
      let outputOffset = 0
      for (let row = 0; row < outputRows; row++) {
        for (let col = 0; col < outputCols; col++) {
          const sourceRow = rowNum == 0 ? row : rowNum - 1
          const sourceCol = colNum == 0 ? col : colNum - 1
          const copyError = copyInputCellToSpill(
            arrayIndex,
            outputOffset,
            base,
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
            return writeLookupError(base, copyError, rangeIndexStack, valueStack, tagStack, kindStack)
          }
          outputOffset += 1
        }
      }
      return writeArrayResult(base, arrayIndex, outputRows, outputCols, rangeIndexStack, valueStack, tagStack, kindStack)
    }

    return copyInputCellToResult(
      base,
      base,
      rowNum - 1,
      colNum - 1,
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

  if (builtinId == BuiltinId.Vlookup && (argc == 3 || argc == 4)) {
    if (kindStack[base] != STACK_KIND_SCALAR || !isTableSlot(kindStack[base + 1]) || kindStack[base + 2] != STACK_KIND_SCALAR) {
      return writeLookupError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    if (tagStack[base] == ValueTag.Error) {
      return writeLookupError(base, <i32>valueStack[base], rangeIndexStack, valueStack, tagStack, kindStack)
    }
    if (tagStack[base + 2] == ValueTag.Error) {
      return writeLookupError(base, <i32>valueStack[base + 2], rangeIndexStack, valueStack, tagStack, kindStack)
    }
    if (argc == 4 && (kindStack[base + 3] != STACK_KIND_SCALAR || tagStack[base + 3] == ValueTag.Error)) {
      return writeLookupError(
        base,
        argc == 4 && tagStack[base + 3] == ValueTag.Error ? <i32>valueStack[base + 3] : ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      )
    }

    const rowCount = inputRowsFromSlot(base + 1, kindStack, rangeIndexStack, rangeRowCounts)
    const colCount = inputColsFromSlot(base + 1, kindStack, rangeIndexStack, rangeColCounts)
    const colIndex = truncToInt(tagStack[base + 2], valueStack[base + 2])
    const rangeLookup = argc == 4 ? coerceLookupBoolean(tagStack[base + 3], valueStack[base + 3]) : 1
    if (rowCount <= 0 || colCount <= 0 || colIndex == i32.MIN_VALUE || colIndex < 1 || colIndex > colCount || rangeLookup < 0) {
      return writeLookupError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }

    let matchedRow = -1
    for (let row = 0; row < rowCount; row += 1) {
      const candidateTag = inputCellTag(
        base + 1,
        row,
        0,
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
        0,
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
      const comparison = compareScalarValues(
        candidateTag,
        candidateValue,
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
        if (rangeLookup == 0) {
          continue
        }
        return writeLookupError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      if (comparison == 0) {
        matchedRow = row
        break
      }
      if (rangeLookup == 1 && comparison < 0) {
        matchedRow = row
        continue
      }
      if (rangeLookup == 1 && comparison > 0) {
        break
      }
    }

    if (matchedRow < 0) {
      return writeLookupError(base, ErrorCode.NA, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    return writeLookupInputCellResult(
      base,
      base + 1,
      matchedRow,
      colIndex - 1,
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

  if (builtinId == BuiltinId.Hlookup && (argc == 3 || argc == 4)) {
    if (kindStack[base] != STACK_KIND_SCALAR || !isTableSlot(kindStack[base + 1]) || kindStack[base + 2] != STACK_KIND_SCALAR) {
      return writeLookupError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    if (tagStack[base] == ValueTag.Error) {
      return writeLookupError(base, <i32>valueStack[base], rangeIndexStack, valueStack, tagStack, kindStack)
    }
    if (tagStack[base + 2] == ValueTag.Error) {
      return writeLookupError(base, <i32>valueStack[base + 2], rangeIndexStack, valueStack, tagStack, kindStack)
    }
    if (argc == 4 && (kindStack[base + 3] != STACK_KIND_SCALAR || tagStack[base + 3] == ValueTag.Error)) {
      return writeLookupError(
        base,
        argc == 4 && tagStack[base + 3] == ValueTag.Error ? <i32>valueStack[base + 3] : ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      )
    }

    const rowCount = inputRowsFromSlot(base + 1, kindStack, rangeIndexStack, rangeRowCounts)
    const colCount = inputColsFromSlot(base + 1, kindStack, rangeIndexStack, rangeColCounts)
    const rowIndex = truncToInt(tagStack[base + 2], valueStack[base + 2])
    const rangeLookup = argc == 4 ? coerceLookupBoolean(tagStack[base + 3], valueStack[base + 3]) : 1
    if (rowCount <= 0 || colCount <= 0 || rowIndex == i32.MIN_VALUE || rowIndex < 1 || rowIndex > rowCount || rangeLookup < 0) {
      return writeLookupError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }

    let matchedCol = -1
    for (let col = 0; col < colCount; col += 1) {
      const candidateTag = inputCellTag(
        base + 1,
        0,
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
        0,
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
      const comparison = compareScalarValues(
        candidateTag,
        candidateValue,
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
        if (rangeLookup == 0) {
          continue
        }
        return writeLookupError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      if (comparison == 0) {
        matchedCol = col
        break
      }
      if (rangeLookup == 1 && comparison < 0) {
        matchedCol = col
        continue
      }
      if (rangeLookup == 1 && comparison > 0) {
        break
      }
    }

    if (matchedCol < 0) {
      return writeLookupError(base, ErrorCode.NA, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    return writeLookupInputCellResult(
      base,
      base + 1,
      rowIndex - 1,
      matchedCol,
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

  return -1
}
