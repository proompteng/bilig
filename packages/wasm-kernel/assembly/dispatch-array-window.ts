import { copyInputCellToResult, copyInputCellToSpill } from './array-materialize'
import { coerceInteger, truncToInt } from './numeric-core'
import { inputCellTag, inputColsFromSlot, inputRowsFromSlot, toNumberExact } from './operands'
import { BuiltinId, ErrorCode, ValueTag } from './protocol'
import { coercePositiveIntegerArg, scalarErrorAt } from './builtin-args'
import { STACK_KIND_ARRAY, STACK_KIND_RANGE, STACK_KIND_SCALAR, writeArrayResult, writeResult } from './result-io'
import { allocateSpillArrayResult, writeSpillArrayValue } from './vm'

function coerceTrimMode(tag: u8, value: f64): i32 {
  const numeric = toNumberExact(tag, value)
  if (!isFinite(numeric)) {
    return i32.MIN_VALUE
  }
  const integer = <i32>numeric
  return integer >= 0 && integer <= 3 ? integer : i32.MIN_VALUE
}

function clipIndex(index: i32, length: i32): i32 {
  if (length <= 0) {
    return i32.MIN_VALUE
  }
  if (index == 0) {
    return i32.MIN_VALUE
  }
  return index < 0 ? max(index, -length) : min(index, length)
}

export function tryApplyArrayWindowBuiltin(
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
  rangeOffsets: Uint32Array,
  rangeLengths: Uint32Array,
  rangeRowCounts: Uint32Array,
  rangeColCounts: Uint32Array,
  rangeMembers: Uint32Array,
): i32 {
  if (builtinId == BuiltinId.Offset && argc >= 3 && argc <= 5) {
    const sourceKind = kindStack[base]
    if (sourceKind != STACK_KIND_SCALAR && sourceKind != STACK_KIND_RANGE && sourceKind != STACK_KIND_ARRAY) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack)
    if (scalarError >= 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, scalarError, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const sourceRows = inputRowsFromSlot(base, kindStack, rangeIndexStack, rangeRowCounts)
    const sourceCols = inputColsFromSlot(base, kindStack, rangeIndexStack, rangeColCounts)
    if (sourceRows <= 0 || sourceCols <= 0 || sourceRows == i32.MIN_VALUE || sourceCols == i32.MIN_VALUE) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const rowOffset = truncToInt(tagStack[base + 1], valueStack[base + 1])
    const colOffset = truncToInt(tagStack[base + 2], valueStack[base + 2])
    const height = argc >= 4 ? coercePositiveIntegerArg(tagStack[base + 3], valueStack[base + 3], true, sourceRows) : sourceRows
    const width = argc >= 5 ? coercePositiveIntegerArg(tagStack[base + 4], valueStack[base + 4], true, sourceCols) : sourceCols
    const areaNumber = argc >= 6 ? truncToInt(tagStack[base + 5], valueStack[base + 5]) : 1
    if (
      rowOffset == i32.MIN_VALUE ||
      colOffset == i32.MIN_VALUE ||
      height == i32.MIN_VALUE ||
      width == i32.MIN_VALUE ||
      areaNumber == i32.MIN_VALUE ||
      areaNumber != 1 ||
      height < 1 ||
      width < 1
    ) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }

    const rowStart = rowOffset < 0 ? sourceRows + rowOffset : rowOffset
    const colStart = colOffset < 0 ? sourceCols + colOffset : colOffset
    if (rowStart < 0 || colStart < 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Ref, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    if (rowStart + height > sourceRows || colStart + width > sourceCols) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Ref, rangeIndexStack, valueStack, tagStack, kindStack)
    }

    if (height == 1 && width == 1) {
      return copyInputCellToResult(
        base,
        base,
        rowStart,
        colStart,
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

    const arrayIndex = allocateSpillArrayResult(height, width)
    for (let row = 0; row < height; row++) {
      for (let col = 0; col < width; col++) {
        const copyError = copyInputCellToSpill(
          arrayIndex,
          row * width + col,
          base,
          rowStart + row,
          colStart + col,
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
      }
    }
    return writeArrayResult(base, arrayIndex, height, width, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.Take && argc >= 1 && argc <= 3) {
    const sourceKind = kindStack[base]
    if (sourceKind != STACK_KIND_SCALAR && sourceKind != STACK_KIND_RANGE && sourceKind != STACK_KIND_ARRAY) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const sourceRows = inputRowsFromSlot(base, kindStack, rangeIndexStack, rangeRowCounts)
    const sourceCols = inputColsFromSlot(base, kindStack, rangeIndexStack, rangeColCounts)
    if (sourceRows <= 0 || sourceCols <= 0 || sourceRows == i32.MIN_VALUE || sourceCols == i32.MIN_VALUE) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack)
    if (scalarError >= 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, scalarError, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const requestedRows = argc >= 2 ? coerceInteger(tagStack[base + 1], valueStack[base + 1]) : sourceRows
    const requestedCols = argc >= 3 ? coerceInteger(tagStack[base + 2], valueStack[base + 2]) : sourceCols
    if ((argc >= 2 && requestedRows == i32.MIN_VALUE) || (argc >= 3 && requestedCols == i32.MIN_VALUE)) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const clippedRows = argc >= 2 ? clipIndex(requestedRows, sourceRows) : sourceRows
    const clippedCols = argc >= 3 ? clipIndex(requestedCols, sourceCols) : sourceCols
    if (clippedRows == i32.MIN_VALUE || clippedCols == i32.MIN_VALUE) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const rowCount = clippedRows > 0 ? min<i32>(clippedRows, sourceRows) : min<i32>(-clippedRows, sourceRows)
    const colCount = clippedCols > 0 ? min<i32>(clippedCols, sourceCols) : min<i32>(-clippedCols, sourceCols)
    if (rowCount == 0 || colCount == 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const rowOffset = clippedRows > 0 ? 0 : max<i32>(sourceRows - rowCount, 0)
    const colOffset = clippedCols > 0 ? 0 : max<i32>(sourceCols - colCount, 0)

    const arrayIndex = allocateSpillArrayResult(rowCount, colCount)
    let outputOffset = 0
    for (let row = 0; row < rowCount; row++) {
      for (let col = 0; col < colCount; col++) {
        const copyError = copyInputCellToSpill(
          arrayIndex,
          outputOffset,
          base,
          rowOffset + row,
          colOffset + col,
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
    return writeArrayResult(base, arrayIndex, rowCount, colCount, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.Drop && argc >= 1 && argc <= 3) {
    const sourceKind = kindStack[base]
    if (sourceKind != STACK_KIND_SCALAR && sourceKind != STACK_KIND_RANGE && sourceKind != STACK_KIND_ARRAY) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const sourceRows = inputRowsFromSlot(base, kindStack, rangeIndexStack, rangeRowCounts)
    const sourceCols = inputColsFromSlot(base, kindStack, rangeIndexStack, rangeColCounts)
    if (sourceRows <= 0 || sourceCols <= 0 || sourceRows == i32.MIN_VALUE || sourceCols == i32.MIN_VALUE) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack)
    if (scalarError >= 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, scalarError, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const requestedRows = argc >= 2 ? coerceInteger(tagStack[base + 1], valueStack[base + 1]) : 0
    const requestedCols = argc >= 3 ? coerceInteger(tagStack[base + 2], valueStack[base + 2]) : 0
    if ((argc >= 2 && requestedRows == i32.MIN_VALUE) || (argc >= 3 && requestedCols == i32.MIN_VALUE)) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const clippedRows = argc >= 2 ? (requestedRows == 0 ? 0 : clipIndex(requestedRows, sourceRows)) : 0
    const clippedCols = argc >= 3 ? (requestedCols == 0 ? 0 : clipIndex(requestedCols, sourceCols)) : 0
    if (clippedRows == i32.MIN_VALUE || clippedCols == i32.MIN_VALUE) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const rowDrop = clippedRows >= 0 ? min<i32>(clippedRows, sourceRows) : min<i32>(-clippedRows, sourceRows)
    const colDrop = clippedCols >= 0 ? min<i32>(clippedCols, sourceCols) : min<i32>(-clippedCols, sourceCols)
    const rowCount = sourceRows - rowDrop
    const colCount = sourceCols - colDrop
    if (rowCount <= 0 || colCount <= 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const rowOffset = clippedRows > 0 ? rowDrop : 0
    const colOffset = clippedCols > 0 ? colDrop : 0

    const arrayIndex = allocateSpillArrayResult(rowCount, colCount)
    let outputOffset = 0
    for (let row = 0; row < rowCount; row++) {
      for (let col = 0; col < colCount; col++) {
        const copyError = copyInputCellToSpill(
          arrayIndex,
          outputOffset,
          base,
          rowOffset + row,
          colOffset + col,
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
    return writeArrayResult(base, arrayIndex, rowCount, colCount, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.Expand && argc >= 2 && argc <= 4) {
    const sourceKind = kindStack[base]
    if (sourceKind != STACK_KIND_SCALAR && sourceKind != STACK_KIND_RANGE && sourceKind != STACK_KIND_ARRAY) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const sourceRows = inputRowsFromSlot(base, kindStack, rangeIndexStack, rangeRowCounts)
    const sourceCols = inputColsFromSlot(base, kindStack, rangeIndexStack, rangeColCounts)
    if (sourceRows <= 0 || sourceCols <= 0 || sourceRows == i32.MIN_VALUE || sourceCols == i32.MIN_VALUE) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    for (let arg = 1; arg < argc; arg++) {
      const slot = base + arg
      if (kindStack[slot] != STACK_KIND_SCALAR) {
        return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      }
    }
    if (tagStack[base + 1] == ValueTag.Error) {
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
    if (argc >= 3 && tagStack[base + 2] == ValueTag.Error) {
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
    const targetRows = coerceInteger(tagStack[base + 1], valueStack[base + 1])
    const targetCols = argc >= 3 ? coerceInteger(tagStack[base + 2], valueStack[base + 2]) : sourceCols
    if (
      targetRows == i32.MIN_VALUE ||
      targetCols == i32.MIN_VALUE ||
      targetRows < sourceRows ||
      targetCols < sourceCols ||
      targetRows < 1 ||
      targetCols < 1
    ) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const padTag = argc >= 4 ? tagStack[base + 3] : <u8>ValueTag.Error
    const padValue = argc >= 4 ? valueStack[base + 3] : ErrorCode.NA
    const arrayIndex = allocateSpillArrayResult(targetRows, targetCols)
    let outputOffset = 0
    for (let row = 0; row < targetRows; row++) {
      for (let col = 0; col < targetCols; col++) {
        if (row < sourceRows && col < sourceCols) {
          const copyError = copyInputCellToSpill(
            arrayIndex,
            outputOffset,
            base,
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
          if (copyError != ErrorCode.None) {
            return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, copyError, rangeIndexStack, valueStack, tagStack, kindStack)
          }
        } else {
          writeSpillArrayValue(arrayIndex, outputOffset, padTag, padValue)
        }
        outputOffset += 1
      }
    }
    return writeArrayResult(base, arrayIndex, targetRows, targetCols, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.Trimrange && argc >= 1 && argc <= 3) {
    const sourceKind = kindStack[base]
    if (sourceKind != STACK_KIND_SCALAR && sourceKind != STACK_KIND_RANGE && sourceKind != STACK_KIND_ARRAY) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const sourceRows = inputRowsFromSlot(base, kindStack, rangeIndexStack, rangeRowCounts)
    const sourceCols = inputColsFromSlot(base, kindStack, rangeIndexStack, rangeColCounts)
    if (sourceRows <= 0 || sourceCols <= 0 || sourceRows == i32.MIN_VALUE || sourceCols == i32.MIN_VALUE) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    for (let arg = 1; arg < argc; arg++) {
      const slot = base + arg
      if (kindStack[slot] != STACK_KIND_SCALAR) {
        return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      if (tagStack[slot] == ValueTag.Error) {
        return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, valueStack[slot], rangeIndexStack, valueStack, tagStack, kindStack)
      }
    }
    const trimRows = argc >= 2 ? coerceTrimMode(tagStack[base + 1], valueStack[base + 1]) : 3
    const trimCols = argc >= 3 ? coerceTrimMode(tagStack[base + 2], valueStack[base + 2]) : 3
    if (trimRows == i32.MIN_VALUE || trimCols == i32.MIN_VALUE) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }

    let startRow = 0
    let endRow = sourceRows - 1
    let startCol = 0
    let endCol = sourceCols - 1

    const trimLeadingRows = trimRows == 1 || trimRows == 3
    const trimTrailingRows = trimRows == 2 || trimRows == 3
    const trimLeadingCols = trimCols == 1 || trimCols == 3
    const trimTrailingCols = trimCols == 2 || trimCols == 3

    if (trimLeadingRows) {
      while (startRow <= endRow) {
        let hasNonEmpty = false
        for (let col = 0; col < sourceCols; col++) {
          if (
            inputCellTag(
              base,
              startRow,
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
            ) != ValueTag.Empty
          ) {
            hasNonEmpty = true
            break
          }
        }
        if (hasNonEmpty) {
          break
        }
        startRow += 1
      }
    }

    if (trimTrailingRows) {
      while (endRow >= startRow) {
        let hasNonEmpty = false
        for (let col = 0; col < sourceCols; col++) {
          if (
            inputCellTag(
              base,
              endRow,
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
            ) != ValueTag.Empty
          ) {
            hasNonEmpty = true
            break
          }
        }
        if (hasNonEmpty) {
          break
        }
        endRow -= 1
      }
    }

    if (startRow > endRow) {
      const arrayIndex = allocateSpillArrayResult(1, 1)
      writeSpillArrayValue(arrayIndex, 0, <u8>ValueTag.Empty, 0)
      return writeArrayResult(base, arrayIndex, 1, 1, rangeIndexStack, valueStack, tagStack, kindStack)
    }

    if (trimLeadingCols) {
      while (startCol <= endCol) {
        let hasNonEmpty = false
        for (let row = startRow; row <= endRow; row++) {
          if (
            inputCellTag(
              base,
              row,
              startCol,
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
            ) != ValueTag.Empty
          ) {
            hasNonEmpty = true
            break
          }
        }
        if (hasNonEmpty) {
          break
        }
        startCol += 1
      }
    }

    if (trimTrailingCols) {
      while (endCol >= startCol) {
        let hasNonEmpty = false
        for (let row = startRow; row <= endRow; row++) {
          if (
            inputCellTag(
              base,
              row,
              endCol,
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
            ) != ValueTag.Empty
          ) {
            hasNonEmpty = true
            break
          }
        }
        if (hasNonEmpty) {
          break
        }
        endCol -= 1
      }
    }

    if (startCol > endCol) {
      const arrayIndex = allocateSpillArrayResult(1, 1)
      writeSpillArrayValue(arrayIndex, 0, <u8>ValueTag.Empty, 0)
      return writeArrayResult(base, arrayIndex, 1, 1, rangeIndexStack, valueStack, tagStack, kindStack)
    }

    const outputRows = endRow - startRow + 1
    const outputCols = endCol - startCol + 1
    const arrayIndex = allocateSpillArrayResult(outputRows, outputCols)
    let outputOffset = 0
    for (let row = startRow; row <= endRow; row++) {
      for (let col = startCol; col <= endCol; col++) {
        const copyError = copyInputCellToSpill(
          arrayIndex,
          outputOffset,
          base,
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
        if (copyError != ErrorCode.None) {
          return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, copyError, rangeIndexStack, valueStack, tagStack, kindStack)
        }
        outputOffset += 1
      }
    }
    return writeArrayResult(base, arrayIndex, outputRows, outputCols, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  return -1
}
