import { BuiltinId, ErrorCode, ValueTag } from './protocol'
import { rangeSupportedScalarOnly, scalarArgsOnly, scalarErrorAt } from './builtin-args'
import { inputCellScalarValue, inputCellTag, inputColsFromSlot, inputRowsFromSlot, memberScalarValue, rangeMemberAt } from './operands'
import { copyInputCellToSpill, materializeSlotResult } from './array-materialize'
import { truncToInt } from './numeric-core'
import { arrayToTextCell } from './text-special'
import { valueNumber } from './comparison'
import { scalarText } from './text-codec'
import {
  copySlotResult,
  STACK_KIND_ARRAY,
  STACK_KIND_RANGE,
  STACK_KIND_SCALAR,
  writeArrayResult,
  writeResult,
  writeStringResult,
} from './result-io'
import { allocateSpillArrayResult } from './vm'

export function tryApplyArrayInfoBuiltin(
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
  if (builtinId == BuiltinId.Areas && argc == 1) {
    if (kindStack[base] != STACK_KIND_RANGE) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, 1, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.Arraytotext && (argc == 1 || argc == 2)) {
    if (kindStack[base] != STACK_KIND_SCALAR && kindStack[base] != STACK_KIND_RANGE) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    let format = 0
    if (argc == 2) {
      if (kindStack[base + 1] != STACK_KIND_SCALAR) {
        return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      format = truncToInt(tagStack[base + 1], valueStack[base + 1])
      if (!(format == 0 || format == 1)) {
        return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      }
    }

    const strict = format == 1
    let text = strict ? '{' : ''
    if (kindStack[base] == STACK_KIND_SCALAR) {
      const cellText = arrayToTextCell(
        tagStack[base],
        valueStack[base],
        strict,
        stringOffsets,
        stringLengths,
        stringData,
        outputStringOffsets,
        outputStringLengths,
        outputStringData,
      )
      if (cellText == null) {
        return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      text += cellText
    } else {
      const rangeIndex = rangeIndexStack[base]
      const rowCount = <i32>rangeRowCounts[rangeIndex]
      const colCount = <i32>rangeColCounts[rangeIndex]
      for (let row = 0; row < rowCount; row += 1) {
        if (row > 0) {
          text += ';'
        }
        for (let col = 0; col < colCount; col += 1) {
          if (col > 0) {
            text += strict ? ', ' : '\t'
          }
          const memberIndex = rangeMemberAt(rangeIndex, row, col, rangeOffsets, rangeLengths, rangeRowCounts, rangeColCounts, rangeMembers)
          if (memberIndex == 0xffffffff) {
            return writeResult(
              base,
              STACK_KIND_SCALAR,
              <u8>ValueTag.Error,
              ErrorCode.Value,
              rangeIndexStack,
              valueStack,
              tagStack,
              kindStack,
            )
          }
          const cellText = arrayToTextCell(
            cellTags[memberIndex],
            memberScalarValue(memberIndex, cellTags, cellNumbers, cellStringIds, cellErrors),
            strict,
            stringOffsets,
            stringLengths,
            stringData,
            outputStringOffsets,
            outputStringLengths,
            outputStringData,
          )
          if (cellText == null) {
            return writeResult(
              base,
              STACK_KIND_SCALAR,
              <u8>ValueTag.Error,
              ErrorCode.Value,
              rangeIndexStack,
              valueStack,
              tagStack,
              kindStack,
            )
          }
          text += cellText
        }
      }
    }
    if (strict) {
      text += '}'
    }
    return writeStringResult(base, text, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.Columns && argc == 1) {
    if (kindStack[base] != STACK_KIND_RANGE) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      rangeColCounts[rangeIndexStack[base]],
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    )
  }

  if (builtinId == BuiltinId.Rows && argc == 1) {
    if (kindStack[base] != STACK_KIND_RANGE) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      rangeRowCounts[rangeIndexStack[base]],
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    )
  }

  if (builtinId == BuiltinId.Transpose && argc == 1) {
    if (kindStack[base] != STACK_KIND_SCALAR && kindStack[base] != STACK_KIND_RANGE && kindStack[base] != STACK_KIND_ARRAY) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const sourceRows = inputRowsFromSlot(base, kindStack, rangeIndexStack, rangeRowCounts)
    const sourceCols = inputColsFromSlot(base, kindStack, rangeIndexStack, rangeColCounts)
    if (sourceRows <= 0 || sourceCols <= 0 || sourceRows == i32.MIN_VALUE || sourceCols == i32.MIN_VALUE) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    if (sourceRows == 1 && sourceCols == 1) {
      const tag = inputCellTag(
        base,
        0,
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
      const value = inputCellScalarValue(
        base,
        0,
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
      if (tag == ValueTag.Error && isNaN(value)) {
        return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      return writeResult(base, STACK_KIND_SCALAR, tag, value, rangeIndexStack, valueStack, tagStack, kindStack)
    }

    const arrayIndex = allocateSpillArrayResult(sourceCols, sourceRows)
    let outputOffset = 0
    for (let col = 0; col < sourceCols; col += 1) {
      for (let row = 0; row < sourceRows; row += 1) {
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
    return writeArrayResult(base, arrayIndex, sourceCols, sourceRows, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.Hstack && argc >= 1) {
    let rowCount = 0
    let totalCols = 0
    for (let index = 0; index < argc; index += 1) {
      const slot = base + index
      const kind = kindStack[slot]
      if (kind != STACK_KIND_SCALAR && kind != STACK_KIND_RANGE && kind != STACK_KIND_ARRAY) {
        return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      const rows = inputRowsFromSlot(slot, kindStack, rangeIndexStack, rangeRowCounts)
      const cols = inputColsFromSlot(slot, kindStack, rangeIndexStack, rangeColCounts)
      if (rows <= 0 || cols <= 0 || rows == i32.MIN_VALUE || cols == i32.MIN_VALUE) {
        return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      rowCount = max(rowCount, rows)
      totalCols += cols
    }
    for (let index = 0; index < argc; index += 1) {
      const rows = inputRowsFromSlot(base + index, kindStack, rangeIndexStack, rangeRowCounts)
      if (rows != 1 && rows != rowCount) {
        return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      }
    }

    const arrayIndex = allocateSpillArrayResult(rowCount, totalCols)
    let outputOffset = 0
    for (let row = 0; row < rowCount; row += 1) {
      for (let index = 0; index < argc; index += 1) {
        const slot = base + index
        const rows = inputRowsFromSlot(slot, kindStack, rangeIndexStack, rangeRowCounts)
        const cols = inputColsFromSlot(slot, kindStack, rangeIndexStack, rangeColCounts)
        const sourceRow = rows == 1 ? 0 : row
        for (let col = 0; col < cols; col += 1) {
          const copyError = copyInputCellToSpill(
            arrayIndex,
            outputOffset,
            slot,
            sourceRow,
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
    }
    return writeArrayResult(base, arrayIndex, rowCount, totalCols, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.Vstack && argc >= 1) {
    let totalRows = 0
    let colCount = 0
    for (let index = 0; index < argc; index += 1) {
      const slot = base + index
      const kind = kindStack[slot]
      if (kind != STACK_KIND_SCALAR && kind != STACK_KIND_RANGE && kind != STACK_KIND_ARRAY) {
        return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      const rows = inputRowsFromSlot(slot, kindStack, rangeIndexStack, rangeRowCounts)
      const cols = inputColsFromSlot(slot, kindStack, rangeIndexStack, rangeColCounts)
      if (rows <= 0 || cols <= 0 || rows == i32.MIN_VALUE || cols == i32.MIN_VALUE) {
        return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      totalRows += rows
      colCount = max(colCount, cols)
    }
    for (let index = 0; index < argc; index += 1) {
      const cols = inputColsFromSlot(base + index, kindStack, rangeIndexStack, rangeColCounts)
      if (cols != 1 && cols != colCount) {
        return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      }
    }

    const arrayIndex = allocateSpillArrayResult(totalRows, colCount)
    let outputOffset = 0
    for (let index = 0; index < argc; index += 1) {
      const slot = base + index
      const rows = inputRowsFromSlot(slot, kindStack, rangeIndexStack, rangeRowCounts)
      const cols = inputColsFromSlot(slot, kindStack, rangeIndexStack, rangeColCounts)
      for (let row = 0; row < rows; row += 1) {
        for (let col = 0; col < colCount; col += 1) {
          const sourceCol = cols == 1 ? 0 : col
          const copyError = copyInputCellToSpill(
            arrayIndex,
            outputOffset,
            slot,
            row,
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
    }
    return writeArrayResult(base, arrayIndex, totalRows, colCount, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.Choose && argc >= 2) {
    const choice = truncToInt(tagStack[base], valueStack[base])
    if (choice == i32.MIN_VALUE || choice < 1 || choice >= argc) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    return materializeSlotResult(
      base,
      base + choice,
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

  if (builtinId == BuiltinId.Textjoin && argc >= 3) {
    if (kindStack[base] != STACK_KIND_SCALAR || kindStack[base + 1] != STACK_KIND_SCALAR) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    if (tagStack[base] == ValueTag.Error) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, valueStack[base], rangeIndexStack, valueStack, tagStack, kindStack)
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

    const delimiter = scalarText(
      tagStack[base],
      valueStack[base],
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    )
    const ignoreEmptyNumeric = valueNumber(
      tagStack[base + 1],
      valueStack[base + 1],
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    )
    if (delimiter == null || !isFinite(ignoreEmptyNumeric)) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }

    const ignoreEmpty = ignoreEmptyNumeric != 0
    let joined = ''
    let hasJoinedValue = false
    for (let slot = base + 2; slot < base + argc; slot += 1) {
      const kind = kindStack[slot]
      if (kind != STACK_KIND_SCALAR && kind != STACK_KIND_RANGE && kind != STACK_KIND_ARRAY) {
        return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      const rows = inputRowsFromSlot(slot, kindStack, rangeIndexStack, rangeRowCounts)
      const cols = inputColsFromSlot(slot, kindStack, rangeIndexStack, rangeColCounts)
      if (rows == i32.MIN_VALUE || cols == i32.MIN_VALUE || rows <= 0 || cols <= 0) {
        return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      for (let row = 0; row < rows; row += 1) {
        for (let col = 0; col < cols; col += 1) {
          const memberTag = inputCellTag(
            slot,
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
          const memberValue = inputCellScalarValue(
            slot,
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
          if (memberTag == ValueTag.Error) {
            return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, memberValue, rangeIndexStack, valueStack, tagStack, kindStack)
          }
          const part = scalarText(
            memberTag,
            memberValue,
            stringOffsets,
            stringLengths,
            stringData,
            outputStringOffsets,
            outputStringLengths,
            outputStringData,
          )
          if (part == null) {
            return writeResult(
              base,
              STACK_KIND_SCALAR,
              <u8>ValueTag.Error,
              ErrorCode.Value,
              rangeIndexStack,
              valueStack,
              tagStack,
              kindStack,
            )
          }
          if (ignoreEmpty && part.length == 0) {
            continue
          }
          if (hasJoinedValue) {
            joined += delimiter
          }
          joined += part
          hasJoinedValue = true
        }
      }
    }
    return writeStringResult(base, joined, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (
    (builtinId == BuiltinId.Na ||
      builtinId == BuiltinId.Iferror ||
      builtinId == BuiltinId.Ifna ||
      builtinId == BuiltinId.T ||
      builtinId == BuiltinId.N ||
      builtinId == BuiltinId.Type ||
      builtinId == BuiltinId.Delta ||
      builtinId == BuiltinId.Gestep) &&
    !rangeSupportedScalarOnly(base, argc, kindStack)
  ) {
    return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.Na && argc == 0) {
    return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.NA, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.Iferror && argc == 2) {
    if (tagStack[base] == ValueTag.Error) {
      return copySlotResult(base, base + 1, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    return base + 1
  }

  if (builtinId == BuiltinId.Ifna && argc == 2) {
    if (tagStack[base] == ValueTag.Error && <i32>valueStack[base] == ErrorCode.NA) {
      return copySlotResult(base, base + 1, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    return base + 1
  }

  if (builtinId == BuiltinId.T && (argc == 0 || argc == 1)) {
    if (argc == 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Empty, 0, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    if (!scalarArgsOnly(base, argc, kindStack)) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    if (tagStack[base] == ValueTag.Error || tagStack[base] == ValueTag.String) {
      return base + 1
    }
    return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Empty, 0, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.N && (argc == 0 || argc == 1)) {
    if (argc == 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, 0, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    if (!scalarArgsOnly(base, argc, kindStack)) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    if (tagStack[base] == ValueTag.Error) {
      return base + 1
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      tagStack[base] == ValueTag.Boolean || tagStack[base] == ValueTag.Number ? valueStack[base] : 0,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    )
  }

  if (builtinId == BuiltinId.Type && (argc == 0 || argc == 1)) {
    if (argc == 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, 1, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const typeCode =
      kindStack[base] == STACK_KIND_ARRAY || kindStack[base] == STACK_KIND_RANGE
        ? 64
        : tagStack[base] == ValueTag.Number || tagStack[base] == ValueTag.Empty
          ? 1
          : tagStack[base] == ValueTag.String
            ? 2
            : tagStack[base] == ValueTag.Boolean
              ? 4
              : 16
    return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, <f64>typeCode, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.Delta && (argc == 1 || argc == 2)) {
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack)
    if (scalarError >= 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, scalarError, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const left = valueNumber(
      tagStack[base],
      valueStack[base],
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    )
    const right =
      argc == 2
        ? valueNumber(
            tagStack[base + 1],
            valueStack[base + 1],
            stringOffsets,
            stringLengths,
            stringData,
            outputStringOffsets,
            outputStringLengths,
            outputStringData,
          )
        : 0.0
    if (isNaN(left) || isNaN(right)) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      left == right ? 1 : 0,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    )
  }

  if (builtinId == BuiltinId.Gestep && (argc == 1 || argc == 2)) {
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack)
    if (scalarError >= 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, scalarError, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const numberValue = valueNumber(
      tagStack[base],
      valueStack[base],
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    )
    const stepValue =
      argc == 2
        ? valueNumber(
            tagStack[base + 1],
            valueStack[base + 1],
            stringOffsets,
            stringLengths,
            stringData,
            outputStringOffsets,
            outputStringLengths,
            outputStringData,
          )
        : 0.0
    if (isNaN(numberValue) || isNaN(stepValue)) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      numberValue >= stepValue ? 1 : 0,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    )
  }

  return -1
}
