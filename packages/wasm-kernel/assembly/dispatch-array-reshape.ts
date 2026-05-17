import { copyInputCellToSpill } from './array-materialize'
import { BuiltinId, ErrorCode, ValueTag } from './protocol'
import { inputCellScalarValue, inputCellTag, inputColsFromSlot, inputRowsFromSlot } from './operands'
import { coerceInteger } from './numeric-core'
import { scalarErrorAt } from './builtin-args'
import { STACK_KIND_ARRAY, STACK_KIND_RANGE, STACK_KIND_SCALAR, writeArrayResult, writeResult } from './result-io'
import { allocateSpillArrayResult, writeSpillArrayValue } from './vm'

function coerceBoolean(tag: u8, value: f64): i32 {
  if (tag == ValueTag.Boolean || tag == ValueTag.Number) {
    return value != 0 ? 1 : 0
  }
  if (tag == ValueTag.Empty) {
    return 0
  }
  return -1
}

function writeValueError(
  base: i32,
  code: f64,
  rangeIndexStack: Uint32Array,
  valueStack: Float64Array,
  tagStack: Uint8Array,
  kindStack: Uint8Array,
): i32 {
  return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, code, rangeIndexStack, valueStack, tagStack, kindStack)
}

function isArrayLike(kind: u8): bool {
  return kind == STACK_KIND_SCALAR || kind == STACK_KIND_RANGE || kind == STACK_KIND_ARRAY
}

function appendFlattenedValues(
  tags: Array<u8>,
  values: Array<f64>,
  base: i32,
  sourceRows: i32,
  sourceCols: i32,
  ignoreMode: i32,
  scanByCol: bool,
  kindStack: Uint8Array,
  valueStack: Float64Array,
  tagStack: Uint8Array,
  rangeIndexStack: Uint32Array,
  rangeOffsets: Uint32Array,
  rangeLengths: Uint32Array,
  rangeRowCounts: Uint32Array,
  rangeColCounts: Uint32Array,
  rangeMembers: Uint32Array,
  cellTags: Uint8Array,
  cellNumbers: Float64Array,
  cellStringIds: Uint32Array,
  cellErrors: Uint16Array,
): f64 {
  if (!scanByCol) {
    for (let row = 0; row < sourceRows; row++) {
      for (let col = 0; col < sourceCols; col++) {
        const sourceTag = inputCellTag(
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
        )
        if (
          (sourceTag == ValueTag.Empty && (ignoreMode == 1 || ignoreMode == 3)) ||
          (sourceTag == ValueTag.Error && (ignoreMode == 2 || ignoreMode == 3))
        ) {
          continue
        }
        const sourceValue = inputCellScalarValue(
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
        if (sourceTag == ValueTag.Error && isNaN(sourceValue)) {
          return ErrorCode.Value
        }
        tags.push(sourceTag)
        values.push(sourceValue)
      }
    }
    return ErrorCode.None
  }

  for (let col = 0; col < sourceCols; col++) {
    for (let row = 0; row < sourceRows; row++) {
      const sourceTag = inputCellTag(
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
      )
      if (
        (sourceTag == ValueTag.Empty && (ignoreMode == 1 || ignoreMode == 3)) ||
        (sourceTag == ValueTag.Error && (ignoreMode == 2 || ignoreMode == 3))
      ) {
        continue
      }
      const sourceValue = inputCellScalarValue(
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
      if (sourceTag == ValueTag.Error && isNaN(sourceValue)) {
        return ErrorCode.Value
      }
      tags.push(sourceTag)
      values.push(sourceValue)
    }
  }

  return ErrorCode.None
}

export function tryApplyArrayReshapeBuiltin(
  builtinId: i32,
  argc: i32,
  base: i32,
  rangeIndexStack: Uint32Array,
  valueStack: Float64Array,
  tagStack: Uint8Array,
  kindStack: Uint8Array,
  rangeOffsets: Uint32Array,
  rangeLengths: Uint32Array,
  rangeRowCounts: Uint32Array,
  rangeColCounts: Uint32Array,
  rangeMembers: Uint32Array,
  cellTags: Uint8Array,
  cellNumbers: Float64Array,
  cellStringIds: Uint32Array,
  cellErrors: Uint16Array,
): i32 {
  if ((builtinId == BuiltinId.Tocol && argc >= 1 && argc <= 3) || (builtinId == BuiltinId.Torow && argc >= 1 && argc <= 3)) {
    const sourceKind = kindStack[base]
    if (!isArrayLike(sourceKind)) {
      return writeValueError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const sourceRows = inputRowsFromSlot(base, kindStack, rangeIndexStack, rangeRowCounts)
    const sourceCols = inputColsFromSlot(base, kindStack, rangeIndexStack, rangeColCounts)
    if (sourceRows <= 0 || sourceCols <= 0 || sourceRows == i32.MIN_VALUE || sourceCols == i32.MIN_VALUE) {
      return writeValueError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack)
    if (scalarError >= 0) {
      return writeValueError(base, scalarError, rangeIndexStack, valueStack, tagStack, kindStack)
    }

    const ignoreValue = argc >= 2 ? coerceInteger(tagStack[base + 1], valueStack[base + 1]) : 0
    const scanByCol = argc >= 3 ? coerceBoolean(tagStack[base + 2], valueStack[base + 2]) : 0
    if (ignoreValue < 0 || ignoreValue > 3 || scanByCol < 0) {
      return writeValueError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }

    const tags = new Array<u8>()
    const values = new Array<f64>()
    const appendError = appendFlattenedValues(
      tags,
      values,
      base,
      sourceRows,
      sourceCols,
      ignoreValue,
      scanByCol != 0,
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
    if (appendError != ErrorCode.None) {
      return writeValueError(base, appendError, rangeIndexStack, valueStack, tagStack, kindStack)
    }

    const outputRows = builtinId == BuiltinId.Tocol ? values.length : 1
    const outputCols = builtinId == BuiltinId.Tocol ? 1 : values.length
    const arrayIndex = allocateSpillArrayResult(outputRows, outputCols)
    for (let offset = 0; offset < values.length; offset++) {
      writeSpillArrayValue(arrayIndex, offset, tags[offset], values[offset])
    }
    return writeArrayResult(base, arrayIndex, outputRows, outputCols, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if ((builtinId == BuiltinId.Wraprows && argc >= 2 && argc <= 4) || (builtinId == BuiltinId.Wrapcols && argc >= 2 && argc <= 4)) {
    const sourceKind = kindStack[base]
    if (!isArrayLike(sourceKind)) {
      return writeValueError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const sourceRows = inputRowsFromSlot(base, kindStack, rangeIndexStack, rangeRowCounts)
    const sourceCols = inputColsFromSlot(base, kindStack, rangeIndexStack, rangeColCounts)
    if (sourceRows <= 0 || sourceCols <= 0 || sourceRows == i32.MIN_VALUE || sourceCols == i32.MIN_VALUE) {
      return writeValueError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack)
    if (scalarError >= 0) {
      return writeValueError(base, scalarError, rangeIndexStack, valueStack, tagStack, kindStack)
    }

    const wrapCount = coerceInteger(tagStack[base + 1], valueStack[base + 1])
    if (wrapCount == i32.MIN_VALUE || wrapCount < 1) {
      return writeValueError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    if (argc >= 4) {
      const padByColBoolean = coerceBoolean(tagStack[base + 3], valueStack[base + 3])
      if (padByColBoolean < 0) {
        return writeValueError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      }
    }

    const sourceLength = sourceRows * sourceCols
    const outputRows = builtinId == BuiltinId.Wraprows ? (sourceLength + wrapCount - 1) / wrapCount : wrapCount
    const outputCols = builtinId == BuiltinId.Wraprows ? wrapCount : (sourceLength + wrapCount - 1) / wrapCount
    const outputLength = outputRows * outputCols
    const arrayIndex = allocateSpillArrayResult(outputRows, outputCols)
    const padTag = argc >= 3 ? tagStack[base + 2] : <u8>ValueTag.Error
    const padValue = argc >= 3 ? valueStack[base + 2] : ErrorCode.NA

    if (builtinId == BuiltinId.Wraprows) {
      for (let outputOffset = 0; outputOffset < sourceLength; outputOffset++) {
        const sourceRow = outputOffset / sourceCols
        const sourceCol = outputOffset - sourceRow * sourceCols
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
          return writeValueError(base, copyError, rangeIndexStack, valueStack, tagStack, kindStack)
        }
      }
      for (let outputOffset = sourceLength; outputOffset < outputLength; outputOffset++) {
        writeSpillArrayValue(arrayIndex, outputOffset, padTag, padValue)
      }
    } else {
      for (let outputOffset = 0; outputOffset < outputLength; outputOffset++) {
        const row = outputOffset / outputCols
        const col = outputOffset - row * outputCols
        const sourceOffset = col * outputRows + row
        if (sourceOffset < sourceLength) {
          const sourceRow = sourceOffset / sourceCols
          const sourceCol = sourceOffset - sourceRow * sourceCols
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
            return writeValueError(base, copyError, rangeIndexStack, valueStack, tagStack, kindStack)
          }
        } else {
          writeSpillArrayValue(arrayIndex, outputOffset, padTag, padValue)
        }
      }
    }

    return writeArrayResult(base, arrayIndex, outputRows, outputCols, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  return -1
}
