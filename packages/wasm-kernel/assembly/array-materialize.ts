import { ErrorCode, ValueTag } from './protocol'
import { scalarText } from './text-codec'
import { inputCellScalarValue, inputCellTag, inputColsFromSlot, inputRowsFromSlot } from './operands'
import { STACK_KIND_SCALAR, copySlotResult, writeArrayResult, writeResult } from './result-io'
import { allocateSpillArrayResult, writeSpillArrayValue } from './vm'

export function copyInputCellToSpill(
  arrayIndex: u32,
  outputOffset: i32,
  slot: i32,
  row: i32,
  col: i32,
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
): i32 {
  const tag = inputCellTag(
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
  const value = inputCellScalarValue(
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
  if (tag == ValueTag.Error && isNaN(value)) {
    return ErrorCode.Value
  }
  writeSpillArrayValue(arrayIndex, outputOffset, tag, value)
  return ErrorCode.None
}

export function materializeSlotResult(
  base: i32,
  sourceSlot: i32,
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
): i32 {
  if (kindStack[sourceSlot] == STACK_KIND_SCALAR) {
    return copySlotResult(base, sourceSlot, rangeIndexStack, valueStack, tagStack, kindStack)
  }
  const rows = inputRowsFromSlot(sourceSlot, kindStack, rangeIndexStack, rangeRowCounts)
  const cols = inputColsFromSlot(sourceSlot, kindStack, rangeIndexStack, rangeColCounts)
  if (rows == i32.MIN_VALUE || cols == i32.MIN_VALUE || rows <= 0 || cols <= 0) {
    return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  const arrayIndex = allocateSpillArrayResult(rows, cols)
  let outputOffset = 0
  for (let row = 0; row < rows; row++) {
    for (let col = 0; col < cols; col++) {
      const copyError = copyInputCellToSpill(
        arrayIndex,
        outputOffset,
        sourceSlot,
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

  return writeArrayResult(base, arrayIndex, rows, cols, rangeIndexStack, valueStack, tagStack, kindStack)
}

export function uniqueScalarKey(
  tag: u8,
  value: f64,
  stringOffsets: Uint32Array,
  stringLengths: Uint32Array,
  stringData: Uint16Array,
  outputStringOffsets: Uint32Array,
  outputStringLengths: Uint32Array,
  outputStringData: Uint16Array,
): string | null {
  if (tag == ValueTag.Empty) {
    return 'E:'
  }
  if (tag == ValueTag.Number) {
    return 'N:' + value.toString()
  }
  if (tag == ValueTag.Boolean) {
    return value != 0 ? 'B:1' : 'B:0'
  }
  if (tag == ValueTag.String) {
    const text = scalarText(
      tag,
      value,
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    )
    return text == null ? null : 'S:' + text.toUpperCase()
  }
  return null
}

export function uniqueRowKey(
  slot: i32,
  row: i32,
  cols: i32,
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
  stringOffsets: Uint32Array,
  stringLengths: Uint32Array,
  stringData: Uint16Array,
  outputStringOffsets: Uint32Array,
  outputStringLengths: Uint32Array,
  outputStringData: Uint16Array,
): string | null {
  let key = ''
  for (let col = 0; col < cols; col++) {
    const tag = inputCellTag(
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
    const value = inputCellScalarValue(
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
    if (tag == ValueTag.Error) {
      return null
    }
    const cellKey = uniqueScalarKey(
      tag,
      value,
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    )
    if (cellKey == null) {
      return null
    }
    if (col > 0) {
      key += '\u0001'
    }
    key += cellKey
  }
  return key
}

export function uniqueColKey(
  slot: i32,
  col: i32,
  rows: i32,
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
  stringOffsets: Uint32Array,
  stringLengths: Uint32Array,
  stringData: Uint16Array,
  outputStringOffsets: Uint32Array,
  outputStringLengths: Uint32Array,
  outputStringData: Uint16Array,
): string | null {
  let key = ''
  for (let row = 0; row < rows; row++) {
    const tag = inputCellTag(
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
    const value = inputCellScalarValue(
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
    if (tag == ValueTag.Error) {
      return null
    }
    const cellKey = uniqueScalarKey(
      tag,
      value,
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    )
    if (cellKey == null) {
      return null
    }
    if (row > 0) {
      key += '\u0001'
    }
    key += cellKey
  }
  return key
}
