import { ValueTag } from './protocol'
import { getTrackedArrayCols as getDynamicArrayCols, getTrackedArrayRows as getDynamicArrayRows } from './dynamic-arrays'
import { readSpillArrayTag, readSpillArrayLength, readSpillArrayNumber } from './vm'

const STACK_KIND_SCALAR: u8 = 0
const STACK_KIND_RANGE: u8 = 1
const STACK_KIND_ARRAY: u8 = 2
const UNRESOLVED_WASM_OPERAND: u32 = 0x00ffffff

export function toNumberOrNaN(tag: u8, value: f64): f64 {
  if (tag == ValueTag.Number || tag == ValueTag.Boolean) return value
  if (tag == ValueTag.Empty) return 0
  return NaN
}

export function toNumberOrZero(tag: u8, value: f64): f64 {
  const numeric = toNumberOrNaN(tag, value)
  return isNaN(numeric) ? 0 : numeric
}

export function toNumberExact(tag: u8, value: f64): f64 {
  if (tag == ValueTag.Number || tag == ValueTag.Boolean) return value
  if (tag == ValueTag.Empty) return 0
  return NaN
}

export function inputRowsFromSlot(slot: i32, kindStack: Uint8Array, rangeIndexStack: Uint32Array, rangeRowCounts: Uint32Array): i32 {
  const kind = kindStack[slot]
  if (kind == STACK_KIND_SCALAR) {
    return 1
  }
  if (kind == STACK_KIND_RANGE) {
    return <i32>rangeRowCounts[rangeIndexStack[slot]]
  }
  if (kind == STACK_KIND_ARRAY) {
    return getDynamicArrayRows(rangeIndexStack[slot])
  }
  return i32.MIN_VALUE
}

export function inputColsFromSlot(slot: i32, kindStack: Uint8Array, rangeIndexStack: Uint32Array, rangeColCounts: Uint32Array): i32 {
  const kind = kindStack[slot]
  if (kind == STACK_KIND_SCALAR) {
    return 1
  }
  if (kind == STACK_KIND_RANGE) {
    return <i32>rangeColCounts[rangeIndexStack[slot]]
  }
  if (kind == STACK_KIND_ARRAY) {
    return getDynamicArrayCols(rangeIndexStack[slot])
  }
  return i32.MIN_VALUE
}

export function memberScalarValue(
  memberIndex: u32,
  cellTags: Uint8Array,
  cellNumbers: Float64Array,
  cellStringIds: Uint32Array,
  cellErrors: Uint16Array,
): f64 {
  const tag = cellTags[memberIndex]
  if (tag == ValueTag.String) {
    return <f64>cellStringIds[memberIndex]
  }
  if (tag == ValueTag.Error) {
    return <f64>cellErrors[memberIndex]
  }
  return cellNumbers[memberIndex]
}

export function rangeMemberAt(
  rangeIndex: u32,
  row: i32,
  col: i32,
  rangeOffsets: Uint32Array,
  rangeLengths: Uint32Array,
  rangeRowCounts: Uint32Array,
  rangeColCounts: Uint32Array,
  rangeMembers: Uint32Array,
): u32 {
  if (rangeIndex == UNRESOLVED_WASM_OPERAND) {
    return 0xffffffff
  }
  const rowCount = <i32>rangeRowCounts[rangeIndex]
  const colCount = <i32>rangeColCounts[rangeIndex]
  const length = <i32>rangeLengths[rangeIndex]
  if (rowCount <= 0 || colCount <= 0 || row < 0 || col < 0 || row >= rowCount || col >= colCount || row * colCount + col >= length) {
    return 0xffffffff
  }
  return rangeMembers[rangeOffsets[rangeIndex] + row * colCount + col]
}

export function inputCellTag(
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
): u8 {
  const kind = kindStack[slot]
  if (row < 0 || col < 0) {
    return <u8>ValueTag.Error
  }
  if (kind == STACK_KIND_SCALAR) {
    if (row != 0 || col != 0) {
      return <u8>ValueTag.Error
    }
    return tagStack[slot]
  }
  if (kind == STACK_KIND_RANGE) {
    const rangeIndex = rangeIndexStack[slot]
    const memberIndex = rangeMemberAt(rangeIndex, row, col, rangeOffsets, rangeLengths, rangeRowCounts, rangeColCounts, rangeMembers)
    if (memberIndex == 0xffffffff) {
      return <u8>ValueTag.Error
    }
    return cellTags[memberIndex]
  }
  if (kind == STACK_KIND_ARRAY) {
    const arrayIndex = rangeIndexStack[slot]
    const arrayRows = getDynamicArrayRows(arrayIndex)
    const arrayCols = getDynamicArrayCols(arrayIndex)
    if (arrayRows < 1 || arrayCols < 1 || row >= arrayRows || col >= arrayCols) {
      return <u8>ValueTag.Error
    }
    const arrayOffset = row * arrayCols + col
    if (arrayOffset >= readSpillArrayLength(arrayIndex)) {
      return <u8>ValueTag.Error
    }
    return readSpillArrayTag(arrayIndex, arrayOffset)
  }
  return <u8>ValueTag.Error
}

export function inputCellScalarValue(
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
): f64 {
  const kind = kindStack[slot]
  if (row < 0 || col < 0) {
    return NaN
  }
  if (kind == STACK_KIND_SCALAR) {
    return row == 0 && col == 0 ? valueStack[slot] : NaN
  }
  if (kind == STACK_KIND_RANGE) {
    const rangeIndex = rangeIndexStack[slot]
    const memberIndex = rangeMemberAt(rangeIndex, row, col, rangeOffsets, rangeLengths, rangeRowCounts, rangeColCounts, rangeMembers)
    if (memberIndex == 0xffffffff) {
      return NaN
    }
    return memberScalarValue(memberIndex, cellTags, cellNumbers, cellStringIds, cellErrors)
  }
  if (kind == STACK_KIND_ARRAY) {
    const arrayIndex = rangeIndexStack[slot]
    const arrayRows = getDynamicArrayRows(arrayIndex)
    const arrayCols = getDynamicArrayCols(arrayIndex)
    if (arrayRows < 1 || arrayCols < 1 || row >= arrayRows || col >= arrayCols) {
      return NaN
    }
    const arrayOffset = row * arrayCols + col
    if (arrayOffset >= readSpillArrayLength(arrayIndex)) {
      return NaN
    }
    return readSpillArrayNumber(arrayIndex, arrayOffset)
  }
  return NaN
}

export function inputCellNumeric(
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
): f64 {
  const kind = kindStack[slot]
  if (row < 0 || col < 0) {
    return NaN
  }
  if (kind == STACK_KIND_SCALAR) {
    return row == 0 && col == 0 ? toNumberOrNaN(tagStack[slot], valueStack[slot]) : NaN
  }
  if (kind == STACK_KIND_RANGE) {
    const rangeIndex = rangeIndexStack[slot]
    const memberIndex = rangeMemberAt(rangeIndex, row, col, rangeOffsets, rangeLengths, rangeRowCounts, rangeColCounts, rangeMembers)
    if (memberIndex == 0xffffffff) {
      return NaN
    }
    return toNumberOrNaN(cellTags[memberIndex], cellNumbers[memberIndex])
  }
  if (kind == STACK_KIND_ARRAY) {
    const arrayIndex = rangeIndexStack[slot]
    const arrayRows = getDynamicArrayRows(arrayIndex)
    const arrayCols = getDynamicArrayCols(arrayIndex)
    if (arrayRows < 1 || arrayCols < 1 || row >= arrayRows || col >= arrayCols) {
      return NaN
    }
    const arrayOffset = row * arrayCols + col
    if (arrayOffset >= readSpillArrayLength(arrayIndex)) {
      return NaN
    }
    return toNumberOrNaN(readSpillArrayTag(arrayIndex, arrayOffset), readSpillArrayNumber(arrayIndex, arrayOffset))
  }
  return NaN
}
