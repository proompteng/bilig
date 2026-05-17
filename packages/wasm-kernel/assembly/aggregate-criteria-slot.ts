import { matchesCriteriaValue } from './criteria'
import { inputCellScalarValue, inputCellTag, inputColsFromSlot, inputRowsFromSlot } from './operands'
import { toNumberOrNaN, toNumberOrZero } from './operands'
import { ValueTag } from './protocol'
import { STACK_KIND_ARRAY, STACK_KIND_RANGE } from './result-io'

export function isCriteriaRangeSlot(kind: u8): bool {
  return kind == STACK_KIND_RANGE || kind == STACK_KIND_ARRAY
}

export function criteriaSlotLength(
  slot: i32,
  kindStack: Uint8Array,
  rangeIndexStack: Uint32Array,
  rangeRowCounts: Uint32Array,
  rangeColCounts: Uint32Array,
): i32 {
  if (!isCriteriaRangeSlot(kindStack[slot])) {
    return i32.MIN_VALUE
  }
  const rows = inputRowsFromSlot(slot, kindStack, rangeIndexStack, rangeRowCounts)
  const cols = inputColsFromSlot(slot, kindStack, rangeIndexStack, rangeColCounts)
  return rows > 0 && cols > 0 ? rows * cols : i32.MIN_VALUE
}

function criteriaSlotRow(slot: i32, offset: i32, kindStack: Uint8Array, rangeIndexStack: Uint32Array, rangeColCounts: Uint32Array): i32 {
  const cols = inputColsFromSlot(slot, kindStack, rangeIndexStack, rangeColCounts)
  return cols > 0 ? offset / cols : i32.MIN_VALUE
}

function criteriaSlotCol(slot: i32, offset: i32, kindStack: Uint8Array, rangeIndexStack: Uint32Array, rangeColCounts: Uint32Array): i32 {
  const cols = inputColsFromSlot(slot, kindStack, rangeIndexStack, rangeColCounts)
  return cols > 0 ? offset % cols : i32.MIN_VALUE
}

export function criteriaSlotTagAt(
  slot: i32,
  offset: i32,
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
  return inputCellTag(
    slot,
    criteriaSlotRow(slot, offset, kindStack, rangeIndexStack, rangeColCounts),
    criteriaSlotCol(slot, offset, kindStack, rangeIndexStack, rangeColCounts),
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
}

export function criteriaSlotValueAt(
  slot: i32,
  offset: i32,
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
  return inputCellScalarValue(
    slot,
    criteriaSlotRow(slot, offset, kindStack, rangeIndexStack, rangeColCounts),
    criteriaSlotCol(slot, offset, kindStack, rangeIndexStack, rangeColCounts),
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

export function criteriaSlotMatches(
  slot: i32,
  offset: i32,
  criteriaTag: u8,
  criteriaValue: f64,
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
): bool {
  return matchesCriteriaValue(
    criteriaSlotTagAt(
      slot,
      offset,
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
    ),
    criteriaSlotValueAt(
      slot,
      offset,
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
    ),
    criteriaTag,
    criteriaValue,
    stringOffsets,
    stringLengths,
    stringData,
    outputStringOffsets,
    outputStringLengths,
    outputStringData,
  )
}

export function criteriaSlotNumberOrZero(
  slot: i32,
  offset: i32,
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
  return toNumberOrZero(
    criteriaSlotTagAt(
      slot,
      offset,
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
    ),
    criteriaSlotValueAt(
      slot,
      offset,
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
    ),
  )
}

export function criteriaSlotNumberOrNaN(
  slot: i32,
  offset: i32,
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
  return toNumberOrNaN(
    criteriaSlotTagAt(
      slot,
      offset,
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
    ),
    criteriaSlotValueAt(
      slot,
      offset,
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
    ),
  )
}

export function criteriaSlotNumberOnlyOrNaN(
  slot: i32,
  offset: i32,
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
  const tag = criteriaSlotTagAt(
    slot,
    offset,
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
  if (tag != ValueTag.Number) {
    return NaN
  }
  return criteriaSlotValueAt(
    slot,
    offset,
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
