import { ValueTag } from './protocol'
import { registerTrackedArrayShape as registerTrackedArrayShapeImpl } from './dynamic-arrays'
import { memberScalarValue } from './operands'
import { allocateOutputString, encodeOutputStringId, writeOutputStringData } from './vm'

export const STACK_KIND_SCALAR: u8 = 0
export const STACK_KIND_RANGE: u8 = 1
export const STACK_KIND_ARRAY: u8 = 2
export const UNRESOLVED_WASM_OPERAND: u32 = 0x00ffffff

export function writeStringResult(
  base: i32,
  text: string,
  rangeIndexStack: Uint32Array,
  valueStack: Float64Array,
  tagStack: Uint8Array,
  kindStack: Uint8Array,
): i32 {
  const outputStringId = allocateOutputString(text.length)
  for (let index = 0; index < text.length; index++) {
    writeOutputStringData(outputStringId, index, <u16>text.charCodeAt(index))
  }
  return writeResult(
    base,
    STACK_KIND_SCALAR,
    <u8>ValueTag.String,
    encodeOutputStringId(outputStringId),
    rangeIndexStack,
    valueStack,
    tagStack,
    kindStack,
  )
}

export function writeMemberResult(
  base: i32,
  memberIndex: u32,
  rangeIndexStack: Uint32Array,
  valueStack: Float64Array,
  tagStack: Uint8Array,
  kindStack: Uint8Array,
  cellTags: Uint8Array,
  cellNumbers: Float64Array,
  cellStringIds: Uint32Array,
  cellErrors: Uint16Array,
): i32 {
  return writeResult(
    base,
    STACK_KIND_SCALAR,
    cellTags[memberIndex],
    memberScalarValue(memberIndex, cellTags, cellNumbers, cellStringIds, cellErrors),
    rangeIndexStack,
    valueStack,
    tagStack,
    kindStack,
  )
}

export function writeResult(
  base: i32,
  kind: u8,
  tag: u8,
  value: f64,
  rangeIndexStack: Uint32Array,
  valueStack: Float64Array,
  tagStack: Uint8Array,
  kindStack: Uint8Array,
): i32 {
  rangeIndexStack[base] = 0
  valueStack[base] = value
  tagStack[base] = tag
  kindStack[base] = kind
  return base + 1
}

export function writeArrayResult(
  base: i32,
  arrayIndex: u32,
  rows: i32,
  cols: i32,
  rangeIndexStack: Uint32Array,
  valueStack: Float64Array,
  tagStack: Uint8Array,
  kindStack: Uint8Array,
): i32 {
  registerTrackedArrayShapeImpl(arrayIndex, rows, cols)
  rangeIndexStack[base] = arrayIndex
  valueStack[base] = 0
  tagStack[base] = ValueTag.Empty
  kindStack[base] = STACK_KIND_ARRAY
  return base + 1
}

export function copySlotResult(
  base: i32,
  sourceSlot: i32,
  rangeIndexStack: Uint32Array,
  valueStack: Float64Array,
  tagStack: Uint8Array,
  kindStack: Uint8Array,
): i32 {
  kindStack[base] = kindStack[sourceSlot]
  tagStack[base] = tagStack[sourceSlot]
  valueStack[base] = valueStack[sourceSlot]
  rangeIndexStack[base] = rangeIndexStack[sourceSlot]
  return base + 1
}

export function vectorSlotLength(
  slot: i32,
  kindStack: Uint8Array,
  rangeIndexStack: Uint32Array,
  rangeLengths: Uint32Array,
  rangeRowCounts: Uint32Array,
  rangeColCounts: Uint32Array,
): i32 {
  if (kindStack[slot] == STACK_KIND_SCALAR) {
    return 1
  }
  if (kindStack[slot] != STACK_KIND_RANGE) {
    return i32.MIN_VALUE
  }
  const rangeIndex = rangeIndexStack[slot]
  const rowCount = <i32>rangeRowCounts[rangeIndex]
  const colCount = <i32>rangeColCounts[rangeIndex]
  if (rowCount <= 0 || colCount <= 0 || (rowCount != 1 && colCount != 1)) {
    return i32.MIN_VALUE
  }
  return <i32>rangeLengths[rangeIndex]
}
