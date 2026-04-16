import { ValueTag } from './protocol'
import { scalarText, trimAsciiWhitespace } from './text-codec'
import { memberScalarValue } from './operands'
import { STACK_KIND_RANGE, STACK_KIND_SCALAR } from './result-io'
import { coerceInteger, truncToInt } from './numeric-core'

export function isWeekendSerial(whole: i32): bool {
  const adjustedWhole = whole < 60 ? whole : whole - 1
  const dayOfWeek = ((adjustedWhole % 7) + 7) % 7
  return dayOfWeek == 0 || dayOfWeek == 6
}

export function weekendSerialDay(whole: i32): i32 {
  const adjustedWhole = whole < 60 ? whole : whole - 1
  return ((adjustedWhole % 7) + 7) % 7
}

export function weekendMaskFromCode(code: i32): i32 {
  if (code == 1) return (1 << 6) | (1 << 0)
  if (code == 2) return (1 << 0) | (1 << 1)
  if (code == 3) return (1 << 1) | (1 << 2)
  if (code == 4) return (1 << 2) | (1 << 3)
  if (code == 5) return (1 << 3) | (1 << 4)
  if (code == 6) return (1 << 4) | (1 << 5)
  if (code == 7) return (1 << 5) | (1 << 6)
  if (code >= 11 && code <= 17) {
    return 1 << (code == 17 ? 0 : code - 10)
  }
  return i32.MIN_VALUE
}

export function weekendMaskFromString(maskText: string): i32 {
  const trimmed = trimAsciiWhitespace(maskText)
  if (trimmed.length != 7) {
    return i32.MIN_VALUE
  }
  let mask = 0
  for (let index = 0; index < 7; index += 1) {
    const char = trimmed.charCodeAt(index)
    if (char != 48 && char != 49) {
      return i32.MIN_VALUE
    }
    if (char != 49) {
      continue
    }
    const day = index == 6 ? 0 : index + 1
    mask |= 1 << day
  }
  return mask != 0x7f ? mask : i32.MIN_VALUE
}

export function coerceWeekendMask(
  hasWeekendArg: bool,
  tag: u8,
  value: f64,
  stringOffsets: Uint32Array,
  stringLengths: Uint32Array,
  stringData: Uint16Array,
  outputStringOffsets: Uint32Array,
  outputStringLengths: Uint32Array,
  outputStringData: Uint16Array,
): i32 {
  if (!hasWeekendArg) {
    return (1 << 0) | (1 << 6)
  }
  if (tag == ValueTag.String) {
    const maskText = scalarText(
      tag,
      value,
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    )
    return maskText == null ? i32.MIN_VALUE : weekendMaskFromString(maskText)
  }
  const code = coerceInteger(tag, value)
  return code == i32.MIN_VALUE ? i32.MIN_VALUE : weekendMaskFromCode(code)
}

function isWeekendWithMask(serial: i32, weekendMask: i32): bool {
  return (weekendMask & (1 << weekendSerialDay(serial))) != 0
}

export function isHolidaySerial(
  serial: i32,
  kind: u8,
  tag: u8,
  value: f64,
  rangeIndex: u32,
  rangeOffsets: Uint32Array,
  rangeLengths: Uint32Array,
  rangeMembers: Uint32Array,
  cellTags: Uint8Array,
  cellNumbers: Float64Array,
  cellStringIds: Uint32Array,
  cellErrors: Uint16Array,
): i32 {
  if (kind == STACK_KIND_SCALAR) {
    if (tag == ValueTag.Error) {
      return -1
    }
    const holiday = truncToInt(tag, value)
    return holiday == i32.MIN_VALUE ? -1 : holiday == serial ? 1 : 0
  }
  if (kind != STACK_KIND_RANGE) {
    return 0
  }
  const start = <i32>rangeOffsets[rangeIndex]
  const length = <i32>rangeLengths[rangeIndex]
  for (let index = 0; index < length; index += 1) {
    const memberIndex = rangeMembers[start + index]
    if (cellTags[memberIndex] == ValueTag.Error) {
      return -1
    }
    const serialCandidate = truncToInt(
      cellTags[memberIndex],
      memberScalarValue(memberIndex, cellTags, cellNumbers, cellStringIds, cellErrors),
    )
    if (serialCandidate == i32.MIN_VALUE) {
      return -1
    }
    if (serialCandidate == serial) {
      return 1
    }
  }
  return 0
}

export function isWorkdaySerial(
  serial: i32,
  kind: u8,
  tag: u8,
  value: f64,
  rangeIndex: u32,
  rangeOffsets: Uint32Array,
  rangeLengths: Uint32Array,
  rangeMembers: Uint32Array,
  cellTags: Uint8Array,
  cellNumbers: Float64Array,
  cellStringIds: Uint32Array,
  cellErrors: Uint16Array,
): i32 {
  if (isWeekendSerial(serial)) {
    return 0
  }
  if (kind == STACK_KIND_SCALAR && tag == ValueTag.Empty) {
    return 1
  }
  if (kind != STACK_KIND_SCALAR && kind != STACK_KIND_RANGE) {
    return 1
  }
  const holiday = isHolidaySerial(
    serial,
    kind,
    tag,
    value,
    rangeIndex,
    rangeOffsets,
    rangeLengths,
    rangeMembers,
    cellTags,
    cellNumbers,
    cellStringIds,
    cellErrors,
  )
  if (holiday < 0) {
    return -1
  }
  return holiday == 1 ? 0 : 1
}

export function isWorkdaySerialWithWeekendMask(
  serial: i32,
  weekendMask: i32,
  kind: u8,
  tag: u8,
  value: f64,
  rangeIndex: u32,
  rangeOffsets: Uint32Array,
  rangeLengths: Uint32Array,
  rangeMembers: Uint32Array,
  cellTags: Uint8Array,
  cellNumbers: Float64Array,
  cellStringIds: Uint32Array,
  cellErrors: Uint16Array,
): i32 {
  if (isWeekendWithMask(serial, weekendMask)) {
    return 0
  }
  if (kind == STACK_KIND_SCALAR && tag == ValueTag.Empty) {
    return 1
  }
  if (kind != STACK_KIND_SCALAR && kind != STACK_KIND_RANGE) {
    return 1
  }
  const holiday = isHolidaySerial(
    serial,
    kind,
    tag,
    value,
    rangeIndex,
    rangeOffsets,
    rangeLengths,
    rangeMembers,
    cellTags,
    cellNumbers,
    cellStringIds,
    cellErrors,
  )
  if (holiday < 0) {
    return -1
  }
  return holiday == 1 ? 0 : 1
}
