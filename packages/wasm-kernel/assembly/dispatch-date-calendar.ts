import { BuiltinId, ErrorCode, ValueTag } from "./protocol";
import { scalarErrorAt } from "./builtin-args";
import {
  coerceWeekendMask,
  isWorkdaySerial,
  isWorkdaySerialWithWeekendMask,
} from "./calendar-workdays";
import {
  excelDays360Value,
  excelSerialWhole,
  excelWeeknumFromSerial,
  excelYearfracValue,
} from "./date-finance";
import { truncToInt } from "./numeric-core";
import { STACK_KIND_SCALAR, writeResult } from "./result-io";

export function tryApplyDateCalendarBuiltin(
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
  rangeMembers: Uint32Array,
  outputStringOffsets: Uint32Array,
  outputStringLengths: Uint32Array,
  outputStringData: Uint16Array,
): i32 {
  if (builtinId == BuiltinId.Days && argc == 2) {
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack);
    if (scalarError >= 0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        scalarError,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const end = excelSerialWhole(tagStack[base], valueStack[base]);
    const start = excelSerialWhole(tagStack[base + 1], valueStack[base + 1]);
    if (end == i32.MIN_VALUE || start == i32.MIN_VALUE) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      <f64>(end - start),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Days360 && (argc == 2 || argc == 3)) {
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack);
    if (scalarError >= 0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        scalarError,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const startWhole = excelSerialWhole(tagStack[base], valueStack[base]);
    const endWhole = excelSerialWhole(tagStack[base + 1], valueStack[base + 1]);
    const method = argc == 3 ? truncToInt(tagStack[base + 2], valueStack[base + 2]) : 0;
    const value =
      startWhole == i32.MIN_VALUE || endWhole == i32.MIN_VALUE || (method != 0 && method != 1)
        ? NaN
        : excelDays360Value(startWhole, endWhole, method);
    if (isNaN(value)) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      value,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Yearfrac && (argc == 2 || argc == 3)) {
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack);
    if (scalarError >= 0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        scalarError,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const startWhole = excelSerialWhole(tagStack[base], valueStack[base]);
    const endWhole = excelSerialWhole(tagStack[base + 1], valueStack[base + 1]);
    const basis = argc == 3 ? truncToInt(tagStack[base + 2], valueStack[base + 2]) : 0;
    const value =
      startWhole == i32.MIN_VALUE || endWhole == i32.MIN_VALUE || basis < 0 || basis > 4
        ? NaN
        : excelYearfracValue(startWhole, endWhole, basis);
    if (isNaN(value)) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      value,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Weeknum && (argc == 1 || argc == 2)) {
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack);
    if (scalarError >= 0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        scalarError,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const returnType = argc == 2 ? truncToInt(tagStack[base + 1], valueStack[base + 1]) : 1;
    const weeknum =
      returnType == i32.MIN_VALUE
        ? i32.MIN_VALUE
        : excelWeeknumFromSerial(tagStack[base], valueStack[base], returnType);
    if (weeknum == i32.MIN_VALUE) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      <f64>weeknum,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Workday && (argc == 2 || argc == 3)) {
    const scalarError = scalarErrorAt(base, min<i32>(argc, 2), kindStack, tagStack, valueStack);
    if (scalarError >= 0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        scalarError,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const start = excelSerialWhole(tagStack[base], valueStack[base]);
    const offset = truncToInt(tagStack[base + 1], valueStack[base + 1]);
    if (start == i32.MIN_VALUE || offset == i32.MIN_VALUE) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    const holidayKind = argc == 3 ? kindStack[base + 2] : STACK_KIND_SCALAR;
    const holidayTag = argc == 3 ? tagStack[base + 2] : <u8>ValueTag.Empty;
    const holidayValue = argc == 3 ? valueStack[base + 2] : 0.0;
    const holidayRangeIndex = argc == 3 ? rangeIndexStack[base + 2] : 0;

    let cursor = start;
    const direction = offset >= 0 ? 1 : -1;
    while (true) {
      const workday = isWorkdaySerial(
        cursor,
        holidayKind,
        holidayTag,
        holidayValue,
        holidayRangeIndex,
        rangeOffsets,
        rangeLengths,
        rangeMembers,
        cellTags,
        cellNumbers,
        cellStringIds,
        cellErrors,
      );
      if (workday < 0) {
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          ErrorCode.Value,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }
      if (workday == 1) {
        break;
      }
      cursor += direction;
    }

    let remaining = offset >= 0 ? offset : -offset;
    while (remaining > 0) {
      cursor += direction;
      const workday = isWorkdaySerial(
        cursor,
        holidayKind,
        holidayTag,
        holidayValue,
        holidayRangeIndex,
        rangeOffsets,
        rangeLengths,
        rangeMembers,
        cellTags,
        cellNumbers,
        cellStringIds,
        cellErrors,
      );
      if (workday < 0) {
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          ErrorCode.Value,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }
      if (workday == 1) {
        remaining -= 1;
      }
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      <f64>cursor,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Networkdays && (argc == 2 || argc == 3)) {
    const scalarError = scalarErrorAt(base, min<i32>(argc, 2), kindStack, tagStack, valueStack);
    if (scalarError >= 0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        scalarError,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const start = excelSerialWhole(tagStack[base], valueStack[base]);
    const end = excelSerialWhole(tagStack[base + 1], valueStack[base + 1]);
    if (start == i32.MIN_VALUE || end == i32.MIN_VALUE) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    const holidayKind = argc == 3 ? kindStack[base + 2] : STACK_KIND_SCALAR;
    const holidayTag = argc == 3 ? tagStack[base + 2] : <u8>ValueTag.Empty;
    const holidayValue = argc == 3 ? valueStack[base + 2] : 0.0;
    const holidayRangeIndex = argc == 3 ? rangeIndexStack[base + 2] : 0;

    const step = start <= end ? 1 : -1;
    let count = 0;
    for (let cursor = start; ; cursor += step) {
      const workday = isWorkdaySerial(
        cursor,
        holidayKind,
        holidayTag,
        holidayValue,
        holidayRangeIndex,
        rangeOffsets,
        rangeLengths,
        rangeMembers,
        cellTags,
        cellNumbers,
        cellStringIds,
        cellErrors,
      );
      if (workday < 0) {
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          ErrorCode.Value,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }
      if (workday == 1) {
        count += step;
      }
      if (cursor == end) {
        break;
      }
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      <f64>count,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.WorkdayIntl && argc >= 2 && argc <= 4) {
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack);
    if (scalarError >= 0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        scalarError,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const start = excelSerialWhole(tagStack[base], valueStack[base]);
    const offset = truncToInt(tagStack[base + 1], valueStack[base + 1]);
    const weekendMask = coerceWeekendMask(
      argc >= 3,
      argc >= 3 ? tagStack[base + 2] : <u8>ValueTag.Empty,
      argc >= 3 ? valueStack[base + 2] : 0.0,
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    );
    if (start == i32.MIN_VALUE || offset == i32.MIN_VALUE || weekendMask == i32.MIN_VALUE) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    const holidayKind = argc == 4 ? kindStack[base + 3] : STACK_KIND_SCALAR;
    const holidayTag = argc == 4 ? tagStack[base + 3] : <u8>ValueTag.Empty;
    const holidayValue = argc == 4 ? valueStack[base + 3] : 0.0;
    const holidayRangeIndex = argc == 4 ? rangeIndexStack[base + 3] : 0;

    let cursor = start;
    const direction = offset >= 0 ? 1 : -1;
    while (true) {
      const workday = isWorkdaySerialWithWeekendMask(
        cursor,
        weekendMask,
        holidayKind,
        holidayTag,
        holidayValue,
        holidayRangeIndex,
        rangeOffsets,
        rangeLengths,
        rangeMembers,
        cellTags,
        cellNumbers,
        cellStringIds,
        cellErrors,
      );
      if (workday < 0) {
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          ErrorCode.Value,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }
      if (workday == 1) {
        break;
      }
      cursor += direction;
    }

    let remaining = offset >= 0 ? offset : -offset;
    while (remaining > 0) {
      cursor += direction;
      const workday = isWorkdaySerialWithWeekendMask(
        cursor,
        weekendMask,
        holidayKind,
        holidayTag,
        holidayValue,
        holidayRangeIndex,
        rangeOffsets,
        rangeLengths,
        rangeMembers,
        cellTags,
        cellNumbers,
        cellStringIds,
        cellErrors,
      );
      if (workday < 0) {
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          ErrorCode.Value,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }
      if (workday == 1) {
        remaining -= 1;
      }
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      <f64>cursor,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.NetworkdaysIntl && argc >= 2 && argc <= 4) {
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack);
    if (scalarError >= 0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        scalarError,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const start = excelSerialWhole(tagStack[base], valueStack[base]);
    const end = excelSerialWhole(tagStack[base + 1], valueStack[base + 1]);
    const weekendMask = coerceWeekendMask(
      argc >= 3,
      argc >= 3 ? tagStack[base + 2] : <u8>ValueTag.Empty,
      argc >= 3 ? valueStack[base + 2] : 0.0,
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    );
    if (start == i32.MIN_VALUE || end == i32.MIN_VALUE || weekendMask == i32.MIN_VALUE) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    const holidayKind = argc == 4 ? kindStack[base + 3] : STACK_KIND_SCALAR;
    const holidayTag = argc == 4 ? tagStack[base + 3] : <u8>ValueTag.Empty;
    const holidayValue = argc == 4 ? valueStack[base + 3] : 0.0;
    const holidayRangeIndex = argc == 4 ? rangeIndexStack[base + 3] : 0;

    const step = start <= end ? 1 : -1;
    let count = 0;
    for (let cursor = start; ; cursor += step) {
      const workday = isWorkdaySerialWithWeekendMask(
        cursor,
        weekendMask,
        holidayKind,
        holidayTag,
        holidayValue,
        holidayRangeIndex,
        rangeOffsets,
        rangeLengths,
        rangeMembers,
        cellTags,
        cellNumbers,
        cellStringIds,
        cellErrors,
      );
      if (workday < 0) {
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          ErrorCode.Value,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }
      if (workday == 1) {
        count += step;
      }
      if (cursor == end) {
        break;
      }
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      <f64>count,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  return -1;
}
