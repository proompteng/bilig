import { BuiltinId, ErrorCode, ValueTag } from "./protocol";
import {
  collectDateCellRangeSeriesFromSlot,
  collectNumericCellRangeSeriesFromSlot,
  sampleCollectionErrorCode,
} from "./statistics-tests";
import {
  hasPositiveAndNegativeSeries,
  mirrCalc,
  solvePeriodicCashflowRateCalc,
  solveXirrCalc,
  xnpvCalc,
} from "./cashflows";
import { STACK_KIND_RANGE, STACK_KIND_SCALAR, writeResult } from "./result-io";
import { toNumberExact, toNumberOrZero } from "./operands";
import { nextVolatileRandomValue, readVolatileNowSerial } from "./vm";

function writeValueError(
  base: i32,
  code: f64,
  rangeIndexStack: Uint32Array,
  valueStack: Float64Array,
  tagStack: Uint8Array,
  kindStack: Uint8Array,
): i32 {
  return writeResult(
    base,
    STACK_KIND_SCALAR,
    <u8>ValueTag.Error,
    code,
    rangeIndexStack,
    valueStack,
    tagStack,
    kindStack,
  );
}

function writeNumericResult(
  base: i32,
  value: f64,
  rangeIndexStack: Uint32Array,
  valueStack: Float64Array,
  tagStack: Uint8Array,
  kindStack: Uint8Array,
): i32 {
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

export function tryApplySpecialRuntimeBuiltin(
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
  if (
    ((builtinId == BuiltinId.Irr || builtinId == BuiltinId.Mirr) &&
      (argc == 1 || argc == 3 || argc == 2)) ||
    (builtinId == BuiltinId.Xnpv && argc == 3) ||
    (builtinId == BuiltinId.Xirr && (argc == 2 || argc == 3))
  ) {
    if (
      (builtinId == BuiltinId.Irr && (argc == 1 || argc == 2)) ||
      (builtinId == BuiltinId.Mirr && argc == 3)
    ) {
      const values = collectNumericCellRangeSeriesFromSlot(
        base,
        kindStack,
        tagStack,
        valueStack,
        rangeIndexStack,
        rangeOffsets,
        rangeLengths,
        rangeMembers,
        cellTags,
        cellNumbers,
        cellErrors,
        false,
      );
      if (values === null) {
        return writeValueError(
          base,
          sampleCollectionErrorCode,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }
      if (!hasPositiveAndNegativeSeries(values)) {
        return writeValueError(
          base,
          builtinId == BuiltinId.Mirr ? ErrorCode.Div0 : ErrorCode.Value,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }
      if (builtinId == BuiltinId.Irr) {
        if (argc == 2 && kindStack[base + 1] != STACK_KIND_SCALAR) {
          return writeValueError(
            base,
            ErrorCode.Value,
            rangeIndexStack,
            valueStack,
            tagStack,
            kindStack,
          );
        }
        const guess = argc == 2 ? toNumberExact(tagStack[base + 1], valueStack[base + 1]) : 0.1;
        const result = isNaN(guess) ? NaN : solvePeriodicCashflowRateCalc(values, guess);
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          isNaN(result) ? <u8>ValueTag.Error : <u8>ValueTag.Number,
          isNaN(result) ? ErrorCode.Value : result,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }

      const financeRate =
        kindStack[base + 1] == STACK_KIND_SCALAR
          ? toNumberExact(tagStack[base + 1], valueStack[base + 1])
          : NaN;
      const reinvestRate =
        kindStack[base + 2] == STACK_KIND_SCALAR
          ? toNumberExact(tagStack[base + 2], valueStack[base + 2])
          : NaN;
      const result = mirrCalc(values, financeRate, reinvestRate);
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        isNaN(result) ? <u8>ValueTag.Error : <u8>ValueTag.Number,
        isNaN(result)
          ? isNaN(financeRate) || isNaN(reinvestRate)
            ? ErrorCode.Value
            : ErrorCode.Div0
          : result,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    if (
      (builtinId == BuiltinId.Xnpv && kindStack[base] != STACK_KIND_SCALAR) ||
      (builtinId == BuiltinId.Xirr && kindStack[base] == STACK_KIND_SCALAR)
    ) {
      return writeValueError(
        base,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const firstNumeric =
      builtinId == BuiltinId.Xnpv ? toNumberExact(tagStack[base], valueStack[base]) : NaN;
    const guess =
      builtinId == BuiltinId.Xirr
        ? argc == 3
          ? kindStack[base + 2] == STACK_KIND_SCALAR
            ? toNumberExact(tagStack[base + 2], valueStack[base + 2])
            : NaN
          : 0.1
        : NaN;
    if (
      (builtinId == BuiltinId.Xnpv && isNaN(firstNumeric)) ||
      (builtinId == BuiltinId.Xirr && isNaN(guess))
    ) {
      return writeValueError(
        base,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const valuesSlot = builtinId == BuiltinId.Xnpv ? base + 1 : base;
    const datesSlot = builtinId == BuiltinId.Xnpv ? base + 2 : base + 1;
    const values = collectNumericCellRangeSeriesFromSlot(
      valuesSlot,
      kindStack,
      tagStack,
      valueStack,
      rangeIndexStack,
      rangeOffsets,
      rangeLengths,
      rangeMembers,
      cellTags,
      cellNumbers,
      cellErrors,
      true,
    );
    if (values === null) {
      return writeValueError(
        base,
        sampleCollectionErrorCode,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const dates = collectDateCellRangeSeriesFromSlot(
      datesSlot,
      kindStack,
      tagStack,
      valueStack,
      rangeIndexStack,
      rangeOffsets,
      rangeLengths,
      rangeMembers,
      cellTags,
      cellNumbers,
      cellErrors,
    );
    if (dates === null) {
      return writeValueError(
        base,
        sampleCollectionErrorCode,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    if (
      values.length != dates.length ||
      values.length == 0 ||
      !hasPositiveAndNegativeSeries(values)
    ) {
      return writeValueError(
        base,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const start = unchecked(dates[0]);
    for (let index = 0; index < dates.length; index += 1) {
      if (unchecked(dates[index]) < start) {
        return writeValueError(
          base,
          ErrorCode.Value,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }
    }
    const result =
      builtinId == BuiltinId.Xnpv
        ? xnpvCalc(firstNumeric, values, dates)
        : solveXirrCalc(values, dates, guess);
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      isNaN(result) ? <u8>ValueTag.Error : <u8>ValueTag.Number,
      isNaN(result) ? ErrorCode.Value : result,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Today || builtinId == BuiltinId.Now) {
    if (argc != 0) {
      return writeValueError(
        base,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const nowSerial = readVolatileNowSerial();
    if (isNaN(nowSerial)) {
      return writeValueError(
        base,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    return writeNumericResult(
      base,
      builtinId == BuiltinId.Today ? Math.floor(nowSerial) : nowSerial,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Rand) {
    if (argc != 0) {
      return writeValueError(
        base,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const next = nextVolatileRandomValue();
    if (!isFinite(next)) {
      return writeValueError(
        base,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    return writeNumericResult(
      base,
      Math.min(Math.max(next, 0), 1 - f64.EPSILON),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Sumproduct) {
    if (argc == 0) {
      return writeValueError(
        base,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const firstRangeIndex = rangeIndexStack[base];
    if (kindStack[base] != STACK_KIND_RANGE) {
      return writeValueError(
        base,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const expectedLength = <i32>rangeLengths[firstRangeIndex];
    for (let index = 0; index < argc; index++) {
      const slot = base + index;
      if (
        kindStack[slot] != STACK_KIND_RANGE ||
        <i32>rangeLengths[rangeIndexStack[slot]] != expectedLength
      ) {
        return writeValueError(
          base,
          ErrorCode.Value,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }
    }

    let sum = 0.0;
    for (let row = 0; row < expectedLength; row++) {
      let product = 1.0;
      for (let index = 0; index < argc; index++) {
        const rangeIndex = rangeIndexStack[base + index];
        const memberIndex = rangeMembers[rangeOffsets[rangeIndex] + row];
        product *= toNumberOrZero(cellTags[memberIndex], cellNumbers[memberIndex]);
      }
      sum += product;
    }
    return writeNumericResult(base, sum, rangeIndexStack, valueStack, tagStack, kindStack);
  }

  return -1;
}
