import { BuiltinId, ErrorCode, ValueTag } from "./protocol";
import {
  cumulativePeriodicPaymentCalc,
  futureValueCalc,
  periodicPaymentCalc,
  presentValueCalc,
  principalPaymentCalc,
  interestPaymentCalc,
  solveRateCalc,
  totalPeriodsCalc,
} from "./cashflows";
import { truncToInt } from "./numeric-core";
import { toNumberExact } from "./operands";
import { collectNumericValuesFromArgs, orderStatisticErrorCode } from "./statistics-tests";
import { paymentType, scalarErrorAt } from "./builtin-args";
import { STACK_KIND_SCALAR, writeResult } from "./result-io";

export function tryApplyFinanceCashflowBuiltin(
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
  if (builtinId == BuiltinId.Pv && argc >= 3 && argc <= 5) {
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
    const rate = toNumberExact(tagStack[base], valueStack[base]);
    const periods = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
    const payment = toNumberExact(tagStack[base + 2], valueStack[base + 2]);
    const future = argc >= 4 ? toNumberExact(tagStack[base + 3], valueStack[base + 3]) : 0.0;
    const paymentTypeValue =
      argc >= 5 ? paymentType(tagStack[base + 4], valueStack[base + 4], true) : 0;
    const present =
      paymentTypeValue < 0
        ? NaN
        : presentValueCalc(rate, periods, payment, future, paymentTypeValue);
    return isNaN(present)
      ? writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          ErrorCode.Value,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        )
      : writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Number,
          present,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
  }

  if (builtinId == BuiltinId.Pmt && argc >= 3 && argc <= 5) {
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
    const rate = toNumberExact(tagStack[base], valueStack[base]);
    const periods = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
    const present = toNumberExact(tagStack[base + 2], valueStack[base + 2]);
    const future = argc >= 4 ? toNumberExact(tagStack[base + 3], valueStack[base + 3]) : 0.0;
    const paymentTypeValue =
      argc >= 5 ? paymentType(tagStack[base + 4], valueStack[base + 4], true) : 0;
    const payment =
      paymentTypeValue < 0
        ? NaN
        : periodicPaymentCalc(rate, periods, present, future, paymentTypeValue);
    return isNaN(payment)
      ? writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          ErrorCode.Value,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        )
      : writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Number,
          payment,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
  }

  if (builtinId == BuiltinId.Nper && argc >= 3 && argc <= 5) {
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
    const rate = toNumberExact(tagStack[base], valueStack[base]);
    const payment = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
    const present = toNumberExact(tagStack[base + 2], valueStack[base + 2]);
    const future = argc >= 4 ? toNumberExact(tagStack[base + 3], valueStack[base + 3]) : 0.0;
    const paymentTypeValue =
      argc >= 5 ? paymentType(tagStack[base + 4], valueStack[base + 4], true) : 0;
    const periods =
      paymentTypeValue < 0
        ? NaN
        : totalPeriodsCalc(rate, payment, present, future, paymentTypeValue);
    return isNaN(periods)
      ? writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          ErrorCode.Value,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        )
      : writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Number,
          periods,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
  }

  if (builtinId == BuiltinId.Npv && argc >= 2) {
    const rate = toNumberExact(tagStack[base], valueStack[base]);
    if (isNaN(rate)) {
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
    const values = collectNumericValuesFromArgs(
      base + 1,
      argc - 1,
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
    );
    if (values === null) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        orderStatisticErrorCode,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    if (values.length == 0) {
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
    let total = 0.0;
    for (let index = 0; index < values.length; index += 1) {
      total += unchecked(values[index]) / Math.pow(1.0 + rate, <f64>(index + 1));
    }
    return !isFinite(total)
      ? writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          ErrorCode.Value,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        )
      : writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Number,
          total,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
  }

  if (builtinId == BuiltinId.Rate && argc >= 3 && argc <= 6) {
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
    const periods = toNumberExact(tagStack[base], valueStack[base]);
    const payment = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
    const present = toNumberExact(tagStack[base + 2], valueStack[base + 2]);
    const future = argc >= 4 ? toNumberExact(tagStack[base + 3], valueStack[base + 3]) : 0.0;
    const paymentTypeValue =
      argc >= 5 ? paymentType(tagStack[base + 4], valueStack[base + 4], true) : 0;
    const guess = argc >= 6 ? toNumberExact(tagStack[base + 5], valueStack[base + 5]) : 0.1;
    const rate =
      paymentTypeValue < 0
        ? NaN
        : solveRateCalc(periods, payment, present, future, paymentTypeValue, guess);
    return isNaN(rate)
      ? writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          ErrorCode.Value,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        )
      : writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Number,
          rate,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
  }

  if ((builtinId == BuiltinId.Ipmt || builtinId == BuiltinId.Ppmt) && argc >= 4 && argc <= 6) {
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
    const rate = toNumberExact(tagStack[base], valueStack[base]);
    const period = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
    const periods = toNumberExact(tagStack[base + 2], valueStack[base + 2]);
    const present = toNumberExact(tagStack[base + 3], valueStack[base + 3]);
    const future = argc >= 5 ? toNumberExact(tagStack[base + 4], valueStack[base + 4]) : 0.0;
    const paymentTypeValue =
      argc >= 6 ? paymentType(tagStack[base + 5], valueStack[base + 5], true) : 0;
    const result =
      paymentTypeValue < 0
        ? NaN
        : builtinId == BuiltinId.Ipmt
          ? interestPaymentCalc(rate, period, periods, present, future, paymentTypeValue)
          : principalPaymentCalc(rate, period, periods, present, future, paymentTypeValue);
    return isNaN(result)
      ? writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          ErrorCode.Value,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        )
      : writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Number,
          result,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
  }

  if (builtinId == BuiltinId.Ispmt && argc == 4) {
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
    const rate = toNumberExact(tagStack[base], valueStack[base]);
    const period = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
    const periods = toNumberExact(tagStack[base + 2], valueStack[base + 2]);
    const present = toNumberExact(tagStack[base + 3], valueStack[base + 3]);
    if (
      isNaN(rate) ||
      isNaN(period) ||
      isNaN(periods) ||
      isNaN(present) ||
      periods <= 0.0 ||
      period < 1.0 ||
      period > periods
    ) {
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
      present * rate * (period / periods - 1.0),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if ((builtinId == BuiltinId.Cumipmt || builtinId == BuiltinId.Cumprinc) && argc == 6) {
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
    const rate = toNumberExact(tagStack[base], valueStack[base]);
    const periods = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
    const present = toNumberExact(tagStack[base + 2], valueStack[base + 2]);
    const startPeriod = truncToInt(tagStack[base + 3], valueStack[base + 3]);
    const endPeriod = truncToInt(tagStack[base + 4], valueStack[base + 4]);
    const paymentTypeValue = paymentType(tagStack[base + 5], valueStack[base + 5], true);
    const total =
      startPeriod == i32.MIN_VALUE || endPeriod == i32.MIN_VALUE || paymentTypeValue < 0
        ? NaN
        : cumulativePeriodicPaymentCalc(
            rate,
            periods,
            present,
            startPeriod,
            endPeriod,
            paymentTypeValue,
            builtinId == BuiltinId.Cumprinc,
          );
    return isNaN(total)
      ? writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          ErrorCode.Value,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        )
      : writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Number,
          total,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
  }

  if (builtinId == BuiltinId.Fv && argc >= 3 && argc <= 5) {
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
    const rate = toNumberExact(tagStack[base], valueStack[base]);
    const periods = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
    const payment = toNumberExact(tagStack[base + 2], valueStack[base + 2]);
    const present = argc >= 4 ? toNumberExact(tagStack[base + 3], valueStack[base + 3]) : 0.0;
    const paymentTypeValue =
      argc >= 5 ? paymentType(tagStack[base + 4], valueStack[base + 4], true) : 0;
    const future =
      paymentTypeValue < 0
        ? NaN
        : futureValueCalc(rate, periods, payment, present, paymentTypeValue);
    return isNaN(future)
      ? writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          ErrorCode.Value,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        )
      : writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Number,
          future,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
  }

  if (builtinId == BuiltinId.Fvschedule && argc >= 2) {
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
    const principal = toNumberExact(tagStack[base], valueStack[base]);
    if (isNaN(principal)) {
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
    let result = principal;
    for (let index = 1; index < argc; index += 1) {
      const rate = toNumberExact(tagStack[base + index], valueStack[base + index]);
      if (isNaN(rate)) {
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
      result *= 1.0 + rate;
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      result,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if ((builtinId == BuiltinId.Effect || builtinId == BuiltinId.Nominal) && argc == 2) {
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
    const rate = toNumberExact(tagStack[base], valueStack[base]);
    const periodsNumeric = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
    const periods = Math.trunc(periodsNumeric);
    if (
      isNaN(rate) ||
      isNaN(periodsNumeric) ||
      !isFinite(periods) ||
      periods < 1.0 ||
      (builtinId == BuiltinId.Nominal && rate <= -1.0)
    ) {
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
    const result =
      builtinId == BuiltinId.Effect
        ? Math.pow(1.0 + rate / periods, periods) - 1.0
        : periods * (Math.pow(1.0 + rate, 1.0 / periods) - 1.0);
    return !isFinite(result)
      ? writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          ErrorCode.Value,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        )
      : writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Number,
          result,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
  }

  if (builtinId == BuiltinId.Pduration && argc == 3) {
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
    const rate = toNumberExact(tagStack[base], valueStack[base]);
    const present = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
    const future = toNumberExact(tagStack[base + 2], valueStack[base + 2]);
    const result =
      isNaN(rate) ||
      isNaN(present) ||
      isNaN(future) ||
      rate <= 0.0 ||
      present <= 0.0 ||
      future <= 0.0
        ? NaN
        : Math.log(future / present) / Math.log(1.0 + rate);
    return !isFinite(result)
      ? writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          ErrorCode.Value,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        )
      : writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Number,
          result,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
  }

  if (builtinId == BuiltinId.Rri && argc == 3) {
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
    const periods = toNumberExact(tagStack[base], valueStack[base]);
    const present = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
    const future = toNumberExact(tagStack[base + 2], valueStack[base + 2]);
    const result =
      isNaN(periods) || isNaN(present) || isNaN(future) || periods <= 0.0 || present == 0.0
        ? NaN
        : Math.pow(future / present, 1.0 / periods) - 1.0;
    return !isFinite(result)
      ? writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          ErrorCode.Value,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        )
      : writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Number,
          result,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
  }

  return -1;
}
