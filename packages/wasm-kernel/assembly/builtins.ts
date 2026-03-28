import { BuiltinId, ErrorCode, ValueTag } from "./protocol";
import {
  getTrackedArrayCols as getDynamicArrayCols,
  getTrackedArrayRows as getDynamicArrayRows,
  registerTrackedArrayShape as registerTrackedArrayShapeImpl,
} from "./dynamic-arrays";
import { scalarText, trimAsciiWhitespace } from "./text-codec";
import { arrayToTextCell, parseNumericText } from "./text-special";
import { coerceLength, coercePositiveStart } from "./text-foundation";
import { formatFixedText } from "./text-format";
import { compareScalarValues, valueNumber } from "./comparison";
import { matchesCriteriaValue } from "./criteria";
import { databaseBuiltinResult } from "./database-criteria";
import { tryApplyAggregateCriteriaBuiltin } from "./dispatch-aggregate-criteria";
import { tryApplyTextMutationBuiltin } from "./dispatch-text-mutation";
import { tryApplyTextFormattingBuiltin } from "./dispatch-text-formatting";
import { tryApplyDateTimeBuiltin } from "./dispatch-date-time";
import { tryApplyDateCalendarBuiltin } from "./dispatch-date-calendar";
import { tryApplyDepreciationBuiltin } from "./dispatch-depreciation";
import { tryApplyFinanceCashflowBuiltin } from "./dispatch-finance-cashflows";
import { tryApplyFinanceSecuritiesBuiltin } from "./dispatch-finance-securities";
import {
  copyInputCellToSpill,
  materializeSlotResult,
  uniqueColKey,
  uniqueRowKey,
  uniqueScalarKey,
} from "./array-materialize";
import {
  coerceNonNegativeShift,
  coerceInteger,
  doubleFactorialCalc,
  evenCalc,
  factorialCalc,
  gcdPairCalc,
  lcmPairCalc,
  oddCalc,
  truncAbs,
  truncToInt,
} from "./numeric-core";
import {
  formatSignedRadixText,
  isValidBaseText,
  parseBaseText,
  parseSignedRadixText,
  toBaseText,
} from "./radix";
import {
  EXCEL_SECONDS_PER_DAY,
  accruedIssueYearfracValue,
  couponDateFromMaturityValue,
  couponDaysByBasisValue,
  couponPeriodDaysValue,
  couponPeriodsRemainingValue,
  couponPriceFromMetricsValue,
  dbDepreciation,
  ddbDepreciation,
  excelSerialWhole,
  macaulayDurationValue,
  maturityIssueYearfracValue,
  oddFirstPriceValue,
  oddFirstYieldValue,
  oddLastPriceValue,
  oddLastYieldValue,
  securityAnnualizedYearfracValue,
  solveCouponYieldValue,
  treasuryBillDaysValue,
  vdbDepreciation,
} from "./date-finance";
import {
  cumulativePeriodicPaymentCalc,
  futureValueCalc,
  hasPositiveAndNegativeSeries,
  interestPaymentCalc,
  mirrCalc,
  periodicCashflowNetPresentValueCalc,
  periodicPaymentCalc,
  presentValueCalc,
  principalPaymentCalc,
  solvePeriodicCashflowRateCalc,
  solveRateCalc,
  solveXirrCalc,
  totalPeriodsCalc,
  xnpvCalc,
} from "./cashflows";
import {
  betaDistributionCdf,
  betaDistributionDensity,
  betaDistributionInverse,
  besselIValue,
  besselJValue,
  besselKValue,
  besselYValue,
  binomialProbability,
  chiSquareCdf,
  chiSquareDensity,
  erfApprox,
  fDistributionCdf,
  fDistributionDensity,
  gammaDistributionCdf,
  gammaDistributionDensity,
  gammaFunction,
  hypergeometricProbability,
  inverseChiSquare,
  inverseFDistribution,
  inverseGammaDistribution,
  inverseStandardNormal,
  inverseStudentT,
  logGamma,
  negativeBinomialProbability,
  poissonProbability,
  regularizedUpperGamma,
  standardNormalCdf,
  standardNormalPdf,
  studentTCdf,
  studentTDensity,
} from "./distributions";
import {
  interpolateSortedPercentRank,
  interpolateSortedPercentile,
  kurtosisOf,
  meanOf,
  modeSingleOf,
  populationVarianceOf,
  sampleVarianceOf,
  skewPopulationOf,
  skewSampleOf,
  sortNumericValues,
  truncateToSignificance,
} from "./statistics-core";
import {
  CONVERT_GROUP_TEMPERATURE,
  convertKelvinToTemperature,
  convertTemperatureToKelvin,
  resolveConvertUnit,
  resolvedConvertFactor,
  resolvedConvertGroup,
  resolvedConvertTemperature,
} from "./unit-convert";
import {
  chiSquareTestPValue,
  collectDateCellRangeSeriesFromSlot,
  collectDateSeriesFromSlot,
  collectNumericCellRangeSeriesFromSlot,
  collectNumericSeriesFromSlot,
  collectNumericValuesFromArgs,
  collectNumericValuesFromSlot,
  collectPairedNumericStats,
  collectSampleNumbersFromSlot,
  fTestPValue,
  orderStatisticErrorCode,
  pairedCenteredCrossProducts,
  pairedCenteredSumSquaresX,
  pairedCenteredSumSquaresY,
  pairedSampleCount,
  pairedSumX,
  pairedSumY,
  sampleCollectionErrorCode,
  tTestPValue,
  zTestPValue,
} from "./statistics-tests";
import { tryApplyScalarTextBuiltin } from "./dispatch-text-scalar";
import {
  inputCellNumeric,
  inputCellScalarValue,
  inputCellTag,
  inputColsFromSlot,
  inputRowsFromSlot,
  memberScalarValue,
  rangeMemberAt,
  toNumberExact,
  toNumberOrNaN,
  toNumberOrZero,
} from "./operands";
import {
  coerceLogical,
  coerceNumberArg,
  coercePositiveIntegerArg,
  collectScalarStatValues,
  collectStatValuesFromArgs,
  isNumericResult,
  lastStatCollectionErrorCode,
  paymentType,
  rangeErrorAt,
  rangeSupportedScalarOnly,
  scalarArgsOnly,
  scalarErrorAt,
  statScalarValue,
} from "./builtin-args";
import {
  STACK_KIND_ARRAY,
  STACK_KIND_RANGE,
  STACK_KIND_SCALAR,
  UNRESOLVED_WASM_OPERAND,
  copySlotResult,
  vectorSlotLength,
  writeArrayResult,
  writeMemberResult,
  writeResult,
  writeStringResult,
} from "./result-io";
import {
  allocateSpillArrayResult,
  nextVolatileRandomValue,
  readSpillArrayTag,
  readSpillArrayLength,
  readSpillArrayNumber,
  readVolatileNowSerial,
  writeSpillArrayNumber,
  writeSpillArrayValue,
} from "./vm";
import { tryApplyArrayFoundationBuiltin } from "./dispatch-array-foundation";
import { tryApplyArrayInfoBuiltin } from "./dispatch-array-info";
import { tryApplyLookupMatchBuiltin } from "./dispatch-lookup-match";
import { tryApplyRegressionBuiltin } from "./dispatch-regression";
import { tryApplyScalarDistributionBuiltin } from "./dispatch-distributions-scalar";
import { tryApplyStatisticsSummaryBuiltin } from "./dispatch-statistics-summary";
import { tryApplyScalarMathBuiltin } from "./dispatch-scalar-math";
import { tryApplyLogicInfoBuiltin } from "./dispatch-logic-info";
import { tryApplyLookupTableBuiltin } from "./dispatch-lookup-table";

export function registerTrackedArrayShape(arrayIndex: u32, rows: i32, cols: i32): void {
  registerTrackedArrayShapeImpl(arrayIndex, rows, cols);
}

const MAX_SAFE_INTEGER_F64: f64 = 9007199254740991.0;

function volatileNowResult(): f64 {
  return readVolatileNowSerial();
}

function coerceBoolean(tag: u8, value: f64): i32 {
  if (tag == ValueTag.Boolean || tag == ValueTag.Number) {
    return value != 0 ? 1 : 0;
  }
  if (tag == ValueTag.Empty) {
    return 0;
  }
  return -1;
}

function columnLabelText(column: i32): string | null {
  if (column < 1) {
    return null;
  }
  let current = column;
  let label = "";
  while (current > 0) {
    const offset = (current - 1) % 26;
    label = String.fromCharCode(65 + offset) + label;
    current = (current - 1) / 26;
  }
  return label;
}

function escapeSheetNameText(value: string): string {
  let output = "";
  for (let index = 0; index < value.length; index++) {
    const char = value.charAt(index);
    output += char;
    if (char == "'") {
      output += "'";
    }
  }
  return output;
}

function parseSignedIntegerText(text: string): i32 {
  if (text.length == 0) {
    return i32.MIN_VALUE;
  }
  let index = 0;
  let sign = 1;
  if (text.charAt(0) == "+") {
    index = 1;
  } else if (text.charAt(0) == "-") {
    sign = -1;
    index = 1;
  }
  if (index >= text.length) {
    return i32.MIN_VALUE;
  }
  let value = 0;
  for (; index < text.length; index += 1) {
    const code = text.charCodeAt(index);
    if (code < 48 || code > 57) {
      return i32.MIN_VALUE;
    }
    value = value * 10 + (code - 48);
  }
  return sign * value;
}

function digitCount(value: i32): i32 {
  if (value <= 0) {
    return 1;
  }
  let current = value;
  let count = 0;
  while (current > 0) {
    count += 1;
    current /= 10;
  }
  return count;
}

function isValidDollarFractionNative(fraction: i32): bool {
  if (fraction <= 0) {
    return false;
  }
  if (fraction == 1) {
    return true;
  }
  let current = fraction;
  while ((current & 1) == 0) {
    current >>= 1;
  }
  return current == 1;
}

function parsePositiveDigits(value: string): i32 {
  if (value.length == 0) {
    return 0;
  }
  let output = 0;
  for (let index = 0; index < value.length; index++) {
    const code = value.charCodeAt(index);
    if (code < 48 || code > 57) {
      return i32.MIN_VALUE;
    }
    output = output * 10 + (code - 48);
  }
  return output;
}

function dollarFractionalNumerator(value: f64): i32 {
  const absoluteText = Math.abs(value).toString();
  const dot = absoluteText.indexOf(".");
  if (dot < 0) {
    return 0;
  }
  return parsePositiveDigits(absoluteText.substring(dot + 1));
}

function signedRadixInputText(
  tag: u8,
  value: f64,
  stringOffsets: Uint32Array,
  stringLengths: Uint32Array,
  stringData: Uint16Array,
  outputStringOffsets: Uint32Array,
  outputStringLengths: Uint32Array,
  outputStringData: Uint16Array,
): string | null {
  if (tag == ValueTag.Error) {
    return null;
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
    );
    if (text == null) {
      return null;
    }
    return trimAsciiWhitespace(text).toUpperCase();
  }
  const numeric = toNumberExact(tag, value);
  if (!isFinite(numeric)) {
    return null;
  }
  return (<i64>numeric).toString().toUpperCase();
}

function roundToPlacesNative(value: f64, places: i32): f64 {
  const scale = Math.pow(10.0, <f64>places);
  return Math.round(value * scale) / scale;
}

function roundToSignificantDigitsNative(value: f64, digits: i32): f64 {
  if (value == 0.0 || !isFinite(value)) {
    return value;
  }
  const exponent = <i32>Math.floor(Math.log(Math.abs(value)) / Math.log(10.0));
  const scale = Math.pow(10.0, <f64>(digits - exponent - 1));
  return Math.round(value * scale) / scale;
}

function euroRateNative(code: string): f64 {
  if (code == "BEF" || code == "LUF") return 40.3399;
  if (code == "DEM") return 1.95583;
  if (code == "ESP") return 166.386;
  if (code == "FRF") return 6.55957;
  if (code == "IEP") return 0.787564;
  if (code == "ITL") return 1936.27;
  if (code == "NLG") return 2.20371;
  if (code == "ATS") return 13.7603;
  if (code == "PTE") return 200.482;
  if (code == "FIM") return 5.94573;
  if (code == "GRD") return 340.75;
  if (code == "SIT") return 239.64;
  if (code == "EUR") return 1.0;
  return NaN;
}

function euroCalculationPrecisionNative(code: string): i32 {
  if (
    code == "BEF" ||
    code == "LUF" ||
    code == "ESP" ||
    code == "ITL" ||
    code == "PTE" ||
    code == "GRD"
  ) {
    return 0;
  }
  if (
    code == "DEM" ||
    code == "FRF" ||
    code == "IEP" ||
    code == "NLG" ||
    code == "ATS" ||
    code == "FIM" ||
    code == "SIT" ||
    code == "EUR"
  ) {
    return 2;
  }
  return i32.MIN_VALUE;
}

function coerceTrimMode(tag: u8, value: f64): i32 {
  const numeric = toNumberExact(tag, value);
  if (!isFinite(numeric)) {
    return i32.MIN_VALUE;
  }
  const integer = <i32>numeric;
  return integer >= 0 && integer <= 3 ? integer : i32.MIN_VALUE;
}

function clipIndex(index: i32, length: i32): i32 {
  if (length <= 0) {
    return i32.MIN_VALUE;
  }
  if (index == 0) {
    return i32.MIN_VALUE;
  }
  return index < 0 ? max(index, -length) : min(index, length);
}

function unresolvedRangeOperandError(
  base: i32,
  argc: i32,
  kindStack: Uint8Array,
  rangeIndexStack: Uint32Array,
): f64 {
  for (let index = 0; index < argc; index++) {
    const slot = base + index;
    if (kindStack[slot] == STACK_KIND_RANGE && rangeIndexStack[slot] == UNRESOLVED_WASM_OPERAND) {
      return ErrorCode.Ref;
    }
  }
  return -1;
}

export function applyBuiltin(
  builtinId: i32,
  argc: i32,
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
  sp: i32,
): i32 {
  const base = sp - argc;
  const unresolvedRangeError = unresolvedRangeOperandError(base, argc, kindStack, rangeIndexStack);
  if (unresolvedRangeError >= 0) {
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Error,
      unresolvedRangeError,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

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
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        sampleCollectionErrorCode,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    if (!hasPositiveAndNegativeSeries(values)) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        builtinId == BuiltinId.Mirr ? ErrorCode.Div0 : ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    if (builtinId == BuiltinId.Irr) {
      if (argc == 2 && kindStack[base + 1] != STACK_KIND_SCALAR) {
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
    (builtinId == BuiltinId.Xnpv && argc == 3) ||
    (builtinId == BuiltinId.Xirr && (argc == 2 || argc == 3))
  ) {
    if (
      (builtinId == BuiltinId.Xnpv && kindStack[base] != STACK_KIND_SCALAR) ||
      (builtinId == BuiltinId.Xirr && kindStack[base] == STACK_KIND_SCALAR)
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
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
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
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
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
    const start = unchecked(dates[0]);
    for (let index = 0; index < dates.length; index += 1) {
      if (unchecked(dates[index]) < start) {
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

  if ((builtinId == 233 || builtinId == 234 || builtinId == 235) && argc == 2) {
    const result = chiSquareTestPValue(
      base,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
      cellTags,
      cellNumbers,
      cellStringIds,
      cellErrors,
      rangeOffsets,
      rangeLengths,
      rangeRowCounts,
      rangeColCounts,
      rangeMembers,
    );
    if (result < 0.0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        -result,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      isNumericResult(result) ? <u8>ValueTag.Number : <u8>ValueTag.Error,
      isNumericResult(result) ? result : ErrorCode.Value,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if ((builtinId == 236 || builtinId == 237) && argc == 2) {
    const result = fTestPValue(
      base,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
      cellTags,
      cellNumbers,
      cellStringIds,
      cellErrors,
      rangeOffsets,
      rangeLengths,
      rangeRowCounts,
      rangeColCounts,
      rangeMembers,
    );
    if (result < 0.0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        -result,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      isNumericResult(result) ? <u8>ValueTag.Number : <u8>ValueTag.Error,
      isNumericResult(result) ? result : ErrorCode.Value,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if ((builtinId == 238 || builtinId == 239) && (argc == 2 || argc == 3)) {
    const result = zTestPValue(
      base,
      argc,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
      cellTags,
      cellNumbers,
      cellStringIds,
      cellErrors,
      rangeOffsets,
      rangeLengths,
      rangeRowCounts,
      rangeColCounts,
      rangeMembers,
    );
    if (result < 0.0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        -result,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      isNumericResult(result) ? <u8>ValueTag.Number : <u8>ValueTag.Error,
      isNumericResult(result) ? result : ErrorCode.Value,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if ((builtinId == BuiltinId.TTest || builtinId == BuiltinId.Ttest) && argc == 4) {
    const result = tTestPValue(
      base,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
      cellTags,
      cellNumbers,
      cellStringIds,
      cellErrors,
      rangeOffsets,
      rangeLengths,
      rangeRowCounts,
      rangeColCounts,
      rangeMembers,
    );
    if (result < 0.0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        -result,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      isNumericResult(result) ? <u8>ValueTag.Number : <u8>ValueTag.Error,
      isNumericResult(result) ? result : ErrorCode.Value,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Today) {
    if (argc != 0)
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
    const nowSerial = volatileNowResult();
    if (isNaN(nowSerial))
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
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      Math.floor(nowSerial),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Now) {
    if (argc != 0)
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
    const nowSerial = volatileNowResult();
    if (isNaN(nowSerial))
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
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      nowSerial,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Rand) {
    if (argc != 0)
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
    const next = nextVolatileRandomValue();
    if (!isFinite(next))
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
    const bounded = Math.min(Math.max(next, 0), 1 - f64.EPSILON);
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      bounded,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  const arrayFoundationResult = tryApplyArrayFoundationBuiltin(
    builtinId,
    argc,
    base,
    rangeIndexStack,
    valueStack,
    tagStack,
    kindStack,
    cellTags,
    cellNumbers,
    cellStringIds,
    cellErrors,
    rangeOffsets,
    rangeLengths,
    rangeRowCounts,
    rangeColCounts,
    rangeMembers,
  );
  if (arrayFoundationResult >= 0) {
    return arrayFoundationResult;
  }

  if (builtinId == BuiltinId.Filter && (argc == 2 || argc == 3)) {
    const sourceKind = kindStack[base];
    const includeKind = kindStack[base + 1];
    if (
      (sourceKind != STACK_KIND_RANGE && sourceKind != STACK_KIND_ARRAY) ||
      (includeKind != STACK_KIND_RANGE && includeKind != STACK_KIND_ARRAY)
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

    const sourceRows = inputRowsFromSlot(base, kindStack, rangeIndexStack, rangeRowCounts);
    const sourceCols = inputColsFromSlot(base, kindStack, rangeIndexStack, rangeColCounts);
    const includeRows = inputRowsFromSlot(base + 1, kindStack, rangeIndexStack, rangeRowCounts);
    const includeCols = inputColsFromSlot(base + 1, kindStack, rangeIndexStack, rangeColCounts);
    if (
      sourceRows <= 0 ||
      sourceCols <= 0 ||
      includeRows <= 0 ||
      includeCols <= 0 ||
      sourceRows == i32.MIN_VALUE ||
      sourceCols == i32.MIN_VALUE ||
      includeRows == i32.MIN_VALUE ||
      includeCols == i32.MIN_VALUE
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

    if (includeRows == sourceRows && includeCols == 1) {
      const keptRows = new Array<i32>();
      for (let row = 0; row < sourceRows; row++) {
        const includeTag = inputCellTag(
          base + 1,
          row,
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
        );
        const includeValue = inputCellScalarValue(
          base + 1,
          row,
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
        );
        if (includeTag == ValueTag.Error) {
          return writeResult(
            base,
            STACK_KIND_SCALAR,
            <u8>ValueTag.Error,
            includeValue,
            rangeIndexStack,
            valueStack,
            tagStack,
            kindStack,
          );
        }
        const keep = coerceBoolean(includeTag, includeValue);
        if (keep < 0) {
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
        if (keep == 1) {
          keptRows.push(row);
        }
      }
      if (keptRows.length == 0) {
        if (argc < 3 || kindStack[base + 2] != STACK_KIND_SCALAR) {
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
          tagStack[base + 2],
          valueStack[base + 2],
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }
      const arrayIndex = allocateSpillArrayResult(keptRows.length, sourceCols);
      let outputOffset = 0;
      for (let index = 0; index < keptRows.length; index++) {
        const sourceRow = unchecked(keptRows[index]);
        for (let col = 0; col < sourceCols; col++) {
          const copyError = copyInputCellToSpill(
            arrayIndex,
            outputOffset,
            base,
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
          );
          if (copyError != ErrorCode.None) {
            return writeResult(
              base,
              STACK_KIND_SCALAR,
              <u8>ValueTag.Error,
              copyError,
              rangeIndexStack,
              valueStack,
              tagStack,
              kindStack,
            );
          }
          outputOffset += 1;
        }
      }
      return writeArrayResult(
        base,
        arrayIndex,
        keptRows.length,
        sourceCols,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    if (includeRows == 1 && includeCols == sourceCols) {
      const keptCols = new Array<i32>();
      for (let col = 0; col < sourceCols; col++) {
        const includeTag = inputCellTag(
          base + 1,
          0,
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
        );
        const includeValue = inputCellScalarValue(
          base + 1,
          0,
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
        );
        if (includeTag == ValueTag.Error) {
          return writeResult(
            base,
            STACK_KIND_SCALAR,
            <u8>ValueTag.Error,
            includeValue,
            rangeIndexStack,
            valueStack,
            tagStack,
            kindStack,
          );
        }
        const keep = coerceBoolean(includeTag, includeValue);
        if (keep < 0) {
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
        if (keep == 1) {
          keptCols.push(col);
        }
      }
      if (keptCols.length == 0) {
        if (argc < 3 || kindStack[base + 2] != STACK_KIND_SCALAR) {
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
          tagStack[base + 2],
          valueStack[base + 2],
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }
      const arrayIndex = allocateSpillArrayResult(sourceRows, keptCols.length);
      let outputOffset = 0;
      for (let row = 0; row < sourceRows; row++) {
        for (let index = 0; index < keptCols.length; index++) {
          const sourceCol = unchecked(keptCols[index]);
          const copyError = copyInputCellToSpill(
            arrayIndex,
            outputOffset,
            base,
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
          );
          if (copyError != ErrorCode.None) {
            return writeResult(
              base,
              STACK_KIND_SCALAR,
              <u8>ValueTag.Error,
              copyError,
              rangeIndexStack,
              valueStack,
              tagStack,
              kindStack,
            );
          }
          outputOffset += 1;
        }
      }
      return writeArrayResult(
        base,
        arrayIndex,
        sourceRows,
        keptCols.length,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

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

  if (builtinId == BuiltinId.Unique && argc >= 1 && argc <= 3) {
    const sourceKind = kindStack[base];
    if (sourceKind != STACK_KIND_RANGE && sourceKind != STACK_KIND_ARRAY) {
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
    if (argc >= 2 && kindStack[base + 1] != STACK_KIND_SCALAR) {
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
    if (argc >= 3 && kindStack[base + 2] != STACK_KIND_SCALAR) {
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
    if (argc >= 2 && tagStack[base + 1] == ValueTag.Error) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        valueStack[base + 1],
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    if (argc >= 3 && tagStack[base + 2] == ValueTag.Error) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        valueStack[base + 2],
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    const byColFlag = argc >= 2 ? coerceBoolean(tagStack[base + 1], valueStack[base + 1]) : 0;
    const exactlyOnceFlag = argc >= 3 ? coerceBoolean(tagStack[base + 2], valueStack[base + 2]) : 0;
    if (byColFlag < 0 || exactlyOnceFlag < 0) {
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

    const sourceRows = inputRowsFromSlot(base, kindStack, rangeIndexStack, rangeRowCounts);
    const sourceCols = inputColsFromSlot(base, kindStack, rangeIndexStack, rangeColCounts);
    if (
      sourceRows <= 0 ||
      sourceCols <= 0 ||
      sourceRows == i32.MIN_VALUE ||
      sourceCols == i32.MIN_VALUE
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

    if (sourceRows == 1 || sourceCols == 1) {
      const vectorLength = sourceRows * sourceCols;
      const keys = new Array<string>();
      for (let index = 0; index < vectorLength; index++) {
        const row = sourceRows == 1 ? 0 : index;
        const col = sourceRows == 1 ? index : 0;
        const tag = inputCellTag(
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
        );
        const value = inputCellScalarValue(
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
        );
        if (tag == ValueTag.Error) {
          return writeResult(
            base,
            STACK_KIND_SCALAR,
            <u8>ValueTag.Error,
            value,
            rangeIndexStack,
            valueStack,
            tagStack,
            kindStack,
          );
        }
        const key = uniqueScalarKey(
          tag,
          value,
          stringOffsets,
          stringLengths,
          stringData,
          outputStringOffsets,
          outputStringLengths,
          outputStringData,
        );
        if (key == null) {
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
        keys.push(key);
      }

      const keptIndexes = new Array<i32>();
      for (let index = 0; index < vectorLength; index++) {
        const key = unchecked(keys[index]);
        let seenEarlier = false;
        for (let prior = 0; prior < index; prior++) {
          if (unchecked(keys[prior]) == key) {
            seenEarlier = true;
            break;
          }
        }
        if (seenEarlier) {
          continue;
        }
        if (exactlyOnceFlag == 1) {
          let count = 0;
          for (let cursor = 0; cursor < vectorLength; cursor++) {
            if (unchecked(keys[cursor]) == key) {
              count += 1;
            }
          }
          if (count != 1) {
            continue;
          }
        }
        keptIndexes.push(index);
      }

      const outputRows = sourceRows == 1 ? 1 : keptIndexes.length;
      const outputCols = sourceRows == 1 ? keptIndexes.length : 1;
      const arrayIndex = allocateSpillArrayResult(outputRows, outputCols);
      for (let outputOffset = 0; outputOffset < keptIndexes.length; outputOffset++) {
        const sourceIndex = unchecked(keptIndexes[outputOffset]);
        const sourceRow = sourceRows == 1 ? 0 : sourceIndex;
        const sourceCol = sourceRows == 1 ? sourceIndex : 0;
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
        );
        if (copyError != ErrorCode.None) {
          return writeResult(
            base,
            STACK_KIND_SCALAR,
            <u8>ValueTag.Error,
            copyError,
            rangeIndexStack,
            valueStack,
            tagStack,
            kindStack,
          );
        }
      }
      return writeArrayResult(
        base,
        arrayIndex,
        outputRows,
        outputCols,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    if (byColFlag == 1) {
      const keys = new Array<string>();
      for (let col = 0; col < sourceCols; col++) {
        const key = uniqueColKey(
          base,
          col,
          sourceRows,
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
          stringOffsets,
          stringLengths,
          stringData,
          outputStringOffsets,
          outputStringLengths,
          outputStringData,
        );
        if (key == null) {
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
        keys.push(key);
      }

      const keptCols = new Array<i32>();
      for (let col = 0; col < sourceCols; col++) {
        const key = unchecked(keys[col]);
        let seenEarlier = false;
        for (let prior = 0; prior < col; prior++) {
          if (unchecked(keys[prior]) == key) {
            seenEarlier = true;
            break;
          }
        }
        if (seenEarlier) {
          continue;
        }
        if (exactlyOnceFlag == 1) {
          let count = 0;
          for (let cursor = 0; cursor < sourceCols; cursor++) {
            if (unchecked(keys[cursor]) == key) {
              count += 1;
            }
          }
          if (count != 1) {
            continue;
          }
        }
        keptCols.push(col);
      }

      const arrayIndex = allocateSpillArrayResult(sourceRows, keptCols.length);
      let outputOffset = 0;
      for (let row = 0; row < sourceRows; row++) {
        for (let index = 0; index < keptCols.length; index++) {
          const sourceCol = unchecked(keptCols[index]);
          const copyError = copyInputCellToSpill(
            arrayIndex,
            outputOffset,
            base,
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
          );
          if (copyError != ErrorCode.None) {
            return writeResult(
              base,
              STACK_KIND_SCALAR,
              <u8>ValueTag.Error,
              copyError,
              rangeIndexStack,
              valueStack,
              tagStack,
              kindStack,
            );
          }
          outputOffset += 1;
        }
      }
      return writeArrayResult(
        base,
        arrayIndex,
        sourceRows,
        keptCols.length,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    const keys = new Array<string>();
    for (let row = 0; row < sourceRows; row++) {
      const key = uniqueRowKey(
        base,
        row,
        sourceCols,
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
        stringOffsets,
        stringLengths,
        stringData,
        outputStringOffsets,
        outputStringLengths,
        outputStringData,
      );
      if (key == null) {
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
      keys.push(key);
    }

    const keptRows = new Array<i32>();
    for (let row = 0; row < sourceRows; row++) {
      const key = unchecked(keys[row]);
      let seenEarlier = false;
      for (let prior = 0; prior < row; prior++) {
        if (unchecked(keys[prior]) == key) {
          seenEarlier = true;
          break;
        }
      }
      if (seenEarlier) {
        continue;
      }
      if (exactlyOnceFlag == 1) {
        let count = 0;
        for (let cursor = 0; cursor < sourceRows; cursor++) {
          if (unchecked(keys[cursor]) == key) {
            count += 1;
          }
        }
        if (count != 1) {
          continue;
        }
      }
      keptRows.push(row);
    }

    const arrayIndex = allocateSpillArrayResult(keptRows.length, sourceCols);
    let outputOffset = 0;
    for (let index = 0; index < keptRows.length; index++) {
      const sourceRow = unchecked(keptRows[index]);
      for (let col = 0; col < sourceCols; col++) {
        const copyError = copyInputCellToSpill(
          arrayIndex,
          outputOffset,
          base,
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
        );
        if (copyError != ErrorCode.None) {
          return writeResult(
            base,
            STACK_KIND_SCALAR,
            <u8>ValueTag.Error,
            copyError,
            rangeIndexStack,
            valueStack,
            tagStack,
            kindStack,
          );
        }
        outputOffset += 1;
      }
    }
    return writeArrayResult(
      base,
      arrayIndex,
      keptRows.length,
      sourceCols,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Offset && argc >= 3 && argc <= 5) {
    const sourceKind = kindStack[base];
    if (
      sourceKind != STACK_KIND_SCALAR &&
      sourceKind != STACK_KIND_RANGE &&
      sourceKind != STACK_KIND_ARRAY
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
    const sourceRows = inputRowsFromSlot(base, kindStack, rangeIndexStack, rangeRowCounts);
    const sourceCols = inputColsFromSlot(base, kindStack, rangeIndexStack, rangeColCounts);
    if (
      sourceRows <= 0 ||
      sourceCols <= 0 ||
      sourceRows == i32.MIN_VALUE ||
      sourceCols == i32.MIN_VALUE
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
    const rowOffset = truncToInt(tagStack[base + 1], valueStack[base + 1]);
    const colOffset = truncToInt(tagStack[base + 2], valueStack[base + 2]);
    const height =
      argc >= 4
        ? coercePositiveIntegerArg(tagStack[base + 3], valueStack[base + 3], true, sourceRows)
        : sourceRows;
    const width =
      argc >= 5
        ? coercePositiveIntegerArg(tagStack[base + 4], valueStack[base + 4], true, sourceCols)
        : sourceCols;
    const areaNumber = argc >= 6 ? truncToInt(tagStack[base + 5], valueStack[base + 5]) : 1;
    if (
      rowOffset == i32.MIN_VALUE ||
      colOffset == i32.MIN_VALUE ||
      height == i32.MIN_VALUE ||
      width == i32.MIN_VALUE ||
      areaNumber == i32.MIN_VALUE ||
      areaNumber != 1 ||
      height < 1 ||
      width < 1
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

    const rowStart = rowOffset < 0 ? sourceRows + rowOffset : rowOffset;
    const colStart = colOffset < 0 ? sourceCols + colOffset : colOffset;
    if (rowStart < 0 || colStart < 0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Ref,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    if (rowStart + height > sourceRows || colStart + width > sourceCols) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Ref,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    if (height == 1 && width == 1) {
      const result = inputCellNumeric(
        base,
        rowStart,
        colStart,
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
      );
      if (isNaN(result)) {
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
        result,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    const arrayIndex = allocateSpillArrayResult(height, width);
    for (let row = 0; row < height; row++) {
      for (let col = 0; col < width; col++) {
        const sourceValue = inputCellNumeric(
          base,
          rowStart + row,
          colStart + col,
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
        );
        if (isNaN(sourceValue)) {
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
        writeSpillArrayNumber(arrayIndex, row * width + col, sourceValue);
      }
    }
    return writeArrayResult(
      base,
      arrayIndex,
      height,
      width,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Take && argc >= 1 && argc <= 3) {
    const sourceKind = kindStack[base];
    if (
      sourceKind != STACK_KIND_SCALAR &&
      sourceKind != STACK_KIND_RANGE &&
      sourceKind != STACK_KIND_ARRAY
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
    const sourceRows = inputRowsFromSlot(base, kindStack, rangeIndexStack, rangeRowCounts);
    const sourceCols = inputColsFromSlot(base, kindStack, rangeIndexStack, rangeColCounts);
    if (
      sourceRows <= 0 ||
      sourceCols <= 0 ||
      sourceRows == i32.MIN_VALUE ||
      sourceCols == i32.MIN_VALUE
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
    const requestedRows =
      argc >= 2 ? coerceInteger(tagStack[base + 1], valueStack[base + 1]) : sourceRows;
    const requestedCols =
      argc >= 3 ? coerceInteger(tagStack[base + 2], valueStack[base + 2]) : sourceCols;
    if (
      (argc >= 2 && requestedRows == i32.MIN_VALUE) ||
      (argc >= 3 && requestedCols == i32.MIN_VALUE)
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
    const clippedRows = argc >= 2 ? clipIndex(requestedRows, sourceRows) : sourceRows;
    const clippedCols = argc >= 3 ? clipIndex(requestedCols, sourceCols) : sourceCols;
    if (clippedRows == i32.MIN_VALUE || clippedCols == i32.MIN_VALUE) {
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
    const rowCount =
      clippedRows > 0 ? min<i32>(clippedRows, sourceRows) : min<i32>(-clippedRows, sourceRows);
    const colCount =
      clippedCols > 0 ? min<i32>(clippedCols, sourceCols) : min<i32>(-clippedCols, sourceCols);
    if (rowCount == 0 || colCount == 0) {
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
    const rowOffset = clippedRows > 0 ? 0 : max<i32>(sourceRows - rowCount, 0);
    const colOffset = clippedCols > 0 ? 0 : max<i32>(sourceCols - colCount, 0);

    const arrayIndex = allocateSpillArrayResult(rowCount, colCount);
    let outputOffset = 0;
    for (let row = 0; row < rowCount; row++) {
      for (let col = 0; col < colCount; col++) {
        const sourceValue = inputCellNumeric(
          base,
          rowOffset + row,
          colOffset + col,
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
        );
        if (isNaN(sourceValue)) {
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
        writeSpillArrayNumber(arrayIndex, outputOffset, sourceValue);
        outputOffset += 1;
      }
    }
    return writeArrayResult(
      base,
      arrayIndex,
      rowCount,
      colCount,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Drop && argc >= 1 && argc <= 3) {
    const sourceKind = kindStack[base];
    if (
      sourceKind != STACK_KIND_SCALAR &&
      sourceKind != STACK_KIND_RANGE &&
      sourceKind != STACK_KIND_ARRAY
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
    const sourceRows = inputRowsFromSlot(base, kindStack, rangeIndexStack, rangeRowCounts);
    const sourceCols = inputColsFromSlot(base, kindStack, rangeIndexStack, rangeColCounts);
    if (
      sourceRows <= 0 ||
      sourceCols <= 0 ||
      sourceRows == i32.MIN_VALUE ||
      sourceCols == i32.MIN_VALUE
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
    const requestedRows = argc >= 2 ? coerceInteger(tagStack[base + 1], valueStack[base + 1]) : 0;
    const requestedCols = argc >= 3 ? coerceInteger(tagStack[base + 2], valueStack[base + 2]) : 0;
    if (
      (argc >= 2 && requestedRows == i32.MIN_VALUE) ||
      (argc >= 3 && requestedCols == i32.MIN_VALUE)
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
    const clippedRows =
      argc >= 2 ? (requestedRows == 0 ? 0 : clipIndex(requestedRows, sourceRows)) : 0;
    const clippedCols =
      argc >= 3 ? (requestedCols == 0 ? 0 : clipIndex(requestedCols, sourceCols)) : 0;
    if (clippedRows == i32.MIN_VALUE || clippedCols == i32.MIN_VALUE) {
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
    const rowDrop =
      clippedRows >= 0 ? min<i32>(clippedRows, sourceRows) : min<i32>(-clippedRows, sourceRows);
    const colDrop =
      clippedCols >= 0 ? min<i32>(clippedCols, sourceCols) : min<i32>(-clippedCols, sourceCols);
    const rowCount = sourceRows - rowDrop;
    const colCount = sourceCols - colDrop;
    if (rowCount <= 0 || colCount <= 0) {
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
    const rowOffset = clippedRows > 0 ? rowDrop : 0;
    const colOffset = clippedCols > 0 ? colDrop : 0;

    const arrayIndex = allocateSpillArrayResult(rowCount, colCount);
    let outputOffset = 0;
    for (let row = 0; row < rowCount; row++) {
      for (let col = 0; col < colCount; col++) {
        const sourceValue = inputCellNumeric(
          base,
          rowOffset + row,
          colOffset + col,
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
        );
        if (isNaN(sourceValue)) {
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
        writeSpillArrayNumber(arrayIndex, outputOffset, sourceValue);
        outputOffset += 1;
      }
    }
    return writeArrayResult(
      base,
      arrayIndex,
      rowCount,
      colCount,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Expand && argc >= 2 && argc <= 4) {
    const sourceKind = kindStack[base];
    if (
      sourceKind != STACK_KIND_SCALAR &&
      sourceKind != STACK_KIND_RANGE &&
      sourceKind != STACK_KIND_ARRAY
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
    const sourceRows = inputRowsFromSlot(base, kindStack, rangeIndexStack, rangeRowCounts);
    const sourceCols = inputColsFromSlot(base, kindStack, rangeIndexStack, rangeColCounts);
    if (
      sourceRows <= 0 ||
      sourceCols <= 0 ||
      sourceRows == i32.MIN_VALUE ||
      sourceCols == i32.MIN_VALUE
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
    for (let arg = 1; arg < argc; arg++) {
      const slot = base + arg;
      if (kindStack[slot] != STACK_KIND_SCALAR) {
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
      );
    }
    if (argc >= 3 && tagStack[base + 2] == ValueTag.Error) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        valueStack[base + 2],
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const targetRows = coerceInteger(tagStack[base + 1], valueStack[base + 1]);
    const targetCols =
      argc >= 3 ? coerceInteger(tagStack[base + 2], valueStack[base + 2]) : sourceCols;
    if (
      targetRows == i32.MIN_VALUE ||
      targetCols == i32.MIN_VALUE ||
      targetRows < sourceRows ||
      targetCols < sourceCols ||
      targetRows < 1 ||
      targetCols < 1
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
    const padTag = argc >= 4 ? tagStack[base + 3] : <u8>ValueTag.Error;
    const padValue = argc >= 4 ? valueStack[base + 3] : ErrorCode.NA;
    const arrayIndex = allocateSpillArrayResult(targetRows, targetCols);
    let outputOffset = 0;
    for (let row = 0; row < targetRows; row++) {
      for (let col = 0; col < targetCols; col++) {
        if (row < sourceRows && col < sourceCols) {
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
          );
          if (copyError != ErrorCode.None) {
            return writeResult(
              base,
              STACK_KIND_SCALAR,
              <u8>ValueTag.Error,
              copyError,
              rangeIndexStack,
              valueStack,
              tagStack,
              kindStack,
            );
          }
        } else {
          writeSpillArrayValue(arrayIndex, outputOffset, padTag, padValue);
        }
        outputOffset += 1;
      }
    }
    return writeArrayResult(
      base,
      arrayIndex,
      targetRows,
      targetCols,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Trimrange && argc >= 1 && argc <= 3) {
    const sourceKind = kindStack[base];
    if (
      sourceKind != STACK_KIND_SCALAR &&
      sourceKind != STACK_KIND_RANGE &&
      sourceKind != STACK_KIND_ARRAY
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
    const sourceRows = inputRowsFromSlot(base, kindStack, rangeIndexStack, rangeRowCounts);
    const sourceCols = inputColsFromSlot(base, kindStack, rangeIndexStack, rangeColCounts);
    if (
      sourceRows <= 0 ||
      sourceCols <= 0 ||
      sourceRows == i32.MIN_VALUE ||
      sourceCols == i32.MIN_VALUE
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
    for (let arg = 1; arg < argc; arg++) {
      const slot = base + arg;
      if (kindStack[slot] != STACK_KIND_SCALAR) {
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
      if (tagStack[slot] == ValueTag.Error) {
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          valueStack[slot],
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }
    }
    const trimRows = argc >= 2 ? coerceTrimMode(tagStack[base + 1], valueStack[base + 1]) : 3;
    const trimCols = argc >= 3 ? coerceTrimMode(tagStack[base + 2], valueStack[base + 2]) : 3;
    if (trimRows == i32.MIN_VALUE || trimCols == i32.MIN_VALUE) {
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

    let startRow = 0;
    let endRow = sourceRows - 1;
    let startCol = 0;
    let endCol = sourceCols - 1;

    const trimLeadingRows = trimRows == 1 || trimRows == 3;
    const trimTrailingRows = trimRows == 2 || trimRows == 3;
    const trimLeadingCols = trimCols == 1 || trimCols == 3;
    const trimTrailingCols = trimCols == 2 || trimCols == 3;

    if (trimLeadingRows) {
      while (startRow <= endRow) {
        let hasNonEmpty = false;
        for (let col = 0; col < sourceCols; col++) {
          if (
            inputCellTag(
              base,
              startRow,
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
            ) != ValueTag.Empty
          ) {
            hasNonEmpty = true;
            break;
          }
        }
        if (hasNonEmpty) {
          break;
        }
        startRow += 1;
      }
    }

    if (trimTrailingRows) {
      while (endRow >= startRow) {
        let hasNonEmpty = false;
        for (let col = 0; col < sourceCols; col++) {
          if (
            inputCellTag(
              base,
              endRow,
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
            ) != ValueTag.Empty
          ) {
            hasNonEmpty = true;
            break;
          }
        }
        if (hasNonEmpty) {
          break;
        }
        endRow -= 1;
      }
    }

    if (startRow > endRow) {
      const arrayIndex = allocateSpillArrayResult(1, 1);
      writeSpillArrayValue(arrayIndex, 0, <u8>ValueTag.Empty, 0);
      return writeArrayResult(
        base,
        arrayIndex,
        1,
        1,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    if (trimLeadingCols) {
      while (startCol <= endCol) {
        let hasNonEmpty = false;
        for (let row = startRow; row <= endRow; row++) {
          if (
            inputCellTag(
              base,
              row,
              startCol,
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
            ) != ValueTag.Empty
          ) {
            hasNonEmpty = true;
            break;
          }
        }
        if (hasNonEmpty) {
          break;
        }
        startCol += 1;
      }
    }

    if (trimTrailingCols) {
      while (endCol >= startCol) {
        let hasNonEmpty = false;
        for (let row = startRow; row <= endRow; row++) {
          if (
            inputCellTag(
              base,
              row,
              endCol,
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
            ) != ValueTag.Empty
          ) {
            hasNonEmpty = true;
            break;
          }
        }
        if (hasNonEmpty) {
          break;
        }
        endCol -= 1;
      }
    }

    if (startCol > endCol) {
      const arrayIndex = allocateSpillArrayResult(1, 1);
      writeSpillArrayValue(arrayIndex, 0, <u8>ValueTag.Empty, 0);
      return writeArrayResult(
        base,
        arrayIndex,
        1,
        1,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    const outputRows = endRow - startRow + 1;
    const outputCols = endCol - startCol + 1;
    const arrayIndex = allocateSpillArrayResult(outputRows, outputCols);
    let outputOffset = 0;
    for (let row = startRow; row <= endRow; row++) {
      for (let col = startCol; col <= endCol; col++) {
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
        );
        if (copyError != ErrorCode.None) {
          return writeResult(
            base,
            STACK_KIND_SCALAR,
            <u8>ValueTag.Error,
            copyError,
            rangeIndexStack,
            valueStack,
            tagStack,
            kindStack,
          );
        }
        outputOffset += 1;
      }
    }
    return writeArrayResult(
      base,
      arrayIndex,
      outputRows,
      outputCols,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Choosecols && argc >= 2) {
    const sourceKind = kindStack[base];
    if (
      sourceKind != STACK_KIND_SCALAR &&
      sourceKind != STACK_KIND_RANGE &&
      sourceKind != STACK_KIND_ARRAY
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
    const sourceRows = inputRowsFromSlot(base, kindStack, rangeIndexStack, rangeRowCounts);
    const sourceCols = inputColsFromSlot(base, kindStack, rangeIndexStack, rangeColCounts);
    if (
      sourceRows <= 0 ||
      sourceCols <= 0 ||
      sourceRows == i32.MIN_VALUE ||
      sourceCols == i32.MIN_VALUE
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
    const scalarError = scalarErrorAt(base, 1, kindStack, tagStack, valueStack);
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
    const selectedCols = new Array<i32>();
    for (let arg = 1; arg < argc; arg++) {
      const argumentSlot = base + arg;
      if (kindStack[argumentSlot] != STACK_KIND_SCALAR) {
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
      const selectedCol = coerceInteger(tagStack[argumentSlot], valueStack[argumentSlot]) - 1;
      if (selectedCol < 0 || selectedCol >= sourceCols || selectedCol == i32.MIN_VALUE - 1) {
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
      selectedCols.push(selectedCol);
    }
    if (selectedCols.length == 0) {
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
    const outputCols = <i32>selectedCols.length;
    const arrayIndex = allocateSpillArrayResult(sourceRows, outputCols);
    let outputOffset = 0;
    for (let row = 0; row < sourceRows; row++) {
      for (let selectedColIndex = 0; selectedColIndex < outputCols; selectedColIndex++) {
        const selectedCol = selectedCols[selectedColIndex];
        const sourceValue = inputCellNumeric(
          base,
          row,
          selectedCol,
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
        );
        if (isNaN(sourceValue)) {
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
        writeSpillArrayNumber(arrayIndex, outputOffset, sourceValue);
        outputOffset += 1;
      }
    }
    return writeArrayResult(
      base,
      arrayIndex,
      sourceRows,
      outputCols,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Chooserows && argc >= 2) {
    const sourceKind = kindStack[base];
    if (
      sourceKind != STACK_KIND_SCALAR &&
      sourceKind != STACK_KIND_RANGE &&
      sourceKind != STACK_KIND_ARRAY
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
    const sourceRows = inputRowsFromSlot(base, kindStack, rangeIndexStack, rangeRowCounts);
    const sourceCols = inputColsFromSlot(base, kindStack, rangeIndexStack, rangeColCounts);
    if (
      sourceRows <= 0 ||
      sourceCols <= 0 ||
      sourceRows == i32.MIN_VALUE ||
      sourceCols == i32.MIN_VALUE
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
    const scalarError = scalarErrorAt(base, 1, kindStack, tagStack, valueStack);
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
    const selectedRows = new Array<i32>();
    for (let arg = 1; arg < argc; arg++) {
      const argumentSlot = base + arg;
      if (kindStack[argumentSlot] != STACK_KIND_SCALAR) {
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
      const selectedRow = coerceInteger(tagStack[argumentSlot], valueStack[argumentSlot]) - 1;
      if (selectedRow < 0 || selectedRow >= sourceRows || selectedRow == i32.MIN_VALUE - 1) {
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
      selectedRows.push(selectedRow);
    }
    if (selectedRows.length == 0) {
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
    const outputRows = <i32>selectedRows.length;
    const arrayIndex = allocateSpillArrayResult(outputRows, sourceCols);
    let outputOffset = 0;
    for (let selectedRowIndex = 0; selectedRowIndex < outputRows; selectedRowIndex++) {
      const selectedRow = selectedRows[selectedRowIndex];
      for (let col = 0; col < sourceCols; col++) {
        const sourceValue = inputCellNumeric(
          base,
          selectedRow,
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
        );
        if (isNaN(sourceValue)) {
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
        writeSpillArrayNumber(arrayIndex, outputOffset, sourceValue);
        outputOffset += 1;
      }
    }
    return writeArrayResult(
      base,
      arrayIndex,
      outputRows,
      sourceCols,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Sort && argc >= 1 && argc <= 4) {
    const sourceKind = kindStack[base];
    if (
      sourceKind != STACK_KIND_SCALAR &&
      sourceKind != STACK_KIND_RANGE &&
      sourceKind != STACK_KIND_ARRAY
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
    const sourceRows = inputRowsFromSlot(base, kindStack, rangeIndexStack, rangeRowCounts);
    const sourceCols = inputColsFromSlot(base, kindStack, rangeIndexStack, rangeColCounts);
    if (
      sourceRows <= 0 ||
      sourceCols <= 0 ||
      sourceRows == i32.MIN_VALUE ||
      sourceCols == i32.MIN_VALUE
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

    const sortIndex = argc >= 2 ? coerceInteger(tagStack[base + 1], valueStack[base + 1]) : 1;
    const sortOrder = argc >= 3 ? coerceInteger(tagStack[base + 2], valueStack[base + 2]) : 1;
    const sortByColBoolean =
      argc >= 4 ? coerceBoolean(tagStack[base + 3], valueStack[base + 3]) : 0;
    if (
      sortIndex == i32.MIN_VALUE ||
      sortIndex < 1 ||
      (sortOrder != 1 && sortOrder != -1) ||
      sortByColBoolean < 0
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
    const sortByCol = sortByColBoolean != 0;

    if (sourceRows == 1 || sourceCols == 1) {
      const length = sourceRows * sourceCols;
      const order = new Array<i32>(length);
      for (let index = 0; index < length; index++) {
        order.push(index);
      }
      for (let index = 1; index < length; index++) {
        const current = order[index];
        const currentRow = current / sourceCols;
        const currentCol = current - currentRow * sourceCols;
        const currentValue = inputCellNumeric(
          base,
          currentRow,
          currentCol,
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
        );
        if (isNaN(currentValue)) {
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
        let cursor = index;
        while (cursor > 0) {
          const previous = order[cursor - 1];
          const previousRow = previous / sourceCols;
          const previousCol = previous - previousRow * sourceCols;
          const previousValue = inputCellNumeric(
            base,
            previousRow,
            previousCol,
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
          );
          if (isNaN(previousValue)) {
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
          const comparison =
            currentValue == previousValue ? 0 : currentValue < previousValue ? -1 : 1;
          if (comparison * sortOrder < 0) {
            order[cursor] = previous;
            cursor -= 1;
            continue;
          }
          break;
        }
        order[cursor] = current;
      }
      const arrayIndex = allocateSpillArrayResult(sourceRows, sourceCols);
      for (let index = 0; index < length; index++) {
        const sourceOffset = order[index];
        const sourceRow = sourceOffset / sourceCols;
        const sourceCol = sourceOffset - sourceRow * sourceCols;
        const sourceValue = inputCellNumeric(
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
        );
        if (isNaN(sourceValue)) {
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
        writeSpillArrayNumber(arrayIndex, index, sourceValue);
      }
      return writeArrayResult(
        base,
        arrayIndex,
        sourceRows,
        sourceCols,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    if (sortByCol) {
      if (sortIndex > sourceRows) {
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
      const rowSort = new Array<i32>(sourceRows);
      for (let row = 0; row < sourceRows; row++) {
        rowSort.push(row);
      }
      const sortCol = sortIndex - 1;
      for (let cursor = 1; cursor < sourceRows; cursor++) {
        const currentRow = rowSort[cursor];
        const currentValue = inputCellNumeric(
          base,
          currentRow,
          sortCol,
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
        );
        if (isNaN(currentValue)) {
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
        let position = cursor;
        while (position > 0) {
          const previousRow = rowSort[position - 1];
          const previousValue = inputCellNumeric(
            base,
            previousRow,
            sortCol,
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
          );
          if (isNaN(previousValue)) {
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
          const comparison =
            currentValue == previousValue ? 0 : currentValue < previousValue ? -1 : 1;
          if (comparison * sortOrder < 0) {
            rowSort[position] = previousRow;
            position -= 1;
            continue;
          }
          break;
        }
        rowSort[position] = currentRow;
      }
      const arrayIndex = allocateSpillArrayResult(sourceRows, sourceCols);
      let outputOffset = 0;
      for (let sortedRow = 0; sortedRow < sourceRows; sortedRow++) {
        const sourceRow = rowSort[sortedRow];
        for (let col = 0; col < sourceCols; col++) {
          const value = inputCellNumeric(
            base,
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
          );
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
          writeSpillArrayNumber(arrayIndex, outputOffset, value);
          outputOffset += 1;
        }
      }
      return writeArrayResult(
        base,
        arrayIndex,
        sourceRows,
        sourceCols,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    if (sortIndex > sourceCols) {
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
    const colSort = new Array<i32>(sourceCols);
    for (let col = 0; col < sourceCols; col++) {
      colSort.push(col);
    }
    const sortRow = sortIndex - 1;
    for (let cursor = 1; cursor < sourceCols; cursor++) {
      const currentCol = colSort[cursor];
      const currentValue = inputCellNumeric(
        base,
        currentCol,
        sortRow,
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
      );
      if (isNaN(currentValue)) {
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
      let position = cursor;
      while (position > 0) {
        const previousCol = colSort[position - 1];
        const previousValue = inputCellNumeric(
          base,
          previousCol,
          sortRow,
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
        );
        if (isNaN(previousValue)) {
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
        const comparison =
          currentValue == previousValue ? 0 : currentValue < previousValue ? -1 : 1;
        if (comparison * sortOrder < 0) {
          colSort[position] = previousCol;
          position -= 1;
          continue;
        }
        break;
      }
      colSort[position] = currentCol;
    }
    const arrayIndex = allocateSpillArrayResult(sourceRows, sourceCols);
    let outputOffset = 0;
    for (let row = 0; row < sourceRows; row++) {
      for (let sortedCol = 0; sortedCol < sourceCols; sortedCol++) {
        const sourceCol = colSort[sortedCol];
        const value = inputCellNumeric(
          base,
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
        );
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
        writeSpillArrayNumber(arrayIndex, outputOffset, value);
        outputOffset += 1;
      }
    }
    return writeArrayResult(
      base,
      arrayIndex,
      sourceRows,
      sourceCols,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Sortby && argc >= 2) {
    const sourceKind = kindStack[base];
    if (
      sourceKind != STACK_KIND_SCALAR &&
      sourceKind != STACK_KIND_RANGE &&
      sourceKind != STACK_KIND_ARRAY
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
    const sourceRows = inputRowsFromSlot(base, kindStack, rangeIndexStack, rangeRowCounts);
    const sourceCols = inputColsFromSlot(base, kindStack, rangeIndexStack, rangeColCounts);
    if (
      sourceRows <= 0 ||
      sourceCols <= 0 ||
      sourceRows == i32.MIN_VALUE ||
      sourceCols == i32.MIN_VALUE
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
    const sourceLength = sourceRows * sourceCols;
    if (sourceRows > 1 && sourceCols > 1) {
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

    const sortBySlots = new Array<i32>();
    const sortByLengths = new Array<i32>();
    const sortByCols = new Array<i32>();
    const sortByOrders = new Array<i32>();
    let arg = 1;
    while (arg < argc) {
      const criterionSlot = base + arg;
      const criterionKind = kindStack[criterionSlot];
      if (
        criterionKind != STACK_KIND_SCALAR &&
        criterionKind != STACK_KIND_RANGE &&
        criterionKind != STACK_KIND_ARRAY
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

      const criterionRows = inputRowsFromSlot(
        criterionSlot,
        kindStack,
        rangeIndexStack,
        rangeRowCounts,
      );
      const criterionCols = inputColsFromSlot(
        criterionSlot,
        kindStack,
        rangeIndexStack,
        rangeColCounts,
      );
      if (
        criterionRows <= 0 ||
        criterionCols <= 0 ||
        criterionRows == i32.MIN_VALUE ||
        criterionCols == i32.MIN_VALUE
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
      const criterionLength = criterionRows * criterionCols;
      if (criterionLength != 1 && criterionLength != sourceLength) {
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

      sortBySlots.push(criterionSlot);
      sortByLengths.push(criterionLength);
      sortByCols.push(criterionCols);

      const nextSlot = criterionSlot + 1;
      if (nextSlot < base + argc) {
        if (kindStack[nextSlot] == STACK_KIND_RANGE || kindStack[nextSlot] == STACK_KIND_ARRAY) {
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
        if (kindStack[nextSlot] == STACK_KIND_SCALAR) {
          const requestedOrder = coerceInteger(tagStack[nextSlot], valueStack[nextSlot]);
          if (requestedOrder != 1 && requestedOrder != -1) {
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
          sortByOrders.push(requestedOrder);
          arg += 2;
          continue;
        }
      }

      sortByOrders.push(1);
      arg += 1;
    }

    if (sortBySlots.length == 0) {
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

    const sourceIndexes = new Array<i32>(sourceLength);
    for (let offset = 0; offset < sourceLength; offset++) {
      sourceIndexes.push(offset);
    }

    for (let cursor = 1; cursor < sourceLength; cursor++) {
      const currentOffset = sourceIndexes[cursor];
      const currentSourceRow = currentOffset / sourceCols;
      const currentSourceCol = currentOffset - currentSourceRow * sourceCols;
      const currentSourceValue = inputCellNumeric(
        base,
        currentSourceRow,
        currentSourceCol,
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
      );
      if (isNaN(currentSourceValue)) {
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

      let position = cursor;
      while (position > 0) {
        const previousOffset = sourceIndexes[position - 1];
        const previousSourceRow = previousOffset / sourceCols;
        const previousSourceCol = previousOffset - previousSourceRow * sourceCols;
        const previousSourceValue = inputCellNumeric(
          base,
          previousSourceRow,
          previousSourceCol,
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
        );
        if (isNaN(previousSourceValue)) {
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

        let comparison = 0;
        for (let sortByIndex = 0; sortByIndex < sortBySlots.length; sortByIndex++) {
          const slot = sortBySlots[sortByIndex];
          const slotLength = sortByLengths[sortByIndex];
          const slotCols = sortByCols[sortByIndex];
          const slotOrder = sortByOrders[sortByIndex];

          const currentCriterionOffset = slotLength == 1 ? 0 : currentOffset;
          const previousCriterionOffset = slotLength == 1 ? 0 : previousOffset;
          const currentCriterionRow = slotLength == 1 ? 0 : currentCriterionOffset / slotCols;
          const currentCriterionCol =
            slotLength == 1 ? 0 : currentCriterionOffset - currentCriterionRow * slotCols;
          const previousCriterionRow = slotLength == 1 ? 0 : previousCriterionOffset / slotCols;
          const previousCriterionCol =
            slotLength == 1 ? 0 : previousCriterionOffset - previousCriterionRow * slotCols;

          const currentCriterionTag = inputCellTag(
            slot,
            currentCriterionRow,
            currentCriterionCol,
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
          );
          const currentCriterionValue = inputCellNumeric(
            slot,
            currentCriterionRow,
            currentCriterionCol,
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
          );
          const previousCriterionTag = inputCellTag(
            slot,
            previousCriterionRow,
            previousCriterionCol,
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
          );
          const previousCriterionValue = inputCellNumeric(
            slot,
            previousCriterionRow,
            previousCriterionCol,
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
          );
          const criterionComparison = compareScalarValues(
            currentCriterionTag,
            currentCriterionValue,
            previousCriterionTag,
            previousCriterionValue,
            null,
            stringOffsets,
            stringLengths,
            stringData,
            outputStringOffsets,
            outputStringLengths,
            outputStringData,
          );
          if (criterionComparison == i32.MIN_VALUE) {
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
          if (criterionComparison != 0) {
            comparison = criterionComparison * slotOrder;
            break;
          }
        }
        if (comparison < 0) {
          sourceIndexes[position] = previousOffset;
          position -= 1;
          continue;
        }
        if (comparison > 0) {
          break;
        }
        break;
      }
      sourceIndexes[position] = currentOffset;
    }

    const arrayIndex = allocateSpillArrayResult(sourceRows, sourceCols);
    let outputOffset = 0;
    for (let index = 0; index < sourceLength; index++) {
      const sortedOffset = sourceIndexes[index];
      const sourceRow = sortedOffset / sourceCols;
      const sourceCol = sortedOffset - sourceRow * sourceCols;
      const sourceValue = inputCellNumeric(
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
      );
      if (isNaN(sourceValue)) {
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
      writeSpillArrayNumber(arrayIndex, outputOffset, sourceValue);
      outputOffset += 1;
    }
    return writeArrayResult(
      base,
      arrayIndex,
      sourceRows,
      sourceCols,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Tocol && argc >= 1 && argc <= 3) {
    const sourceKind = kindStack[base];
    if (
      sourceKind != STACK_KIND_SCALAR &&
      sourceKind != STACK_KIND_RANGE &&
      sourceKind != STACK_KIND_ARRAY
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
    const sourceRows = inputRowsFromSlot(base, kindStack, rangeIndexStack, rangeRowCounts);
    const sourceCols = inputColsFromSlot(base, kindStack, rangeIndexStack, rangeColCounts);
    if (
      sourceRows <= 0 ||
      sourceCols <= 0 ||
      sourceRows == i32.MIN_VALUE ||
      sourceCols == i32.MIN_VALUE
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
    const ignoreValue = argc >= 2 ? coerceInteger(tagStack[base + 1], valueStack[base + 1]) : 0;
    const scanByCol = argc >= 3 ? coerceBoolean(tagStack[base + 2], valueStack[base + 2]) : 1;
    if (ignoreValue < 0 || (ignoreValue != 0 && ignoreValue != 1) || scanByCol < 0) {
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
    const ignoreEmpty = ignoreValue == 1;

    const values = new Array<f64>();
    if (scanByCol == 0) {
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
          );
          if (ignoreEmpty && sourceTag == ValueTag.Empty) {
            continue;
          }
          const sourceValue = inputCellNumeric(
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
          );
          if (isNaN(sourceValue)) {
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
          values.push(sourceValue);
        }
      }
    } else {
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
          );
          if (ignoreEmpty && sourceTag == ValueTag.Empty) {
            continue;
          }
          const sourceValue = inputCellNumeric(
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
          );
          if (isNaN(sourceValue)) {
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
          values.push(sourceValue);
        }
      }
    }

    const outputRows = values.length;
    const arrayIndex = allocateSpillArrayResult(outputRows, 1);
    for (let offset = 0; offset < outputRows; offset++) {
      writeSpillArrayNumber(arrayIndex, offset, values[offset]);
    }
    return writeArrayResult(
      base,
      arrayIndex,
      outputRows,
      1,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Torow && argc >= 1 && argc <= 3) {
    const sourceKind = kindStack[base];
    if (
      sourceKind != STACK_KIND_SCALAR &&
      sourceKind != STACK_KIND_RANGE &&
      sourceKind != STACK_KIND_ARRAY
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
    const sourceRows = inputRowsFromSlot(base, kindStack, rangeIndexStack, rangeRowCounts);
    const sourceCols = inputColsFromSlot(base, kindStack, rangeIndexStack, rangeColCounts);
    if (
      sourceRows <= 0 ||
      sourceCols <= 0 ||
      sourceRows == i32.MIN_VALUE ||
      sourceCols == i32.MIN_VALUE
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
    const ignoreValue = argc >= 2 ? coerceInteger(tagStack[base + 1], valueStack[base + 1]) : 0;
    const scanByCol = argc >= 3 ? coerceBoolean(tagStack[base + 2], valueStack[base + 2]) : 0;
    if (ignoreValue < 0 || (ignoreValue != 0 && ignoreValue != 1) || scanByCol < 0) {
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
    const ignoreEmpty = ignoreValue == 1;

    const values = new Array<f64>();
    if (scanByCol == 0) {
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
          );
          if (ignoreEmpty && sourceTag == ValueTag.Empty) {
            continue;
          }
          const sourceValue = inputCellNumeric(
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
          );
          if (isNaN(sourceValue)) {
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
          values.push(sourceValue);
        }
      }
    } else {
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
          );
          if (ignoreEmpty && sourceTag == ValueTag.Empty) {
            continue;
          }
          const sourceValue = inputCellNumeric(
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
          );
          if (isNaN(sourceValue)) {
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
          values.push(sourceValue);
        }
      }
    }

    const outputCols = values.length;
    const arrayIndex = allocateSpillArrayResult(1, outputCols);
    for (let offset = 0; offset < outputCols; offset++) {
      writeSpillArrayNumber(arrayIndex, offset, values[offset]);
    }
    return writeArrayResult(
      base,
      arrayIndex,
      1,
      outputCols,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Wraprows && argc >= 2 && argc <= 4) {
    const sourceKind = kindStack[base];
    if (
      sourceKind != STACK_KIND_SCALAR &&
      sourceKind != STACK_KIND_RANGE &&
      sourceKind != STACK_KIND_ARRAY
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
    const sourceRows = inputRowsFromSlot(base, kindStack, rangeIndexStack, rangeRowCounts);
    const sourceCols = inputColsFromSlot(base, kindStack, rangeIndexStack, rangeColCounts);
    if (
      sourceRows <= 0 ||
      sourceCols <= 0 ||
      sourceRows == i32.MIN_VALUE ||
      sourceCols == i32.MIN_VALUE
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

    const wrapCount = coerceInteger(tagStack[base + 1], valueStack[base + 1]);
    if (wrapCount == i32.MIN_VALUE || wrapCount < 1) {
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

    const needsPad = (sourceRows * sourceCols) % wrapCount != 0;
    if (needsPad && argc < 3) {
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

    if (argc >= 4) {
      const padByColBoolean = coerceBoolean(tagStack[base + 3], valueStack[base + 3]);
      if (padByColBoolean < 0) {
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
    }

    const defaultPadValue = argc >= 3 ? toNumberOrNaN(tagStack[base + 2], valueStack[base + 2]) : 0;
    if (argc >= 3 && isNaN(defaultPadValue)) {
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

    const sourceLength = sourceRows * sourceCols;
    const outputCols = wrapCount;
    const outputRows = (sourceLength + wrapCount - 1) / wrapCount;
    const outputLength = outputRows * outputCols;

    const arrayIndex = allocateSpillArrayResult(outputRows, outputCols);
    for (let outputOffset = 0; outputOffset < sourceLength; outputOffset++) {
      const sourceRow = outputOffset / sourceCols;
      const sourceCol = outputOffset - sourceRow * sourceCols;
      const sourceValue = inputCellNumeric(
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
      );
      if (isNaN(sourceValue)) {
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
      writeSpillArrayNumber(arrayIndex, outputOffset, sourceValue);
    }
    const padValue = needsPad && argc >= 3 ? defaultPadValue : 0;
    for (let outputOffset = sourceLength; outputOffset < outputLength; outputOffset++) {
      writeSpillArrayNumber(arrayIndex, outputOffset, padValue);
    }
    return writeArrayResult(
      base,
      arrayIndex,
      outputRows,
      outputCols,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Wrapcols && argc >= 2 && argc <= 4) {
    const sourceKind = kindStack[base];
    if (
      sourceKind != STACK_KIND_SCALAR &&
      sourceKind != STACK_KIND_RANGE &&
      sourceKind != STACK_KIND_ARRAY
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
    const sourceRows = inputRowsFromSlot(base, kindStack, rangeIndexStack, rangeRowCounts);
    const sourceCols = inputColsFromSlot(base, kindStack, rangeIndexStack, rangeColCounts);
    if (
      sourceRows <= 0 ||
      sourceCols <= 0 ||
      sourceRows == i32.MIN_VALUE ||
      sourceCols == i32.MIN_VALUE
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

    const wrapCount = coerceInteger(tagStack[base + 1], valueStack[base + 1]);
    if (wrapCount == i32.MIN_VALUE || wrapCount < 1) {
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

    if (argc >= 4) {
      const padByColBoolean = coerceBoolean(tagStack[base + 3], valueStack[base + 3]);
      if (padByColBoolean < 0) {
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
    }

    const needsPad = (sourceRows * sourceCols) % wrapCount != 0;
    if (needsPad && argc < 3) {
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

    const defaultPadValue = argc >= 3 ? toNumberOrNaN(tagStack[base + 2], valueStack[base + 2]) : 0;
    if (argc >= 3 && isNaN(defaultPadValue)) {
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

    const sourceLength = sourceRows * sourceCols;
    const outputRows = wrapCount;
    const outputCols = (sourceLength + wrapCount - 1) / wrapCount;
    const outputLength = outputRows * outputCols;

    const arrayIndex = allocateSpillArrayResult(outputRows, outputCols);
    for (let outputOffset = 0; outputOffset < outputLength; outputOffset++) {
      const sourceOffset = (outputOffset / outputRows) * outputRows + (outputOffset % outputRows);
      let sourceValue = defaultPadValue;
      if (sourceOffset < sourceLength) {
        const sourceRow = sourceOffset / sourceCols;
        const sourceCol = sourceOffset - sourceRow * sourceCols;
        sourceValue = inputCellNumeric(
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
        );
      }
      if (isNaN(sourceValue)) {
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
      writeSpillArrayNumber(arrayIndex, outputOffset, sourceValue);
    }
    return writeArrayResult(
      base,
      arrayIndex,
      outputRows,
      outputCols,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  const lookupTableResult = tryApplyLookupTableBuiltin(
    builtinId,
    argc,
    base,
    rangeIndexStack,
    valueStack,
    tagStack,
    kindStack,
    cellTags,
    cellNumbers,
    cellStringIds,
    cellErrors,
    stringOffsets,
    stringLengths,
    stringData,
    rangeOffsets,
    rangeLengths,
    rangeRowCounts,
    rangeColCounts,
    rangeMembers,
    outputStringOffsets,
    outputStringLengths,
    outputStringData,
  );
  if (lookupTableResult >= 0) {
    return lookupTableResult;
  }

  const aggregateCriteriaResult = tryApplyAggregateCriteriaBuiltin(
    builtinId,
    argc,
    base,
    rangeIndexStack,
    valueStack,
    tagStack,
    kindStack,
    cellTags,
    cellNumbers,
    cellStringIds,
    cellErrors,
    stringOffsets,
    stringLengths,
    stringData,
    rangeOffsets,
    rangeLengths,
    rangeRowCounts,
    rangeColCounts,
    rangeMembers,
    outputStringOffsets,
    outputStringLengths,
    outputStringData,
  );
  if (aggregateCriteriaResult >= 0) {
    return aggregateCriteriaResult;
  }

  if (
    (builtinId == BuiltinId.Daverage ||
      builtinId == BuiltinId.Dcount ||
      builtinId == BuiltinId.Dcounta ||
      builtinId == BuiltinId.Dget ||
      builtinId == BuiltinId.Dmax ||
      builtinId == BuiltinId.Dmin ||
      builtinId == BuiltinId.Dproduct ||
      builtinId == BuiltinId.Dstdev ||
      builtinId == BuiltinId.Dstdevp ||
      builtinId == BuiltinId.Dsum ||
      builtinId == BuiltinId.Dvar ||
      builtinId == BuiltinId.Dvarp) &&
    argc == 3
  ) {
    return databaseBuiltinResult(
      <u16>builtinId,
      base,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
      cellTags,
      cellNumbers,
      cellStringIds,
      cellErrors,
      stringOffsets,
      stringLengths,
      stringData,
      rangeOffsets,
      rangeLengths,
      rangeRowCounts,
      rangeColCounts,
      rangeMembers,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    );
  }

  if (builtinId == BuiltinId.Sumproduct) {
    if (argc == 0) {
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

    const firstRangeIndex = rangeIndexStack[base];
    if (kindStack[base] != STACK_KIND_RANGE) {
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
    const expectedLength = <i32>rangeLengths[firstRangeIndex];
    for (let index = 0; index < argc; index++) {
      const slot = base + index;
      if (
        kindStack[slot] != STACK_KIND_RANGE ||
        <i32>rangeLengths[rangeIndexStack[slot]] != expectedLength
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
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      sum,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  const regressionResult = tryApplyRegressionBuiltin(
    builtinId,
    argc,
    base,
    rangeIndexStack,
    valueStack,
    tagStack,
    kindStack,
    cellTags,
    cellNumbers,
    cellStringIds,
    cellErrors,
    rangeOffsets,
    rangeLengths,
    rangeRowCounts,
    rangeColCounts,
    rangeMembers,
  );
  if (regressionResult >= 0) {
    return regressionResult;
  }

  const statisticsSummaryResult = tryApplyStatisticsSummaryBuiltin(
    builtinId,
    argc,
    base,
    rangeIndexStack,
    valueStack,
    tagStack,
    kindStack,
    cellTags,
    cellNumbers,
    cellStringIds,
    cellErrors,
    rangeOffsets,
    rangeLengths,
    rangeRowCounts,
    rangeColCounts,
    rangeMembers,
  );
  if (statisticsSummaryResult >= 0) {
    return statisticsSummaryResult;
  }

  const lookupMatchResult = tryApplyLookupMatchBuiltin(
    builtinId,
    argc,
    base,
    rangeIndexStack,
    valueStack,
    tagStack,
    kindStack,
    cellTags,
    cellNumbers,
    cellStringIds,
    cellErrors,
    stringOffsets,
    stringLengths,
    stringData,
    rangeOffsets,
    rangeLengths,
    rangeRowCounts,
    rangeColCounts,
    rangeMembers,
    outputStringOffsets,
    outputStringLengths,
    outputStringData,
  );
  if (lookupMatchResult >= 0) {
    return lookupMatchResult;
  }

  const arrayInfoResult = tryApplyArrayInfoBuiltin(
    builtinId,
    argc,
    base,
    rangeIndexStack,
    valueStack,
    tagStack,
    kindStack,
    cellTags,
    cellNumbers,
    cellStringIds,
    cellErrors,
    stringOffsets,
    stringLengths,
    stringData,
    rangeOffsets,
    rangeLengths,
    rangeRowCounts,
    rangeColCounts,
    rangeMembers,
    outputStringOffsets,
    outputStringLengths,
    outputStringData,
  );
  if (arrayInfoResult >= 0) {
    return arrayInfoResult;
  }

  const scalarDistributionResult = tryApplyScalarDistributionBuiltin(
    builtinId,
    argc,
    base,
    rangeIndexStack,
    valueStack,
    tagStack,
    kindStack,
    stringOffsets,
    stringLengths,
    stringData,
    outputStringOffsets,
    outputStringLengths,
    outputStringData,
  );
  if (scalarDistributionResult >= 0) {
    return scalarDistributionResult;
  }

  if (!rangeSupportedScalarOnly(base, argc, kindStack)) {
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

  if (
    (builtinId == BuiltinId.Expondist ||
      builtinId == BuiltinId.ExponDist ||
      builtinId == BuiltinId.Poisson ||
      builtinId == BuiltinId.PoissonDist ||
      builtinId == BuiltinId.Negbinomdist) &&
    argc == 3
  ) {
    if (!rangeSupportedScalarOnly(base, argc, kindStack)) {
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
    let result = NaN;
    if (builtinId == BuiltinId.Expondist || builtinId == BuiltinId.ExponDist) {
      const x = toNumberExact(tagStack[base], valueStack[base]);
      const lambda = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
      const cumulative = coerceBoolean(tagStack[base + 2], valueStack[base + 2]);
      result =
        isNaN(x) || isNaN(lambda) || cumulative < 0 || x < 0.0 || lambda <= 0.0
          ? NaN
          : cumulative == 1
            ? 1.0 - Math.exp(-lambda * x)
            : lambda * Math.exp(-lambda * x);
    } else if (builtinId == BuiltinId.Poisson || builtinId == BuiltinId.PoissonDist) {
      const eventsRaw = toNumberExact(tagStack[base], valueStack[base]);
      const mean = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
      const cumulative = coerceBoolean(tagStack[base + 2], valueStack[base + 2]);
      const events = <i32>eventsRaw;
      if (!isNaN(eventsRaw) && mean >= 0.0 && cumulative >= 0 && events >= 0) {
        if (cumulative == 1) {
          result = 0.0;
          for (let index = 0; index <= events; index += 1) {
            result += poissonProbability(index, mean);
          }
        } else {
          result = poissonProbability(events, mean);
        }
      }
    } else {
      const failuresRaw = toNumberExact(tagStack[base], valueStack[base]);
      const successesRaw = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
      const probability = toNumberExact(tagStack[base + 2], valueStack[base + 2]);
      const failures = <i32>failuresRaw;
      const successes = <i32>successesRaw;
      if (!isNaN(failuresRaw) && !isNaN(successesRaw)) {
        result = negativeBinomialProbability(failures, successes, probability);
      }
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      isNumericResult(result) ? <u8>ValueTag.Number : <u8>ValueTag.Error,
      isNumericResult(result) ? result : ErrorCode.Value,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (
    (builtinId == BuiltinId.Weibull ||
      builtinId == BuiltinId.WeibullDist ||
      builtinId == BuiltinId.Gammadist ||
      builtinId == BuiltinId.GammaDist ||
      builtinId == BuiltinId.Binomdist ||
      builtinId == BuiltinId.BinomDist ||
      builtinId == BuiltinId.NegbinomDist) &&
    argc == 4
  ) {
    if (!rangeSupportedScalarOnly(base, argc, kindStack)) {
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
    let result = NaN;
    if (builtinId == BuiltinId.Weibull || builtinId == BuiltinId.WeibullDist) {
      const x = toNumberExact(tagStack[base], valueStack[base]);
      const alpha = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
      const beta = toNumberExact(tagStack[base + 2], valueStack[base + 2]);
      const cumulative = coerceBoolean(tagStack[base + 3], valueStack[base + 3]);
      if (
        !isNaN(x) &&
        !isNaN(alpha) &&
        !isNaN(beta) &&
        cumulative >= 0 &&
        x >= 0.0 &&
        alpha > 0.0 &&
        beta > 0.0
      ) {
        if (cumulative == 1) {
          result = 1.0 - Math.exp(-Math.pow(x / beta, alpha));
        } else if (x == 0.0) {
          result = alpha == 1.0 ? 1.0 / beta : alpha < 1.0 ? Infinity : 0.0;
        } else {
          result =
            (alpha / Math.pow(beta, alpha)) *
            Math.pow(x, alpha - 1.0) *
            Math.exp(-Math.pow(x / beta, alpha));
        }
      }
    } else if (builtinId == BuiltinId.Gammadist || builtinId == BuiltinId.GammaDist) {
      const x = toNumberExact(tagStack[base], valueStack[base]);
      const alpha = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
      const beta = toNumberExact(tagStack[base + 2], valueStack[base + 2]);
      const cumulative = coerceBoolean(tagStack[base + 3], valueStack[base + 3]);
      if (
        !isNaN(x) &&
        !isNaN(alpha) &&
        !isNaN(beta) &&
        cumulative >= 0 &&
        x >= 0.0 &&
        alpha > 0.0 &&
        beta > 0.0
      ) {
        result =
          cumulative == 1
            ? gammaDistributionCdf(x, alpha, beta)
            : gammaDistributionDensity(x, alpha, beta);
      }
    } else if (builtinId == BuiltinId.Binomdist || builtinId == BuiltinId.BinomDist) {
      const successesRaw = toNumberExact(tagStack[base], valueStack[base]);
      const trialsRaw = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
      const probability = toNumberExact(tagStack[base + 2], valueStack[base + 2]);
      const cumulative = coerceBoolean(tagStack[base + 3], valueStack[base + 3]);
      const successes = <i32>successesRaw;
      const trials = <i32>trialsRaw;
      if (!isNaN(successesRaw) && !isNaN(trialsRaw) && cumulative >= 0) {
        if (cumulative == 1) {
          result = 0.0;
          for (let index = 0; index <= successes; index += 1) {
            result += binomialProbability(index, trials, probability);
          }
        } else {
          result = binomialProbability(successes, trials, probability);
        }
      }
    } else {
      const failuresRaw = toNumberExact(tagStack[base], valueStack[base]);
      const successesRaw = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
      const probability = toNumberExact(tagStack[base + 2], valueStack[base + 2]);
      const cumulative = coerceBoolean(tagStack[base + 3], valueStack[base + 3]);
      const failures = <i32>failuresRaw;
      const successes = <i32>successesRaw;
      if (!isNaN(failuresRaw) && !isNaN(successesRaw) && cumulative >= 0) {
        if (cumulative == 1) {
          result = 0.0;
          for (let index = 0; index <= failures; index += 1) {
            result += negativeBinomialProbability(index, successes, probability);
          }
        } else {
          result = negativeBinomialProbability(failures, successes, probability);
        }
      }
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      isNumericResult(result) ? <u8>ValueTag.Number : <u8>ValueTag.Error,
      isNumericResult(result) ? result : ErrorCode.Value,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (
    (builtinId == BuiltinId.Chidist ||
      builtinId == BuiltinId.LegacyChidist ||
      builtinId == BuiltinId.ChisqDistRt ||
      builtinId == BuiltinId.Chisqdist) &&
    argc == 2
  ) {
    if (!rangeSupportedScalarOnly(base, argc, kindStack)) {
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
    const x = toNumberExact(tagStack[base], valueStack[base]);
    const degrees = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
    const result =
      !isNaN(x) && !isNaN(degrees) && x >= 0.0 && degrees >= 1.0
        ? regularizedUpperGamma(degrees / 2.0, x / 2.0)
        : NaN;
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      isNumericResult(result) ? <u8>ValueTag.Number : <u8>ValueTag.Error,
      isNumericResult(result) ? result : ErrorCode.Value,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.ChisqDist && argc == 3) {
    if (!rangeSupportedScalarOnly(base, argc, kindStack)) {
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
    const x = toNumberExact(tagStack[base], valueStack[base]);
    const degrees = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
    const cumulative = coerceBoolean(tagStack[base + 2], valueStack[base + 2]);
    const result =
      !isNaN(x) && !isNaN(degrees) && cumulative >= 0 && x >= 0.0 && degrees >= 1.0
        ? cumulative == 1
          ? chiSquareCdf(x, degrees)
          : chiSquareDensity(x, degrees)
        : NaN;
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      isNumericResult(result) ? <u8>ValueTag.Number : <u8>ValueTag.Error,
      isNumericResult(result) ? result : ErrorCode.Value,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (
    (builtinId == BuiltinId.Chiinv ||
      builtinId == BuiltinId.ChisqInvRt ||
      builtinId == BuiltinId.Chisqinv ||
      builtinId == BuiltinId.LegacyChiinv) &&
    argc == 2
  ) {
    if (!rangeSupportedScalarOnly(base, argc, kindStack)) {
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
    const probability = toNumberExact(tagStack[base], valueStack[base]);
    const degrees = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
    const valid =
      !isNaN(probability) &&
      !isNaN(degrees) &&
      probability > 0.0 &&
      probability < 1.0 &&
      degrees >= 1.0;
    let result = NaN;
    if (valid) {
      result = inverseChiSquare(1.0 - probability, degrees);
    }
    const resultTag = isNumericResult(result) ? <u8>ValueTag.Number : <u8>ValueTag.Error;
    const resultValue = isNumericResult(result) ? result : valid ? ErrorCode.NA : ErrorCode.Value;
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      resultTag,
      resultValue,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.ChisqInv && argc == 2) {
    if (!rangeSupportedScalarOnly(base, argc, kindStack)) {
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
    const probability = toNumberExact(tagStack[base], valueStack[base]);
    const degrees = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
    const valid =
      !isNaN(probability) &&
      !isNaN(degrees) &&
      probability > 0.0 &&
      probability < 1.0 &&
      degrees >= 1.0;
    let result = NaN;
    if (valid) {
      result = inverseChiSquare(probability, degrees);
    }
    const resultTag = isNumericResult(result) ? <u8>ValueTag.Number : <u8>ValueTag.Error;
    const resultValue = isNumericResult(result) ? result : valid ? ErrorCode.NA : ErrorCode.Value;
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      resultTag,
      resultValue,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (
    (builtinId == BuiltinId.BetaDist && argc >= 4 && argc <= 6) ||
    (builtinId == BuiltinId.Betadist && argc >= 3 && argc <= 5)
  ) {
    if (!rangeSupportedScalarOnly(base, argc, kindStack)) {
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
    const x = toNumberExact(tagStack[base], valueStack[base]);
    const alpha = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
    const beta = toNumberExact(tagStack[base + 2], valueStack[base + 2]);
    const modern = builtinId == BuiltinId.BetaDist;
    const cumulative = modern ? coerceBoolean(tagStack[base + 3], valueStack[base + 3]) : 1;
    const lowerBound = modern
      ? argc >= 5
        ? toNumberExact(tagStack[base + 4], valueStack[base + 4])
        : 0.0
      : argc >= 4
        ? toNumberExact(tagStack[base + 3], valueStack[base + 3])
        : 0.0;
    const upperBound = modern
      ? argc >= 6
        ? toNumberExact(tagStack[base + 5], valueStack[base + 5])
        : 1.0
      : argc >= 5
        ? toNumberExact(tagStack[base + 4], valueStack[base + 4])
        : 1.0;
    const result =
      !isNaN(x) &&
      !isNaN(alpha) &&
      !isNaN(beta) &&
      cumulative >= 0 &&
      !isNaN(lowerBound) &&
      !isNaN(upperBound)
        ? cumulative == 1 || !modern
          ? betaDistributionCdf(x, alpha, beta, lowerBound, upperBound)
          : betaDistributionDensity(x, alpha, beta, lowerBound, upperBound)
        : NaN;
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      isNumericResult(result) ? <u8>ValueTag.Number : <u8>ValueTag.Error,
      isNumericResult(result) ? result : ErrorCode.Value,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (
    (builtinId == BuiltinId.BetaInv || builtinId == BuiltinId.Betainv) &&
    argc >= 3 &&
    argc <= 5
  ) {
    if (!rangeSupportedScalarOnly(base, argc, kindStack)) {
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
    const probability = toNumberExact(tagStack[base], valueStack[base]);
    const alpha = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
    const beta = toNumberExact(tagStack[base + 2], valueStack[base + 2]);
    const lowerBound = argc >= 4 ? toNumberExact(tagStack[base + 3], valueStack[base + 3]) : 0.0;
    const upperBound = argc >= 5 ? toNumberExact(tagStack[base + 4], valueStack[base + 4]) : 1.0;
    const result = betaDistributionInverse(probability, alpha, beta, lowerBound, upperBound);
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      isNumericResult(result) ? <u8>ValueTag.Number : <u8>ValueTag.Error,
      isNumericResult(result) ? result : ErrorCode.Value,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.FDist && argc == 4) {
    if (!rangeSupportedScalarOnly(base, argc, kindStack)) {
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
    const x = toNumberExact(tagStack[base], valueStack[base]);
    const degrees1Raw = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
    const degrees2Raw = toNumberExact(tagStack[base + 2], valueStack[base + 2]);
    const cumulative = coerceBoolean(tagStack[base + 3], valueStack[base + 3]);
    const degrees1 = Math.floor(degrees1Raw);
    const degrees2 = Math.floor(degrees2Raw);
    const result =
      !isNaN(x) && !isNaN(degrees1Raw) && !isNaN(degrees2Raw) && cumulative >= 0
        ? cumulative == 1
          ? fDistributionCdf(x, degrees1, degrees2)
          : fDistributionDensity(x, degrees1, degrees2)
        : NaN;
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      isNumericResult(result) ? <u8>ValueTag.Number : <u8>ValueTag.Error,
      isNumericResult(result) ? result : ErrorCode.Value,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (
    (builtinId == BuiltinId.FDistRt ||
      builtinId == BuiltinId.Fdist ||
      builtinId == BuiltinId.LegacyFdist) &&
    argc == 3
  ) {
    if (!rangeSupportedScalarOnly(base, argc, kindStack)) {
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
    const x = toNumberExact(tagStack[base], valueStack[base]);
    const degrees1Raw = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
    const degrees2Raw = toNumberExact(tagStack[base + 2], valueStack[base + 2]);
    const degrees1 = Math.floor(degrees1Raw);
    const degrees2 = Math.floor(degrees2Raw);
    const cdf = fDistributionCdf(x, degrees1, degrees2);
    const result =
      !isNaN(x) && !isNaN(degrees1Raw) && !isNaN(degrees2Raw) && isFinite(cdf) ? 1.0 - cdf : NaN;
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      isNumericResult(result) ? <u8>ValueTag.Number : <u8>ValueTag.Error,
      isNumericResult(result) ? result : ErrorCode.Value,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (
    (builtinId == BuiltinId.FInv ||
      builtinId == BuiltinId.FInvRt ||
      builtinId == BuiltinId.Finv ||
      builtinId == BuiltinId.LegacyFinv) &&
    argc == 3
  ) {
    if (!rangeSupportedScalarOnly(base, argc, kindStack)) {
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
    const probabilityRaw = toNumberExact(tagStack[base], valueStack[base]);
    const degrees1Raw = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
    const degrees2Raw = toNumberExact(tagStack[base + 2], valueStack[base + 2]);
    const degrees1 = Math.floor(degrees1Raw);
    const degrees2 = Math.floor(degrees2Raw);
    const probability = builtinId == BuiltinId.FInv ? probabilityRaw : 1.0 - probabilityRaw;
    const result = inverseFDistribution(probability, degrees1, degrees2);
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      isNumericResult(result) ? <u8>ValueTag.Number : <u8>ValueTag.Error,
      isNumericResult(result) ? result : ErrorCode.Value,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.TDist && argc == 3) {
    if (!rangeSupportedScalarOnly(base, argc, kindStack)) {
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
    const x = toNumberExact(tagStack[base], valueStack[base]);
    const degreesRaw = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
    const cumulative = coerceBoolean(tagStack[base + 2], valueStack[base + 2]);
    const degrees = Math.floor(degreesRaw);
    const result =
      !isNaN(x) && !isNaN(degreesRaw) && cumulative >= 0
        ? cumulative == 1
          ? studentTCdf(x, degrees)
          : studentTDensity(x, degrees)
        : NaN;
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      isNumericResult(result) ? <u8>ValueTag.Number : <u8>ValueTag.Error,
      isNumericResult(result) ? result : ErrorCode.Value,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if ((builtinId == BuiltinId.TDistRt || builtinId == BuiltinId.TDist2T) && argc == 2) {
    if (!rangeSupportedScalarOnly(base, argc, kindStack)) {
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
    const x = toNumberExact(tagStack[base], valueStack[base]);
    const degreesRaw = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
    const degrees = Math.floor(degreesRaw);
    const upperTail = 1.0 - studentTCdf(x, degrees);
    const result =
      !isNaN(x) &&
      !isNaN(degreesRaw) &&
      (builtinId != BuiltinId.TDist2T || x >= 0.0) &&
      isFinite(upperTail)
        ? builtinId == BuiltinId.TDistRt
          ? upperTail
          : min(1.0, upperTail * 2.0)
        : NaN;
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      isNumericResult(result) ? <u8>ValueTag.Number : <u8>ValueTag.Error,
      isNumericResult(result) ? result : ErrorCode.Value,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Tdist && argc == 3) {
    if (!rangeSupportedScalarOnly(base, argc, kindStack)) {
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
    const x = toNumberExact(tagStack[base], valueStack[base]);
    const degreesRaw = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
    const tailsRaw = toNumberExact(tagStack[base + 2], valueStack[base + 2]);
    const degrees = Math.floor(degreesRaw);
    const tails = <i32>tailsRaw;
    const upperTail = 1.0 - studentTCdf(x, degrees);
    const result =
      !isNaN(x) &&
      !isNaN(degreesRaw) &&
      !isNaN(tailsRaw) &&
      tailsRaw == <f64>tails &&
      x >= 0.0 &&
      (tails == 1 || tails == 2) &&
      isFinite(upperTail)
        ? tails == 1
          ? upperTail
          : min(1.0, upperTail * 2.0)
        : NaN;
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      isNumericResult(result) ? <u8>ValueTag.Number : <u8>ValueTag.Error,
      isNumericResult(result) ? result : ErrorCode.Value,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (
    (builtinId == BuiltinId.TInv || builtinId == BuiltinId.TInv2T || builtinId == BuiltinId.Tinv) &&
    argc == 2
  ) {
    if (!rangeSupportedScalarOnly(base, argc, kindStack)) {
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
    const probabilityRaw = toNumberExact(tagStack[base], valueStack[base]);
    const degreesRaw = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
    const degrees = Math.floor(degreesRaw);
    const probability = builtinId == BuiltinId.TInv ? probabilityRaw : 1.0 - probabilityRaw / 2.0;
    const result = inverseStudentT(probability, degrees);
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      isNumericResult(result) ? <u8>ValueTag.Number : <u8>ValueTag.Error,
      isNumericResult(result) ? result : ErrorCode.Value,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (
    (builtinId == BuiltinId.BinomDistRange && (argc == 3 || argc == 4)) ||
    (builtinId == BuiltinId.Critbinom && argc == 3) ||
    (builtinId == BuiltinId.BinomInv && argc == 3) ||
    (builtinId == BuiltinId.Hypgeomdist && argc == 4) ||
    (builtinId == BuiltinId.HypgeomDist && argc == 5)
  ) {
    if (!rangeSupportedScalarOnly(base, argc, kindStack)) {
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
    let result = NaN;
    if (builtinId == BuiltinId.BinomDistRange) {
      const trialsRaw = toNumberExact(tagStack[base], valueStack[base]);
      const probability = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
      const lowerRaw = toNumberExact(tagStack[base + 2], valueStack[base + 2]);
      const upperRaw =
        argc == 4 ? toNumberExact(tagStack[base + 3], valueStack[base + 3]) : lowerRaw;
      const trials = <i32>trialsRaw;
      const lower = <i32>lowerRaw;
      const upper = <i32>upperRaw;
      if (!isNaN(trialsRaw) && !isNaN(lowerRaw) && !isNaN(upperRaw) && lower <= upper) {
        result = 0.0;
        for (let index = lower; index <= upper; index += 1) {
          result += binomialProbability(index, trials, probability);
        }
      }
    } else if (builtinId == BuiltinId.Critbinom || builtinId == BuiltinId.BinomInv) {
      const trialsRaw = toNumberExact(tagStack[base], valueStack[base]);
      const probability = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
      const alpha = toNumberExact(tagStack[base + 2], valueStack[base + 2]);
      const trials = <i32>trialsRaw;
      if (
        !isNaN(trialsRaw) &&
        !isNaN(probability) &&
        !isNaN(alpha) &&
        trials >= 0 &&
        probability >= 0.0 &&
        probability <= 1.0 &&
        alpha > 0.0 &&
        alpha < 1.0
      ) {
        let cumulative = 0.0;
        for (let index = 0; index <= trials; index += 1) {
          cumulative += binomialProbability(index, trials, probability);
          if (cumulative >= alpha) {
            result = <f64>index;
            break;
          }
        }
        if (isNaN(result)) {
          result = <f64>trials;
        }
      }
    } else if (builtinId == BuiltinId.Hypgeomdist) {
      const sampleSuccessesRaw = toNumberExact(tagStack[base], valueStack[base]);
      const sampleSizeRaw = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
      const populationSuccessesRaw = toNumberExact(tagStack[base + 2], valueStack[base + 2]);
      const populationSizeRaw = toNumberExact(tagStack[base + 3], valueStack[base + 3]);
      if (
        !isNaN(sampleSuccessesRaw) &&
        !isNaN(sampleSizeRaw) &&
        !isNaN(populationSuccessesRaw) &&
        !isNaN(populationSizeRaw)
      ) {
        result = hypergeometricProbability(
          <i32>sampleSuccessesRaw,
          <i32>sampleSizeRaw,
          <i32>populationSuccessesRaw,
          <i32>populationSizeRaw,
        );
      }
    } else {
      const sampleSuccessesRaw = toNumberExact(tagStack[base], valueStack[base]);
      const sampleSizeRaw = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
      const populationSuccessesRaw = toNumberExact(tagStack[base + 2], valueStack[base + 2]);
      const populationSizeRaw = toNumberExact(tagStack[base + 3], valueStack[base + 3]);
      const cumulative = coerceBoolean(tagStack[base + 4], valueStack[base + 4]);
      const sampleSuccesses = <i32>sampleSuccessesRaw;
      const sampleSize = <i32>sampleSizeRaw;
      const populationSuccesses = <i32>populationSuccessesRaw;
      const populationSize = <i32>populationSizeRaw;
      if (
        !isNaN(sampleSuccessesRaw) &&
        !isNaN(sampleSizeRaw) &&
        !isNaN(populationSuccessesRaw) &&
        !isNaN(populationSizeRaw) &&
        cumulative >= 0
      ) {
        if (cumulative == 1) {
          result = 0.0;
          const minimum = max<i32>(0, sampleSize - (populationSize - populationSuccesses));
          for (let index = minimum; index <= sampleSuccesses; index += 1) {
            result += hypergeometricProbability(
              index,
              sampleSize,
              populationSuccesses,
              populationSize,
            );
          }
        } else {
          result = hypergeometricProbability(
            sampleSuccesses,
            sampleSize,
            populationSuccesses,
            populationSize,
          );
        }
      }
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      isNumericResult(result) ? <u8>ValueTag.Number : <u8>ValueTag.Error,
      isNumericResult(result) ? result : ErrorCode.Value,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (
    builtinId != BuiltinId.Valuetotext &&
    builtinId != BuiltinId.Address &&
    builtinId != BuiltinId.Dollar &&
    builtinId != BuiltinId.Dollarde &&
    builtinId != BuiltinId.Dollarfr
  ) {
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
  }

  if (builtinId == BuiltinId.Address && argc >= 2 && argc <= 5) {
    const row = coercePositiveIntegerArg(tagStack[base], valueStack[base], true, 1);
    const column = coercePositiveIntegerArg(tagStack[base + 1], valueStack[base + 1], true, 1);
    if (row == i32.MIN_VALUE || column == i32.MIN_VALUE) {
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
    const absNumeric = argc >= 3 ? toNumberExact(tagStack[base + 2], valueStack[base + 2]) : 1.0;
    const refStyleNumeric =
      argc >= 4 ? toNumberExact(tagStack[base + 3], valueStack[base + 3]) : 1.0;
    if (!isFinite(absNumeric) || !isFinite(refStyleNumeric)) {
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
    const absNum = <i32>absNumeric;
    const refStyle = <i32>refStyleNumeric;
    if (absNum < 1 || absNum > 4 || (refStyle != 1 && refStyle != 2)) {
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
    let sheetPrefix = "";
    if (argc == 5) {
      if (tagStack[base + 4] == ValueTag.Empty || tagStack[base + 4] != ValueTag.String) {
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
      const sheetText = scalarText(
        tagStack[base + 4],
        valueStack[base + 4],
        stringOffsets,
        stringLengths,
        stringData,
        outputStringOffsets,
        outputStringLengths,
        outputStringData,
      );
      if (sheetText == null) {
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
      sheetPrefix = `'${escapeSheetNameText(sheetText)}'!`;
    }
    const columnLabel = columnLabelText(column);
    if (columnLabel == null) {
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
    if (refStyle == 2) {
      const rowLabel = absNum == 1 || absNum == 2 ? row.toString() : `[${row}]`;
      const colLabel = absNum == 1 || absNum == 3 ? column.toString() : `[${column}]`;
      return writeStringResult(
        base,
        `${sheetPrefix}R${rowLabel}C${colLabel}`,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const rowLabel = absNum == 1 || absNum == 2 ? `$${row.toString()}` : row.toString();
    const colLabel = absNum == 1 || absNum == 3 ? `$${columnLabel}` : columnLabel;
    return writeStringResult(
      base,
      `${sheetPrefix}${colLabel}${rowLabel}`,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Dollar && argc >= 1 && argc <= 3) {
    const value = toNumberExact(tagStack[base], valueStack[base]);
    const decimalsNumeric =
      argc >= 2 ? toNumberExact(tagStack[base + 1], valueStack[base + 1]) : 2.0;
    let noCommasValue = 0.0;
    if (argc >= 3) {
      const numeric = toNumberExact(tagStack[base + 2], valueStack[base + 2]);
      noCommasValue = isNaN(numeric) ? 0.0 : numeric;
    }
    if (!isFinite(value) || !isFinite(decimalsNumeric)) {
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
    const text = formatFixedText(value, <i32>decimalsNumeric, noCommasValue == 0.0);
    if (text == null || text.length == 0) {
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
    const normalizedText = text.startsWith("-") ? text.slice(1) : text;
    return writeStringResult(
      base,
      value < 0.0 ? `-$${normalizedText}` : `$${text}`,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if ((builtinId == BuiltinId.Dollarde || builtinId == BuiltinId.Dollarfr) && argc == 2) {
    const value = toNumberExact(tagStack[base], valueStack[base]);
    const fractionNumeric = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
    if (!isFinite(value) || !isFinite(fractionNumeric)) {
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
    const fraction = <i32>fractionNumeric;
    if (!isValidDollarFractionNative(fraction)) {
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
    if (builtinId == BuiltinId.Dollarde) {
      const integerPart = <i32>Math.floor(Math.abs(value));
      const fractionalNumerator = dollarFractionalNumerator(value);
      if (fractionalNumerator == i32.MIN_VALUE || fractionalNumerator >= fraction) {
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
      const sign = value < 0.0 ? -1.0 : 1.0;
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Number,
        sign * (<f64>integerPart + <f64>fractionalNumerator / <f64>fraction),
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const sign = value < 0.0 ? -1.0 : 1.0;
    const absolute = Math.abs(value);
    const integerPart = <i32>Math.floor(absolute);
    const fractional = absolute - <f64>integerPart;
    const width = digitCount(fraction);
    const scaledNumerator = <i32>Math.round(fractional * <f64>fraction);
    const carry = scaledNumerator / fraction;
    const numerator = scaledNumerator - carry * fraction;
    const outputValue = <f64>(integerPart + carry) + <f64>numerator / Math.pow(10.0, <f64>width);
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      sign * outputValue,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Base && (argc == 2 || argc == 3)) {
    const numberNumeric = toNumberExact(tagStack[base], valueStack[base]);
    const radixNumeric = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
    const minLengthNumeric =
      argc == 3 ? toNumberExact(tagStack[base + 2], valueStack[base + 2]) : 0.0;
    if (!isFinite(numberNumeric) || !isFinite(radixNumeric) || !isFinite(minLengthNumeric)) {
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
    const numberValue = <i64>numberNumeric;
    const radixValue = <i32>radixNumeric;
    const minLength = <i32>minLengthNumeric;
    if (numberValue < 0 || radixValue < 2 || radixValue > 36 || minLength < 0) {
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
    return writeStringResult(
      base,
      toBaseText(numberValue, radixValue, minLength),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Decimal && argc == 2) {
    if (tagStack[base] == ValueTag.Error) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        valueStack[base],
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const radixNumeric = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
    if (!isFinite(radixNumeric)) {
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
    const radixValue = <i32>radixNumeric;
    if (radixValue < 2 || radixValue > 36) {
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
    let raw = "";
    if (tagStack[base] == ValueTag.String) {
      const text = scalarText(
        tagStack[base],
        valueStack[base],
        stringOffsets,
        stringLengths,
        stringData,
        outputStringOffsets,
        outputStringLengths,
        outputStringData,
      );
      if (text == null) {
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
      raw = trimAsciiWhitespace(text);
    } else {
      const numeric = toNumberExact(tagStack[base], valueStack[base]);
      if (!isFinite(numeric)) {
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
      raw = (<i64>numeric).toString();
    }
    if (raw.length == 0 || !isValidBaseText(raw, radixValue)) {
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
      parseBaseText(raw, radixValue),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (
    (builtinId == BuiltinId.Bin2dec ||
      builtinId == BuiltinId.Hex2dec ||
      builtinId == BuiltinId.Oct2dec) &&
    argc == 1
  ) {
    const radix = builtinId == BuiltinId.Bin2dec ? 2 : builtinId == BuiltinId.Hex2dec ? 16 : 8;
    const raw = signedRadixInputText(
      tagStack[base],
      valueStack[base],
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    );
    if (raw == null) {
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
    const numeric = parseSignedRadixText(raw, radix, 10);
    if (numeric == i64.MIN_VALUE) {
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
      <f64>numeric,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (
    (builtinId == BuiltinId.Bin2hex ||
      builtinId == BuiltinId.Bin2oct ||
      builtinId == BuiltinId.Dec2bin ||
      builtinId == BuiltinId.Dec2hex ||
      builtinId == BuiltinId.Dec2oct ||
      builtinId == BuiltinId.Hex2bin ||
      builtinId == BuiltinId.Hex2oct ||
      builtinId == BuiltinId.Oct2bin ||
      builtinId == BuiltinId.Oct2hex) &&
    (argc == 1 || argc == 2)
  ) {
    let numeric: i64 = i64.MIN_VALUE;
    let radix = 10;
    let negativeWidth = 10;
    let minValue: i64 = -549755813888;
    let maxValue: i64 = 549755813887;

    if (builtinId == BuiltinId.Bin2hex || builtinId == BuiltinId.Bin2oct) {
      const raw = signedRadixInputText(
        tagStack[base],
        valueStack[base],
        stringOffsets,
        stringLengths,
        stringData,
        outputStringOffsets,
        outputStringLengths,
        outputStringData,
      );
      if (raw != null) {
        numeric = parseSignedRadixText(raw, 2, 10);
      }
      if (builtinId == BuiltinId.Bin2oct) {
        radix = 8;
        minValue = -536870912;
        maxValue = 536870911;
      } else {
        radix = 16;
      }
    } else if (builtinId == BuiltinId.Hex2bin || builtinId == BuiltinId.Hex2oct) {
      const raw = signedRadixInputText(
        tagStack[base],
        valueStack[base],
        stringOffsets,
        stringLengths,
        stringData,
        outputStringOffsets,
        outputStringLengths,
        outputStringData,
      );
      if (raw != null) {
        numeric = parseSignedRadixText(raw, 16, 10);
      }
      if (builtinId == BuiltinId.Hex2bin) {
        radix = 2;
        minValue = -512;
        maxValue = 511;
      } else {
        radix = 8;
        minValue = -536870912;
        maxValue = 536870911;
      }
    } else if (builtinId == BuiltinId.Oct2bin || builtinId == BuiltinId.Oct2hex) {
      const raw = signedRadixInputText(
        tagStack[base],
        valueStack[base],
        stringOffsets,
        stringLengths,
        stringData,
        outputStringOffsets,
        outputStringLengths,
        outputStringData,
      );
      if (raw != null) {
        numeric = parseSignedRadixText(raw, 8, 10);
      }
      if (builtinId == BuiltinId.Oct2bin) {
        radix = 2;
        minValue = -512;
        maxValue = 511;
      } else {
        radix = 16;
      }
    } else {
      const inputNumeric = toNumberExact(tagStack[base], valueStack[base]);
      if (isFinite(inputNumeric) && Math.abs(inputNumeric) <= MAX_SAFE_INTEGER_F64) {
        numeric = <i64>inputNumeric;
      }
      if (builtinId == BuiltinId.Dec2bin) {
        radix = 2;
        minValue = -512;
        maxValue = 511;
      } else if (builtinId == BuiltinId.Dec2oct) {
        radix = 8;
        minValue = -536870912;
        maxValue = 536870911;
      } else {
        radix = 16;
      }
    }

    const places = argc == 2 ? coerceNonNegativeShift(tagStack[base + 1], valueStack[base + 1]) : 0;
    if (numeric == i64.MIN_VALUE || places == i64.MIN_VALUE) {
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
    const text = formatSignedRadixText(
      numeric,
      radix,
      <i32>places,
      negativeWidth,
      minValue,
      maxValue,
    );
    if (text == null) {
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
    return writeStringResult(base, text, rangeIndexStack, valueStack, tagStack, kindStack);
  }

  if (builtinId == BuiltinId.Convert && argc == 3) {
    const numeric = toNumberExact(tagStack[base], valueStack[base]);
    if (
      !isFinite(numeric) ||
      tagStack[base + 1] != ValueTag.String ||
      tagStack[base + 2] != ValueTag.String
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
    const fromText = scalarText(
      tagStack[base + 1],
      valueStack[base + 1],
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    );
    const toText = scalarText(
      tagStack[base + 2],
      valueStack[base + 2],
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    );
    if (fromText == null || toText == null || !resolveConvertUnit(fromText)) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.NA,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const fromGroup = resolvedConvertGroup;
    const fromFactor = resolvedConvertFactor;
    const fromTemperature = resolvedConvertTemperature;
    if (!resolveConvertUnit(toText)) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.NA,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const toGroup = resolvedConvertGroup;
    const toFactor = resolvedConvertFactor;
    const toTemperature = resolvedConvertTemperature;
    let result = NaN;
    if (fromGroup == CONVERT_GROUP_TEMPERATURE || toGroup == CONVERT_GROUP_TEMPERATURE) {
      if (fromGroup != CONVERT_GROUP_TEMPERATURE || toGroup != CONVERT_GROUP_TEMPERATURE) {
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          ErrorCode.NA,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }
      result = convertKelvinToTemperature(
        toTemperature,
        convertTemperatureToKelvin(fromTemperature, numeric),
      );
    } else {
      if (fromGroup != toGroup) {
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          ErrorCode.NA,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }
      result = (numeric * fromFactor) / toFactor;
    }
    if (!isFinite(result)) {
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
      result,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Euroconvert && argc >= 3 && argc <= 5) {
    const numeric = toNumberExact(tagStack[base], valueStack[base]);
    if (
      !isFinite(numeric) ||
      tagStack[base + 1] != ValueTag.String ||
      tagStack[base + 2] != ValueTag.String
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
    const sourceText = scalarText(
      tagStack[base + 1],
      valueStack[base + 1],
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    );
    const targetText = scalarText(
      tagStack[base + 2],
      valueStack[base + 2],
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    );
    const fullPrecisionNumeric =
      argc >= 4 ? toNumberExact(tagStack[base + 3], valueStack[base + 3]) : 0.0;
    const triangulationNumeric =
      argc == 5 ? toNumberExact(tagStack[base + 4], valueStack[base + 4]) : NaN;
    if (
      sourceText == null ||
      targetText == null ||
      !isFinite(fullPrecisionNumeric) ||
      (argc == 5 && (!isFinite(triangulationNumeric) || triangulationNumeric < 3.0))
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
    const sourceRate = euroRateNative(sourceText);
    const targetRate = euroRateNative(targetText);
    const targetPrecision = euroCalculationPrecisionNative(targetText);
    if (!isFinite(sourceRate) || !isFinite(targetRate) || targetPrecision == i32.MIN_VALUE) {
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
    if (sourceText == targetText) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Number,
        numeric,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    let euroValue = sourceText == "EUR" ? numeric : numeric / sourceRate;
    if (sourceText != "EUR" && argc == 5) {
      euroValue = roundToSignificantDigitsNative(euroValue, <i32>triangulationNumeric);
    }
    let result = targetText == "EUR" ? euroValue : euroValue * targetRate;
    if (fullPrecisionNumeric == 0.0) {
      result = roundToPlacesNative(result, targetPrecision);
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

  const scalarMathResult = tryApplyScalarMathBuiltin(
    builtinId,
    argc,
    base,
    rangeIndexStack,
    valueStack,
    tagStack,
    kindStack,
  );
  if (scalarMathResult >= 0) {
    return scalarMathResult;
  }

  const logicInfoResult = tryApplyLogicInfoBuiltin(
    builtinId,
    argc,
    base,
    rangeIndexStack,
    valueStack,
    tagStack,
    kindStack,
    stringOffsets,
    stringLengths,
    stringData,
    outputStringOffsets,
    outputStringLengths,
    outputStringData,
  );
  if (logicInfoResult >= 0) {
    return logicInfoResult;
  }

  const scalarTextResult = tryApplyScalarTextBuiltin(
    builtinId,
    argc,
    base,
    rangeIndexStack,
    valueStack,
    tagStack,
    kindStack,
    stringOffsets,
    stringLengths,
    stringData,
    outputStringOffsets,
    outputStringLengths,
    outputStringData,
  );
  if (scalarTextResult >= 0) {
    return scalarTextResult;
  }

  const textFormattingResult = tryApplyTextFormattingBuiltin(
    builtinId,
    argc,
    base,
    rangeIndexStack,
    valueStack,
    tagStack,
    kindStack,
    cellTags,
    cellNumbers,
    cellStringIds,
    cellErrors,
    stringOffsets,
    stringLengths,
    stringData,
    rangeOffsets,
    rangeLengths,
    rangeMembers,
    outputStringOffsets,
    outputStringLengths,
    outputStringData,
  );
  if (textFormattingResult >= 0) {
    return textFormattingResult;
  }

  const textMutationResult = tryApplyTextMutationBuiltin(
    builtinId,
    argc,
    base,
    rangeIndexStack,
    valueStack,
    tagStack,
    kindStack,
    stringOffsets,
    stringLengths,
    stringData,
    outputStringOffsets,
    outputStringLengths,
    outputStringData,
  );
  if (textMutationResult >= 0) {
    return textMutationResult;
  }

  const dateTimeResult = tryApplyDateTimeBuiltin(
    builtinId,
    argc,
    base,
    rangeIndexStack,
    valueStack,
    tagStack,
    kindStack,
    stringOffsets,
    stringLengths,
    stringData,
    outputStringOffsets,
    outputStringLengths,
    outputStringData,
  );
  if (dateTimeResult >= 0) {
    return dateTimeResult;
  }

  const dateCalendarResult = tryApplyDateCalendarBuiltin(
    builtinId,
    argc,
    base,
    rangeIndexStack,
    valueStack,
    tagStack,
    kindStack,
    cellTags,
    cellNumbers,
    cellStringIds,
    cellErrors,
    stringOffsets,
    stringLengths,
    stringData,
    rangeOffsets,
    rangeLengths,
    rangeMembers,
    outputStringOffsets,
    outputStringLengths,
    outputStringData,
  );
  if (dateCalendarResult >= 0) {
    return dateCalendarResult;
  }

  const financeSecuritiesResult = tryApplyFinanceSecuritiesBuiltin(
    builtinId,
    argc,
    base,
    rangeIndexStack,
    valueStack,
    tagStack,
    kindStack,
  );
  if (financeSecuritiesResult >= 0) {
    return financeSecuritiesResult;
  }

  const financeCashflowResult = tryApplyFinanceCashflowBuiltin(
    builtinId,
    argc,
    base,
    rangeIndexStack,
    valueStack,
    tagStack,
    kindStack,
    cellTags,
    cellNumbers,
    cellStringIds,
    cellErrors,
    rangeOffsets,
    rangeLengths,
    rangeRowCounts,
    rangeColCounts,
    rangeMembers,
  );
  if (financeCashflowResult >= 0) {
    return financeCashflowResult;
  }

  const depreciationResult = tryApplyDepreciationBuiltin(
    builtinId,
    argc,
    base,
    rangeIndexStack,
    valueStack,
    tagStack,
    kindStack,
  );
  if (depreciationResult >= 0) {
    return depreciationResult;
  }

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
