import { BuiltinId, ErrorCode, ValueTag } from "./protocol";
import {
  getTrackedArrayCols as getDynamicArrayCols,
  getTrackedArrayRows as getDynamicArrayRows,
  registerTrackedArrayShape as registerTrackedArrayShapeImpl,
} from "./dynamic-arrays";
import { scalarText } from "./text-codec";
import { arrayToTextCell, parseNumericText } from "./text-special";
import { coerceLength, coercePositiveStart } from "./text-foundation";
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
  interestPaymentCalc,
  periodicCashflowNetPresentValueCalc,
  periodicPaymentCalc,
  presentValueCalc,
  principalPaymentCalc,
  solveRateCalc,
  totalPeriodsCalc,
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
  chiSquareTestPValue,
  collectDateSeriesFromSlot,
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
  readSpillArrayTag,
  readSpillArrayLength,
  readSpillArrayNumber,
  writeSpillArrayNumber,
  writeSpillArrayValue,
} from "./vm";
import { tryApplyArrayFoundationBuiltin } from "./dispatch-array-foundation";
import { tryApplyArrayInfoBuiltin } from "./dispatch-array-info";
import { tryApplyArrayOrderingBuiltin } from "./dispatch-array-ordering";
import { tryApplyArrayReshapeBuiltin } from "./dispatch-array-reshape";
import { tryApplyLookupMatchBuiltin } from "./dispatch-lookup-match";
import { tryApplyRegressionBuiltin } from "./dispatch-regression";
import { tryApplyScalarDistributionBuiltin } from "./dispatch-distributions-scalar";
import { tryApplyExtendedDistributionBuiltin } from "./dispatch-distributions-extended";
import { tryApplyFormatConvertBuiltin } from "./dispatch-format-convert";
import { tryApplyStatisticsSummaryBuiltin } from "./dispatch-statistics-summary";
import { tryApplyScalarMathBuiltin } from "./dispatch-scalar-math";
import { tryApplyLogicInfoBuiltin } from "./dispatch-logic-info";
import { tryApplyLookupTableBuiltin } from "./dispatch-lookup-table";
import { tryApplySpecialRuntimeBuiltin } from "./dispatch-special-runtime";

export function registerTrackedArrayShape(arrayIndex: u32, rows: i32, cols: i32): void {
  registerTrackedArrayShapeImpl(arrayIndex, rows, cols);
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

  const arrayOrderingResult = tryApplyArrayOrderingBuiltin(
    builtinId,
    argc,
    base,
    rangeIndexStack,
    valueStack,
    tagStack,
    kindStack,
    rangeOffsets,
    rangeLengths,
    rangeRowCounts,
    rangeColCounts,
    rangeMembers,
    cellTags,
    cellNumbers,
    stringOffsets,
    stringLengths,
    stringData,
    outputStringOffsets,
    outputStringLengths,
    outputStringData,
  );
  if (arrayOrderingResult >= 0) {
    return arrayOrderingResult;
  }

  const arrayReshapeResult = tryApplyArrayReshapeBuiltin(
    builtinId,
    argc,
    base,
    rangeIndexStack,
    valueStack,
    tagStack,
    kindStack,
    rangeOffsets,
    rangeLengths,
    rangeRowCounts,
    rangeColCounts,
    rangeMembers,
    cellTags,
    cellNumbers,
  );
  if (arrayReshapeResult >= 0) {
    return arrayReshapeResult;
  }

  const specialRuntimeResult = tryApplySpecialRuntimeBuiltin(
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
  if (specialRuntimeResult >= 0) {
    return specialRuntimeResult;
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

  const extendedDistributionResult = tryApplyExtendedDistributionBuiltin(
    builtinId,
    argc,
    base,
    rangeIndexStack,
    valueStack,
    tagStack,
    kindStack,
  );
  if (extendedDistributionResult >= 0) {
    return extendedDistributionResult;
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

  const formatConvertResult = tryApplyFormatConvertBuiltin(
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
  if (formatConvertResult >= 0) {
    return formatConvertResult;
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
