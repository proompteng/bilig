import { BuiltinId, ErrorCode, ValueTag } from './protocol'
import { registerTrackedArrayShape as registerTrackedArrayShapeImpl } from './dynamic-arrays'
import { databaseBuiltinResult } from './database-criteria'
import { tryApplyAggregateCriteriaBuiltin } from './dispatch-aggregate-criteria'
import { tryApplyTextMutationBuiltin } from './dispatch-text-mutation'
import { tryApplyTextFormattingBuiltin } from './dispatch-text-formatting'
import { tryApplyDateTimeBuiltin } from './dispatch-date-time'
import { tryApplyDateCalendarBuiltin } from './dispatch-date-calendar'
import { tryApplyDepreciationBuiltin } from './dispatch-depreciation'
import { tryApplyFinanceCashflowBuiltin } from './dispatch-finance-cashflows'
import { tryApplyFinanceSecuritiesBuiltin } from './dispatch-finance-securities'
import { tryApplyScalarTextBuiltin } from './dispatch-text-scalar'
import { rangeSupportedScalarOnly, scalarErrorAt } from './builtin-args'
import { STACK_KIND_RANGE, STACK_KIND_SCALAR, UNRESOLVED_WASM_OPERAND, writeResult } from './result-io'
import { tryApplyArrayFoundationBuiltin } from './dispatch-array-foundation'
import { tryApplyArrayFilterBuiltin } from './dispatch-array-filter'
import { tryApplyArrayInfoBuiltin } from './dispatch-array-info'
import { tryApplyArrayOrderingBuiltin } from './dispatch-array-ordering'
import { tryApplyArrayReshapeBuiltin } from './dispatch-array-reshape'
import { tryApplyArrayUniqueBuiltin } from './dispatch-array-unique'
import { tryApplyArrayWindowBuiltin } from './dispatch-array-window'
import { tryApplyLookupMatchBuiltin } from './dispatch-lookup-match'
import { tryApplyRegressionBuiltin } from './dispatch-regression'
import { tryApplyScalarDistributionBuiltin } from './dispatch-distributions-scalar'
import { tryApplyExtendedDistributionBuiltin } from './dispatch-distributions-extended'
import { tryApplyFormatConvertBuiltin } from './dispatch-format-convert'
import { tryApplyStatisticsSummaryBuiltin } from './dispatch-statistics-summary'
import { tryApplyStatisticalTestBuiltin } from './dispatch-statistics-tests'
import { tryApplyScalarMathBuiltin } from './dispatch-scalar-math'
import { tryApplyLogicInfoBuiltin } from './dispatch-logic-info'
import { tryApplyLookupTableBuiltin } from './dispatch-lookup-table'
import { tryApplySpecialRuntimeBuiltin } from './dispatch-special-runtime'

export function registerTrackedArrayShape(arrayIndex: u32, rows: i32, cols: i32): void {
  registerTrackedArrayShapeImpl(arrayIndex, rows, cols)
}

function unresolvedRangeOperandError(base: i32, argc: i32, kindStack: Uint8Array, rangeIndexStack: Uint32Array): f64 {
  for (let index = 0; index < argc; index++) {
    const slot = base + index
    if (kindStack[slot] == STACK_KIND_RANGE && rangeIndexStack[slot] == UNRESOLVED_WASM_OPERAND) {
      return ErrorCode.Ref
    }
  }
  return -1
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
  const base = sp - argc
  const unresolvedRangeError = unresolvedRangeOperandError(base, argc, kindStack, rangeIndexStack)
  if (unresolvedRangeError >= 0) {
    return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, unresolvedRangeError, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  const statisticalTestResult = tryApplyStatisticalTestBuiltin(
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
  )
  if (statisticalTestResult >= 0) {
    return statisticalTestResult
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
  )
  if (arrayFoundationResult >= 0) {
    return arrayFoundationResult
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
    cellStringIds,
    cellErrors,
    stringOffsets,
    stringLengths,
    stringData,
    outputStringOffsets,
    outputStringLengths,
    outputStringData,
  )
  if (arrayOrderingResult >= 0) {
    return arrayOrderingResult
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
    cellStringIds,
    cellErrors,
  )
  if (arrayReshapeResult >= 0) {
    return arrayReshapeResult
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
  )
  if (specialRuntimeResult >= 0) {
    return specialRuntimeResult
  }

  const arrayFilterResult = tryApplyArrayFilterBuiltin(
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
  )
  if (arrayFilterResult >= 0) {
    return arrayFilterResult
  }

  const arrayUniqueResult = tryApplyArrayUniqueBuiltin(
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
  )
  if (arrayUniqueResult >= 0) {
    return arrayUniqueResult
  }

  const arrayWindowResult = tryApplyArrayWindowBuiltin(
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
  )
  if (arrayWindowResult >= 0) {
    return arrayWindowResult
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
  )
  if (lookupTableResult >= 0) {
    return lookupTableResult
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
  )
  if (aggregateCriteriaResult >= 0) {
    return aggregateCriteriaResult
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
    )
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
  )
  if (regressionResult >= 0) {
    return regressionResult
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
  )
  if (statisticsSummaryResult >= 0) {
    return statisticsSummaryResult
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
  )
  if (lookupMatchResult >= 0) {
    return lookupMatchResult
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
  )
  if (arrayInfoResult >= 0) {
    return arrayInfoResult
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
  )
  if (scalarDistributionResult >= 0) {
    return scalarDistributionResult
  }

  const extendedDistributionResult = tryApplyExtendedDistributionBuiltin(
    builtinId,
    argc,
    base,
    rangeIndexStack,
    valueStack,
    tagStack,
    kindStack,
  )
  if (extendedDistributionResult >= 0) {
    return extendedDistributionResult
  }

  if (!rangeSupportedScalarOnly(base, argc, kindStack)) {
    return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (
    builtinId != BuiltinId.Valuetotext &&
    builtinId != BuiltinId.Address &&
    builtinId != BuiltinId.Dollar &&
    builtinId != BuiltinId.Dollarde &&
    builtinId != BuiltinId.Dollarfr
  ) {
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack)
    if (scalarError >= 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, scalarError, rangeIndexStack, valueStack, tagStack, kindStack)
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
  )
  if (formatConvertResult >= 0) {
    return formatConvertResult
  }

  const scalarMathResult = tryApplyScalarMathBuiltin(
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
  )
  if (scalarMathResult >= 0) {
    return scalarMathResult
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
  )
  if (logicInfoResult >= 0) {
    return logicInfoResult
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
  )
  if (scalarTextResult >= 0) {
    return scalarTextResult
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
  )
  if (textFormattingResult >= 0) {
    return textFormattingResult
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
  )
  if (textMutationResult >= 0) {
    return textMutationResult
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
  )
  if (dateTimeResult >= 0) {
    return dateTimeResult
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
  )
  if (dateCalendarResult >= 0) {
    return dateCalendarResult
  }

  const financeSecuritiesResult = tryApplyFinanceSecuritiesBuiltin(builtinId, argc, base, rangeIndexStack, valueStack, tagStack, kindStack)
  if (financeSecuritiesResult >= 0) {
    return financeSecuritiesResult
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
  )
  if (financeCashflowResult >= 0) {
    return financeCashflowResult
  }

  const depreciationResult = tryApplyDepreciationBuiltin(builtinId, argc, base, rangeIndexStack, valueStack, tagStack, kindStack)
  if (depreciationResult >= 0) {
    return depreciationResult
  }

  return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
}
