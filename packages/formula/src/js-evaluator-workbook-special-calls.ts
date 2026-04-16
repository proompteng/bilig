import { ErrorCode, type CellValue } from '@bilig/protocol'
import { parseCellAddress, parseRangeAddress } from './addressing.js'
import { evaluateGroupBy, evaluatePivotBy } from './group-pivot-evaluator.js'
import { isArrayValue } from './runtime-values.js'
import type { EvaluationContext, ReferenceOperand, StackValue } from './js-evaluator.js'

interface MatrixLikeValue {
  rows: number
  cols: number
  values: readonly CellValue[]
}

interface WorkbookSpecialCallDeps {
  error: (code: ErrorCode) => CellValue
  stackScalar: (value: CellValue) => StackValue
  toStringValue: (value: CellValue) => string
  isSingleCellValue: (value: StackValue) => CellValue | undefined
  matrixFromStackValue: (value: StackValue) => MatrixLikeValue | undefined
  scalarIntegerArgument: (value: StackValue | undefined) => number | undefined
  vectorIntegerArgument: (value: StackValue | undefined) => number[] | undefined
  aggregateRangeSubset: (
    functionArg: StackValue,
    subset: readonly CellValue[],
    context: EvaluationContext,
    totalSet?: readonly CellValue[],
  ) => CellValue
  referenceTopLeftAddress: (ref: ReferenceOperand | undefined) => string | undefined
  referenceSheetName: (ref: ReferenceOperand | undefined, context: EvaluationContext) => string | undefined
  coerceScalarTextArgument: (value: StackValue | undefined) => string | CellValue
  coerceOptionalBooleanArgument: (value: StackValue | undefined, fallback: boolean) => boolean | CellValue
  isCellValueError: (value: number | boolean | string | CellValue) => value is CellValue
}

export function evaluateWorkbookSpecialCall(
  callee: string,
  rawArgs: StackValue[],
  context: EvaluationContext,
  argRefs: readonly (ReferenceOperand | undefined)[],
  deps: WorkbookSpecialCallDeps,
): StackValue | undefined {
  switch (callee) {
    case 'GETPIVOTDATA': {
      if (rawArgs.length < 2 || (rawArgs.length - 2) % 2 !== 0) {
        return deps.stackScalar(deps.error(ErrorCode.Value))
      }
      const dataFieldValue = deps.isSingleCellValue(rawArgs[0]!)
      const address = deps.referenceTopLeftAddress(argRefs[1])
      const sheetName = deps.referenceSheetName(argRefs[1], context)
      if (!dataFieldValue) {
        return deps.stackScalar(deps.error(ErrorCode.Value))
      }
      if (!address || !sheetName) {
        return deps.stackScalar(deps.error(ErrorCode.Ref))
      }
      const filters: Array<{ field: string; item: CellValue }> = []
      for (let index = 2; index < rawArgs.length; index += 2) {
        const fieldValue = deps.isSingleCellValue(rawArgs[index]!)
        const itemValue = deps.isSingleCellValue(rawArgs[index + 1]!)
        if (!fieldValue || !itemValue) {
          return deps.stackScalar(deps.error(ErrorCode.Value))
        }
        filters.push({ field: deps.toStringValue(fieldValue), item: itemValue })
      }
      return deps.stackScalar(
        context.resolvePivotData?.({
          dataField: deps.toStringValue(dataFieldValue),
          sheetName,
          address,
          filters,
        }) ?? deps.error(ErrorCode.Ref),
      )
    }
    case 'GROUPBY': {
      if (rawArgs.length < 3 || rawArgs.length > 8) {
        return deps.stackScalar(deps.error(ErrorCode.Value))
      }
      const rowFields = deps.matrixFromStackValue(rawArgs[0]!)
      const values = deps.matrixFromStackValue(rawArgs[1]!)
      if (!rowFields || !values) {
        return deps.stackScalar(deps.error(ErrorCode.Value))
      }
      const sortOrder =
        deps.vectorIntegerArgument(rawArgs[5]) ?? (rawArgs[5] ? [deps.scalarIntegerArgument(rawArgs[5]) ?? Number.NaN] : undefined)
      const fieldHeadersMode = deps.scalarIntegerArgument(rawArgs[3])
      const totalDepth = deps.scalarIntegerArgument(rawArgs[4])
      const filterArray = rawArgs[6] ? deps.matrixFromStackValue(rawArgs[6]) : undefined
      const fieldRelationship = deps.scalarIntegerArgument(rawArgs[7])
      const result = evaluateGroupBy(rowFields, values, {
        aggregate: (subset: readonly CellValue[], totalSet?: readonly CellValue[]) =>
          deps.aggregateRangeSubset(rawArgs[2]!, subset, context, totalSet),
        ...(fieldHeadersMode !== undefined ? { fieldHeadersMode } : {}),
        ...(totalDepth !== undefined ? { totalDepth } : {}),
        ...(sortOrder?.every(Number.isFinite) ? { sortOrder } : {}),
        ...(filterArray !== undefined ? { filterArray } : {}),
        ...(fieldRelationship !== undefined ? { fieldRelationship } : {}),
      })
      return isArrayValue(result) ? result : deps.stackScalar(result)
    }
    case 'PIVOTBY': {
      if (rawArgs.length < 4 || rawArgs.length > 11) {
        return deps.stackScalar(deps.error(ErrorCode.Value))
      }
      const rowFields = deps.matrixFromStackValue(rawArgs[0]!)
      const colFields = deps.matrixFromStackValue(rawArgs[1]!)
      const values = deps.matrixFromStackValue(rawArgs[2]!)
      if (!rowFields || !colFields || !values) {
        return deps.stackScalar(deps.error(ErrorCode.Value))
      }
      const rowSortOrder =
        deps.vectorIntegerArgument(rawArgs[6]) ?? (rawArgs[6] ? [deps.scalarIntegerArgument(rawArgs[6]) ?? Number.NaN] : undefined)
      const colSortOrder =
        deps.vectorIntegerArgument(rawArgs[8]) ?? (rawArgs[8] ? [deps.scalarIntegerArgument(rawArgs[8]) ?? Number.NaN] : undefined)
      const fieldHeadersMode = deps.scalarIntegerArgument(rawArgs[4])
      const rowTotalDepth = deps.scalarIntegerArgument(rawArgs[5])
      const colTotalDepth = deps.scalarIntegerArgument(rawArgs[7])
      const filterArray = rawArgs[9] ? deps.matrixFromStackValue(rawArgs[9]) : undefined
      const relativeTo = deps.scalarIntegerArgument(rawArgs[10])
      const result = evaluatePivotBy(rowFields, colFields, values, {
        aggregate: (subset: readonly CellValue[], totalSet?: readonly CellValue[]) =>
          deps.aggregateRangeSubset(rawArgs[3]!, subset, context, totalSet),
        ...(fieldHeadersMode !== undefined ? { fieldHeadersMode } : {}),
        ...(rowTotalDepth !== undefined ? { rowTotalDepth } : {}),
        ...(rowSortOrder?.every(Number.isFinite) ? { rowSortOrder } : {}),
        ...(colTotalDepth !== undefined ? { colTotalDepth } : {}),
        ...(colSortOrder?.every(Number.isFinite) ? { colSortOrder } : {}),
        ...(filterArray !== undefined ? { filterArray } : {}),
        ...(relativeTo !== undefined ? { relativeTo } : {}),
      })
      return isArrayValue(result) ? result : deps.stackScalar(result)
    }
    case 'MULTIPLE.OPERATIONS': {
      if (rawArgs.length !== 3 && rawArgs.length !== 5) {
        return deps.stackScalar(deps.error(ErrorCode.Value))
      }
      const formulaAddress = deps.referenceTopLeftAddress(argRefs[0])
      const formulaSheetName = deps.referenceSheetName(argRefs[0], context)
      const rowCellAddress = deps.referenceTopLeftAddress(argRefs[1])
      const rowCellSheetName = deps.referenceSheetName(argRefs[1], context)
      const rowReplacementAddress = deps.referenceTopLeftAddress(argRefs[2])
      const rowReplacementSheetName = deps.referenceSheetName(argRefs[2], context)
      if (
        !formulaAddress ||
        !formulaSheetName ||
        !rowCellAddress ||
        !rowCellSheetName ||
        !rowReplacementAddress ||
        !rowReplacementSheetName
      ) {
        return deps.stackScalar(deps.error(ErrorCode.Ref))
      }
      const columnCellAddress = rawArgs.length === 5 ? deps.referenceTopLeftAddress(argRefs[3]) : undefined
      const columnCellSheetName = rawArgs.length === 5 ? deps.referenceSheetName(argRefs[3], context) : undefined
      const columnReplacementAddress = rawArgs.length === 5 ? deps.referenceTopLeftAddress(argRefs[4]) : undefined
      const columnReplacementSheetName = rawArgs.length === 5 ? deps.referenceSheetName(argRefs[4], context) : undefined
      if (
        rawArgs.length === 5 &&
        (!columnCellAddress || !columnCellSheetName || !columnReplacementAddress || !columnReplacementSheetName)
      ) {
        return deps.stackScalar(deps.error(ErrorCode.Ref))
      }
      return deps.stackScalar(
        context.resolveMultipleOperations?.({
          formulaSheetName,
          formulaAddress,
          rowCellSheetName,
          rowCellAddress,
          rowReplacementSheetName,
          rowReplacementAddress,
          ...(columnCellSheetName ? { columnCellSheetName } : {}),
          ...(columnCellAddress ? { columnCellAddress } : {}),
          ...(columnReplacementSheetName ? { columnReplacementSheetName } : {}),
          ...(columnReplacementAddress ? { columnReplacementAddress } : {}),
        }) ?? deps.error(ErrorCode.Ref),
      )
    }
    case 'INDIRECT': {
      if (rawArgs.length < 1 || rawArgs.length > 2) {
        return deps.stackScalar(deps.error(ErrorCode.Value))
      }
      const refText = deps.coerceScalarTextArgument(rawArgs[0])
      if (deps.isCellValueError(refText)) {
        return deps.stackScalar(refText)
      }
      const a1Mode = deps.coerceOptionalBooleanArgument(rawArgs[1], true)
      if (deps.isCellValueError(a1Mode)) {
        return deps.stackScalar(a1Mode)
      }
      if (!a1Mode) {
        return deps.stackScalar(deps.error(ErrorCode.Value))
      }
      const normalizedRefText = refText.trim()
      if (normalizedRefText === '') {
        return deps.stackScalar(deps.error(ErrorCode.Ref))
      }

      try {
        const cell = parseCellAddress(normalizedRefText, context.sheetName)
        return deps.stackScalar(context.resolveCell(cell.sheetName ?? context.sheetName, cell.text))
      } catch {
        // fall through
      }

      try {
        const range = parseRangeAddress(normalizedRefText, context.sheetName)
        if (range.kind !== 'cells') {
          return deps.stackScalar(deps.error(ErrorCode.Ref))
        }
        const targetSheetName = range.sheetName ?? context.sheetName
        const values = context.resolveRange(targetSheetName, range.start.text, range.end.text, 'cells')
        return {
          kind: 'range',
          values,
          refKind: 'cells',
          rows: range.end.row - range.start.row + 1,
          cols: range.end.col - range.start.col + 1,
        }
      } catch {
        // fall through
      }

      const resolvedName = context.resolveName?.(normalizedRefText)
      return deps.stackScalar(resolvedName ?? deps.error(ErrorCode.Ref))
    }
    default:
      return undefined
  }
}
