import { ErrorCode, MAX_COLS, MAX_ROWS, ValueTag, type CellValue } from '@bilig/protocol'
import { formatAddress, parseCellAddress, parseRangeAddress } from './addressing.js'
import { getBuiltin } from './builtins.js'
import { getLookupBuiltin, type RangeBuiltinArgument } from './builtins/lookup.js'
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

function valueError(): CellValue {
  return { tag: ValueTag.Error, code: ErrorCode.Value }
}

type WholeAxisRefKind = 'rows' | 'cols'
type WholeAxisRangeAddress = Extract<ReturnType<typeof parseRangeAddress>, { kind: WholeAxisRefKind }>

interface WholeAxisReference {
  sheetName: string
  start: string
  end: string
  refKind: WholeAxisRefKind
  parsed: WholeAxisRangeAddress
}

type ScalarArgumentResult = { kind: 'ok'; value: CellValue | undefined } | { kind: 'error'; value: CellValue }
type IntegerArgumentResult = { kind: 'omitted' } | { kind: 'ok'; value: number } | { kind: 'error'; value: CellValue }

function isWholeAxisRefKind(refKind: 'cells' | 'rows' | 'cols' | undefined): refKind is WholeAxisRefKind {
  return refKind === 'rows' || refKind === 'cols'
}

function wholeAxisReferenceFromArg(
  value: StackValue | undefined,
  ref: ReferenceOperand | undefined,
  context: EvaluationContext,
  deps: WorkbookSpecialCallDeps,
): WholeAxisReference | undefined {
  const refKind = isWholeAxisRefKind(ref?.refKind)
    ? ref.refKind
    : value?.kind === 'range' && isWholeAxisRefKind(value.refKind)
      ? value.refKind
      : undefined
  if (!refKind) {
    return undefined
  }

  const start = (ref?.kind === 'range' ? ref.start : undefined) ?? (value?.kind === 'range' ? value.start : undefined)
  const end = (ref?.kind === 'range' ? ref.end : undefined) ?? (value?.kind === 'range' ? value.end : undefined)
  if (!start || !end) {
    return undefined
  }

  const sheetName = deps.referenceSheetName(ref, context) ?? (value?.kind === 'range' ? value.sheetName : undefined) ?? context.sheetName
  try {
    const parsed = parseRangeAddress(`${start}:${end}`, sheetName)
    if (parsed.kind !== refKind) {
      return undefined
    }
    return {
      sheetName,
      start,
      end,
      refKind,
      parsed,
    }
  } catch {
    return undefined
  }
}

function scalarLookupArgument(value: StackValue | undefined, deps: WorkbookSpecialCallDeps): ScalarArgumentResult {
  if (!value || (value.kind === 'omitted' && value.source === 'argument')) {
    return { kind: 'ok', value: undefined }
  }
  const scalar = deps.isSingleCellValue(value)
  if (!scalar) {
    return { kind: 'error', value: valueError() }
  }
  return { kind: 'ok', value: scalar }
}

function integerIndexArgument(value: StackValue | undefined, deps: WorkbookSpecialCallDeps): IntegerArgumentResult {
  if (!value || (value.kind === 'omitted' && value.source === 'argument')) {
    return { kind: 'omitted' }
  }
  const scalar = deps.isSingleCellValue(value)
  if (!scalar) {
    return { kind: 'error', value: valueError() }
  }
  if (scalar.tag === ValueTag.Error) {
    return { kind: 'error', value: scalar }
  }
  const integer = deps.scalarIntegerArgument(value)
  return integer === undefined ? { kind: 'error', value: valueError() } : { kind: 'ok', value: integer }
}

function wholeAxisLookupRange(
  reference: WholeAxisReference,
  context: EvaluationContext,
  deps: WorkbookSpecialCallDeps,
): RangeBuiltinArgument | CellValue {
  const values = context.resolveRange(reference.sheetName, reference.start, reference.end, reference.refKind)
  if (reference.parsed.kind === 'rows') {
    const rowCount = reference.parsed.end.row - reference.parsed.start.row + 1
    if (rowCount !== 1) {
      return deps.error(ErrorCode.Value)
    }
    return {
      kind: 'range',
      values,
      refKind: 'cells',
      rows: 1,
      cols: values.length,
      sheetName: reference.sheetName,
      start: formatAddress(reference.parsed.start.row, 0),
      end: formatAddress(reference.parsed.start.row, Math.max(values.length - 1, 0)),
    }
  }

  const colCount = reference.parsed.end.col - reference.parsed.start.col + 1
  if (colCount !== 1) {
    return deps.error(ErrorCode.Value)
  }
  return {
    kind: 'range',
    values,
    refKind: 'cells',
    rows: values.length,
    cols: 1,
    sheetName: reference.sheetName,
    start: formatAddress(0, reference.parsed.start.col),
    end: formatAddress(Math.max(values.length - 1, 0), reference.parsed.start.col),
  }
}

function evaluateWholeAxisMatch(
  callee: 'MATCH' | 'XMATCH',
  rawArgs: StackValue[],
  context: EvaluationContext,
  argRefs: readonly (ReferenceOperand | undefined)[],
  deps: WorkbookSpecialCallDeps,
): StackValue | undefined {
  const reference = wholeAxisReferenceFromArg(rawArgs[1], argRefs[1], context, deps)
  if (!reference) {
    return undefined
  }
  if (rawArgs.length < 2 || (callee === 'MATCH' ? rawArgs.length > 3 : rawArgs.length > 4)) {
    return deps.stackScalar(deps.error(ErrorCode.Value))
  }
  const lookupValue = deps.isSingleCellValue(rawArgs[0]!)
  if (!lookupValue) {
    return deps.stackScalar(deps.error(ErrorCode.Value))
  }
  const lookupRange = wholeAxisLookupRange(reference, context, deps)
  if ('tag' in lookupRange) {
    return deps.stackScalar(lookupRange)
  }

  const lookupBuiltin = context.resolveLookupBuiltin?.(callee) ?? getLookupBuiltin(callee)
  if (!lookupBuiltin) {
    return undefined
  }

  const firstOptional = scalarLookupArgument(rawArgs[2], deps)
  if (firstOptional.kind === 'error') {
    return deps.stackScalar(firstOptional.value)
  }
  if (callee === 'MATCH') {
    const result = lookupBuiltin(lookupValue, lookupRange, firstOptional.value)
    return isArrayValue(result) ? result : deps.stackScalar(result)
  }

  const secondOptional = scalarLookupArgument(rawArgs[3], deps)
  if (secondOptional.kind === 'error') {
    return deps.stackScalar(secondOptional.value)
  }
  const result = lookupBuiltin(lookupValue, lookupRange, firstOptional.value, secondOptional.value)
  return isArrayValue(result) ? result : deps.stackScalar(result)
}

function evaluateWholeAxisIndex(
  rawArgs: StackValue[],
  context: EvaluationContext,
  argRefs: readonly (ReferenceOperand | undefined)[],
  deps: WorkbookSpecialCallDeps,
): StackValue | undefined {
  const reference = wholeAxisReferenceFromArg(rawArgs[0], argRefs[0], context, deps)
  if (!reference) {
    return undefined
  }
  if (rawArgs.length < 1 || rawArgs.length > 3) {
    return deps.stackScalar(deps.error(ErrorCode.Value))
  }

  const rowArg = integerIndexArgument(rawArgs[1], deps)
  if (rowArg.kind === 'error') {
    return deps.stackScalar(rowArg.value)
  }
  const colArg = integerIndexArgument(rawArgs[2], deps)
  if (colArg.kind === 'error') {
    return deps.stackScalar(colArg.value)
  }

  const rawRowNum = rowArg.kind === 'ok' ? rowArg.value : undefined
  const rawColNum = colArg.kind === 'ok' ? colArg.value : undefined
  if (rawRowNum === undefined && rawColNum === undefined) {
    return deps.stackScalar(deps.error(ErrorCode.Value))
  }

  if (reference.parsed.kind === 'cols') {
    const colCount = reference.parsed.end.col - reference.parsed.start.col + 1
    const rowNum = rawRowNum ?? 0
    const colNum = rawColNum ?? (colCount === 1 && rowNum !== 0 ? 1 : 0)
    if (rowNum <= 0 || colNum <= 0) {
      return deps.stackScalar(deps.error(ErrorCode.Value))
    }
    if (rowNum > MAX_ROWS || colNum > colCount) {
      return deps.stackScalar(deps.error(ErrorCode.Ref))
    }
    const targetRow = rowNum - 1
    const targetCol = reference.parsed.start.col + colNum - 1
    if (targetCol >= MAX_COLS) {
      return deps.stackScalar(deps.error(ErrorCode.Ref))
    }
    return deps.stackScalar(context.resolveCell(reference.sheetName, formatAddress(targetRow, targetCol)))
  }

  const rowCount = reference.parsed.end.row - reference.parsed.start.row + 1
  let rowNum = rawRowNum ?? 0
  let colNum = rawColNum ?? 0
  if (rowCount === 1 && rawColNum === undefined && rawRowNum !== undefined && rawRowNum !== 0) {
    rowNum = 1
    colNum = rawRowNum
  }
  if (rowNum <= 0 || colNum <= 0) {
    return deps.stackScalar(deps.error(ErrorCode.Value))
  }
  if (rowNum > rowCount || colNum > MAX_COLS) {
    return deps.stackScalar(deps.error(ErrorCode.Ref))
  }
  const targetRow = reference.parsed.start.row + rowNum - 1
  const targetCol = colNum - 1
  if (targetRow >= MAX_ROWS) {
    return deps.stackScalar(deps.error(ErrorCode.Ref))
  }
  return deps.stackScalar(context.resolveCell(reference.sheetName, formatAddress(targetRow, targetCol)))
}

function hiddenAwareSubtotalValues(value: StackValue, ref: ReferenceOperand | undefined, context: EvaluationContext): CellValue[] {
  if (value.kind === 'scalar') {
    if (ref?.kind !== 'cell' || !ref.address) {
      return [value.value]
    }
    const sheetName = ref.sheetName ?? context.sheetName
    const row = parseCellAddress(ref.address, sheetName).row
    return context.isRowHidden?.(sheetName, row) === true ? [] : [value.value]
  }
  if (value.kind === 'omitted' || value.kind === 'lambda') {
    return [valueError()]
  }
  if (value.kind === 'array') {
    return [...value.values]
  }
  if (value.refKind !== 'cells') {
    return [...value.values]
  }

  const sheetName = ref?.sheetName ?? value.sheetName ?? context.sheetName
  const start = ref?.kind === 'range' ? ref.start : value.start
  const end = ref?.kind === 'range' ? ref.end : value.end
  if (!start || !end) {
    return [...value.values]
  }

  let startRow = 0
  let cols = value.cols
  try {
    const parsed = parseRangeAddress(`${start}:${end}`, sheetName)
    if (parsed.kind !== 'cells') {
      return [...value.values]
    }
    startRow = parsed.start.row
    cols = parsed.end.col - parsed.start.col + 1
  } catch {
    return [...value.values]
  }

  const visibleValues: CellValue[] = []
  for (let index = 0; index < value.values.length; index += 1) {
    const row = startRow + Math.floor(index / cols)
    if (context.isRowHidden?.(sheetName, row) === true) {
      continue
    }
    visibleValues.push(value.values[index]!)
  }
  return visibleValues
}

export function evaluateWorkbookSpecialCall(
  callee: string,
  rawArgs: StackValue[],
  context: EvaluationContext,
  argRefs: readonly (ReferenceOperand | undefined)[],
  deps: WorkbookSpecialCallDeps,
): StackValue | undefined {
  switch (callee) {
    case 'MATCH':
    case 'XMATCH':
      return evaluateWholeAxisMatch(callee, rawArgs, context, argRefs, deps)
    case 'INDEX':
      return evaluateWholeAxisIndex(rawArgs, context, argRefs, deps)
    case 'SUBTOTAL': {
      const functionNum = deps.scalarIntegerArgument(rawArgs[0])
      if (functionNum === undefined) {
        return deps.stackScalar(deps.error(ErrorCode.Value))
      }
      if (functionNum <= 100 || !context.isRowHidden) {
        return undefined
      }
      const subtotal = getBuiltin('SUBTOTAL')
      if (!subtotal) {
        return undefined
      }
      const values = rawArgs.slice(1).flatMap((value, index) => hiddenAwareSubtotalValues(value, argRefs[index + 1], context))
      const result = subtotal({ tag: ValueTag.Number, value: functionNum }, ...values)
      return isArrayValue(result) ? result : deps.stackScalar(result)
    }
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
