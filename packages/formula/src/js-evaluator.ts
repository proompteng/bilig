import { ErrorCode, ValueTag, formatErrorCode, type CellValue } from '@bilig/protocol'
import type { FormulaNode } from './ast.js'
import { parseRangeAddress } from './addressing.js'
import { getBuiltin } from './builtins.js'
import { getLookupBuiltin, type LookupBuiltin, type RangeBuiltinArgument } from './builtins/lookup.js'
import type { MatrixValue } from './group-pivot-evaluator.js'
import { evaluateArraySpecialCall } from './js-evaluator-array-special-calls.js'
import { evaluateContextSpecialCall } from './js-evaluator-context-special-calls.js'
import {
  absoluteAddress,
  cellTypeCode,
  currentCellReference,
  referenceColumnNumber,
  referenceRowNumber,
  referenceSheetName,
  referenceTopLeftAddress,
  sheetIndexByName,
  sheetNames,
} from './js-evaluator-reference-context.js'
import { evaluateWorkbookSpecialCall } from './js-evaluator-workbook-special-calls.js'
import { lowerToPlan } from './js-plan-lowering.js'
import { isArrayValue, scalarFromEvaluationResult, type ArrayValue, type EvaluationResult, type RangeLikeValue } from './runtime-values.js'

export interface EvaluationContext {
  sheetName: string
  currentAddress?: string
  resolveCell: (sheetName: string, address: string) => CellValue
  resolveRange: (sheetName: string, start: string, end: string, refKind: 'cells' | 'rows' | 'cols') => CellValue[]
  resolveName?: (name: string) => CellValue
  resolveFormula?: (sheetName: string, address: string) => string | undefined
  resolvePivotData?: (request: {
    dataField: string
    sheetName: string
    address: string
    filters: ReadonlyArray<{ field: string; item: CellValue }>
  }) => CellValue | undefined
  resolveMultipleOperations?: (request: {
    formulaSheetName: string
    formulaAddress: string
    rowCellSheetName: string
    rowCellAddress: string
    rowReplacementSheetName: string
    rowReplacementAddress: string
    columnCellSheetName?: string
    columnCellAddress?: string
    columnReplacementSheetName?: string
    columnReplacementAddress?: string
  }) => CellValue | undefined
  resolveExactVectorMatch?: (request: {
    lookupValue: CellValue
    sheetName: string
    start: string
    end: string
    startRow: number
    endRow: number
    startCol: number
    endCol: number
    searchMode: 1 | -1
  }) => ExactVectorMatchResult
  resolveApproximateVectorMatch?: (request: {
    lookupValue: CellValue
    sheetName: string
    start: string
    end: string
    startRow: number
    endRow: number
    startCol: number
    endCol: number
    matchMode: 1 | -1
  }) => ApproximateVectorMatchResult
  noteRangeMaterialization?: (cellCount: number) => void
  noteExactLookupDirect?: () => void
  noteExactLookupFallback?: () => void
  listSheetNames?: () => string[]
  resolveBuiltin?: (name: string) => ((...args: CellValue[]) => EvaluationResult) | undefined
  resolveLookupBuiltin?: (name: string) => LookupBuiltin | undefined
}

export type ExactVectorMatchResult = { handled: false } | { handled: true; position: number | undefined }

export type ApproximateVectorMatchResult = ExactVectorMatchResult

export interface ReferenceOperand {
  kind: 'cell' | 'range' | 'row' | 'col'
  sheetName?: string
  address?: string
  start?: string
  end?: string
  refKind?: 'cells' | 'rows' | 'cols'
}

export type JsPlanInstruction =
  | { opcode: 'push-number'; value: number }
  | { opcode: 'push-boolean'; value: boolean }
  | { opcode: 'push-string'; value: string }
  | { opcode: 'push-error'; code: ErrorCode }
  | { opcode: 'push-name'; name: string }
  | { opcode: 'push-cell'; sheetName?: string; address: string }
  | {
      opcode: 'push-range'
      sheetName?: string
      start: string
      end: string
      refKind: 'cells' | 'rows' | 'cols'
    }
  | {
      opcode: 'lookup-exact-match'
      callee: 'MATCH' | 'XMATCH'
      sheetName?: string
      start: string
      end: string
      startRow: number
      endRow: number
      startCol: number
      endCol: number
      refKind: 'cells'
      searchMode: 1 | -1
    }
  | {
      opcode: 'lookup-approximate-match'
      callee: 'MATCH' | 'XMATCH'
      sheetName?: string
      start: string
      end: string
      startRow: number
      endRow: number
      startCol: number
      endCol: number
      refKind: 'cells'
      matchMode: 1 | -1
    }
  | { opcode: 'push-lambda'; params: string[]; body: JsPlanInstruction[] }
  | { opcode: 'unary'; operator: '+' | '-' }
  | {
      opcode: 'binary'
      operator: '+' | '-' | '*' | '/' | '^' | '&' | '=' | '<>' | '>' | '>=' | '<' | '<='
    }
  | {
      opcode: 'call'
      callee: string
      argc: number
      argRefs?: Array<ReferenceOperand | undefined>
    }
  | { opcode: 'invoke'; argc: number }
  | { opcode: 'begin-scope' }
  | { opcode: 'bind-name'; name: string }
  | { opcode: 'end-scope' }
  | { opcode: 'jump-if-false'; target: number }
  | { opcode: 'jump'; target: number }
  | { opcode: 'return' }

export type StackValue =
  | { kind: 'scalar'; value: CellValue }
  | { kind: 'omitted' }
  | {
      kind: 'range'
      values: CellValue[]
      refKind: 'cells' | 'rows' | 'cols'
      rows: number
      cols: number
      sheetName?: string
      start?: string
      end?: string
    }
  | {
      kind: 'lambda'
      params: string[]
      body: JsPlanInstruction[]
      scopes: Array<Map<string, StackValue>>
    }
  | ArrayValue
type BinaryOperator = Extract<JsPlanInstruction, { opcode: 'binary' }>['operator']
export { lowerToPlan } from './js-plan-lowering.js'

function emptyValue(): CellValue {
  return { tag: ValueTag.Empty }
}

function error(code: ErrorCode): CellValue {
  return { tag: ValueTag.Error, code }
}

function numberValue(value: number): CellValue {
  return { tag: ValueTag.Number, value }
}

function stringValue(value: string): CellValue {
  return { tag: ValueTag.String, value, stringId: 0 }
}

function toNumber(value: CellValue): number | undefined {
  switch (value.tag) {
    case ValueTag.Number:
      return value.value
    case ValueTag.Boolean:
      return value.value ? 1 : 0
    case ValueTag.Empty:
      return 0
    case ValueTag.String:
    case ValueTag.Error:
      return undefined
    default:
      return undefined
  }
}

function toStringValue(value: CellValue): string {
  switch (value.tag) {
    case ValueTag.Empty:
      return ''
    case ValueTag.Number:
      return String(value.value)
    case ValueTag.Boolean:
      return value.value ? 'TRUE' : 'FALSE'
    case ValueTag.String:
      return value.value
    case ValueTag.Error:
      return formatErrorCode(value.code)
  }
}

function isTextLike(value: CellValue): boolean {
  return value.tag === ValueTag.String || value.tag === ValueTag.Empty
}

function compareText(left: string, right: string): number {
  const normalizedLeft = left.toUpperCase()
  const normalizedRight = right.toUpperCase()
  if (normalizedLeft === normalizedRight) {
    return 0
  }
  return normalizedLeft < normalizedRight ? -1 : 1
}

function compareScalars(left: CellValue, right: CellValue): number | undefined {
  if (isTextLike(left) && isTextLike(right)) {
    return compareText(toStringValue(left), toStringValue(right))
  }

  const leftNum = toNumber(left)
  const rightNum = toNumber(right)
  if (leftNum === undefined || rightNum === undefined) {
    return undefined
  }
  if (leftNum === rightNum) {
    return 0
  }
  return leftNum < rightNum ? -1 : 1
}

function truthy(value: CellValue): boolean {
  return (toNumber(value) ?? 0) !== 0
}

function popScalar(stack: StackValue[]): CellValue {
  const value = stack.pop()
  if (!value) {
    return error(ErrorCode.Value)
  }
  if (value.kind === 'scalar') {
    return value.value
  }
  if (value.kind === 'omitted') {
    return error(ErrorCode.Value)
  }
  if (value.kind === 'lambda') {
    return error(ErrorCode.Value)
  }
  return value.values[0] ?? emptyValue()
}

function popArgument(stack: StackValue[]): StackValue {
  return stack.pop() ?? { kind: 'scalar', value: error(ErrorCode.Value) }
}

function toEvaluationResult(value: StackValue | undefined): EvaluationResult {
  if (!value) {
    return error(ErrorCode.Value)
  }
  if (value.kind === 'scalar') {
    return value.value
  }
  if (value.kind === 'omitted') {
    return error(ErrorCode.Value)
  }
  if (value.kind === 'lambda') {
    return error(ErrorCode.Value)
  }
  if (value.kind === 'range') {
    return {
      kind: 'array',
      rows: value.rows,
      cols: value.cols,
      values: value.values,
    }
  }
  return value
}

function cloneStackValue(value: StackValue): StackValue {
  if (value.kind === 'scalar') {
    return { kind: 'scalar', value: value.value }
  }
  if (value.kind === 'omitted') {
    return { kind: 'omitted' }
  }
  if (value.kind === 'range') {
    return {
      kind: 'range',
      values: value.values,
      refKind: value.refKind,
      rows: value.rows,
      cols: value.cols,
      ...(value.sheetName ? { sheetName: value.sheetName } : {}),
      ...(value.start ? { start: value.start } : {}),
      ...(value.end ? { end: value.end } : {}),
    }
  }
  if (value.kind === 'lambda') {
    return {
      kind: 'lambda',
      params: [...value.params],
      body: value.body,
      scopes: cloneScopes(value.scopes),
    }
  }
  return { kind: 'array', values: value.values, rows: value.rows, cols: value.cols }
}

function cloneScopes(scopes: readonly Map<string, StackValue>[]): Array<Map<string, StackValue>> {
  return scopes.map((scope) => new Map([...scope.entries()].map(([name, value]) => [name, cloneStackValue(value)])))
}

function toRangeLike(value: StackValue): RangeLikeValue {
  if (value.kind === 'omitted') {
    return { kind: 'range', values: [error(ErrorCode.Value)], rows: 1, cols: 1, refKind: 'cells' }
  }
  if (value.kind === 'lambda') {
    return { kind: 'range', values: [error(ErrorCode.Value)], rows: 1, cols: 1, refKind: 'cells' }
  }
  if (value.kind === 'range') {
    return value
  }
  if (value.kind === 'array') {
    return {
      kind: 'range',
      values: value.values,
      rows: value.rows,
      cols: value.cols,
      refKind: 'cells',
    }
  }
  return { kind: 'range', values: [value.value], rows: 1, cols: 1, refKind: 'cells' }
}

function scalarBinary(operator: BinaryOperator, leftValue: CellValue, rightValue: CellValue): CellValue {
  if (leftValue.tag === ValueTag.Error) {
    return leftValue
  }
  if (rightValue.tag === ValueTag.Error) {
    return rightValue
  }

  if (operator === '&') {
    return {
      tag: ValueTag.String,
      value: `${toStringValue(leftValue)}${toStringValue(rightValue)}`,
      stringId: 0,
    }
  }

  if (['+', '-', '*', '/', '^'].includes(operator)) {
    const left = toNumber(leftValue)
    const right = toNumber(rightValue)
    if (left === undefined || right === undefined) {
      return error(ErrorCode.Value)
    }
    if (operator === '/' && right === 0) {
      return error(ErrorCode.Div0)
    }
    const value =
      operator === '+'
        ? left + right
        : operator === '-'
          ? left - right
          : operator === '*'
            ? left * right
            : operator === '/'
              ? left / right
              : left ** right
    return { tag: ValueTag.Number, value }
  }

  const comparison = compareScalars(leftValue, rightValue)
  if (comparison === undefined) {
    return error(ErrorCode.Value)
  }
  return {
    tag: ValueTag.Boolean,
    value:
      operator === '='
        ? comparison === 0
        : operator === '<>'
          ? comparison !== 0
          : operator === '>'
            ? comparison > 0
            : operator === '>='
              ? comparison >= 0
              : operator === '<'
                ? comparison < 0
                : comparison <= 0,
  }
}

function evaluateBinary(operator: BinaryOperator, leftValue: StackValue, rightValue: StackValue): EvaluationResult {
  if (leftValue.kind === 'scalar' && rightValue.kind === 'scalar') {
    return scalarBinary(operator, leftValue.value, rightValue.value)
  }

  const leftRange = toRangeLike(leftValue)
  const rightRange = toRangeLike(rightValue)
  const rows =
    leftRange.rows === rightRange.rows
      ? leftRange.rows
      : leftRange.rows === 1
        ? rightRange.rows
        : rightRange.rows === 1
          ? leftRange.rows
          : 0
  const cols =
    leftRange.cols === rightRange.cols
      ? leftRange.cols
      : leftRange.cols === 1
        ? rightRange.cols
        : rightRange.cols === 1
          ? leftRange.cols
          : 0
  if (rows === 0 || cols === 0) {
    return error(ErrorCode.Value)
  }

  const values: CellValue[] = []
  for (let row = 0; row < rows; row += 1) {
    for (let col = 0; col < cols; col += 1) {
      const leftIndex = Math.min(row, leftRange.rows - 1) * leftRange.cols + Math.min(col, leftRange.cols - 1)
      const rightIndex = Math.min(row, rightRange.rows - 1) * rightRange.cols + Math.min(col, rightRange.cols - 1)
      values.push(scalarBinary(operator, leftRange.values[leftIndex] ?? emptyValue(), rightRange.values[rightIndex] ?? emptyValue()))
    }
  }
  return rows === 1 && cols === 1 ? (values[0] ?? emptyValue()) : { kind: 'array', values, rows, cols }
}

function stackScalar(value: CellValue): StackValue {
  return { kind: 'scalar', value }
}

function normalizeScopeName(name: string): string {
  return name.toUpperCase()
}

function isSingleCellValue(value: StackValue): CellValue | undefined {
  if (value.kind === 'scalar') {
    return value.value
  }
  if (value.kind === 'omitted') {
    return undefined
  }
  if (value.kind === 'lambda') {
    return undefined
  }
  return value.rows * value.cols === 1 ? (value.values[0] ?? emptyValue()) : undefined
}

function toRangeArgument(value: StackValue): CellValue | RangeBuiltinArgument {
  if (value.kind === 'scalar') {
    return value.value
  }
  if (value.kind === 'omitted') {
    return error(ErrorCode.Value)
  }
  if (value.kind === 'lambda') {
    return error(ErrorCode.Value)
  }
  return {
    kind: 'range',
    values: value.values,
    refKind: value.kind === 'range' ? value.refKind : 'cells',
    rows: value.rows,
    cols: value.cols,
    ...(value.kind === 'range' && value.sheetName ? { sheetName: value.sheetName } : {}),
    ...(value.kind === 'range' && value.start ? { start: value.start } : {}),
    ...(value.kind === 'range' && value.end ? { end: value.end } : {}),
  }
}

function toPositiveInteger(value: StackValue | undefined): number | undefined {
  if (value === undefined) {
    return undefined
  }
  const scalar = isSingleCellValue(value)
  const numeric = scalar ? toNumber(scalar) : undefined
  if (numeric === undefined || !Number.isFinite(numeric)) {
    return undefined
  }
  const integer = Math.trunc(numeric)
  return integer >= 1 ? integer : undefined
}

function getRangeCell(range: RangeLikeValue, row: number, col: number): CellValue {
  return range.values[row * range.cols + col] ?? emptyValue()
}

function getBroadcastShape(values: readonly StackValue[]): { rows: number; cols: number } | undefined {
  const ranges = values.map(toRangeLike)
  const rows = Math.max(...ranges.map((range) => range.rows))
  const cols = Math.max(...ranges.map((range) => range.cols))
  const compatible = ranges.every((range) => (range.rows === rows || range.rows === 1) && (range.cols === cols || range.cols === 1))
  return compatible ? { rows, cols } : undefined
}

function coerceScalarTextArgument(value: StackValue | undefined): string | CellValue {
  if (value === undefined) {
    return error(ErrorCode.Value)
  }
  const scalar = isSingleCellValue(value)
  if (!scalar) {
    return error(ErrorCode.Value)
  }
  if (scalar.tag === ValueTag.Error) {
    return scalar
  }
  return toStringValue(scalar)
}

function coerceOptionalBooleanArgument(value: StackValue | undefined, fallback: boolean): boolean | CellValue {
  if (value === undefined) {
    return fallback
  }
  const scalar = isSingleCellValue(value)
  if (!scalar) {
    return error(ErrorCode.Value)
  }
  if (scalar.tag === ValueTag.Error) {
    return scalar
  }
  if (scalar.tag === ValueTag.Boolean) {
    return scalar.value
  }
  const numeric = toNumber(scalar)
  return numeric === undefined ? error(ErrorCode.Value) : numeric !== 0
}

function coerceOptionalMatchModeArgument(value: StackValue | undefined, fallback: 0 | 1): 0 | 1 | CellValue {
  if (value === undefined) {
    return fallback
  }
  const scalar = isSingleCellValue(value)
  if (!scalar) {
    return error(ErrorCode.Value)
  }
  if (scalar.tag === ValueTag.Error) {
    return scalar
  }
  const numeric = toNumber(scalar)
  if (numeric === undefined || !Number.isFinite(numeric)) {
    return error(ErrorCode.Value)
  }
  const integer = Math.trunc(numeric)
  return integer === 0 || integer === 1 ? integer : error(ErrorCode.Value)
}

function coerceOptionalPositiveIntegerArgument(value: StackValue | undefined, fallback: number): number | CellValue {
  if (value === undefined) {
    return fallback
  }
  const scalar = isSingleCellValue(value)
  if (!scalar) {
    return error(ErrorCode.Value)
  }
  if (scalar.tag === ValueTag.Error) {
    return scalar
  }
  const numeric = toNumber(scalar)
  if (numeric === undefined || !Number.isFinite(numeric)) {
    return error(ErrorCode.Value)
  }
  const integer = Math.trunc(numeric)
  return integer >= 1 ? integer : error(ErrorCode.Value)
}

function coerceOptionalTrimModeArgument(value: StackValue | undefined, fallback: 0 | 1 | 2 | 3): 0 | 1 | 2 | 3 | CellValue {
  if (value === undefined) {
    return fallback
  }
  const scalar = isSingleCellValue(value)
  if (!scalar) {
    return error(ErrorCode.Value)
  }
  if (scalar.tag === ValueTag.Error) {
    return scalar
  }
  const numeric = toNumber(scalar)
  if (numeric === undefined || !Number.isFinite(numeric)) {
    return error(ErrorCode.Value)
  }
  const integer = Math.trunc(numeric)
  switch (integer) {
    case 0:
    case 1:
    case 2:
    case 3:
      return integer
    default:
      return error(ErrorCode.Value)
  }
}

function isCellValueError(value: number | boolean | string | CellValue): value is CellValue {
  return typeof value === 'object' && value !== null && 'tag' in value
}

function makeArrayStack(rows: number, cols: number, values: CellValue[]): StackValue {
  return { kind: 'array', rows, cols, values }
}

function matrixFromStackValue(value: StackValue): MatrixValue | undefined {
  if (value.kind === 'omitted' || value.kind === 'lambda') {
    return undefined
  }
  if (value.kind === 'scalar') {
    return { rows: 1, cols: 1, values: [value.value] }
  }
  return { rows: value.rows, cols: value.cols, values: value.values }
}

function scalarIntegerArgument(value: StackValue | undefined): number | undefined {
  const scalar = value ? isSingleCellValue(value) : undefined
  const numeric = scalar ? toNumber(scalar) : undefined
  return numeric === undefined || !Number.isFinite(numeric) ? undefined : Math.trunc(numeric)
}

function vectorIntegerArgument(value: StackValue | undefined): number[] | undefined {
  if (!value) {
    return undefined
  }
  const matrix = matrixFromStackValue(value)
  if (!matrix || !(matrix.rows === 1 || matrix.cols === 1)) {
    return undefined
  }
  const values: number[] = []
  for (let index = 0; index < matrix.rows * matrix.cols; index += 1) {
    const numeric = toNumber(matrix.values[index] ?? emptyValue())
    if (numeric === undefined || !Number.isFinite(numeric)) {
      return undefined
    }
    values.push(Math.trunc(numeric))
  }
  return values
}

function aggregateRangeSubset(
  functionArg: StackValue,
  subset: readonly CellValue[],
  context: EvaluationContext,
  totalSet?: readonly CellValue[],
): CellValue {
  if (functionArg.kind === 'lambda') {
    const args: StackValue[] = [makeArrayStack(Math.max(subset.length, 1), 1, [...subset])]
    if (functionArg.params.length >= 2) {
      args.push(makeArrayStack(Math.max(totalSet?.length ?? 0, 1), 1, [...(totalSet ?? [emptyValue()])]))
    }
    const result = applyLambda(functionArg, args, context)
    return isSingleCellValue(result) ?? error(ErrorCode.Value)
  }
  const scalar = isSingleCellValue(functionArg)
  if (scalar?.tag !== ValueTag.String) {
    return scalar?.tag === ValueTag.Error ? scalar : error(ErrorCode.Value)
  }
  const name = scalar.value.trim().toUpperCase()
  if (subset.length === 0) {
    if (name === 'SUM' || name === 'COUNT' || name === 'COUNTA') {
      return numberValue(0)
    }
    if (name === 'AVERAGE' || name === 'AVG') {
      return error(ErrorCode.Div0)
    }
    return numberValue(0)
  }
  const builtin = context.resolveBuiltin?.(name) ?? getBuiltin(name)
  if (!builtin) {
    return error(ErrorCode.Name)
  }
  const result = builtin(...subset)
  return isArrayValue(result) ? scalarFromEvaluationResult(result) : result
}

function applyLambda(lambdaValue: StackValue, args: StackValue[], context: EvaluationContext): StackValue {
  if (lambdaValue.kind !== 'lambda') {
    return stackScalar(
      lambdaValue.kind === 'scalar' && lambdaValue.value.tag === ValueTag.Error ? lambdaValue.value : error(ErrorCode.Value),
    )
  }
  if (args.length > lambdaValue.params.length) {
    return stackScalar(error(ErrorCode.Value))
  }
  const parameterScope = new Map<string, StackValue>()
  lambdaValue.params.forEach((name: string, index: number) => {
    parameterScope.set(normalizeScopeName(name), index < args.length ? cloneStackValue(args[index]!) : { kind: 'omitted' })
  })
  return executePlan(lambdaValue.body, context, [...cloneScopes(lambdaValue.scopes), parameterScope]) ?? stackScalar(error(ErrorCode.Value))
}

function evaluateSpecialCall(
  callee: string,
  rawArgs: StackValue[],
  context: EvaluationContext,
  argRefs: readonly (ReferenceOperand | undefined)[] = [],
): StackValue | undefined {
  switch (callee) {
    default:
      return (
        evaluateWorkbookSpecialCall(callee, rawArgs, context, argRefs, {
          error,
          stackScalar,
          toStringValue,
          isSingleCellValue,
          matrixFromStackValue,
          scalarIntegerArgument,
          vectorIntegerArgument,
          aggregateRangeSubset,
          referenceTopLeftAddress,
          referenceSheetName,
          coerceScalarTextArgument,
          coerceOptionalBooleanArgument,
          isCellValueError,
        }) ??
        evaluateContextSpecialCall(callee, rawArgs, context, argRefs, {
          error,
          emptyValue,
          numberValue,
          stringValue,
          stackScalar,
          cloneStackValue,
          toNumber,
          toStringValue,
          isSingleCellValue,
          currentCellReference,
          referenceSheetName,
          referenceTopLeftAddress,
          referenceRowNumber,
          referenceColumnNumber,
          absoluteAddress,
          cellTypeCode,
          sheetNames,
          sheetIndexByName,
        }) ??
        evaluateArraySpecialCall(callee, rawArgs, context, {
          error,
          emptyValue,
          numberValue,
          stringValue,
          stackScalar,
          toRangeLike,
          getRangeCell,
          getBroadcastShape,
          makeArrayStack,
          applyLambda,
          toPositiveInteger,
          coerceScalarTextArgument,
          coerceOptionalBooleanArgument,
          coerceOptionalMatchModeArgument,
          coerceOptionalPositiveIntegerArgument,
          coerceOptionalTrimModeArgument,
          isCellValueError,
          isSingleCellValue,
        })
      )
  }
}

function executePlan(
  plan: readonly JsPlanInstruction[],
  context: EvaluationContext,
  initialScopes: readonly Map<string, StackValue>[] = [],
): StackValue | undefined {
  const stack: StackValue[] = []
  const scopes: Array<Map<string, StackValue>> = cloneScopes(initialScopes)
  let pc = 0

  while (pc < plan.length) {
    const instruction = plan[pc]!
    switch (instruction.opcode) {
      case 'push-number':
        stack.push({ kind: 'scalar', value: { tag: ValueTag.Number, value: instruction.value } })
        break
      case 'push-boolean':
        stack.push({ kind: 'scalar', value: { tag: ValueTag.Boolean, value: instruction.value } })
        break
      case 'push-string':
        stack.push({
          kind: 'scalar',
          value: { tag: ValueTag.String, value: instruction.value, stringId: 0 },
        })
        break
      case 'push-error':
        stack.push({ kind: 'scalar', value: error(instruction.code) })
        break
      case 'push-name':
        {
          let scopedValue: StackValue | undefined
          for (let index = scopes.length - 1; index >= 0; index -= 1) {
            const found = scopes[index]!.get(normalizeScopeName(instruction.name))
            if (found) {
              scopedValue = found
              break
            }
          }
          stack.push(
            scopedValue
              ? cloneStackValue(scopedValue)
              : {
                  kind: 'scalar',
                  value: context.resolveName?.(instruction.name) ?? error(ErrorCode.Name),
                },
          )
        }
        break
      case 'push-cell':
        stack.push({
          kind: 'scalar',
          value: context.resolveCell(instruction.sheetName ?? context.sheetName, instruction.address),
        })
        break
      case 'push-range':
        {
          const values = context.resolveRange(
            instruction.sheetName ?? context.sheetName,
            instruction.start,
            instruction.end,
            instruction.refKind,
          )
          let rows = values.length
          let cols = 1
          if (instruction.refKind === 'cells') {
            try {
              const sheetPrefix = instruction.sheetName ? `${instruction.sheetName}!` : ''
              const range = parseRangeAddress(`${sheetPrefix}${instruction.start}:${instruction.end}`)
              if (range.kind === 'cells') {
                rows = range.end.row - range.start.row + 1
                cols = range.end.col - range.start.col + 1
              }
            } catch {
              rows = values.length
              cols = 1
            }
          }
          context.noteRangeMaterialization?.(values.length)
          stack.push({
            kind: 'range',
            values,
            refKind: instruction.refKind,
            rows,
            cols,
            sheetName: instruction.sheetName ?? context.sheetName,
            start: instruction.start,
            end: instruction.end,
          })
        }
        break
      case 'lookup-exact-match': {
        const lookupOperand = popArgument(stack)
        const lookupValue = isSingleCellValue(lookupOperand)
        if (!lookupValue) {
          stack.push({ kind: 'scalar', value: error(ErrorCode.Value) })
          break
        }

        const sheetName = instruction.sheetName ?? context.sheetName
        const directMatch = context.resolveExactVectorMatch?.({
          lookupValue,
          sheetName,
          start: instruction.start,
          end: instruction.end,
          startRow: instruction.startRow,
          endRow: instruction.endRow,
          startCol: instruction.startCol,
          endCol: instruction.endCol,
          searchMode: instruction.searchMode,
        })
        if (directMatch?.handled) {
          context.noteExactLookupDirect?.()
          stack.push({
            kind: 'scalar',
            value: directMatch.position === undefined ? error(ErrorCode.NA) : { tag: ValueTag.Number, value: directMatch.position },
          })
          break
        }

        context.noteExactLookupFallback?.()
        const values = context.resolveRange(sheetName, instruction.start, instruction.end, instruction.refKind)
        context.noteRangeMaterialization?.(values.length)
        let rows = values.length
        let cols = 1
        rows = instruction.endRow - instruction.startRow + 1
        cols = instruction.endCol - instruction.startCol + 1

        const rangeArg: RangeBuiltinArgument = {
          kind: 'range',
          values,
          refKind: instruction.refKind,
          rows,
          cols,
          sheetName,
          start: instruction.start,
          end: instruction.end,
        }
        const lookupBuiltin = context.resolveLookupBuiltin?.(instruction.callee) ?? getLookupBuiltin(instruction.callee)
        if (!lookupBuiltin) {
          stack.push({ kind: 'scalar', value: error(ErrorCode.Name) })
          break
        }

        const result =
          instruction.callee === 'MATCH'
            ? lookupBuiltin(lookupValue, rangeArg, { tag: ValueTag.Number, value: 0 })
            : lookupBuiltin(
                lookupValue,
                rangeArg,
                { tag: ValueTag.Number, value: 0 },
                { tag: ValueTag.Number, value: instruction.searchMode },
              )
        stack.push(isArrayValue(result) ? result : { kind: 'scalar', value: result })
        break
      }
      case 'lookup-approximate-match': {
        const lookupOperand = popArgument(stack)
        const lookupValue = isSingleCellValue(lookupOperand)
        if (!lookupValue) {
          stack.push({ kind: 'scalar', value: error(ErrorCode.Value) })
          break
        }

        const sheetName = instruction.sheetName ?? context.sheetName
        const directMatch = context.resolveApproximateVectorMatch?.({
          lookupValue,
          sheetName,
          start: instruction.start,
          end: instruction.end,
          startRow: instruction.startRow,
          endRow: instruction.endRow,
          startCol: instruction.startCol,
          endCol: instruction.endCol,
          matchMode: instruction.matchMode,
        })
        if (directMatch?.handled) {
          stack.push({
            kind: 'scalar',
            value: directMatch.position === undefined ? error(ErrorCode.NA) : { tag: ValueTag.Number, value: directMatch.position },
          })
          break
        }

        const values = context.resolveRange(sheetName, instruction.start, instruction.end, instruction.refKind)
        context.noteRangeMaterialization?.(values.length)
        const rows = instruction.endRow - instruction.startRow + 1
        const cols = instruction.endCol - instruction.startCol + 1

        const rangeArg: RangeBuiltinArgument = {
          kind: 'range',
          values,
          refKind: instruction.refKind,
          rows,
          cols,
          sheetName,
          start: instruction.start,
          end: instruction.end,
        }
        const lookupBuiltin = context.resolveLookupBuiltin?.(instruction.callee) ?? getLookupBuiltin(instruction.callee)
        if (!lookupBuiltin) {
          stack.push({ kind: 'scalar', value: error(ErrorCode.Name) })
          break
        }

        const matchModeValue = { tag: ValueTag.Number, value: instruction.matchMode } as const
        const result =
          instruction.callee === 'MATCH'
            ? lookupBuiltin(lookupValue, rangeArg, matchModeValue)
            : lookupBuiltin(lookupValue, rangeArg, matchModeValue)
        stack.push(isArrayValue(result) ? result : { kind: 'scalar', value: result })
        break
      }
      case 'push-lambda':
        stack.push({
          kind: 'lambda',
          params: [...instruction.params],
          body: instruction.body,
          scopes: cloneScopes(scopes),
        })
        break
      case 'unary': {
        const value = popScalar(stack)
        const numeric = toNumber(value)
        stack.push({
          kind: 'scalar',
          value:
            numeric === undefined
              ? error(ErrorCode.Value)
              : { tag: ValueTag.Number, value: instruction.operator === '-' ? -numeric : numeric },
        })
        break
      }
      case 'binary': {
        const right = popArgument(stack)
        const left = popArgument(stack)
        const result = evaluateBinary(instruction.operator, left, right)
        stack.push(isArrayValue(result) ? result : { kind: 'scalar', value: result })
        break
      }
      case 'begin-scope':
        scopes.push(new Map())
        break
      case 'bind-name': {
        const scope = scopes[scopes.length - 1]
        if (!scope) {
          stack.push({ kind: 'scalar', value: error(ErrorCode.Value) })
          break
        }
        scope.set(normalizeScopeName(instruction.name), cloneStackValue(popArgument(stack)))
        break
      }
      case 'end-scope':
        scopes.pop()
        break
      case 'call': {
        const rawArgs: StackValue[] = []
        for (let index = 0; index < instruction.argc; index += 1) {
          rawArgs.unshift(popArgument(stack))
        }
        const specialResult = evaluateSpecialCall(instruction.callee, rawArgs, context, instruction.argRefs)
        if (specialResult) {
          stack.push(specialResult)
          break
        }
        const lookupBuiltin = context.resolveLookupBuiltin?.(instruction.callee) ?? getLookupBuiltin(instruction.callee)
        if (lookupBuiltin) {
          const args: Array<CellValue | RangeBuiltinArgument> = []
          for (const rawArg of rawArgs) {
            args.push(toRangeArgument(rawArg))
          }
          const result = lookupBuiltin(...args)
          stack.push(isArrayValue(result) ? result : { kind: 'scalar', value: result })
          break
        }

        const builtin = context.resolveBuiltin?.(instruction.callee) ?? getBuiltin(instruction.callee)
        if (!builtin) {
          stack.push({ kind: 'scalar', value: error(ErrorCode.Name) })
          break
        }
        const args: CellValue[] = []
        for (const rawArg of rawArgs) {
          if (rawArg.kind === 'scalar') {
            args.push(rawArg.value)
            continue
          }
          if (rawArg.kind === 'omitted') {
            args.push(error(ErrorCode.Value))
            continue
          }
          if (rawArg.kind === 'lambda') {
            args.push(error(ErrorCode.Value))
            continue
          }
          args.push(...rawArg.values)
        }
        const result = builtin(...args)
        stack.push(isArrayValue(result) ? result : { kind: 'scalar', value: result })
        break
      }
      case 'invoke': {
        const args: StackValue[] = []
        for (let index = 0; index < instruction.argc; index += 1) {
          args.unshift(popArgument(stack))
        }
        const callee = popArgument(stack)
        stack.push(applyLambda(callee, args, context))
        break
      }
      case 'jump-if-false': {
        const value = popScalar(stack)
        if (value.tag === ValueTag.Error) {
          return stackScalar(value)
        }
        if (!truthy(value)) {
          pc = instruction.target
          continue
        }
        break
      }
      case 'jump':
        pc = instruction.target
        continue
      case 'return':
        return stack.pop()
    }
    pc += 1
  }

  return stack.pop()
}

export function evaluatePlanResult(plan: readonly JsPlanInstruction[], context: EvaluationContext): EvaluationResult {
  return toEvaluationResult(executePlan(plan, context))
}

export function evaluatePlan(plan: readonly JsPlanInstruction[], context: EvaluationContext): CellValue {
  return scalarFromEvaluationResult(evaluatePlanResult(plan, context))
}

export function evaluateAst(node: FormulaNode, context: EvaluationContext): CellValue {
  return evaluatePlan(lowerToPlan(node), context)
}

export function evaluateAstResult(node: FormulaNode, context: EvaluationContext): EvaluationResult {
  return evaluatePlanResult(lowerToPlan(node), context)
}
