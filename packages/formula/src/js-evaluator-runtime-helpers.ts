import { ErrorCode, ValueTag, formatErrorCode, formatGeneralNumberValue, type CellValue } from '@bilig/protocol'
import type { RangeBuiltinArgument } from './builtins/lookup.js'
import { normalizeExactLookupNumber } from './builtins/lookup-core-helpers.js'
import type { MatrixValue } from './group-pivot-evaluator.js'
import { excelPower } from './excel-power.js'
import { emptyValue, error, numberValue } from './js-evaluator-cell-values.js'
import type { JsPlanInstruction, StackValue } from './js-evaluator-types.js'
import { parseNumericText } from './numeric-text.js'
import type { EvaluationResult, RangeLikeValue } from './runtime-values.js'

type BinaryOperator = Extract<JsPlanInstruction, { opcode: 'binary' }>['operator']

export function toNumber(value: CellValue): number | undefined {
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

export function toArithmeticNumber(value: CellValue): number | undefined {
  if (value.tag === ValueTag.String) {
    const trimmed = value.value.trim()
    if (trimmed === '') {
      return 0
    }
    return parseNumericText(trimmed)
  }
  return toNumber(value)
}

export function toStringValue(value: CellValue): string {
  switch (value.tag) {
    case ValueTag.Empty:
      return ''
    case ValueTag.Number:
      return formatGeneralNumberValue(value.value)
    case ValueTag.Boolean:
      return value.value ? 'TRUE' : 'FALSE'
    case ValueTag.String:
      return value.value
    case ValueTag.Error:
      return formatErrorCode(value.code)
  }
}

function isNumberLike(value: CellValue): boolean {
  return value.tag === ValueTag.Number || value.tag === ValueTag.Boolean
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
  if (left.tag === ValueTag.String && right.tag === ValueTag.String) {
    return compareText(left.value, right.value)
  }
  if (left.tag === ValueTag.Empty && right.tag === ValueTag.Empty) {
    return 0
  }
  if (left.tag === ValueTag.String && right.tag === ValueTag.Empty) {
    return compareText(left.value, '')
  }
  if (left.tag === ValueTag.Empty && right.tag === ValueTag.String) {
    return compareText('', right.value)
  }
  if (left.tag === ValueTag.String && isNumberLike(right)) {
    return 1
  }
  if (isNumberLike(left) && right.tag === ValueTag.String) {
    return -1
  }

  const leftNum = comparableNumber(left)
  const rightNum = comparableNumber(right)
  if (leftNum === undefined || rightNum === undefined) {
    return undefined
  }
  const normalizedLeft = normalizeExactLookupNumber(leftNum)
  const normalizedRight = normalizeExactLookupNumber(rightNum)
  if (normalizedLeft === normalizedRight) {
    return 0
  }
  return normalizedLeft < normalizedRight ? -1 : 1
}

function comparableNumber(value: CellValue): number | undefined {
  switch (value.tag) {
    case ValueTag.Number:
      return value.value
    case ValueTag.Boolean:
      return value.value ? 1 : 0
    case ValueTag.Empty:
      return 0
    case ValueTag.String:
    case ValueTag.Error:
    default:
      return undefined
  }
}

export function truthy(value: CellValue): boolean {
  return (toNumber(value) ?? 0) !== 0
}

export function popScalar(stack: StackValue[]): CellValue {
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

export function popArgument(stack: StackValue[]): StackValue {
  return stack.pop() ?? { kind: 'scalar', value: error(ErrorCode.Value) }
}

export function toEvaluationResult(value: StackValue | undefined): EvaluationResult {
  if (!value) {
    return error(ErrorCode.Value)
  }
  if (value.kind === 'scalar') {
    if (value.blankReference === true && value.value.tag === ValueTag.Empty) {
      return numberValue(0)
    }
    return value.value
  }
  if (value.kind === 'omitted') {
    return error(ErrorCode.Value)
  }
  if (value.kind === 'lambda') {
    return error(ErrorCode.Value)
  }
  if (value.kind === 'range') {
    if (value.blankReference === true && value.rows === 1 && value.cols === 1 && value.values[0]?.tag === ValueTag.Empty) {
      return numberValue(0)
    }
    return {
      kind: 'array',
      rows: value.rows,
      cols: value.cols,
      values: value.values,
    }
  }
  return value
}

export function cloneStackValue(value: StackValue): StackValue {
  if (value.kind === 'scalar') {
    return value.blankReference === true
      ? { kind: 'scalar', value: value.value, blankReference: true }
      : { kind: 'scalar', value: value.value }
  }
  if (value.kind === 'omitted') {
    return value.source === undefined ? { kind: 'omitted' } : { kind: 'omitted', source: value.source }
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
      ...(value.blankReference === true ? { blankReference: true } : {}),
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

export function cloneScopes(scopes: readonly Map<string, StackValue>[]): Array<Map<string, StackValue>> {
  return scopes.map((scope) => new Map([...scope.entries()].map(([name, value]) => [name, cloneStackValue(value)])))
}

export function toRangeLike(value: StackValue): RangeLikeValue {
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
    const left = toArithmeticNumber(leftValue)
    const right = toArithmeticNumber(rightValue)
    if (left === undefined || right === undefined) {
      return error(ErrorCode.Value)
    }
    if (operator === '/' && right === 0) {
      return error(ErrorCode.Div0)
    }
    if (operator === '^') {
      const value = excelPower(left, right)
      return Number.isFinite(value) ? { tag: ValueTag.Number, value } : error(ErrorCode.Value)
    }
    const value = operator === '+' ? left + right : operator === '-' ? left - right : operator === '*' ? left * right : left / right
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

export function evaluateBinary(operator: BinaryOperator, leftValue: StackValue, rightValue: StackValue): EvaluationResult {
  if (operator === ':') {
    return error(ErrorCode.Value)
  }

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

export function stackScalar(value: CellValue, blankReference = false): StackValue {
  return blankReference ? { kind: 'scalar', value, blankReference: true } : { kind: 'scalar', value }
}

export function normalizeScopeName(name: string): string {
  return name.toUpperCase()
}

export function isSingleCellValue(value: StackValue): CellValue | undefined {
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

export function toRangeArgument(value: StackValue): CellValue | RangeBuiltinArgument {
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

export function toPositiveInteger(value: StackValue | undefined): number | undefined {
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

export function getRangeCell(range: RangeLikeValue, row: number, col: number): CellValue {
  return range.values[row * range.cols + col] ?? emptyValue()
}

export function getBroadcastShape(values: readonly StackValue[]): { rows: number; cols: number } | undefined {
  const ranges = values.map(toRangeLike)
  const rows = Math.max(...ranges.map((range) => range.rows))
  const cols = Math.max(...ranges.map((range) => range.cols))
  const compatible = ranges.every((range) => (range.rows === rows || range.rows === 1) && (range.cols === cols || range.cols === 1))
  return compatible ? { rows, cols } : undefined
}

export function coerceScalarTextArgument(value: StackValue | undefined): string | CellValue {
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

export function coerceOptionalBooleanArgument(value: StackValue | undefined, fallback: boolean): boolean | CellValue {
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

export function coerceOptionalMatchModeArgument(value: StackValue | undefined, fallback: 0 | 1): 0 | 1 | CellValue {
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

export function coerceOptionalPositiveIntegerArgument(value: StackValue | undefined, fallback: number): number | CellValue {
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

export function coerceOptionalTrimModeArgument(value: StackValue | undefined, fallback: 0 | 1 | 2 | 3): 0 | 1 | 2 | 3 | CellValue {
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

export function isCellValueError(value: number | boolean | string | CellValue): value is CellValue {
  return typeof value === 'object' && value !== null && 'tag' in value
}

export function makeArrayStack(rows: number, cols: number, values: CellValue[]): StackValue {
  return { kind: 'array', rows, cols, values }
}

export function matrixFromStackValue(value: StackValue): MatrixValue | undefined {
  if (value.kind === 'omitted' || value.kind === 'lambda') {
    return undefined
  }
  if (value.kind === 'scalar') {
    return { rows: 1, cols: 1, values: [value.value] }
  }
  return { rows: value.rows, cols: value.cols, values: value.values }
}

export function scalarIntegerArgument(value: StackValue | undefined): number | undefined {
  const scalar = value ? isSingleCellValue(value) : undefined
  const numeric = scalar ? toNumber(scalar) : undefined
  return numeric === undefined || !Number.isFinite(numeric) ? undefined : Math.trunc(numeric)
}

export function vectorIntegerArgument(value: StackValue | undefined): number[] | undefined {
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
