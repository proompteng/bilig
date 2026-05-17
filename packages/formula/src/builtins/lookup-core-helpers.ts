import { ErrorCode, ValueTag, formatGeneralNumberValue, type CellValue } from '@bilig/protocol'
import { parseNumericText } from '../numeric-text.js'
import type { ArrayValue, EvaluationResult } from '../runtime-values.js'

export interface RangeBuiltinArgument {
  kind: 'range'
  values: CellValue[]
  refKind: 'cells' | 'rows' | 'cols'
  rows: number
  cols: number
  sheetName?: string
  start?: string
  end?: string
}

export type LookupBuiltinArgument = CellValue | RangeBuiltinArgument | undefined
export type LookupBuiltin = (...args: LookupBuiltinArgument[]) => EvaluationResult
export interface LookupBuiltinResolverOptions {
  resolveIndexedExactMatch?: (lookupValue: CellValue, range: RangeBuiltinArgument) => number | undefined
}

const exactLookupNumberSignificantDigits = 15

export function normalizeExactLookupNumber(value: number): number {
  if (Object.is(value, -0)) {
    return 0
  }
  if (!Number.isFinite(value)) {
    return value
  }
  const normalized = Number(value.toPrecision(exactLookupNumberSignificantDigits))
  return Object.is(normalized, -0) ? 0 : normalized
}

export function exactLookupNumberKey(value: number): string {
  return `n:${normalizeExactLookupNumber(value)}`
}

export function sameExactLookupNumber(left: number, right: number): boolean {
  return normalizeExactLookupNumber(left) === normalizeExactLookupNumber(right)
}

export function errorValue(code: ErrorCode): CellValue {
  return { tag: ValueTag.Error, code }
}

export function numberResult(value: number): CellValue {
  return { tag: ValueTag.Number, value }
}

export function isError(value: LookupBuiltinArgument | undefined): value is Extract<CellValue, { tag: ValueTag.Error }> {
  return value !== undefined && !isRangeArg(value) && value.tag === ValueTag.Error
}

export function isRangeArg(value: LookupBuiltinArgument | undefined): value is RangeBuiltinArgument {
  return typeof value === 'object' && value !== null && 'kind' in value && value.kind === 'range'
}

export function findFirstNonRange(values: readonly LookupBuiltinArgument[]): CellValue | undefined {
  for (const value of values) {
    if (value !== undefined && !isRangeArg(value)) {
      return value
    }
  }
  return undefined
}

export function areRangeArgs(values: readonly LookupBuiltinArgument[]): values is RangeBuiltinArgument[] {
  return values.every((value) => isRangeArg(value))
}

export function toNumber(value: CellValue | undefined): number | undefined {
  if (value === undefined) {
    return undefined
  }
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

export function toInteger(value: CellValue | undefined): number | undefined {
  const numeric = toNumber(value)
  if (numeric === undefined || !Number.isFinite(numeric)) {
    return undefined
  }
  return Math.trunc(numeric)
}

export function toBoolean(value: CellValue | undefined): boolean | undefined {
  if (value === undefined) {
    return undefined
  }
  switch (value.tag) {
    case ValueTag.Boolean:
      return value.value
    case ValueTag.Number:
      return value.value !== 0
    case ValueTag.Empty:
      return false
    case ValueTag.String:
    case ValueTag.Error:
      return undefined
    default:
      return undefined
  }
}

export function toStringValue(value: CellValue | undefined): string {
  if (value === undefined) {
    return ''
  }
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
      return ''
  }
}

export function compareScalars(left: CellValue, right: CellValue): number | undefined {
  if ((left.tag === ValueTag.String || left.tag === ValueTag.Empty) && (right.tag === ValueTag.String || right.tag === ValueTag.Empty)) {
    const normalizedLeft = toStringValue(left).toUpperCase()
    const normalizedRight = toStringValue(right).toUpperCase()
    if (normalizedLeft === normalizedRight) {
      return 0
    }
    return normalizedLeft < normalizedRight ? -1 : 1
  }

  const leftNum = toNumber(left)
  const rightNum = toNumber(right)
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

export function requireCellVector(arg: LookupBuiltinArgument): RangeBuiltinArgument | CellValue {
  if (!isRangeArg(arg)) {
    return errorValue(ErrorCode.Value)
  }
  if (arg.rows !== 1 && arg.cols !== 1) {
    return errorValue(ErrorCode.NA)
  }
  return arg
}

export function requireCellRange(arg: LookupBuiltinArgument): RangeBuiltinArgument | CellValue {
  if (!isRangeArg(arg) || arg.refKind !== 'cells') {
    return errorValue(ErrorCode.Value)
  }
  return arg
}

export function getRangeValue(range: RangeBuiltinArgument, row: number, col: number): CellValue {
  const index = row * range.cols + col
  return range.values[index] ?? { tag: ValueTag.Empty }
}

export function arrayResult(values: CellValue[], rows: number, cols: number): ArrayValue {
  return { kind: 'array', values, rows, cols }
}

export function collectNumericSeries(arg: LookupBuiltinArgument, mode: 'lenient' | 'strict'): number[] | CellValue {
  if (arg === undefined) {
    return errorValue(ErrorCode.Value)
  }
  const values: number[] = []
  const cells = isRangeArg(arg) ? arg.values : [arg]
  if (isRangeArg(arg) && arg.refKind !== 'cells') {
    return errorValue(ErrorCode.Value)
  }
  for (const cell of cells) {
    if (cell.tag === ValueTag.Error) {
      return cell
    }
    if (cell.tag === ValueTag.Number) {
      values.push(cell.value)
      continue
    }
    if (mode === 'strict') {
      return errorValue(ErrorCode.Value)
    }
  }
  return values
}

export function numericAggregateCandidate(value: CellValue | undefined): number | undefined {
  return value?.tag === ValueTag.Number ? value.value : undefined
}

export function toCellRange(arg: LookupBuiltinArgument): RangeBuiltinArgument | CellValue {
  if (arg === undefined) {
    return errorValue(ErrorCode.Value)
  }
  if (!isRangeArg(arg)) {
    return { kind: 'range', values: [arg], refKind: 'cells', rows: 1, cols: 1 }
  }
  if (arg.refKind !== 'cells') {
    return errorValue(ErrorCode.Value)
  }
  return arg
}

export function toNumericMatrix(arg: LookupBuiltinArgument): number[][] | CellValue {
  const range = toCellRange(arg)
  if (!isRangeArg(range)) {
    return range
  }
  const matrix: number[][] = []
  for (let row = 0; row < range.rows; row += 1) {
    const rowValues: number[] = []
    for (let col = 0; col < range.cols; col += 1) {
      const numeric = toNumber(getRangeValue(range, row, col))
      if (numeric === undefined) {
        return errorValue(ErrorCode.Value)
      }
      rowValues.push(numeric)
    }
    matrix.push(rowValues)
  }
  return matrix
}

export function flattenNumbers(arg: LookupBuiltinArgument): number[] | CellValue {
  if (!isRangeArg(arg)) {
    const numeric = toNumber(arg)
    return numeric === undefined ? errorValue(ErrorCode.Value) : [numeric]
  }
  const values: number[] = []
  for (const value of arg.values) {
    const numeric = toNumber(value)
    if (numeric === undefined) {
      return errorValue(ErrorCode.Value)
    }
    values.push(numeric)
  }
  return values
}

export function pickRangeRow(range: RangeBuiltinArgument, row: number): CellValue[] {
  const values: CellValue[] = []
  for (let col = 0; col < range.cols; col += 1) {
    values.push(getRangeValue(range, row, col))
  }
  return values
}

export function firstLookupError(args: readonly LookupBuiltinArgument[]): CellValue | undefined {
  return args.find((arg) => isError(arg))
}

export type CriteriaOperator = '=' | '<>' | '>' | '>=' | '<' | '<='

export interface CompiledCriteriaMatcher {
  readonly operator: CriteriaOperator
  readonly operand: CellValue
  readonly wildcardPattern?: RegExp
}

export function matchesCompiledCriteria(value: CellValue, compiled: CompiledCriteriaMatcher): boolean {
  if (isError(value)) {
    return false
  }
  if (compiled.wildcardPattern) {
    const matches = value.tag === ValueTag.String && compiled.wildcardPattern.test(toStringValue(value))
    return compiled.operator === '=' ? matches : !matches
  }
  if (value.tag === ValueTag.Empty && compiled.operand.tag === ValueTag.Number && compiled.operator !== '=' && compiled.operator !== '<>') {
    return false
  }
  const comparison = compareScalars(value, compiled.operand)
  if (comparison === undefined) {
    return false
  }
  switch (compiled.operator) {
    case '=':
      return comparison === 0
    case '<>':
      return comparison !== 0
    case '>':
      return comparison > 0
    case '>=':
      return comparison >= 0
    case '<':
      return comparison < 0
    case '<=':
      return comparison <= 0
  }
}

export function compileCriteriaMatcher(criteria: CellValue): CompiledCriteriaMatcher {
  let { operator, operand } = parseCriteria(criteria)
  let wildcardPattern: RegExp | undefined
  if (operand.tag === ValueTag.String && (operator === '=' || operator === '<>') && hasWildcardPattern(operand.value)) {
    wildcardPattern = wildcardPatternToRegExp(operand.value)
  } else if (operand.tag === ValueTag.String && operand.value.includes('~')) {
    operand = {
      tag: ValueTag.String,
      value: unescapeCriteriaPattern(operand.value),
      stringId: operand.stringId,
    }
  }
  return {
    operator,
    operand,
    ...(wildcardPattern ? { wildcardPattern } : {}),
  }
}

export function matchesCriteria(value: CellValue, criteria: CellValue): boolean {
  return matchesCompiledCriteria(value, compileCriteriaMatcher(criteria))
}

function isCriteriaOperator(value: string): value is CriteriaOperator {
  return value === '=' || value === '<>' || value === '>' || value === '>=' || value === '<' || value === '<='
}

function hasWildcardPattern(pattern: string): boolean {
  for (let index = 0; index < pattern.length; index += 1) {
    const char = pattern[index]
    if (char === '~') {
      index += 1
      continue
    }
    if (char === '*' || char === '?') {
      return true
    }
  }
  return false
}

function wildcardPatternToRegExp(pattern: string): RegExp {
  let source = '^'
  for (let index = 0; index < pattern.length; index += 1) {
    const char = pattern[index]
    if (char === undefined) {
      continue
    }
    if (char === '~') {
      const escaped = pattern[index + 1]
      if (escaped !== undefined) {
        source += escapeRegexFragment(escaped)
        index += 1
        continue
      }
      source += escapeRegexFragment(char)
      continue
    }
    if (char === '*') {
      source += '.*'
      continue
    }
    if (char === '?') {
      source += '.'
      continue
    }
    source += escapeRegexFragment(char)
  }
  source += '$'
  return new RegExp(source, 'i')
}

function escapeRegexFragment(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
}

function unescapeCriteriaPattern(pattern: string): string {
  let unescaped = ''
  for (let index = 0; index < pattern.length; index += 1) {
    const char = pattern[index]
    if (char === undefined) {
      continue
    }
    if (char === '~') {
      const escaped = pattern[index + 1]
      if (escaped !== undefined) {
        unescaped += escaped
        index += 1
        continue
      }
    }
    unescaped += char
  }
  return unescaped
}

function parseCriteria(criteria: CellValue): { operator: CriteriaOperator; operand: CellValue } {
  if (criteria.tag !== ValueTag.String) {
    return { operator: '=', operand: criteria }
  }

  const match = /^(<=|>=|<>|=|<|>)(.*)$/.exec(criteria.value)
  if (!match) {
    return { operator: '=', operand: criteria }
  }

  const operator = match[1] ?? '='
  return {
    operator: isCriteriaOperator(operator) ? operator : '=',
    operand: parseCriteriaOperand(match[2] ?? ''),
  }
}

function parseCriteriaOperand(raw: string): CellValue {
  const trimmed = raw.trim()
  if (trimmed === '') {
    return { tag: ValueTag.String, value: '', stringId: 0 }
  }
  const upper = trimmed.toUpperCase()
  if (upper === 'TRUE' || upper === 'FALSE') {
    return { tag: ValueTag.Boolean, value: upper === 'TRUE' }
  }
  const numeric = parseCriteriaNumericOperand(trimmed)
  if (numeric !== undefined) {
    return { tag: ValueTag.Number, value: numeric }
  }
  return { tag: ValueTag.String, value: trimmed, stringId: 0 }
}

function parseCriteriaNumericOperand(trimmed: string): number | undefined {
  return parseNumericText(trimmed)
}
