import { ValueTag, type CellValue, type ErrorCode, type LiteralInput } from '@bilig/protocol'
import { isArrayValue, type EvaluationResult, type FormulaNode } from '@bilig/formula'
import { WorkPaperInvalidArgumentsError } from './work-paper-errors.js'
import type { RawCellContent, WorkPaperAddressLike, WorkPaperCellRange, WorkPaperSheet } from './work-paper-types.js'

export function emptyValue(): CellValue {
  return { tag: ValueTag.Empty }
}

export function errorValue(code: ErrorCode): CellValue {
  return { tag: ValueTag.Error, code }
}

export function scalarValueFromLiteral(value: LiteralInput): CellValue {
  if (value === null) {
    return emptyValue()
  }
  if (typeof value === 'number') {
    return { tag: ValueTag.Number, value }
  }
  if (typeof value === 'boolean') {
    return { tag: ValueTag.Boolean, value }
  }
  return { tag: ValueTag.String, value, stringId: 0 }
}

export function scalarFromResult(result: EvaluationResult | CellValue): CellValue {
  if (!isArrayValue(result)) {
    return result
  }
  return result.values[0] ?? emptyValue()
}

const SIMPLE_NUMERIC_FORMULA_RE = /^[+-]?(?:(?:\d+(?:\.\d*)?)|(?:\.\d+))(?:[Ee][+-]?\d+)?$/

export function tryReadSimpleScalarFormulaBody(raw: RawCellContent): string | undefined {
  if (typeof raw !== 'string') {
    return undefined
  }
  const trimmed = raw.trim()
  if (!trimmed.startsWith('=')) {
    return undefined
  }
  const body = trimmed.slice(1).trim()
  return body.length === 0 ? undefined : body
}

export function tryEvaluateSimpleScalarFormulaBody(body: string): CellValue | undefined {
  if (SIMPLE_NUMERIC_FORMULA_RE.test(body)) {
    const value = Number(body)
    return Number.isFinite(value) ? { tag: ValueTag.Number, value: Object.is(value, -0) ? 0 : value } : undefined
  }
  const upper = body.toUpperCase()
  if (upper === 'TRUE') {
    return { tag: ValueTag.Boolean, value: true }
  }
  if (upper === 'FALSE') {
    return { tag: ValueTag.Boolean, value: false }
  }
  if (body.length >= 2 && body.startsWith('"') && body.endsWith('"')) {
    return { tag: ValueTag.String, value: body.slice(1, -1).replaceAll('""', '"'), stringId: 0 }
  }
  return undefined
}

export function tryEvaluateSimpleNamedExpression(raw: RawCellContent): CellValue | undefined {
  if (raw === null || typeof raw === 'number' || typeof raw === 'boolean') {
    return scalarValueFromLiteral(raw)
  }
  if (typeof raw === 'string' && !raw.trim().startsWith('=')) {
    return scalarValueFromLiteral(raw)
  }
  const body = tryReadSimpleScalarFormulaBody(raw)
  return body === undefined ? undefined : tryEvaluateSimpleScalarFormulaBody(body)
}

export function valuesEqual(left: CellValue, right: CellValue): boolean {
  if (left.tag !== right.tag) {
    return false
  }
  switch (left.tag) {
    case ValueTag.Number:
      return right.tag === ValueTag.Number && left.value === right.value
    case ValueTag.Boolean:
      return right.tag === ValueTag.Boolean && left.value === right.value
    case ValueTag.String:
      return right.tag === ValueTag.String && left.value === right.value
    case ValueTag.Error:
      return right.tag === ValueTag.Error && left.code === right.code
    case ValueTag.Empty:
      return true
    default:
      return false
  }
}

function isCellValueMatrix(value: CellValue | CellValue[][]): value is CellValue[][] {
  return Array.isArray(value)
}

export function matrixValuesEqual(left: CellValue | CellValue[][] | undefined, right: CellValue | CellValue[][] | undefined): boolean {
  if (!left && !right) {
    return true
  }
  if (!left || !right) {
    return false
  }
  if (isCellValueMatrix(left) !== isCellValueMatrix(right)) {
    return false
  }
  if (!isCellValueMatrix(left) && !isCellValueMatrix(right)) {
    return valuesEqual(left, right)
  }
  if (!isCellValueMatrix(left) || !isCellValueMatrix(right)) {
    return false
  }
  const leftMatrix = left
  const rightMatrix = right
  if (leftMatrix.length !== rightMatrix.length) {
    return false
  }
  return leftMatrix.every((row: CellValue[], rowIndex: number) => {
    const otherRow = rightMatrix[rowIndex]
    if (!otherRow || row.length !== otherRow.length) {
      return false
    }
    return row.every((value: CellValue, columnIndex: number) => {
      const otherValue = otherRow[columnIndex]
      if (!otherValue) {
        return false
      }
      return valuesEqual(value, otherValue)
    })
  })
}

export function normalizeName(name: string): string {
  return name.trim().toUpperCase()
}

export function makeNamedExpressionKey(name: string, scope?: number): string {
  return `${scope ?? 'workbook'}:${normalizeName(name)}`
}

export function makeInternalScopedName(scope: number, name: string): string {
  return `__BILIG_WORKPAPER_SCOPE_${scope}_${normalizeName(name)}`
}

export function isFormulaContent(content: RawCellContent): content is string {
  return typeof content === 'string' && content.trim().startsWith('=')
}

export function isBlankRawCellContent(content: RawCellContent | undefined): content is null | undefined {
  return content === null || content === undefined
}

export function isWorkPaperSheetMatrix(value: RawCellContent | WorkPaperSheet): value is WorkPaperSheet {
  return Array.isArray(value)
}

export function matrixContainsFormulaContent(content: WorkPaperSheet): boolean {
  return content.some((row) => row.some((cell) => isFormulaContent(cell)))
}

export function isDeferredBatchLiteralContent(content: RawCellContent): boolean {
  return content === null || typeof content === 'boolean' || typeof content === 'number' || typeof content === 'string'
}

export function stripLeadingEquals(formula: string): string {
  return formula.trim().startsWith('=') ? formula.trim().slice(1) : formula.trim()
}

export function assertRowAndColumn(value: number, label: string): void {
  if (!Number.isInteger(value) || value < 0) {
    throw new WorkPaperInvalidArgumentsError(`${label} to be a non-negative integer`)
  }
}

export function assertRange(range: WorkPaperCellRange): void {
  assertRowAndColumn(range.start.sheet, 'start.sheet')
  assertRowAndColumn(range.start.row, 'start.row')
  assertRowAndColumn(range.start.col, 'start.col')
  assertRowAndColumn(range.end.sheet, 'end.sheet')
  assertRowAndColumn(range.end.row, 'end.row')
  assertRowAndColumn(range.end.col, 'end.col')
  if (range.start.sheet !== range.end.sheet) {
    throw new WorkPaperInvalidArgumentsError('Ranges must stay on a single sheet')
  }
}

export function isCellRange(value: WorkPaperAddressLike): value is WorkPaperCellRange {
  return 'start' in value && 'end' in value
}

export function cloneCellValue(value: CellValue): CellValue {
  switch (value.tag) {
    case ValueTag.Empty:
      return emptyValue()
    case ValueTag.Number:
      return { tag: ValueTag.Number, value: value.value }
    case ValueTag.Boolean:
      return { tag: ValueTag.Boolean, value: value.value }
    case ValueTag.String:
      return { tag: ValueTag.String, value: value.value, stringId: value.stringId }
    case ValueTag.Error:
      return { tag: ValueTag.Error, code: value.code }
    default:
      return emptyValue()
  }
}

export function transformFormulaNode(node: FormulaNode, transform: (current: FormulaNode) => FormulaNode): FormulaNode {
  const current = transform(node)
  switch (current.kind) {
    case 'BooleanLiteral':
    case 'CellRef':
    case 'ColumnRef':
    case 'ErrorLiteral':
    case 'NameRef':
    case 'NumberLiteral':
    case 'OmittedArgument':
    case 'RangeRef':
    case 'RowRef':
    case 'SpillRef':
    case 'StringLiteral':
    case 'StructuredRef':
      return current
    case 'ArrayConstant':
      return {
        ...current,
        rows: current.rows.map((row) => row.map((entry) => transformFormulaNode(entry, transform))),
      }
    case 'UnaryExpr':
      return {
        ...current,
        argument: transformFormulaNode(current.argument, transform),
      }
    case 'BinaryExpr':
      return {
        ...current,
        left: transformFormulaNode(current.left, transform),
        right: transformFormulaNode(current.right, transform),
      }
    case 'CallExpr':
      return {
        ...current,
        args: current.args.map((argument) => transformFormulaNode(argument, transform)),
      }
    case 'InvokeExpr':
      return {
        ...current,
        callee: transformFormulaNode(current.callee, transform),
        args: current.args.map((argument) => transformFormulaNode(argument, transform)),
      }
    default:
      return current
  }
}

export function collectFormulaNameRefs(node: FormulaNode, output: Set<string>): void {
  switch (node.kind) {
    case 'BooleanLiteral':
    case 'CellRef':
    case 'ColumnRef':
    case 'ErrorLiteral':
    case 'NameRef':
      if (node.kind === 'NameRef') {
        output.add(node.name)
      }
      return
    case 'NumberLiteral':
    case 'OmittedArgument':
    case 'RangeRef':
    case 'RowRef':
    case 'SpillRef':
    case 'StringLiteral':
    case 'StructuredRef':
      return
    case 'ArrayConstant':
      node.rows.forEach((row) => row.forEach((entry) => collectFormulaNameRefs(entry, output)))
      return
    case 'UnaryExpr':
      collectFormulaNameRefs(node.argument, output)
      return
    case 'BinaryExpr':
      collectFormulaNameRefs(node.left, output)
      collectFormulaNameRefs(node.right, output)
      return
    case 'CallExpr':
      node.args.forEach((argument) => collectFormulaNameRefs(argument, output))
      return
    case 'InvokeExpr':
      collectFormulaNameRefs(node.callee, output)
      node.args.forEach((argument) => collectFormulaNameRefs(argument, output))
      return
    default:
      return
  }
}

function isAbsoluteCellReference(value: string): boolean {
  return /^\$[A-Z]+\$[1-9][0-9]*$/.test(value.toUpperCase())
}

function isAbsoluteRowReference(value: string): boolean {
  return /^\$[1-9][0-9]*$/.test(value)
}

function isAbsoluteColumnReference(value: string): boolean {
  return /^\$[A-Z]+$/.test(value.toUpperCase())
}

export function formulaHasRelativeReferences(node: FormulaNode): boolean {
  switch (node.kind) {
    case 'BooleanLiteral':
    case 'ErrorLiteral':
    case 'NameRef':
    case 'NumberLiteral':
    case 'OmittedArgument':
    case 'StringLiteral':
    case 'StructuredRef':
      return false
    case 'ArrayConstant':
      return node.rows.some((row) => row.some(formulaHasRelativeReferences))
    case 'CellRef':
    case 'SpillRef':
      return !isAbsoluteCellReference(node.ref)
    case 'RowRef':
      return !isAbsoluteRowReference(node.ref)
    case 'ColumnRef':
      return !isAbsoluteColumnReference(node.ref)
    case 'RangeRef':
      if (node.refKind === 'cells') {
        return !isAbsoluteCellReference(node.start) || !isAbsoluteCellReference(node.end)
      }
      if (node.refKind === 'rows') {
        return !isAbsoluteRowReference(node.start) || !isAbsoluteRowReference(node.end)
      }
      return !isAbsoluteColumnReference(node.start) || !isAbsoluteColumnReference(node.end)
    case 'UnaryExpr':
      return formulaHasRelativeReferences(node.argument)
    case 'BinaryExpr':
      return formulaHasRelativeReferences(node.left) || formulaHasRelativeReferences(node.right)
    case 'CallExpr':
      return node.args.some((argument) => formulaHasRelativeReferences(argument))
    case 'InvokeExpr':
      return formulaHasRelativeReferences(node.callee) || node.args.some((argument) => formulaHasRelativeReferences(argument))
    default:
      return false
  }
}
