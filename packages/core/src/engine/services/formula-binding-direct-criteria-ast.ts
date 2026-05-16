import { type FormulaNode, parseRangeAddress } from '@bilig/formula'
import { MAX_ROWS, ValueTag, type CellValue } from '@bilig/protocol'
import type { RuntimeDirectCriteriaDescriptor, RuntimeDirectCriteriaResultTransform } from '../runtime-state.js'

export interface DirectCriteriaResolvedRange {
  readonly sheetName: string
  readonly rowStart: number
  readonly rowEnd: number
  readonly col: number
  readonly length: number
}

export function staticCellValue(node: FormulaNode | undefined): CellValue | undefined {
  if (!node) {
    return undefined
  }
  switch (node.kind) {
    case 'BooleanLiteral':
      return { tag: ValueTag.Boolean, value: node.value }
    case 'ErrorLiteral':
      return { tag: ValueTag.Error, code: node.code }
    case 'NumberLiteral':
      return { tag: ValueTag.Number, value: node.value }
    case 'StringLiteral':
      return { tag: ValueTag.String, value: node.value, stringId: 0 }
    case 'UnaryExpr':
      if (node.operator === '-' && node.argument.kind === 'NumberLiteral') {
        return { tag: ValueTag.Number, value: -node.argument.value }
      }
      return undefined
    case 'ArrayConstant':
      return undefined
    case 'CellRef':
    case 'CallExpr':
    case 'BinaryExpr':
    case 'ColumnRef':
    case 'InvokeExpr':
    case 'NameRef':
    case 'OmittedArgument':
    case 'RangeRef':
    case 'RowRef':
    case 'SpillRef':
    case 'StructuredRef':
      return undefined
  }
}

export function flattenCriteriaProduct(node: FormulaNode): FormulaNode[] {
  if (node.kind !== 'BinaryExpr' || node.operator !== '*') {
    return [node]
  }
  return [...flattenCriteriaProduct(node.left), ...flattenCriteriaProduct(node.right)]
}

export function resolveDirectCriteriaRange(node: FormulaNode | undefined, ownerSheetName: string): DirectCriteriaResolvedRange | undefined {
  if (!node || node.kind !== 'RangeRef' || node.sheetEndName !== undefined) {
    return undefined
  }
  const parsed = parseRangeAddress(`${node.start}:${node.end}`, node.sheetName ?? ownerSheetName)
  const sheetName = parsed.sheetName ?? node.sheetName ?? ownerSheetName
  if (parsed.kind === 'cols') {
    if (parsed.start.col !== parsed.end.col) {
      return undefined
    }
    return {
      sheetName,
      rowStart: 0,
      rowEnd: MAX_ROWS - 1,
      col: parsed.start.col,
      length: MAX_ROWS,
    }
  }
  if (parsed.kind !== 'cells' || parsed.start.col !== parsed.end.col) {
    return undefined
  }
  return {
    sheetName,
    rowStart: parsed.start.row,
    rowEnd: parsed.end.row,
    col: parsed.start.col,
    length: parsed.end.row - parsed.start.row + 1,
  }
}

export function callName(node: FormulaNode | undefined): string | undefined {
  return node?.kind === 'CallExpr' ? node.callee.trim().toUpperCase() : undefined
}

export function appendDirectCriteriaResultTransform(
  descriptor: RuntimeDirectCriteriaDescriptor,
  transform: RuntimeDirectCriteriaResultTransform,
): RuntimeDirectCriteriaDescriptor {
  return {
    ...descriptor,
    resultTransforms: [...(descriptor.resultTransforms ?? []), transform],
  }
}
