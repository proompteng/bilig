import { ErrorCode } from '@bilig/protocol'
import type { BinaryExprNode, FormulaNode } from './ast.js'
import { formatSheetPrefix, quoteSheetNameIfNeeded } from './translation-reference-utils.js'

const BINARY_PRECEDENCE: Record<BinaryExprNode['operator'], number> = {
  '=': 1,
  '<>': 1,
  '>': 1,
  '>=': 1,
  '<': 1,
  '<=': 1,
  '&': 2,
  '+': 3,
  '-': 3,
  '*': 4,
  '/': 4,
  '^': 5,
}

const ERROR_LITERAL_TEXT: Record<number, string> = {
  [ErrorCode.Ref]: '#REF!',
  [ErrorCode.Name]: '#NAME?',
  [ErrorCode.Div0]: '#DIV/0!',
  [ErrorCode.NA]: '#N/A',
  [ErrorCode.Value]: '#VALUE!',
  [ErrorCode.Cycle]: '#CYCLE!',
  [ErrorCode.Spill]: '#SPILL!',
  [ErrorCode.Blocked]: '#BLOCKED!',
}

function formatRangeSheetPrefix(sheetName: string | undefined, sheetEndName: string | undefined): string {
  if (sheetEndName === undefined) {
    return formatSheetPrefix(sheetName)
  }
  if (sheetName === undefined) {
    return ''
  }
  return `${quoteSheetNameIfNeeded(sheetName)}:${quoteSheetNameIfNeeded(sheetEndName)}!`
}

export function serializeFormula(node: FormulaNode, parentPrecedence = 0, parentAssociativity: 'left' | 'right' | null = null): string {
  switch (node.kind) {
    case 'NumberLiteral':
      return String(node.value)
    case 'BooleanLiteral':
      return node.value ? 'TRUE' : 'FALSE'
    case 'StringLiteral':
      return `"${node.value.replaceAll('"', '""')}"`
    case 'ErrorLiteral':
      return ERROR_LITERAL_TEXT[node.code] ?? '#ERROR!'
    case 'OmittedArgument':
      return ''
    case 'ArrayConstant':
      return `{${node.rows.map((row) => row.map((entry) => serializeFormula(entry)).join(',')).join(';')}}`
    case 'NameRef':
      return node.name
    case 'StructuredRef':
      return `${node.tableName}[${node.columnName}]`
    case 'CellRef':
      return `${formatSheetPrefix(node.sheetName)}${node.ref}`
    case 'SpillRef':
      return `${formatSheetPrefix(node.sheetName)}${node.ref}#`
    case 'ColumnRef':
      return `${formatSheetPrefix(node.sheetName)}${node.ref}`
    case 'RowRef':
      return `${formatSheetPrefix(node.sheetName)}${node.ref}`
    case 'RangeRef':
      return `${formatRangeSheetPrefix(node.sheetName, node.sheetEndName)}${node.start}:${node.end}`
    case 'UnaryExpr':
      return `${node.operator}${serializeFormula(node.argument, 6)}`
    case 'CallExpr':
      return `${node.callee}(${node.args.map((arg) => serializeFormula(arg)).join(',')})`
    case 'InvokeExpr': {
      const callee =
        node.callee.kind === 'CallExpr' || node.callee.kind === 'InvokeExpr'
          ? serializeFormula(node.callee)
          : `(${serializeFormula(node.callee)})`
      return `${callee}(${node.args.map((arg) => serializeFormula(arg)).join(',')})`
    }
    case 'BinaryExpr': {
      const precedence = BINARY_PRECEDENCE[node.operator]
      const isRightAssociative = node.operator === '^'
      const left = serializeFormula(node.left, precedence, 'left')
      const right = serializeFormula(node.right, precedence, 'right')
      const output = `${left}${node.operator}${right}`
      const needsParens =
        precedence < parentPrecedence ||
        (precedence === parentPrecedence &&
          ((parentAssociativity === 'left' && isRightAssociative) || (parentAssociativity === 'right' && !isRightAssociative)))
      return needsParens ? `(${output})` : output
    }
  }
}
