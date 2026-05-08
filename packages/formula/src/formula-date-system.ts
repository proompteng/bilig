import type { FormulaNode } from './ast.js'
import { normalizeBuiltinLookupName } from './builtins.js'

const dateSystemSensitiveBuiltins = new Set([
  'DATE',
  'DATEVALUE',
  'YEAR',
  'MONTH',
  'DAY',
  'WEEKDAY',
  'WEEKNUM',
  'DAYS360',
  'ISOWEEKNUM',
  'YEARFRAC',
  'WORKDAY',
  'WORKDAY.INTL',
  'NETWORKDAYS',
  'NETWORKDAYS.INTL',
  'TODAY',
  'NOW',
  'EDATE',
  'EOMONTH',
  'DATEDIF',
  'TEXT',
])

export function formulaContainsDateSystemSensitiveBuiltin(node: FormulaNode): boolean {
  switch (node.kind) {
    case 'NumberLiteral':
    case 'BooleanLiteral':
    case 'StringLiteral':
    case 'ErrorLiteral':
    case 'OmittedArgument':
    case 'NameRef':
    case 'StructuredRef':
    case 'CellRef':
    case 'SpillRef':
    case 'RowRef':
    case 'ColumnRef':
    case 'RangeRef':
      return false
    case 'ArrayConstant':
      return node.rows.some((row) => row.some(formulaContainsDateSystemSensitiveBuiltin))
    case 'UnaryExpr':
      return formulaContainsDateSystemSensitiveBuiltin(node.argument)
    case 'BinaryExpr':
      return formulaContainsDateSystemSensitiveBuiltin(node.left) || formulaContainsDateSystemSensitiveBuiltin(node.right)
    case 'CallExpr':
      return (
        dateSystemSensitiveBuiltins.has(normalizeBuiltinLookupName(node.callee)) ||
        node.args.some(formulaContainsDateSystemSensitiveBuiltin)
      )
    case 'InvokeExpr':
      return formulaContainsDateSystemSensitiveBuiltin(node.callee) || node.args.some(formulaContainsDateSystemSensitiveBuiltin)
  }
}
