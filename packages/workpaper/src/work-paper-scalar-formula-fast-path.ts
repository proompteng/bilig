import {
  evaluateAstResult,
  evaluatePlanResult,
  compareScalars,
  coerceLogicalValue,
  getBuiltin,
  getDateSystemBuiltin,
  excelExponentiation,
  isArrayValue,
  lowerToPlan,
  normalizeBuiltinLookupName,
  parseArithmeticScalarText,
  parseFormula,
  coerceDirectNumericTextAggregateArgument,
  coerceDirectNumericTextStatisticArgument,
  referencedScalarBuiltinAStyleValues,
  referencedScalarBuiltinAggregateValues,
  referencedScalarBuiltinLogicalValues,
  referencedScalarBuiltinStatisticValues,
  scalarFromEvaluationResult,
  type EvaluationContext,
  type EvaluationResult,
  type ExcelDateSystem,
  type FormulaNode,
} from '@bilig/formula'
import { ErrorCode, ValueTag, type CellValue, type LiteralInput } from '@bilig/protocol'
import {
  emptyValue,
  errorValue,
  makeNamedExpressionKey,
  stripLeadingEquals,
  tryEvaluateSimpleNamedExpression,
} from './work-paper-runtime-helpers.js'
import type {
  SerializedWorkPaperNamedExpression,
  WorkPaperCompiledScalarFormula,
  WorkPaperConfig,
  WorkPaperScalarFormulaEnvironment,
  WorkPaperScalarFormulaVariableValue,
} from './work-paper-types.js'

const scratchRequiredCalls = new Set([
  'AGGREGATE',
  'CELL',
  'COLUMN',
  'COLUMNS',
  'FORMULA',
  'FORMULATEXT',
  'GETPIVOTDATA',
  'GROUPBY',
  'HLOOKUP',
  'INDEX',
  'INDIRECT',
  'MATCH',
  'MULTIPLE.OPERATIONS',
  'OFFSET',
  'ROW',
  'ROWS',
  'SHEET',
  'SHEETS',
  'SINGLE',
  'SUBTOTAL',
  'VLOOKUP',
  'XMATCH',
])

interface DirectScalarRuntime {
  readonly resolveCell: (address: string) => CellValue
  readonly resolveName: (name: string) => CellValue
}

type DirectScalarEvaluator = (runtime: DirectScalarRuntime) => CellValue

export function tryCalculatePureWorkPaperFormula(args: {
  readonly config: WorkPaperConfig
  readonly formula: string
  readonly namedExpressions: readonly SerializedWorkPaperNamedExpression[]
  readonly scope?: number
}): CellValue | CellValue[][] | undefined {
  if (!args.formula.trim().startsWith('=') || (args.config.functionPlugins?.length ?? 0) > 0) {
    return undefined
  }

  const ast = parseFormula(stripLeadingEquals(args.formula))
  const nameRefs = new Set<string>()
  const dateSystem = args.config.calculationSettings?.dateSystem
  if (!isPureScalarFormulaNode(ast, nameRefs, dateSystem)) {
    return undefined
  }

  const namedScalars = collectSimpleNamedScalarValues(args.namedExpressions, args.scope)
  for (const name of nameRefs) {
    if (!namedScalars.has(makeNamedExpressionKey(name, args.scope)) && !namedScalars.has(makeNamedExpressionKey(name))) {
      return undefined
    }
  }

  const context: EvaluationContext = {
    sheetName: '__WORKPAPER_SCALAR__',
    resolveCell: () => errorValue(ErrorCode.Ref),
    resolveRange: () => [],
    resolveName: (name) =>
      namedScalars.get(makeNamedExpressionKey(name, args.scope)) ??
      namedScalars.get(makeNamedExpressionKey(name)) ??
      errorValue(ErrorCode.Name),
    ...(dateSystem !== undefined ? { dateSystem } : {}),
  }
  const result = evaluateAstResult(ast, context)
  if (!isArrayValue(result)) {
    return result
  }
  const matrix: CellValue[][] = []
  for (let row = 0; row < result.rows; row += 1) {
    const values: CellValue[] = []
    for (let col = 0; col < result.cols; col += 1) {
      values.push(result.values[row * result.cols + col] ?? emptyValue())
    }
    matrix.push(values)
  }
  return matrix
}

export function tryCompilePureWorkPaperScalarFormula(args: {
  readonly config: WorkPaperConfig
  readonly formula: string
  readonly namedExpressions: readonly SerializedWorkPaperNamedExpression[]
  readonly scope?: number
}): WorkPaperCompiledScalarFormula | undefined {
  if (!args.formula.trim().startsWith('=') || (args.config.functionPlugins?.length ?? 0) > 0) {
    return undefined
  }

  const ast = parseFormula(stripLeadingEquals(args.formula))
  const variableRefs = new Set<string>()
  const dateSystem = args.config.calculationSettings?.dateSystem
  if (!isScalarFormulaNode(ast, variableRefs, dateSystem, true)) {
    return undefined
  }

  const plan = lowerToPlan(ast)
  const directEvaluator = compileDirectScalarEvaluator(ast, dateSystem)
  const namedScalars = collectSimpleNamedScalarValues(args.namedExpressions, args.scope)
  const variables = [...variableRefs].toSorted()
  const formula = args.formula

  return {
    formula,
    variables,
    evaluate(environment: WorkPaperScalarFormulaEnvironment = {}) {
      const runtime: DirectScalarRuntime = {
        resolveCell: (address) => readScalarEnvironmentValue(environment, address) ?? errorValue(ErrorCode.Ref),
        resolveName: (name) =>
          readScalarEnvironmentValue(environment, name) ??
          namedScalars.get(makeNamedExpressionKey(name, args.scope)) ??
          namedScalars.get(makeNamedExpressionKey(name)) ??
          errorValue(ErrorCode.Name),
      }
      if (directEvaluator !== undefined) {
        return directEvaluator(runtime)
      }

      const context: EvaluationContext = {
        sheetName: '__WORKPAPER_SCALAR__',
        resolveCell: (sheetName, address) =>
          readScalarEnvironmentValue(environment, `${sheetName}!${address}`) ??
          readScalarEnvironmentValue(environment, address) ??
          errorValue(ErrorCode.Ref),
        resolveRange: () => [],
        resolveName: (name, sheetName) =>
          readScalarEnvironmentValue(environment, sheetName ? `${sheetName}!${name}` : name) ??
          readScalarEnvironmentValue(environment, name) ??
          namedScalars.get(makeNamedExpressionKey(name, args.scope)) ??
          namedScalars.get(makeNamedExpressionKey(name)) ??
          errorValue(ErrorCode.Name),
        ...(dateSystem !== undefined ? { dateSystem } : {}),
      }
      const result = evaluatePlanResult(plan, context)
      return isArrayValue(result) ? matrixFromArrayValue(result) : result
    },
  }
}

function collectSimpleNamedScalarValues(
  namedExpressions: readonly SerializedWorkPaperNamedExpression[],
  scope: number | undefined,
): Map<string, CellValue> {
  const values = new Map<string, CellValue>()
  for (const expression of namedExpressions) {
    if (expression.scope !== undefined && expression.scope !== scope) {
      continue
    }
    const value = tryEvaluateSimpleNamedExpression(expression.expression)
    if (value !== undefined) {
      values.set(makeNamedExpressionKey(expression.name, expression.scope), value)
    }
  }
  return values
}

function isPureScalarFormulaNode(node: FormulaNode, nameRefs: Set<string>, dateSystem: '1900' | '1904' | undefined): boolean {
  return isScalarFormulaNode(node, nameRefs, dateSystem, false)
}

function isScalarFormulaNode(
  node: FormulaNode,
  variableRefs: Set<string>,
  dateSystem: '1900' | '1904' | undefined,
  allowCellRefs: boolean,
): boolean {
  switch (node.kind) {
    case 'BooleanLiteral':
    case 'ErrorLiteral':
    case 'NumberLiteral':
    case 'OmittedArgument':
    case 'StringLiteral':
      return true
    case 'ArrayConstant':
      return node.rows.every((row) => row.every((entry) => isScalarFormulaNode(entry, variableRefs, dateSystem, allowCellRefs)))
    case 'NameRef':
      if (node.sheetName !== undefined) {
        return false
      }
      variableRefs.add(normalizeScalarVariableKey(node.name))
      return true
    case 'CellRef':
      if (!allowCellRefs || node.sheetName !== undefined) {
        return false
      }
      variableRefs.add(normalizeScalarVariableKey(node.ref))
      return true
    case 'ColumnRef':
    case 'RangeRef':
    case 'RowRef':
    case 'SpillRef':
    case 'StructuredRef':
      return false
    case 'UnaryExpr':
      return isScalarFormulaNode(node.argument, variableRefs, dateSystem, allowCellRefs)
    case 'BinaryExpr':
      return (
        node.operator !== ':' &&
        isScalarFormulaNode(node.left, variableRefs, dateSystem, allowCellRefs) &&
        isScalarFormulaNode(node.right, variableRefs, dateSystem, allowCellRefs)
      )
    case 'CallExpr': {
      const callee = normalizeBuiltinLookupName(node.callee)
      return (
        !scratchRequiredCalls.has(callee) &&
        isKnownFormulaBuiltin(callee, dateSystem) &&
        node.args.every((argument) => isScalarFormulaNode(argument, variableRefs, dateSystem, allowCellRefs))
      )
    }
    case 'InvokeExpr':
      return false
    default:
      return false
  }
}

function isKnownFormulaBuiltin(callee: string, dateSystem: '1900' | '1904' | undefined): boolean {
  return (dateSystem !== undefined ? getDateSystemBuiltin(callee, dateSystem) : getBuiltin(callee)) !== undefined
}

function compileDirectScalarEvaluator(node: FormulaNode, dateSystem: '1900' | '1904' | undefined): DirectScalarEvaluator | undefined {
  switch (node.kind) {
    case 'NumberLiteral': {
      const value: CellValue = { tag: ValueTag.Number, value: node.value }
      return () => value
    }
    case 'BooleanLiteral': {
      const value: CellValue = { tag: ValueTag.Boolean, value: node.value }
      return () => value
    }
    case 'StringLiteral': {
      const value: CellValue = { tag: ValueTag.String, value: node.value, stringId: 0 }
      return () => value
    }
    case 'ErrorLiteral': {
      const value = errorValue(node.code)
      return () => value
    }
    case 'OmittedArgument':
      return () => emptyValue()
    case 'CellRef':
      if (node.sheetName !== undefined) {
        return undefined
      }
      return (runtime) => runtime.resolveCell(node.ref)
    case 'NameRef':
      if (node.sheetName !== undefined) {
        return undefined
      }
      return (runtime) => runtime.resolveName(node.name)
    case 'UnaryExpr': {
      const argument = compileDirectScalarEvaluator(node.argument, dateSystem)
      if (argument === undefined) {
        return undefined
      }
      return (runtime) => evaluateDirectUnary(node.operator, argument(runtime), dateSystem ?? '1900')
    }
    case 'BinaryExpr': {
      if (node.operator === ':') {
        return undefined
      }
      const left = compileDirectScalarEvaluator(node.left, dateSystem)
      const right = compileDirectScalarEvaluator(node.right, dateSystem)
      if (left === undefined || right === undefined) {
        return undefined
      }
      return (runtime) => evaluateDirectBinary(node.operator, left(runtime), right(runtime), dateSystem ?? '1900')
    }
    case 'CallExpr': {
      const callee = normalizeBuiltinLookupName(node.callee)
      if (callee === 'IF') {
        if (node.args.length < 2) {
          return () => errorValue(ErrorCode.Value)
        }
        const condition = compileDirectScalarEvaluator(node.args[0]!, dateSystem)
        const whenTrue = compileDirectIfBranchEvaluator(node.args[1], dateSystem, errorValue(ErrorCode.Value))
        const missingFalseValue: CellValue = { tag: ValueTag.Boolean, value: false }
        const whenFalse =
          node.args.length >= 3
            ? compileDirectIfBranchEvaluator(node.args[2], dateSystem, { tag: ValueTag.Number, value: 0 })
            : () => missingFalseValue
        if (condition === undefined || whenTrue === undefined || whenFalse === undefined) {
          return undefined
        }
        return (runtime) => {
          const conditionValue = condition(runtime)
          const coerced = coerceLogicalValue(conditionValue)
          if (!coerced.ok) {
            return coerced.error
          }
          return coerced.value ? whenTrue(runtime) : whenFalse(runtime)
        }
      }
      const builtin = dateSystem !== undefined ? getDateSystemBuiltin(callee, dateSystem) : getBuiltin(callee)
      if (builtin === undefined) {
        return undefined
      }
      const evaluators: DirectScalarEvaluator[] = []
      for (let index = 0; index < node.args.length; index += 1) {
        const evaluator = compileDirectScalarEvaluator(node.args[index]!, dateSystem)
        if (evaluator === undefined) {
          return undefined
        }
        evaluators.push(evaluator)
      }
      return (runtime) => {
        const values = Array.from({ length: evaluators.length }, () => emptyValue())
        for (let index = 0; index < evaluators.length; index += 1) {
          values[index] = evaluators[index]!(runtime)
        }
        return scalarFromDirectResult(builtin(...directScalarCallArguments(callee, values, node.args)))
      }
    }
    case 'ArrayConstant':
    case 'ColumnRef':
    case 'InvokeExpr':
    case 'RangeRef':
    case 'RowRef':
    case 'SpillRef':
    case 'StructuredRef':
      return undefined
  }
}

function directScalarCallArguments(callee: string, values: readonly CellValue[], argNodes: readonly FormulaNode[]): CellValue[] {
  const args: CellValue[] = []
  for (let index = 0; index < values.length; index += 1) {
    const value = values[index] ?? emptyValue()
    const isReferenceArg = argNodes[index]?.kind === 'CellRef'
    if (isReferenceArg) {
      const aggregateValues = referencedScalarBuiltinAggregateValues(callee, value)
      if (aggregateValues) {
        args.push(...aggregateValues)
        continue
      }
      const statisticValues = referencedScalarBuiltinStatisticValues(callee, value)
      if (statisticValues) {
        args.push(...statisticValues)
        continue
      }
      const logicalValues = referencedScalarBuiltinLogicalValues(callee, value)
      if (logicalValues) {
        args.push(...logicalValues)
        continue
      }
      const aStyleValues = referencedScalarBuiltinAStyleValues(callee, value)
      if (aStyleValues) {
        args.push(...aStyleValues)
        continue
      }
    }
    args.push(
      coerceDirectNumericTextStatisticArgument(
        callee,
        coerceDirectNumericTextAggregateArgument(callee, value, isReferenceArg),
        isReferenceArg,
      ),
    )
  }
  return args
}

function compileDirectIfBranchEvaluator(
  node: FormulaNode | undefined,
  dateSystem: '1900' | '1904' | undefined,
  missingValue: CellValue,
): DirectScalarEvaluator | undefined {
  if (node === undefined) {
    return () => missingValue
  }
  if (node.kind === 'OmittedArgument') {
    const value: CellValue = { tag: ValueTag.Number, value: 0 }
    return () => value
  }
  return compileDirectScalarEvaluator(node, dateSystem)
}

function evaluateDirectUnary(operator: string, value: CellValue, dateSystem: ExcelDateSystem): CellValue {
  if (value.tag === ValueTag.Error) {
    return value
  }
  if (operator === '+') {
    const numeric = directNumber(value, dateSystem)
    return typeof numeric === 'number' ? { tag: ValueTag.Number, value: numeric } : numeric
  }
  if (operator === '-') {
    const numeric = directNumber(value, dateSystem)
    return typeof numeric === 'number' ? { tag: ValueTag.Number, value: -numeric } : numeric
  }
  return errorValue(ErrorCode.Value)
}

function evaluateDirectBinary(operator: string, left: CellValue, right: CellValue, dateSystem: ExcelDateSystem): CellValue {
  if (left.tag === ValueTag.Error) {
    return left
  }
  if (right.tag === ValueTag.Error) {
    return right
  }
  if (operator === '&') {
    return { tag: ValueTag.String, value: `${directText(left)}${directText(right)}`, stringId: 0 }
  }
  if (operator === '=' || operator === '<>' || operator === '>' || operator === '>=' || operator === '<' || operator === '<=') {
    const comparison = directCompare(left, right)
    if (comparison === undefined) {
      return errorValue(ErrorCode.Value)
    }
    const value =
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
                : comparison <= 0
    return { tag: ValueTag.Boolean, value }
  }

  const leftNumber = directNumber(left, dateSystem)
  if (typeof leftNumber !== 'number') {
    return leftNumber
  }
  const rightNumber = directNumber(right, dateSystem)
  if (typeof rightNumber !== 'number') {
    return rightNumber
  }
  switch (operator) {
    case '+':
      return { tag: ValueTag.Number, value: leftNumber + rightNumber }
    case '-':
      return { tag: ValueTag.Number, value: leftNumber - rightNumber }
    case '*':
      return { tag: ValueTag.Number, value: leftNumber * rightNumber }
    case '/':
      return rightNumber === 0 ? errorValue(ErrorCode.Div0) : { tag: ValueTag.Number, value: leftNumber / rightNumber }
    case '^': {
      const value = excelExponentiation(leftNumber, rightNumber)
      return Number.isFinite(value) ? { tag: ValueTag.Number, value } : errorValue(ErrorCode.Num)
    }
    default:
      return errorValue(ErrorCode.Value)
  }
}

function directNumber(value: CellValue, dateSystem: ExcelDateSystem = '1900'): number | CellValue {
  switch (value.tag) {
    case ValueTag.Number:
      return value.value
    case ValueTag.Boolean:
      return value.value ? 1 : 0
    case ValueTag.Empty:
      return 0
    case ValueTag.String: {
      const numeric = parseArithmeticScalarText(value.value, dateSystem)
      return numeric === undefined ? errorValue(ErrorCode.Value) : numeric
    }
    case ValueTag.Error:
      return value
  }
}

function directText(value: CellValue): string {
  switch (value.tag) {
    case ValueTag.Number:
      return String(value.value)
    case ValueTag.Boolean:
      return value.value ? 'TRUE' : 'FALSE'
    case ValueTag.String:
      return value.value
    case ValueTag.Empty:
      return ''
    case ValueTag.Error:
      return ''
  }
}

function directCompare(left: CellValue, right: CellValue): number | undefined {
  return compareScalars(left, right)
}

function scalarFromDirectResult(result: EvaluationResult): CellValue {
  return isArrayValue(result) ? scalarFromEvaluationResult(result) : result
}

function matrixFromArrayValue(result: {
  readonly rows: number
  readonly cols: number
  readonly values: readonly CellValue[]
}): CellValue[][] {
  const matrix: CellValue[][] = []
  for (let row = 0; row < result.rows; row += 1) {
    const values: CellValue[] = []
    for (let col = 0; col < result.cols; col += 1) {
      values.push(result.values[row * result.cols + col] ?? emptyValue())
    }
    matrix.push(values)
  }
  return matrix
}

function readScalarEnvironmentValue(environment: WorkPaperScalarFormulaEnvironment, rawKey: string): CellValue | undefined {
  const directValue = readOwnEnvironmentValue(environment, rawKey)
  if (directValue !== undefined) {
    return directValue
  }

  const normalizedKey = normalizeScalarVariableKey(rawKey)
  if (normalizedKey !== rawKey) {
    const normalizedValue = readOwnEnvironmentValue(environment, normalizedKey)
    if (normalizedValue !== undefined) {
      return normalizedValue
    }
  }

  for (const [key, value] of Object.entries(environment)) {
    if (normalizeScalarVariableKey(key) === normalizedKey) {
      return scalarVariableValueToCellValue(value)
    }
  }

  return undefined
}

function readOwnEnvironmentValue(environment: WorkPaperScalarFormulaEnvironment, key: string): CellValue | undefined {
  if (!Object.prototype.hasOwnProperty.call(environment, key)) {
    return undefined
  }
  return scalarVariableValueToCellValue(environment[key]!)
}

function scalarVariableValueToCellValue(value: WorkPaperScalarFormulaVariableValue): CellValue {
  if (isCellValue(value)) {
    return value
  }
  return literalToCellValue(value)
}

function literalToCellValue(value: LiteralInput): CellValue {
  if (typeof value === 'number') {
    return { tag: ValueTag.Number, value }
  }
  if (typeof value === 'boolean') {
    return { tag: ValueTag.Boolean, value }
  }
  if (typeof value === 'string') {
    return { tag: ValueTag.String, value, stringId: 0 }
  }
  return emptyValue()
}

function isCellValue(value: WorkPaperScalarFormulaVariableValue): value is CellValue {
  return typeof value === 'object' && value !== null && 'tag' in value
}

function normalizeScalarVariableKey(key: string): string {
  return key.trim().replace(/\$/g, '').toUpperCase()
}
