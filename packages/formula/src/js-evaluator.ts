import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import type { FormulaNode } from './ast.js'
import { formatAddress, parseRangeAddress } from './addressing.js'
import { getBuiltin, getDateSystemBuiltin, normalizeBuiltinLookupName } from './builtins.js'
import { getLookupBuiltin, type RangeBuiltinArgument } from './builtins/lookup.js'
import { evaluateArraySpecialCall } from './js-evaluator-array-special-calls.js'
import { emptyValue, error, numberValue, stringValue } from './js-evaluator-cell-values.js'
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
import {
  cloneScopes,
  cloneStackValue,
  coerceOptionalBooleanArgument,
  coerceOptionalMatchModeArgument,
  coerceOptionalPositiveIntegerArgument,
  coerceOptionalTrimModeArgument,
  coerceScalarTextArgument,
  evaluateBinary,
  getBroadcastShape,
  getRangeCell,
  isCellValueError,
  isSingleCellValue,
  makeArrayStack,
  matrixFromStackValue,
  normalizeScopeName,
  popArgument,
  popScalar,
  scalarIntegerArgument,
  stackScalar,
  toArithmeticNumber,
  toEvaluationResult,
  toNumber,
  toPositiveInteger,
  toRangeArgument,
  toRangeLike,
  toStringValue,
  truthy,
  vectorIntegerArgument,
} from './js-evaluator-runtime-helpers.js'
import { evaluateWorkbookSpecialCall } from './js-evaluator-workbook-special-calls.js'
import type { EvaluationContext, JsPlanInstruction, ReferenceOperand, StackValue } from './js-evaluator-types.js'
import { lowerToPlan } from './js-plan-lowering.js'
import { isArrayValue, scalarFromEvaluationResult, type EvaluationResult } from './runtime-values.js'

export type {
  ApproximateVectorMatchResult,
  EvaluationContext,
  ExactVectorMatchResult,
  JsPlanInstruction,
  ReferenceOperand,
  StackValue,
} from './js-evaluator-types.js'
export { lowerToPlan } from './js-plan-lowering.js'

function resolveBuiltinForContext(name: string, context: EvaluationContext): ((...args: CellValue[]) => EvaluationResult) | undefined {
  const override = context.resolveBuiltin?.(name)
  if (override) {
    return override
  }
  return context.dateSystem ? getDateSystemBuiltin(name, context.dateSystem) : getBuiltin(name)
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
  const builtin = resolveBuiltinForContext(name, context)
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
    parameterScope.set(
      normalizeScopeName(name),
      index < args.length ? cloneStackValue(args[index]!) : { kind: 'omitted', source: 'binding' },
    )
  })
  return executePlan(lambdaValue.body, context, [...cloneScopes(lambdaValue.scopes), parameterScope]) ?? stackScalar(error(ErrorCode.Value))
}

function evaluateSpecialCall(
  callee: string,
  rawArgs: StackValue[],
  context: EvaluationContext,
  argRefs: readonly (ReferenceOperand | undefined)[] = [],
): StackValue | undefined {
  const normalizedCallee = normalizeBuiltinLookupName(callee)
  switch (callee) {
    default:
      return (
        evaluateWorkbookSpecialCall(normalizedCallee, rawArgs, context, argRefs, {
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
        evaluateContextSpecialCall(normalizedCallee, rawArgs, context, argRefs, {
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
        evaluateArraySpecialCall(normalizedCallee, rawArgs, context, {
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

function coerceDirectNumericTextAggregateArgument(callee: string, value: CellValue, argRef: ReferenceOperand | undefined): CellValue {
  if (callee !== 'SUM' || argRef !== undefined || value.tag !== ValueTag.String) {
    return value
  }
  const numeric = toArithmeticNumber(value)
  return numeric === undefined ? error(ErrorCode.Value) : numberValue(numeric)
}

function toLookupBuiltinArgument(callee: string, rawArg: StackValue): CellValue | RangeBuiltinArgument | undefined {
  if (rawArg.kind === 'omitted') {
    return undefined
  }
  if (
    callee === 'SUMPRODUCT' &&
    rawArg.kind === 'scalar' &&
    (rawArg.value.tag === ValueTag.Number || rawArg.value.tag === ValueTag.Boolean || rawArg.value.tag === ValueTag.Empty)
  ) {
    return { kind: 'range', values: [rawArg.value], refKind: 'cells', rows: 1, cols: 1 }
  }
  return toRangeArgument(rawArg)
}

const arrayLiftedScalarBuiltinArities = new Map<string, number>([
  ['IFERROR', 2],
  ['IFNA', 2],
  ['ISBLANK', 1],
  ['ISERR', 1],
  ['ISERROR', 1],
  ['ISEVEN', 1],
  ['ISFORMULA', 1],
  ['ISLOGICAL', 1],
  ['ISNA', 1],
  ['ISNONTEXT', 1],
  ['ISNUMBER', 1],
  ['ISODD', 1],
  ['ISREF', 1],
  ['ISTEXT', 1],
  ['NOT', 1],
  ['ROUND', 2],
])

function evaluateArrayLiftedScalarBuiltin(
  callee: string,
  rawArgs: readonly StackValue[],
  builtin: (...args: CellValue[]) => EvaluationResult,
): StackValue | undefined {
  const expectedArity = arrayLiftedScalarBuiltinArities.get(callee)
  if (
    expectedArity === undefined ||
    rawArgs.length !== expectedArity ||
    !rawArgs.some((arg) => arg.kind === 'array' || arg.kind === 'range')
  ) {
    return undefined
  }
  const shape = getBroadcastShape(rawArgs)
  if (!shape) {
    return stackScalar(error(ErrorCode.Value))
  }
  const ranges = rawArgs.map(toRangeLike)
  const values: CellValue[] = []
  for (let row = 0; row < shape.rows; row += 1) {
    for (let col = 0; col < shape.cols; col += 1) {
      const args = ranges.map((range) => getRangeCell(range, Math.min(row, range.rows - 1), Math.min(col, range.cols - 1)))
      const result = builtin(...args)
      values.push(isArrayValue(result) ? scalarFromEvaluationResult(result) : result)
    }
  }
  return shape.rows === 1 && shape.cols === 1 ? stackScalar(values[0] ?? emptyValue()) : makeArrayStack(shape.rows, shape.cols, values)
}

function evaluateUnary(operator: Extract<JsPlanInstruction, { opcode: 'unary' }>['operator'], value: StackValue): StackValue {
  const coerce = (cellValue: CellValue): CellValue => {
    const numeric = toArithmeticNumber(cellValue)
    return numeric === undefined ? error(ErrorCode.Value) : numberValue(operator === '-' ? -numeric : numeric)
  }

  if (value.kind === 'scalar') {
    return stackScalar(coerce(value.value))
  }
  if (value.kind === 'omitted' || value.kind === 'lambda') {
    return stackScalar(error(ErrorCode.Value))
  }

  const range = toRangeLike(value)
  return makeArrayStack(range.rows, range.cols, range.values.map(coerce))
}

function evaluateDynamicRange(leftValue: StackValue, rightValue: StackValue, context: EvaluationContext): StackValue {
  const left = dynamicRangeEndpoint(leftValue, context)
  if ('tag' in left) {
    return stackScalar(left)
  }
  const right = dynamicRangeEndpoint(rightValue, context)
  if ('tag' in right) {
    return stackScalar(right)
  }
  if (left.sheetName !== right.sheetName) {
    return stackScalar(error(ErrorCode.Value))
  }

  const start = formatAddress(left.row, left.col)
  const end = formatAddress(right.row, right.col)
  const parsed = parseRangeAddress(`${start}:${end}`, left.sheetName)
  if (parsed.kind !== 'cells') {
    return stackScalar(error(ErrorCode.Value))
  }

  const values = context.resolveRange(left.sheetName, start, end, 'cells')
  context.noteRangeMaterialization?.(values.length)
  return {
    kind: 'range',
    values,
    refKind: 'cells',
    rows: parsed.end.row - parsed.start.row + 1,
    cols: parsed.end.col - parsed.start.col + 1,
    sheetName: left.sheetName,
    start: parsed.start.text,
    end: parsed.end.text,
  }
}

function dynamicRangeEndpoint(value: StackValue, context: EvaluationContext): { sheetName: string; row: number; col: number } | CellValue {
  if (value.kind !== 'range' || value.refKind !== 'cells' || value.rows !== 1 || value.cols !== 1 || !value.start || !value.end) {
    return error(ErrorCode.Value)
  }
  const sheetName = value.sheetName ?? context.sheetName
  const parsed = parseRangeAddress(`${value.start}:${value.end}`, sheetName)
  if (parsed.kind !== 'cells' || parsed.start.row !== parsed.end.row || parsed.start.col !== parsed.end.col) {
    return error(ErrorCode.Value)
  }
  return {
    sheetName: parsed.sheetName ?? sheetName,
    row: parsed.start.row,
    col: parsed.start.col,
  }
}

function sheetNamesInRange(context: EvaluationContext, sheetName: string, sheetEndName: string): string[] | undefined {
  const names = context.listSheetNames?.()
  if (!names) {
    return undefined
  }
  const startIndex = names.indexOf(sheetName)
  const endIndex = names.indexOf(sheetEndName)
  if (startIndex === -1 || endIndex === -1 || startIndex > endIndex) {
    return undefined
  }
  return names.slice(startIndex, endIndex + 1)
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
    context.checkEvaluationBudget?.()
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
                  value: context.resolveName?.(instruction.name, instruction.sheetName) ?? error(ErrorCode.Name),
                },
          )
        }
        break
      case 'push-omitted':
        stack.push({ kind: 'omitted', source: 'argument' })
        break
      case 'make-array': {
        const values = Array<CellValue>(instruction.rows * instruction.cols)
        for (let index = values.length - 1; index >= 0; index -= 1) {
          const scalar = isSingleCellValue(popArgument(stack))
          values[index] = scalar ?? error(ErrorCode.Value)
        }
        stack.push(makeArrayStack(instruction.rows, instruction.cols, values))
        break
      }
      case 'push-cell':
        {
          const value = context.resolveCell(instruction.sheetName ?? context.sheetName, instruction.address)
          stack.push(stackScalar(value, value.tag === ValueTag.Empty))
        }
        break
      case 'push-range':
        {
          if (instruction.sheetEndName !== undefined) {
            const startSheetName = instruction.sheetName ?? context.sheetName
            const rangeSheetNames = sheetNamesInRange(context, startSheetName, instruction.sheetEndName)
            if (!rangeSheetNames || instruction.refKind !== 'cells') {
              stack.push({ kind: 'scalar', value: error(ErrorCode.Ref) })
              break
            }
            const values = rangeSheetNames.flatMap((sheetName) =>
              context.resolveRange(sheetName, instruction.start, instruction.end, instruction.refKind),
            )
            context.noteRangeMaterialization?.(values.length)
            stack.push({
              kind: 'range',
              values,
              refKind: instruction.refKind,
              rows: values.length,
              cols: 1,
              sheetName: startSheetName,
              sheetEndName: instruction.sheetEndName,
              start: instruction.start,
              end: instruction.end,
            })
            break
          }
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
        stack.push(evaluateUnary(instruction.operator, popArgument(stack)))
        break
      }
      case 'binary': {
        const right = popArgument(stack)
        const left = popArgument(stack)
        if (instruction.operator === ':') {
          stack.push(evaluateDynamicRange(left, right, context))
          break
        }
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
          const args: Array<CellValue | RangeBuiltinArgument | undefined> = []
          for (const rawArg of rawArgs) {
            args.push(toLookupBuiltinArgument(instruction.callee, rawArg))
          }
          const result = lookupBuiltin(...args)
          stack.push(isArrayValue(result) ? result : { kind: 'scalar', value: result })
          break
        }

        const builtin = resolveBuiltinForContext(instruction.callee, context)
        if (!builtin) {
          stack.push({ kind: 'scalar', value: error(ErrorCode.Name) })
          break
        }
        const liftedResult = evaluateArrayLiftedScalarBuiltin(instruction.callee, rawArgs, builtin)
        if (liftedResult) {
          stack.push(liftedResult)
          break
        }
        const args: CellValue[] = []
        for (const [index, rawArg] of rawArgs.entries()) {
          if (rawArg.kind === 'scalar') {
            args.push(coerceDirectNumericTextAggregateArgument(instruction.callee, rawArg.value, instruction.argRefs?.[index]))
            continue
          }
          if (rawArg.kind === 'omitted') {
            if (rawArg.source === 'argument') {
              args.push(emptyValue())
              continue
            }
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
