import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import type { FormulaNode } from './ast.js'
import { parseRangeAddress } from './addressing.js'
import { getBuiltin, normalizeBuiltinLookupName } from './builtins.js'
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
