import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import { getLookupBuiltin, type RangeBuiltinArgument } from './builtins/lookup.js'
import { hasLookupWildcardSyntax } from './builtins/lookup-reference-search.js'
import { error } from './js-evaluator-cell-values.js'
import { isSingleCellValue } from './js-evaluator-runtime-helpers.js'
import type { EvaluationContext, JsPlanInstruction, StackValue } from './js-evaluator-types.js'
import { isArrayValue } from './runtime-values.js'

type LookupMatchInstruction = Extract<JsPlanInstruction, { opcode: 'lookup-exact-match' | 'lookup-approximate-match' }>

export function evaluateLookupMatchOpcode(args: {
  readonly instruction: LookupMatchInstruction
  readonly lookupOperand: StackValue
  readonly context: EvaluationContext
}): StackValue {
  const lookupValue = isSingleCellValue(args.lookupOperand)
  if (!lookupValue) {
    return { kind: 'scalar', value: error(ErrorCode.Value) }
  }

  return args.instruction.opcode === 'lookup-exact-match'
    ? evaluateExactLookupMatchOpcode(args.instruction, lookupValue, args.context)
    : evaluateApproximateLookupMatchOpcode(args.instruction, lookupValue, args.context)
}

function evaluateExactLookupMatchOpcode(
  instruction: Extract<LookupMatchInstruction, { opcode: 'lookup-exact-match' }>,
  lookupValue: CellValue,
  context: EvaluationContext,
): StackValue {
  const sheetName = instruction.sheetName ?? context.sheetName
  const directMatch =
    instruction.callee === 'MATCH' && hasLookupWildcardSyntax(lookupValue)
      ? undefined
      : context.resolveExactVectorMatch?.({
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
    return {
      kind: 'scalar',
      value: directMatch.position === undefined ? error(ErrorCode.NA) : { tag: ValueTag.Number, value: directMatch.position },
    }
  }

  context.noteExactLookupFallback?.()
  const rangeArg = materializeLookupRange(instruction, sheetName, context)
  const lookupBuiltin = context.resolveLookupBuiltin?.(instruction.callee) ?? getLookupBuiltin(instruction.callee)
  if (!lookupBuiltin) {
    return { kind: 'scalar', value: error(ErrorCode.Name) }
  }

  const result =
    instruction.callee === 'MATCH'
      ? lookupBuiltin(lookupValue, rangeArg, { tag: ValueTag.Number, value: 0 })
      : lookupBuiltin(lookupValue, rangeArg, { tag: ValueTag.Number, value: 0 }, { tag: ValueTag.Number, value: instruction.searchMode })
  return isArrayValue(result) ? result : { kind: 'scalar', value: result }
}

function evaluateApproximateLookupMatchOpcode(
  instruction: Extract<LookupMatchInstruction, { opcode: 'lookup-approximate-match' }>,
  lookupValue: CellValue,
  context: EvaluationContext,
): StackValue {
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
    return {
      kind: 'scalar',
      value: directMatch.position === undefined ? error(ErrorCode.NA) : { tag: ValueTag.Number, value: directMatch.position },
    }
  }

  const rangeArg = materializeLookupRange(instruction, sheetName, context)
  const lookupBuiltin = context.resolveLookupBuiltin?.(instruction.callee) ?? getLookupBuiltin(instruction.callee)
  if (!lookupBuiltin) {
    return { kind: 'scalar', value: error(ErrorCode.Name) }
  }

  const matchModeValue = { tag: ValueTag.Number, value: instruction.matchMode } as const
  const result =
    instruction.callee === 'MATCH'
      ? lookupBuiltin(lookupValue, rangeArg, matchModeValue)
      : lookupBuiltin(lookupValue, rangeArg, matchModeValue)
  return isArrayValue(result) ? result : { kind: 'scalar', value: result }
}

function materializeLookupRange(instruction: LookupMatchInstruction, sheetName: string, context: EvaluationContext): RangeBuiltinArgument {
  const values = context.resolveRange(sheetName, instruction.start, instruction.end, instruction.refKind)
  context.noteRangeMaterialization?.(values.length)
  return {
    kind: 'range',
    values,
    refKind: instruction.refKind,
    rows: instruction.endRow - instruction.startRow + 1,
    cols: instruction.endCol - instruction.startCol + 1,
    sheetName,
    start: instruction.start,
    end: instruction.end,
  }
}
