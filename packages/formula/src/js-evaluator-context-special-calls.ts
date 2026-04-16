import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import type { EvaluationContext, ReferenceOperand, StackValue } from './js-evaluator.js'

interface ContextSpecialCallDeps {
  error: (code: ErrorCode) => CellValue
  emptyValue: () => CellValue
  numberValue: (value: number) => CellValue
  stringValue: (value: string) => CellValue
  stackScalar: (value: CellValue) => StackValue
  cloneStackValue: (value: StackValue) => StackValue
  toNumber: (value: CellValue) => number | undefined
  toStringValue: (value: CellValue) => string
  isSingleCellValue: (value: StackValue) => CellValue | undefined
  currentCellReference: (context: EvaluationContext) => ReferenceOperand | undefined
  referenceSheetName: (ref: ReferenceOperand | undefined, context: EvaluationContext) => string | undefined
  referenceTopLeftAddress: (ref: ReferenceOperand | undefined) => string | undefined
  referenceRowNumber: (ref: ReferenceOperand | undefined, context: EvaluationContext) => number | undefined
  referenceColumnNumber: (ref: ReferenceOperand | undefined, context: EvaluationContext) => number | undefined
  absoluteAddress: (ref: ReferenceOperand | undefined, context: EvaluationContext) => string | undefined
  cellTypeCode: (value: CellValue) => string
  sheetNames: (context: EvaluationContext) => string[]
  sheetIndexByName: (name: string, context: EvaluationContext) => number | undefined
}

export function evaluateContextSpecialCall(
  callee: string,
  rawArgs: StackValue[],
  context: EvaluationContext,
  argRefs: readonly (ReferenceOperand | undefined)[],
  deps: ContextSpecialCallDeps,
): StackValue | undefined {
  switch (callee) {
    case 'ROW': {
      if (rawArgs.length > 1) {
        return deps.stackScalar(deps.error(ErrorCode.Value))
      }
      const row = deps.referenceRowNumber(argRefs[0], context)
      return deps.stackScalar(row === undefined ? deps.error(ErrorCode.Value) : deps.numberValue(row))
    }
    case 'COLUMN': {
      if (rawArgs.length > 1) {
        return deps.stackScalar(deps.error(ErrorCode.Value))
      }
      const column = deps.referenceColumnNumber(argRefs[0], context)
      return deps.stackScalar(column === undefined ? deps.error(ErrorCode.Value) : deps.numberValue(column))
    }
    case 'ISOMITTED': {
      if (rawArgs.length !== 1) {
        return deps.stackScalar(deps.error(ErrorCode.Value))
      }
      return deps.stackScalar({ tag: ValueTag.Boolean, value: rawArgs[0]?.kind === 'omitted' })
    }
    case 'FORMULATEXT':
    case 'FORMULA': {
      if (rawArgs.length !== 1) {
        return deps.stackScalar(deps.error(ErrorCode.Value))
      }
      const address = deps.referenceTopLeftAddress(argRefs[0])
      const sheetName = deps.referenceSheetName(argRefs[0], context)
      if (!address || !sheetName) {
        return deps.stackScalar(deps.error(ErrorCode.Ref))
      }
      const formula = context.resolveFormula?.(sheetName, address)
      return deps.stackScalar(formula ? deps.stringValue(formula.startsWith('=') ? formula : `=${formula}`) : deps.error(ErrorCode.NA))
    }
    case 'PHONETIC': {
      if (rawArgs.length !== 1) {
        return deps.stackScalar(deps.error(ErrorCode.Value))
      }
      const target = rawArgs[0]!
      if (target.kind === 'scalar') {
        return deps.stackScalar(deps.stringValue(deps.toStringValue(target.value)))
      }
      if (target.kind === 'range') {
        return deps.stackScalar(deps.stringValue(deps.toStringValue(target.values[0] ?? deps.emptyValue())))
      }
      return deps.stackScalar(deps.error(ErrorCode.Value))
    }
    case 'CHOOSE': {
      if (rawArgs.length < 2) {
        return deps.stackScalar(deps.error(ErrorCode.Value))
      }
      const indexValue = deps.isSingleCellValue(rawArgs[0]!)
      const choice = indexValue ? deps.toNumber(indexValue) : undefined
      if (choice === undefined || !Number.isFinite(choice)) {
        return deps.stackScalar(deps.error(ErrorCode.Value))
      }
      const truncated = Math.trunc(choice)
      if (truncated < 1 || truncated >= rawArgs.length) {
        return deps.stackScalar(deps.error(ErrorCode.Value))
      }
      return deps.cloneStackValue(rawArgs[truncated]!)
    }
    case 'SHEET': {
      if (rawArgs.length > 1) {
        return deps.stackScalar(deps.error(ErrorCode.Value))
      }
      if (rawArgs.length === 0) {
        const index = deps.sheetIndexByName(context.sheetName, context)
        return deps.stackScalar(index === undefined ? deps.error(ErrorCode.NA) : deps.numberValue(index))
      }
      if (argRefs[0]) {
        const index = deps.sheetIndexByName(deps.referenceSheetName(argRefs[0], context) ?? context.sheetName, context)
        return deps.stackScalar(index === undefined ? deps.error(ErrorCode.NA) : deps.numberValue(index))
      }
      const scalar = deps.isSingleCellValue(rawArgs[0]!)
      if (scalar?.tag !== ValueTag.String) {
        return deps.stackScalar(deps.error(ErrorCode.NA))
      }
      const index = deps.sheetIndexByName(scalar.value, context)
      return deps.stackScalar(index === undefined ? deps.error(ErrorCode.NA) : deps.numberValue(index))
    }
    case 'SHEETS': {
      if (rawArgs.length > 1) {
        return deps.stackScalar(deps.error(ErrorCode.Value))
      }
      if (rawArgs.length === 0) {
        return deps.stackScalar(deps.numberValue(deps.sheetNames(context).length))
      }
      if (argRefs[0]) {
        return deps.stackScalar(deps.numberValue(1))
      }
      const scalar = deps.isSingleCellValue(rawArgs[0]!)
      if (scalar?.tag !== ValueTag.String) {
        return deps.stackScalar(deps.error(ErrorCode.NA))
      }
      return deps.stackScalar(deps.sheetIndexByName(scalar.value, context) === undefined ? deps.error(ErrorCode.NA) : deps.numberValue(1))
    }
    case 'CELL': {
      if (rawArgs.length < 1 || rawArgs.length > 2) {
        return deps.stackScalar(deps.error(ErrorCode.Value))
      }
      const infoType = deps.isSingleCellValue(rawArgs[0]!)
      if (infoType?.tag !== ValueTag.String) {
        return deps.stackScalar(deps.error(ErrorCode.Value))
      }
      const ref = rawArgs.length === 2 ? argRefs[1] : deps.currentCellReference(context)
      if (!ref) {
        return deps.stackScalar(deps.error(ErrorCode.Value))
      }
      const normalizedInfoType = infoType.value.trim().toLowerCase()
      switch (normalizedInfoType) {
        case 'address': {
          const address = deps.absoluteAddress(ref, context)
          return deps.stackScalar(address ? deps.stringValue(address) : deps.error(ErrorCode.Value))
        }
        case 'row': {
          const row = deps.referenceRowNumber(ref, context)
          return deps.stackScalar(row === undefined ? deps.error(ErrorCode.Value) : deps.numberValue(row))
        }
        case 'col': {
          const column = deps.referenceColumnNumber(ref, context)
          return deps.stackScalar(column === undefined ? deps.error(ErrorCode.Value) : deps.numberValue(column))
        }
        case 'contents': {
          const address = deps.referenceTopLeftAddress(ref)
          const sheetName = deps.referenceSheetName(ref, context)
          if (!address || !sheetName) {
            return deps.stackScalar(deps.error(ErrorCode.Value))
          }
          return deps.stackScalar(context.resolveCell(sheetName, address))
        }
        case 'type': {
          const address = deps.referenceTopLeftAddress(ref)
          const sheetName = deps.referenceSheetName(ref, context)
          if (!address || !sheetName) {
            return deps.stackScalar(deps.error(ErrorCode.Value))
          }
          return deps.stackScalar(deps.stringValue(deps.cellTypeCode(context.resolveCell(sheetName, address))))
        }
        case 'filename':
          return deps.stackScalar(deps.stringValue(''))
        default:
          return deps.stackScalar(deps.error(ErrorCode.Value))
      }
    }
    default:
      return undefined
  }
}
