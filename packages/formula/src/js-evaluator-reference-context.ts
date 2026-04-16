import { ValueTag, type CellValue } from '@bilig/protocol'
import { indexToColumn, parseCellAddress } from './addressing.js'
import type { EvaluationContext, ReferenceOperand } from './js-evaluator.js'

export function currentCellReference(context: EvaluationContext): ReferenceOperand | undefined {
  return context.currentAddress ? { kind: 'cell', sheetName: context.sheetName, address: context.currentAddress } : undefined
}

export function referenceSheetName(ref: ReferenceOperand | undefined, context: EvaluationContext): string | undefined {
  return ref?.sheetName ?? context.sheetName
}

export function referenceTopLeftAddress(ref: ReferenceOperand | undefined): string | undefined {
  if (!ref) {
    return undefined
  }
  switch (ref.kind) {
    case 'cell':
    case 'row':
    case 'col':
      return ref.address
    case 'range':
      return ref.start
  }
}

export function referenceRowNumber(ref: ReferenceOperand | undefined, context: EvaluationContext): number | undefined {
  const target = ref ?? currentCellReference(context)
  if (!target) {
    return undefined
  }
  switch (target.kind) {
    case 'cell':
      return parseCellAddress(target.address!, referenceSheetName(target, context)).row + 1
    case 'range':
      if (target.refKind === 'rows') {
        return Number.parseInt(target.start!, 10)
      }
      if (target.refKind === 'cells') {
        return parseCellAddress(target.start!, referenceSheetName(target, context)).row + 1
      }
      return undefined
    case 'row':
      return Number.parseInt(target.address!, 10)
    case 'col':
      return undefined
  }
}

export function referenceColumnNumber(ref: ReferenceOperand | undefined, context: EvaluationContext): number | undefined {
  const target = ref ?? currentCellReference(context)
  if (!target) {
    return undefined
  }
  switch (target.kind) {
    case 'cell':
      return parseCellAddress(target.address!, referenceSheetName(target, context)).col + 1
    case 'range':
      if (target.refKind === 'cols') {
        return parseCellAddress(`${target.start!}1`, referenceSheetName(target, context)).col + 1
      }
      if (target.refKind === 'cells') {
        return parseCellAddress(target.start!, referenceSheetName(target, context)).col + 1
      }
      return undefined
    case 'row':
      return undefined
    case 'col':
      return parseCellAddress(`${target.address!}1`, referenceSheetName(target, context)).col + 1
  }
}

export function absoluteAddress(ref: ReferenceOperand | undefined, context: EvaluationContext): string | undefined {
  const row = referenceRowNumber(ref, context)
  const col = referenceColumnNumber(ref, context)
  return row === undefined || col === undefined ? undefined : `$${indexToColumn(col - 1)}$${row}`
}

export function cellTypeCode(value: CellValue): string {
  switch (value.tag) {
    case ValueTag.Empty:
      return 'b'
    case ValueTag.String:
      return 'l'
    case ValueTag.Number:
    case ValueTag.Boolean:
    case ValueTag.Error:
      return 'v'
  }
}

export function sheetNames(context: EvaluationContext): string[] {
  return context.listSheetNames?.() ?? [context.sheetName]
}

export function sheetIndexByName(name: string, context: EvaluationContext): number | undefined {
  const index = sheetNames(context).findIndex((sheetName) => sheetName.trim().toUpperCase() === name.trim().toUpperCase())
  return index === -1 ? undefined : index + 1
}
