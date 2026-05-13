import { ErrorCode, ValueTag } from '@bilig/protocol'
import { parseNumericText } from '@bilig/formula'
import { CellFlags } from '../../cell-store.js'
import type { EngineRuntimeState, RuntimeDirectScalarDescriptor, RuntimeDirectScalarOperand } from '../runtime-state.js'
import { directScalarCellNumber, directScalarValueNumber } from './direct-scalar-helpers.js'
import type { DirectScalarCurrentOperand, PendingNumericCellValues } from './direct-formula-index-collection.js'

export function createOperationDirectFormulaValues(args: { readonly state: Pick<EngineRuntimeState, 'workbook' | 'strings'> }) {
  const readDirectScalarCellNumber = (cellIndex: number): number | undefined =>
    directScalarValueNumber(args.state.workbook.cellStore.getValue(cellIndex, (id) => args.state.strings.get(id)))

  const directScalarCellNumericValue = (cellIndex: number | undefined): number | undefined =>
    directScalarCellNumber(args.state.workbook.cellStore, cellIndex)

  const readCellTextNumber = (cellIndex: number): number | undefined => {
    const stringId = args.state.workbook.cellStore.stringIds[cellIndex] ?? 0
    const text = args.state.strings.get(stringId)
    return text.trim() === '' ? 0 : parseNumericText(text)
  }

  const directScalarCurrentResultMatchesCell = (cellIndex: number, result: DirectScalarCurrentOperand): boolean => {
    const cellStore = args.state.workbook.cellStore
    const currentTag = (cellStore.tags[cellIndex] as ValueTag | undefined) ?? ValueTag.Empty
    if (result.kind === 'number') {
      return currentTag === ValueTag.Number && Object.is(cellStore.numbers[cellIndex] ?? 0, result.value)
    }
    return currentTag === ValueTag.Error && ((cellStore.errors[cellIndex] as ErrorCode | undefined) ?? ErrorCode.None) === result.code
  }

  const directScalarNumericResultMatchesCell = (cellIndex: number, result: number): boolean => {
    const cellStore = args.state.workbook.cellStore
    return cellStore.tags[cellIndex] === ValueTag.Number && Object.is(cellStore.numbers[cellIndex] ?? 0, result)
  }

  const applyDirectFormulaCurrentResult = (cellIndex: number, result: DirectScalarCurrentOperand): boolean => {
    const cellStore = args.state.workbook.cellStore
    const beforeTag = (cellStore.tags[cellIndex] as ValueTag | undefined) ?? ValueTag.Empty
    const beforeNumber = cellStore.numbers[cellIndex] ?? 0
    const beforeError = (cellStore.errors[cellIndex] as ErrorCode | undefined) ?? ErrorCode.None
    cellStore.flags[cellIndex] = (cellStore.flags[cellIndex] ?? 0) & ~(CellFlags.SpillChild | CellFlags.PivotOutput)
    cellStore.versions[cellIndex] = (cellStore.versions[cellIndex] ?? 0) + 1
    cellStore.stringIds[cellIndex] = 0
    if (result.kind === 'number') {
      cellStore.tags[cellIndex] = ValueTag.Number
      cellStore.numbers[cellIndex] = result.value
      cellStore.errors[cellIndex] = ErrorCode.None
      if (cellStore.onSetValue) {
        cellStore.onSetValue(cellIndex)
      } else if (beforeTag !== ValueTag.Number || !Object.is(beforeNumber, result.value)) {
        args.state.workbook.notifyCellValueWritten(cellIndex)
      }
      return true
    }
    cellStore.tags[cellIndex] = ValueTag.Error
    cellStore.numbers[cellIndex] = 0
    cellStore.errors[cellIndex] = result.code
    if (cellStore.onSetValue) {
      cellStore.onSetValue(cellIndex)
    } else if (beforeTag !== ValueTag.Error || beforeError !== result.code) {
      args.state.workbook.notifyCellValueWritten(cellIndex)
    }
    return true
  }

  const applyDirectFormulaNumericResult = (cellIndex: number, value: number): void => {
    const cellStore = args.state.workbook.cellStore
    cellStore.flags[cellIndex] = (cellStore.flags[cellIndex] ?? 0) & ~(CellFlags.SpillChild | CellFlags.PivotOutput)
    cellStore.versions[cellIndex] = (cellStore.versions[cellIndex] ?? 0) + 1
    cellStore.stringIds[cellIndex] = 0
    cellStore.tags[cellIndex] = ValueTag.Number
    cellStore.numbers[cellIndex] = value
    cellStore.errors[cellIndex] = ErrorCode.None
    if (cellStore.onSetValue) {
      cellStore.onSetValue(cellIndex)
    } else {
      args.state.workbook.notifyCellValueWritten(cellIndex)
    }
  }

  const applyTerminalDirectFormulaNumericResult = (cellIndex: number, value: number): void => {
    const cellStore = args.state.workbook.cellStore
    cellStore.flags[cellIndex] = (cellStore.flags[cellIndex] ?? 0) & ~(CellFlags.SpillChild | CellFlags.PivotOutput)
    cellStore.versions[cellIndex] = (cellStore.versions[cellIndex] ?? 0) + 1
    cellStore.stringIds[cellIndex] = 0
    cellStore.tags[cellIndex] = ValueTag.Number
    cellStore.numbers[cellIndex] = value
    cellStore.errors[cellIndex] = ErrorCode.None
  }

  const writeNumericLiteralToCellStore = (cellIndex: number, value: number): void => {
    const cellStore = args.state.workbook.cellStore
    cellStore.flags[cellIndex] = (cellStore.flags[cellIndex] ?? 0) & ~(CellFlags.SpillChild | CellFlags.PivotOutput)
    cellStore.versions[cellIndex] = (cellStore.versions[cellIndex] ?? 0) + 1
    cellStore.stringIds[cellIndex] = 0
    cellStore.tags[cellIndex] = ValueTag.Number
    cellStore.numbers[cellIndex] = value
    cellStore.errors[cellIndex] = ErrorCode.None
  }

  const readDirectScalarCurrentOperand = (operand: RuntimeDirectScalarOperand): DirectScalarCurrentOperand | undefined => {
    switch (operand.kind) {
      case 'literal-number':
        return { kind: 'number', value: operand.value }
      case 'error':
        return { kind: 'error', code: operand.code }
      case 'cell': {
        const cellStore = args.state.workbook.cellStore
        const tag = (cellStore.tags[operand.cellIndex] as ValueTag | undefined) ?? ValueTag.Empty
        switch (tag) {
          case ValueTag.Number:
            return { kind: 'number', value: cellStore.numbers[operand.cellIndex] ?? 0 }
          case ValueTag.Boolean:
            return { kind: 'number', value: (cellStore.numbers[operand.cellIndex] ?? 0) !== 0 ? 1 : 0 }
          case ValueTag.Empty:
            return { kind: 'number', value: 0 }
          case ValueTag.Error:
            return { kind: 'error', code: (cellStore.errors[operand.cellIndex] as ErrorCode | undefined) ?? ErrorCode.None }
          case ValueTag.String: {
            const numeric = readCellTextNumber(operand.cellIndex)
            return numeric === undefined ? { kind: 'error', code: ErrorCode.Value } : { kind: 'number', value: numeric }
          }
        }
      }
    }
  }

  const evaluateDirectScalarCurrentValue = (directScalar: RuntimeDirectScalarDescriptor): DirectScalarCurrentOperand | undefined => {
    if (directScalar.kind === 'abs') {
      const operand = readDirectScalarCurrentOperand(directScalar.operand)
      return operand?.kind === 'number' ? { kind: 'number', value: Math.abs(operand.value) } : operand
    }
    const left = readDirectScalarCurrentOperand(directScalar.left)
    const right = readDirectScalarCurrentOperand(directScalar.right)
    if (!left || !right) {
      return undefined
    }
    if (left.kind === 'error') {
      return left
    }
    if (right.kind === 'error') {
      return right
    }
    let result: number
    switch (directScalar.operator) {
      case '+':
        result = left.value + right.value
        break
      case '-':
        result = left.value - right.value
        break
      case '*':
        result = left.value * right.value
        break
      case '/':
        if (right.value === 0) {
          return { kind: 'error', code: ErrorCode.Div0 }
        }
        result = left.value / right.value
        break
    }
    return { kind: 'number', value: result + (directScalar.resultOffset ?? 0) }
  }

  const applyDirectScalarCurrentValue = (cellIndex: number, directScalar: RuntimeDirectScalarDescriptor): boolean => {
    const result = evaluateDirectScalarCurrentValue(directScalar)
    if (!result) {
      return false
    }
    return applyDirectFormulaCurrentResult(cellIndex, result)
  }

  const tryEvaluateDirectScalarWithPendingNumbers = (
    directScalar: RuntimeDirectScalarDescriptor,
    pendingNumbers: PendingNumericCellValues,
  ): DirectScalarCurrentOperand | undefined => {
    const readOperand = (operand: RuntimeDirectScalarOperand): DirectScalarCurrentOperand | undefined => {
      switch (operand.kind) {
        case 'literal-number':
          return { kind: 'number', value: operand.value }
        case 'error':
          return { kind: 'error', code: operand.code }
        case 'cell': {
          const pending = pendingNumbers.get(operand.cellIndex)
          if (pending !== undefined) {
            return { kind: 'number', value: pending }
          }
          const cellStore = args.state.workbook.cellStore
          const tag = (cellStore.tags[operand.cellIndex] as ValueTag | undefined) ?? ValueTag.Empty
          switch (tag) {
            case ValueTag.Number:
              return { kind: 'number', value: cellStore.numbers[operand.cellIndex] ?? 0 }
            case ValueTag.Boolean:
              return { kind: 'number', value: (cellStore.numbers[operand.cellIndex] ?? 0) !== 0 ? 1 : 0 }
            case ValueTag.Empty:
              return { kind: 'number', value: 0 }
            case ValueTag.Error:
              return { kind: 'error', code: (cellStore.errors[operand.cellIndex] as ErrorCode | undefined) ?? ErrorCode.None }
            case ValueTag.String: {
              const numeric = readCellTextNumber(operand.cellIndex)
              return numeric === undefined ? { kind: 'error', code: ErrorCode.Value } : { kind: 'number', value: numeric }
            }
          }
        }
      }
    }

    if (directScalar.kind === 'abs') {
      const operand = readOperand(directScalar.operand)
      return operand?.kind === 'number' ? { kind: 'number', value: Math.abs(operand.value) } : operand
    }
    const left = readOperand(directScalar.left)
    const right = readOperand(directScalar.right)
    if (!left || !right) {
      return undefined
    }
    if (left.kind === 'error') {
      return left
    }
    if (right.kind === 'error') {
      return right
    }
    let result: number
    switch (directScalar.operator) {
      case '+':
        result = left.value + right.value
        break
      case '-':
        result = left.value - right.value
        break
      case '*':
        result = left.value * right.value
        break
      case '/':
        if (right.value === 0) {
          return { kind: 'error', code: ErrorCode.Div0 }
        }
        result = left.value / right.value
        break
    }
    return { kind: 'number', value: result + (directScalar.resultOffset ?? 0) }
  }

  const tryEvaluateDirectScalarNumericWithPendingNumbers = (
    directScalar: RuntimeDirectScalarDescriptor,
    pendingNumbers: PendingNumericCellValues,
  ): number | undefined => {
    const readOperand = (operand: RuntimeDirectScalarOperand): number | undefined => {
      switch (operand.kind) {
        case 'literal-number':
          return operand.value
        case 'error':
          return undefined
        case 'cell': {
          const pending = pendingNumbers.get(operand.cellIndex)
          if (pending !== undefined) {
            return pending
          }
          const cellStore = args.state.workbook.cellStore
          switch ((cellStore.tags[operand.cellIndex] as ValueTag | undefined) ?? ValueTag.Empty) {
            case ValueTag.Number:
              return cellStore.numbers[operand.cellIndex] ?? 0
            case ValueTag.Boolean:
              return (cellStore.numbers[operand.cellIndex] ?? 0) !== 0 ? 1 : 0
            case ValueTag.Empty:
              return 0
            case ValueTag.Error:
            case ValueTag.String:
              return undefined
          }
        }
      }
    }

    if (directScalar.kind === 'abs') {
      const operand = readOperand(directScalar.operand)
      return operand === undefined ? undefined : Math.abs(operand)
    }
    const left = readOperand(directScalar.left)
    const right = readOperand(directScalar.right)
    if (left === undefined || right === undefined) {
      return undefined
    }
    let result: number
    switch (directScalar.operator) {
      case '+':
        result = left + right
        break
      case '-':
        result = left - right
        break
      case '*':
        result = left * right
        break
      case '/':
        if (right === 0) {
          return undefined
        }
        result = left / right
        break
    }
    return result + (directScalar.resultOffset ?? 0)
  }

  return {
    readDirectScalarCellNumber,
    directScalarCellNumericValue,
    directScalarCurrentResultMatchesCell,
    directScalarNumericResultMatchesCell,
    applyDirectFormulaCurrentResult,
    applyDirectFormulaNumericResult,
    applyTerminalDirectFormulaNumericResult,
    writeNumericLiteralToCellStore,
    evaluateDirectScalarCurrentValue,
    applyDirectScalarCurrentValue,
    tryEvaluateDirectScalarWithPendingNumbers,
    tryEvaluateDirectScalarNumericWithPendingNumbers,
  } as const
}
