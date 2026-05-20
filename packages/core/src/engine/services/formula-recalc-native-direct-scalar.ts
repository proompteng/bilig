import { ErrorCode, ValueTag } from '@bilig/protocol'
import { CellFlags } from '../../cell-store.js'
import { addEngineCounter } from '../../perf/engine-counters.js'
import type {
  EngineRuntimeState,
  RuntimeDirectScalarDescriptor,
  RuntimeDirectScalarOperand,
  RuntimeFormula,
  U32,
} from '../runtime-state.js'
import {
  createWrittenColumnTracker,
  markWrittenColumn,
  materializeWrittenColumns,
  type WrittenColumnTracker,
} from '../../written-column-tracker.js'

const OP_ADD = 1
const OP_SUB = 2
const OP_MUL = 3
const OP_DIV = 4
const OP_ABS = 5
const BATCH_REF_NONE = 0xffffffff

export const MIN_RECALC_NATIVE_DIRECT_SCALAR_BATCH_SIZE = 64
export const MAX_RECALC_NATIVE_DIRECT_SCALAR_BATCH_SIZE = 4096

type RecalcNativeDirectScalarState = Pick<EngineRuntimeState, 'workbook' | 'wasm' | 'counters'>

export interface RecalcNativeDirectScalarBatch {
  readonly count: number
  readonly add: (cellIndex: number, formula: RuntimeFormula) => boolean
  readonly evaluate: () => U32 | undefined
  readonly reset: () => void
}

export function createRecalcNativeDirectScalarBatch(args: {
  readonly state: RecalcNativeDirectScalarState
  readonly capacity: number
}): RecalcNativeDirectScalarBatch {
  const capacity = Math.max(1, args.capacity)
  const targets = new Uint32Array(capacity)
  const operators = new Uint8Array(capacity)
  const leftBatchRefs = new Uint32Array(capacity)
  const leftTags = new Uint8Array(capacity)
  const leftValues = new Float64Array(capacity)
  const leftErrors = new Uint16Array(capacity)
  const rightBatchRefs = new Uint32Array(capacity)
  const rightTags = new Uint8Array(capacity)
  const rightValues = new Float64Array(capacity)
  const rightErrors = new Uint16Array(capacity)
  const resultOffsets = new Float64Array(capacity)
  const outTags = new Uint8Array(capacity)
  const outNumbers = new Float64Array(capacity)
  const outErrors = new Uint16Array(capacity)
  const targetOrdinalByCellIndex = new Map<number, number>()
  const columnsBySheetId = new Map<number, WrittenColumnTracker>()
  let count = 0

  const noteColumn = (sheetId: number, col: number): void => {
    let tracker = columnsBySheetId.get(sheetId)
    if (!tracker) {
      tracker = createWrittenColumnTracker()
      columnsBySheetId.set(sheetId, tracker)
    }
    markWrittenColumn(tracker, col)
  }

  const writeOperand = (
    index: number,
    operand: RuntimeDirectScalarOperand,
    batchRefs: Uint32Array,
    tags: Uint8Array,
    values: Float64Array,
    errors: Uint16Array,
  ): boolean => {
    switch (operand.kind) {
      case 'cell': {
        const batchRef = targetOrdinalByCellIndex.get(operand.cellIndex)
        if (batchRef !== undefined) {
          batchRefs[index] = batchRef
          tags[index] = ValueTag.Empty
          values[index] = 0
          errors[index] = ErrorCode.None
          return true
        }
        const cellStore = args.state.workbook.cellStore
        const tag = (cellStore.tags[operand.cellIndex] as ValueTag | undefined) ?? ValueTag.Empty
        if (tag === ValueTag.String) {
          return false
        }
        batchRefs[index] = BATCH_REF_NONE
        tags[index] = tag
        values[index] = cellStore.numbers[operand.cellIndex] ?? 0
        errors[index] = (cellStore.errors[operand.cellIndex] as ErrorCode | undefined) ?? ErrorCode.None
        return true
      }
      case 'literal-number':
        batchRefs[index] = BATCH_REF_NONE
        tags[index] = ValueTag.Number
        values[index] = operand.value
        errors[index] = ErrorCode.None
        return true
      case 'error':
        batchRefs[index] = BATCH_REF_NONE
        tags[index] = ValueTag.Error
        values[index] = 0
        errors[index] = operand.code
        return true
    }
  }

  const operatorCode = (directScalar: RuntimeDirectScalarDescriptor): number => {
    if (directScalar.kind === 'abs') {
      return OP_ABS
    }
    switch (directScalar.operator) {
      case '+':
        return OP_ADD
      case '-':
        return OP_SUB
      case '*':
        return OP_MUL
      case '/':
        return OP_DIV
    }
  }

  return {
    get count() {
      return count
    },
    add(cellIndex, formula) {
      const directScalar = formula.directScalar
      if (count >= capacity || directScalar === undefined || formula.compiled.producesSpill) {
        return false
      }
      const cellStore = args.state.workbook.cellStore
      const sheetId = cellStore.sheetIds[cellIndex]
      const col = cellStore.cols[cellIndex]
      if (sheetId === undefined || col === undefined) {
        return false
      }
      const index = count
      operators[index] = operatorCode(directScalar)
      resultOffsets[index] = directScalar.kind === 'binary' ? (directScalar.resultOffset ?? 0) : 0
      if (directScalar.kind === 'abs') {
        if (!writeOperand(index, directScalar.operand, leftBatchRefs, leftTags, leftValues, leftErrors)) {
          return false
        }
        rightBatchRefs[index] = BATCH_REF_NONE
        rightTags[index] = ValueTag.Number
        rightValues[index] = 0
        rightErrors[index] = ErrorCode.None
      } else if (
        !writeOperand(index, directScalar.left, leftBatchRefs, leftTags, leftValues, leftErrors) ||
        !writeOperand(index, directScalar.right, rightBatchRefs, rightTags, rightValues, rightErrors)
      ) {
        return false
      }
      targets[index] = cellIndex
      targetOrdinalByCellIndex.set(cellIndex, index)
      noteColumn(sheetId, col)
      count += 1
      return true
    },
    evaluate() {
      if (count < MIN_RECALC_NATIVE_DIRECT_SCALAR_BATCH_SIZE || !args.state.wasm.initSyncIfPossible()) {
        return undefined
      }
      const targetView = targets.subarray(0, count)
      if (
        !args.state.wasm.evalDirectScalarValueBatch({
          operators: operators.subarray(0, count),
          leftBatchRefs: leftBatchRefs.subarray(0, count),
          leftTags: leftTags.subarray(0, count),
          leftValues: leftValues.subarray(0, count),
          leftErrors: leftErrors.subarray(0, count),
          rightBatchRefs: rightBatchRefs.subarray(0, count),
          rightTags: rightTags.subarray(0, count),
          rightValues: rightValues.subarray(0, count),
          rightErrors: rightErrors.subarray(0, count),
          resultOffsets: resultOffsets.subarray(0, count),
          outTags: outTags.subarray(0, count),
          outNumbers: outNumbers.subarray(0, count),
          outErrors: outErrors.subarray(0, count),
        })
      ) {
        return undefined
      }
      const cellStore = args.state.workbook.cellStore
      for (let index = 0; index < count; index += 1) {
        const cellIndex = targets[index]!
        const tag = (outTags[index] as ValueTag | undefined) ?? ValueTag.Empty
        cellStore.flags[cellIndex] = (cellStore.flags[cellIndex] ?? 0) & ~(CellFlags.SpillChild | CellFlags.PivotOutput)
        cellStore.tags[cellIndex] = tag
        cellStore.errors[cellIndex] =
          tag === ValueTag.Error ? ((outErrors[index] as ErrorCode | undefined) ?? ErrorCode.None) : ErrorCode.None
        cellStore.stringIds[cellIndex] = 0
        cellStore.numbers[cellIndex] = tag === ValueTag.Number || tag === ValueTag.Boolean ? (outNumbers[index] ?? 0) : 0
        cellStore.versions[cellIndex] = (cellStore.versions[cellIndex] ?? 0) + 1
      }
      columnsBySheetId.forEach((tracker, sheetId) => {
        args.state.workbook.notifyColumnsWritten(sheetId, materializeWrittenColumns(tracker))
      })
      addEngineCounter(args.state.counters, 'nativeDirectScalarRecalcEvaluations', count)
      return targetView
    },
    reset() {
      count = 0
      targetOrdinalByCellIndex.clear()
      columnsBySheetId.clear()
    },
  }
}
