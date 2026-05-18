import { ErrorCode, ValueTag } from '@bilig/protocol'
import type { EngineRuntimeState, RuntimeDirectScalarDescriptor, RuntimeDirectScalarOperand, U32 } from '../runtime-state.js'
import { addEngineCounter } from '../../perf/engine-counters.js'
import { CellFlags } from '../../cell-store.js'

const OP_ADD = 1
const OP_SUB = 2
const OP_MUL = 3
const OP_DIV = 4
const OP_ABS = 5
const BATCH_REF_NONE = 0xffffffff
export const MIN_INITIAL_NATIVE_DIRECT_SCALAR_BATCH_SIZE = 1024
export const MAX_INITIAL_NATIVE_DIRECT_SCALAR_BATCH_SIZE = 2500

interface InitialNativeDirectScalarBatchState {
  readonly workbook: EngineRuntimeState['workbook']
  readonly wasm: EngineRuntimeState['wasm']
  readonly counters: EngineRuntimeState['counters']
}

export interface InitialNativeDirectScalarBatch {
  readonly count: number
  readonly add: (
    prepared: { readonly cellIndex: number; readonly sheetId: number; readonly col: number },
    directScalar: RuntimeDirectScalarDescriptor,
  ) => boolean
  readonly evaluate: () => U32 | undefined
}

export function createInitialNativeDirectScalarBatch(args: {
  readonly state: InitialNativeDirectScalarBatchState
  readonly capacity: number
}): InitialNativeDirectScalarBatch {
  const targets = new Uint32Array(args.capacity)
  const operators = new Uint8Array(args.capacity)
  const leftBatchRefs = new Uint32Array(args.capacity)
  const leftTags = new Uint8Array(args.capacity)
  const leftValues = new Float64Array(args.capacity)
  const leftErrors = new Uint16Array(args.capacity)
  const rightBatchRefs = new Uint32Array(args.capacity)
  const rightTags = new Uint8Array(args.capacity)
  const rightValues = new Float64Array(args.capacity)
  const rightErrors = new Uint16Array(args.capacity)
  const resultOffsets = new Float64Array(args.capacity)
  const outTags = new Uint8Array(args.capacity)
  const outNumbers = new Float64Array(args.capacity)
  const outErrors = new Uint16Array(args.capacity)
  const targetOrdinalByCellIndex = new Map<number, number>()
  const columnsBySheetId = new Map<number, Set<number>>()
  let count = 0

  const noteColumn = (sheetId: number, col: number): void => {
    let columns = columnsBySheetId.get(sheetId)
    if (!columns) {
      columns = new Set()
      columnsBySheetId.set(sheetId, columns)
    }
    columns.add(col)
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
        batchRefs[index] = BATCH_REF_NONE
        tags[index] = (cellStore.tags[operand.cellIndex] as ValueTag | undefined) ?? ValueTag.Empty
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
    add(prepared, directScalar) {
      if (count >= args.capacity) {
        return false
      }
      const index = count
      targets[index] = prepared.cellIndex
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
      } else {
        if (
          !writeOperand(index, directScalar.left, leftBatchRefs, leftTags, leftValues, leftErrors) ||
          !writeOperand(index, directScalar.right, rightBatchRefs, rightTags, rightValues, rightErrors)
        ) {
          return false
        }
      }
      targetOrdinalByCellIndex.set(prepared.cellIndex, index)
      noteColumn(prepared.sheetId, prepared.col)
      count += 1
      return true
    },
    evaluate() {
      if (count === 0 || !args.state.wasm.initSyncIfPossible()) {
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
      columnsBySheetId.forEach((columns, sheetId) => {
        args.state.workbook.notifyColumnsWritten(sheetId, Uint32Array.from([...columns].toSorted((left, right) => left - right)))
      })
      addEngineCounter(args.state.counters, 'nativeDirectScalarInitialEvaluations', count)
      return targetView
    },
  }
}
