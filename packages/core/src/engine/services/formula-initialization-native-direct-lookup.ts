import { ErrorCode, ValueTag } from '@bilig/protocol'
import { normalizeExactLookupNumber } from '@bilig/formula'
import { CellFlags } from '../../cell-store.js'
import { addEngineCounter, type EngineCounterKey } from '../../perf/engine-counters.js'
import type { EngineRuntimeState, RuntimeDirectLookupDescriptor, U32 } from '../runtime-state.js'
import { directLookupVersionMatches } from './direct-lookup-helpers.js'
import {
  createWrittenColumnTracker,
  markWrittenColumn,
  materializeWrittenColumns,
  type WrittenColumnTracker,
} from '../../written-column-tracker.js'

const LOOKUP_KIND_EXACT_UNIFORM_NUMERIC = 1
const LOOKUP_KIND_APPROXIMATE_UNIFORM_NUMERIC = 2
const MATCH_MODE_ASCENDING = 1
const MATCH_MODE_DESCENDING = 2

export const MIN_INITIAL_NATIVE_DIRECT_LOOKUP_BATCH_SIZE = 256
export const MAX_INITIAL_NATIVE_DIRECT_LOOKUP_BATCH_SIZE = 65_536
export const MIN_RECALC_NATIVE_DIRECT_LOOKUP_BATCH_SIZE = 64
export const MAX_RECALC_NATIVE_DIRECT_LOOKUP_BATCH_SIZE = 4096

interface InitialNativeDirectLookupBatchState {
  readonly workbook: EngineRuntimeState['workbook']
  readonly wasm: EngineRuntimeState['wasm']
  readonly counters: EngineRuntimeState['counters']
}

type NativeDirectLookup = Extract<RuntimeDirectLookupDescriptor, { kind: 'exact-uniform-numeric' | 'approximate-uniform-numeric' }>

export interface NativeDirectLookupBatch {
  readonly count: number
  readonly add: (
    prepared: { readonly cellIndex: number; readonly sheetId: number; readonly col: number },
    directLookup: RuntimeDirectLookupDescriptor,
  ) => boolean
  readonly evaluate: () => U32 | undefined
  readonly reset: () => void
}

export function createInitialNativeDirectLookupBatch(args: {
  readonly state: InitialNativeDirectLookupBatchState
  readonly capacity: number
}): NativeDirectLookupBatch {
  return createNativeDirectLookupBatch({
    ...args,
    counterName: 'nativeDirectLookupInitialEvaluations',
    minBatchSize: 1,
  })
}

export function createRecalcNativeDirectLookupBatch(args: {
  readonly state: InitialNativeDirectLookupBatchState
  readonly capacity: number
}): NativeDirectLookupBatch {
  return createNativeDirectLookupBatch({
    ...args,
    counterName: 'nativeDirectLookupRecalcEvaluations',
    minBatchSize: MIN_RECALC_NATIVE_DIRECT_LOOKUP_BATCH_SIZE,
  })
}

function createNativeDirectLookupBatch(args: {
  readonly state: InitialNativeDirectLookupBatchState
  readonly capacity: number
  readonly counterName: Extract<EngineCounterKey, 'nativeDirectLookupInitialEvaluations' | 'nativeDirectLookupRecalcEvaluations'>
  readonly minBatchSize: number
}): NativeDirectLookupBatch {
  const targets = new Uint32Array(args.capacity)
  const kinds = new Uint8Array(args.capacity)
  const matchModes = new Uint8Array(args.capacity)
  const starts = new Float64Array(args.capacity)
  const steps = new Float64Array(args.capacity)
  const lengths = new Uint32Array(args.capacity)
  const repeatedRunLengths = new Uint32Array(args.capacity)
  const lookupTags = new Uint8Array(args.capacity)
  const lookupNumbers = new Float64Array(args.capacity)
  const outTags = new Uint8Array(args.capacity)
  const outNumbers = new Float64Array(args.capacity)
  const outErrors = new Uint16Array(args.capacity)
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

  const writeLookupValue = (index: number, directLookup: NativeDirectLookup): boolean => {
    const cellStore = args.state.workbook.cellStore
    const tag = (cellStore.tags[directLookup.operandCellIndex] as ValueTag | undefined) ?? ValueTag.Empty
    if (directLookup.kind === 'exact-uniform-numeric') {
      if (tag !== ValueTag.Number) {
        return false
      }
      const normalized = normalizeExactLookupNumber(cellStore.numbers[directLookup.operandCellIndex] ?? 0)
      if (!Number.isFinite(normalized)) {
        return false
      }
      lookupTags[index] = ValueTag.Number
      lookupNumbers[index] = normalized
      return true
    }
    switch (tag) {
      case ValueTag.Empty:
        lookupTags[index] = ValueTag.Empty
        lookupNumbers[index] = 0
        return true
      case ValueTag.Number:
        lookupTags[index] = ValueTag.Number
        lookupNumbers[index] = Object.is(cellStore.numbers[directLookup.operandCellIndex] ?? 0, -0)
          ? 0
          : (cellStore.numbers[directLookup.operandCellIndex] ?? 0)
        return true
      case ValueTag.Boolean:
        lookupTags[index] = ValueTag.Boolean
        lookupNumbers[index] = (cellStore.numbers[directLookup.operandCellIndex] ?? 0) !== 0 ? 1 : 0
        return true
      case ValueTag.String:
      case ValueTag.Error:
        return false
    }
  }

  const canUseNativeLookupShape = (directLookup: NativeDirectLookup): boolean => {
    if (directLookup.tailPatch !== undefined || directLookup.length <= 0 || directLookup.step === 0) {
      return false
    }
    const lookupSheet = args.state.workbook.getSheetById(directLookup.sheetId)
    if (!directLookupVersionMatches(lookupSheet, directLookup)) {
      return false
    }
    if (directLookup.kind === 'exact-uniform-numeric') {
      const start = normalizeExactLookupNumber(directLookup.start)
      const step = normalizeExactLookupNumber(directLookup.step)
      return Number.isFinite(start) && Number.isFinite(step) && Number.isInteger(start) && Number.isInteger(step)
    }
    if (directLookup.matchMode === 1) {
      return directLookup.step > 0
    }
    return directLookup.step < 0
  }

  return {
    get count() {
      return count
    },
    add(prepared, directLookup) {
      if (
        count >= args.capacity ||
        (directLookup.kind !== 'exact-uniform-numeric' && directLookup.kind !== 'approximate-uniform-numeric') ||
        !canUseNativeLookupShape(directLookup)
      ) {
        return false
      }
      const index = count
      if (!writeLookupValue(index, directLookup)) {
        return false
      }
      targets[index] = prepared.cellIndex
      kinds[index] =
        directLookup.kind === 'exact-uniform-numeric' ? LOOKUP_KIND_EXACT_UNIFORM_NUMERIC : LOOKUP_KIND_APPROXIMATE_UNIFORM_NUMERIC
      matchModes[index] =
        directLookup.kind === 'approximate-uniform-numeric' && directLookup.matchMode === -1 ? MATCH_MODE_DESCENDING : MATCH_MODE_ASCENDING
      starts[index] = directLookup.kind === 'exact-uniform-numeric' ? normalizeExactLookupNumber(directLookup.start) : directLookup.start
      steps[index] = directLookup.kind === 'exact-uniform-numeric' ? normalizeExactLookupNumber(directLookup.step) : directLookup.step
      lengths[index] = directLookup.length
      repeatedRunLengths[index] = directLookup.kind === 'approximate-uniform-numeric' ? (directLookup.repeatedRunLength ?? 0) : 0
      noteColumn(prepared.sheetId, prepared.col)
      count += 1
      return true
    },
    evaluate() {
      if (count < args.minBatchSize || !args.state.wasm.initSyncIfPossible()) {
        return undefined
      }
      const targetView = targets.subarray(0, count)
      if (
        !args.state.wasm.evalUniformNumericLookupBatch({
          kinds: kinds.subarray(0, count),
          matchModes: matchModes.subarray(0, count),
          starts: starts.subarray(0, count),
          steps: steps.subarray(0, count),
          lengths: lengths.subarray(0, count),
          repeatedRunLengths: repeatedRunLengths.subarray(0, count),
          lookupTags: lookupTags.subarray(0, count),
          lookupNumbers: lookupNumbers.subarray(0, count),
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
      addEngineCounter(args.state.counters, args.counterName, count)
      return targetView
    },
    reset() {
      count = 0
      columnsBySheetId.clear()
    },
  }
}
