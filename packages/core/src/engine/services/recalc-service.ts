import { Effect } from 'effect'
import {
  ErrorCode,
  FormulaMode,
  ValueTag,
  type CellSnapshot,
  type CellValue,
  type EngineChangedCell,
  type EngineEvent,
  type WorkbookSnapshot,
} from '@bilig/protocol'
import { makeCellKey } from '../../workbook-store.js'
import { CellFlags } from '../../cell-store.js'
import { errorValue } from '../../engine-value-utils.js'
import type { EngineRuntimeState, RecalcVolatileState, RuntimeFormula, SpillMaterialization, U32 } from '../runtime-state.js'
import { EngineRecalcError } from '../errors.js'
import type { WorkbookPivotRecord } from '../../workbook-store.js'
import { parseCellAddress, utcDateToExcelSerial } from '@bilig/formula'
import type { EngineDirtyFrontierSchedulerService } from './dirty-frontier-scheduler-service.js'
import type { EnginePatch } from '../../patches/patch-types.js'

const TRACKED_CELL_PATCH_LIMIT = 2_048

export interface DirtyRegion {
  readonly sheetName: string
  readonly rowStart: number
  readonly rowEnd: number
  readonly colStart: number
  readonly colEnd: number
}

export interface EngineRecalcService {
  readonly recalculateNow: () => Effect.Effect<number[], EngineRecalcError>
  readonly recalculateDirty: (dirtyRegions: ReadonlyArray<DirtyRegion>) => Effect.Effect<number[], EngineRecalcError>
  readonly recalculateDifferential: () => Effect.Effect<{ js: CellSnapshot[]; wasm: CellSnapshot[]; drift: string[] }, EngineRecalcError>
  readonly recalculatePreordered: (
    changedRoots: readonly number[] | U32,
    orderedFormulaCellIndices: readonly number[] | U32,
    orderedFormulaCount: number,
    kernelSyncRoots?: readonly number[] | U32,
  ) => Effect.Effect<U32, EngineRecalcError>
  readonly recalculate: (
    changedRoots: readonly number[] | U32,
    kernelSyncRoots?: readonly number[] | U32,
  ) => Effect.Effect<U32, EngineRecalcError>
  readonly reconcilePivotOutputs: (baseChanged: U32, forceAllPivots?: boolean) => Effect.Effect<U32, EngineRecalcError>
  readonly recalculatePreorderedNowSync: (
    changedRoots: readonly number[] | U32,
    orderedFormulaCellIndices: readonly number[] | U32,
    orderedFormulaCount: number,
    kernelSyncRoots?: readonly number[] | U32,
  ) => U32
  readonly recalculateNowSync: (changedRoots: readonly number[] | U32, kernelSyncRoots?: readonly number[] | U32) => U32
  readonly reconcilePivotOutputsNow: (baseChanged: U32, forceAllPivots?: boolean) => U32
}

function createRecalcVolatileState(now: () => Date): RecalcVolatileState {
  return {
    nowSerial: utcDateToExcelSerial(now()),
    randomValues: [],
    randomCursor: 0,
  }
}

function ensureVolatileRandomValues(state: RecalcVolatileState, count: number, random: () => number): void {
  const needed = state.randomCursor + count - state.randomValues.length
  if (needed <= 0) {
    return
  }
  for (let index = 0; index < needed; index += 1) {
    state.randomValues.push(random())
  }
}

function consumeVolatileRandomValues(state: RecalcVolatileState, count: number, random: () => number): Float64Array {
  ensureVolatileRandomValues(state, count, random)
  const values = state.randomValues.slice(state.randomCursor, state.randomCursor + count)
  state.randomCursor += count
  return Float64Array.from(values)
}

function toOrderedUint32(ordered: readonly number[] | U32, orderedCount: number): U32 {
  if (ordered instanceof Uint32Array) {
    return ordered
  }
  const next = new Uint32Array(orderedCount)
  for (let index = 0; index < orderedCount; index += 1) {
    next[index] = ordered[index] ?? 0
  }
  return next
}

export function createEngineRecalcService(args: {
  readonly state: Pick<
    EngineRuntimeState,
    'workbook' | 'strings' | 'wasm' | 'formulas' | 'ranges' | 'events' | 'getLastMetrics' | 'setLastMetrics'
  >
  readonly getCellByIndex: (cellIndex: number) => CellSnapshot
  readonly exportSnapshot: () => WorkbookSnapshot
  readonly importSnapshot: (snapshot: WorkbookSnapshot) => void
  readonly beginMutationCollection: () => void
  readonly markInputChanged: (cellIndex: number, count: number) => number
  readonly markFormulaChanged: (cellIndex: number, count: number) => number
  readonly markExplicitChanged: (cellIndex: number, count: number) => number
  readonly composeMutationRoots: (changedInputCount: number, formulaChangedCount: number) => U32
  readonly composeEventChanges: (recalculated: U32, explicitChangedCount: number) => U32
  readonly captureChangedCells: (changedCellIndices: readonly number[] | U32) => readonly EngineChangedCell[]
  readonly captureChangedPatches: (
    changedCellIndices: readonly number[] | U32,
    request?: {
      invalidation?: 'cells' | 'full'
      invalidatedRanges?: readonly {
        sheetName: string
        startAddress: string
        endAddress: string
      }[]
      invalidatedRows?: readonly { sheetName: string; startIndex: number; endIndex: number }[]
      invalidatedColumns?: readonly { sheetName: string; startIndex: number; endIndex: number }[]
    },
  ) => readonly EnginePatch[]
  readonly unionChangedSets: (...sets: Array<readonly number[] | U32>) => U32
  readonly composeChangedRootsAndOrdered: (changedRoots: readonly number[] | U32, ordered: U32, orderedCount: number) => U32
  readonly emptyChangedSet: () => U32
  readonly ensureRecalcScratchCapacity: (size: number) => void
  readonly getPendingKernelSync: () => U32
  readonly getDeferredKernelSyncCount: () => number
  readonly setDeferredKernelSyncCount: (next: number) => void
  readonly getDeferredKernelSyncEpoch: () => number
  readonly setDeferredKernelSyncEpoch: (next: number) => void
  readonly getDeferredKernelSyncSeen: () => U32
  readonly getWasmBatch: () => U32
  readonly getChangedInputBuffer: () => U32
  readonly now: () => Date
  readonly random: () => number
  readonly performanceNow: () => number
  readonly dirtyScheduler: EngineDirtyFrontierSchedulerService
  readonly materializeSpill: (cellIndex: number, arrayValue: { values: CellValue[]; rows: number; cols: number }) => SpillMaterialization
  readonly clearOwnedSpill: (cellIndex: number) => number[]
  readonly evaluateDirectLookupFormula: (cellIndex: number) => number[] | undefined
  readonly evaluateUnsupportedFormula: (cellIndex: number) => number[]
  readonly materializePivot: (pivot: WorkbookPivotRecord) => number[]
}): EngineRecalcService {
  const captureTrackedPatchesForCells = (changed: readonly number[] | U32): readonly EnginePatch[] | undefined =>
    changed.length <= TRACKED_CELL_PATCH_LIMIT
      ? args.captureChangedPatches(changed, {
          invalidation: 'cells',
          invalidatedRanges: [],
          invalidatedRows: [],
          invalidatedColumns: [],
        })
      : undefined

  const shouldRefreshPivot = (pivot: WorkbookPivotRecord, changed: readonly number[] | U32): boolean => {
    const ownerSheet = args.state.workbook.getSheet(pivot.source.sheetName)
    if (!ownerSheet) {
      return true
    }
    const ownerStart = parseCellAddress(pivot.source.startAddress, pivot.source.sheetName)
    const ownerEnd = parseCellAddress(pivot.source.endAddress, pivot.source.sheetName)
    for (let index = 0; index < changed.length; index += 1) {
      const cellIndex = changed[index]!
      const sheetId = args.state.workbook.cellStore.sheetIds[cellIndex]
      if (sheetId === undefined || sheetId !== ownerSheet.id) {
        continue
      }
      const position = args.state.workbook.getCellPosition(cellIndex)
      const row = position?.row ?? args.state.workbook.cellStore.rows[cellIndex] ?? -1
      const col = position?.col ?? args.state.workbook.cellStore.cols[cellIndex] ?? -1
      if (row >= ownerStart.row && row <= ownerEnd.row && col >= ownerStart.col && col <= ownerEnd.col) {
        return true
      }
    }
    return false
  }

  const refreshPivotOutputs = (changed: readonly number[] | U32, forceAll: boolean): U32 => {
    const pivots = args.state.workbook.listPivots()
    if (pivots.length === 0 || (!forceAll && changed.length === 0)) {
      return args.emptyChangedSet()
    }

    const changedCellIndices: number[] = []
    const changedSeen = new Set<number>()
    for (let index = 0; index < pivots.length; index += 1) {
      const pivot = pivots[index]!
      if (!forceAll && !shouldRefreshPivot(pivot, changed)) {
        continue
      }
      const pivotChanges = args.materializePivot(pivot)
      for (let changeIndex = 0; changeIndex < pivotChanges.length; changeIndex += 1) {
        const cellIndex = pivotChanges[changeIndex]!
        if (changedSeen.has(cellIndex)) {
          continue
        }
        changedSeen.add(cellIndex)
        changedCellIndices.push(cellIndex)
      }
    }

    return changedCellIndices.length === 0 ? args.emptyChangedSet() : Uint32Array.from(changedCellIndices)
  }

  const recalculateInternal = (
    changedRoots: readonly number[] | U32,
    kernelSyncRoots: readonly number[] | U32,
    firstPassOrder?: {
      orderedFormulaCellIndices: readonly number[] | U32
      orderedFormulaCount: number
    },
  ): U32 => {
    const started = args.performanceNow()
    args.ensureRecalcScratchCapacity(args.state.workbook.cellStore.size + 1)
    let pendingKernelSync = args.getPendingKernelSync()
    let wasmBatch = args.getWasmBatch()
    let deferredKernelSyncCount = args.getDeferredKernelSyncCount()
    let deferredKernelSyncEpoch = args.getDeferredKernelSyncEpoch() + 1
    const deferredKernelSyncSeen = args.getDeferredKernelSyncSeen()
    if (deferredKernelSyncEpoch === 0xffff_ffff) {
      deferredKernelSyncEpoch = 1
      deferredKernelSyncSeen.fill(0)
    }
    args.setDeferredKernelSyncEpoch(deferredKernelSyncEpoch)
    for (let index = 0; index < deferredKernelSyncCount; index += 1) {
      const cellIndex = pendingKernelSync[index]
      if (cellIndex !== undefined) {
        deferredKernelSyncSeen[cellIndex] = deferredKernelSyncEpoch
      }
    }
    if (args.state.wasm.ready) {
      args.state.wasm.syncStringPool(args.state.strings.exportLayout())
    }

    const allChangedRoots = [...changedRoots]
    const allOrdered: number[] = []
    let singlePassOrdered: readonly number[] | U32 | null = null
    let singlePassOrderedCount = 0
    let pendingFirstPassOrder = firstPassOrder
    let passRoots = [...changedRoots]
    let passKernelRoots = [...kernelSyncRoots]
    let totalOrderedCount = 0
    let totalRangeNodeVisits = 0
    let wasmCount = 0
    let jsCount = 0
    let pendingKernelSyncCount = deferredKernelSyncCount
    const volatileState = createRecalcVolatileState(args.now)

    const flushWasmBatch = (batchCount: number, hasVolatile: boolean, randCount: number): number => {
      if (batchCount === 0) {
        return 0
      }
      args.state.wasm.syncFromStore(args.state.workbook.cellStore, pendingKernelSync.subarray(0, pendingKernelSyncCount))
      pendingKernelSyncCount = 0
      deferredKernelSyncCount = 0
      args.setDeferredKernelSyncCount(0)
      deferredKernelSyncEpoch += 1
      if (deferredKernelSyncEpoch === 0xffff_ffff) {
        deferredKernelSyncEpoch = 1
        deferredKernelSyncSeen.fill(0)
      }
      args.setDeferredKernelSyncEpoch(deferredKernelSyncEpoch)
      if (hasVolatile) {
        args.state.wasm.uploadVolatileNowSerial(volatileState.nowSerial)
        args.state.wasm.uploadVolatileRandomValues(consumeVolatileRandomValues(volatileState, randCount, args.random))
      }
      const batchIndices = wasmBatch.subarray(0, batchCount)
      args.state.wasm.evalBatch(batchIndices)
      args.state.wasm.syncToStore(args.state.workbook.cellStore, batchIndices, args.state.strings, (cellIndex) =>
        args.state.workbook.notifyCellValueWritten(cellIndex),
      )
      return batchCount
    }

    while (pendingFirstPassOrder || passRoots.length > 0) {
      let ordered: readonly number[] | U32
      let orderedCount: number
      let rangeNodeVisits = 0
      if (pendingFirstPassOrder) {
        ordered = pendingFirstPassOrder.orderedFormulaCellIndices
        orderedCount = pendingFirstPassOrder.orderedFormulaCount
        pendingFirstPassOrder = undefined
      } else {
        const scheduled = args.dirtyScheduler.collectDirty(passRoots)
        ordered = scheduled.orderedFormulaCellIndices
        orderedCount = scheduled.orderedFormulaCount
        rangeNodeVisits = scheduled.rangeNodeVisits
      }
      totalOrderedCount += orderedCount
      totalRangeNodeVisits += rangeNodeVisits
      if (singlePassOrdered === null && allOrdered.length === 0) {
        singlePassOrdered = ordered
        singlePassOrderedCount = orderedCount
      } else {
        if (singlePassOrdered !== null) {
          for (let orderedIndex = 0; orderedIndex < singlePassOrderedCount; orderedIndex += 1) {
            const cellIndex = singlePassOrdered[orderedIndex]
            if (cellIndex !== undefined) {
              allOrdered.push(cellIndex)
            }
          }
          singlePassOrdered = null
          singlePassOrderedCount = 0
        }
        for (let orderedIndex = 0; orderedIndex < orderedCount; orderedIndex += 1) {
          allOrdered.push(ordered[orderedIndex]!)
        }
      }

      for (let index = 0; index < passKernelRoots.length; index += 1) {
        const cellIndex = passKernelRoots[index]!
        if (deferredKernelSyncSeen[cellIndex] === deferredKernelSyncEpoch) {
          continue
        }
        deferredKernelSyncSeen[cellIndex] = deferredKernelSyncEpoch
        pendingKernelSync[pendingKernelSyncCount] = cellIndex
        pendingKernelSyncCount += 1
      }

      let wasmBatchCount = 0
      let wasmBatchHasVolatile = false
      let wasmBatchRandCount = 0
      const spillChangedRoots: number[] = []
      const spillChangedSeen = new Set<number>()
      const noteSpillChanges = (changedCellIndices: readonly number[]): void => {
        for (let spillIndex = 0; spillIndex < changedCellIndices.length; spillIndex += 1) {
          const changedCellIndex = changedCellIndices[spillIndex]!
          if (spillChangedSeen.has(changedCellIndex)) {
            continue
          }
          spillChangedSeen.add(changedCellIndex)
          spillChangedRoots.push(changedCellIndex)
        }
      }
      const queueKernelSync = (cellIndex: number): void => {
        if (deferredKernelSyncSeen[cellIndex] === deferredKernelSyncEpoch) {
          return
        }
        deferredKernelSyncSeen[cellIndex] = deferredKernelSyncEpoch
        pendingKernelSync[pendingKernelSyncCount] = cellIndex
        pendingKernelSyncCount += 1
      }
      const hasCycleDependency = (formula: RuntimeFormula): boolean => {
        for (let dependencyIndex = 0; dependencyIndex < formula.dependencyIndices.length; dependencyIndex += 1) {
          const dependencyCellIndex = formula.dependencyIndices[dependencyIndex]!
          if (((args.state.workbook.cellStore.flags[dependencyCellIndex] ?? 0) & CellFlags.InCycle) !== 0) {
            return true
          }
        }
        return false
      }
      const materializeCycleDependentError = (cellIndex: number): void => {
        const spillChanges = args.clearOwnedSpill(cellIndex)
        args.state.workbook.cellStore.setValue(cellIndex, errorValue(ErrorCode.Cycle))
        queueKernelSync(cellIndex)
        noteSpillChanges(spillChanges)
        for (let spillIndex = 0; spillIndex < spillChanges.length; spillIndex += 1) {
          queueKernelSync(spillChanges[spillIndex]!)
        }
      }
      const evaluateWasmSpillFormula = (cellIndex: number, formula: RuntimeFormula): number => {
        args.state.wasm.syncFromStore(args.state.workbook.cellStore, pendingKernelSync.subarray(0, pendingKernelSyncCount))
        pendingKernelSyncCount = 0
        deferredKernelSyncCount = 0
        args.setDeferredKernelSyncCount(0)
        deferredKernelSyncEpoch += 1
        if (deferredKernelSyncEpoch === 0xffff_ffff) {
          deferredKernelSyncEpoch = 1
          deferredKernelSyncSeen.fill(0)
        }
        args.setDeferredKernelSyncEpoch(deferredKernelSyncEpoch)
        if (formula.compiled.volatile) {
          args.state.wasm.uploadVolatileNowSerial(volatileState.nowSerial)
          args.state.wasm.uploadVolatileRandomValues(
            consumeVolatileRandomValues(volatileState, formula.compiled.randCallCount, args.random),
          )
        }
        const batchIndices = Uint32Array.of(cellIndex)
        args.state.wasm.evalBatch(batchIndices)
        args.state.wasm.syncToStore(args.state.workbook.cellStore, batchIndices, args.state.strings, (changedCellIndex) =>
          args.state.workbook.notifyCellValueWritten(changedCellIndex),
        )
        const spill = args.state.wasm.readSpill(cellIndex, args.state.strings)
        const spillMaterialization = spill
          ? args.materializeSpill(cellIndex, {
              rows: spill.rows,
              cols: spill.cols,
              values: spill.values,
            })
          : {
              changedCellIndices: args.clearOwnedSpill(cellIndex),
              ownerValue: args.state.workbook.cellStore.getValue(cellIndex, (id) => args.state.strings.get(id)),
            }
        args.state.workbook.cellStore.setValue(
          cellIndex,
          spillMaterialization.ownerValue,
          spillMaterialization.ownerValue.tag === ValueTag.String ? args.state.strings.intern(spillMaterialization.ownerValue.value) : 0,
        )
        queueKernelSync(cellIndex)
        for (let spillIndex = 0; spillIndex < spillMaterialization.changedCellIndices.length; spillIndex += 1) {
          queueKernelSync(spillMaterialization.changedCellIndices[spillIndex]!)
        }
        noteSpillChanges(spillMaterialization.changedCellIndices)
        return 1
      }

      for (let index = 0; index < orderedCount; index += 1) {
        const cellIndex = ordered[index]!
        const formula = args.state.formulas.get(cellIndex)
        if (!formula) {
          continue
        }
        if (((args.state.workbook.cellStore.flags[cellIndex] ?? 0) & CellFlags.InCycle) !== 0) {
          continue
        }
        if (hasCycleDependency(formula)) {
          jsCount += 1
          materializeCycleDependentError(cellIndex)
          continue
        }
        if (
          formula.directLookup !== undefined ||
          formula.directAggregate !== undefined ||
          formula.directScalar !== undefined ||
          formula.directCriteria !== undefined
        ) {
          if (wasmBatchCount > 0) {
            wasmCount += flushWasmBatch(wasmBatchCount, wasmBatchHasVolatile, wasmBatchRandCount)
            wasmBatchCount = 0
            wasmBatchHasVolatile = false
            wasmBatchRandCount = 0
          }
          const directLookupChanges = args.evaluateDirectLookupFormula(cellIndex)
          if (directLookupChanges !== undefined) {
            if (
              formula.compiled.mode === FormulaMode.WasmFastPath &&
              (formula.directScalar !== undefined || formula.directAggregate !== undefined)
            ) {
              wasmCount += 1
            } else if (
              formula.compiled.mode !== FormulaMode.WasmFastPath &&
              (formula.directScalar !== undefined || formula.directAggregate !== undefined)
            ) {
              jsCount += 1
            }
            noteSpillChanges(directLookupChanges)
            queueKernelSync(cellIndex)
            continue
          }
        }
        if (formula.compiled.mode === FormulaMode.WasmFastPath && args.state.wasm.ready) {
          if (formula.compiled.producesSpill) {
            wasmCount += flushWasmBatch(wasmBatchCount, wasmBatchHasVolatile, wasmBatchRandCount)
            wasmBatchCount = 0
            wasmBatchHasVolatile = false
            wasmBatchRandCount = 0
            wasmCount += evaluateWasmSpillFormula(cellIndex, formula)
            continue
          }
          wasmBatch[wasmBatchCount] = cellIndex
          wasmBatchCount += 1
          wasmBatchHasVolatile = wasmBatchHasVolatile || formula.compiled.volatile
          wasmBatchRandCount += formula.compiled.randCallCount
          continue
        }
        wasmCount += flushWasmBatch(wasmBatchCount, wasmBatchHasVolatile, wasmBatchRandCount)
        wasmBatchCount = 0
        wasmBatchHasVolatile = false
        wasmBatchRandCount = 0
        jsCount += 1
        const spillChanges = args.evaluateUnsupportedFormula(cellIndex)
        noteSpillChanges(spillChanges)
        queueKernelSync(cellIndex)
      }

      wasmCount += flushWasmBatch(wasmBatchCount, wasmBatchHasVolatile, wasmBatchRandCount)
      args.setDeferredKernelSyncCount(pendingKernelSyncCount)
      deferredKernelSyncCount = pendingKernelSyncCount

      if (spillChangedRoots.length === 0) {
        break
      }
      if (singlePassOrdered !== null) {
        for (let orderedIndex = 0; orderedIndex < singlePassOrderedCount; orderedIndex += 1) {
          const cellIndex = singlePassOrdered[orderedIndex]
          if (cellIndex !== undefined) {
            allOrdered.push(cellIndex)
          }
        }
        singlePassOrdered = null
        singlePassOrderedCount = 0
      }
      allChangedRoots.push(...spillChangedRoots)
      passRoots = spillChangedRoots
      passKernelRoots = spillChangedRoots
    }

    const lastMetrics = { ...args.state.getLastMetrics() }
    lastMetrics.dirtyFormulaCount = totalOrderedCount
    lastMetrics.jsFormulaCount = jsCount
    lastMetrics.wasmFormulaCount = wasmCount
    lastMetrics.rangeNodeVisits = totalRangeNodeVisits
    lastMetrics.recalcMs = args.performanceNow() - started
    args.state.setLastMetrics(lastMetrics)
    args.setDeferredKernelSyncCount(pendingKernelSyncCount)
    if (singlePassOrdered !== null) {
      return totalOrderedCount === 0 && allChangedRoots.length === 0
        ? args.emptyChangedSet()
        : args.composeChangedRootsAndOrdered(
            allChangedRoots,
            toOrderedUint32(singlePassOrdered, singlePassOrderedCount),
            singlePassOrderedCount,
          )
    }
    return totalOrderedCount === 0 && allChangedRoots.length === 0
      ? args.emptyChangedSet()
      : args.composeChangedRootsAndOrdered(allChangedRoots, Uint32Array.from(allOrdered), allOrdered.length)
  }

  const recalculate = (changedRoots: readonly number[] | U32, kernelSyncRoots: readonly number[] | U32 = changedRoots): U32 =>
    recalculateInternal(changedRoots, kernelSyncRoots)

  const recalculatePreordered = (
    changedRoots: readonly number[] | U32,
    orderedFormulaCellIndices: readonly number[] | U32,
    orderedFormulaCount: number,
    kernelSyncRoots: readonly number[] | U32 = changedRoots,
  ): U32 =>
    recalculateInternal(changedRoots, kernelSyncRoots, {
      orderedFormulaCellIndices,
      orderedFormulaCount,
    })

  const reconcilePivotOutputs = (baseChanged: U32, forceAllPivots = false): U32 => {
    let aggregate = baseChanged
    let pending = baseChanged
    let forceAll = forceAllPivots

    for (let iteration = 0; iteration < 4; iteration += 1) {
      const pivotChanged = refreshPivotOutputs(pending, forceAll)
      if (pivotChanged.length === 0) {
        break
      }
      aggregate = aggregate.length === 0 ? pivotChanged : args.unionChangedSets(aggregate, pivotChanged)
      pending = recalculate(pivotChanged, pivotChanged)
      aggregate = pending.length === 0 ? aggregate : args.unionChangedSets(aggregate, pending)
      forceAll = false
    }

    return aggregate
  }

  return {
    recalculatePreordered(changedRoots, orderedFormulaCellIndices, orderedFormulaCount, kernelSyncRoots = changedRoots) {
      return Effect.try({
        try: () => recalculatePreordered(changedRoots, orderedFormulaCellIndices, orderedFormulaCount, kernelSyncRoots),
        catch: (cause) =>
          new EngineRecalcError({
            message: 'Failed to recalculate workbook state from a preordered formula batch',
            cause,
          }),
      })
    },
    recalculate(changedRoots, kernelSyncRoots = changedRoots) {
      return Effect.try({
        try: () => recalculate(changedRoots, kernelSyncRoots),
        catch: (cause) =>
          new EngineRecalcError({
            message: 'Failed to recalculate workbook state',
            cause,
          }),
      })
    },
    reconcilePivotOutputs(baseChanged, forceAllPivots = false) {
      return Effect.try({
        try: () => reconcilePivotOutputs(baseChanged, forceAllPivots),
        catch: (cause) =>
          new EngineRecalcError({
            message: 'Failed to reconcile pivot outputs',
            cause,
          }),
      })
    },
    recalculateNow() {
      return Effect.try({
        try: () => {
          args.beginMutationCollection()
          args.state.workbook.setVolatileContext({
            recalcEpoch: args.state.workbook.getVolatileContext().recalcEpoch + 1,
          })
          let formulaChangedCount = 0
          let explicitChangedCount = 0
          args.state.formulas.forEach((_formula, cellIndex) => {
            formulaChangedCount = args.markFormulaChanged(cellIndex, formulaChangedCount)
            explicitChangedCount = args.markExplicitChanged(cellIndex, explicitChangedCount)
          })
          const recalculated = reconcilePivotOutputs(
            recalculate(args.composeMutationRoots(0, formulaChangedCount), args.emptyChangedSet()),
            true,
          )
          const changed = args.composeEventChanges(recalculated, explicitChangedCount)
          const lastMetrics = { ...args.state.getLastMetrics() }
          lastMetrics.batchId += 1
          lastMetrics.changedInputCount = formulaChangedCount
          args.state.setLastMetrics(lastMetrics)
          const event: EngineEvent & {
            explicitChangedCount: number
          } = {
            kind: 'batch',
            invalidation: 'cells',
            changedCellIndices: changed,
            changedCells: args.captureChangedCells(changed),
            invalidatedRanges: [],
            invalidatedRows: [],
            invalidatedColumns: [],
            metrics: lastMetrics,
            explicitChangedCount,
          }
          args.state.events.emit(event, changed, (cellIndex) => args.state.workbook.getQualifiedAddress(cellIndex))
          if (args.state.events.hasTrackedListeners()) {
            const patches = captureTrackedPatchesForCells(changed)
            args.state.events.emitTracked({
              kind: 'batch',
              invalidation: 'cells',
              changedCellIndices: changed,
              ...(patches ? { patches } : {}),
              invalidatedRanges: [],
              invalidatedRows: [],
              invalidatedColumns: [],
              metrics: lastMetrics,
              explicitChangedCount,
            })
          }
          return Array.from(changed)
        },
        catch: (cause) =>
          new EngineRecalcError({
            message: 'Failed to recalculate all formulas',
            cause,
          }),
      })
    },
    recalculateDirty(dirtyRegions) {
      return Effect.try({
        try: () => {
          args.beginMutationCollection()
          let changedInputCount = 0
          let explicitChangedCount = 0

          for (const region of dirtyRegions) {
            const sheet = args.state.workbook.getSheet(region.sheetName)
            if (!sheet) {
              continue
            }

            for (let row = region.rowStart; row <= region.rowEnd; row += 1) {
              for (let col = region.colStart; col <= region.colEnd; col += 1) {
                const cellIndex = args.state.workbook.cellKeyToIndex.get(makeCellKey(sheet.id, row, col))
                if (cellIndex !== undefined) {
                  changedInputCount = args.markInputChanged(cellIndex, changedInputCount)
                  explicitChangedCount = args.markExplicitChanged(cellIndex, explicitChangedCount)
                }
              }
            }
          }

          const changedInputArray = args.getChangedInputBuffer().subarray(0, changedInputCount)
          const recalculated = reconcilePivotOutputs(recalculate(args.composeMutationRoots(changedInputCount, 0), changedInputArray), false)
          const changed = args.composeEventChanges(recalculated, explicitChangedCount)
          const lastMetrics = { ...args.state.getLastMetrics() }
          lastMetrics.batchId += 1
          lastMetrics.changedInputCount = changedInputCount
          args.state.setLastMetrics(lastMetrics)
          const event: EngineEvent & {
            explicitChangedCount: number
          } = {
            kind: 'batch',
            invalidation: 'cells',
            changedCellIndices: changed,
            changedCells: args.captureChangedCells(changed),
            invalidatedRanges: [],
            invalidatedRows: [],
            invalidatedColumns: [],
            metrics: lastMetrics,
            explicitChangedCount,
          }
          args.state.events.emit(event, changed, (cellIndex) => args.state.workbook.getQualifiedAddress(cellIndex))
          if (args.state.events.hasTrackedListeners()) {
            const patches = captureTrackedPatchesForCells(changed)
            args.state.events.emitTracked({
              kind: 'batch',
              invalidation: 'cells',
              changedCellIndices: changed,
              ...(patches ? { patches } : {}),
              invalidatedRanges: [],
              invalidatedRows: [],
              invalidatedColumns: [],
              metrics: lastMetrics,
              explicitChangedCount,
            })
          }
          return Array.from(changed)
        },
        catch: (cause) =>
          new EngineRecalcError({
            message: 'Failed to recalculate dirty regions',
            cause,
          }),
      })
    },
    recalculateDifferential() {
      return Effect.try({
        try: () => {
          const originalSnapshot = args.exportSnapshot()
          args.state.formulas.forEach((formula) => {
            formula.compiled.mode = FormulaMode.JsOnly
          })
          const jsChanged = Effect.runSync(this.recalculateNow())
          const jsResults = jsChanged.map((idx) => args.getCellByIndex(idx))

          args.importSnapshot(originalSnapshot)
          const wasmChanged = Effect.runSync(this.recalculateNow())
          const wasmResults = wasmChanged.map((idx) => args.getCellByIndex(idx))

          const drift: string[] = []
          const jsMap = new Map(jsResults.map((result) => [`${result.sheetName}!${result.address}`, result]))
          const wasmMap = new Map(wasmResults.map((result) => [`${result.sheetName}!${result.address}`, result]))

          for (const [addr, jsCell] of jsMap) {
            const wasmCell = wasmMap.get(addr)
            if (!wasmCell) {
              drift.push(`${addr}: Calculated in JS but MISSING in WASM`)
              continue
            }
            if (JSON.stringify(jsCell.value) !== JSON.stringify(wasmCell.value)) {
              drift.push(`${addr}: JS=${JSON.stringify(jsCell.value)} WASM=${JSON.stringify(wasmCell.value)}`)
            }
          }

          for (const addr of wasmMap.keys()) {
            if (!jsMap.has(addr)) {
              drift.push(`${addr}: Calculated in WASM but MISSING in JS`)
            }
          }

          return { js: jsResults, wasm: wasmResults, drift }
        },
        catch: (cause) =>
          new EngineRecalcError({
            message: 'Failed to run differential recalculation',
            cause,
          }),
      })
    },
    recalculatePreorderedNowSync: recalculatePreordered,
    recalculateNowSync: recalculate,
    reconcilePivotOutputsNow: reconcilePivotOutputs,
  }
}
