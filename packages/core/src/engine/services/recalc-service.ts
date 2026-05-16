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
import { areCellValuesEqual, emptyValue, errorValue } from '../../engine-value-utils.js'
import type { EngineRuntimeState, RuntimeFormula, SpillMaterialization, U32 } from '../runtime-state.js'
import { EngineRecalcError } from '../errors.js'
import type { WorkbookPivotRecord } from '../../workbook-store.js'
import { parseCellAddress } from '@bilig/formula'
import type { EngineDirtyFrontierSchedulerService } from './dirty-frontier-scheduler-service.js'
import type { EnginePatch } from '../../patches/patch-types.js'
import { buildCycleEvaluationNodes, type CycleEvaluationNode } from './recalc-cycle-evaluation.js'
import { consumeVolatileRandomValues, createRecalcVolatileState, toOrderedUint32 } from './recalc-evaluation-state.js'
import { resolveRecalcIterationSettings } from './recalc-iteration-settings.js'

const TRACKED_CELL_PATCH_LIMIT = 64

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
  readonly flushWasmProgramSync: () => void
  readonly beginEvaluationBudget: (startedAtMs: number) => void
  readonly endEvaluationBudget: () => void
  readonly checkEvaluationBudget: (stepCost?: number) => void
  readonly now: () => Date
  readonly random: () => number
  readonly performanceNow: () => number
  readonly dirtyScheduler: EngineDirtyFrontierSchedulerService
  readonly materializeSpill: (cellIndex: number, arrayValue: { values: CellValue[]; rows: number; cols: number }) => SpillMaterialization
  readonly clearOwnedSpill: (cellIndex: number) => number[]
  readonly evaluateDirectLookupFormula: (cellIndex: number) => number[] | undefined
  readonly evaluateUnsupportedFormula: (cellIndex: number) => number[]
  readonly materializePivot: (pivot: WorkbookPivotRecord) => number[]
  readonly forEachFormulaDependencyCell: (cellIndex: number, fn: (dependencyCellIndex: number) => void) => void
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
    args.beginEvaluationBudget(started)
    try {
      args.checkEvaluationBudget()
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
      const iterationSettings = resolveRecalcIterationSettings(args.state.workbook.getCalculationSettings())
      let wasmProgramFlushed = false
      const ensureWasmProgramFlushed = (): void => {
        if (wasmProgramFlushed) {
          return
        }
        args.flushWasmProgramSync()
        wasmProgramFlushed = true
      }

      if (changedRoots.length === 0 && kernelSyncRoots.length > 0) {
        for (let index = 0; index < kernelSyncRoots.length; index += 1) {
          const cellIndex = kernelSyncRoots[index]!
          if (deferredKernelSyncSeen[cellIndex] === deferredKernelSyncEpoch) {
            continue
          }
          deferredKernelSyncSeen[cellIndex] = deferredKernelSyncEpoch
          pendingKernelSync[pendingKernelSyncCount] = cellIndex
          pendingKernelSyncCount += 1
        }
        const lastMetrics = { ...args.state.getLastMetrics() }
        lastMetrics.dirtyFormulaCount = 0
        lastMetrics.jsFormulaCount = 0
        lastMetrics.wasmFormulaCount = 0
        lastMetrics.rangeNodeVisits = 0
        lastMetrics.recalcMs = args.performanceNow() - started
        args.state.setLastMetrics(lastMetrics)
        args.setDeferredKernelSyncCount(pendingKernelSyncCount)
        return args.emptyChangedSet()
      }

      const flushWasmBatch = (batchCount: number, hasVolatile: boolean, randCount: number): number => {
        if (batchCount === 0) {
          return 0
        }
        ensureWasmProgramFlushed()
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
        args.checkEvaluationBudget(batchCount)
        args.state.wasm.evalBatch(batchIndices)
        args.state.wasm.syncToStore(args.state.workbook.cellStore, batchIndices, args.state.strings, (cellIndex) =>
          args.state.workbook.notifyCellValueWritten(cellIndex),
        )
        args.checkEvaluationBudget(batchCount)
        return batchCount
      }

      while (pendingFirstPassOrder || passRoots.length > 0) {
        args.checkEvaluationBudget()
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
        const noteQueuedSpillChanges = (changedCellIndices: readonly number[]): void => {
          noteSpillChanges(changedCellIndices)
          for (let spillIndex = 0; spillIndex < changedCellIndices.length; spillIndex += 1) {
            queueKernelSync(changedCellIndices[spillIndex]!)
          }
        }
        const flushPendingWasmBatch = (): void => {
          wasmCount += flushWasmBatch(wasmBatchCount, wasmBatchHasVolatile, wasmBatchRandCount)
          wasmBatchCount = 0
          wasmBatchHasVolatile = false
          wasmBatchRandCount = 0
        }
        const readStoredValue = (cellIndex: number): CellValue =>
          args.state.workbook.cellStore.getValue(cellIndex, (id) => (id === 0 ? '' : args.state.strings.get(id)))
        const clearDerivedFormulaFlags = (cellIndex: number): boolean => {
          const currentFlags = args.state.workbook.cellStore.flags[cellIndex] ?? 0
          const nextFlags = currentFlags & ~(CellFlags.SpillChild | CellFlags.PivotOutput)
          if (nextFlags === currentFlags) {
            return false
          }
          args.state.workbook.cellStore.flags[cellIndex] = nextFlags
          return true
        }
        let hasAnyCycleFormula = false
        args.state.formulas.forEach((_formula, formulaCellIndex) => {
          hasAnyCycleFormula ||= ((args.state.workbook.cellStore.flags[formulaCellIndex] ?? 0) & CellFlags.InCycle) !== 0
        })
        const hasCycleDependency = (cellIndex: number): boolean => {
          if (!hasAnyCycleFormula) {
            return false
          }
          let found = false
          args.forEachFormulaDependencyCell(cellIndex, (dependencyCellIndex) => {
            if (!found && ((args.state.workbook.cellStore.flags[dependencyCellIndex] ?? 0) & CellFlags.InCycle) !== 0) {
              found = true
            }
          })
          return found
        }
        const materializeCycleFormulaError = (cellIndex: number): void => {
          const beforeValue = readStoredValue(cellIndex)
          const spillChanges = args.clearOwnedSpill(cellIndex)
          const flagsChanged = clearDerivedFormulaFlags(cellIndex)
          const nextValue = errorValue(ErrorCode.Cycle)
          if (!flagsChanged && spillChanges.length === 0 && areCellValuesEqual(beforeValue, nextValue)) {
            return
          }
          args.state.workbook.cellStore.setValue(cellIndex, nextValue)
          args.state.workbook.notifyCellValueWritten(cellIndex)
          queueKernelSync(cellIndex)
          noteQueuedSpillChanges(spillChanges)
        }
        const seedCycleFormulaCell = (cellIndex: number): void => {
          const currentValue = readStoredValue(cellIndex)
          if (currentValue.tag !== ValueTag.Error || currentValue.code !== ErrorCode.Cycle) {
            return
          }
          const spillChanges = args.clearOwnedSpill(cellIndex)
          clearDerivedFormulaFlags(cellIndex)
          args.state.workbook.cellStore.setValue(cellIndex, emptyValue())
          args.state.workbook.notifyCellValueWritten(cellIndex)
          queueKernelSync(cellIndex)
          noteQueuedSpillChanges(spillChanges)
        }
        const cycleIterationDrift = (beforeValue: CellValue, afterValue: CellValue): number => {
          if (beforeValue.tag === ValueTag.Number && afterValue.tag === ValueTag.Number) {
            if (Object.is(beforeValue.value, afterValue.value)) {
              return 0
            }
            const drift = Math.abs(afterValue.value - beforeValue.value)
            return Number.isFinite(drift) ? drift : Number.POSITIVE_INFINITY
          }
          return areCellValuesEqual(beforeValue, afterValue) ? 0 : Number.POSITIVE_INFINITY
        }
        const evaluateWasmSpillFormula = (cellIndex: number, formula: RuntimeFormula): number => {
          ensureWasmProgramFlushed()
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
          args.checkEvaluationBudget()
          args.state.wasm.evalBatch(batchIndices)
          args.state.wasm.syncToStore(args.state.workbook.cellStore, batchIndices, args.state.strings, (changedCellIndex) =>
            args.state.workbook.notifyCellValueWritten(changedCellIndex),
          )
          args.checkEvaluationBudget()
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
          noteQueuedSpillChanges(spillMaterialization.changedCellIndices)
          return 1
        }

        const evaluateFormulaCell = (
          cellIndex: number,
          formula: RuntimeFormula,
          options: {
            readonly allowCycleDependencyError: boolean
            readonly treatCycleFormulaAsError: boolean
            readonly forceJs: boolean
          },
        ): void => {
          if (options.treatCycleFormulaAsError && ((args.state.workbook.cellStore.flags[cellIndex] ?? 0) & CellFlags.InCycle) !== 0) {
            jsCount += 1
            materializeCycleFormulaError(cellIndex)
            return
          }
          if (options.allowCycleDependencyError && hasCycleDependency(cellIndex)) {
            jsCount += 1
            materializeCycleFormulaError(cellIndex)
            return
          }
          if (
            formula.directLookup !== undefined ||
            formula.directAggregate !== undefined ||
            formula.directScalar !== undefined ||
            formula.directCriteria !== undefined
          ) {
            flushPendingWasmBatch()
            args.checkEvaluationBudget()
            const directLookupChanges = args.evaluateDirectLookupFormula(cellIndex)
            args.checkEvaluationBudget()
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
              noteQueuedSpillChanges(directLookupChanges)
              queueKernelSync(cellIndex)
              return
            }
          }
          if (!options.forceJs && formula.compiled.mode === FormulaMode.WasmFastPath && args.state.wasm.ready) {
            if (formula.compiled.producesSpill) {
              flushPendingWasmBatch()
              wasmCount += evaluateWasmSpillFormula(cellIndex, formula)
              return
            }
            wasmBatch[wasmBatchCount] = cellIndex
            wasmBatchCount += 1
            wasmBatchHasVolatile = wasmBatchHasVolatile || formula.compiled.volatile
            wasmBatchRandCount += formula.compiled.randCallCount
            return
          }
          flushPendingWasmBatch()
          jsCount += 1
          args.checkEvaluationBudget()
          const spillChanges = args.evaluateUnsupportedFormula(cellIndex)
          args.checkEvaluationBudget()
          noteQueuedSpillChanges(spillChanges)
          queueKernelSync(cellIndex)
        }
        const evaluateCycleNode = (node: CycleEvaluationNode): void => {
          for (let formulaIndex = 0; formulaIndex < node.formulaCellIndices.length; formulaIndex += 1) {
            seedCycleFormulaCell(node.formulaCellIndices[formulaIndex]!)
          }
          const previousValues = node.formulaCellIndices.map((cellIndex) => readStoredValue(cellIndex))
          for (let iterationIndex = 0; iterationIndex < iterationSettings.count; iterationIndex += 1) {
            args.checkEvaluationBudget()
            for (let formulaIndex = 0; formulaIndex < node.formulaCellIndices.length; formulaIndex += 1) {
              const cellIndex = node.formulaCellIndices[formulaIndex]!
              const formula = args.state.formulas.get(cellIndex)
              if (!formula) {
                continue
              }
              evaluateFormulaCell(cellIndex, formula, {
                allowCycleDependencyError: false,
                treatCycleFormulaAsError: false,
                forceJs: true,
              })
            }

            let converged = true
            for (let formulaIndex = 0; formulaIndex < node.formulaCellIndices.length; formulaIndex += 1) {
              const cellIndex = node.formulaCellIndices[formulaIndex]!
              const currentValue = readStoredValue(cellIndex)
              if (cycleIterationDrift(previousValues[formulaIndex]!, currentValue) > iterationSettings.delta) {
                converged = false
              }
              previousValues[formulaIndex] = currentValue
            }
            if (converged) {
              break
            }
          }
        }
        const cycleEvaluationNodes = iterationSettings.enabled
          ? buildCycleEvaluationNodes({
              ordered,
              orderedCount,
              formulas: args.state.formulas,
              cycleGroupIds: args.state.workbook.cellStore.cycleGroupIds,
              forEachFormulaDependencyCell: args.forEachFormulaDependencyCell,
            })
          : undefined

        if (cycleEvaluationNodes) {
          for (let nodeIndex = 0; nodeIndex < cycleEvaluationNodes.length; nodeIndex += 1) {
            args.checkEvaluationBudget()
            const node = cycleEvaluationNodes[nodeIndex]!
            if (node.kind === 'cycle') {
              flushPendingWasmBatch()
              evaluateCycleNode(node)
              continue
            }
            for (let formulaIndex = 0; formulaIndex < node.formulaCellIndices.length; formulaIndex += 1) {
              const cellIndex = node.formulaCellIndices[formulaIndex]!
              const formula = args.state.formulas.get(cellIndex)
              if (!formula) {
                continue
              }
              evaluateFormulaCell(cellIndex, formula, {
                allowCycleDependencyError: false,
                treatCycleFormulaAsError: false,
                forceJs: false,
              })
            }
          }
        } else {
          for (let index = 0; index < orderedCount; index += 1) {
            args.checkEvaluationBudget()
            const cellIndex = ordered[index]!
            const formula = args.state.formulas.get(cellIndex)
            if (!formula) {
              continue
            }
            evaluateFormulaCell(cellIndex, formula, {
              allowCycleDependencyError: !iterationSettings.enabled,
              treatCycleFormulaAsError: !iterationSettings.enabled,
              forceJs: false,
            })
          }
        }

        flushPendingWasmBatch()
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
    } finally {
      args.endEvaluationBudget()
    }
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
          let canUseFullFormulaOrder = true
          args.state.formulas.forEach((formula, cellIndex) => {
            formulaChangedCount = args.markFormulaChanged(cellIndex, formulaChangedCount)
            explicitChangedCount = args.markExplicitChanged(cellIndex, explicitChangedCount)
            if (formula.compiled.producesSpill || formula.directLookup !== undefined || formula.directCriteria !== undefined) {
              canUseFullFormulaOrder = false
            }
          })
          const mutationRoots = args.composeMutationRoots(0, formulaChangedCount)
          let recalculatedBase: U32
          if (canUseFullFormulaOrder) {
            const fullFormulaOrder = args.dirtyScheduler.collectAll()
            recalculatedBase = recalculatePreordered(
              mutationRoots,
              fullFormulaOrder.orderedFormulaCellIndices,
              fullFormulaOrder.orderedFormulaCount,
              args.emptyChangedSet(),
            )
          } else {
            recalculatedBase = recalculate(mutationRoots, args.emptyChangedSet())
          }
          const recalculated = reconcilePivotOutputs(recalculatedBase, true)
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
