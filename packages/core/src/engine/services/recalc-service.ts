import { Effect } from 'effect'
import { ErrorCode, FormulaMode, ValueTag, type CellValue } from '@bilig/protocol'
import { CellFlags } from '../../cell-store.js'
import { areCellValuesEqual, emptyValue, errorValue } from '../../engine-value-utils.js'
import type { RuntimeFormula, U32 } from '../runtime-state.js'
import { EngineRecalcError } from '../errors.js'
import { buildCycleEvaluationNodes, type CycleEvaluationNode } from './recalc-cycle-evaluation.js'
import { consumeVolatileRandomValues, createRecalcVolatileState, toOrderedUint32 } from './recalc-evaluation-state.js'
import { emitRecalcBatchEvents } from './recalc-event-emission.js'
import { hasQueuedFormulaDependency } from './recalc-formula-dependency-helpers.js'
import { resolveRecalcIterationSettings } from './recalc-iteration-settings.js'
import { createRecalcNativeDirectScalarBatch, MAX_RECALC_NATIVE_DIRECT_SCALAR_BATCH_SIZE } from './formula-recalc-native-direct-scalar.js'
import {
  createRecalcNativeDirectLookupBatch,
  MAX_RECALC_NATIVE_DIRECT_LOOKUP_BATCH_SIZE,
} from './formula-initialization-native-direct-lookup.js'
import { refreshPivotOutputsForChangedCells } from './recalc-pivot-refresh.js'
import { filterSkippedCachedFormulaCells } from './recalc-skipped-cached-formula-cells.js'
import { capturePivotOutputValues, notePivotValueChanges } from './recalc-pivot-value-changes.js'
import { createRecalcValueChangeCollector, type RecalcValueChangeCollector } from './recalc-value-change-collector.js'
import type { EngineRecalcService, EngineRecalcServiceArgs } from './recalc-service-types.js'
import { forEachDirtyRegionCell } from './recalc-dirty-region-cells.js'

export type { DirtyRegion, EngineRecalcService } from './recalc-service-types.js'

interface RecalculateInternalOptions {
  readonly orderedFormulaCellIndices?: readonly number[] | U32
  readonly orderedFormulaCount?: number
  readonly preserveCachedValuesOnFullRecalc?: boolean
  readonly valueChangeCollector?: RecalcValueChangeCollector
}

interface RecalculateAllNowOptions {
  readonly collectValueChangesForRebuild?: boolean
}

export function createEngineRecalcService(args: EngineRecalcServiceArgs): EngineRecalcService {
  const refreshPivotOutputs = (
    changed: readonly number[] | U32,
    forceAll: boolean,
    valueChangeCollector?: RecalcValueChangeCollector,
  ): U32 =>
    refreshPivotOutputsForChangedCells({
      changed,
      forceAll,
      workbook: args.state.workbook,
      materializePivot: (pivot) => {
        const beforeValues =
          valueChangeCollector === undefined
            ? undefined
            : capturePivotOutputValues({ pivot, workbook: args.state.workbook, strings: args.state.strings })
        const pivotChanged = args.materializePivot(pivot)
        if (beforeValues !== undefined && valueChangeCollector !== undefined) {
          notePivotValueChanges({
            changed: pivotChanged,
            before: beforeValues,
            workbook: args.state.workbook,
            strings: args.state.strings,
            collector: valueChangeCollector,
          })
        }
        return pivotChanged
      },
      emptyChangedSet: args.emptyChangedSet,
    })

  const recalculateInternal = (
    changedRoots: readonly number[] | U32,
    kernelSyncRoots: readonly number[] | U32,
    options: RecalculateInternalOptions = {},
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

      const precisionAsDisplayed = args.state.workbook.getCalculationSettings().fullPrecision === false
      const skippedCachedFormulaCells = options.preserveCachedValuesOnFullRecalc === true ? new Set<number>() : undefined
      const valueChangeCollector = options.valueChangeCollector
      const noteValueChanged = (cellIndex: number): void => {
        valueChangeCollector?.note(cellIndex)
      }
      const notifyCellValueWritten = (cellIndex: number): void => {
        args.state.workbook.notifyCellValueWritten(cellIndex)
        noteValueChanged(cellIndex)
      }
      const allChangedRoots = [...changedRoots]
      const allOrdered: number[] = []
      let singlePassOrdered: readonly number[] | U32 | null = null
      let singlePassOrderedCount = 0
      let pendingFirstPassOrder =
        options.orderedFormulaCellIndices !== undefined && options.orderedFormulaCount !== undefined
          ? {
              orderedFormulaCellIndices: options.orderedFormulaCellIndices,
              orderedFormulaCount: options.orderedFormulaCount,
            }
          : undefined
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

      if (changedRoots.length === 0 && kernelSyncRoots.length > 0 && pendingFirstPassOrder === undefined) {
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
        args.state.wasm.syncToStore(args.state.workbook.cellStore, batchIndices, args.state.strings, notifyCellValueWritten)
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
            noteValueChanged(changedCellIndex)
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
        const nativeDirectScalarBatch = createRecalcNativeDirectScalarBatch({
          state: args.state,
          capacity: Math.min(Math.max(orderedCount, 1), MAX_RECALC_NATIVE_DIRECT_SCALAR_BATCH_SIZE),
          ...(valueChangeCollector === undefined ? {} : { onCellValueChanged: noteValueChanged }),
        })
        const nativeDirectScalarCells: number[] = []
        const nativeDirectLookupBatch = createRecalcNativeDirectLookupBatch({
          state: args.state,
          capacity: Math.min(Math.max(orderedCount, 1), MAX_RECALC_NATIVE_DIRECT_LOOKUP_BATCH_SIZE),
          ...(valueChangeCollector === undefined ? {} : { onCellValueChanged: noteValueChanged }),
        })
        const nativeDirectLookupCells: number[] = []
        const nativeDirectLookupQueued = new Set<number>()
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
          notifyCellValueWritten(cellIndex)
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
          args.state.wasm.syncToStore(args.state.workbook.cellStore, batchIndices, args.state.strings, notifyCellValueWritten)
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

        const evaluateDirectFormulaImmediately = (cellIndex: number, formula: RuntimeFormula): boolean => {
          args.checkEvaluationBudget()
          const beforeValue = valueChangeCollector === undefined ? undefined : readStoredValue(cellIndex)
          const directLookupChanges = args.evaluateDirectLookupFormula(cellIndex)
          args.checkEvaluationBudget()
          if (directLookupChanges === undefined) {
            return false
          }
          if (beforeValue !== undefined && !areCellValuesEqual(beforeValue, readStoredValue(cellIndex))) {
            noteValueChanged(cellIndex)
          }
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
          return true
        }

        const evaluateUnsupportedFormulaWithValueCollection = (cellIndex: number): number[] => {
          const beforeValue = valueChangeCollector === undefined ? undefined : readStoredValue(cellIndex)
          const spillChanges = args.evaluateUnsupportedFormula(cellIndex)
          if (beforeValue !== undefined && !areCellValuesEqual(beforeValue, readStoredValue(cellIndex))) {
            noteValueChanged(cellIndex)
          }
          return spillChanges
        }

        const flushNativeDirectLookupBatch = (): void => {
          if (nativeDirectLookupBatch.count === 0) {
            return
          }
          const evaluated = nativeDirectLookupBatch.evaluate()
          if (evaluated !== undefined) {
            wasmCount += evaluated.length
            for (let index = 0; index < evaluated.length; index += 1) {
              queueKernelSync(evaluated[index]!)
            }
          } else {
            for (let index = 0; index < nativeDirectLookupCells.length; index += 1) {
              const cellIndex = nativeDirectLookupCells[index]!
              const formula = args.state.formulas.get(cellIndex)
              if (formula && !evaluateDirectFormulaImmediately(cellIndex, formula)) {
                jsCount += 1
                const spillChanges = evaluateUnsupportedFormulaWithValueCollection(cellIndex)
                noteQueuedSpillChanges(spillChanges)
                queueKernelSync(cellIndex)
              }
            }
          }
          nativeDirectLookupCells.length = 0
          nativeDirectLookupQueued.clear()
          nativeDirectLookupBatch.reset()
        }

        const flushNativeDirectScalarBatch = (): void => {
          if (nativeDirectScalarBatch.count === 0) {
            return
          }
          const evaluated = nativeDirectScalarBatch.evaluate()
          if (evaluated !== undefined) {
            wasmCount += evaluated.length
            for (let index = 0; index < evaluated.length; index += 1) {
              queueKernelSync(evaluated[index]!)
            }
          } else {
            for (let index = 0; index < nativeDirectScalarCells.length; index += 1) {
              const cellIndex = nativeDirectScalarCells[index]!
              const formula = args.state.formulas.get(cellIndex)
              if (formula && !evaluateDirectFormulaImmediately(cellIndex, formula)) {
                jsCount += 1
                const spillChanges = evaluateUnsupportedFormulaWithValueCollection(cellIndex)
                noteQueuedSpillChanges(spillChanges)
                queueKernelSync(cellIndex)
              }
            }
          }
          nativeDirectScalarCells.length = 0
          nativeDirectScalarBatch.reset()
        }

        const evaluateFormulaCell = (
          cellIndex: number,
          formula: RuntimeFormula,
          evaluationOptions: {
            readonly allowCycleDependencyError: boolean
            readonly treatCycleFormulaAsError: boolean
            readonly forceJs: boolean
          },
        ): void => {
          if (hasQueuedFormulaDependency(cellIndex, nativeDirectLookupQueued, args.forEachFormulaDependencyCell)) {
            flushNativeDirectLookupBatch()
          }
          if (skippedCachedFormulaCells && formula.preserveCachedValueOnFullRecalc === true) {
            skippedCachedFormulaCells.add(cellIndex)
            queueKernelSync(cellIndex)
            return
          }
          if (
            evaluationOptions.treatCycleFormulaAsError &&
            ((args.state.workbook.cellStore.flags[cellIndex] ?? 0) & CellFlags.InCycle) !== 0
          ) {
            flushPendingWasmBatch()
            flushNativeDirectScalarBatch()
            jsCount += 1
            materializeCycleFormulaError(cellIndex)
            return
          }
          if (evaluationOptions.allowCycleDependencyError && hasCycleDependency(cellIndex)) {
            flushPendingWasmBatch()
            flushNativeDirectScalarBatch()
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
            if (
              !evaluationOptions.forceJs &&
              formula.directLookup !== undefined &&
              formula.directScalar === undefined &&
              formula.directAggregate === undefined &&
              formula.directCriteria === undefined &&
              !formula.compiled.producesSpill
            ) {
              if (nativeDirectLookupBatch.count >= MAX_RECALC_NATIVE_DIRECT_LOOKUP_BATCH_SIZE) {
                flushNativeDirectLookupBatch()
              }
              const sheetId = args.state.workbook.cellStore.sheetIds[cellIndex]
              const col = args.state.workbook.cellStore.cols[cellIndex]
              if (
                sheetId !== undefined &&
                col !== undefined &&
                nativeDirectLookupBatch.add({ cellIndex, sheetId, col }, formula.directLookup)
              ) {
                nativeDirectLookupCells.push(cellIndex)
                nativeDirectLookupQueued.add(cellIndex)
                return
              }
              flushNativeDirectLookupBatch()
            }
            if (
              !evaluationOptions.forceJs &&
              formula.directScalar !== undefined &&
              formula.directLookup === undefined &&
              formula.directAggregate === undefined &&
              formula.directCriteria === undefined &&
              formula.compiled.mode === FormulaMode.WasmFastPath &&
              !formula.compiled.producesSpill
            ) {
              if (nativeDirectScalarBatch.count >= MAX_RECALC_NATIVE_DIRECT_SCALAR_BATCH_SIZE) {
                flushNativeDirectScalarBatch()
              }
              if (nativeDirectScalarBatch.add(cellIndex, formula)) {
                nativeDirectScalarCells.push(cellIndex)
                return
              }
              flushNativeDirectScalarBatch()
            } else {
              flushNativeDirectScalarBatch()
            }
            if (evaluateDirectFormulaImmediately(cellIndex, formula)) {
              return
            }
          }
          if (!evaluationOptions.forceJs && formula.compiled.mode === FormulaMode.WasmFastPath && args.state.wasm.ready) {
            flushNativeDirectScalarBatch()
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
          flushNativeDirectLookupBatch()
          flushNativeDirectScalarBatch()
          jsCount += 1
          args.checkEvaluationBudget()
          const spillChanges = evaluateUnsupportedFormulaWithValueCollection(cellIndex)
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
              flushNativeDirectLookupBatch()
              flushNativeDirectScalarBatch()
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
                forceJs: precisionAsDisplayed,
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
              forceJs: precisionAsDisplayed,
            })
          }
        }

        flushNativeDirectLookupBatch()
        flushNativeDirectScalarBatch()
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
        const ordered = filterSkippedCachedFormulaCells(
          toOrderedUint32(singlePassOrdered, singlePassOrderedCount),
          singlePassOrderedCount,
          skippedCachedFormulaCells,
        )
        return totalOrderedCount === 0 && allChangedRoots.length === 0
          ? args.emptyChangedSet()
          : args.composeChangedRootsAndOrdered(allChangedRoots, ordered, ordered.length)
      }
      const ordered = filterSkippedCachedFormulaCells(Uint32Array.from(allOrdered), allOrdered.length, skippedCachedFormulaCells)
      return totalOrderedCount === 0 && allChangedRoots.length === 0
        ? args.emptyChangedSet()
        : args.composeChangedRootsAndOrdered(allChangedRoots, ordered, ordered.length)
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
    options: Pick<RecalculateInternalOptions, 'preserveCachedValuesOnFullRecalc' | 'valueChangeCollector'> = {},
  ): U32 =>
    recalculateInternal(changedRoots, kernelSyncRoots, {
      orderedFormulaCellIndices,
      orderedFormulaCount,
      ...options,
    })

  const reconcilePivotOutputs = (baseChanged: U32, forceAllPivots = false, valueChangeCollector?: RecalcValueChangeCollector): U32 => {
    let aggregate = baseChanged
    let pending = baseChanged
    let forceAll = forceAllPivots

    for (let iteration = 0; iteration < 4; iteration += 1) {
      const pivotChanged = refreshPivotOutputs(pending, forceAll, valueChangeCollector)
      if (pivotChanged.length === 0) {
        break
      }
      aggregate = aggregate.length === 0 ? pivotChanged : args.unionChangedSets(aggregate, pivotChanged)
      pending =
        valueChangeCollector === undefined
          ? recalculateInternal(pivotChanged, pivotChanged)
          : recalculateInternal(pivotChanged, pivotChanged, { valueChangeCollector })
      aggregate = pending.length === 0 ? aggregate : args.unionChangedSets(aggregate, pending)
      forceAll = false
    }

    return aggregate
  }

  const recalculateAllNow = (options: RecalculateAllNowOptions = {}): number[] => {
    args.beginMutationCollection()
    args.state.workbook.setVolatileContext({
      recalcEpoch: args.state.workbook.getVolatileContext().recalcEpoch + 1,
    })
    const valueChangeCollector = options.collectValueChangesForRebuild
      ? createRecalcValueChangeCollector(args.state.workbook.cellStore.size)
      : undefined
    let formulaChangedCount = 0
    let explicitChangedCount = 0
    let canUseFullFormulaOrder = true
    args.state.formulas.forEach((formula, cellIndex) => {
      if (formula.preserveCachedValueOnFullRecalc === true) {
        return
      }
      formulaChangedCount = args.markFormulaChanged(cellIndex, formulaChangedCount)
      if (valueChangeCollector === undefined) {
        explicitChangedCount = args.markExplicitChanged(cellIndex, explicitChangedCount)
      }
      if (formula.compiled.producesSpill || formula.directLookup !== undefined || formula.directCriteria !== undefined) {
        canUseFullFormulaOrder = false
      }
    })
    const mutationRoots = args.composeMutationRoots(0, formulaChangedCount)
    let recalculatedBase: U32
    const fullRecalcOptions =
      valueChangeCollector === undefined
        ? { preserveCachedValuesOnFullRecalc: true as const }
        : { preserveCachedValuesOnFullRecalc: true as const, valueChangeCollector }
    if (canUseFullFormulaOrder) {
      const fullFormulaOrder = args.dirtyScheduler.collectAll()
      recalculatedBase = recalculatePreordered(
        mutationRoots,
        fullFormulaOrder.orderedFormulaCellIndices,
        fullFormulaOrder.orderedFormulaCount,
        args.emptyChangedSet(),
        fullRecalcOptions,
      )
    } else {
      recalculatedBase = recalculateInternal(mutationRoots, args.emptyChangedSet(), fullRecalcOptions)
    }
    const recalculated = reconcilePivotOutputs(recalculatedBase, true, valueChangeCollector)
    const changed =
      valueChangeCollector === undefined
        ? args.composeEventChanges(recalculated, explicitChangedCount)
        : valueChangeCollector.toChangedSet()
    const lastMetrics = { ...args.state.getLastMetrics() }
    lastMetrics.batchId += 1
    lastMetrics.changedInputCount = formulaChangedCount
    args.state.setLastMetrics(lastMetrics)
    emitRecalcBatchEvents({
      state: args.state,
      changed,
      captureChangedCells: args.captureChangedCells,
      metrics: lastMetrics,
      explicitChangedCount,
      captureChangedPatches: args.captureChangedPatches,
    })
    return Array.from(changed)
  }

  const recalculateAllNowSync = (): number[] => {
    try {
      return recalculateAllNow()
    } catch (cause) {
      throw new EngineRecalcError({
        message: 'Failed to recalculate all formulas',
        cause,
      })
    }
  }

  const recalculateChangedValuesNowForRebuildSync = (): number[] => {
    try {
      return recalculateAllNow({ collectValueChangesForRebuild: true })
    } catch (cause) {
      throw new EngineRecalcError({
        message: 'Failed to recalculate changed formula values for rebuild',
        cause,
      })
    }
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
        try: recalculateAllNow,
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
            forEachDirtyRegionCell(args.state.workbook, region, (cellIndex) => {
              changedInputCount = args.markInputChanged(cellIndex, changedInputCount)
              explicitChangedCount = args.markExplicitChanged(cellIndex, explicitChangedCount)
            })
          }

          const changedInputArray = args.getChangedInputBuffer().subarray(0, changedInputCount)
          const recalculated = reconcilePivotOutputs(recalculate(args.composeMutationRoots(changedInputCount, 0), changedInputArray), false)
          const changed = args.composeEventChanges(recalculated, explicitChangedCount)
          const lastMetrics = { ...args.state.getLastMetrics() }
          lastMetrics.batchId += 1
          lastMetrics.changedInputCount = changedInputCount
          args.state.setLastMetrics(lastMetrics)
          emitRecalcBatchEvents({
            state: args.state,
            changed,
            captureChangedCells: args.captureChangedCells,
            metrics: lastMetrics,
            explicitChangedCount,
            captureChangedPatches: args.captureChangedPatches,
          })
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
    recalculateAllNowSync,
    recalculateChangedValuesNowForRebuildSync,
    recalculateNowSync: recalculate,
    reconcilePivotOutputsNow: reconcilePivotOutputs,
  }
}
