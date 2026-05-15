import type { CellRangeRef, EngineEvent } from '@bilig/protocol'
import { addEngineCounter } from '../../perf/engine-counters.js'
import type { U32 } from '../runtime-state.js'
import type { DirectFormulaIndexCollection } from './direct-formula-index-collection.js'
import {
  canEvaluatePostRecalcDirectFormulasWithoutKernel,
  countDirectFormulaDeltaSkip,
  directFormulaChangesAreDisjointFromInputs,
  hasCompleteDirectFormulaDeltas,
} from './direct-formula-recalc-helpers.js'
import { mergeChangedCellIndices } from './operation-change-helpers.js'
import { shouldMaterializeOperationChangedCells } from './operation-event-emission.js'
import { finalizeOperationMutationEvents } from './operation-mutation-event-finalizer.js'
import {
  applyPostRecalcDirectFormulaChanges,
  type ApplyPostRecalcDirectFormulaChangesArgs,
  type DirectFormulaMetricCounts,
} from './operation-post-recalc-direct-formulas.js'
import type { CreateEngineOperationServiceArgs } from './operation-service-types.js'

type OperationPostRecalcDirectFormulaCallbacks = Pick<
  ApplyPostRecalcDirectFormulaChangesArgs,
  | 'applyDirectFormulaCurrentResult'
  | 'applyDirectFormulaNumericDelta'
  | 'applyDirectScalarCurrentValue'
  | 'tryApplyDirectScalarDeltas'
  | 'tryApplyDirectFormulaDeltas'
  | 'countPostRecalcDirectFormulaMetric'
>

type OperationDirtyTraversalSkip = (
  changedInputCellIndices: U32,
  changedInputCount: number,
  postRecalcDirectFormulaIndices?: DirectFormulaIndexCollection,
  options?: {
    readonly lookupHandledInputCellIndices?: readonly number[]
  },
) => boolean

type OperationChangedInputsNeedRegionQueryIndices = (
  changedInputCellIndices: U32,
  changedInputCount: number,
  postRecalcDirectFormulaIndices?: DirectFormulaIndexCollection,
) => boolean

export interface FinalizeOperationRecalcAndEventsArgs {
  readonly serviceArgs: CreateEngineOperationServiceArgs
  readonly isRestore: boolean
  readonly topologyChanged: boolean
  readonly sheetDeleted: boolean
  readonly structuralInvalidation: boolean
  readonly refreshAllPivots: boolean
  readonly changedInputCount: number
  readonly formulaChangedCount: number
  readonly explicitChangedCount: number
  readonly compileMs: number
  readonly precomputedKernelSyncCellIndices: readonly number[]
  readonly postRecalcDirectFormulaIndices: DirectFormulaIndexCollection
  readonly postRecalcDirectFormulaMetrics: DirectFormulaMetricCounts
  readonly lookupHandledInputCellIndices: readonly number[]
  readonly invalidatedRanges: CellRangeRef[]
  readonly invalidatedRows: EngineEvent['invalidatedRows']
  readonly invalidatedColumns: EngineEvent['invalidatedColumns']
  readonly hadCycleMembersBeforeNow: () => boolean
  readonly markCycleMemberInputsChanged: (changedInputCount: number) => number
  readonly canSkipDirtyTraversalForChangedInputs: OperationDirtyTraversalSkip
  readonly changedInputsNeedRegionQueryIndices: OperationChangedInputsNeedRegionQueryIndices
  readonly directFormulaCallbacks: OperationPostRecalcDirectFormulaCallbacks
  readonly shouldMaterializeChangedCells?: (args: {
    readonly changedLength: number
    readonly hasGeneralEventListeners: boolean
    readonly invalidation: EngineEvent['invalidation']
  }) => boolean
}

export function finalizeOperationRecalcAndEvents(input: FinalizeOperationRecalcAndEventsArgs): void {
  let changedInputCount = input.changedInputCount
  let formulaChangedCount = input.formulaChangedCount

  if (input.topologyChanged) {
    const repaired =
      !input.hadCycleMembersBeforeNow() &&
      !input.sheetDeleted &&
      !input.structuralInvalidation &&
      formulaChangedCount > 0 &&
      input.serviceArgs.repairTopoRanks(input.serviceArgs.getChangedFormulaBuffer().subarray(0, formulaChangedCount))
    if (!repaired) {
      input.serviceArgs.rebuildTopoRanks()
      input.serviceArgs.detectCycles()
      changedInputCount = input.markCycleMemberInputsChanged(changedInputCount)
    }
  }

  const hasActiveFormulas = input.serviceArgs.state.formulas.size > 0
  const hasActivePivots = input.serviceArgs.state.workbook.hasPivots()
  const hasGeneralEventListeners = input.serviceArgs.state.events.hasListeners()
  const hasTrackedEventListeners = input.serviceArgs.state.events.hasTrackedListeners()
  const hasWatchedCellListeners = input.serviceArgs.state.events.hasCellListeners()
  const requiresChangedSet = hasGeneralEventListeners || hasTrackedEventListeners || hasWatchedCellListeners
  const hasRecalcWork =
    changedInputCount > 0 ||
    formulaChangedCount > 0 ||
    input.precomputedKernelSyncCellIndices.length > 0 ||
    input.postRecalcDirectFormulaIndices.size > 0
  const hasVolatileFormulaWork =
    hasActiveFormulas && (input.serviceArgs.hasVolatileFormulas ? input.serviceArgs.hasVolatileFormulas() : true)
  const shouldRefreshPivots = input.refreshAllPivots && hasActivePivots
  let recalculated: U32 = new Uint32Array()
  let didRunRecalc = false
  let didFastDeferKernelSyncOnly = false
  let canComposeDisjointEventChanges = false

  const canFastDeferPrecomputedStructuralKernelSync =
    hasActiveFormulas &&
    changedInputCount === 0 &&
    input.explicitChangedCount > 0 &&
    formulaChangedCount === 0 &&
    input.precomputedKernelSyncCellIndices.length > 0 &&
    input.postRecalcDirectFormulaIndices.size === 0 &&
    input.invalidatedRanges.length === 0 &&
    (input.invalidatedRows.length > 0 || input.invalidatedColumns.length > 0) &&
    !input.topologyChanged &&
    !input.structuralInvalidation &&
    !shouldRefreshPivots &&
    !hasActivePivots &&
    !hasVolatileFormulaWork

  if (
    hasActiveFormulas &&
    changedInputCount > 0 &&
    formulaChangedCount === 0 &&
    input.precomputedKernelSyncCellIndices.length === 0 &&
    input.postRecalcDirectFormulaIndices.size === 0 &&
    !input.refreshAllPivots &&
    !hasActivePivots &&
    !hasVolatileFormulaWork
  ) {
    const changedInputArray = input.serviceArgs.getChangedInputBuffer().subarray(0, changedInputCount)
    if (
      input.canSkipDirtyTraversalForChangedInputs(changedInputArray, changedInputCount, input.postRecalcDirectFormulaIndices, {
        lookupHandledInputCellIndices: input.lookupHandledInputCellIndices,
      })
    ) {
      addEngineCounter(input.serviceArgs.state.counters, 'kernelSyncOnlyRecalcSkips')
      input.serviceArgs.deferKernelSync(changedInputArray)
      didFastDeferKernelSyncOnly = true
    }
  }

  if (!didFastDeferKernelSyncOnly && canFastDeferPrecomputedStructuralKernelSync) {
    addEngineCounter(input.serviceArgs.state.counters, 'kernelSyncOnlyRecalcSkips')
    input.serviceArgs.deferKernelSync(Uint32Array.from(input.precomputedKernelSyncCellIndices))
    didFastDeferKernelSyncOnly = true
  }

  if (
    !didFastDeferKernelSyncOnly &&
    ((hasActiveFormulas && (hasRecalcWork || hasVolatileFormulaWork)) || (hasActivePivots && hasRecalcWork) || shouldRefreshPivots)
  ) {
    formulaChangedCount = input.serviceArgs.markVolatileFormulasChanged(formulaChangedCount)
    const changedInputArray = input.serviceArgs.getChangedInputBuffer().subarray(0, changedInputCount)
    if (input.changedInputsNeedRegionQueryIndices(changedInputArray, changedInputCount, input.postRecalcDirectFormulaIndices)) {
      input.serviceArgs.prepareRegionQueryIndices()
    }
    const canUseKernelSyncOnlyRecalc =
      formulaChangedCount === 0 &&
      changedInputCount > 0 &&
      input.precomputedKernelSyncCellIndices.length === 0 &&
      !input.refreshAllPivots &&
      input.canSkipDirtyTraversalForChangedInputs(changedInputArray, changedInputCount, input.postRecalcDirectFormulaIndices, {
        lookupHandledInputCellIndices: input.lookupHandledInputCellIndices,
      })
    const canDeferKernelSyncOnlyRecalc = canUseKernelSyncOnlyRecalc && input.postRecalcDirectFormulaIndices.size === 0
    const canSkipKernelSyncOnlyRecalc = canUseKernelSyncOnlyRecalc && input.postRecalcDirectFormulaIndices.size > 0
    const canSkipRecalcForDirectDeltas = canSkipKernelSyncOnlyRecalc && hasCompleteDirectFormulaDeltas(input.postRecalcDirectFormulaIndices)
    const canSkipRecalcForDirectEvaluation =
      canSkipKernelSyncOnlyRecalc &&
      canEvaluatePostRecalcDirectFormulasWithoutKernel(input.serviceArgs.state.formulas, input.postRecalcDirectFormulaIndices)
    const canUseDisjointDirectEventChanges =
      (canSkipRecalcForDirectDeltas || canSkipRecalcForDirectEvaluation) &&
      input.explicitChangedCount === changedInputCount &&
      !hasActivePivots &&
      !input.refreshAllPivots &&
      directFormulaChangesAreDisjointFromInputs(changedInputArray, changedInputCount, input.postRecalcDirectFormulaIndices)

    if (canDeferKernelSyncOnlyRecalc) {
      addEngineCounter(input.serviceArgs.state.counters, 'kernelSyncOnlyRecalcSkips')
      input.serviceArgs.deferKernelSync(changedInputArray)
    } else if (!canSkipKernelSyncOnlyRecalc) {
      input.serviceArgs.prepareRegionQueryIndices()
      const changedRoots = canUseKernelSyncOnlyRecalc
        ? new Uint32Array()
        : input.serviceArgs.composeMutationRoots(changedInputCount, formulaChangedCount)
      const kernelSyncRoots =
        input.precomputedKernelSyncCellIndices.length === 0
          ? changedInputArray
          : Uint32Array.from([...changedInputArray, ...input.precomputedKernelSyncCellIndices])
      recalculated = input.serviceArgs.recalculate(changedRoots, kernelSyncRoots)
      didRunRecalc = true
    } else if (canSkipRecalcForDirectDeltas) {
      countDirectFormulaDeltaSkip(input.serviceArgs.state.formulas, input.postRecalcDirectFormulaIndices, input.serviceArgs.state.counters)
      input.serviceArgs.deferKernelSync(changedInputArray)
    } else if (canSkipRecalcForDirectEvaluation) {
      addEngineCounter(input.serviceArgs.state.counters, 'directFormulaKernelSyncOnlyRecalcSkips')
      input.serviceArgs.deferKernelSync(changedInputArray)
    } else {
      input.serviceArgs.prepareRegionQueryIndices()
      input.serviceArgs.recalculate(new Uint32Array(), changedInputArray)
    }

    if (input.postRecalcDirectFormulaIndices.size > 0) {
      recalculated = applyPostRecalcDirectFormulaChanges({
        state: input.serviceArgs.state,
        collection: input.postRecalcDirectFormulaIndices,
        recalculated,
        didRunRecalc,
        captureChanged: requiresChangedSet || hasActivePivots || shouldRefreshPivots,
        metrics: input.postRecalcDirectFormulaMetrics,
        ...input.directFormulaCallbacks,
        evaluateDirectFormula: input.serviceArgs.evaluateDirectFormula,
      })
    }

    if (hasActivePivots || shouldRefreshPivots) {
      const pivotRefreshRoots =
        shouldRefreshPivots || changedInputArray.length === 0 ? recalculated : mergeChangedCellIndices(recalculated, changedInputArray)
      recalculated = input.serviceArgs.reconcilePivotOutputs(pivotRefreshRoots, shouldRefreshPivots)
    } else if (canUseDisjointDirectEventChanges) {
      canComposeDisjointEventChanges = true
    }
  }

  const invalidation = input.isRestore || input.sheetDeleted || input.structuralInvalidation ? 'full' : 'cells'

  finalizeOperationMutationEvents({
    serviceArgs: input.serviceArgs,
    suppressChangedSet: input.isRestore || invalidation === 'full' || !requiresChangedSet,
    canComposeDisjointEventChanges,
    recalculated,
    explicitChangedCount: input.explicitChangedCount,
    changedInputCount,
    formulaChangedCount,
    compileMs: input.compileMs,
    didRunRecalc,
    directFormulaMetrics: input.postRecalcDirectFormulaMetrics,
    invalidation,
    invalidatedRanges: input.invalidatedRanges,
    invalidatedRows: input.invalidatedRows,
    invalidatedColumns: input.invalidatedColumns,
    hasGeneralEventListeners,
    hasTrackedEventListeners,
    hasWatchedCellListeners,
    shouldMaterializeChangedCells: (changedLength) =>
      input.shouldMaterializeChangedCells?.({
        changedLength,
        hasGeneralEventListeners,
        invalidation,
      }) ??
      shouldMaterializeOperationChangedCells({
        changedLength,
        hasGeneralEventListeners,
        invalidation,
        invalidations: {
          invalidatedRanges: input.invalidatedRanges,
          invalidatedRows: input.invalidatedRows,
          invalidatedColumns: input.invalidatedColumns,
        },
      }),
  })
}
