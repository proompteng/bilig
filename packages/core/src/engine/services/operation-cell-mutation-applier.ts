import type { EngineOpBatch } from '@bilig/workbook-domain'
import type { CellValue } from '@bilig/protocol'
import type { EngineCellMutationRef } from '../../cell-mutations-at.js'
import { batchOpOrder, markBatchApplied, type OpOrder } from '../../replica-state.js'
import { DirectFormulaIndexCollection } from './direct-formula-index-collection.js'
import { aggregateColumnDependencyKey } from './direct-formula-recalc-helpers.js'
import { assertNever } from './operation-change-helpers.js'
import { createOperationTrackedColumnDependencyFlagResolver } from './operation-column-dependency-tracker.js'
import { createOperationSheetNameResolver, resolveOperationExistingMutationCellIndex } from './operation-cell-address-resolver.js'
import type { OperationLookupAccess } from './operation-lookup-access.js'
import type { ExactLookupImpactCaches, OperationLookupDirtyMarkerService } from './operation-lookup-dirty-markers.js'
import type { OperationColumnDependencyTrackerService } from './operation-column-dependency-tracker.js'
import type { OperationLookupPlanner } from './operation-lookup-planner.js'
import type { OperationDirectRangeDependentService } from './operation-direct-range-dependents.js'
import type { DirectFormulaMetricCounts } from './operation-post-recalc-direct-formulas.js'
import { finalizeOperationRecalcAndEvents } from './operation-recalc-finalizer.js'
import type { CreateEngineOperationServiceArgs, MutationSource } from './operation-service-types.js'
import { createOperationFreshDirectAggregateFormulaBatchFastPath } from './operation-fresh-direct-aggregate-formula-batch-fast-path.js'
import { createOperationFreshDirectScalarFormulaBatchFastPath } from './operation-fresh-direct-scalar-formula-batch-fast-path.js'
import { createOperationExistingRectangularLiteralBatchFastPath } from './operation-existing-rectangular-literal-batch-fast-path.js'
import { applyClearCellMutation } from './operation-clear-cell-mutation.js'
import { applySetCellValueMutation } from './operation-set-cell-value-mutation.js'
import { applySetCellFormulaMutation } from './operation-set-cell-formula-mutation.js'

type OperationCellMutationSource = Exclude<MutationSource, 'remote'>
type OperationCellDirectFormulaCallbacks = Parameters<typeof finalizeOperationRecalcAndEvents>[0]['directFormulaCallbacks']
type OperationCellDirtyTraversalSkip = Parameters<typeof finalizeOperationRecalcAndEvents>[0]['canSkipDirtyTraversalForChangedInputs']
type OperationCellCycleInputMarker = Parameters<typeof finalizeOperationRecalcAndEvents>[0]['markCycleMemberInputsChanged']

interface CreateOperationCellMutationApplierArgs {
  readonly serviceArgs: CreateEngineOperationServiceArgs
  readonly emitBatch: (batch: EngineOpBatch) => void
  readonly setCellEntityVersion: (sheetName: string, address: string, order: OpOrder) => void
  readonly isNullLiteralWriteNoOp: (cellIndex: number) => boolean
  readonly isClearCellNoOp: (cellIndex: number) => boolean
  readonly canFastPathLiteralOverwrite: (cellIndex: number) => boolean
  readonly tryApplySingleExistingDirectLiteralMutation: (
    refs: readonly EngineCellMutationRef[],
    batch: EngineOpBatch | null,
    source: OperationCellMutationSource,
  ) => boolean
  readonly tryApplyCoalescedDirectScalarLiteralBatch: (
    refs: readonly EngineCellMutationRef[],
    batch: EngineOpBatch | null,
    source: OperationCellMutationSource,
    potentialNewCells?: number,
  ) => boolean
  readonly tryApplyFreshDenseRectangularNumericLiteralBatch: (
    refs: readonly EngineCellMutationRef[],
    batch: EngineOpBatch | null,
    potentialNewCells?: number,
  ) => boolean
  readonly writeNumericLiteralToCellStore: (cellIndex: number, value: number) => void
  readonly readCellValueForLookup: OperationLookupAccess['readCellValueForLookup']
  readonly readApproximateNumericValueForLookup: OperationLookupAccess['readApproximateNumericValueForLookup']
  readonly readExactNumericValueForLookup: OperationLookupAccess['readExactNumericValueForLookup']
  readonly hasTrackedExactLookupDependents: OperationColumnDependencyTrackerService['hasTrackedExactLookupDependents']
  readonly hasTrackedSortedLookupDependents: OperationColumnDependencyTrackerService['hasTrackedSortedLookupDependents']
  readonly hasTrackedDirectRangeDependents: OperationColumnDependencyTrackerService['hasTrackedDirectRangeDependents']
  readonly planExactLookupNumericColumnWrite: OperationLookupPlanner['planExactLookupNumericColumnWrite']
  readonly planApproximateLookupNumericColumnWrite: OperationLookupPlanner['planApproximateLookupNumericColumnWrite']
  readonly noteExactLookupLiteralWriteWhenDirty: OperationLookupDirtyMarkerService['noteExactLookupLiteralWriteWhenDirty']
  readonly noteSortedLookupLiteralWriteWhenDirty: OperationLookupDirtyMarkerService['noteSortedLookupLiteralWriteWhenDirty']
  readonly markAffectedDirectRangeDependents: OperationDirectRangeDependentService['markAffectedDirectRangeDependents']
  readonly collectAffectedDirectRangeDependents: OperationDirectRangeDependentService['collectAffectedDirectRangeDependents']
  readonly markPostRecalcDirectFormulaDependents: (
    cellIndex: number,
    postRecalcDirectFormulaIndices: DirectFormulaIndexCollection,
    oldValue?: CellValue,
    newValue?: CellValue,
  ) => boolean
  readonly markDirectScalarDeltaClosure: (
    rootCellIndex: number,
    oldValue: CellValue,
    newValue: CellValue,
    postRecalcDirectFormulaIndices: DirectFormulaIndexCollection,
  ) => void
  readonly markPostRecalcDirectScalarNumericDependents: (
    cellIndex: number,
    oldNumber: number,
    newNumber: number,
    postRecalcDirectFormulaIndices: DirectFormulaIndexCollection,
    exactLookupValue?: number,
    approximateLookupValue?: number,
  ) => boolean
  readonly markPostRecalcDirectLookupCurrentDependentsFromNumeric: (
    cellIndex: number,
    exactLookupValue: number | undefined,
    approximateLookupValue: number | undefined,
    postRecalcDirectFormulaIndices: DirectFormulaIndexCollection,
  ) => boolean
  readonly directScalarCellNumericValue: (cellIndex: number) => number | undefined
  readonly tryApplyFormulaReplacementAsDirectScalarDeltaRoot: (request: {
    readonly cellIndex: number
    readonly oldNumber: number | undefined
    readonly changedTopology: boolean
    readonly postRecalcDirectFormulaIndices: DirectFormulaIndexCollection
  }) => boolean
  readonly rebindDynamicFormulaDependents: (cellIndex: number, formulaChangedCount: number) => number
  readonly markCycleMemberInputsChanged: OperationCellCycleInputMarker
  readonly hasCycleMembersNow: () => boolean
  readonly canSkipDirtyTraversalForChangedInputs: OperationCellDirtyTraversalSkip
  readonly directFormulaCallbacks: OperationCellDirectFormulaCallbacks
  readonly pruneCellIfOrphaned: (cellIndex: number) => void
  readonly normalizeHistoryDependencyPlaceholder: (cellIndex: number, source: MutationSource) => void
}

export function createOperationCellMutationApplier(input: CreateOperationCellMutationApplierArgs) {
  const {
    serviceArgs: args,
    emitBatch,
    setCellEntityVersion,
    isNullLiteralWriteNoOp,
    isClearCellNoOp,
    canFastPathLiteralOverwrite,
    tryApplySingleExistingDirectLiteralMutation,
    tryApplyCoalescedDirectScalarLiteralBatch,
    tryApplyFreshDenseRectangularNumericLiteralBatch,
    writeNumericLiteralToCellStore,
    readCellValueForLookup,
    readApproximateNumericValueForLookup,
    readExactNumericValueForLookup,
    hasTrackedExactLookupDependents,
    hasTrackedSortedLookupDependents,
    hasTrackedDirectRangeDependents,
    planExactLookupNumericColumnWrite,
    planApproximateLookupNumericColumnWrite,
    noteExactLookupLiteralWriteWhenDirty,
    noteSortedLookupLiteralWriteWhenDirty,
    markAffectedDirectRangeDependents,
    collectAffectedDirectRangeDependents,
    markPostRecalcDirectFormulaDependents,
    markDirectScalarDeltaClosure,
    markPostRecalcDirectScalarNumericDependents,
    markPostRecalcDirectLookupCurrentDependentsFromNumeric,
    rebindDynamicFormulaDependents,
    directScalarCellNumericValue,
    tryApplyFormulaReplacementAsDirectScalarDeltaRoot,
    markCycleMemberInputsChanged,
    hasCycleMembersNow,
    canSkipDirtyTraversalForChangedInputs,
    directFormulaCallbacks: {
      applyDirectFormulaCurrentResult,
      applyDirectFormulaNumericDelta,
      applyDirectScalarCurrentValue,
      tryApplyDirectScalarDeltas,
      tryApplyDirectFormulaDeltas,
      countPostRecalcDirectFormulaMetric,
    },
    pruneCellIfOrphaned,
    normalizeHistoryDependencyPlaceholder,
  } = input

  const freshDirectAggregateFormulaBatchFastPath =
    args.bindPreparedFormula === undefined || args.compileTemplateFormula === undefined
      ? undefined
      : createOperationFreshDirectAggregateFormulaBatchFastPath({
          state: args.state,
          emitBatch,
          setCellEntityVersion,
          hasTrackedExactLookupDependents,
          hasTrackedSortedLookupDependents,
          hasTrackedDirectRangeDependents,
          ...(args.hasRegionFormulaSubscriptionsOverlappingRange === undefined
            ? {}
            : { hasRegionFormulaSubscriptionsOverlappingRange: args.hasRegionFormulaSubscriptionsOverlappingRange }),
          ...(args.getRegionFormulaSubscriptionCount === undefined
            ? {}
            : { getRegionFormulaSubscriptionCount: args.getRegionFormulaSubscriptionCount }),
          hasRegionFormulaSubscriptionsIntersectingRect: args.hasRegionFormulaSubscriptionsIntersectingRect,
          bindPreparedFormula: args.bindPreparedFormula,
          bindFreshDirectAggregateFormulaRun: args.bindFreshDirectAggregateFormulaRun,
          registerFreshFormulaFamilyRun: args.registerFreshFormulaFamilyRun,
          upsertFormulaFamilyRun: args.upsertFormulaFamilyRun,
          upsertFreshFormulaInstances: args.upsertFreshFormulaInstances,
          compileTemplateFormula: args.compileTemplateFormula,
          materializeDeferredStructuralFormulaSources: args.materializeDeferredStructuralFormulaSources,
          checkEvaluationBudget: args.checkEvaluationBudget,
          beginMutationCollection: args.beginMutationCollection,
          ensureRecalcScratchCapacity: args.ensureRecalcScratchCapacity,
          resetMaterializedCellScratch: args.resetMaterializedCellScratch,
          getBatchMutationDepth: args.getBatchMutationDepth,
          setBatchMutationDepth: args.setBatchMutationDepth,
          markInputChanged: args.markInputChanged,
          markExplicitChanged: args.markExplicitChanged,
          getChangedInputBuffer: args.getChangedInputBuffer,
          deferKernelSync: args.deferKernelSync,
          captureChangedCells: args.captureChangedCells,
          applyDirectFormulaCurrentResult,
        })
  const freshDirectScalarFormulaBatchFastPath =
    args.bindFreshDirectScalarFormulaRun === undefined || args.compileTemplateFormula === undefined
      ? undefined
      : createOperationFreshDirectScalarFormulaBatchFastPath({
          state: args.state,
          emitBatch,
          setCellEntityVersion,
          hasTrackedExactLookupDependents,
          hasTrackedSortedLookupDependents,
          hasTrackedDirectRangeDependents,
          ...(args.hasRegionFormulaSubscriptionsOverlappingRange === undefined
            ? {}
            : { hasRegionFormulaSubscriptionsOverlappingRange: args.hasRegionFormulaSubscriptionsOverlappingRange }),
          ...(args.getRegionFormulaSubscriptionCount === undefined
            ? {}
            : { getRegionFormulaSubscriptionCount: args.getRegionFormulaSubscriptionCount }),
          hasRegionFormulaSubscriptionsIntersectingRect: args.hasRegionFormulaSubscriptionsIntersectingRect,
          bindFreshDirectScalarFormulaRun: args.bindFreshDirectScalarFormulaRun,
          registerFreshFormulaFamilyRun: args.registerFreshFormulaFamilyRun,
          upsertFormulaFamilyRun: args.upsertFormulaFamilyRun,
          upsertFreshFormulaInstances: args.upsertFreshFormulaInstances,
          compileTemplateFormula: args.compileTemplateFormula,
          materializeDeferredStructuralFormulaSources: args.materializeDeferredStructuralFormulaSources,
          checkEvaluationBudget: args.checkEvaluationBudget,
          beginMutationCollection: args.beginMutationCollection,
          ensureRecalcScratchCapacity: args.ensureRecalcScratchCapacity,
          resetMaterializedCellScratch: args.resetMaterializedCellScratch,
          getBatchMutationDepth: args.getBatchMutationDepth,
          setBatchMutationDepth: args.setBatchMutationDepth,
          markInputChanged: args.markInputChanged,
          markExplicitChanged: args.markExplicitChanged,
          getChangedInputBuffer: args.getChangedInputBuffer,
          deferKernelSync: args.deferKernelSync,
          captureChangedCells: args.captureChangedCells,
          applyDirectFormulaCurrentResult,
        })
  const existingRectangularLiteralBatchFastPath = createOperationExistingRectangularLiteralBatchFastPath({
    state: args.state,
    emitBatch,
    canFastPathLiteralOverwrite,
    writeNumericLiteralToCellStore,
    materializeDeferredStructuralFormulaSources: args.materializeDeferredStructuralFormulaSources,
    getBatchMutationDepth: args.getBatchMutationDepth,
    setBatchMutationDepth: args.setBatchMutationDepth,
    deferKernelSync: args.deferKernelSync,
    captureChangedCells: args.captureChangedCells,
  })

  return function applyCellMutationsAtNow(
    refs: readonly EngineCellMutationRef[],
    batch: EngineOpBatch | null,
    source: 'local' | 'restore' | 'undo' | 'redo',
    potentialNewCells?: number,
  ): void {
    const isRestore = source === 'restore'
    if (tryApplySingleExistingDirectLiteralMutation(refs, batch, source)) {
      return
    }
    if (freshDirectAggregateFormulaBatchFastPath?.tryApplyFreshDirectAggregateFormulaMatrixBatch(refs, batch, source, potentialNewCells)) {
      return
    }
    if (freshDirectScalarFormulaBatchFastPath?.tryApplyFreshDirectScalarFormulaMatrixBatch(refs, batch, source, potentialNewCells)) {
      return
    }
    if (
      source === 'local' &&
      existingRectangularLiteralBatchFastPath.tryApplyExistingDenseRectangularNumericLiteralBatch(refs, batch, potentialNewCells)
    ) {
      return
    }
    if (source === 'local' && tryApplyFreshDenseRectangularNumericLiteralBatch(refs, batch, potentialNewCells)) {
      return
    }
    if (tryApplyCoalescedDirectScalarLiteralBatch(refs, batch, source, potentialNewCells)) {
      return
    }
    if (freshDirectAggregateFormulaBatchFastPath?.tryApplyFreshDirectAggregateFormulaBatch(refs, batch, source, potentialNewCells)) {
      return
    }
    args.materializeDeferredStructuralFormulaSources()
    args.beginMutationCollection()
    let changedInputCount = 0
    let formulaChangedCount = 0
    let explicitChangedCount = 0
    let topologyChanged = false
    let compileMs = 0
    const postRecalcDirectFormulaIndices = new DirectFormulaIndexCollection()
    const postRecalcDirectFormulaMetrics: DirectFormulaMetricCounts = {
      wasmFormulaCount: 0,
      jsFormulaCount: 0,
    }
    const lookupHandledInputCellIndices: number[] = []
    const pendingExactLookupInvalidations = new Map<number, { sheetName: string; col: number }>()
    const pendingSortedLookupInvalidations = new Map<number, { sheetName: string; col: number }>()
    const queueHandledLookupInvalidation = (sheetId: number, sheetName: string, col: number, exact: boolean, sorted: boolean): void => {
      const key = aggregateColumnDependencyKey(sheetId, col)
      if (exact) {
        pendingExactLookupInvalidations.set(key, { sheetName, col })
      }
      if (sorted) {
        pendingSortedLookupInvalidations.set(key, { sheetName, col })
      }
    }
    const hasGeneralEventListeners = args.state.events.hasListeners()
    const hasTrackedEventListeners = args.state.events.hasTrackedListeners()
    const hasWatchedCellListeners = args.state.events.hasCellListeners()
    const requiresChangedSet = hasGeneralEventListeners || hasTrackedEventListeners || hasWatchedCellListeners
    const trackExplicitChanges = !isRestore && requiresChangedSet
    const batchMayNeedFreshAggregateInputCoverage = refs.some(
      (ref) => ref.mutation.kind === 'setCellValue' || ref.mutation.kind === 'clearCell',
    )
    let hadCycleMembersBefore: boolean | undefined
    const hadCycleMembersBeforeNow = (): boolean => (hadCycleMembersBefore ??= hasCycleMembersNow())
    const reservedNewCells = potentialNewCells ?? refs.length
    args.state.workbook.cellStore.ensureCapacity(args.state.workbook.cellStore.size + reservedNewCells)
    args.ensureRecalcScratchCapacity(args.state.workbook.cellStore.size + reservedNewCells + 1)
    args.resetMaterializedCellScratch(reservedNewCells)

    const sheetNameResolver = createOperationSheetNameResolver(args.state.workbook)
    const trackedColumnDependencyFlags = createOperationTrackedColumnDependencyFlagResolver({
      hasTrackedExactLookupDependents,
      hasTrackedSortedLookupDependents,
      hasTrackedDirectRangeDependents,
    })
    const exactLookupImpactCaches: ExactLookupImpactCaches = new Map()
    const clearTrackedColumnDependencyFlagCache = (): void => {
      trackedColumnDependencyFlags.clear()
      exactLookupImpactCaches.clear()
    }
    args.setBatchMutationDepth(args.getBatchMutationDepth() + 1)
    try {
      args.state.workbook.withBatchedColumnVersionUpdates(() => {
        refs.forEach((ref, refIndex) => {
          args.checkEvaluationBudget()
          const { sheetId, mutation } = ref
          const order = args.state.trackReplicaVersions && batch ? batchOpOrder(batch, refIndex) : undefined
          const existingIndex = resolveOperationExistingMutationCellIndex(args.state.workbook, ref)

          switch (mutation.kind) {
            case 'setCellValue': {
              const sheetName = sheetNameResolver.resolve(sheetId)
              const setValueResult = applySetCellValueMutation({
                serviceArgs: args,
                sheetId,
                sheetName,
                mutation,
                existingIndex,
                isRestore,
                trackExplicitChanges,
                order,
                changedInputCount,
                formulaChangedCount,
                explicitChangedCount,
                topologyChanged,
                dependencyFlags: trackedColumnDependencyFlags.resolve(sheetId, mutation.col),
                postRecalcDirectFormulaIndices,
                exactLookupImpactCaches,
                setCellEntityVersion,
                isNullLiteralWriteNoOp,
                canFastPathLiteralOverwrite,
                readCellValueForLookup,
                readApproximateNumericValueForLookup,
                readExactNumericValueForLookup,
                planExactLookupNumericColumnWrite,
                planApproximateLookupNumericColumnWrite,
                rebindValueSensitiveFormulaDependents: (cellIndex, counts) => {
                  const reboundCount = counts.formulaChangedCount
                  const nextFormulaChangedCount = rebindDynamicFormulaDependents(cellIndex, counts.formulaChangedCount)
                  return {
                    ...counts,
                    formulaChangedCount: nextFormulaChangedCount,
                    topologyChanged: counts.topologyChanged || nextFormulaChangedCount !== reboundCount,
                  }
                },
                markPostRecalcDirectFormulaDependents,
                markDirectScalarDeltaClosure,
                markPostRecalcDirectScalarNumericDependents,
                markPostRecalcDirectLookupCurrentDependentsFromNumeric,
                directScalarCellNumericValue,
                noteExactLookupLiteralWriteWhenDirty,
                noteSortedLookupLiteralWriteWhenDirty,
                markAffectedDirectRangeDependents,
                queueHandledLookupInvalidation,
                noteHandledLookupInputCellIndex: (cellIndex) => lookupHandledInputCellIndices.push(cellIndex),
                clearTrackedColumnDependencyFlagCache,
                pruneCellIfOrphaned,
              })
              changedInputCount = setValueResult.changedInputCount
              formulaChangedCount = setValueResult.formulaChangedCount
              explicitChangedCount = setValueResult.explicitChangedCount
              topologyChanged = setValueResult.topologyChanged
              break
            }
            case 'setCellFormula': {
              const sheetName = sheetNameResolver.resolve(sheetId)
              const formulaResult = applySetCellFormulaMutation({
                serviceArgs: args,
                sheetId,
                sheetName,
                mutation,
                existingIndex,
                isRestore,
                trackExplicitChanges,
                order,
                changedInputCount,
                formulaChangedCount,
                explicitChangedCount,
                topologyChanged,
                compileMs,
                dependencyFlags: trackedColumnDependencyFlags.resolve(sheetId, mutation.col),
                batchMayNeedFreshAggregateInputCoverage,
                postRecalcDirectFormulaIndices,
                setCellEntityVersion,
                readExactNumericValueForLookup,
                clearTrackedColumnDependencyFlagCache,
                applyDirectFormulaCurrentResult,
                tryApplyFormulaReplacementAsDirectScalarDeltaRoot,
                collectAffectedDirectRangeDependents,
              })
              changedInputCount = formulaResult.changedInputCount
              formulaChangedCount = formulaResult.formulaChangedCount
              explicitChangedCount = formulaResult.explicitChangedCount
              topologyChanged = formulaResult.topologyChanged
              compileMs = formulaResult.compileMs
              break
            }
            case 'clearCell': {
              const clearResult = applyClearCellMutation({
                serviceArgs: args,
                sheetId,
                mutation,
                existingIndex,
                source,
                isRestore,
                trackExplicitChanges,
                order,
                changedInputCount,
                formulaChangedCount,
                explicitChangedCount,
                topologyChanged,
                dependencyFlags: trackedColumnDependencyFlags.resolve(sheetId, mutation.col),
                postRecalcDirectFormulaIndices,
                exactLookupImpactCaches,
                resolveSheetName: (id) => sheetNameResolver.resolve(id),
                setCellEntityVersion,
                isClearCellNoOp,
                canFastPathLiteralOverwrite,
                readCellValueForLookup,
                rebindValueSensitiveFormulaDependents: (cellIndex, counts) => {
                  const reboundCount = counts.formulaChangedCount
                  const nextFormulaChangedCount = rebindDynamicFormulaDependents(cellIndex, counts.formulaChangedCount)
                  return {
                    ...counts,
                    formulaChangedCount: nextFormulaChangedCount,
                    topologyChanged: counts.topologyChanged || nextFormulaChangedCount !== reboundCount,
                  }
                },
                markPostRecalcDirectFormulaDependents,
                markDirectScalarDeltaClosure,
                noteExactLookupLiteralWriteWhenDirty,
                noteSortedLookupLiteralWriteWhenDirty,
                markAffectedDirectRangeDependents,
                clearTrackedColumnDependencyFlagCache,
                pruneCellIfOrphaned,
                normalizeHistoryDependencyPlaceholder,
              })
              changedInputCount = clearResult.changedInputCount
              formulaChangedCount = clearResult.formulaChangedCount
              explicitChangedCount = clearResult.explicitChangedCount
              topologyChanged = clearResult.topologyChanged
              break
            }
            default:
              assertNever(mutation)
          }
        })
      })

      const reboundCount = formulaChangedCount
      formulaChangedCount = args.syncDynamicRanges(formulaChangedCount)
      topologyChanged = topologyChanged || formulaChangedCount !== reboundCount
    } finally {
      args.setBatchMutationDepth(args.getBatchMutationDepth() - 1)
    }

    pendingExactLookupInvalidations.forEach((entry) => args.invalidateExactLookupColumn(entry))
    pendingSortedLookupInvalidations.forEach((entry) => args.invalidateSortedLookupColumn(entry))

    if (batch) {
      markBatchApplied(args.state.replicaState, batch)
    }
    if (refs.length === 0) {
      if (!isRestore && batch) {
        emitBatch(batch)
      }
      return
    }

    finalizeOperationRecalcAndEvents({
      serviceArgs: args,
      isRestore,
      topologyChanged,
      sheetDeleted: false,
      structuralInvalidation: false,
      refreshAllPivots: false,
      explicitChangedCount,
      changedInputCount,
      formulaChangedCount,
      compileMs,
      precomputedKernelSyncCellIndices: [],
      postRecalcDirectFormulaIndices,
      postRecalcDirectFormulaMetrics,
      lookupHandledInputCellIndices,
      invalidatedRanges: [],
      invalidatedRows: [],
      invalidatedColumns: [],
      invalidatedRowCount: 0,
      invalidatedColumnCount: 0,
      hadCycleMembersBeforeNow,
      markCycleMemberInputsChanged,
      canSkipDirtyTraversalForChangedInputs,
      directFormulaCallbacks: {
        applyDirectFormulaCurrentResult,
        applyDirectFormulaNumericDelta,
        applyDirectScalarCurrentValue,
        tryApplyDirectScalarDeltas,
        tryApplyDirectFormulaDeltas,
        countPostRecalcDirectFormulaMetric,
      },
      shouldMaterializeChangedCells: (listenerState) => listenerState.hasGeneralEventListeners,
    })
    if (isRestore && !hasTrackedEventListeners) {
      return
    }
    if (batch) {
      void args.state.getSyncClientConnection()?.send(batch)
      emitBatch(batch)
    }
  }
}
