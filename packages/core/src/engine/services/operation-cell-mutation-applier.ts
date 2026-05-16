import { formatAddress } from '@bilig/formula'
import type { EngineOpBatch } from '@bilig/workbook-domain'
import type { CellValue } from '@bilig/protocol'
import type { EngineCellMutationRef } from '../../cell-mutations-at.js'
import { batchOpOrder, markBatchApplied, type OpOrder } from '../../replica-state.js'
import { CellFlags } from '../../cell-store.js'
import { emptyValue, literalToValue, writeLiteralToCellStore } from '../../engine-value-utils.js'
import { exactLookupLiteralNumericValue, withOptionalLookupStringIds } from './direct-lookup-helpers.js'
import { DirectFormulaIndexCollection } from './direct-formula-index-collection.js'
import { directScalarLiteralNumericValue } from './direct-scalar-helpers.js'
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
import { analyzeFreshDirectAggregateFormula, bindFreshTemplateFormula } from './operation-fresh-direct-aggregate.js'

type OperationCellMutationSource = Exclude<MutationSource, 'remote'>
type OperationCellDirectFormulaCallbacks = Parameters<typeof finalizeOperationRecalcAndEvents>[0]['directFormulaCallbacks']
type OperationCellDirtyTraversalSkip = Parameters<typeof finalizeOperationRecalcAndEvents>[0]['canSkipDirtyTraversalForChangedInputs']
type OperationCellChangedInputsNeedRegionQueryIndices = Parameters<
  typeof finalizeOperationRecalcAndEvents
>[0]['changedInputsNeedRegionQueryIndices']
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
  readonly readCellValueForLookup: OperationLookupAccess['readCellValueForLookup']
  readonly readApproximateNumericValueForLookup: OperationLookupAccess['readApproximateNumericValueForLookup']
  readonly readExactNumericValueForLookup: OperationLookupAccess['readExactNumericValueForLookup']
  readonly hasTrackedExactLookupDependents: OperationColumnDependencyTrackerService['hasTrackedExactLookupDependents']
  readonly hasTrackedSortedLookupDependents: OperationColumnDependencyTrackerService['hasTrackedSortedLookupDependents']
  readonly hasTrackedDirectRangeDependents: OperationColumnDependencyTrackerService['hasTrackedDirectRangeDependents']
  readonly canSkipExactLookupNumericColumnWrite: OperationLookupPlanner['canSkipExactLookupNumericColumnWrite']
  readonly canSkipApproximateLookupNumericColumnWrite: OperationLookupPlanner['canSkipApproximateLookupNumericColumnWrite']
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
    readonly postRecalcDirectFormulaMetrics: DirectFormulaMetricCounts
  }) => boolean
  readonly rebindDynamicFormulaDependents: (cellIndex: number, formulaChangedCount: number) => number
  readonly markCycleMemberInputsChanged: OperationCellCycleInputMarker
  readonly hasCycleMembersNow: () => boolean
  readonly canSkipDirtyTraversalForChangedInputs: OperationCellDirtyTraversalSkip
  readonly changedInputsNeedRegionQueryIndices: OperationCellChangedInputsNeedRegionQueryIndices
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
    readCellValueForLookup,
    readApproximateNumericValueForLookup,
    readExactNumericValueForLookup,
    hasTrackedExactLookupDependents,
    hasTrackedSortedLookupDependents,
    hasTrackedDirectRangeDependents,
    canSkipExactLookupNumericColumnWrite,
    canSkipApproximateLookupNumericColumnWrite,
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
    changedInputsNeedRegionQueryIndices,
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
    if (tryApplyCoalescedDirectScalarLiteralBatch(refs, batch, source, potentialNewCells)) {
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
    let hadCycleMembersBefore: boolean | undefined
    const hadCycleMembersBeforeNow = (): boolean => (hadCycleMembersBefore ??= hasCycleMembersNow())
    const rebindValueSensitiveFormulaDependents = (cellIndex: number): void => {
      const reboundCount = formulaChangedCount
      formulaChangedCount = rebindDynamicFormulaDependents(cellIndex, formulaChangedCount)
      topologyChanged = topologyChanged || formulaChangedCount !== reboundCount
    }
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
          const { sheetId, mutation } = ref
          const order = args.state.trackReplicaVersions && batch ? batchOpOrder(batch, refIndex) : undefined
          const existingIndex = resolveOperationExistingMutationCellIndex(args.state.workbook, ref)

          switch (mutation.kind) {
            case 'setCellValue': {
              const sheetName = sheetNameResolver.resolve(sheetId)
              const { hasExactLookupDependents, hasSortedLookupDependents, hasAggregateDependents } = trackedColumnDependencyFlags.resolve(
                sheetId,
                mutation.col,
              )
              if (mutation.value === null && !isRestore && (existingIndex === undefined || isNullLiteralWriteNoOp(existingIndex))) {
                break
              }
              const canFastOverwriteExisting = existingIndex !== undefined && canFastPathLiteralOverwrite(existingIndex)
              const needsDirectLookupNumericValue = canFastOverwriteExisting
              const oldExactLookupNumber =
                canFastOverwriteExisting && hasExactLookupDependents ? readExactNumericValueForLookup(existingIndex) : undefined
              const newExactLookupNumber =
                hasExactLookupDependents || needsDirectLookupNumericValue ? exactLookupLiteralNumericValue(mutation.value) : undefined
              const oldApproximateLookupNumber =
                canFastOverwriteExisting && hasSortedLookupDependents ? readApproximateNumericValueForLookup(existingIndex) : undefined
              const newApproximateLookupNumber =
                hasSortedLookupDependents || needsDirectLookupNumericValue ? directScalarLiteralNumericValue(mutation.value) : undefined
              const exactLookupDependentsHandled =
                !isRestore &&
                hasExactLookupDependents &&
                !hasAggregateDependents &&
                oldExactLookupNumber !== undefined &&
                newExactLookupNumber !== undefined &&
                canSkipExactLookupNumericColumnWrite(sheetId, mutation.col, mutation.row, oldExactLookupNumber, newExactLookupNumber)
              const sortedLookupDependentsHandled =
                !isRestore &&
                hasSortedLookupDependents &&
                oldApproximateLookupNumber !== undefined &&
                newApproximateLookupNumber !== undefined &&
                canSkipApproximateLookupNumericColumnWrite(
                  sheetId,
                  sheetName,
                  mutation.col,
                  mutation.row,
                  oldApproximateLookupNumber,
                  newApproximateLookupNumber,
                )
              const needsLookupValueRead =
                hasAggregateDependents ||
                (hasExactLookupDependents && !exactLookupDependentsHandled) ||
                (hasSortedLookupDependents && !sortedLookupDependentsHandled)
              const needsLookupOwnerInvalidation =
                (hasExactLookupDependents && exactLookupDependentsHandled) || (hasSortedLookupDependents && sortedLookupDependentsHandled)
              let directDependentsHandled = false
              if (!isRestore && canFastOverwriteExisting) {
                const oldNumber = directScalarCellNumericValue(existingIndex)
                const newNumber = directScalarLiteralNumericValue(mutation.value)
                if (oldNumber !== undefined && newNumber !== undefined) {
                  directDependentsHandled = markPostRecalcDirectScalarNumericDependents(
                    existingIndex,
                    oldNumber,
                    newNumber,
                    postRecalcDirectFormulaIndices,
                    newExactLookupNumber,
                    newApproximateLookupNumber,
                  )
                }
              }
              const canUseDirectLookupCurrent =
                !isRestore &&
                canFastOverwriteExisting &&
                (newExactLookupNumber !== undefined || newApproximateLookupNumber !== undefined) &&
                !needsLookupValueRead &&
                !directDependentsHandled
              if (canUseDirectLookupCurrent) {
                directDependentsHandled = markPostRecalcDirectLookupCurrentDependentsFromNumeric(
                  existingIndex,
                  newExactLookupNumber,
                  newApproximateLookupNumber,
                  postRecalcDirectFormulaIndices,
                )
              }
              let prior = needsLookupValueRead || !directDependentsHandled ? readCellValueForLookup(existingIndex) : undefined
              if (canFastOverwriteExisting) {
                writeLiteralToCellStore(args.state.workbook.cellStore, existingIndex, mutation.value, args.state.strings)
                args.state.workbook.notifyCellValueWritten(existingIndex)
                if (!isRestore) {
                  rebindValueSensitiveFormulaDependents(existingIndex)
                }
                if (needsLookupOwnerInvalidation) {
                  queueHandledLookupInvalidation(
                    sheetId,
                    sheetName,
                    mutation.col,
                    hasExactLookupDependents && exactLookupDependentsHandled,
                    hasSortedLookupDependents && sortedLookupDependentsHandled,
                  )
                  if (!needsLookupValueRead) {
                    lookupHandledInputCellIndices.push(existingIndex)
                  }
                }
                const newValue =
                  needsLookupValueRead || !directDependentsHandled ? literalToValue(mutation.value, args.state.strings) : undefined
                if (!isRestore && !directDependentsHandled && newValue) {
                  prior ??= readCellValueForLookup(existingIndex)
                  const genericDirectDependentsHandled = markPostRecalcDirectFormulaDependents(
                    existingIndex,
                    postRecalcDirectFormulaIndices,
                    prior.value,
                    newValue,
                  )
                  if (!genericDirectDependentsHandled) {
                    markDirectScalarDeltaClosure(existingIndex, prior.value, newValue, postRecalcDirectFormulaIndices)
                  }
                }
                if (needsLookupValueRead) {
                  const newStringId =
                    typeof mutation.value === 'string' ? args.state.workbook.cellStore.stringIds[existingIndex] : undefined
                  const priorLookup = prior ?? readCellValueForLookup(existingIndex)
                  const newLookupValue = newValue ?? literalToValue(mutation.value, args.state.strings)
                  if (hasExactLookupDependents || hasAggregateDependents) {
                    const exactLookupRequest = withOptionalLookupStringIds({
                      sheetName,
                      row: mutation.row,
                      col: mutation.col,
                      oldValue: priorLookup.value,
                      newValue: newLookupValue,
                      oldStringId: priorLookup.stringId,
                      newStringId,
                      inputCellIndex: existingIndex,
                    })
                    if (hasExactLookupDependents && !exactLookupDependentsHandled) {
                      formulaChangedCount = noteExactLookupLiteralWriteWhenDirty(
                        exactLookupRequest,
                        formulaChangedCount,
                        exactLookupImpactCaches,
                      )
                    }
                    if (hasAggregateDependents) {
                      args.noteAggregateLiteralWrite({
                        sheetName: exactLookupRequest.sheetName,
                        row: exactLookupRequest.row,
                        col: exactLookupRequest.col,
                        oldValue: exactLookupRequest.oldValue,
                        newValue: exactLookupRequest.newValue,
                      })
                      formulaChangedCount = markAffectedDirectRangeDependents(
                        exactLookupRequest,
                        formulaChangedCount,
                        postRecalcDirectFormulaIndices,
                      )
                    }
                  }
                  if (hasSortedLookupDependents) {
                    const sortedLookupRequest = withOptionalLookupStringIds({
                      sheetName,
                      row: mutation.row,
                      col: mutation.col,
                      oldValue: priorLookup.value,
                      newValue: newLookupValue,
                      oldStringId: priorLookup.stringId,
                      newStringId,
                    })
                    if (!sortedLookupDependentsHandled) {
                      formulaChangedCount = noteSortedLookupLiteralWriteWhenDirty(sortedLookupRequest, formulaChangedCount)
                    }
                  }
                }
                changedInputCount = args.markInputChanged(existingIndex, changedInputCount)
                if (trackExplicitChanges) {
                  explicitChangedCount = args.markExplicitChanged(existingIndex, explicitChangedCount)
                }
                if (!isRestore && args.state.trackReplicaVersions) {
                  setCellEntityVersion(sheetName, formatAddress(mutation.row, mutation.col), order!)
                }
                break
              }
              if (existingIndex !== undefined) {
                changedInputCount = args.markPivotRootsChanged(args.clearPivotForCell(existingIndex), changedInputCount)
              }
              const cellIndex = args.state.workbook.ensureCellAt(sheetId, mutation.row, mutation.col).cellIndex
              if (!isRestore && existingIndex !== undefined) {
                changedInputCount = args.markSpillRootsChanged(args.clearOwnedSpill(cellIndex), changedInputCount)
                const removedFormula = args.removeFormula(cellIndex)
                topologyChanged = removedFormula || topologyChanged
                if (removedFormula) {
                  args.invalidateAggregateColumn({ sheetName, col: mutation.col })
                }
                if (removedFormula) {
                  clearTrackedColumnDependencyFlagCache()
                }
              }
              writeLiteralToCellStore(args.state.workbook.cellStore, cellIndex, mutation.value, args.state.strings)
              args.state.workbook.notifyCellValueWritten(cellIndex)
              if (!isRestore) {
                rebindValueSensitiveFormulaDependents(cellIndex)
              }
              const newValue =
                needsLookupValueRead || !directDependentsHandled ? literalToValue(mutation.value, args.state.strings) : undefined
              if (!isRestore && !directDependentsHandled && newValue) {
                prior ??= readCellValueForLookup(existingIndex)
                const genericDirectDependentsHandled = markPostRecalcDirectFormulaDependents(
                  cellIndex,
                  postRecalcDirectFormulaIndices,
                  prior.value,
                  newValue,
                )
                if (!genericDirectDependentsHandled) {
                  markDirectScalarDeltaClosure(cellIndex, prior.value, newValue, postRecalcDirectFormulaIndices)
                }
              }
              if (needsLookupValueRead) {
                const newStringId = typeof mutation.value === 'string' ? args.state.workbook.cellStore.stringIds[cellIndex] : undefined
                const priorLookup = prior ?? readCellValueForLookup(existingIndex)
                const newLookupValue = newValue ?? literalToValue(mutation.value, args.state.strings)
                if (hasExactLookupDependents || hasAggregateDependents) {
                  const exactLookupRequest = withOptionalLookupStringIds({
                    sheetName,
                    row: mutation.row,
                    col: mutation.col,
                    oldValue: priorLookup.value,
                    newValue: newLookupValue,
                    oldStringId: priorLookup.stringId,
                    newStringId,
                    inputCellIndex: cellIndex,
                  })
                  if (hasExactLookupDependents && !exactLookupDependentsHandled) {
                    formulaChangedCount = noteExactLookupLiteralWriteWhenDirty(
                      exactLookupRequest,
                      formulaChangedCount,
                      exactLookupImpactCaches,
                    )
                  }
                  if (hasAggregateDependents) {
                    args.noteAggregateLiteralWrite({
                      sheetName: exactLookupRequest.sheetName,
                      row: exactLookupRequest.row,
                      col: exactLookupRequest.col,
                      oldValue: exactLookupRequest.oldValue,
                      newValue: exactLookupRequest.newValue,
                    })
                    formulaChangedCount = markAffectedDirectRangeDependents(
                      exactLookupRequest,
                      formulaChangedCount,
                      postRecalcDirectFormulaIndices,
                    )
                  }
                }
                if (hasSortedLookupDependents) {
                  const sortedLookupRequest = withOptionalLookupStringIds({
                    sheetName,
                    row: mutation.row,
                    col: mutation.col,
                    oldValue: priorLookup.value,
                    newValue: newLookupValue,
                    oldStringId: priorLookup.stringId,
                    newStringId,
                  })
                  if (!sortedLookupDependentsHandled) {
                    formulaChangedCount = noteSortedLookupLiteralWriteWhenDirty(sortedLookupRequest, formulaChangedCount)
                  }
                }
              }
              args.state.workbook.cellStore.flags[cellIndex] =
                (args.state.workbook.cellStore.flags[cellIndex] ?? 0) &
                ~(CellFlags.HasFormula | CellFlags.JsOnly | CellFlags.InCycle | CellFlags.SpillChild | CellFlags.PivotOutput)
              if (!isRestore && mutation.value === null) {
                pruneCellIfOrphaned(cellIndex)
              }
              changedInputCount = args.markInputChanged(cellIndex, changedInputCount)
              if (trackExplicitChanges) {
                explicitChangedCount = args.markExplicitChanged(cellIndex, explicitChangedCount)
              }
              if (!isRestore && args.state.trackReplicaVersions) {
                setCellEntityVersion(sheetName, formatAddress(mutation.row, mutation.col), order!)
              }
              break
            }
            case 'setCellFormula': {
              const sheetName = sheetNameResolver.resolve(sheetId)
              const { hasExactLookupDependents, hasSortedLookupDependents, hasAggregateDependents } = trackedColumnDependencyFlags.resolve(
                sheetId,
                mutation.col,
              )
              if (hasExactLookupDependents) {
                args.invalidateExactLookupColumn({ sheetName, col: mutation.col })
              }
              if (hasSortedLookupDependents) {
                args.invalidateSortedLookupColumn({ sheetName, col: mutation.col })
              }
              if (!isRestore && existingIndex !== undefined) {
                changedInputCount = args.markPivotRootsChanged(args.clearPivotForCell(existingIndex), changedInputCount)
              }
              const cellIndex = args.state.workbook.ensureCellAt(sheetId, mutation.row, mutation.col).cellIndex
              if (!isRestore && existingIndex !== undefined) {
                changedInputCount = args.markSpillRootsChanged(args.clearOwnedSpill(cellIndex), changedInputCount)
              }
              const priorHadFormula = args.state.formulas.get(cellIndex) !== undefined
              const oldFormulaNumber = !isRestore && priorHadFormula ? readExactNumericValueForLookup(cellIndex) : undefined
              const compileStarted = isRestore ? 0 : performance.now()
              try {
                const priorDirectScalarFormula = args.state.formulas.get(cellIndex)?.directScalar !== undefined
                const canRewriteFormulaPreservingBinding =
                  !isRestore &&
                  priorDirectScalarFormula &&
                  !hasExactLookupDependents &&
                  !hasSortedLookupDependents &&
                  !hasAggregateDependents &&
                  args.rewriteFormulaSourcePreservingBinding !== undefined
                const canAssumeFreshFormula =
                  !isRestore &&
                  existingIndex === undefined &&
                  !priorHadFormula &&
                  args.state.workbook.metadata.definedNames.size === 0 &&
                  args.bindPreparedFormula !== undefined &&
                  args.compileTemplateFormula !== undefined
                const changedTopology = canRewriteFormulaPreservingBinding
                  ? args.rewriteFormulaSourcePreservingBinding(cellIndex, sheetName, mutation.formula)
                    ? false
                    : args.bindFormula(cellIndex, sheetName, mutation.formula)
                  : canAssumeFreshFormula
                    ? bindFreshTemplateFormula(args, cellIndex, sheetName, mutation)
                    : args.bindFormula(cellIndex, sheetName, mutation.formula)
                const runtimeFormula = args.state.formulas.get(cellIndex)
                if (hasAggregateDependents) {
                  args.invalidateAggregateColumn({ sheetName, col: mutation.col })
                }
                if (!isRestore) {
                  compileMs += performance.now() - compileStarted
                }
                clearTrackedColumnDependencyFlagCache()
                changedInputCount = args.markInputChanged(cellIndex, changedInputCount)
                const freshDirectAggregateAnalysis = analyzeFreshDirectAggregateFormula(args, {
                  priorHadFormula,
                  formulaCellIndex: cellIndex,
                  formula: runtimeFormula,
                })
                const canSkipTopoRepair = freshDirectAggregateAnalysis.canSkipTopoRepair
                const freshDirectFormulaResult = freshDirectAggregateAnalysis.currentResult
                const evaluatedFreshDirectFormula =
                  freshDirectFormulaResult !== undefined
                    ? applyDirectFormulaCurrentResult(cellIndex, freshDirectFormulaResult)
                    : canSkipTopoRepair && args.evaluateDirectFormula(cellIndex) !== undefined
                const handledFormulaReplacementAsDirectDelta =
                  priorHadFormula &&
                  !hasExactLookupDependents &&
                  !hasSortedLookupDependents &&
                  !hasAggregateDependents &&
                  tryApplyFormulaReplacementAsDirectScalarDeltaRoot({
                    cellIndex,
                    oldNumber: oldFormulaNumber,
                    changedTopology,
                    postRecalcDirectFormulaIndices,
                    postRecalcDirectFormulaMetrics,
                  })
                if (!handledFormulaReplacementAsDirectDelta) {
                  if (!evaluatedFreshDirectFormula) {
                    formulaChangedCount = args.markFormulaChanged(cellIndex, formulaChangedCount)
                  }
                }
                topologyChanged = topologyChanged || (changedTopology && !canSkipTopoRepair)
                const aggregateDependents = hasAggregateDependents
                  ? collectAffectedDirectRangeDependents({
                      sheetName,
                      row: mutation.row,
                      col: mutation.col,
                    }).filter((candidate) => candidate !== cellIndex)
                  : []
                if (aggregateDependents.length > 0) {
                  formulaChangedCount = args.rebindFormulaCells(aggregateDependents, formulaChangedCount)
                  for (let index = 0; index < aggregateDependents.length; index += 1) {
                    postRecalcDirectFormulaIndices.add(aggregateDependents[index]!)
                    formulaChangedCount = args.markFormulaChanged(aggregateDependents[index]!, formulaChangedCount)
                    changedInputCount = args.markInputChanged(aggregateDependents[index]!, changedInputCount)
                  }
                  topologyChanged = true
                }
              } catch {
                if (!isRestore) {
                  compileMs += performance.now() - compileStarted
                }
                const removedFormula = args.removeFormula(cellIndex)
                topologyChanged = removedFormula || topologyChanged
                clearTrackedColumnDependencyFlagCache()
                args.setInvalidFormulaValue(cellIndex)
                changedInputCount = args.markInputChanged(cellIndex, changedInputCount)
              }
              if (trackExplicitChanges) {
                explicitChangedCount = args.markExplicitChanged(cellIndex, explicitChangedCount)
              }
              if (!isRestore && args.state.trackReplicaVersions) {
                setCellEntityVersion(sheetName, formatAddress(mutation.row, mutation.col), order!)
              }
              break
            }
            case 'clearCell': {
              const { hasExactLookupDependents, hasSortedLookupDependents, hasAggregateDependents, needsLookupValueRead } =
                trackedColumnDependencyFlags.resolve(sheetId, mutation.col)
              const prior = readCellValueForLookup(existingIndex)
              if (existingIndex !== undefined && isClearCellNoOp(existingIndex)) {
                break
              }
              if (existingIndex !== undefined && canFastPathLiteralOverwrite(existingIndex)) {
                args.state.workbook.cellStore.setValue(existingIndex, emptyValue())
                args.state.workbook.notifyCellValueWritten(existingIndex)
                if (!isRestore) {
                  rebindValueSensitiveFormulaDependents(existingIndex)
                }
                if (!isRestore) {
                  const nextValue = emptyValue()
                  const directDependentsHandled = markPostRecalcDirectFormulaDependents(
                    existingIndex,
                    postRecalcDirectFormulaIndices,
                    prior.value,
                    nextValue,
                  )
                  if (!directDependentsHandled) {
                    markDirectScalarDeltaClosure(existingIndex, prior.value, nextValue, postRecalcDirectFormulaIndices)
                  }
                }
                if (needsLookupValueRead) {
                  const sheetName = sheetNameResolver.resolve(sheetId)
                  if (hasExactLookupDependents || hasAggregateDependents) {
                    const exactLookupRequest = withOptionalLookupStringIds({
                      sheetName,
                      row: mutation.row,
                      col: mutation.col,
                      oldValue: prior.value,
                      newValue: emptyValue(),
                      oldStringId: prior.stringId,
                      newStringId: undefined,
                      inputCellIndex: existingIndex,
                    })
                    if (hasExactLookupDependents) {
                      formulaChangedCount = noteExactLookupLiteralWriteWhenDirty(
                        exactLookupRequest,
                        formulaChangedCount,
                        exactLookupImpactCaches,
                      )
                    }
                    if (hasAggregateDependents) {
                      args.noteAggregateLiteralWrite({
                        sheetName: exactLookupRequest.sheetName,
                        row: exactLookupRequest.row,
                        col: exactLookupRequest.col,
                        oldValue: exactLookupRequest.oldValue,
                        newValue: exactLookupRequest.newValue,
                      })
                      formulaChangedCount = markAffectedDirectRangeDependents(
                        exactLookupRequest,
                        formulaChangedCount,
                        postRecalcDirectFormulaIndices,
                      )
                    }
                  }
                  if (hasSortedLookupDependents) {
                    const sortedLookupRequest = withOptionalLookupStringIds({
                      sheetName,
                      row: mutation.row,
                      col: mutation.col,
                      oldValue: prior.value,
                      newValue: emptyValue(),
                      oldStringId: prior.stringId,
                      newStringId: undefined,
                    })
                    formulaChangedCount = noteSortedLookupLiteralWriteWhenDirty(sortedLookupRequest, formulaChangedCount)
                  }
                }
                changedInputCount = args.markInputChanged(existingIndex, changedInputCount)
                if (trackExplicitChanges) {
                  explicitChangedCount = args.markExplicitChanged(existingIndex, explicitChangedCount)
                }
                if (!isRestore && args.state.trackReplicaVersions) {
                  setCellEntityVersion(sheetNameResolver.resolve(sheetId), formatAddress(mutation.row, mutation.col), order!)
                }
                break
              }
              if (existingIndex === undefined) {
                if (!isRestore && args.state.trackReplicaVersions) {
                  setCellEntityVersion(sheetNameResolver.resolve(sheetId), formatAddress(mutation.row, mutation.col), order!)
                }
                break
              }
              changedInputCount = args.markPivotRootsChanged(args.clearPivotForCell(existingIndex), changedInputCount)
              changedInputCount = args.markSpillRootsChanged(args.clearOwnedSpill(existingIndex), changedInputCount)
              const removedFormula = args.removeFormula(existingIndex)
              topologyChanged = removedFormula || topologyChanged
              if (removedFormula) {
                args.invalidateAggregateColumn({ sheetName: sheetNameResolver.resolve(sheetId), col: mutation.col })
              }
              if (removedFormula) {
                clearTrackedColumnDependencyFlagCache()
              }
              args.state.workbook.cellStore.setValue(existingIndex, emptyValue())
              args.state.workbook.notifyCellValueWritten(existingIndex)
              if (!isRestore) {
                rebindValueSensitiveFormulaDependents(existingIndex)
              }
              if (!isRestore) {
                const nextValue = emptyValue()
                const directDependentsHandled = markPostRecalcDirectFormulaDependents(
                  existingIndex,
                  postRecalcDirectFormulaIndices,
                  prior.value,
                  nextValue,
                )
                if (!directDependentsHandled) {
                  markDirectScalarDeltaClosure(existingIndex, prior.value, nextValue, postRecalcDirectFormulaIndices)
                }
              }
              if (needsLookupValueRead) {
                const sheetName = sheetNameResolver.resolve(sheetId)
                if (hasExactLookupDependents || hasAggregateDependents) {
                  const exactLookupRequest = withOptionalLookupStringIds({
                    sheetName,
                    row: mutation.row,
                    col: mutation.col,
                    oldValue: prior.value,
                    newValue: emptyValue(),
                    oldStringId: prior.stringId,
                    newStringId: undefined,
                    inputCellIndex: existingIndex,
                  })
                  if (hasExactLookupDependents) {
                    formulaChangedCount = noteExactLookupLiteralWriteWhenDirty(
                      exactLookupRequest,
                      formulaChangedCount,
                      exactLookupImpactCaches,
                    )
                  }
                  if (hasAggregateDependents) {
                    args.noteAggregateLiteralWrite({
                      sheetName: exactLookupRequest.sheetName,
                      row: exactLookupRequest.row,
                      col: exactLookupRequest.col,
                      oldValue: exactLookupRequest.oldValue,
                      newValue: exactLookupRequest.newValue,
                    })
                    formulaChangedCount = markAffectedDirectRangeDependents(
                      exactLookupRequest,
                      formulaChangedCount,
                      postRecalcDirectFormulaIndices,
                    )
                  }
                }
                if (hasSortedLookupDependents) {
                  const sortedLookupRequest = withOptionalLookupStringIds({
                    sheetName,
                    row: mutation.row,
                    col: mutation.col,
                    oldValue: prior.value,
                    newValue: emptyValue(),
                    oldStringId: prior.stringId,
                    newStringId: undefined,
                  })
                  formulaChangedCount = noteSortedLookupLiteralWriteWhenDirty(sortedLookupRequest, formulaChangedCount)
                }
              }
              args.state.workbook.cellStore.flags[existingIndex] =
                (args.state.workbook.cellStore.flags[existingIndex] ?? 0) &
                ~(
                  CellFlags.AuthoredBlank |
                  CellFlags.HasFormula |
                  CellFlags.JsOnly |
                  CellFlags.InCycle |
                  CellFlags.SpillChild |
                  CellFlags.PivotOutput
                )
              normalizeHistoryDependencyPlaceholder(existingIndex, source)
              if (!isRestore) {
                pruneCellIfOrphaned(existingIndex)
              }
              changedInputCount = args.markInputChanged(existingIndex, changedInputCount)
              if (trackExplicitChanges) {
                explicitChangedCount = args.markExplicitChanged(existingIndex, explicitChangedCount)
              }
              if (!isRestore && args.state.trackReplicaVersions) {
                setCellEntityVersion(sheetNameResolver.resolve(sheetId), formatAddress(mutation.row, mutation.col), order!)
              }
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
      hadCycleMembersBeforeNow,
      markCycleMemberInputsChanged,
      canSkipDirtyTraversalForChangedInputs,
      changedInputsNeedRegionQueryIndices,
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
