import { Effect } from 'effect'
import type { LiteralInput } from '@bilig/protocol'
import type { EngineExistingNumericCellMutationResult } from '../../cell-mutations-at.js'
import type { SheetRecord } from '../../workbook-store.js'
import type { U32 } from '../runtime-state.js'
import { EngineMutationError } from '../errors.js'
import {
  hasOperationCycleMembers,
  markOperationCycleMemberInputsChanged,
  normalizeOperationHistoryDependencyPlaceholder,
  collectOperationDynamicFormulaDependents,
  pruneOperationCellIfOrphaned,
  refreshDependentRangesAndRebindOperationFormulaDependents,
  rebindOperationDynamicFormulaDependents,
} from './operation-cell-lifecycle-helpers.js'
import type { DirectFormulaIndexCollection } from './direct-formula-index-collection.js'
import {
  aggregateColumnDependencyKey,
  canEvaluatePostRecalcDirectFormulasWithoutKernel,
  collectTrackedDependents,
  composeSingleDisjointExplicitEventChanges,
  countDirectFormulaDeltaSkip,
  directAggregateNumericContribution,
  directCriteriaTouchesPoint,
  directFormulaChangesAreDisjointFromInputs,
  hasCompleteDirectFormulaDeltas,
  lookupImpactCacheKey,
} from './direct-formula-recalc-helpers.js'
import {
  cellRange,
  makeCompactExistingNumericMutationResult,
  makeExistingNumericMutationResult,
  mergeChangedCellIndices,
  mutationErrorMessage,
  rangesIntersect,
  tagTrustedPhysicalTrackedChanges,
  throwProtectionBlocked,
} from './operation-change-helpers.js'
import { createOperationReplicaVersionWriter } from './operation-replica-helpers.js'
import { createOperationDirectFormulaDeltas } from './operation-direct-formula-deltas.js'
import { createOperationDirectFormulaValues } from './operation-direct-formula-values.js'
import { createOperationDirectRangeDependentService } from './operation-direct-range-dependents.js'
import { createOperationDirectLookupCurrentService } from './operation-direct-lookup-current.js'
import { createOperationColumnDependencyTrackerService } from './operation-column-dependency-tracker.js'
import { createOperationLookupDirtyMarkerService } from './operation-lookup-dirty-markers.js'
import { createOperationDirectPostRecalcMarkers } from './operation-direct-post-recalc-markers.js'
import { createOperationDerivedOpApplier } from './operation-derived-op-helpers.js'
import {
  canSkipOperationDirtyTraversalForChangedInputs,
  operationChangedInputsNeedRegionQueryIndices,
} from './operation-dirty-traversal-helpers.js'
import { countOperationPostRecalcDirectFormulaMetric, type DirectFormulaMetricCounts } from './operation-post-recalc-direct-formulas.js'
import {
  tryApplySingleDirectAggregateLiteralMutationFastPath as tryApplySingleDirectAggregateLiteralMutationFastPathWithArgs,
  tryApplyTrustedSingleRangeDirectAggregateExistingNumericMutation as tryApplyTrustedSingleRangeDirectAggregateExistingNumericMutationWithArgs,
} from './operation-direct-aggregate-literal-fast-path.js'
import { tryApplyTrustedDirectScalarClosureExistingNumericMutation as tryApplyTrustedDirectScalarClosureExistingNumericMutationWithArgs } from './operation-direct-scalar-closure-fast-path.js'
import {
  tryApplyFormulaLeafExistingLiteralMutation as tryApplyFormulaLeafExistingLiteralMutationWithArgs,
  tryApplyTrustedFormulaLeafExistingNumericMutation as tryApplyTrustedFormulaLeafExistingNumericMutationWithArgs,
} from './operation-formula-leaf-existing-numeric-fast-path.js'
import { createOperationDirectScalarBatchFastPaths } from './operation-direct-scalar-batch-fast-paths.js'
import { tryApplySingleDirectScalarLiteralMutationWithoutEvents as tryApplySingleDirectScalarLiteralMutationWithoutEventsWithArgs } from './operation-direct-scalar-literal-fast-path.js'
import { tryApplySingleDirectFormulaLiteralMutationWithoutEvents as tryApplySingleDirectFormulaLiteralMutationWithoutEventsWithArgs } from './operation-direct-formula-literal-fast-path.js'
import { createOperationSingleExistingLiteralFastPath } from './operation-single-existing-literal-fast-path.js'
import { createOperationLookupAccess } from './operation-lookup-access.js'
import { createOperationLookupPlanner } from './operation-lookup-planner.js'
import { tryApplySingleDirectLookupOperandMutationFastPath as tryApplySingleDirectLookupOperandMutationFastPathWithArgs } from './operation-direct-lookup-operand-fast-path.js'
import { canSkipOperationFormulaColumnVersion } from './operation-literal-write-helpers.js'
import { createOperationServiceRuntimeHelpers } from './operation-service-runtime-helpers.js'
import { tryApplyOperationFormulaReplacementAsDirectScalarDeltaRoot } from './operation-formula-replacement-direct-scalar.js'
import { tryApplySingleKernelSyncOnlyLiteralMutationFastPath as tryApplySingleKernelSyncOnlyLiteralMutationFastPathWithArgs } from './operation-kernel-sync-literal-fast-path.js'
import { createOperationBatchApplier } from './operation-batch-applier.js'
import { createOperationCellMutationApplier } from './operation-cell-mutation-applier.js'
import {
  ENGINE_OPERATION_TEST_HOOKS_ENABLED,
  type CreateEngineOperationServiceArgs,
  type EngineOperationService,
  type MutationSource,
} from './operation-service-types.js'

export type { EngineOperationService } from './operation-service-types.js'

const DIRECT_RANGE_POST_RECALC_LIMIT = 16_384
const DIRECT_SCALAR_DELTA_CLOSURE_LIMIT = 4_096

export const operationServiceTestHooks = {
  aggregateColumnDependencyKey,
  canEvaluatePostRecalcDirectFormulasWithoutKernel,
  cellRange,
  collectTrackedDependents,
  composeSingleDisjointExplicitEventChanges,
  countDirectFormulaDeltaSkip,
  directAggregateNumericContribution,
  directCriteriaTouchesPoint,
  directFormulaChangesAreDisjointFromInputs,
  getConstantDirectFormulaDeltas: hasCompleteDirectFormulaDeltas,
  lookupImpactCacheKey,
  makeCompactExistingNumericMutationResult,
  makeExistingNumericMutationResult,
  mergeChangedCellIndices,
  rangesIntersect,
  tagTrustedPhysicalTrackedChanges,
  throwProtectionBlocked,
}

export function createEngineOperationService(args: CreateEngineOperationServiceArgs): EngineOperationService {
  const {
    emitBatch,
    deferSingleCellKernelSync,
    makeSingleLiteralSkipMetrics,
    writeNumericLiteralToExistingCell,
    writeTrustedExistingNumericLiteralToCell,
    writeFastPathLiteralToExistingCell,
    cellsShareVersionColumn,
    withOptionalColumnVersionBatch,
    canFastPathLiteralOverwrite,
    isNullLiteralWriteNoOp,
    isClearCellNoOp,
  } = createOperationServiceRuntimeHelpers({
    state: args.state,
    deferKernelSync: args.deferKernelSync,
  })

  const replicaVersionWriter = createOperationReplicaVersionWriter({
    trackReplicaVersions: args.state.trackReplicaVersions,
    entityVersions: args.state.entityVersions,
    sheetDeleteVersions: args.state.sheetDeleteVersions,
  })
  const { setCellEntityVersion, setEntityVersionForOp } = replicaVersionWriter

  const {
    readCellValueForLookup,
    readApproximateNumericValueForLookup,
    readExactNumericValueForLookup,
    readCellValueAtForLookup,
    readDirectCriteriaOperandValue,
    directCriteriaMatchesChangedAggregateRow,
    tryDirectCriteriaSumDelta,
    readApproximateNumericValueAtForLookup,
    isLocallySortedNumericWrite,
    isLocallySortedTextWrite,
  } = createOperationLookupAccess({
    workbook: args.state.workbook,
    strings: args.state.strings,
  })

  const {
    planSingleExactLookupNumericColumnWrite,
    planExactLookupNumericColumnWrite,
    canSkipExactLookupNumericColumnWrite,
    planSingleApproximateLookupNumericColumnWrite,
    planApproximateLookupNumericColumnWrite,
    canSkipApproximateLookupNumericColumnWrite,
    canSkipApproximateLookupNewNumericColumnWrite,
    canPatchUniformLookupTailWrite,
    patchUniformLookupTailWrites,
    canSkipApproximateLookupDirtyMark,
  } = createOperationLookupPlanner({
    state: args.state,
    access: {
      readExactNumericValueForLookup,
      readApproximateNumericValueForLookup,
      readCellValueForLookup,
      isLocallySortedNumericWrite,
      isLocallySortedTextWrite,
    },
    getSingleEntityDependent: args.getSingleEntityDependent,
    getEntityDependents: args.getEntityDependents,
  })

  const { markAffectedApproximateLookupDependents, noteExactLookupLiteralWriteWhenDirty, noteSortedLookupLiteralWriteWhenDirty } =
    createOperationLookupDirtyMarkerService({
      state: args.state,
      getEntityDependents: args.getEntityDependents,
      getSingleEntityDependent: args.getSingleEntityDependent,
      markFormulaChanged: args.markFormulaChanged,
      readCellValueForLookup,
      canSkipApproximateLookupDirtyMark,
      noteExactLookupLiteralWrite: args.noteExactLookupLiteralWrite,
      noteSortedLookupLiteralWrite: args.noteSortedLookupLiteralWrite,
      lookupImpactCacheKey,
    })

  const pruneCellIfOrphaned = (cellIndex: number): void => {
    pruneOperationCellIfOrphaned({
      workbook: args.state.workbook,
      cellIndex,
      collectFormulaDependents: args.collectFormulaDependents,
    })
  }

  const normalizeHistoryDependencyPlaceholder = (cellIndex: number, source: MutationSource): void => {
    normalizeOperationHistoryDependencyPlaceholder({
      state: args.state,
      source,
      cellIndex,
      collectFormulaDependents: args.collectFormulaDependents,
    })
  }

  const markCycleMemberInputsChanged = (changedInputCount: number): number => {
    return markOperationCycleMemberInputsChanged({
      formulas: args.state.formulas,
      cellStore: args.state.workbook.cellStore,
      changedInputCount,
      markInputChanged: args.markInputChanged,
    })
  }

  const hasCycleMembersNow = (): boolean => {
    return hasOperationCycleMembers({
      counters: args.state.counters,
      formulas: args.state.formulas,
      cellStore: args.state.workbook.cellStore,
    })
  }

  const rebindDynamicFormulaDependents = (cellIndex: number, formulaChangedCount: number): number =>
    rebindOperationDynamicFormulaDependents({
      cellIndex,
      formulaChangedCount,
      formulas: args.state.formulas,
      collectFormulaDependents: args.collectFormulaDependents,
      rebindFormulaCells: args.rebindFormulaCells,
    })

  const hasDynamicFormulaDependents = (cellIndex: number): boolean =>
    collectOperationDynamicFormulaDependents({
      cellIndex,
      formulas: args.state.formulas,
      collectFormulaDependents: args.collectFormulaDependents,
    }).length > 0

  const {
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
  } = createOperationDirectFormulaValues({
    state: args.state,
  })

  const {
    tryDirectUniformLookupCurrentResult,
    tryDirectUniformLookupCurrentResultFromNumeric,
    tryDirectUniformLookupNumericResultFromDescriptor,
    canEvaluateDirectUniformLookupCurrentResultFromNumeric,
    tryDirectApproximateLookupCurrentResultFromNumeric,
    tryDirectExactLookupCurrentResult,
  } = createOperationDirectLookupCurrentService({
    state: args.state,
    exactLookup: args.exactLookup,
    sortedLookup: args.sortedLookup,
  })

  const {
    hasTrackedExactLookupDependents,
    hasTrackedSortedLookupDependents,
    hasTrackedDirectRangeDependents,
    hasTrackedColumnDependents,
    hasNoCellDependents,
  } = createOperationColumnDependencyTrackerService({
    reverseState: args.reverseState,
    workbook: args.state.workbook,
    hasRegionFormulaSubscriptionsForColumn: args.hasRegionFormulaSubscriptionsForColumn,
    hasRegionFormulaSubscriptionsForColumnAt: args.hasRegionFormulaSubscriptionsForColumnAt,
  })

  const {
    collectAffectedDirectRangeDependents,
    collectSingleAffectedDirectRangeDependent,
    canApplyDirectAggregateLiteralDelta,
    canApplyDirectAggregateLiteralDeltaForRequest,
    collectSingleApplicableDirectAggregateDependent,
    markAffectedDirectRangeDependents,
  } = createOperationDirectRangeDependentService({
    workbook: args.state.workbook,
    formulas: args.state.formulas,
    reverseAggregateColumnEdges: args.reverseState.reverseAggregateColumnEdges,
    collectRegionFormulaDependentsForCell: args.collectRegionFormulaDependentsForCell,
    collectSingleRegionFormulaDependentForCell: args.collectSingleRegionFormulaDependentForCell,
    collectSingleRegionFormulaDependentForCellAt: args.collectSingleRegionFormulaDependentForCellAt,
    hasNoCellDependents,
    getSingleEntityDependent: args.getSingleEntityDependent,
    markFormulaChanged: args.markFormulaChanged,
    tryDirectCriteriaSumDelta,
    postRecalcLimit: DIRECT_RANGE_POST_RECALC_LIMIT,
  })

  const canSkipDirtyTraversalForChangedInputs = (
    changedInputCellIndices: U32,
    changedInputCount: number,
    postRecalcDirectFormulaIndices?: DirectFormulaIndexCollection,
    options: {
      readonly lookupHandledInputCellIndices?: readonly number[]
    } = {},
  ): boolean => {
    return canSkipOperationDirtyTraversalForChangedInputs({
      changedInputCellIndices,
      changedInputCount,
      postRecalcDirectFormulaIndices,
      options,
      access: {
        workbook: args.state.workbook,
        getSingleEntityDependent: args.getSingleEntityDependent,
        getEntityDependents: args.getEntityDependents,
        collectRegionFormulaDependentsForCell: args.collectRegionFormulaDependentsForCell,
        collectAffectedDirectRangeDependents,
        hasTrackedExactLookupDependents,
        hasTrackedSortedLookupDependents,
        hasTrackedDirectRangeDependents,
      },
    })
  }

  const changedInputsNeedRegionQueryIndices = (
    changedInputCellIndices: U32,
    changedInputCount: number,
    postRecalcDirectFormulaIndices?: DirectFormulaIndexCollection,
  ): boolean => {
    return operationChangedInputsNeedRegionQueryIndices({
      changedInputCellIndices,
      changedInputCount,
      postRecalcDirectFormulaIndices,
      access: {
        workbook: args.state.workbook,
        hasTrackedDirectRangeDependents,
      },
    })
  }

  const canSkipFormulaColumnVersion = (cellIndex: number): boolean => {
    return canSkipOperationFormulaColumnVersion({
      workbook: args.state.workbook,
      cellIndex,
      hasTrackedColumnDependents,
    })
  }

  const {
    canUseDirectFormulaPostRecalc,
    markDirectScalarDeltaClosure,
    markPostRecalcDirectFormulaDependents,
    markPostRecalcDirectLookupCurrentDependentsFromNumeric,
    markPostRecalcDirectScalarNumericDependents,
    tryDirectScalarNumericDeltaFromNumbers,
    tryMarkDirectScalarLinearDeltaClosure,
  } = createOperationDirectPostRecalcMarkers({
    state: args.state,
    getSingleEntityDependent: args.getSingleEntityDependent,
    getEntityDependents: args.getEntityDependents,
    hasNoCellDependents,
    canSkipDirectFormulaColumnVersion: canSkipFormulaColumnVersion,
    readDirectScalarCellNumber,
    directScalarCellNumericValue,
    directScalarCurrentResultMatchesCell,
    lookupCurrent: {
      canEvaluateDirectUniformLookupCurrentResultFromNumeric,
      tryDirectExactLookupCurrentResult,
      tryDirectUniformLookupCurrentResult,
      tryDirectUniformLookupCurrentResultFromNumeric,
    },
    scalarDeltaClosureLimit: DIRECT_SCALAR_DELTA_CLOSURE_LIMIT,
  })

  const {
    applyDirectFormulaNumericDelta,
    applyTerminalDirectFormulaNumericDeltaAndReturn,
    tryApplyDirectFormulaDeltas,
    tryApplyDirectScalarDeltas,
  } = createOperationDirectFormulaDeltas({
    state: args.state,
    canSkipTerminalFormulaColumnVersion: canSkipFormulaColumnVersion,
    canSkipDirectFormulaColumnVersion: canSkipFormulaColumnVersion,
  })

  const tryApplySingleDirectAggregateLiteralMutationFastPath = (request: {
    existingIndex: number
    sheetId?: number
    sheetName: string
    row: number
    col: number
    value: LiteralInput
    delta: number
    emitTracked: boolean
    singleRangeEntityDependent?: number
  }): EngineExistingNumericCellMutationResult | null =>
    tryApplySingleDirectAggregateLiteralMutationFastPathWithArgs(
      {
        state: args.state,
        directRangePostRecalcLimit: DIRECT_RANGE_POST_RECALC_LIMIT,
        getSingleEntityDependent: args.getSingleEntityDependent,
        collectAffectedDirectRangeDependents,
        collectSingleApplicableDirectAggregateDependent,
        canApplyDirectAggregateLiteralDeltaForRequest,
        canApplyDirectAggregateLiteralDelta,
        writeFastPathLiteralToExistingCell,
        writeTrustedExistingNumericLiteralToCell,
        applyTerminalDirectFormulaNumericDeltaAndReturn,
        applyDirectFormulaNumericDelta,
        cellsShareVersionColumn,
        withOptionalColumnVersionBatch,
        deferSingleCellKernelSync,
        makeSingleLiteralSkipMetrics,
      },
      request,
    )

  const tryApplyTrustedSingleRangeDirectAggregateExistingNumericMutation = (request: {
    existingIndex: number
    rangeEntityDependent: number
    sheet: SheetRecord
    sheetId: number
    col: number
    value: number
    delta: number
    hasExactLookupDependents: boolean
    hasSortedLookupDependents: boolean
  }): EngineExistingNumericCellMutationResult | null =>
    tryApplyTrustedSingleRangeDirectAggregateExistingNumericMutationWithArgs(
      {
        state: args.state,
        directRangePostRecalcLimit: DIRECT_RANGE_POST_RECALC_LIMIT,
        getSingleEntityDependent: args.getSingleEntityDependent,
        collectAffectedDirectRangeDependents,
        collectSingleApplicableDirectAggregateDependent,
        canApplyDirectAggregateLiteralDeltaForRequest,
        canApplyDirectAggregateLiteralDelta,
        writeFastPathLiteralToExistingCell,
        writeTrustedExistingNumericLiteralToCell,
        applyTerminalDirectFormulaNumericDeltaAndReturn,
        applyDirectFormulaNumericDelta,
        cellsShareVersionColumn,
        withOptionalColumnVersionBatch,
        deferSingleCellKernelSync,
        makeSingleLiteralSkipMetrics,
      },
      request,
    )

  const tryApplyTrustedDirectScalarClosureExistingNumericMutation = (request: {
    existingIndex: number
    sheet: SheetRecord
    sheetId: number
    col: number
    value: number
    oldNumber: number
    hasTrackedEventListeners: boolean
  }): EngineExistingNumericCellMutationResult | null =>
    tryApplyTrustedDirectScalarClosureExistingNumericMutationWithArgs(
      {
        state: args.state,
        tryMarkDirectScalarLinearDeltaClosure,
        writeTrustedExistingNumericLiteralToCell,
        tryApplyDirectScalarDeltas,
        deferSingleCellKernelSync,
        makeSingleLiteralSkipMetrics,
      },
      request,
    )

  const tryApplyTrustedFormulaLeafExistingNumericMutation = (request: {
    existingIndex: number
    formulaCellIndex: number
    sheet: SheetRecord
    col: number
    value: number
    hasTrackedEventListeners: boolean
  }): EngineExistingNumericCellMutationResult | null =>
    tryApplyTrustedFormulaLeafExistingNumericMutationWithArgs(
      {
        state: args.state,
        getSingleEntityDependent: args.getSingleEntityDependent,
        writeTrustedExistingNumericLiteralToCell,
        evaluateFormulaCell: args.evaluateFormulaCell,
        deferSingleCellKernelSync,
        makeSingleLiteralSkipMetrics,
      },
      request,
    )

  const tryApplyFormulaLeafExistingLiteralMutation = (request: {
    existingIndex: number
    formulaCellIndex: number
    value: LiteralInput
    hasTrackedEventListeners: boolean
  }): boolean =>
    tryApplyFormulaLeafExistingLiteralMutationWithArgs(
      {
        state: args.state,
        getSingleEntityDependent: args.getSingleEntityDependent,
        writeFastPathLiteralToExistingCell,
        evaluateFormulaCell: args.evaluateFormulaCell,
        deferSingleCellKernelSync,
        makeSingleLiteralSkipMetrics,
      },
      request,
    )

  const tryApplySingleDirectScalarLiteralMutationWithoutEvents = (request: {
    existingIndex: number
    value: LiteralInput
    oldNumber: number
    newNumber: number
  }): boolean =>
    tryApplySingleDirectScalarLiteralMutationWithoutEventsWithArgs(
      {
        state: args.state,
        getSingleEntityDependent: args.getSingleEntityDependent,
        getEntityDependents: args.getEntityDependents,
        canUseDirectFormulaPostRecalc,
        tryDirectScalarNumericDeltaFromNumbers,
        applyDirectFormulaNumericDelta,
        deferSingleCellKernelSync,
        makeSingleLiteralSkipMetrics,
      },
      request,
    )

  const tryApplySingleDirectLookupOperandMutationFastPath = (request: {
    existingIndex: number
    formulaCellIndex: number
    value: LiteralInput
    exactLookupValue: number | undefined
    approximateLookupValue: number | undefined
    emitTracked: boolean
    lookupSheetHint?: SheetRecord | undefined
    trustedInputSheet?: SheetRecord | undefined
    trustedInputCol?: number | undefined
  }): EngineExistingNumericCellMutationResult | null =>
    tryApplySingleDirectLookupOperandMutationFastPathWithArgs(
      {
        state: args.state,
        hasNoCellDependents,
        directScalarNumericResultMatchesCell,
        directScalarCurrentResultMatchesCell,
        tryDirectUniformLookupNumericResultFromDescriptor,
        tryDirectApproximateLookupCurrentResultFromNumeric,
        tryDirectUniformLookupCurrentResultFromNumeric,
        writeTrustedExistingNumericLiteralToCell,
        writeNumericLiteralToExistingCell,
        writeFastPathLiteralToExistingCell,
        applyTerminalDirectFormulaNumericResult,
        applyDirectFormulaCurrentResult,
        cellsShareVersionColumn,
        withOptionalColumnVersionBatch,
        deferSingleCellKernelSync,
        makeSingleLiteralSkipMetrics,
        evaluateDirectFormula: args.evaluateDirectFormula,
      },
      request,
    )

  const tryApplySingleDirectFormulaLiteralMutationWithoutEvents = (request: {
    existingIndex: number
    formulaCellIndex: number
    value: LiteralInput
    oldNumber: number
    newNumber: number
    exactLookupValue: number | undefined
    approximateLookupValue: number | undefined
  }): boolean =>
    tryApplySingleDirectFormulaLiteralMutationWithoutEventsWithArgs(
      {
        state: args.state,
        hasNoCellDependents,
        tryDirectUniformLookupNumericResultFromDescriptor,
        directScalarNumericResultMatchesCell,
        tryDirectUniformLookupCurrentResultFromNumeric,
        directScalarCurrentResultMatchesCell,
        canUseDirectFormulaPostRecalc,
        tryDirectScalarNumericDeltaFromNumbers,
        writeFastPathLiteralToExistingCell,
        applyTerminalDirectFormulaNumericResult,
        applyDirectFormulaCurrentResult,
        applyDirectFormulaNumericDelta,
        cellsShareVersionColumn,
        withOptionalColumnVersionBatch,
        deferSingleCellKernelSync,
        makeSingleLiteralSkipMetrics,
      },
      request,
    )

  const tryApplySingleKernelSyncOnlyLiteralMutationFastPath = (request: {
    existingIndex: number
    value: LiteralInput
    emitTracked: boolean
    afterWrite?: () => void
  }): boolean =>
    tryApplySingleKernelSyncOnlyLiteralMutationFastPathWithArgs(
      {
        state: args.state,
        writeFastPathLiteralToExistingCell,
        deferSingleCellKernelSync,
        makeSingleLiteralSkipMetrics,
      },
      request,
    )

  const tryApplyFormulaReplacementAsDirectScalarDeltaRoot = (request: {
    cellIndex: number
    oldNumber: number | undefined
    changedTopology: boolean
    postRecalcDirectFormulaIndices: DirectFormulaIndexCollection
    postRecalcDirectFormulaMetrics: DirectFormulaMetricCounts
  }): boolean =>
    tryApplyOperationFormulaReplacementAsDirectScalarDeltaRoot(
      {
        state: args.state,
        getSingleEntityDependent: args.getSingleEntityDependent,
        evaluateDirectScalarCurrentValue,
        tryMarkDirectScalarLinearDeltaClosure,
        applyDirectFormulaCurrentResult,
        countPostRecalcDirectFormulaMetric,
      },
      request,
    )

  const {
    tryApplyCoalescedDirectScalarLiteralBatch,
    tryApplyDenseRowPairDirectScalarLiteralBatch,
    tryApplyLookupOnlyNumericColumnLiteralBatch,
  } = createOperationDirectScalarBatchFastPaths({
    state: args.state,
    emitBatch,
    hasVolatileFormulas: args.hasVolatileFormulas,
    hasTrackedExactLookupDependents,
    hasTrackedSortedLookupDependents,
    hasTrackedDirectRangeDependents,
    canFastPathLiteralOverwrite: (cellIndex) => canFastPathLiteralOverwrite(cellIndex),
    canUseDirectFormulaPostRecalc,
    canSkipFormulaColumnVersion,
    directScalarCellNumericValue,
    writeNumericLiteralToCellStore,
    applyTerminalDirectFormulaNumericResult,
    applyDirectFormulaNumericResult,
    applyDirectFormulaCurrentResult,
    tryEvaluateDirectScalarWithPendingNumbers,
    tryEvaluateDirectScalarNumericWithPendingNumbers,
    planExactLookupNumericColumnWrite,
    planApproximateLookupNumericColumnWrite,
    getSingleEntityDependent: args.getSingleEntityDependent,
    getEntityDependents: args.getEntityDependents,
    materializeDeferredStructuralFormulaSources: args.materializeDeferredStructuralFormulaSources,
    beginMutationCollection: args.beginMutationCollection,
    ensureRecalcScratchCapacity: args.ensureRecalcScratchCapacity,
    resetMaterializedCellScratch: args.resetMaterializedCellScratch,
    getBatchMutationDepth: args.getBatchMutationDepth,
    setBatchMutationDepth: args.setBatchMutationDepth,
    markInputChanged: args.markInputChanged,
    markExplicitChanged: args.markExplicitChanged,
    getChangedInputBuffer: args.getChangedInputBuffer,
    deferKernelSync: args.deferKernelSync,
    composeDisjointEventChanges: args.composeDisjointEventChanges,
    captureChangedCells: args.captureChangedCells,
    invalidateExactLookupColumn: args.invalidateExactLookupColumn,
    invalidateSortedLookupColumn: args.invalidateSortedLookupColumn,
  })

  const { tryApplySingleExistingDirectLiteralMutation, applyExistingNumericCellMutationAtNow } =
    createOperationSingleExistingLiteralFastPath({
      state: args.state,
      hasVolatileFormulas: args.hasVolatileFormulas,
      getSingleEntityDependent: args.getSingleEntityDependent,
      noteAggregateLiteralWrite: args.noteAggregateLiteralWrite,
      evaluateDirectFormula: args.evaluateDirectFormula,
      invalidateExactLookupColumn: args.invalidateExactLookupColumn,
      invalidateSortedLookupColumn: args.invalidateSortedLookupColumn,
      hasTrackedExactLookupDependents,
      hasTrackedSortedLookupDependents,
      hasTrackedDirectRangeDependents,
      canSkipApproximateLookupNewNumericColumnWrite,
      writeNumericLiteralToExistingCell,
      deferSingleCellKernelSync,
      makeSingleLiteralSkipMetrics,
      canFastPathLiteralOverwrite: (cellIndex) => canFastPathLiteralOverwrite(cellIndex),
      directScalarCellNumericValue,
      tryApplyTrustedSingleRangeDirectAggregateExistingNumericMutation,
      tryApplyTrustedDirectScalarClosureExistingNumericMutation,
      tryApplyTrustedFormulaLeafExistingNumericMutation,
      tryApplyFormulaLeafExistingLiteralMutation,
      tryApplySingleDirectAggregateLiteralMutationFastPath,
      planExactLookupNumericColumnWrite,
      planApproximateLookupNumericColumnWrite,
      patchUniformLookupTailWrites,
      tryApplySingleKernelSyncOnlyLiteralMutationFastPath,
      tryApplySingleDirectFormulaLiteralMutationWithoutEvents,
      tryApplySingleDirectScalarLiteralMutationWithoutEvents,
      tryApplySingleDirectLookupOperandMutationFastPath,
      markPostRecalcDirectScalarNumericDependents,
      tryMarkDirectScalarLinearDeltaClosure,
      collectSingleAffectedDirectRangeDependent,
      collectAffectedDirectRangeDependents,
      applyDirectFormulaCurrentResult,
      applyDirectFormulaNumericDelta,
      applyDirectScalarCurrentValue,
      tryApplyDirectScalarDeltas,
      tryApplyDirectFormulaDeltas,
      countPostRecalcDirectFormulaMetric: (cellIndex, counts) => countPostRecalcDirectFormulaMetric(cellIndex, counts),
      hasDynamicFormulaDependents,
    })

  const countPostRecalcDirectFormulaMetric = (cellIndex: number, counts: DirectFormulaMetricCounts): void => {
    countOperationPostRecalcDirectFormulaMetric({
      formulas: args.state.formulas,
      cellIndex,
      counts,
    })
  }

  const refreshDependentRangesAndRebindFormulaDependents = (cellIndex: number, formulaChangedCount: number): number => {
    return refreshDependentRangesAndRebindOperationFormulaDependents({
      cellIndex,
      formulaChangedCount,
      getEntityDependents: args.getEntityDependents,
      collectFormulaDependents: args.collectFormulaDependents,
      refreshRangeDependencies: args.refreshRangeDependencies,
      rebindFormulaCells: args.rebindFormulaCells,
    })
  }

  const { applyDerivedOpNow, applySpillRangeOp, applyPivotUpsertOp, applyPivotDeleteOp } = createOperationDerivedOpApplier({
    state: {
      workbook: args.state.workbook,
      replicaState: args.state.replicaState,
    },
    reverseSpillEdges: args.reverseState.reverseSpillEdges,
    setEntityVersionForOp,
    materializePivot: args.materializePivot,
    clearOwnedPivot: args.clearOwnedPivot,
    rebindFormulaCells: args.rebindFormulaCells,
  })

  const applyBatchNow = createOperationBatchApplier({
    serviceArgs: args,
    emitBatch,
    replicaVersionWriter,
    isNullLiteralWriteNoOp,
    isClearCellNoOp,
    readCellValueForLookup,
    readExactNumericValueForLookup,
    hasTrackedExactLookupDependents,
    hasTrackedSortedLookupDependents,
    hasTrackedDirectRangeDependents,
    noteExactLookupLiteralWriteWhenDirty,
    noteSortedLookupLiteralWriteWhenDirty,
    markAffectedDirectRangeDependents,
    markPostRecalcDirectFormulaDependents,
    markDirectScalarDeltaClosure,
    collectAffectedDirectRangeDependents,
    tryApplyFormulaReplacementAsDirectScalarDeltaRoot,
    rebindDynamicFormulaDependents,
    refreshDependentRangesAndRebindFormulaDependents,
    pruneCellIfOrphaned,
    normalizeHistoryDependencyPlaceholder,
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
    applySpillRangeOp,
    applyPivotUpsertOp,
    applyPivotDeleteOp,
  })

  const applyCellMutationsAtNow = createOperationCellMutationApplier({
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
  })

  const __testHooks: Record<string, unknown> = ENGINE_OPERATION_TEST_HOOKS_ENABLED
    ? {
        canPatchUniformLookupTailWrite,
        canSkipApproximateLookupDirtyMark,
        canSkipApproximateLookupNewNumericColumnWrite,
        canSkipApproximateLookupNumericColumnWrite,
        canSkipExactLookupNumericColumnWrite,
        collectAffectedDirectRangeDependents,
        collectSingleAffectedDirectRangeDependent,
        directCriteriaMatchesChangedAggregateRow,
        isLocallySortedNumericWrite,
        isLocallySortedTextWrite,
        markAffectedApproximateLookupDependents,
        markAffectedDirectRangeDependents,
        patchUniformLookupTailWrites,
        planApproximateLookupNumericColumnWrite,
        planExactLookupNumericColumnWrite,
        planSingleApproximateLookupNumericColumnWrite,
        planSingleExactLookupNumericColumnWrite,
        readApproximateNumericValueAtForLookup,
        readApproximateNumericValueForLookup,
        readCellValueAtForLookup,
        readCellValueForLookup,
        readDirectCriteriaOperandValue,
        readExactNumericValueForLookup,
        tryApplyDenseRowPairDirectScalarLiteralBatch,
        tryApplyLookupOnlyNumericColumnLiteralBatch,
        tryApplySingleExistingDirectLiteralMutation,
        tryApplySingleDirectAggregateLiteralMutationFastPath,
        tryApplySingleDirectFormulaLiteralMutationWithoutEvents,
        tryApplySingleDirectLookupOperandMutationFastPath,
        tryApplySingleDirectScalarLiteralMutationWithoutEvents,
        tryApplySingleKernelSyncOnlyLiteralMutationFastPath,
        tryDirectCriteriaSumDelta,
      }
    : {}

  return {
    __testHooks,
    applyBatch(batch, source, potentialNewCells, preparedCellAddressesByOpIndex) {
      return Effect.try({
        try: () => {
          applyBatchNow(batch, source, potentialNewCells, preparedCellAddressesByOpIndex)
        },
        catch: (cause) =>
          new EngineMutationError({
            message: mutationErrorMessage(`Failed to apply ${source} batch`, cause),
            cause,
          }),
      })
    },
    applyCellMutationsAt(refs, batch, source, potentialNewCells) {
      return Effect.try({
        try: () => {
          applyCellMutationsAtNow(refs, batch, source, potentialNewCells)
        },
        catch: (cause) =>
          new EngineMutationError({
            message: mutationErrorMessage(`Failed to apply ${source} cell mutations`, cause),
            cause,
          }),
      })
    },
    applyCellMutationsAtNow,
    applyExistingNumericCellMutationAtNow,
    applyDerivedOp(op) {
      return Effect.try({
        try: () => applyDerivedOpNow(op),
        catch: (cause) =>
          new EngineMutationError({
            message: mutationErrorMessage(`Failed to apply derived operation ${op.kind}`, cause),
            cause,
          }),
      })
    },
  }
}
