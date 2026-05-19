import type { EngineOp, EngineOpBatch } from '@bilig/workbook-domain'
import type { CellValue } from '@bilig/protocol'
import type { DirectFormulaIndexCollection } from './direct-formula-index-collection.js'
import type { OperationReplicaVersionWriter } from './operation-replica-helpers.js'
import type { DirectFormulaMetricCounts } from './operation-post-recalc-direct-formulas.js'
import type { OperationLookupAccess } from './operation-lookup-access.js'
import type { OperationLookupPlanner } from './operation-lookup-planner.js'
import type { OperationLookupDirtyMarkerService } from './operation-lookup-dirty-markers.js'
import type { OperationColumnDependencyTrackerService } from './operation-column-dependency-tracker.js'
import type { OperationDirectRangeDependentService } from './operation-direct-range-dependents.js'
import type { finalizeOperationRecalcAndEvents } from './operation-recalc-finalizer.js'
import type { OpOrder } from '../../replica-state.js'
import type { CreateEngineOperationServiceArgs, MutationSource } from './operation-service-types.js'

export type OperationBatchDirectFormulaCallbacks = Parameters<typeof finalizeOperationRecalcAndEvents>[0]['directFormulaCallbacks']
export type OperationBatchDirtyTraversalSkip = Parameters<
  typeof finalizeOperationRecalcAndEvents
>[0]['canSkipDirtyTraversalForChangedInputs']
export type OperationBatchCycleInputMarker = Parameters<typeof finalizeOperationRecalcAndEvents>[0]['markCycleMemberInputsChanged']
export type OperationBatchDerivedOp<K extends EngineOp['kind']> = Extract<EngineOp, { kind: K }>

export interface CreateOperationBatchApplierArgs {
  readonly serviceArgs: CreateEngineOperationServiceArgs
  readonly emitBatch: (batch: EngineOpBatch) => void
  readonly replicaVersionWriter: OperationReplicaVersionWriter
  readonly isNullLiteralWriteNoOp: (cellIndex: number) => boolean
  readonly isClearCellNoOp: (cellIndex: number) => boolean
  readonly readCellValueForLookup: OperationLookupAccess['readCellValueForLookup']
  readonly readExactNumericValueForLookup: OperationLookupAccess['readExactNumericValueForLookup']
  readonly hasTrackedExactLookupDependents: OperationColumnDependencyTrackerService['hasTrackedExactLookupDependents']
  readonly hasTrackedSortedLookupDependents: OperationColumnDependencyTrackerService['hasTrackedSortedLookupDependents']
  readonly hasTrackedDirectRangeDependents: OperationColumnDependencyTrackerService['hasTrackedDirectRangeDependents']
  readonly planExactLookupNumericColumnWrite: OperationLookupPlanner['planExactLookupNumericColumnWrite']
  readonly planApproximateLookupNumericColumnWrite: OperationLookupPlanner['planApproximateLookupNumericColumnWrite']
  readonly noteExactLookupLiteralWriteWhenDirty: OperationLookupDirtyMarkerService['noteExactLookupLiteralWriteWhenDirty']
  readonly noteSortedLookupLiteralWriteWhenDirty: OperationLookupDirtyMarkerService['noteSortedLookupLiteralWriteWhenDirty']
  readonly markAffectedDirectRangeDependents: OperationDirectRangeDependentService['markAffectedDirectRangeDependents']
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
  readonly collectAffectedDirectRangeDependents: OperationDirectRangeDependentService['collectAffectedDirectRangeDependents']
  readonly tryApplyFormulaReplacementAsDirectScalarDeltaRoot: (request: {
    readonly cellIndex: number
    readonly oldNumber: number | undefined
    readonly changedTopology: boolean
    readonly postRecalcDirectFormulaIndices: DirectFormulaIndexCollection
    readonly postRecalcDirectFormulaMetrics: DirectFormulaMetricCounts
  }) => boolean
  readonly rebindDynamicFormulaDependents: (cellIndex: number, formulaChangedCount: number) => number
  readonly refreshDependentRangesAndRebindFormulaDependents: (cellIndex: number, formulaChangedCount: number) => number
  readonly pruneCellIfOrphaned: (cellIndex: number) => void
  readonly normalizeHistoryDependencyPlaceholder: (cellIndex: number, source: MutationSource) => void
  readonly markCycleMemberInputsChanged: OperationBatchCycleInputMarker
  readonly hasCycleMembersNow: () => boolean
  readonly canSkipDirtyTraversalForChangedInputs: OperationBatchDirtyTraversalSkip
  readonly directFormulaCallbacks: OperationBatchDirectFormulaCallbacks
  readonly applySpillRangeOp: (op: OperationBatchDerivedOp<'upsertSpillRange' | 'deleteSpillRange'>, order: OpOrder) => number[]
  readonly applyPivotUpsertOp: (op: OperationBatchDerivedOp<'upsertPivotTable'>, order: OpOrder) => number[]
  readonly applyPivotDeleteOp: (op: OperationBatchDerivedOp<'deletePivotTable'>, order: OpOrder) => number[]
}
