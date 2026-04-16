import { Data } from 'effect'

interface EngineServiceErrorDetails {
  readonly message: string
  readonly cause?: unknown
}

const TaggedEngineMutationError = Data.TaggedError('EngineMutationError')<EngineServiceErrorDetails>
const TaggedEngineRecalcError = Data.TaggedError('EngineRecalcError')<EngineServiceErrorDetails>
const TaggedEngineSnapshotError = Data.TaggedError('EngineSnapshotError')<EngineServiceErrorDetails>
const TaggedEngineSyncError = Data.TaggedError('EngineSyncError')<EngineServiceErrorDetails>
const TaggedEngineHistoryError = Data.TaggedError('EngineHistoryError')<EngineServiceErrorDetails>
const TaggedEnginePivotError = Data.TaggedError('EnginePivotError')<EngineServiceErrorDetails>
const TaggedEngineStructureError = Data.TaggedError('EngineStructureError')<EngineServiceErrorDetails>
const TaggedEngineFormulaBindingError = Data.TaggedError('EngineFormulaBindingError')<EngineServiceErrorDetails>
const TaggedEngineFormulaGraphError = Data.TaggedError('EngineFormulaGraphError')<EngineServiceErrorDetails>
const TaggedEngineFormulaEvaluationError = Data.TaggedError('EngineFormulaEvaluationError')<EngineServiceErrorDetails>
const TaggedEngineCellStateError = Data.TaggedError('EngineCellStateError')<EngineServiceErrorDetails>
const TaggedEngineTraversalError = Data.TaggedError('EngineTraversalError')<EngineServiceErrorDetails>
const TaggedEngineMaintenanceError = Data.TaggedError('EngineMaintenanceError')<EngineServiceErrorDetails>
const TaggedEngineRuntimeScratchError = Data.TaggedError('EngineRuntimeScratchError')<EngineServiceErrorDetails>

export class EngineMutationError extends TaggedEngineMutationError {}

export class EngineRecalcError extends TaggedEngineRecalcError {}

export class EngineSnapshotError extends TaggedEngineSnapshotError {}

export class EngineSyncError extends TaggedEngineSyncError {}

export class EngineHistoryError extends TaggedEngineHistoryError {}

export class EnginePivotError extends TaggedEnginePivotError {}

export class EngineStructureError extends TaggedEngineStructureError {}

export class EngineFormulaBindingError extends TaggedEngineFormulaBindingError {}

export class EngineFormulaGraphError extends TaggedEngineFormulaGraphError {}

export class EngineFormulaEvaluationError extends TaggedEngineFormulaEvaluationError {}

export class EngineCellStateError extends TaggedEngineCellStateError {}

export class EngineTraversalError extends TaggedEngineTraversalError {}

export class EngineMaintenanceError extends TaggedEngineMaintenanceError {}

export class EngineRuntimeScratchError extends TaggedEngineRuntimeScratchError {}
