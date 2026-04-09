import { Data } from "effect";

interface EngineServiceErrorDetails {
  readonly message: string;
  readonly cause?: unknown;
}

export class EngineMutationError extends Data.TaggedError("EngineMutationError")<EngineServiceErrorDetails> {}

export class EngineRecalcError extends Data.TaggedError("EngineRecalcError")<EngineServiceErrorDetails> {}

export class EngineSnapshotError extends Data.TaggedError("EngineSnapshotError")<EngineServiceErrorDetails> {}

export class EngineSyncError extends Data.TaggedError("EngineSyncError")<EngineServiceErrorDetails> {}

export class EngineHistoryError extends Data.TaggedError("EngineHistoryError")<EngineServiceErrorDetails> {}

export class EnginePivotError extends Data.TaggedError("EnginePivotError")<EngineServiceErrorDetails> {}

export class EngineStructureError extends Data.TaggedError("EngineStructureError")<EngineServiceErrorDetails> {}

export class EngineFormulaBindingError extends Data.TaggedError("EngineFormulaBindingError")<EngineServiceErrorDetails> {}

export class EngineFormulaGraphError extends Data.TaggedError("EngineFormulaGraphError")<EngineServiceErrorDetails> {}

export class EngineFormulaEvaluationError extends Data.TaggedError("EngineFormulaEvaluationError")<EngineServiceErrorDetails> {}
