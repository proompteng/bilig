import type { EngineEvent } from "@bilig/protocol";

export interface TrackedEngineEvent {
  invalidation: EngineEvent["invalidation"];
  changedCellIndices: Uint32Array;
  changedInputCount: number;
  hasInvalidatedRanges: boolean;
  hasInvalidatedRows: boolean;
  hasInvalidatedColumns: boolean;
}

export function captureTrackedEngineEvent(event: EngineEvent): TrackedEngineEvent {
  return {
    invalidation: event.invalidation,
    changedCellIndices: Uint32Array.from(event.changedCellIndices),
    changedInputCount: event.metrics.changedInputCount,
    hasInvalidatedRanges: event.invalidatedRanges.length > 0,
    hasInvalidatedRows: event.invalidatedRows.length > 0,
    hasInvalidatedColumns: event.invalidatedColumns.length > 0,
  };
}
