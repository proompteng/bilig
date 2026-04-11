import type { EngineEvent } from "@bilig/protocol";

export interface TrackedEngineEvent {
  invalidation: EngineEvent["invalidation"];
  changedCellIndices: Uint32Array;
  changedInputCount: number;
  explicitChangedCount?: number;
  hasInvalidatedRanges: boolean;
  hasInvalidatedRows: boolean;
  hasInvalidatedColumns: boolean;
}

function readExplicitChangedCount(event: EngineEvent): number | undefined {
  const explicitChangedCount = Reflect.get(event, "explicitChangedCount");
  return typeof explicitChangedCount === "number" && explicitChangedCount >= 0
    ? explicitChangedCount
    : undefined;
}

export function captureTrackedEngineEvent(event: EngineEvent): TrackedEngineEvent {
  return {
    invalidation: event.invalidation,
    changedCellIndices: Uint32Array.from(event.changedCellIndices),
    changedInputCount: event.metrics.changedInputCount,
    explicitChangedCount: readExplicitChangedCount(event),
    hasInvalidatedRanges: event.invalidatedRanges.length > 0,
    hasInvalidatedRows: event.invalidatedRows.length > 0,
    hasInvalidatedColumns: event.invalidatedColumns.length > 0,
  };
}
