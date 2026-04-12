import type { EngineChangedCell, EngineEvent } from "@bilig/protocol";

export interface TrackedEngineEvent {
  invalidation: EngineEvent["invalidation"];
  changedCells: readonly EngineChangedCell[];
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
    changedCells: event.changedCells,
    changedInputCount: event.metrics.changedInputCount,
    explicitChangedCount: readExplicitChangedCount(event),
    hasInvalidatedRanges: event.invalidatedRanges.length > 0,
    hasInvalidatedRows: event.invalidatedRows.length > 0,
    hasInvalidatedColumns: event.invalidatedColumns.length > 0,
  };
}
