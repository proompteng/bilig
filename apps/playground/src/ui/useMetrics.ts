import { useMemo, useSyncExternalStore } from "react";
import type { SpreadsheetEngine } from "@bilig/core";

export function useMetrics(engine: SpreadsheetEngine) {
  const revision = useSyncExternalStore(
    engine.subscribe.bind(engine),
    () => engine.getLastMetrics().batchId,
    () => engine.getLastMetrics().batchId
  );

  return useMemo(() => ({ ...engine.getLastMetrics() }), [engine, revision]);
}
