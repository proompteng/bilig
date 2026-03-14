import { useMemo, useSyncExternalStore } from "react";
import { selectors, type SpreadsheetEngine } from "@bilig/core";

export function useMetrics(engine: SpreadsheetEngine) {
  const revision = useSyncExternalStore(
    engine.subscribe.bind(engine),
    () => selectors.selectMetrics(engine).batchId,
    () => selectors.selectMetrics(engine).batchId
  );

  return useMemo(() => ({ ...selectors.selectMetrics(engine) }), [engine, revision]);
}
