import { useSyncExternalStore } from "react";
import type { SpreadsheetEngine } from "@bilig/core";

export function useMetrics(engine: SpreadsheetEngine) {
  return useSyncExternalStore(
    engine.subscribe.bind(engine),
    () => engine.getLastMetrics(),
    () => engine.getLastMetrics()
  );
}
