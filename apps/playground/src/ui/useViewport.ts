import { useMemo, useSyncExternalStore } from "react";
import type { SpreadsheetEngine } from "@bilig/core";
import type { Viewport } from "@bilig/protocol";
import { selectors } from "@bilig/core";

export function useViewport(engine: SpreadsheetEngine, sheetName: string, viewport: Viewport) {
  const revision = useSyncExternalStore(
    engine.subscribe.bind(engine),
    () => engine.getLastMetrics().batchId,
    () => engine.getLastMetrics().batchId
  );

  return useMemo(
    () => selectors.selectViewportCells(engine, sheetName, viewport),
    [engine, revision, sheetName, viewport]
  );
}
