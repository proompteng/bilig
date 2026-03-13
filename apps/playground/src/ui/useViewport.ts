import { useSyncExternalStore } from "react";
import type { SpreadsheetEngine } from "@bilig/core";
import type { Viewport } from "@bilig/protocol";
import { selectors } from "@bilig/core";

export function useViewport(engine: SpreadsheetEngine, sheetName: string, viewport: Viewport) {
  return useSyncExternalStore(
    engine.subscribe.bind(engine),
    () => selectors.selectViewportCells(engine, sheetName, viewport),
    () => selectors.selectViewportCells(engine, sheetName, viewport)
  );
}
