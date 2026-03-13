import { useSyncExternalStore } from "react";
import type { SpreadsheetEngine } from "@bilig/core";

export function useCell(engine: SpreadsheetEngine, sheetName: string, addr: string) {
  return useSyncExternalStore(
    engine.subscribe.bind(engine),
    () => engine.getCell(sheetName, addr),
    () => engine.getCell(sheetName, addr)
  );
}
