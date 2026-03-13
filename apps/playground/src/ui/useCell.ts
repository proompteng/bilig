import { useMemo, useSyncExternalStore } from "react";
import type { SpreadsheetEngine } from "@bilig/core";

export function useCell(engine: SpreadsheetEngine, sheetName: string, addr: string) {
  const revision = useSyncExternalStore(
    (listener) => engine.subscribeCell(sheetName, addr, listener),
    () => engine.getLastMetrics().batchId,
    () => engine.getLastMetrics().batchId
  );

  return useMemo(() => engine.getCell(sheetName, addr), [addr, engine, revision, sheetName]);
}
