import { useMemo, useSyncExternalStore } from "react";
import { selectors, type SpreadsheetEngine } from "@bilig/core";

export function useCell(engine: SpreadsheetEngine, sheetName: string, addr: string) {
  const revision = useSyncExternalStore(
    (listener) => engine.subscribeCell(sheetName, addr, listener),
    () => engine.getLastMetrics().batchId,
    () => engine.getLastMetrics().batchId
  );

  return useMemo(() => selectors.selectCellSnapshot(engine, sheetName, addr), [addr, engine, revision, sheetName]);
}
