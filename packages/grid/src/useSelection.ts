import { useCallback, useSyncExternalStore } from "react";
import type { SpreadsheetEngine } from "@bilig/core";

export function useSelection(engine: SpreadsheetEngine) {
  const selection = useSyncExternalStore(
    (listener) => engine.subscribeSelection(listener),
    () => engine.getSelectionState(),
    () => engine.getSelectionState()
  );
  const select = useCallback((nextSheetName: string, nextAddress: string | null) => {
    engine.setSelection(nextSheetName, nextAddress);
  }, [engine]);

  return {
    sheetName: selection.sheetName,
    address: selection.address,
    select
  };
}
