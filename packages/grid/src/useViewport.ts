import { useMemo, useSyncExternalStore } from "react";
import type { SpreadsheetEngine } from "@bilig/core";
import type { Viewport } from "@bilig/protocol";
import { formatAddress } from "@bilig/formula";
import { selectors } from "@bilig/core";

export function useSheetViewport(engine: SpreadsheetEngine, sheetName: string, viewport: Viewport) {
  const watchedAddresses = useMemo(() => {
    const addresses: string[] = [];
    for (let row = viewport.rowStart; row <= viewport.rowEnd; row += 1) {
      for (let col = viewport.colStart; col <= viewport.colEnd; col += 1) {
        addresses.push(formatAddress(row, col));
      }
    }
    return addresses;
  }, [viewport]);

  useSyncExternalStore(
    (listener) => engine.subscribeCells(sheetName, watchedAddresses, listener),
    () => engine.getLastMetrics().batchId,
    () => engine.getLastMetrics().batchId
  );

  return selectors.selectViewportCells(engine, sheetName, viewport);
}

export const useViewport = useSheetViewport;
