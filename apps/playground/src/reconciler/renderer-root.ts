import type { ReactNode } from "react";
import type { SpreadsheetEngine } from "@bilig/core";
import type { WorkbookContainer } from "./host-config.js";
import { createFiberRoot, updateFiberRoot } from "./compat.js";

export interface WorkbookRendererRoot {
  render(element: ReactNode): Promise<void>;
  unmount(): Promise<void>;
}

export function createWorkbookRendererRoot(engine: SpreadsheetEngine): WorkbookRendererRoot {
  const container: WorkbookContainer = {
    engine,
    root: null,
    pendingOps: [],
    shouldSyncSheetOrders: false,
    lastError: null
  };

  const fiberRoot = createFiberRoot(container);

  return {
    render(element: ReactNode) {
      container.lastError = null;
      return new Promise<void>((resolve, reject) => {
        try {
          updateFiberRoot(fiberRoot, element, () => {
            if (container.lastError) {
              const error = container.lastError;
              container.lastError = null;
              reject(error);
              return;
            }
            resolve();
          });
        } catch (error) {
          reject(error);
        }
      });
    },
    unmount() {
      container.lastError = null;
      return new Promise<void>((resolve, reject) => {
        try {
          updateFiberRoot(fiberRoot, null, () => {
            if (container.lastError) {
              const error = container.lastError;
              container.lastError = null;
              reject(error);
              return;
            }
            resolve();
          });
        } catch (error) {
          reject(error);
        }
      });
    }
  };
}
