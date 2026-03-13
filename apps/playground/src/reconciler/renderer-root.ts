import type { ReactNode } from "react";
import type { SpreadsheetEngine } from "@bilig/core";
import { WorkbookReconciler, type WorkbookContainer } from "./host-config.js";

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

  const fiberRoot = (WorkbookReconciler as any).createContainer(
    container,
    1,
    null,
    false,
    null,
    "",
    console.error,
    console.error,
    console.error,
    null
  );

  return {
    render(element: ReactNode) {
      container.lastError = null;
      return new Promise<void>((resolve, reject) => {
        try {
          (WorkbookReconciler as any).updateContainer(element, fiberRoot, null, () => {
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
          (WorkbookReconciler as any).updateContainer(null, fiberRoot, null, () => {
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
