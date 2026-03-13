import type { ReactNode } from "react";
import type { SpreadsheetEngine } from "@bilig/core";
import { WorkbookReconciler, type WorkbookContainer } from "./host-config.js";
import { emptyRenderModel } from "./descriptors.js";

export interface WorkbookRendererRoot {
  render(element: ReactNode): Promise<void>;
  unmount(): Promise<void>;
}

export function createWorkbookRendererRoot(engine: SpreadsheetEngine): WorkbookRendererRoot {
  const container: WorkbookContainer = {
    engine,
    root: null,
    model: emptyRenderModel()
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
      return new Promise<void>((resolve) => {
        (WorkbookReconciler as any).updateContainer(element, fiberRoot, null, resolve);
      });
    },
    unmount() {
      return new Promise<void>((resolve) => {
        (WorkbookReconciler as any).updateContainer(null, fiberRoot, null, resolve);
      });
    }
  };
}
