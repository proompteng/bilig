import type { ReactNode } from "react";
import { WorkbookReconciler, type WorkbookContainer } from "./host-config.js";

function bindRendererError(container: WorkbookContainer) {
  return (error: unknown) => {
    container.lastError = error instanceof Error ? error : new Error(String(error));
  };
}

export function createFiberRoot(container: WorkbookContainer): unknown {
  const onError = bindRendererError(container);
  return WorkbookReconciler.createContainer(
    container,
    1,
    null,
    false,
    null,
    "",
    onError,
    onError,
    onError,
    null
  );
}

export function updateFiberRoot(root: unknown, element: ReactNode, callback: () => void): void {
  if (typeof root !== "object" || root === null) {
    throw new Error("Invalid fiber root");
  }
  WorkbookReconciler.updateContainer(element, root, null, callback);
}
