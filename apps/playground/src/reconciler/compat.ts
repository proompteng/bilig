import type { ReactNode } from "react";
import { WorkbookReconciler, type WorkbookContainer } from "./host-config.js";

interface ReconcilerCompat {
  createContainer(
    containerInfo: WorkbookContainer,
    tag: number,
    hydrationCallbacks: unknown,
    isStrictMode: boolean,
    concurrentUpdatesByDefaultOverride: unknown,
    identifierPrefix: string,
    onUncaughtError: (error: unknown) => void,
    onCaughtError: (error: unknown) => void,
    onRecoverableError: (error: unknown) => void,
    transitionCallbacks: unknown
  ): unknown;
  updateContainer(
    element: ReactNode,
    container: unknown,
    parentComponent: unknown,
    callback: () => void
  ): void;
}

const reconcilerCompat = WorkbookReconciler as unknown as ReconcilerCompat;

export function createFiberRoot(container: WorkbookContainer): unknown {
  return reconcilerCompat.createContainer(
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
}

export function updateFiberRoot(root: unknown, element: ReactNode, callback: () => void): void {
  reconcilerCompat.updateContainer(element, root, null, callback);
}
