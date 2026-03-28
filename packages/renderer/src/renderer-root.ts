import React, { type ReactNode } from "react";
import type { SpreadsheetEngine } from "@bilig/core";
import type { WorkbookContainer } from "./host-config.js";
import { collectDeleteOps } from "./commit-log.js";
import { createFiberRoot, updateFiberRoot } from "./compat.js";
import { RENDERER_KIND_PROP, type RendererKind } from "./renderer-kind.js";

export interface WorkbookRendererRoot {
  render(element: ReactNode): Promise<void>;
  unmount(): Promise<void>;
}

const COMMIT_SETTLEMENT_TURNS = 5;

function scheduleCommitSettlement(
  container: WorkbookContainer,
  finish: () => void,
  remainingTurns = COMMIT_SETTLEMENT_TURNS,
): void {
  // React/compat can surface commit errors after the update callback returns,
  // and the exact turn count is not stable under the full repo test load.
  // Wait through a bounded macrotask quiet window so deferred microtask errors
  // from compat land before we resolve the render/unmount promise.
  setTimeout(() => {
    if (container.lastError !== null || remainingTurns === 0) {
      finish();
      return;
    }
    scheduleCommitSettlement(container, finish, remainingTurns - 1);
  });
}

function schedulePostCommitObservation(container: WorkbookContainer, finish: () => void): void {
  // Let compat finish queuing any synchronous post-callback work before we start
  // watching the macrotask window for deferred errors.
  queueMicrotask(() => {
    scheduleCommitSettlement(container, finish);
  });
}

function isNamedComponentType(value: unknown): value is { displayName?: string; name?: string } {
  return typeof value === "function";
}

function rendererKindOf(value: unknown): RendererKind | null {
  if (typeof value === "function") {
    const marker = Reflect.get(value, RENDERER_KIND_PROP);
    if (marker === "Workbook" || marker === "Sheet" || marker === "Cell") {
      return marker;
    }
  }
  return null;
}

function kindOfNode(node: React.ReactElement): "Workbook" | "Sheet" | "Cell" | "wrapper" | null {
  if (typeof node.type === "string") {
    if (node.type === "Workbook" || node.type === "Sheet" || node.type === "Cell") {
      return node.type;
    }
    return "wrapper";
  }

  const markedKind = rendererKindOf(node.type);
  if (markedKind !== null) {
    return markedKind;
  }

  if (node.type === React.Fragment || node.type === React.StrictMode) {
    return "wrapper";
  }

  if (isNamedComponentType(node.type)) {
    const name = node.type.displayName ?? node.type.name;
    if (name === "Workbook" || name === "Sheet" || name === "Cell") {
      return name;
    }
    return "wrapper";
  }

  return "wrapper";
}

function toRenderableChildren(node: ReactNode): React.ReactNode[] {
  return React.Children.toArray(node);
}

type RendererElementProps = {
  name?: string;
  addr?: string;
  value?: unknown;
  formula?: unknown;
  children?: ReactNode;
};

function isRendererElement(node: ReactNode): node is React.ReactElement<RendererElementProps> {
  return React.isValidElement<RendererElementProps>(node);
}

function validateSheetChildren(node: ReactNode): void {
  for (const child of toRenderableChildren(node)) {
    if (typeof child === "string" || typeof child === "number") {
      throw new Error("Workbook DSL does not support text nodes.");
    }
    if (!isRendererElement(child)) {
      continue;
    }
    const kind = kindOfNode(child);
    if (kind === "wrapper") {
      validateSheetChildren(child.props.children);
      continue;
    }
    if (kind !== "Cell") {
      throw new Error("Only <Cell> can be nested inside <Sheet>.");
    }
    if (!child.props.addr) {
      throw new Error("<Cell> requires an addr prop.");
    }
    if (child.props.value !== undefined && child.props.formula !== undefined) {
      throw new Error("<Cell> cannot specify both value and formula.");
    }
  }
}

function validateWorkbookChildren(node: ReactNode): void {
  for (const child of toRenderableChildren(node)) {
    if (typeof child === "string" || typeof child === "number") {
      throw new Error("Workbook DSL does not support text nodes.");
    }
    if (!isRendererElement(child)) {
      continue;
    }
    const kind = kindOfNode(child);
    if (kind === "wrapper") {
      validateWorkbookChildren(child.props.children);
      continue;
    }
    if (kind !== "Sheet") {
      throw new Error("Only <Sheet> nodes can exist under <Workbook>.");
    }
    if (!child.props.name) {
      throw new Error("<Sheet> requires a name prop.");
    }
    validateSheetChildren(child.props.children);
  }
}

function validateRootElement(node: ReactNode): void {
  for (const child of toRenderableChildren(node)) {
    if (typeof child === "string" || typeof child === "number") {
      throw new Error("Workbook DSL does not support text nodes.");
    }
    if (!isRendererElement(child)) {
      continue;
    }
    const kind = kindOfNode(child);
    if (kind === "wrapper") {
      validateRootElement(child.props.children);
      return;
    }
    if (kind !== "Workbook") {
      throw new Error("Root descriptor must be a Workbook.");
    }
    validateWorkbookChildren(child.props.children);
    return;
  }
}

function normalizeNodeForRenderer(node: ReactNode): ReactNode {
  if (!isRendererElement(node)) {
    return node;
  }
  const kind = kindOfNode(node);
  const normalizedChildren = normalizeChildrenForRenderer(node.props.children);
  if (kind === "wrapper") {
    return null;
  }
  return React.cloneElement(node, undefined, ...normalizedChildren);
}

function normalizeChildrenForRenderer(node: ReactNode): React.ReactNode[] {
  const normalized: React.ReactNode[] = [];
  for (const child of toRenderableChildren(node)) {
    if (isRendererElement(child) && kindOfNode(child) === "wrapper") {
      normalized.push(...normalizeChildrenForRenderer(child.props.children));
      continue;
    }
    const next = normalizeNodeForRenderer(child);
    if (next === null || next === undefined || next === false) {
      continue;
    }
    normalized.push(next);
  }
  return normalized;
}

function normalizeRootElement(node: ReactNode): ReactNode {
  const normalizedChildren = normalizeChildrenForRenderer(node);
  if (normalizedChildren.length === 0) {
    return null;
  }
  if (normalizedChildren.length === 1) {
    return normalizedChildren[0];
  }
  return React.createElement(React.Fragment, null, ...normalizedChildren);
}

export function createWorkbookRendererRoot(engine: SpreadsheetEngine): WorkbookRendererRoot {
  const container: WorkbookContainer = {
    engine,
    root: null,
    pendingOps: [],
    shouldSyncSheetOrders: false,
    lastError: null,
  };

  const fiberRoot = createFiberRoot(container);

  return {
    render(element: ReactNode) {
      if (element === null || element === undefined || element === false) {
        return this.unmount();
      }
      container.lastError = null;
      try {
        validateRootElement(element);
      } catch (error) {
        return Promise.reject(error);
      }
      const normalizedElement = normalizeRootElement(element);
      return new Promise<void>((resolve, reject) => {
        let settled = false;
        const finish = () => {
          if (settled) {
            return;
          }
          settled = true;
          const error = container.lastError;
          container.lastError = null;
          if (error) {
            reject(error);
            return;
          }
          resolve();
        };
        try {
          updateFiberRoot(fiberRoot, normalizedElement, () => {
            schedulePostCommitObservation(container, finish);
          });
        } catch (error) {
          settled = true;
          reject(error);
        }
      });
    },
    unmount() {
      container.lastError = null;
      if (container.root) {
        const deleteOps = collectDeleteOps(container.root);
        container.root = null;
        if (deleteOps.length > 0) {
          container.engine.renderCommit(deleteOps);
        }
      }
      return new Promise<void>((resolve, reject) => {
        let settled = false;
        const finish = () => {
          if (settled) {
            return;
          }
          settled = true;
          const error = container.lastError;
          container.lastError = null;
          if (error) {
            reject(error);
            return;
          }
          resolve();
        };
        try {
          updateFiberRoot(fiberRoot, null, () => {
            schedulePostCommitObservation(container, finish);
          });
        } catch (error) {
          settled = true;
          reject(error);
        }
      });
    },
  };
}
