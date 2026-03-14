import React, { type ReactNode } from "react";
import type { SpreadsheetEngine } from "@bilig/core";
import type { WorkbookContainer } from "./host-config.js";
import { createFiberRoot, updateFiberRoot } from "./compat.js";

export interface WorkbookRendererRoot {
  render(element: ReactNode): Promise<void>;
  unmount(): Promise<void>;
}

function kindOfNode(node: React.ReactElement): "Workbook" | "Sheet" | "Cell" | "wrapper" | null {
  if (typeof node.type === "string") {
    if (node.type === "Workbook" || node.type === "Sheet" || node.type === "Cell") {
      return node.type;
    }
    return "wrapper";
  }

  if (node.type === React.Fragment || node.type === React.StrictMode) {
    return "wrapper";
  }

  if (typeof node.type === "function") {
    const name = (node.type as { displayName?: string; name?: string }).displayName
      ?? (node.type as { name?: string }).name;
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
      try {
        validateRootElement(element);
      } catch (error) {
        return Promise.reject(error);
      }
      return new Promise<void>((resolve, reject) => {
        const finish = () => {
          const error = container.lastError;
          container.lastError = null;
          if (error) {
            reject(error);
            return;
          }
          resolve();
        };
        try {
          updateFiberRoot(fiberRoot, element, () => {
            setTimeout(finish, 0);
          });
        } catch (error) {
          reject(error);
        }
      });
    },
    unmount() {
      container.lastError = null;
      return new Promise<void>((resolve, reject) => {
        const finish = () => {
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
            setTimeout(finish, 0);
          });
        } catch (error) {
          reject(error);
        }
      });
    }
  };
}
