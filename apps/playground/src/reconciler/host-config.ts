import React from "react";
import Reconciler from "react-reconciler";
import { DefaultEventPriority } from "react-reconciler/constants";
import type { CommitOp, SpreadsheetEngine } from "@bilig/core";
import {
  collectDeleteOps,
  collectModelDeleteOps,
  collectMountOps,
  collectSheetOrderOps,
  normalizeCommitOps
} from "./commit-log.js";
import type { CellProps, Descriptor, RenderModel, SheetProps, WorkbookDescriptor, WorkbookProps } from "./descriptors.js";
import { buildRenderModel } from "./validation.js";

export interface WorkbookContainer {
  engine: SpreadsheetEngine;
  root: WorkbookDescriptor | null;
  model: RenderModel;
  pendingOps: import("@bilig/core").CommitOp[];
  shouldSyncSheetOrders: boolean;
  lastError: Error | null;
}

let currentUpdatePriority = DefaultEventPriority;
const rootHostContext = Object.freeze({ kind: "workbook-root" });

function insertChild(parent: Descriptor, child: Descriptor, before?: Descriptor): void {
  child.parent = parent;
  const list = "children" in parent ? parent.children : [];
  if (!before) {
    (list as Descriptor[]).push(child);
    return;
  }
  const index = list.indexOf(before as never);
  if (index === -1) {
    (list as Descriptor[]).push(child);
  } else {
    (list as Descriptor[]).splice(index, 0, child);
  }
}

function removeChild(parent: Descriptor, child: Descriptor): void {
  if (!("children" in parent)) return;
  const index = parent.children.indexOf(child as never);
  if (index >= 0) parent.children.splice(index, 1);
  child.parent = null;
}

function containerFor(descriptor: Descriptor): WorkbookContainer {
  return descriptor.container as WorkbookContainer;
}

function pushCollectedOps(container: WorkbookContainer, collector: () => CommitOp[]): void {
  try {
    container.pendingOps.push(...collector());
  } catch (error) {
    container.pendingOps = [];
    container.shouldSyncSheetOrders = false;
    container.lastError = error instanceof Error ? error : new Error(String(error));
  }
}

function pushCellUpsert(
  ops: CommitOp[],
  sheetName: string,
  props: { addr: string; value?: CellProps["value"]; formula?: string }
): void {
  const op: CommitOp = {
    kind: "upsertCell",
    sheetName,
    addr: props.addr
  };
  if (props.value !== undefined) op.value = props.value;
  if (props.formula !== undefined) op.formula = props.formula;
  ops.push(op);
}

const hostConfig = {
  rendererPackageName: "bilig-playground-reconciler",
  rendererVersion: "0.1.0",
  extraDevToolsConfig: null,
  supportsMutation: true,
  supportsPersistence: false,
  supportsHydration: false,
  supportsMicrotasks: true,
  isPrimaryRenderer: false,
  now: Date.now,
  getRootHostContext() {
    return rootHostContext;
  },
  getChildHostContext() {
    return rootHostContext;
  },
  getPublicInstance(instance: Descriptor) {
    return instance;
  },
  prepareForCommit(container: WorkbookContainer) {
    container.pendingOps = [];
    container.shouldSyncSheetOrders = false;
    return rootHostContext;
  },
  resetAfterCommit(container: WorkbookContainer) {
    let nextModel: RenderModel;
    try {
      nextModel = buildRenderModel(container.root);
    } catch (error) {
      container.pendingOps = [];
      container.shouldSyncSheetOrders = false;
      container.lastError = error instanceof Error ? error : new Error(String(error));
      return;
    }
    if (container.shouldSyncSheetOrders) {
      container.pendingOps.push(...collectSheetOrderOps(container.root));
    }
    const ops = normalizeCommitOps(container.pendingOps);
    container.pendingOps = [];
    container.shouldSyncSheetOrders = false;
    container.model = nextModel;
    if (ops.length > 0) {
      container.engine.renderCommit(ops);
    }
  },
  preparePortalMount() {},
  createInstance(type: string, props: WorkbookProps | SheetProps | CellProps, container: WorkbookContainer): Descriptor {
    switch (type) {
      case "Workbook":
        return { kind: "Workbook", props: props as WorkbookProps, children: [], parent: null, container };
      case "Sheet":
        return { kind: "Sheet", props: props as SheetProps, children: [], parent: null, container };
      case "Cell":
        return { kind: "Cell", props: props as CellProps, parent: null, container };
      default:
        throw new Error(`Unknown workbook host type: ${type}`);
    }
  },
  appendInitialChild(parent: Descriptor, child: Descriptor) {
    insertChild(parent, child);
  },
  finalizeInitialChildren() {
    return false;
  },
  shouldSetTextContent() {
    return false;
  },
  createTextInstance() {
    throw new Error("Workbook DSL does not support text nodes.");
  },
  scheduleTimeout: setTimeout,
  cancelTimeout: clearTimeout,
  noTimeout: -1,
  scheduleMicrotask: queueMicrotask,
  getCurrentEventPriority() {
    return DefaultEventPriority;
  },
  setCurrentUpdatePriority(priority: number) {
    currentUpdatePriority = priority;
  },
  getCurrentUpdatePriority() {
    return currentUpdatePriority;
  },
  resolveUpdatePriority() {
    return currentUpdatePriority;
  },
  trackSchedulerEvent() {},
  resolveEventType() {
    return null;
  },
  resolveEventTimeStamp() {
    return Date.now();
  },
  shouldAttemptEagerTransition() {
    return false;
  },
  detachDeletedInstance() {},
  maySuspendCommit() {
    return false;
  },
  maySuspendCommitOnUpdate() {
    return false;
  },
  maySuspendCommitInSyncRender() {
    return false;
  },
  preloadInstance() {},
  startSuspendingCommit() {},
  suspendInstance() {},
  waitForCommitToBeReady() {
    return null;
  },
  NotPendingTransition: null,
  HostTransitionContext: React.createContext(null),
  resetFormInstance() {},
  bindToConsole() {
    return console.log.bind(console);
  },
  supportsTestSelectors: false,
  appendChild(parent: Descriptor, child: Descriptor) {
    insertChild(parent, child);
    const container = containerFor(parent);
    pushCollectedOps(container, () => collectMountOps(child));
    if (container.lastError === null && parent.kind === "Workbook" && child.kind === "Sheet") {
      container.shouldSyncSheetOrders = true;
    }
  },
  appendChildToContainer(container: WorkbookContainer, child: WorkbookDescriptor) {
    container.root = child;
    pushCollectedOps(container, () => collectMountOps(child));
    if (container.lastError === null) {
      container.shouldSyncSheetOrders = true;
    }
  },
  insertBefore(parent: Descriptor, child: Descriptor, beforeChild: Descriptor) {
    removeChild(parent, child);
    insertChild(parent, child, beforeChild);
    const container = containerFor(parent);
    pushCollectedOps(container, () => collectMountOps(child));
    if (container.lastError === null && parent.kind === "Workbook" && child.kind === "Sheet") {
      container.shouldSyncSheetOrders = true;
    }
  },
  insertInContainerBefore(container: WorkbookContainer, child: WorkbookDescriptor) {
    container.root = child;
    pushCollectedOps(container, () => collectMountOps(child));
    if (container.lastError === null) {
      container.shouldSyncSheetOrders = true;
    }
  },
  removeChild(parent: Descriptor, child: Descriptor) {
    removeChild(parent, child);
    const container = containerFor(parent);
    pushCollectedOps(container, () => collectDeleteOps(child));
    if (container.lastError === null && parent.kind === "Workbook" && child.kind === "Sheet") {
      container.shouldSyncSheetOrders = true;
    }
  },
  removeChildFromContainer(container: WorkbookContainer, child: WorkbookDescriptor) {
    pushCollectedOps(container, () => collectDeleteOps(child));
    if (container.root === child) {
      container.root = null;
    }
  },
  prepareUpdate(_instance: Descriptor, _type: string, oldProps: unknown, newProps: unknown) {
    return oldProps === newProps ? null : true;
  },
  commitUpdate(
    instance: Descriptor,
    _type: string,
    previousProps: WorkbookProps | SheetProps | CellProps,
    newProps: WorkbookProps | SheetProps | CellProps
  ) {
    const container = containerFor(instance);

    if (instance.kind === "Workbook") {
      instance.props = newProps as WorkbookProps;
      if ((previousProps as WorkbookProps).name !== instance.props.name) {
        container.pendingOps.push({
          kind: "upsertWorkbook",
          name: instance.props.name ?? "Workbook"
        });
      }
      return;
    }

    if (instance.kind === "Sheet") {
      instance.props = newProps as SheetProps;
      const workbook = instance.parent;
      const order = workbook?.kind === "Workbook" ? workbook.children.indexOf(instance) : 0;
      const previousName = (previousProps as SheetProps).name;
      const previousOrder = container.model.sheets.get(previousName)?.order;
      if ((previousProps as SheetProps).name !== instance.props.name) {
        container.pendingOps.push({ kind: "deleteSheet", name: previousName });
        container.pendingOps.push({
          kind: "upsertSheet",
          name: instance.props.name,
          order: Math.max(order, 0)
        });
        instance.children.forEach((cell) => {
          pushCellUpsert(container.pendingOps, instance.props.name, cell.props);
        });
        container.shouldSyncSheetOrders = true;
      } else {
        const nextOrder = Math.max(order, 0);
        if (previousOrder !== undefined && previousOrder !== nextOrder) {
          container.pendingOps.push({
            kind: "upsertSheet",
            name: instance.props.name,
            order: nextOrder
          });
          container.shouldSyncSheetOrders = true;
        }
      }
      return;
    }

    instance.props = newProps as CellProps;
    const sheet = instance.parent;
    if (sheet?.kind !== "Sheet") {
      return;
    }
    if ((previousProps as CellProps).addr !== instance.props.addr) {
      container.pendingOps.push({
        kind: "deleteCell",
        sheetName: sheet.props.name,
        addr: (previousProps as CellProps).addr
      });
    }
    if (
      (previousProps as CellProps).addr !== instance.props.addr ||
      (previousProps as CellProps).value !== instance.props.value ||
      (previousProps as CellProps).formula !== instance.props.formula
    ) {
      pushCellUpsert(container.pendingOps, sheet.props.name, instance.props);
    }
  },
  commitTextUpdate() {},
  commitMount() {},
  resetTextContent() {},
  hideInstance() {},
  hideTextInstance() {},
  unhideInstance() {},
  unhideTextInstance() {},
  clearContainer(container: WorkbookContainer) {
    pushCollectedOps(container, () => collectModelDeleteOps(container.model));
    container.root = null;
  }
};

export const WorkbookReconciler = Reconciler(hostConfig as never);
