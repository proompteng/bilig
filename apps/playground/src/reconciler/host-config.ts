import React from "react";
import Reconciler from "react-reconciler";
import { DefaultEventPriority } from "react-reconciler/constants";
import type { SpreadsheetEngine } from "@bilig/core";
import { diffModels } from "./commit-log.js";
import type { CellDescriptor, CellProps, Descriptor, RenderModel, SheetDescriptor, SheetProps, WorkbookDescriptor, WorkbookProps } from "./descriptors.js";
import { buildRenderModel } from "./validation.js";
import { emptyRenderModel } from "./descriptors.js";

export interface WorkbookContainer {
  engine: SpreadsheetEngine;
  root: WorkbookDescriptor | null;
  model: RenderModel;
}

let currentUpdatePriority = DefaultEventPriority;

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
    return null;
  },
  getChildHostContext() {
    return null;
  },
  getPublicInstance(instance: Descriptor) {
    return instance;
  },
  prepareForCommit() {
    return null;
  },
  resetAfterCommit(container: WorkbookContainer) {
    const nextModel = buildRenderModel(container.root);
    const ops = diffModels(container.model, nextModel);
    container.model = nextModel;
    if (ops.length > 0) {
      container.engine.renderCommit(ops);
    }
  },
  preparePortalMount() {},
  createInstance(type: string, props: WorkbookProps | SheetProps | CellProps, container: WorkbookContainer): Descriptor {
    switch (type) {
      case "Workbook":
        return { kind: "Workbook", props: props as WorkbookProps, children: [], parent: null };
      case "Sheet":
        return { kind: "Sheet", props: props as SheetProps, children: [], parent: null };
      case "Cell":
        return { kind: "Cell", props: props as CellProps, parent: null };
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
  },
  appendChildToContainer(container: WorkbookContainer, child: WorkbookDescriptor) {
    container.root = child;
  },
  insertBefore(parent: Descriptor, child: Descriptor, beforeChild: Descriptor) {
    removeChild(parent, child);
    insertChild(parent, child, beforeChild);
  },
  insertInContainerBefore(container: WorkbookContainer, child: WorkbookDescriptor) {
    container.root = child;
  },
  removeChild(parent: Descriptor, child: Descriptor) {
    removeChild(parent, child);
  },
  removeChildFromContainer(container: WorkbookContainer, child: WorkbookDescriptor) {
    if (container.root === child) {
      container.root = null;
    }
  },
  prepareUpdate(_instance: Descriptor, _type: string, oldProps: unknown, newProps: unknown) {
    return oldProps === newProps ? null : true;
  },
  commitUpdate(instance: Descriptor, _payload: unknown, type: string, _oldProps: unknown, newProps: WorkbookProps | SheetProps | CellProps) {
    if (instance.kind === "Workbook") instance.props = newProps as WorkbookProps;
    if (instance.kind === "Sheet") instance.props = newProps as SheetProps;
    if (instance.kind === "Cell") instance.props = newProps as CellProps;
  },
  commitTextUpdate() {},
  commitMount() {},
  resetTextContent() {},
  hideInstance() {},
  hideTextInstance() {},
  unhideInstance() {},
  unhideTextInstance() {},
  clearContainer(container: WorkbookContainer) {
    container.root = null;
    container.model = emptyRenderModel();
  }
};

export const WorkbookReconciler = Reconciler(hostConfig as never);
