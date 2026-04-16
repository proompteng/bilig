import React from 'react'
import Reconciler from 'react-reconciler'
import { DefaultEventPriority } from 'react-reconciler/constants'
import type { CommitOp, SpreadsheetEngine } from '@bilig/core'
import { collectDeleteOps, collectMountOps, collectSheetOrderOps, normalizeCommitOps } from './commit-log.js'
import type { CellProps, Descriptor, SheetProps, WorkbookDescriptor, WorkbookProps } from './descriptors.js'
import { validateDescriptorTree } from './validation.js'

export interface WorkbookContainer {
  engine: SpreadsheetEngine
  root: WorkbookDescriptor | null
  pendingOps: import('@bilig/core').CommitOp[]
  shouldSyncSheetOrders: boolean
  lastError: Error | null
}

type WorkbookReconcilerInstance = {
  createContainer(...args: unknown[]): unknown
  updateContainer(...args: unknown[]): void
}

let currentUpdatePriority = DefaultEventPriority
const rootHostContext = Object.freeze({ kind: 'workbook-root' })
type WorkbookHostProps = WorkbookProps | SheetProps | CellProps

function isWorkbookProps(props: WorkbookHostProps): props is WorkbookProps {
  return !('addr' in props)
}

function isSheetProps(props: WorkbookHostProps): props is SheetProps {
  return !('addr' in props) && typeof props.name === 'string'
}

function isCellProps(props: WorkbookHostProps): props is CellProps {
  return 'addr' in props && typeof props.addr === 'string'
}

function isWorkbookContainer(value: unknown): value is WorkbookContainer {
  return (
    typeof value === 'object' &&
    value !== null &&
    'engine' in value &&
    'pendingOps' in value &&
    'shouldSyncSheetOrders' in value &&
    'lastError' in value
  )
}

function insertChild(parent: Descriptor, child: Descriptor, before?: Descriptor): void {
  if (parent.kind === 'Cell') {
    return
  }
  child.parent = parent
  if (parent.kind === 'Workbook') {
    if (child.kind !== 'Sheet') {
      throw new Error(`Cannot append ${child.kind} to ${parent.kind}.`)
    }
    if (!before) {
      parent.children.push(child)
      return
    }
    const index = parent.children.findIndex((candidate) => candidate === before)
    if (index === -1) {
      parent.children.push(child)
    } else {
      parent.children.splice(index, 0, child)
    }
    return
  }

  if (child.kind !== 'Cell') {
    throw new Error(`Cannot append ${child.kind} to ${parent.kind}.`)
  }
  if (!before) {
    parent.children.push(child)
    return
  }
  const index = parent.children.findIndex((candidate) => candidate === before)
  if (index === -1) {
    parent.children.push(child)
  } else {
    parent.children.splice(index, 0, child)
  }
}

function removeChild(parent: Descriptor, child: Descriptor): void {
  if (parent.kind === 'Cell') return
  if (parent.kind === 'Workbook' && child.kind !== 'Sheet') {
    return
  }
  if (parent.kind === 'Sheet' && child.kind !== 'Cell') {
    return
  }
  const index = parent.children.findIndex((candidate) => candidate === child)
  if (index >= 0) parent.children.splice(index, 1)
  child.parent = null
}

function containerFor(descriptor: Descriptor): WorkbookContainer {
  if (!isWorkbookContainer(descriptor.container)) {
    throw new Error('Descriptor is not attached to a workbook container.')
  }
  return descriptor.container
}

function pushCollectedOps(container: WorkbookContainer, collector: () => CommitOp[]): void {
  try {
    container.pendingOps.push(...collector())
  } catch (error) {
    container.pendingOps = []
    container.shouldSyncSheetOrders = false
    container.lastError = error instanceof Error ? error : new Error(String(error))
  }
}

function pushCellUpsert(
  ops: CommitOp[],
  sheetName: string,
  props: { addr: string; value?: CellProps['value']; formula?: string; format?: string },
): void {
  const op: CommitOp = {
    kind: 'upsertCell',
    sheetName,
    addr: props.addr,
  }
  if (props.value !== undefined) op.value = props.value
  if (props.formula !== undefined) op.formula = props.formula
  if (props.format !== undefined) op.format = props.format
  ops.push(op)
}

export const workbookHostConfig = {
  rendererPackageName: 'bilig-renderer',
  rendererVersion: '0.1.0',
  extraDevToolsConfig: null,
  supportsMutation: true,
  supportsPersistence: false,
  supportsHydration: false,
  supportsMicrotasks: true,
  isPrimaryRenderer: false,
  now: Date.now,
  getRootHostContext() {
    return rootHostContext
  },
  getChildHostContext() {
    return rootHostContext
  },
  getPublicInstance(instance: Descriptor) {
    return instance
  },
  prepareForCommit(container: WorkbookContainer) {
    container.pendingOps = []
    container.shouldSyncSheetOrders = false
    return rootHostContext
  },
  resetAfterCommit(container: WorkbookContainer) {
    try {
      validateDescriptorTree(container.root)
    } catch (error) {
      container.pendingOps = []
      container.shouldSyncSheetOrders = false
      container.lastError = error instanceof Error ? error : new Error(String(error))
      return
    }
    if (container.shouldSyncSheetOrders) {
      container.pendingOps.push(...collectSheetOrderOps(container.root))
    }
    const ops = normalizeCommitOps(container.pendingOps)
    container.pendingOps = []
    container.shouldSyncSheetOrders = false
    if (ops.length > 0) {
      container.engine.renderCommit(ops)
    }
  },
  preparePortalMount() {},
  createInstance(type: Descriptor['kind'], props: WorkbookHostProps, container: WorkbookContainer): Descriptor {
    switch (type) {
      case 'Workbook':
        if (!isWorkbookProps(props)) {
          throw new Error('Workbook props must not include cell fields.')
        }
        return { kind: 'Workbook', props, children: [], parent: null, container }
      case 'Sheet':
        if (!isSheetProps(props)) {
          throw new Error('Sheet props require a sheet name.')
        }
        return { kind: 'Sheet', props, children: [], parent: null, container }
      case 'Cell':
        if (!isCellProps(props)) {
          throw new Error('Cell props require an address.')
        }
        return { kind: 'Cell', props, parent: null, container }
      default:
        throw new Error('Unknown workbook host type.')
    }
  },
  appendInitialChild(parent: Descriptor, child: Descriptor) {
    insertChild(parent, child)
  },
  finalizeInitialChildren() {
    return false
  },
  shouldSetTextContent() {
    return false
  },
  createTextInstance() {
    throw new Error('Workbook DSL does not support text nodes.')
  },
  scheduleTimeout: setTimeout,
  cancelTimeout: clearTimeout,
  noTimeout: -1,
  scheduleMicrotask: queueMicrotask,
  getCurrentEventPriority() {
    return DefaultEventPriority
  },
  setCurrentUpdatePriority(priority: number) {
    currentUpdatePriority = priority
  },
  getCurrentUpdatePriority() {
    return currentUpdatePriority
  },
  resolveUpdatePriority() {
    return currentUpdatePriority
  },
  trackSchedulerEvent() {},
  resolveEventType() {
    return null
  },
  resolveEventTimeStamp() {
    return Date.now()
  },
  shouldAttemptEagerTransition() {
    return false
  },
  detachDeletedInstance() {},
  maySuspendCommit() {
    return false
  },
  maySuspendCommitOnUpdate() {
    return false
  },
  maySuspendCommitInSyncRender() {
    return false
  },
  preloadInstance() {},
  startSuspendingCommit() {},
  suspendInstance() {},
  waitForCommitToBeReady() {
    return null
  },
  NotPendingTransition: null,
  HostTransitionContext: React.createContext(null),
  resetFormInstance() {},
  bindToConsole() {
    return console.log.bind(console)
  },
  supportsTestSelectors: false,
  appendChild(parent: Descriptor, child: Descriptor) {
    insertChild(parent, child)
    const container = containerFor(parent)
    pushCollectedOps(container, () => collectMountOps(child))
    if (container.lastError === null && parent.kind === 'Workbook' && child.kind === 'Sheet') {
      container.shouldSyncSheetOrders = true
    }
  },
  appendChildToContainer(container: WorkbookContainer, child: Descriptor) {
    if (child.kind !== 'Workbook') {
      throw new Error('Only workbook descriptors can be attached to the root container.')
    }
    container.root = child
    pushCollectedOps(container, () => collectMountOps(child))
    if (container.lastError === null) {
      container.shouldSyncSheetOrders = true
    }
  },
  insertBefore(parent: Descriptor, child: Descriptor, beforeChild: Descriptor) {
    removeChild(parent, child)
    insertChild(parent, child, beforeChild)
    const container = containerFor(parent)
    pushCollectedOps(container, () => collectMountOps(child))
    if (container.lastError === null && parent.kind === 'Workbook' && child.kind === 'Sheet') {
      container.shouldSyncSheetOrders = true
    }
  },
  insertInContainerBefore(container: WorkbookContainer, child: Descriptor) {
    if (child.kind !== 'Workbook') {
      throw new Error('Only workbook descriptors can be attached to the root container.')
    }
    container.root = child
    pushCollectedOps(container, () => collectMountOps(child))
    if (container.lastError === null) {
      container.shouldSyncSheetOrders = true
    }
  },
  removeChild(parent: Descriptor, child: Descriptor) {
    const container = containerFor(parent)
    pushCollectedOps(container, () => collectDeleteOps(child))
    removeChild(parent, child)
    if (container.lastError === null && parent.kind === 'Workbook' && child.kind === 'Sheet') {
      container.shouldSyncSheetOrders = true
    }
  },
  removeChildFromContainer(container: WorkbookContainer, child: Descriptor) {
    if (child.kind !== 'Workbook') {
      throw new Error('Only workbook descriptors can be removed from the root container.')
    }
    pushCollectedOps(container, () => collectDeleteOps(child))
    if (container.root === child) {
      container.root = null
    }
  },
  prepareUpdate(_instance: Descriptor, _type: string, oldProps: unknown, newProps: unknown) {
    return oldProps === newProps ? null : true
  },
  commitUpdate(
    instance: Descriptor,
    _type: string,
    previousProps: WorkbookProps | SheetProps | CellProps,
    newProps: WorkbookProps | SheetProps | CellProps,
  ) {
    const container = containerFor(instance)

    if (instance.kind === 'Workbook') {
      if (!isWorkbookProps(previousProps) || !isWorkbookProps(newProps)) {
        throw new Error('Workbook updates require workbook props.')
      }
      instance.props = newProps
      if (previousProps.name !== instance.props.name) {
        container.pendingOps.push({
          kind: 'upsertWorkbook',
          name: instance.props.name ?? 'Workbook',
        })
      }
      return
    }

    if (instance.kind === 'Sheet') {
      if (!isSheetProps(previousProps) || !isSheetProps(newProps)) {
        throw new Error('Sheet updates require sheet props.')
      }
      instance.props = newProps
      const workbook = instance.parent
      const order = workbook?.kind === 'Workbook' ? workbook.children.indexOf(instance) : 0
      const previousName = previousProps.name
      if (previousProps.name !== instance.props.name) {
        container.pendingOps.push({ kind: 'deleteSheet', name: previousName })
        container.pendingOps.push({
          kind: 'upsertSheet',
          name: instance.props.name,
          order: Math.max(order, 0),
        })
        instance.children.forEach((cell) => {
          pushCellUpsert(container.pendingOps, instance.props.name, cell.props)
        })
        container.shouldSyncSheetOrders = true
      }
      return
    }

    if (!isCellProps(previousProps) || !isCellProps(newProps)) {
      throw new Error('Cell updates require cell props.')
    }
    instance.props = newProps
    const sheet = instance.parent
    if (sheet?.kind !== 'Sheet') {
      return
    }
    if (previousProps.addr !== instance.props.addr) {
      container.pendingOps.push({
        kind: 'deleteCell',
        sheetName: sheet.props.name,
        addr: previousProps.addr,
      })
    }
    if (
      previousProps.addr !== instance.props.addr ||
      previousProps.value !== instance.props.value ||
      previousProps.formula !== instance.props.formula ||
      previousProps.format !== instance.props.format
    ) {
      pushCellUpsert(container.pendingOps, sheet.props.name, instance.props)
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
    if (container.root) {
      pushCollectedOps(container, () => collectDeleteOps(container.root!))
    }
    container.root = null
  },
}

export const WorkbookReconciler: WorkbookReconcilerInstance = Reconciler(workbookHostConfig)
