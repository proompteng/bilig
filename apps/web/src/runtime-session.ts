import type { Zero } from '@rocicorp/zero'
import { createWorkerEngineClient, type MessagePortLike } from '@bilig/worker-transport'
import { parseCellAddress } from '@bilig/formula'
import { isAuthoritativeWorkbookEventBatch, type AuthoritativeWorkbookEventBatch } from '@bilig/zero-sync'
import {
  isCellSnapshot,
  isWorkbookSnapshot,
  ValueTag,
  type CellSnapshot,
  type RecalcMetrics,
  type Viewport,
  type WorkbookSnapshot,
} from '@bilig/protocol'
import type { InstallAuthoritativeSnapshotInput, WorkbookWorkerBootstrapResult, WorkbookWorkerStateSnapshot } from './worker-runtime.js'
import type { WorkbookPerfSession } from './perf/workbook-perf.js'
import { ProjectedViewportStore } from './projected-viewport-store.js'
import { ZeroWorkbookRevisionSync, type WorkbookRevisionState } from './runtime-zero-revision-sync.js'

export interface WorkerHandle {
  readonly viewportStore: ProjectedViewportStore
}

export interface WorkerRuntimeSelection {
  readonly sheetName: string
  readonly address: string
}

export type WorkerRuntimeSessionPhase = 'hydratingLocal' | 'syncing' | 'reconciling' | 'recovering' | 'steady'

export interface WorkerRuntimeSessionCallbacks {
  readonly onRuntimeState: (runtimeState: WorkbookWorkerStateSnapshot) => void
  readonly onSelection: (selection: WorkerRuntimeSelection) => void
  readonly onError: (message: string) => void
  readonly onPhase?: (phase: WorkerRuntimeSessionPhase) => void
}

export type ZeroClient = Zero

export interface ZeroWorkbookSyncSource {
  materialize(query: unknown): unknown
}

export interface CreateWorkerRuntimeSessionInput {
  readonly documentId: string
  readonly replicaId: string
  readonly persistState: boolean
  readonly initialSelection: WorkerRuntimeSelection
  readonly perfSession?: WorkbookPerfSession
  readonly zero?: ZeroWorkbookSyncSource
  readonly fetchImpl?: typeof fetch
  readonly createWorker?: () => WorkerSessionPort
}

export interface WorkerRuntimeSessionController {
  readonly handle: WorkerHandle
  readonly runtimeState: WorkbookWorkerStateSnapshot
  readonly selection: WorkerRuntimeSelection
  readonly invoke: (method: string, ...args: unknown[]) => Promise<unknown>
  readonly setSelection: (selection: WorkerRuntimeSelection) => Promise<void>
  readonly subscribeViewport: (
    sheetName: string,
    viewport: Viewport,
    listener: (damage?: readonly { cell: readonly [number, number] }[]) => void,
    options?: { readonly initialPatch?: 'full' | 'none' },
    sheetViewId?: string,
  ) => () => void
  readonly dispose: () => void
}

interface WorkerSessionPort extends MessagePortLike {
  terminate?: () => void
}

const EMPTY_METRICS: RecalcMetrics = {
  batchId: 0,
  changedInputCount: 0,
  dirtyFormulaCount: 0,
  wasmFormulaCount: 0,
  jsFormulaCount: 0,
  rangeNodeVisits: 0,
  recalcMs: 0,
  compileMs: 0,
}
const EMPTY_UNSUBSCRIBE = () => {}
const BACKGROUND_RUNTIME_STATE_REFRESH_DELAY_MS = 96

function createInitialRuntimeState(documentId: string): WorkbookWorkerStateSnapshot {
  return {
    workbookName: documentId,
    sheets: [{ id: 1, name: 'Sheet1', order: 0 }],
    sheetNames: ['Sheet1'],
    definedNames: [],
    metrics: EMPTY_METRICS,
    syncState: 'syncing',
    localPersistenceMode: 'ephemeral',
    pendingMutationSummary: {
      activeCount: 0,
      failedCount: 0,
      firstFailed: null,
    },
  }
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function isWorkbookWorkerStateSnapshot(value: unknown): value is WorkbookWorkerStateSnapshot {
  if (!isRecord(value)) {
    return false
  }
  return (
    typeof value['workbookName'] === 'string' &&
    Array.isArray(value['sheetNames']) &&
    value['sheetNames'].every((sheetName) => typeof sheetName === 'string') &&
    (value['sheets'] === undefined ||
      (Array.isArray(value['sheets']) &&
        value['sheets'].every(
          (sheet) =>
            isRecord(sheet) &&
            typeof sheet['id'] === 'number' &&
            Number.isInteger(sheet['id']) &&
            sheet['id'] >= 0 &&
            typeof sheet['name'] === 'string' &&
            Number.isInteger(sheet['order']),
        ))) &&
    Array.isArray(value['definedNames']) &&
    typeof value['metrics'] === 'object' &&
    value['metrics'] !== null &&
    typeof value['syncState'] === 'string' &&
    (value['localPersistenceMode'] === undefined ||
      value['localPersistenceMode'] === 'persistent' ||
      value['localPersistenceMode'] === 'ephemeral' ||
      value['localPersistenceMode'] === 'follower')
  )
}

function isWorkbookWorkerBootstrapResult(value: unknown): value is WorkbookWorkerBootstrapResult {
  if (!isRecord(value)) {
    return false
  }
  return (
    typeof value['restoredFromPersistence'] === 'boolean' &&
    typeof value['requiresAuthoritativeHydrate'] === 'boolean' &&
    (value['localPersistenceMode'] === undefined ||
      value['localPersistenceMode'] === 'persistent' ||
      value['localPersistenceMode'] === 'ephemeral' ||
      value['localPersistenceMode'] === 'follower') &&
    isWorkbookWorkerStateSnapshot(value['runtimeState'])
  )
}

function isNumber(value: unknown): value is number {
  return typeof value === 'number'
}

function isInstallAuthoritativeSnapshotInput(value: unknown): value is InstallAuthoritativeSnapshotInput {
  if (!isRecord(value)) {
    return false
  }
  return (
    isWorkbookSnapshot(value['snapshot']) &&
    typeof value['authoritativeRevision'] === 'number' &&
    (value['mode'] === 'bootstrap' || value['mode'] === 'reconcile')
  )
}

async function loadAuthoritativeEventBatch(
  documentId: string,
  afterRevision: number,
  fetchImpl: typeof fetch,
): Promise<AuthoritativeWorkbookEventBatch> {
  const response = await fetchImpl(`/v2/documents/${encodeURIComponent(documentId)}/events?afterRevision=${String(afterRevision)}`, {
    headers: {
      accept: 'application/json',
    },
    cache: 'no-store',
  })
  if (!response.ok) {
    throw new Error(`Failed to load authoritative events (${response.status})`)
  }
  const parsed: unknown = JSON.parse(await response.text())
  if (!isAuthoritativeWorkbookEventBatch(parsed)) {
    throw new Error('Authoritative event payload does not match the expected schema')
  }
  return parsed
}

async function invokeWorkerMethod<T>(
  client: ReturnType<typeof createWorkerEngineClient>,
  method: string,
  guard: (value: unknown) => value is T,
  ...args: unknown[]
): Promise<T> {
  const value = await client.invoke(method, ...args)
  if (!guard(value)) {
    throw new Error(`Worker method ${method} returned an unexpected payload`)
  }
  return value
}

function emptyCellSnapshot(selection: WorkerRuntimeSelection): CellSnapshot {
  return {
    sheetName: selection.sheetName,
    address: selection.address,
    value: { tag: ValueTag.Empty },
    flags: 0,
    version: 0,
  }
}

function selectionViewport(selection: WorkerRuntimeSelection): Viewport {
  const parsed = parseCellAddress(selection.address, selection.sheetName)
  return {
    rowStart: parsed.row,
    rowEnd: parsed.row,
    colStart: parsed.col,
    colEnd: parsed.col,
  }
}

function toErrorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error)
}

function sameSelection(left: WorkerRuntimeSelection, right: WorkerRuntimeSelection): boolean {
  return left.sheetName === right.sheetName && left.address === right.address
}

interface LatestWorkbookSnapshot {
  readonly snapshot: WorkbookSnapshot
  readonly revision: number | null
}

function reconcileSelection(selection: WorkerRuntimeSelection, sheetNames: readonly string[]): WorkerRuntimeSelection {
  if (sheetNames.length === 0) {
    return selection
  }
  if (sheetNames.includes(selection.sheetName)) {
    return selection
  }
  return {
    sheetName: sheetNames[0]!,
    address: 'A1',
  }
}

function parseSnapshotRevisionHeader(value: string | null): number | null {
  if (!value) {
    return null
  }
  const parsed = Number.parseInt(value, 10)
  return Number.isFinite(parsed) && parsed >= 0 ? parsed : null
}

async function loadLatestWorkbookSnapshot(documentId: string, fetchImpl: typeof fetch): Promise<LatestWorkbookSnapshot | null> {
  const response = await fetchImpl(`/v2/documents/${encodeURIComponent(documentId)}/snapshot/latest`, {
    headers: {
      accept: 'application/json, application/vnd.bilig.workbook+json',
    },
    cache: 'no-store',
  })
  if (response.status === 204 || response.status === 404) {
    return null
  }
  if (!response.ok) {
    throw new Error(`Failed to load workbook snapshot (${response.status})`)
  }
  const parsed: unknown = JSON.parse(await response.text())
  if (!isWorkbookSnapshot(parsed)) {
    throw new Error('Workbook snapshot payload does not match the expected schema')
  }
  return {
    snapshot: parsed,
    revision: parseSnapshotRevisionHeader(response.headers.get('x-bilig-snapshot-cursor')),
  }
}

function createWorkbookWorker(): WorkerSessionPort {
  return new Worker(new URL('./workbook.worker.ts', import.meta.url), {
    type: 'module',
  }) as WorkerSessionPort
}

export async function createWorkerRuntimeSessionController(
  input: CreateWorkerRuntimeSessionInput,
  callbacks: WorkerRuntimeSessionCallbacks,
): Promise<WorkerRuntimeSessionController> {
  const workerPort = (input.createWorker ?? createWorkbookWorker)()
  const client = createWorkerEngineClient({ port: workerPort })
  const viewportStore = new ProjectedViewportStore(client)
  const handle: WorkerHandle = { viewportStore }
  const fetchImpl = input.fetchImpl ?? fetch
  let currentSelection = input.initialSelection
  let currentRuntimeState = createInitialRuntimeState(input.documentId)
  let disposed = false
  let bootstrapped = false
  let currentAuthoritativeRevision = 0
  let requestedAuthoritativeRevision = 0
  let pendingRevisionState: WorkbookRevisionState | null = null
  let rebaseQueue = Promise.resolve()
  let selectionViewportCleanup = EMPTY_UNSUBSCRIBE
  let currentPhase: WorkerRuntimeSessionPhase = 'hydratingLocal'
  let runtimeStateRefreshTimer: ReturnType<typeof setTimeout> | null = null
  const liveSync = input.zero
    ? new ZeroWorkbookRevisionSync({
        zero: input.zero,
        documentId: input.documentId,
        onRevisionState(revisionState) {
          pendingRevisionState = revisionState
          if (bootstrapped) {
            queueAuthoritativeRebase(revisionState)
          }
        },
      })
    : null

  const publishPhase = (phase: WorkerRuntimeSessionPhase) => {
    if (currentPhase === phase) {
      return
    }
    currentPhase = phase
    callbacks.onPhase?.(phase)
  }

  const publishRuntimeState = (runtimeState: WorkbookWorkerStateSnapshot) => {
    currentRuntimeState = runtimeState
    viewportStore.setKnownSheets(runtimeState.sheetNames)
    callbacks.onRuntimeState(runtimeState)
  }

  const subscribeProjectedViewport = (
    sheetName: string,
    viewport: Viewport,
    listener: (damage?: readonly { cell: readonly [number, number] }[]) => void,
    options?: { readonly initialPatch?: 'full' | 'none' },
  ): (() => void) => {
    return viewportStore.subscribeViewport(sheetName, viewport, listener, options)
  }

  const updateSelectionViewport = (selection: WorkerRuntimeSelection): void => {
    selectionViewportCleanup()
    selectionViewportCleanup = viewportStore.subscribeAuxiliaryViewport(selection.sheetName, selectionViewport(selection), () => {}, {
      initialPatch: 'none',
    })
  }

  const applySelection = async (selection: WorkerRuntimeSelection): Promise<CellSnapshot> => {
    currentSelection = selection
    callbacks.onSelection(selection)
    const snapshot = await loadSelectionCellSnapshot(selection)
    viewportStore.setCellSnapshot(snapshot ?? emptyCellSnapshot(selection))
    updateSelectionViewport(selection)
    return snapshot ?? emptyCellSnapshot(selection)
  }

  const loadSelectionCellSnapshot = async (selection: WorkerRuntimeSelection): Promise<CellSnapshot | null> => {
    return await invokeWorkerMethod(client, 'getCell', isCellSnapshot, selection.sheetName, selection.address)
  }

  const refreshRuntimeState = async (): Promise<void> => {
    const runtimeState = await invokeWorkerMethod(client, 'getRuntimeState', isWorkbookWorkerStateSnapshot)
    publishRuntimeState(runtimeState)
    const reconciledSelection = reconcileSelection(currentSelection, runtimeState.sheetNames)
    if (reconciledSelection.sheetName !== currentSelection.sheetName || reconciledSelection.address !== currentSelection.address) {
      await applySelection(reconciledSelection)
    }
  }

  const queueRuntimeStateRefresh = (): void => {
    if (runtimeStateRefreshTimer) {
      clearTimeout(runtimeStateRefreshTimer)
    }
    runtimeStateRefreshTimer = setTimeout(() => {
      runtimeStateRefreshTimer = null
      void (async () => {
        try {
          await refreshRuntimeState()
        } catch (error) {
          if (!disposed) {
            callbacks.onError(toErrorMessage(error))
          }
        }
      })()
    }, BACKGROUND_RUNTIME_STATE_REFRESH_DELAY_MS)
  }

  const syncSelectionAfterRuntimeState = async (runtimeState: WorkbookWorkerStateSnapshot): Promise<void> => {
    const reconciledSelection = reconcileSelection(currentSelection, runtimeState.sheetNames)
    if (reconciledSelection.sheetName !== currentSelection.sheetName || reconciledSelection.address !== currentSelection.address) {
      await applySelection(reconciledSelection)
      return
    }
    updateSelectionViewport(reconciledSelection)
  }

  const applyAuthoritativeEventBatch = async (eventBatch: AuthoritativeWorkbookEventBatch): Promise<boolean> => {
    if (
      eventBatch.events.length === 0 ||
      eventBatch.headRevision <= currentAuthoritativeRevision ||
      eventBatch.calculatedRevision < eventBatch.headRevision
    ) {
      return false
    }

    const runtimeState = await invokeWorkerMethod(
      client,
      'applyAuthoritativeEvents',
      isWorkbookWorkerStateSnapshot,
      eventBatch.events,
      eventBatch.headRevision,
    )
    currentAuthoritativeRevision = eventBatch.headRevision
    requestedAuthoritativeRevision = Math.max(requestedAuthoritativeRevision, currentAuthoritativeRevision)
    publishRuntimeState(runtimeState)
    input.perfSession?.markFirstAuthoritativePatchVisible()
    await syncSelectionAfterRuntimeState(runtimeState)
    return true
  }

  const runAuthoritativeRefresh = async (): Promise<void> => {
    if (disposed) {
      return
    }
    const eventBatch = await loadAuthoritativeEventBatch(input.documentId, currentAuthoritativeRevision, fetchImpl)
    if (await applyAuthoritativeEventBatch(eventBatch)) {
      await runAuthoritativeRebase()
    }
  }

  const queueAuthoritativeRefresh = (): void => {
    if (!liveSync) {
      return
    }
    const previousRebaseQueue = rebaseQueue
    rebaseQueue = (async () => {
      await previousRebaseQueue.catch(() => undefined)
      input.perfSession?.markFirstReconcileStarted()
      publishPhase('reconciling')
      try {
        await runAuthoritativeRefresh()
      } catch (error) {
        if (!disposed) {
          callbacks.onError(toErrorMessage(error))
        }
      } finally {
        if (!disposed) {
          publishPhase('steady')
          input.perfSession?.markFirstReconcileSettled()
        }
      }
    })()
  }

  const runAuthoritativeRebase = async (): Promise<void> => {
    if (disposed) {
      return
    }
    const targetRevision = requestedAuthoritativeRevision
    if (targetRevision <= currentAuthoritativeRevision) {
      return
    }
    const eventBatch = await loadAuthoritativeEventBatch(input.documentId, currentAuthoritativeRevision, fetchImpl)
    if (
      eventBatch.events.length > 0 &&
      eventBatch.headRevision >= targetRevision &&
      eventBatch.headRevision > currentAuthoritativeRevision
    ) {
      if (await applyAuthoritativeEventBatch(eventBatch)) {
        return runAuthoritativeRebase()
      }
    }
    publishPhase('recovering')
    const latestSnapshot = await loadLatestWorkbookSnapshot(input.documentId, fetchImpl)
    if (!latestSnapshot) {
      throw new Error('Authoritative workbook snapshot was not available for rebase')
    }
    const snapshotRevision = latestSnapshot.revision ?? targetRevision
    if (latestSnapshot.revision !== null && snapshotRevision <= currentAuthoritativeRevision) {
      throw new Error('Authoritative workbook snapshot was not newer than the current runtime state')
    }
    const runtimeState = await invokeWorkerMethod(client, 'installAuthoritativeSnapshot', isWorkbookWorkerStateSnapshot, {
      snapshot: latestSnapshot.snapshot,
      authoritativeRevision: snapshotRevision,
      mode: 'reconcile',
    } satisfies InstallAuthoritativeSnapshotInput)
    currentAuthoritativeRevision = snapshotRevision
    requestedAuthoritativeRevision = Math.max(requestedAuthoritativeRevision, currentAuthoritativeRevision)
    publishRuntimeState(runtimeState)
    await syncSelectionAfterRuntimeState(runtimeState)
    return runAuthoritativeRebase()
  }

  const queueAuthoritativeRebase = (revisionState: WorkbookRevisionState | null): void => {
    if (
      revisionState === null ||
      revisionState.calculatedRevision < revisionState.headRevision ||
      revisionState.headRevision <= currentAuthoritativeRevision
    ) {
      return
    }
    requestedAuthoritativeRevision = Math.max(requestedAuthoritativeRevision, revisionState.headRevision)
    const previousRebaseQueue = rebaseQueue
    rebaseQueue = (async () => {
      await previousRebaseQueue.catch(() => undefined)
      input.perfSession?.markFirstReconcileStarted()
      publishPhase('reconciling')
      try {
        await runAuthoritativeRebase()
      } catch (error) {
        if (!disposed) {
          callbacks.onError(toErrorMessage(error))
        }
      } finally {
        if (!disposed) {
          publishPhase('steady')
          input.perfSession?.markFirstReconcileSettled()
        }
      }
    })()
  }

  callbacks.onRuntimeState(currentRuntimeState)
  callbacks.onSelection(currentSelection)
  callbacks.onPhase?.(currentPhase)

  try {
    const bootstrap = await invokeWorkerMethod(client, 'bootstrap', isWorkbookWorkerBootstrapResult, {
      documentId: input.documentId,
      replicaId: input.replicaId,
      persistState: input.persistState,
    })
    publishRuntimeState(bootstrap.runtimeState)
    input.perfSession?.noteBootstrapResult(bootstrap)
    currentAuthoritativeRevision = await invokeWorkerMethod(client, 'getAuthoritativeRevision', isNumber)
    requestedAuthoritativeRevision = currentAuthoritativeRevision

    const shouldHydrateFromAuthoritativeState =
      !bootstrap.restoredFromPersistence ||
      bootstrap.requiresAuthoritativeHydrate ||
      (liveSync !== null && (bootstrap.runtimeState.pendingMutationSummary?.activeCount ?? 0) === 0)

    if (shouldHydrateFromAuthoritativeState) {
      publishPhase('syncing')
      const latestSnapshot = await loadLatestWorkbookSnapshot(input.documentId, fetchImpl)
      if (latestSnapshot) {
        const snapshotRevision = latestSnapshot.revision ?? currentAuthoritativeRevision
        const hydratedState = await invokeWorkerMethod(client, 'installAuthoritativeSnapshot', isWorkbookWorkerStateSnapshot, {
          snapshot: latestSnapshot.snapshot,
          authoritativeRevision: snapshotRevision,
          mode: 'bootstrap',
        } satisfies InstallAuthoritativeSnapshotInput)
        currentAuthoritativeRevision = Math.max(currentAuthoritativeRevision, snapshotRevision)
        requestedAuthoritativeRevision = Math.max(requestedAuthoritativeRevision, currentAuthoritativeRevision)
        publishRuntimeState(hydratedState)
        input.perfSession?.markFirstAuthoritativePatchVisible()
      }
    }

    bootstrapped = true
    publishPhase('steady')
    if (pendingRevisionState) {
      queueAuthoritativeRebase(pendingRevisionState)
    }
    await applySelection(reconcileSelection(currentSelection, currentRuntimeState.sheetNames))
    input.perfSession?.markFirstSelectionVisible()
  } catch (error) {
    liveSync?.dispose()
    client.dispose()
    workerPort.terminate?.()
    throw error
  }

  return {
    get handle() {
      return handle
    },
    get runtimeState() {
      return currentRuntimeState
    },
    get selection() {
      return currentSelection
    },
    async invoke(method, ...args) {
      if (method === 'refreshAuthoritativeEvents') {
        queueAuthoritativeRefresh()
        return undefined
      }
      try {
        if (method === 'installAuthoritativeSnapshot' && !isInstallAuthoritativeSnapshotInput(args[0])) {
          throw new Error('installAuthoritativeSnapshot requires a valid authoritative snapshot input')
        }
        const result = await client.invoke(method, ...args)
        if (method === 'enqueuePendingMutation') {
          queueRuntimeStateRefresh()
        } else if (method === 'markPendingMutationSubmitted') {
          queueRuntimeStateRefresh()
          queueAuthoritativeRefresh()
        } else if (method === 'markPendingMutationFailed' || method === 'retryPendingMutation') {
          await refreshRuntimeState()
        }
        if (method === 'renderCommit' || method === 'installAuthoritativeSnapshot') {
          await refreshRuntimeState()
        }
        return result
      } catch (error) {
        if (!disposed) {
          callbacks.onError(toErrorMessage(error))
        }
        throw error
      }
    },
    async setSelection(selection) {
      try {
        if (sameSelection(selection, currentSelection)) {
          return
        }
        await applySelection(selection)
      } catch (error) {
        if (!disposed) {
          callbacks.onError(toErrorMessage(error))
        }
        throw error
      }
    },
    subscribeViewport(sheetName, viewport, listener, options) {
      return subscribeProjectedViewport(sheetName, viewport, listener, options)
    },
    dispose() {
      if (disposed) {
        return
      }
      disposed = true
      if (runtimeStateRefreshTimer) {
        clearTimeout(runtimeStateRefreshTimer)
        runtimeStateRefreshTimer = null
      }
      selectionViewportCleanup()
      selectionViewportCleanup = EMPTY_UNSUBSCRIBE
      liveSync?.dispose()
      client.dispose()
      workerPort.terminate?.()
    },
  }
}
