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
import { isInstallBenchmarkCorpusResult } from './benchmark-corpus-result.js'
import { isWorkerRuntimeStateSnapshot, normalizeWorkerRuntimeStateSnapshot } from './worker-runtime-state.js'
import type { WorkbookPerfSession } from './perf/workbook-perf.js'
import { ProjectedViewportStore } from './projected-viewport-store.js'
import { ZeroWorkbookRevisionSync, type WorkbookRevisionState } from './runtime-zero-revision-sync.js'
import { loadPersistedWorkbookMutationJournal, persistWorkbookMutationJournal } from './workbook-local-mutation-journal-persistence.js'
import { isPendingWorkbookMutationList } from './workbook-sync.js'

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
  readonly authoritativeSyncEnabled?: boolean
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
const MUTATION_JOURNAL_METHODS = new Set([
  'enqueuePendingMutation',
  'recordPendingMutationAttempt',
  'markPendingMutationSubmitted',
  'markPendingMutationFailed',
  'retryPendingMutation',
  'ackPendingMutation',
  'installAuthoritativeSnapshot',
])

function createInitialRuntimeState(documentId: string): WorkbookWorkerStateSnapshot {
  return {
    workbookName: documentId,
    sheets: [{ id: 1, name: 'Sheet1', order: 0 }],
    sheetNames: ['Sheet1'],
    definedNames: [],
    metrics: EMPTY_METRICS,
    syncState: 'syncing',
    localHistoryState: {
      canUndo: false,
      canRedo: false,
    },
    authoritativeRevision: 0,
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
  return isWorkerRuntimeStateSnapshot(value)
}

function isWorkbookWorkerBootstrapResult(value: unknown): value is WorkbookWorkerBootstrapResult {
  if (!isRecord(value)) {
    return false
  }
  return (
    typeof value['restoredFromPersistence'] === 'boolean' &&
    typeof value['requiresAuthoritativeHydrate'] === 'boolean' &&
    (value['localPersistenceMode'] === undefined || value['localPersistenceMode'] === 'ephemeral') &&
    isWorkbookWorkerStateSnapshot(value['runtimeState'])
  )
}

function isNonNegativeInteger(value: unknown): value is number {
  return typeof value === 'number' && Number.isInteger(value) && value >= 0
}

function isInstallAuthoritativeSnapshotInput(value: unknown): value is InstallAuthoritativeSnapshotInput {
  if (!isRecord(value)) {
    return false
  }
  return (
    isWorkbookSnapshot(value['snapshot']) &&
    isNonNegativeInteger(value['authoritativeRevision']) &&
    (value['mode'] === 'bootstrap' || value['mode'] === 'reconcile')
  )
}

async function parseJsonResponse(response: Response, context: string): Promise<unknown> {
  try {
    return JSON.parse(await response.text()) as unknown
  } catch {
    throw new Error(`${context} response returned malformed JSON`)
  }
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
  const parsed = await parseJsonResponse(response, 'Authoritative events')
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

function shouldForceSelectionHydration(current: CellSnapshot | undefined, incoming: CellSnapshot): boolean {
  if (!current) {
    return false
  }
  if ((current.input !== undefined || current.formula !== undefined) && current.version > incoming.version) {
    return false
  }
  return true
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

export function parseSnapshotRevisionHeader(value: string | null): number | null {
  const trimmed = value?.trim()
  if (!trimmed || !/^(0|[1-9]\d*)$/u.test(trimmed)) {
    return null
  }
  const parsed = Number(trimmed)
  return Number.isSafeInteger(parsed) ? parsed : null
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
  const parsed = await parseJsonResponse(response, 'Workbook snapshot')
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
  const authoritativeSyncEnabled = input.authoritativeSyncEnabled ?? true
  const restoredMutationJournal = input.persistState ? loadPersistedWorkbookMutationJournal(input.documentId) : null
  let currentSelection = input.initialSelection
  let currentRuntimeState = createInitialRuntimeState(input.documentId)
  viewportStore.setKnownSheets(currentRuntimeState.sheetNames)
  viewportStore.setSheetIdentities(currentRuntimeState.sheets ?? [])
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

  const publishRuntimeState = (runtimeState: WorkbookWorkerStateSnapshot): WorkbookWorkerStateSnapshot => {
    const normalizedRuntimeState = normalizeWorkerRuntimeStateSnapshot(runtimeState) ?? runtimeState
    const runtimeStateWithRevision: WorkbookWorkerStateSnapshot = {
      ...normalizedRuntimeState,
      authoritativeRevision: normalizedRuntimeState.authoritativeRevision ?? currentAuthoritativeRevision,
    }
    currentRuntimeState = runtimeStateWithRevision
    viewportStore.setKnownSheets(runtimeStateWithRevision.sheetNames)
    viewportStore.setSheetIdentities(runtimeStateWithRevision.sheets ?? [])
    callbacks.onRuntimeState(runtimeStateWithRevision)
    return runtimeStateWithRevision
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

  const hydrateSelectionCell = async (
    selection: WorkerRuntimeSelection,
    options: { readonly force?: boolean; readonly forceOptimistic?: boolean } = {},
  ): Promise<CellSnapshot> => {
    const snapshot = await loadSelectionCellSnapshot(selection)
    const nextSnapshot = snapshot ?? emptyCellSnapshot(selection)
    viewportStore.setCellSnapshot(nextSnapshot, {
      force:
        options.force === true ||
        shouldForceSelectionHydration(viewportStore.peekCell(selection.sheetName, selection.address), nextSnapshot),
      forceOptimistic: options.forceOptimistic === true,
    })
    updateSelectionViewport(selection)
    return nextSnapshot
  }

  const applySelection = async (selection: WorkerRuntimeSelection): Promise<CellSnapshot> => {
    currentSelection = selection
    callbacks.onSelection(selection)
    return hydrateSelectionCell(selection)
  }

  const loadSelectionCellSnapshot = async (selection: WorkerRuntimeSelection): Promise<CellSnapshot | null> => {
    return await invokeWorkerMethod(client, 'getCell', isCellSnapshot, selection.sheetName, selection.address)
  }

  const refreshRuntimeState = async (): Promise<void> => {
    const runtimeState = await invokeWorkerMethod(client, 'getRuntimeState', isWorkbookWorkerStateSnapshot)
    const publishedRuntimeState = publishRuntimeState(runtimeState)
    const reconciledSelection = reconcileSelection(currentSelection, publishedRuntimeState.sheetNames)
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
    const hasActivePendingMutation = (runtimeState.pendingMutationSummary?.activeCount ?? 0) > 0
    await hydrateSelectionCell(reconciledSelection, {
      force: !hasActivePendingMutation,
      forceOptimistic: !hasActivePendingMutation,
    })
  }

  const persistMutationJournal = async (): Promise<void> => {
    if (!input.persistState) {
      return
    }
    const entries = await invokeWorkerMethod(client, 'listMutationJournalEntries', isPendingWorkbookMutationList)
    persistWorkbookMutationJournal(input.documentId, entries)
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
    const publishedRuntimeState = publishRuntimeState(runtimeState)
    input.perfSession?.markFirstAuthoritativePatchVisible()
    await syncSelectionAfterRuntimeState(publishedRuntimeState)
    await persistMutationJournal()
    return true
  }

  const runAuthoritativeRefresh = async (): Promise<void> => {
    if (disposed || !authoritativeSyncEnabled) {
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

  const refreshAuthoritativeEventsNow = async (targetRevision: number | null): Promise<void> => {
    if (!authoritativeSyncEnabled) {
      return
    }
    if (targetRevision !== null) {
      requestedAuthoritativeRevision = Math.max(requestedAuthoritativeRevision, targetRevision)
    }
    const previousRebaseQueue = rebaseQueue
    rebaseQueue = (async () => {
      await previousRebaseQueue.catch(() => undefined)
      input.perfSession?.markFirstReconcileStarted()
      publishPhase('reconciling')
      try {
        await runAuthoritativeRefresh()
        await runAuthoritativeRebase()
      } catch (error) {
        if (!disposed) {
          callbacks.onError(toErrorMessage(error))
        }
        throw error
      } finally {
        if (!disposed) {
          publishPhase('steady')
          input.perfSession?.markFirstReconcileSettled()
        }
      }
    })()
    await rebaseQueue
  }

  const runAuthoritativeRebase = async (): Promise<void> => {
    if (disposed || !authoritativeSyncEnabled) {
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
    const publishedRuntimeState = publishRuntimeState(runtimeState)
    await syncSelectionAfterRuntimeState(publishedRuntimeState)
    await persistMutationJournal()
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
      ...(restoredMutationJournal
        ? {
            mutationJournalEntries: restoredMutationJournal.mutationJournalEntries,
            nextPendingMutationSeq: restoredMutationJournal.nextPendingMutationSeq,
          }
        : {}),
    })
    const bootstrapRuntimeState = publishRuntimeState(bootstrap.runtimeState)
    input.perfSession?.noteBootstrapResult(bootstrap)
    currentAuthoritativeRevision = await invokeWorkerMethod(client, 'getAuthoritativeRevision', isNonNegativeInteger)
    requestedAuthoritativeRevision = currentAuthoritativeRevision

    const activePendingMutationCount = bootstrapRuntimeState.pendingMutationSummary?.activeCount ?? 0
    const requiresAuthoritativeSnapshot = !bootstrap.restoredFromPersistence || bootstrap.requiresAuthoritativeHydrate
    const shouldRefreshCleanAuthoritativeState = activePendingMutationCount === 0

    if (authoritativeSyncEnabled && (requiresAuthoritativeSnapshot || shouldRefreshCleanAuthoritativeState)) {
      publishPhase('syncing')
      let latestSnapshot: LatestWorkbookSnapshot | null = null
      try {
        latestSnapshot = await loadLatestWorkbookSnapshot(input.documentId, fetchImpl)
      } catch (error) {
        if (requiresAuthoritativeSnapshot) {
          throw error
        }
        callbacks.onError(toErrorMessage(error))
      }
      if (latestSnapshot) {
        const snapshotRevision = latestSnapshot.revision ?? currentAuthoritativeRevision
        const shouldInstallSnapshot = latestSnapshot.revision === null || snapshotRevision >= currentAuthoritativeRevision
        if (shouldInstallSnapshot) {
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
        const targetRevision = typeof args[0] === 'number' && Number.isInteger(args[0]) && args[0] >= 0 ? args[0] : null
        await refreshAuthoritativeEventsNow(targetRevision)
        await persistMutationJournal()
        return undefined
      }
      try {
        if (method === 'installAuthoritativeSnapshot' && !isInstallAuthoritativeSnapshotInput(args[0])) {
          throw new Error('installAuthoritativeSnapshot requires a valid authoritative snapshot input')
        }
        if (method === 'installBenchmarkCorpus' && typeof args[0] !== 'string') {
          throw new Error('installBenchmarkCorpus requires a benchmark corpus id')
        }
        if (method === 'installAuthoritativeSnapshot' || method === 'installBenchmarkCorpus') {
          viewportStore.resetProjectionState()
        }
        const result = await client.invoke(method, ...args)
        if (MUTATION_JOURNAL_METHODS.has(method)) {
          await persistMutationJournal()
        }
        if (method === 'installBenchmarkCorpus' && !isInstallBenchmarkCorpusResult(result)) {
          throw new Error('installBenchmarkCorpus returned an unexpected payload')
        }
        if (method === 'enqueuePendingMutation') {
          queueRuntimeStateRefresh()
        } else if (method === 'markPendingMutationSubmitted') {
          queueRuntimeStateRefresh()
          queueAuthoritativeRefresh()
        } else if (method === 'markPendingMutationFailed' || method === 'retryPendingMutation') {
          await refreshRuntimeState()
        }
        if (
          method === 'renderCommit' ||
          method === 'undoLocalChange' ||
          method === 'redoLocalChange' ||
          method === 'installAuthoritativeSnapshot' ||
          method === 'installBenchmarkCorpus'
        ) {
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
          await applySelection(selection)
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
