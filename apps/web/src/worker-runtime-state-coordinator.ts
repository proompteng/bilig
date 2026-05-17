import type { RecalcMetrics, SyncState } from '@bilig/protocol'
import { buildWorkerRuntimeLocalHistoryState } from './worker-runtime-local-history.js'
import {
  EMPTY_RUNTIME_METRICS,
  buildCachedWorkerRuntimeState,
  buildWorkerRuntimeStateFromEngine,
  cloneRuntimeMetrics,
  cloneWorkerRuntimeState,
  withExternalSyncState,
  type WorkbookPendingMutationSummarySnapshot,
  type WorkbookWorkerStateSnapshot,
  type WorkerRuntimeStateEngine,
} from './worker-runtime-state.js'

export type { WorkerRuntimeStateEngine } from './worker-runtime-state.js'

export class WorkerRuntimeStateCoordinator {
  private externalSyncState: SyncState | null = null
  private runtimeStateCache: WorkbookWorkerStateSnapshot | null = null
  private readonly localPersistenceMode = 'ephemeral' as const

  constructor(
    private readonly options: {
      readonly getEngine: () => WorkerRuntimeStateEngine | null
      readonly getAuthoritativeRevision: () => number
      readonly buildPendingMutationSummary: () => WorkbookPendingMutationSummarySnapshot
    },
  ) {}

  reset(): void {
    this.externalSyncState = null
    this.runtimeStateCache = null
  }

  getLocalPersistenceMode(): 'ephemeral' {
    return this.localPersistenceMode
  }

  setExternalSyncState(syncState: SyncState | null, requireEngine?: () => WorkerRuntimeStateEngine): WorkbookWorkerStateSnapshot {
    this.externalSyncState = syncState
    return this.getRuntimeState(requireEngine)
  }

  getRuntimeState(
    requireEngine: () => WorkerRuntimeStateEngine = () => {
      const engine = this.options.getEngine()
      if (!engine) {
        throw new Error('Workbook worker runtime has not been bootstrapped')
      }
      return engine
    },
  ): WorkbookWorkerStateSnapshot {
    const cachedState = this.runtimeStateCache
    if (cachedState) {
      return buildCachedWorkerRuntimeState({
        cachedState,
        externalSyncState: this.externalSyncState,
        localHistoryState: buildWorkerRuntimeLocalHistoryState(this.options.getEngine()),
        authoritativeRevision: this.options.getAuthoritativeRevision(),
        pendingMutationSummary: this.options.buildPendingMutationSummary(),
        localPersistenceMode: this.localPersistenceMode,
      })
    }
    return this.storeRuntimeState({
      ...buildWorkerRuntimeStateFromEngine(requireEngine()),
      localPersistenceMode: this.localPersistenceMode,
    })
  }

  storeRuntimeState(state: WorkbookWorkerStateSnapshot): WorkbookWorkerStateSnapshot {
    this.runtimeStateCache = cloneWorkerRuntimeState({
      ...state,
      authoritativeRevision: this.options.getAuthoritativeRevision(),
      pendingMutationSummary: this.options.buildPendingMutationSummary(),
    })
    return withExternalSyncState(this.runtimeStateCache, this.externalSyncState)
  }

  updateRuntimeStateFromEngine(engine: WorkerRuntimeStateEngine): WorkbookWorkerStateSnapshot {
    return this.storeRuntimeState(buildWorkerRuntimeStateFromEngine(engine))
  }

  getCurrentMetrics(): RecalcMetrics {
    const engine = this.options.getEngine()
    if (engine) {
      return cloneRuntimeMetrics(engine.getLastMetrics())
    }
    return cloneRuntimeMetrics(this.runtimeStateCache?.metrics ?? EMPTY_RUNTIME_METRICS)
  }
}
