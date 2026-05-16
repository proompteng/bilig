import type { RecalcMetrics, SyncState, WorkbookDefinedNameSnapshot, WorkbookSnapshot } from '@bilig/protocol'
import type { WorkbookBenchmarkCorpusId, WorkbookBenchmarkCorpusViewport } from './worker-runtime-benchmark-corpus.js'
import type { WorkerRuntimeLocalHistoryState } from './worker-runtime-local-history.js'
import type { WorkbookRuntimeSheetSnapshot } from './worker-runtime-state.js'

export const DEFERRED_PROJECTION_ENGINE_MIN_CELL_COUNT = 100_000

export interface WorkbookWorkerBootstrapOptions {
  documentId: string
  replicaId: string
  persistState: boolean
}

export interface WorkbookWorkerStateSnapshot {
  workbookName: string
  sheets?: WorkbookRuntimeSheetSnapshot[] | undefined
  sheetNames: string[]
  definedNames: WorkbookDefinedNameSnapshot[]
  metrics: RecalcMetrics
  syncState: SyncState
  localHistoryState: WorkerRuntimeLocalHistoryState
  authoritativeRevision?: number | undefined
  pendingMutationSummary?: WorkbookPendingMutationSummarySnapshot
  localPersistenceMode?: 'persistent' | 'ephemeral' | 'follower'
}

export interface WorkbookFailedPendingMutationSnapshot {
  readonly id: string
  readonly method: string
  readonly failureMessage: string
  readonly attemptCount: number
}

export interface WorkbookPendingMutationSummarySnapshot {
  readonly activeCount: number
  readonly failedCount: number
  readonly firstFailed: WorkbookFailedPendingMutationSnapshot | null
}

export interface WorkbookWorkerBootstrapResult {
  runtimeState: WorkbookWorkerStateSnapshot
  restoredFromPersistence: boolean
  requiresAuthoritativeHydrate: boolean
  localPersistenceMode?: 'persistent' | 'ephemeral' | 'follower'
}

export interface InstallAuthoritativeSnapshotInput {
  readonly snapshot: WorkbookSnapshot
  readonly authoritativeRevision: number
  readonly mode: 'bootstrap' | 'reconcile'
}

export interface InstallBenchmarkCorpusResult {
  readonly id: WorkbookBenchmarkCorpusId
  readonly materializedCellCount: number
  readonly primaryViewport: WorkbookBenchmarkCorpusViewport
}
