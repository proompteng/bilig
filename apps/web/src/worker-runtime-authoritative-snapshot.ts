import type { EngineReplicaSnapshot, SpreadsheetEngine } from '@bilig/core'
import type { WorkbookSnapshot } from '@bilig/protocol'
import type { PendingWorkbookMutation } from './workbook-sync.js'
import { createProjectionEngineFromState, createWorkbookEngineFromState } from './worker-runtime-engine-state.js'
import type { InstallAuthoritativeSnapshotInput } from './worker-runtime-state.js'
import type { ProjectionOverlayScope } from './worker-local-overlay.js'
import type { WorkerEngine } from './worker-runtime-support.js'

export interface PreparedAuthoritativeSnapshotProjection {
  readonly projectionEngine: SpreadsheetEngine & WorkerEngine
  readonly projectionOverlayScope: ProjectionOverlayScope | null
  readonly authoritativeEngine: SpreadsheetEngine | null
  readonly authoritativeSnapshot: WorkbookSnapshot
  readonly authoritativeReplica: EngineReplicaSnapshot | null
  readonly shouldMarkPendingMutationsRebased: boolean
}

export async function prepareAuthoritativeSnapshotProjection(input: {
  readonly documentId: string
  readonly replicaId: string
  readonly snapshot: WorkbookSnapshot
  readonly mode: InstallAuthoritativeSnapshotInput['mode']
  readonly pendingMutations: readonly PendingWorkbookMutation[]
}): Promise<PreparedAuthoritativeSnapshotProjection> {
  if (input.mode === 'bootstrap' && input.pendingMutations.length === 0) {
    const { engine, overlayScope } = await createProjectionEngineFromState({
      workbookName: input.documentId,
      replicaId: input.replicaId,
      snapshot: input.snapshot,
      replica: null,
      pendingMutations: [],
    })
    return {
      projectionEngine: engine,
      projectionOverlayScope: overlayScope,
      authoritativeEngine: null,
      authoritativeSnapshot: input.snapshot,
      authoritativeReplica: engine.exportReplicaSnapshot(),
      shouldMarkPendingMutationsRebased: false,
    }
  }

  const authoritativeEngine = await createWorkbookEngineFromState({
    workbookName: input.documentId,
    replicaId: input.replicaId,
    snapshot: input.snapshot,
    replica: null,
  })
  const authoritativeReplica = authoritativeEngine.exportReplicaSnapshot()
  const { engine, overlayScope } = await createProjectionEngineFromState({
    workbookName: input.documentId,
    replicaId: input.replicaId,
    snapshot: input.snapshot,
    replica: authoritativeReplica,
    pendingMutations: input.pendingMutations,
  })
  return {
    projectionEngine: engine,
    projectionOverlayScope: overlayScope,
    authoritativeEngine,
    authoritativeSnapshot: input.snapshot,
    authoritativeReplica,
    shouldMarkPendingMutationsRebased: input.mode === 'reconcile' && input.pendingMutations.length > 0,
  }
}
