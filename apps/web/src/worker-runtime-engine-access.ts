import type { EngineReplicaSnapshot, SpreadsheetEngine } from '@bilig/core'
import type { WorkbookSnapshot } from '@bilig/protocol'
import { createProjectionEngineFromState, createWorkbookEngineFromState } from './worker-runtime-engine-state.js'
import type { ProjectionOverlayScope } from './worker-local-overlay.js'
import type { PendingWorkbookMutation } from './workbook-sync.js'

export interface WorkerRuntimeAuthoritativeStateInput {
  snapshot: WorkbookSnapshot | null
  replica: EngineReplicaSnapshot | null
}

interface AuthoritativeStateSnapshotCache {
  installAuthoritativeState(snapshot: WorkbookSnapshot | null, replica: EngineReplicaSnapshot | null): void
  resolveAuthoritativeState(input: {
    exportSnapshot: (() => WorkbookSnapshot) | null
    exportReplica: (() => EngineReplicaSnapshot) | null
  }): WorkerRuntimeAuthoritativeStateInput
  storeAuthoritativeSnapshot(snapshot: WorkbookSnapshot): WorkbookSnapshot
  storeAuthoritativeReplica(replica: EngineReplicaSnapshot): EngineReplicaSnapshot
}

interface ExportableAuthoritativeEngine {
  exportSnapshot(): WorkbookSnapshot
  exportReplicaSnapshot(): EngineReplicaSnapshot
}

export function installAuthoritativeEngineState(
  snapshotCaches: AuthoritativeStateSnapshotCache,
  engine: ExportableAuthoritativeEngine,
  snapshot: WorkbookSnapshot | null,
  replica: EngineReplicaSnapshot | null,
): void {
  snapshotCaches.installAuthoritativeState(snapshot ?? engine.exportSnapshot(), replica ?? engine.exportReplicaSnapshot())
}

export function installRestoredAuthoritativeState(
  snapshotCaches: AuthoritativeStateSnapshotCache,
  snapshot: WorkbookSnapshot | null,
  replica: EngineReplicaSnapshot | null,
): void {
  snapshotCaches.installAuthoritativeState(snapshot, replica)
}

export async function resolveAuthoritativeStateInput(args: {
  authoritativeStateSource: 'none' | 'memory'
  snapshotCaches: AuthoritativeStateSnapshotCache
  authoritativeEngine: ExportableAuthoritativeEngine | null
}): Promise<WorkerRuntimeAuthoritativeStateInput> {
  return args.snapshotCaches.resolveAuthoritativeState({
    exportSnapshot: args.authoritativeEngine ? () => args.authoritativeEngine!.exportSnapshot() : null,
    exportReplica: args.authoritativeEngine ? () => args.authoritativeEngine!.exportReplicaSnapshot() : null,
  })
}

export async function ensureAuthoritativeEngine(args: {
  authoritativeEngine: SpreadsheetEngine | null
  documentId: string
  replicaId: string
  snapshotCaches: AuthoritativeStateSnapshotCache
  resolveAuthoritativeStateInput: () => Promise<WorkerRuntimeAuthoritativeStateInput>
}): Promise<SpreadsheetEngine> {
  if (args.authoritativeEngine) {
    return args.authoritativeEngine
  }
  const authoritativeState = await args.resolveAuthoritativeStateInput()
  const engine = await createWorkbookEngineFromState({
    workbookName: args.documentId,
    replicaId: args.replicaId,
    snapshot: authoritativeState.snapshot,
    replica: authoritativeState.replica,
  })
  if (!authoritativeState.snapshot) {
    args.snapshotCaches.storeAuthoritativeSnapshot(engine.exportSnapshot())
  }
  if (!authoritativeState.replica) {
    args.snapshotCaches.storeAuthoritativeReplica(engine.exportReplicaSnapshot())
  }
  return engine
}

export async function rebuildProjectionEngine(args: {
  documentId: string
  replicaId: string
  pendingMutations: readonly PendingWorkbookMutation[]
  resolveAuthoritativeStateInput: () => Promise<WorkerRuntimeAuthoritativeStateInput>
}): Promise<{
  engine: SpreadsheetEngine
  overlayScope: ProjectionOverlayScope | null
}> {
  const authoritativeState = await args.resolveAuthoritativeStateInput()
  return await createProjectionEngineFromState({
    workbookName: args.documentId,
    replicaId: args.replicaId,
    snapshot: authoritativeState.snapshot,
    replica: authoritativeState.replica,
    pendingMutations: args.pendingMutations,
  })
}
