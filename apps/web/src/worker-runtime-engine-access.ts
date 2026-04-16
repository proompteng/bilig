import type { EngineReplicaSnapshot } from '@bilig/core'
import { isEngineReplicaSnapshot, SpreadsheetEngine } from '@bilig/core'
import { parseCellAddress } from '@bilig/formula'
import { isWorkbookSnapshot, type CellSnapshot, type WorkbookSnapshot } from '@bilig/protocol'
import type { WorkbookLocalStore, WorkbookStoredState } from '@bilig/storage-browser'
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

interface LocalStoreStateReader {
  loadState(): Promise<WorkbookStoredState | null>
}

interface ExportableAuthoritativeEngine {
  exportSnapshot(): WorkbookSnapshot
  exportReplicaSnapshot(): EngineReplicaSnapshot
}

interface ViewportTileReader {
  readViewport(input: {
    localStore: Pick<WorkbookLocalStore, 'readViewportProjection'>
    sheetName: string
    viewport: {
      rowStart: number
      rowEnd: number
      colStart: number
      colEnd: number
    }
  }): {
    readonly cells: readonly { readonly snapshot: CellSnapshot }[]
  } | null
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
  authoritativeStateSource: 'none' | 'memory' | 'localStore'
  localStore: LocalStoreStateReader | null
  snapshotCaches: AuthoritativeStateSnapshotCache
  authoritativeEngine: ExportableAuthoritativeEngine | null
  installRestoredAuthoritativeState: (snapshot: WorkbookSnapshot | null, replica: EngineReplicaSnapshot | null) => void
}): Promise<WorkerRuntimeAuthoritativeStateInput> {
  if (args.authoritativeStateSource === 'localStore') {
    const restoredState = args.localStore ? await args.localStore.loadState() : null
    const restoredSnapshot = isWorkbookSnapshot(restoredState?.snapshot) ? restoredState.snapshot : null
    const restoredReplica = isEngineReplicaSnapshot(restoredState?.replica) ? restoredState.replica : null
    args.installRestoredAuthoritativeState(restoredSnapshot, restoredReplica)
  }
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

export function readProjectedCellFromLocalStore(args: {
  canReadLocalProjectionForViewport: boolean
  localStore: Pick<WorkbookLocalStore, 'readViewportProjection'> | null
  viewportTileStore: ViewportTileReader
  sheetName: string
  address: string
}): CellSnapshot | null {
  if (!args.canReadLocalProjectionForViewport || !args.localStore) {
    return null
  }
  const parsed = parseCellAddress(args.address, args.sheetName)
  const localBase = args.viewportTileStore.readViewport({
    localStore: args.localStore,
    sheetName: args.sheetName,
    viewport: {
      rowStart: parsed.row,
      rowEnd: parsed.row,
      colStart: parsed.col,
      colEnd: parsed.col,
    },
  })
  const cell = localBase?.cells.find((entry) => entry.snapshot.address === args.address)
  return cell ? structuredClone(cell.snapshot) : null
}
