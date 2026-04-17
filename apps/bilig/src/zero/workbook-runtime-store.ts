import { SpreadsheetEngine, type EngineReplicaSnapshot } from '@bilig/core'
import { isWorkbookSnapshot, type WorkbookSnapshot } from '@bilig/protocol'
import { applyWorkbookEvent } from '@bilig/zero-sync'
import { parseCheckpointPayload, parseCheckpointReplicaState, parseInteger } from './store-support.js'
import { loadWorkbookEventRecordsAfter, type Queryable, type WorkbookRuntimeMetadata, type WorkbookRuntimeState } from './store.js'

interface WorkbookCheckpointRecord {
  revision: number
  checkpointPayload: WorkbookSnapshot
  replicaState: EngineReplicaSnapshot | null
}

async function loadLatestWorkbookCheckpoint(db: Queryable, documentId: string): Promise<WorkbookCheckpointRecord | null> {
  const result = await db.query<{
    revision: number | string | null
    payload: unknown
    replica_snapshot: unknown
  }>(
    `
      SELECT revision, payload, replica_snapshot
      FROM workbook_snapshot
      WHERE workbook_id = $1
        AND format = 'json-v1'
      ORDER BY revision DESC
      LIMIT 1
    `,
    [documentId],
  )
  const row = result.rows[0]
  if (!row || !isWorkbookSnapshot(row.payload)) {
    return null
  }
  return {
    revision: parseInteger(row.revision),
    checkpointPayload: row.payload,
    replicaState: parseCheckpointReplicaState(row.replica_snapshot),
  }
}

export async function loadWorkbookState(db: Queryable, documentId: string): Promise<WorkbookRuntimeState> {
  const result = await db.query<{
    snapshot: unknown
    replica_snapshot: unknown
    head_revision: number | string | null
    calculated_revision: number | string | null
    owner_user_id: string | null
  }>(`SELECT snapshot, replica_snapshot, head_revision, calculated_revision, owner_user_id FROM workbooks WHERE id = $1 LIMIT 1`, [
    documentId,
  ])
  const row = result.rows[0]
  const headRevision = parseInteger(row?.head_revision)
  const calculatedRevision = parseInteger(row?.calculated_revision)
  const ownerUserId = row?.owner_user_id ?? 'system'
  const inlineCheckpointPayload = isWorkbookSnapshot(row?.snapshot) ? row.snapshot : null
  const inlineReplicaState = parseCheckpointReplicaState(row?.replica_snapshot)

  if (inlineCheckpointPayload) {
    return {
      snapshot: inlineCheckpointPayload,
      replicaSnapshot: inlineReplicaState,
      headRevision,
      calculatedRevision,
      ownerUserId,
    }
  }

  const checkpoint = await loadLatestWorkbookCheckpoint(db, documentId)
  const baseRevision = checkpoint?.revision ?? 0
  const baseCheckpointPayload = parseCheckpointPayload(checkpoint?.checkpointPayload, documentId)
  const baseReplicaState = parseCheckpointReplicaState(checkpoint?.replicaState)

  if (headRevision <= baseRevision) {
    return {
      snapshot: baseCheckpointPayload,
      replicaSnapshot: baseReplicaState,
      headRevision,
      calculatedRevision,
      ownerUserId,
    }
  }

  const engine = new SpreadsheetEngine({
    workbookName: documentId,
    replicaId: `checkpoint-replay:${documentId}:${headRevision}`,
  })
  await engine.ready()
  engine.importSnapshot(baseCheckpointPayload)
  if (baseReplicaState) {
    engine.importReplicaSnapshot(baseReplicaState)
  }
  const events = await loadWorkbookEventRecordsAfter(db, documentId, baseRevision)
  for (const event of events) {
    applyWorkbookEvent(engine, event.payload)
  }

  return {
    snapshot: engine.exportSnapshot(),
    replicaSnapshot: null,
    headRevision,
    calculatedRevision,
    ownerUserId,
  }
}

export async function loadWorkbookRuntimeMetadata(db: Queryable, documentId: string): Promise<WorkbookRuntimeMetadata> {
  const result = await db.query<{
    head_revision: number | string | null
    calculated_revision: number | string | null
    owner_user_id: string | null
  }>(`SELECT head_revision, calculated_revision, owner_user_id FROM workbooks WHERE id = $1 LIMIT 1`, [documentId])
  const row = result.rows[0]
  return {
    headRevision: parseInteger(row?.head_revision),
    calculatedRevision: parseInteger(row?.calculated_revision),
    ownerUserId: row?.owner_user_id ?? 'system',
  }
}

export async function acquireWorkbookMutationLock(db: Queryable, documentId: string): Promise<void> {
  await db.query(`SELECT pg_advisory_xact_lock(hashtext($1))`, [documentId])
}
