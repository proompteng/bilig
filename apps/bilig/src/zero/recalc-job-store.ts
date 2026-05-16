import type { EngineReplicaSnapshot } from '@bilig/core'
import type { WorkbookSnapshot } from '@bilig/protocol'
import type { DirtyRegion } from '@bilig/zero-sync'
import type { CellEvalRow } from './projection.js'
import { isDirtyRegion, parseNonNegativeInteger, parsePositiveInteger } from './store-support.js'
import { shouldPersistWorkbookCheckpointRevision, type Queryable, type QueryResultRow } from './store.js'
import { persistCellEvalDiff, persistCellEvalIncremental, persistWorkbookCheckpoint } from './workbook-calculation-store.js'

const RECALC_LEASE_MS = 30_000
const MAX_RECALC_ATTEMPTS = 3

export interface RecalcJobLease {
  id: string
  workbookId: string
  fromRevision: number
  toRevision: number
  dirtyRegions: DirtyRegion[] | null
  attempts: number
}

interface RecalcJobLeaseRow extends QueryResultRow {
  readonly id?: unknown
  readonly workbook_id?: unknown
  readonly from_revision?: unknown
  readonly to_revision?: unknown
  readonly dirty_regions_json?: unknown
  readonly attempts?: unknown
}

function normalizeRecalcJobLease(row: RecalcJobLeaseRow): RecalcJobLease | null {
  const fromRevision = parseNonNegativeInteger(row.from_revision)
  const toRevision = parsePositiveInteger(row.to_revision)
  const attempts = parsePositiveInteger(row.attempts)
  if (
    typeof row.id !== 'string' ||
    row.id.length === 0 ||
    typeof row.workbook_id !== 'string' ||
    row.workbook_id.length === 0 ||
    fromRevision === null ||
    toRevision === null ||
    toRevision <= fromRevision ||
    attempts === null
  ) {
    return null
  }

  const dirtyRegions = Array.isArray(row.dirty_regions_json) ? row.dirty_regions_json.filter(isDirtyRegion) : null
  return {
    id: row.id,
    workbookId: row.workbook_id,
    fromRevision,
    toRevision,
    dirtyRegions: dirtyRegions && dirtyRegions.length > 0 ? dirtyRegions : null,
    attempts,
  }
}

async function markMalformedRecalcJobFailed(db: Queryable, row: RecalcJobLeaseRow, reason: string): Promise<void> {
  if (typeof row.id !== 'string' || row.id.length === 0) {
    return
  }
  await db.query(
    `
      UPDATE recalc_job
      SET status = 'failed',
          lease_until = NULL,
          lease_owner = NULL,
          last_error = $2,
          updated_at = NOW()
      WHERE id = $1
    `,
    [row.id, reason],
  )
}

export async function leaseNextRecalcJob(db: Queryable, workerId: string): Promise<RecalcJobLease | null> {
  const result = await db.query<RecalcJobLeaseRow>(
    `
      WITH candidate AS (
        SELECT id
        FROM recalc_job
        WHERE status = 'pending'
           OR (status = 'running' AND lease_until IS NOT NULL AND lease_until < NOW())
        ORDER BY created_at ASC
        LIMIT 1
        FOR UPDATE SKIP LOCKED
      )
      UPDATE recalc_job
      SET status = 'running',
          attempts = attempts + 1,
          lease_owner = $1,
          lease_until = NOW() + ($2 * INTERVAL '1 millisecond'),
          updated_at = NOW()
      WHERE id IN (SELECT id FROM candidate)
      RETURNING id, workbook_id, from_revision, to_revision, dirty_regions_json, attempts
    `,
    [workerId, RECALC_LEASE_MS],
  )
  const row = result.rows[0]
  if (!row) {
    return null
  }
  const lease = normalizeRecalcJobLease(row)
  if (!lease) {
    await markMalformedRecalcJobFailed(db, row, 'Malformed recalc job lease row')
    return null
  }
  return lease
}

export async function markRecalcJobCompleted(
  db: Queryable,
  lease: RecalcJobLease,
  nextRows: readonly CellEvalRow[],
  snapshot: WorkbookSnapshot | null,
  replicaSnapshot: EngineReplicaSnapshot | null,
  isIncremental = false,
): Promise<boolean> {
  const revisionResult = await db.query<{ head_revision: number | string | null }>(
    `SELECT head_revision FROM workbooks WHERE id = $1 LIMIT 1`,
    [lease.workbookId],
  )
  const headRevision = parsePositiveInteger(revisionResult.rows[0]?.head_revision)
  if (headRevision === null) {
    throw new Error(`Invalid workbook head revision while completing recalc job ${lease.id}`)
  }
  if (headRevision !== lease.toRevision) {
    await markRecalcJobSuperseded(db, lease)
    return false
  }

  if (isIncremental) {
    await persistCellEvalIncremental(db, lease.workbookId, nextRows)
  } else {
    await persistCellEvalDiff(db, lease.workbookId, nextRows)
  }
  await db.query(
    `
      UPDATE workbooks
      SET calculated_revision = $2
      WHERE id = $1 AND head_revision = $2
    `,
    [lease.workbookId, lease.toRevision],
  )
  await db.query(
    `
      UPDATE recalc_job
      SET status = 'completed',
          lease_until = NULL,
          lease_owner = NULL,
          updated_at = NOW()
      WHERE id = $1
    `,
    [lease.id],
  )

  if (shouldPersistWorkbookCheckpointRevision(lease.toRevision) && snapshot) {
    await persistWorkbookCheckpoint(db, lease.workbookId, lease.toRevision, snapshot, replicaSnapshot)
  }
  return true
}

export async function markRecalcJobSuperseded(db: Queryable, lease: RecalcJobLease): Promise<void> {
  await db.query(
    `
      UPDATE recalc_job
      SET status = 'superseded',
          lease_until = NULL,
          lease_owner = NULL,
          updated_at = NOW()
      WHERE id = $1
    `,
    [lease.id],
  )
}

export async function markRecalcJobFailed(db: Queryable, lease: RecalcJobLease, error: unknown): Promise<void> {
  const message = error instanceof Error ? (error.stack ?? error.message) : String(error)
  const exhausted = lease.attempts >= MAX_RECALC_ATTEMPTS
  await db.query(
    `
      UPDATE recalc_job
      SET status = $2,
          lease_until = NULL,
          lease_owner = NULL,
          last_error = $3,
          updated_at = NOW()
      WHERE id = $1
    `,
    [lease.id, exhausted ? 'failed' : 'pending', message],
  )
}
