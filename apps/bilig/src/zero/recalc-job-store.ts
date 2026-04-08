import type { EngineReplicaSnapshot } from "@bilig/core";
import type { WorkbookSnapshot } from "@bilig/protocol";
import type { CellEvalRow } from "./projection.js";
import { isDirtyRegion, parseInteger } from "./store-support.js";
import {
  persistCellEvalDiff,
  persistCellEvalIncremental,
  persistWorkbookCheckpoint,
  shouldPersistWorkbookCheckpointRevision,
  type Queryable,
} from "./store.js";

const RECALC_LEASE_MS = 30_000;
const MAX_RECALC_ATTEMPTS = 3;

export interface RecalcJobLease {
  id: string;
  workbookId: string;
  fromRevision: number;
  toRevision: number;
  dirtyRegions: import("@bilig/zero-sync").DirtyRegion[] | null;
  attempts: number;
}

export async function leaseNextRecalcJob(
  db: Queryable,
  workerId: string,
): Promise<RecalcJobLease | null> {
  const result = await db.query<{
    id: string;
    workbook_id: string;
    from_revision: number | string | null;
    to_revision: number | string | null;
    dirty_regions_json: unknown;
    attempts: number | string | null;
  }>(
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
  );
  const row = result.rows[0];
  if (!row) {
    return null;
  }
  const dirtyRegions = Array.isArray(row.dirty_regions_json)
    ? row.dirty_regions_json.filter(isDirtyRegion)
    : null;
  return {
    id: row.id,
    workbookId: row.workbook_id,
    fromRevision: parseInteger(row.from_revision),
    toRevision: parseInteger(row.to_revision),
    dirtyRegions: dirtyRegions && dirtyRegions.length > 0 ? dirtyRegions : null,
    attempts: parseInteger(row.attempts),
  };
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
  );
  if (parseInteger(revisionResult.rows[0]?.head_revision) !== lease.toRevision) {
    await markRecalcJobSuperseded(db, lease);
    return false;
  }

  if (isIncremental) {
    await persistCellEvalIncremental(db, lease.workbookId, nextRows);
  } else {
    await persistCellEvalDiff(db, lease.workbookId, nextRows);
  }
  await db.query(
    `
      UPDATE workbooks
      SET calculated_revision = $2
      WHERE id = $1 AND head_revision = $2
    `,
    [lease.workbookId, lease.toRevision],
  );
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
  );

  if (shouldPersistWorkbookCheckpointRevision(lease.toRevision) && snapshot) {
    await persistWorkbookCheckpoint(
      db,
      lease.workbookId,
      lease.toRevision,
      snapshot,
      replicaSnapshot,
    );
  }
  return true;
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
  );
}

export async function markRecalcJobFailed(
  db: Queryable,
  lease: RecalcJobLease,
  error: unknown,
): Promise<void> {
  const message = error instanceof Error ? (error.stack ?? error.message) : String(error);
  const exhausted = lease.attempts >= MAX_RECALC_ATTEMPTS;
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
    [lease.id, exhausted ? "failed" : "pending", message],
  );
}
