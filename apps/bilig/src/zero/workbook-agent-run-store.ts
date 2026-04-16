import {
  decodeWorkbookAgentPreviewSummary,
  isWorkbookAgentContextRef,
  isWorkbookAgentCommand,
  type WorkbookAgentCommand,
  type WorkbookAgentExecutionRecord,
} from '@bilig/agent-api'
import type { QueryResultRow, Queryable } from './store.js'

interface WorkbookAgentRunRow extends QueryResultRow {
  readonly id?: unknown
  readonly bundleId?: unknown
  readonly workbookId?: unknown
  readonly threadId?: unknown
  readonly turnId?: unknown
  readonly actorUserId?: unknown
  readonly goalText?: unknown
  readonly planText?: unknown
  readonly summary?: unknown
  readonly scope?: unknown
  readonly riskClass?: unknown
  readonly acceptedScope?: unknown
  readonly appliedBy?: unknown
  readonly baseRevision?: unknown
  readonly appliedRevision?: unknown
  readonly createdAtUnixMs?: unknown
  readonly appliedAtUnixMs?: unknown
  readonly contextJson?: unknown
  readonly commandsJson?: unknown
  readonly previewJson?: unknown
}

function parseNumericValue(value: unknown): number | null {
  if (typeof value === 'number' && Number.isFinite(value)) {
    return Math.trunc(value)
  }
  if (typeof value === 'string' && value.length > 0) {
    const parsed = Number.parseInt(value, 10)
    return Number.isFinite(parsed) ? parsed : null
  }
  return null
}

function parseCommands(value: unknown): WorkbookAgentCommand[] | null {
  if (!Array.isArray(value)) {
    return null
  }
  const commands = value.flatMap((entry) => (isWorkbookAgentCommand(entry) ? [entry] : []))
  return commands.length === value.length ? commands : null
}

function normalizeExecutionRecord(row: WorkbookAgentRunRow): WorkbookAgentExecutionRecord | null {
  const baseRevision = parseNumericValue(row.baseRevision)
  const appliedRevision = parseNumericValue(row.appliedRevision)
  const createdAtUnixMs = parseNumericValue(row.createdAtUnixMs)
  const appliedAtUnixMs = parseNumericValue(row.appliedAtUnixMs)
  const commands = parseCommands(row.commandsJson)
  if (
    typeof row.id !== 'string' ||
    typeof row.bundleId !== 'string' ||
    typeof row.workbookId !== 'string' ||
    typeof row.threadId !== 'string' ||
    typeof row.turnId !== 'string' ||
    typeof row.actorUserId !== 'string' ||
    typeof row.goalText !== 'string' ||
    typeof row.summary !== 'string' ||
    (row.scope !== 'selection' && row.scope !== 'sheet' && row.scope !== 'workbook') ||
    (row.riskClass !== 'low' && row.riskClass !== 'medium' && row.riskClass !== 'high') ||
    (row.acceptedScope !== 'full' && row.acceptedScope !== 'partial') ||
    (row.appliedBy !== 'user' && row.appliedBy !== 'auto') ||
    baseRevision === null ||
    appliedRevision === null ||
    createdAtUnixMs === null ||
    appliedAtUnixMs === null ||
    commands === null
  ) {
    return null
  }
  const context =
    row.contextJson === null || row.contextJson === undefined ? null : isWorkbookAgentContextRef(row.contextJson) ? row.contextJson : null
  const preview = row.previewJson === null || row.previewJson === undefined ? null : decodeWorkbookAgentPreviewSummary(row.previewJson)
  return {
    id: row.id,
    bundleId: row.bundleId,
    documentId: row.workbookId,
    threadId: row.threadId,
    turnId: row.turnId,
    actorUserId: row.actorUserId,
    goalText: row.goalText,
    planText: typeof row.planText === 'string' ? row.planText : null,
    summary: row.summary,
    scope: row.scope,
    riskClass: row.riskClass,
    acceptedScope: row.acceptedScope,
    appliedBy: row.appliedBy,
    baseRevision,
    appliedRevision,
    createdAtUnixMs,
    appliedAtUnixMs,
    context,
    commands,
    preview,
  }
}

export async function ensureWorkbookAgentRunSchema(db: Queryable): Promise<void> {
  await db.query(`
    CREATE TABLE IF NOT EXISTS workbook_agent_run (
      id TEXT PRIMARY KEY,
      bundle_id TEXT NOT NULL,
      workbook_id TEXT NOT NULL,
      thread_id TEXT NOT NULL,
      turn_id TEXT NOT NULL,
      actor_user_id TEXT NOT NULL,
      goal_text TEXT NOT NULL,
      plan_text TEXT,
      summary TEXT NOT NULL,
      scope TEXT NOT NULL,
      risk_class TEXT NOT NULL,
      accepted_scope TEXT NOT NULL DEFAULT 'full',
      applied_by TEXT NOT NULL DEFAULT 'user',
      base_revision BIGINT NOT NULL,
      applied_revision BIGINT NOT NULL,
      created_at_unix_ms BIGINT NOT NULL,
      applied_at_unix_ms BIGINT NOT NULL,
      context_json JSONB,
      commands_json JSONB NOT NULL,
      preview_json JSONB
    )
  `)
  await db.query(`ALTER TABLE workbook_agent_run ADD COLUMN IF NOT EXISTS bundle_id TEXT;`)
  await db.query(`
    ALTER TABLE workbook_agent_run
      ADD COLUMN IF NOT EXISTS accepted_scope TEXT NOT NULL DEFAULT 'full';
  `)
  await db.query(`
    ALTER TABLE workbook_agent_run
      ADD COLUMN IF NOT EXISTS applied_by TEXT NOT NULL DEFAULT 'user';
  `)
  await db.query(`
    CREATE INDEX IF NOT EXISTS workbook_agent_run_workbook_actor_applied_idx
      ON workbook_agent_run (workbook_id, actor_user_id, applied_at_unix_ms DESC)
  `)
}

export async function appendWorkbookAgentRun(db: Queryable, record: WorkbookAgentExecutionRecord): Promise<void> {
  await db.query(
    `
      INSERT INTO workbook_agent_run (
        id,
        bundle_id,
        workbook_id,
        thread_id,
        turn_id,
        actor_user_id,
        goal_text,
        plan_text,
        summary,
        scope,
        risk_class,
        accepted_scope,
        applied_by,
        base_revision,
        applied_revision,
        created_at_unix_ms,
        applied_at_unix_ms,
        context_json,
        commands_json,
        preview_json
      )
      VALUES (
        $1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11, $12, $13, $14, $15, $16, $17, $18::jsonb, $19::jsonb, $20::jsonb
      )
      ON CONFLICT (id)
      DO UPDATE SET
        bundle_id = EXCLUDED.bundle_id,
        workbook_id = EXCLUDED.workbook_id,
        thread_id = EXCLUDED.thread_id,
        turn_id = EXCLUDED.turn_id,
        actor_user_id = EXCLUDED.actor_user_id,
        goal_text = EXCLUDED.goal_text,
        plan_text = EXCLUDED.plan_text,
        summary = EXCLUDED.summary,
        scope = EXCLUDED.scope,
        risk_class = EXCLUDED.risk_class,
        accepted_scope = EXCLUDED.accepted_scope,
        applied_by = EXCLUDED.applied_by,
        base_revision = EXCLUDED.base_revision,
        applied_revision = EXCLUDED.applied_revision,
        created_at_unix_ms = EXCLUDED.created_at_unix_ms,
        applied_at_unix_ms = EXCLUDED.applied_at_unix_ms,
        context_json = EXCLUDED.context_json,
        commands_json = EXCLUDED.commands_json,
        preview_json = EXCLUDED.preview_json
    `,
    [
      record.id,
      record.bundleId,
      record.documentId,
      record.threadId,
      record.turnId,
      record.actorUserId,
      record.goalText,
      record.planText,
      record.summary,
      record.scope,
      record.riskClass,
      record.acceptedScope,
      record.appliedBy,
      record.baseRevision,
      record.appliedRevision,
      record.createdAtUnixMs,
      record.appliedAtUnixMs,
      JSON.stringify(record.context),
      JSON.stringify(record.commands),
      JSON.stringify(record.preview),
    ],
  )
}

export async function listWorkbookAgentRuns(
  db: Queryable,
  input: {
    documentId: string
    actorUserId: string
    limit?: number
  },
): Promise<WorkbookAgentExecutionRecord[]> {
  const result = await db.query<WorkbookAgentRunRow>(
    `
      SELECT
        id AS "id",
        bundle_id AS "bundleId",
        workbook_id AS "workbookId",
        thread_id AS "threadId",
        turn_id AS "turnId",
        actor_user_id AS "actorUserId",
        goal_text AS "goalText",
        plan_text AS "planText",
        summary AS "summary",
        scope AS "scope",
        risk_class AS "riskClass",
        accepted_scope AS "acceptedScope",
        applied_by AS "appliedBy",
        base_revision AS "baseRevision",
        applied_revision AS "appliedRevision",
        created_at_unix_ms AS "createdAtUnixMs",
        applied_at_unix_ms AS "appliedAtUnixMs",
        context_json AS "contextJson",
        commands_json AS "commandsJson",
        preview_json AS "previewJson"
      FROM workbook_agent_run
      WHERE workbook_id = $1 AND actor_user_id = $2
      ORDER BY applied_at_unix_ms DESC
      LIMIT $3
    `,
    [input.documentId, input.actorUserId, input.limit ?? 20],
  )
  return result.rows.flatMap((row) => {
    const record = normalizeExecutionRecord(row)
    return record ? [record] : []
  })
}

export async function listWorkbookAgentThreadRuns(
  db: Queryable,
  input: {
    documentId: string
    actorUserId: string
    threadId: string
    limit?: number
  },
): Promise<WorkbookAgentExecutionRecord[]> {
  const result = await db.query<WorkbookAgentRunRow>(
    `
      SELECT
        run.id AS "id",
        run.bundle_id AS "bundleId",
        run.workbook_id AS "workbookId",
        run.thread_id AS "threadId",
        run.turn_id AS "turnId",
        run.actor_user_id AS "actorUserId",
        run.goal_text AS "goalText",
        run.plan_text AS "planText",
        run.summary AS "summary",
        run.scope AS "scope",
        run.risk_class AS "riskClass",
        run.accepted_scope AS "acceptedScope",
        run.applied_by AS "appliedBy",
        run.base_revision AS "baseRevision",
        run.applied_revision AS "appliedRevision",
        run.created_at_unix_ms AS "createdAtUnixMs",
        run.applied_at_unix_ms AS "appliedAtUnixMs",
        run.context_json AS "contextJson",
        run.commands_json AS "commandsJson",
        run.preview_json AS "previewJson"
      FROM workbook_agent_run AS run
      WHERE run.workbook_id = $1
        AND run.thread_id = $2
        AND (
          run.actor_user_id = $3
          OR EXISTS (
            SELECT 1
            FROM workbook_chat_thread AS thread
            WHERE thread.workbook_id = run.workbook_id
              AND thread.thread_id = run.thread_id
              AND thread.scope = 'shared'
          )
        )
      ORDER BY run.applied_at_unix_ms DESC
      LIMIT $4
    `,
    [input.documentId, input.threadId, input.actorUserId, input.limit ?? 20],
  )
  return result.rows.flatMap((row) => {
    const record = normalizeExecutionRecord(row)
    return record ? [record] : []
  })
}
