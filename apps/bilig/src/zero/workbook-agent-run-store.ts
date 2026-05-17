import {
  decodeWorkbookAgentPreviewSummary,
  isWorkbookAgentContextRef,
  isWorkbookAgentCommand,
  type WorkbookAgentCommand,
  type WorkbookAgentExecutionRecord,
} from '@bilig/agent-api'
import { queries } from '@bilig/zero-sync'
import type { Row } from '@rocicorp/zero'
import { addDefaultedColumnIfMissing, enforceDefaultedNotNullColumn } from './schema-upgrade.js'
import type { QueryResultRow, Queryable, ZeroQueryRunner } from './store.js'
import { parseNonNegativeInteger } from './store-support.js'
import { ensureZeroSchemaTable } from './zero-schema-ddl.js'

type ZeroWorkbookAgentRunRow = Row['workbook_agent_run']

export interface WorkbookAgentRunStoreConnection extends Queryable {
  listWorkbookAgentRunRows(input: {
    readonly documentId: string
    readonly actorUserId: string
    readonly threadId?: string
  }): Promise<readonly ZeroWorkbookAgentRunRow[]>
}

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

export function createWorkbookAgentRunStoreConnection(db: Queryable & ZeroQueryRunner): WorkbookAgentRunStoreConnection {
  return {
    query: (text, values) => db.query(text, values),
    listWorkbookAgentRunRows: ({ actorUserId, documentId, threadId }) =>
      threadId === undefined
        ? db.run(queries.workbookAgentRun.byWorkbook.fn({ args: { documentId }, ctx: { userID: actorUserId } }))
        : db.run(queries.workbookAgentRun.byThread.fn({ args: { documentId, threadId }, ctx: { userID: actorUserId } })),
  }
}

function toAgentRunRow(row: ZeroWorkbookAgentRunRow): WorkbookAgentRunRow {
  return {
    id: row.id,
    bundleId: row.bundleId,
    workbookId: row.workbookId,
    threadId: row.threadId,
    turnId: row.turnId,
    actorUserId: row.actorUserId,
    goalText: row.goalText,
    planText: row.planText,
    summary: row.summary,
    scope: row.scope,
    riskClass: row.riskClass,
    acceptedScope: row.acceptedScope,
    appliedBy: row.appliedBy,
    baseRevision: row.baseRevision,
    appliedRevision: row.appliedRevision,
    createdAtUnixMs: row.createdAtUnixMs,
    appliedAtUnixMs: row.appliedAtUnixMs,
    contextJson: row.context,
    commandsJson: row.commands,
    previewJson: row.preview,
  }
}

function parseCommands(value: unknown): WorkbookAgentCommand[] | null {
  if (!Array.isArray(value)) {
    return null
  }
  const commands = value.flatMap((entry) => (isWorkbookAgentCommand(entry) ? [entry] : []))
  return commands.length === value.length ? commands : null
}

function normalizeExecutionRecord(row: WorkbookAgentRunRow): WorkbookAgentExecutionRecord | null {
  const baseRevision = parseNonNegativeInteger(row.baseRevision)
  const appliedRevision = parseNonNegativeInteger(row.appliedRevision)
  const createdAtUnixMs = parseNonNegativeInteger(row.createdAtUnixMs)
  const appliedAtUnixMs = parseNonNegativeInteger(row.appliedAtUnixMs)
  const commands = parseCommands(row.commandsJson)
  const bundleId = typeof row.bundleId === 'string' && row.bundleId.length > 0 ? row.bundleId : row.id
  if (
    typeof row.id !== 'string' ||
    typeof bundleId !== 'string' ||
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
    appliedRevision < baseRevision ||
    appliedAtUnixMs < createdAtUnixMs ||
    commands === null
  ) {
    return null
  }
  const context =
    row.contextJson === null || row.contextJson === undefined ? null : isWorkbookAgentContextRef(row.contextJson) ? row.contextJson : null
  const preview = row.previewJson === null || row.previewJson === undefined ? null : decodeWorkbookAgentPreviewSummary(row.previewJson)
  return {
    id: row.id,
    bundleId,
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

function normalizeExecutionRows(rows: readonly ZeroWorkbookAgentRunRow[], limit: number): WorkbookAgentExecutionRecord[] {
  const records: WorkbookAgentExecutionRecord[] = []
  for (const row of rows) {
    const record = normalizeExecutionRecord(toAgentRunRow(row))
    if (record) {
      records.push(record)
      if (records.length >= limit) {
        break
      }
    }
  }
  return records
}

export async function ensureWorkbookAgentRunSchema(db: Queryable): Promise<void> {
  await ensureZeroSchemaTable(db, 'workbook_agent_run', {
    columnOverrides: {
      acceptedScope: { defaultSql: "'full'" },
      appliedBy: { defaultSql: "'user'" },
    },
  })
  await db.query(`ALTER TABLE workbook_agent_run ADD COLUMN IF NOT EXISTS bundle_id TEXT;`)
  await db.query(`
    UPDATE workbook_agent_run
    SET bundle_id = id
    WHERE bundle_id IS NULL OR bundle_id = '';
  `)
  await db.query(`
    ALTER TABLE workbook_agent_run
      ALTER COLUMN bundle_id SET NOT NULL;
  `)
  await addDefaultedColumnIfMissing(db, {
    tableName: 'workbook_agent_run',
    columnName: 'accepted_scope',
    dataType: 'TEXT',
    defaultSql: "'full'",
  })
  await db.query(`
    UPDATE workbook_agent_run
    SET accepted_scope = 'full'
    WHERE accepted_scope IS NULL
       OR accepted_scope NOT IN ('full', 'partial');
  `)
  await enforceDefaultedNotNullColumn(db, {
    tableName: 'workbook_agent_run',
    columnName: 'accepted_scope',
    dataType: 'TEXT',
    defaultSql: "'full'",
  })
  await addDefaultedColumnIfMissing(db, {
    tableName: 'workbook_agent_run',
    columnName: 'applied_by',
    dataType: 'TEXT',
    defaultSql: "'user'",
  })
  await db.query(`
    UPDATE workbook_agent_run
    SET applied_by = 'user'
    WHERE applied_by IS NULL
       OR applied_by NOT IN ('user', 'auto');
  `)
  await enforceDefaultedNotNullColumn(db, {
    tableName: 'workbook_agent_run',
    columnName: 'applied_by',
    dataType: 'TEXT',
    defaultSql: "'user'",
  })
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
  db: WorkbookAgentRunStoreConnection,
  input: {
    documentId: string
    actorUserId: string
    limit?: number
  },
): Promise<WorkbookAgentExecutionRecord[]> {
  return normalizeExecutionRows(await db.listWorkbookAgentRunRows(input), input.limit ?? 20)
}

export async function listWorkbookAgentThreadRuns(
  db: WorkbookAgentRunStoreConnection,
  input: {
    documentId: string
    actorUserId: string
    threadId: string
    limit?: number
  },
): Promise<WorkbookAgentExecutionRecord[]> {
  return normalizeExecutionRows(await db.listWorkbookAgentRunRows(input), input.limit ?? 20)
}
