import type { WorkbookAgentReviewQueueItem } from '@bilig/agent-api'
import type {
  WorkbookAgentExecutionPolicy,
  WorkbookAgentThreadSummary,
  WorkbookAgentTimelineEntry,
  WorkbookAgentUiContext,
} from '@bilig/contracts'
import { queries } from '@bilig/zero-sync'
import {
  addColumnIfMissing,
  addDefaultedColumnIfMissing,
  enforceDefaultedNotNullColumn,
  ensureDefaultedNotNullColumn,
} from './schema-upgrade.js'
import type { Queryable, QueryResultRow, ZeroQueryRunner } from './store.js'
import { runQueryableTransaction, runSequentially } from './transaction-support.js'
import {
  defaultExecutionPolicyForScope,
  hasToolCallState,
  normalizeZeroWorkbookChatThread,
  normalizeReviewQueueItem,
  normalizeThreadSummary,
  normalizeTimelineEntry,
  normalizeTimelineToolName,
  normalizeToolCallRow,
  type WorkbookChatItemRow,
  type NormalizedWorkbookChatThreadModel,
  type WorkbookChatThreadScope,
  type WorkbookChatToolCallRow,
  type WorkbookReviewQueueItemRow,
  type ZeroWorkbookChatThreadRow,
} from './workbook-chat-thread-normalizers.js'

export type { WorkbookChatThreadScope } from './workbook-chat-thread-normalizers.js'

export interface WorkbookAgentThreadStateRecord {
  readonly documentId: string
  readonly threadId: string
  readonly actorUserId: string
  readonly scope: WorkbookChatThreadScope
  readonly executionPolicy: WorkbookAgentExecutionPolicy
  readonly context: WorkbookAgentUiContext | null
  readonly entries: readonly WorkbookAgentTimelineEntry[]
  readonly reviewQueueItems: readonly WorkbookAgentReviewQueueItem[]
  readonly updatedAtUnixMs: number
}

export interface WorkbookChatThreadStoreConnection extends Queryable {
  listWorkbookChatThreadRows(input: {
    readonly documentId: string
    readonly actorUserId: string
  }): Promise<readonly ZeroWorkbookChatThreadRow[]>
  listWorkbookChatItemRows(input: {
    readonly documentId: string
    readonly threadId: string
    readonly actorUserId: string
  }): Promise<readonly QueryResultRow[]>
  listWorkbookChatToolCallRows(input: {
    readonly documentId: string
    readonly threadId: string
    readonly actorUserId: string
  }): Promise<readonly QueryResultRow[]>
  listWorkbookReviewQueueItemRows(input: {
    readonly documentId: string
    readonly threadId: string
    readonly actorUserId: string
  }): Promise<readonly QueryResultRow[]>
}

export function createWorkbookChatThreadStoreConnection(db: Queryable & ZeroQueryRunner): WorkbookChatThreadStoreConnection {
  return {
    query: (text, values) => db.query(text, values),
    listWorkbookChatThreadRows: ({ actorUserId, documentId }) =>
      db.run(queries.workbookChatThread.byWorkbook.fn({ args: { documentId }, ctx: { userID: actorUserId } })),
    listWorkbookChatItemRows: ({ actorUserId, documentId, threadId }) =>
      db.run(queries.workbookChatItem.byThread.fn({ args: { documentId, threadId }, ctx: { userID: actorUserId } })),
    listWorkbookChatToolCallRows: ({ actorUserId, documentId, threadId }) =>
      db.run(queries.workbookChatToolCall.byThread.fn({ args: { documentId, threadId }, ctx: { userID: actorUserId } })),
    listWorkbookReviewQueueItemRows: ({ actorUserId, documentId, threadId }) =>
      db.run(queries.workbookReviewQueueItem.byThread.fn({ args: { documentId, threadId }, ctx: { userID: actorUserId } })),
  }
}

function dedupeTimelineEntries(entries: readonly WorkbookAgentTimelineEntry[]): WorkbookAgentTimelineEntry[] {
  const deduped: WorkbookAgentTimelineEntry[] = []
  const indexById = new Map<string, number>()
  for (const entry of entries) {
    const existingIndex = indexById.get(entry.id)
    if (existingIndex === undefined) {
      indexById.set(entry.id, deduped.length)
      deduped.push(entry)
      continue
    }
    deduped[existingIndex] = entry
  }
  return deduped
}

async function loadZeroWorkbookChatThreads(
  db: WorkbookChatThreadStoreConnection,
  input: {
    readonly documentId: string
    readonly actorUserId: string
  },
): Promise<NormalizedWorkbookChatThreadModel[]> {
  const rows = await db.listWorkbookChatThreadRows(input)
  return rows.flatMap((row) => {
    const normalized = normalizeZeroWorkbookChatThread(row)
    return normalized ? [normalized] : []
  })
}

function isVisibleThreadForActor(thread: NormalizedWorkbookChatThreadModel, actorUserId: string): boolean {
  return thread.actorUserId === actorUserId || thread.scope === 'shared'
}

function compareThreadVisibilityPreference(actorUserId: string) {
  return (left: NormalizedWorkbookChatThreadModel, right: NormalizedWorkbookChatThreadModel): number => {
    const leftActorRank = left.actorUserId === actorUserId ? 0 : 1
    const rightActorRank = right.actorUserId === actorUserId ? 0 : 1
    if (leftActorRank !== rightActorRank) {
      return leftActorRank - rightActorRank
    }
    return right.updatedAtUnixMs - left.updatedAtUnixMs
  }
}

function selectVisibleThread(
  threads: readonly NormalizedWorkbookChatThreadModel[],
  input: {
    readonly threadId: string
    readonly actorUserId: string
  },
): NormalizedWorkbookChatThreadModel | null {
  return (
    threads
      .filter((thread) => thread.threadId === input.threadId && isVisibleThreadForActor(thread, input.actorUserId))
      .toSorted(compareThreadVisibilityPreference(input.actorUserId))[0] ?? null
  )
}

function toChatItemRow(row: QueryResultRow): WorkbookChatItemRow {
  return {
    entryId: row['entryId'],
    turnId: row['turnId'],
    kind: row['kind'],
    text: row['text'],
    phase: row['phase'],
    toolName: row['toolName'],
    toolStatus: row['toolStatus'],
    argumentsText: row['argumentsText'],
    outputText: row['outputText'],
    success: row['success'],
    citationsJson: row['citations'],
    sortOrder: row['sortOrder'],
  }
}

function toChatToolCallRow(row: QueryResultRow): WorkbookChatToolCallRow {
  return {
    entryId: row['entryId'],
    turnId: row['turnId'],
    toolName: row['toolName'],
    toolStatus: row['toolStatus'],
    argumentsText: row['argumentsText'],
    outputText: row['outputText'],
    success: row['success'],
    sortOrder: row['sortOrder'],
  }
}

function toReviewQueueItemRow(row: QueryResultRow): WorkbookReviewQueueItemRow {
  return {
    reviewItemId: row['reviewItemId'],
    workbookId: row['workbookId'],
    threadId: row['threadId'],
    actorUserId: row['actorUserId'],
    turnId: row['turnId'],
    goalText: row['goalText'],
    summary: row['summary'],
    scope: row['scope'],
    riskClass: row['riskClass'],
    reviewMode: row['reviewMode'],
    ownerUserId: row['ownerUserId'],
    status: row['status'],
    decidedByUserId: row['decidedByUserId'],
    decidedAtUnixMs: row['decidedAtUnixMs'],
    baseRevision: row['baseRevision'],
    createdAtUnixMs: row['createdAtUnixMs'],
    contextJson: row['context'],
    commandsJson: row['commands'],
    affectedRangesJson: row['affectedRanges'],
    estimatedAffectedCells: row['estimatedAffectedCells'],
    recommendationsJson: row['recommendations'],
  }
}

export async function ensureWorkbookChatThreadSchema(db: Queryable): Promise<void> {
  await db.query(`
    CREATE TABLE IF NOT EXISTS workbook_chat_thread (
      workbook_id TEXT NOT NULL,
      thread_id TEXT NOT NULL,
      actor_user_id TEXT NOT NULL,
      scope TEXT NOT NULL DEFAULT 'private',
      execution_policy TEXT NOT NULL DEFAULT 'autoApplyAll',
      context_json JSONB,
      entry_count BIGINT NOT NULL DEFAULT 0,
      review_queue_item_count BIGINT NOT NULL DEFAULT 0,
      latest_entry_text TEXT,
      updated_at_unix_ms BIGINT NOT NULL,
      PRIMARY KEY (workbook_id, thread_id, actor_user_id)
    )
  `)
  await ensureDefaultedNotNullColumn(db, {
    tableName: 'workbook_chat_thread',
    columnName: 'scope',
    dataType: 'TEXT',
    defaultSql: "'private'",
  })
  await addColumnIfMissing(db, {
    tableName: 'workbook_chat_thread',
    columnName: 'context_json',
    dataType: 'JSONB',
  })
  await ensureDefaultedNotNullColumn(db, {
    tableName: 'workbook_chat_thread',
    columnName: 'updated_at_unix_ms',
    dataType: 'BIGINT',
    defaultSql: '0',
  })
  await db.query(`
    ALTER TABLE workbook_chat_thread
      ADD COLUMN IF NOT EXISTS execution_policy TEXT;
  `)
  await db.query(`
    UPDATE workbook_chat_thread
    SET execution_policy = CASE WHEN scope = 'shared' THEN 'ownerReview' ELSE 'autoApplyAll' END
    WHERE execution_policy IS NULL;
  `)
  await db.query(`
    ALTER TABLE workbook_chat_thread
      ALTER COLUMN execution_policy SET DEFAULT 'autoApplyAll';
  `)
  await db.query(`
    ALTER TABLE workbook_chat_thread
      ALTER COLUMN execution_policy SET NOT NULL;
  `)
  await addDefaultedColumnIfMissing(db, {
    tableName: 'workbook_chat_thread',
    columnName: 'entry_count',
    dataType: 'BIGINT',
    defaultSql: '0',
  })
  await addDefaultedColumnIfMissing(db, {
    tableName: 'workbook_chat_thread',
    columnName: 'review_queue_item_count',
    dataType: 'BIGINT',
    defaultSql: '0',
  })
  await addColumnIfMissing(db, { tableName: 'workbook_chat_thread', columnName: 'latest_entry_text', dataType: 'TEXT' })
  await db.query(`
    CREATE TABLE IF NOT EXISTS workbook_chat_item (
      workbook_id TEXT NOT NULL,
      thread_id TEXT NOT NULL,
      actor_user_id TEXT NOT NULL,
      entry_id TEXT NOT NULL,
      sort_order INTEGER NOT NULL,
      turn_id TEXT,
      kind TEXT NOT NULL,
      text TEXT,
      phase TEXT,
      tool_name TEXT,
      tool_status TEXT,
      arguments_text TEXT,
      output_text TEXT,
      success BOOLEAN,
      citations_json JSONB,
      PRIMARY KEY (workbook_id, thread_id, actor_user_id, entry_id)
    )
  `)
  await db.query(`
    CREATE TABLE IF NOT EXISTS workbook_chat_tool_call (
      workbook_id TEXT NOT NULL,
      thread_id TEXT NOT NULL,
      actor_user_id TEXT NOT NULL,
      entry_id TEXT NOT NULL,
      sort_order INTEGER NOT NULL,
      turn_id TEXT,
      tool_name TEXT,
      tool_status TEXT,
      arguments_text TEXT,
      output_text TEXT,
      success BOOLEAN,
      PRIMARY KEY (workbook_id, thread_id, actor_user_id, entry_id)
    )
  `)
  await db.query(`
    ALTER TABLE workbook_chat_item
      ADD COLUMN IF NOT EXISTS citations_json JSONB;
  `)
  await db.query(`
    CREATE TABLE IF NOT EXISTS workbook_review_queue_item (
      workbook_id TEXT NOT NULL,
      thread_id TEXT NOT NULL,
      actor_user_id TEXT NOT NULL,
      review_item_id TEXT NOT NULL,
      turn_id TEXT NOT NULL,
      goal_text TEXT NOT NULL,
      summary TEXT NOT NULL,
      scope TEXT NOT NULL,
      risk_class TEXT NOT NULL,
      review_mode TEXT NOT NULL,
      owner_user_id TEXT,
      status TEXT NOT NULL,
      decided_by_user_id TEXT,
      decided_at_unix_ms BIGINT,
      base_revision BIGINT NOT NULL,
      created_at_unix_ms BIGINT NOT NULL,
      context_json JSONB,
      commands_json JSONB NOT NULL,
      affected_ranges_json JSONB NOT NULL,
      estimated_affected_cells BIGINT,
      recommendations_json JSONB NOT NULL DEFAULT '[]'::jsonb,
      PRIMARY KEY (workbook_id, thread_id, actor_user_id, review_item_id)
    )
  `)
  await reconcileWorkbookChatThreadSummaryColumns(db)
  await enforceDefaultedNotNullColumn(db, {
    tableName: 'workbook_chat_thread',
    columnName: 'entry_count',
    dataType: 'BIGINT',
    defaultSql: '0',
  })
  await enforceDefaultedNotNullColumn(db, {
    tableName: 'workbook_chat_thread',
    columnName: 'review_queue_item_count',
    dataType: 'BIGINT',
    defaultSql: '0',
  })
  await db.query(`
    CREATE INDEX IF NOT EXISTS workbook_chat_thread_document_actor_updated_idx
      ON workbook_chat_thread (workbook_id, actor_user_id, updated_at_unix_ms DESC)
  `)
  await db.query(`
    CREATE INDEX IF NOT EXISTS workbook_chat_thread_document_scope_updated_idx
      ON workbook_chat_thread (workbook_id, scope, updated_at_unix_ms DESC)
  `)
  await db.query(`
    CREATE INDEX IF NOT EXISTS workbook_chat_tool_call_thread_order_idx
      ON workbook_chat_tool_call (workbook_id, thread_id, actor_user_id, sort_order ASC)
  `)
  await db.query(`
    CREATE INDEX IF NOT EXISTS workbook_review_queue_item_thread_created_idx
      ON workbook_review_queue_item (workbook_id, thread_id, actor_user_id, created_at_unix_ms ASC)
  `)
}

async function reconcileWorkbookChatThreadSummaryColumns(db: Queryable): Promise<void> {
  await db.query(`
    WITH entry_stats AS (
      SELECT
        workbook_id,
        thread_id,
        actor_user_id,
        COUNT(*)::bigint AS entry_count,
        (ARRAY_AGG(text ORDER BY sort_order DESC, entry_id DESC) FILTER (
          WHERE text IS NOT NULL AND btrim(text) <> ''
        ))[1] AS latest_entry_text
      FROM workbook_chat_item
      GROUP BY workbook_id, thread_id, actor_user_id
    ),
    review_stats AS (
      SELECT
        workbook_id,
        thread_id,
        actor_user_id,
        COUNT(*)::bigint AS review_queue_item_count
      FROM workbook_review_queue_item
      GROUP BY workbook_id, thread_id, actor_user_id
    ),
    thread_stats AS (
      SELECT
        thread.workbook_id,
        thread.thread_id,
        thread.actor_user_id,
        COALESCE(entry_stats.entry_count, 0)::bigint AS entry_count,
        COALESCE(review_stats.review_queue_item_count, 0)::bigint AS review_queue_item_count,
        entry_stats.latest_entry_text
      FROM workbook_chat_thread AS thread
      LEFT JOIN entry_stats
        ON entry_stats.workbook_id = thread.workbook_id
        AND entry_stats.thread_id = thread.thread_id
        AND entry_stats.actor_user_id = thread.actor_user_id
      LEFT JOIN review_stats
        ON review_stats.workbook_id = thread.workbook_id
        AND review_stats.thread_id = thread.thread_id
        AND review_stats.actor_user_id = thread.actor_user_id
    )
    UPDATE workbook_chat_thread AS thread
    SET
      entry_count = thread_stats.entry_count,
      review_queue_item_count = thread_stats.review_queue_item_count,
      latest_entry_text = thread_stats.latest_entry_text
    FROM thread_stats
    WHERE thread.workbook_id = thread_stats.workbook_id
      AND thread.thread_id = thread_stats.thread_id
      AND thread.actor_user_id = thread_stats.actor_user_id
      AND (
        thread.entry_count IS DISTINCT FROM thread_stats.entry_count
        OR thread.review_queue_item_count IS DISTINCT FROM thread_stats.review_queue_item_count
        OR thread.latest_entry_text IS DISTINCT FROM thread_stats.latest_entry_text
      )
  `)
}

export async function saveWorkbookAgentThreadState(db: Queryable, record: WorkbookAgentThreadStateRecord): Promise<void> {
  await runQueryableTransaction(db, async (transactionDb) => {
    await persistWorkbookAgentThreadState(transactionDb, record)
  })
}

async function persistWorkbookAgentThreadState(db: Queryable, record: WorkbookAgentThreadStateRecord): Promise<void> {
  const persistedEntries = dedupeTimelineEntries(record.entries)
  const latestEntryText =
    [...persistedEntries].toReversed().find((entry) => typeof entry.text === 'string' && entry.text.trim().length > 0)?.text ?? null
  await db.query(
    `
      INSERT INTO workbook_chat_thread (
        workbook_id,
        thread_id,
        actor_user_id,
        scope,
        execution_policy,
        context_json,
        entry_count,
        review_queue_item_count,
        latest_entry_text,
        updated_at_unix_ms
      )
      VALUES ($1, $2, $3, $4, $5, $6::jsonb, $7, $8, $9, $10)
      ON CONFLICT (workbook_id, thread_id, actor_user_id)
      DO UPDATE SET
        scope = EXCLUDED.scope,
        execution_policy = EXCLUDED.execution_policy,
        context_json = EXCLUDED.context_json,
        entry_count = EXCLUDED.entry_count,
        review_queue_item_count = EXCLUDED.review_queue_item_count,
        latest_entry_text = EXCLUDED.latest_entry_text,
        updated_at_unix_ms = EXCLUDED.updated_at_unix_ms
    `,
    [
      record.documentId,
      record.threadId,
      record.actorUserId,
      record.scope,
      record.executionPolicy,
      JSON.stringify(record.context),
      persistedEntries.length,
      record.reviewQueueItems.length,
      latestEntryText,
      record.updatedAtUnixMs,
    ],
  )
  await db.query(
    `
      DELETE FROM workbook_chat_item
      WHERE workbook_id = $1 AND thread_id = $2 AND actor_user_id = $3
    `,
    [record.documentId, record.threadId, record.actorUserId],
  )
  await db.query(
    `
      DELETE FROM workbook_chat_tool_call
      WHERE workbook_id = $1 AND thread_id = $2 AND actor_user_id = $3
    `,
    [record.documentId, record.threadId, record.actorUserId],
  )
  await runSequentially(persistedEntries, async (entry, index) => {
    await db.query(
      `
          INSERT INTO workbook_chat_item (
            workbook_id,
            thread_id,
            actor_user_id,
            entry_id,
            sort_order,
            turn_id,
            kind,
            text,
            phase,
            tool_name,
            tool_status,
            arguments_text,
            output_text,
            success,
            citations_json
          )
          VALUES (
            $1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11, $12, $13, $14, $15::jsonb
          )
          ON CONFLICT (workbook_id, thread_id, actor_user_id, entry_id)
          DO UPDATE SET
            sort_order = EXCLUDED.sort_order,
            turn_id = EXCLUDED.turn_id,
            kind = EXCLUDED.kind,
            text = EXCLUDED.text,
            phase = EXCLUDED.phase,
            tool_name = EXCLUDED.tool_name,
            tool_status = EXCLUDED.tool_status,
            arguments_text = EXCLUDED.arguments_text,
            output_text = EXCLUDED.output_text,
            success = EXCLUDED.success,
            citations_json = EXCLUDED.citations_json
        `,
      [
        record.documentId,
        record.threadId,
        record.actorUserId,
        entry.id,
        index,
        entry.turnId,
        entry.kind,
        entry.text,
        entry.phase,
        normalizeTimelineToolName(entry.toolName),
        entry.toolStatus,
        entry.argumentsText,
        entry.outputText,
        entry.success,
        JSON.stringify(entry.citations),
      ],
    )
    if (hasToolCallState(entry)) {
      await db.query(
        `
          INSERT INTO workbook_chat_tool_call (
            workbook_id,
            thread_id,
            actor_user_id,
            entry_id,
            sort_order,
            turn_id,
            tool_name,
            tool_status,
            arguments_text,
            output_text,
            success
          )
          VALUES (
            $1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11
          )
          ON CONFLICT (workbook_id, thread_id, actor_user_id, entry_id)
          DO UPDATE SET
            sort_order = EXCLUDED.sort_order,
            turn_id = EXCLUDED.turn_id,
            tool_name = EXCLUDED.tool_name,
            tool_status = EXCLUDED.tool_status,
            arguments_text = EXCLUDED.arguments_text,
            output_text = EXCLUDED.output_text,
            success = EXCLUDED.success
        `,
        [
          record.documentId,
          record.threadId,
          record.actorUserId,
          entry.id,
          index,
          entry.turnId,
          normalizeTimelineToolName(entry.toolName),
          entry.toolStatus,
          entry.argumentsText,
          entry.outputText,
          entry.success,
        ],
      )
    }
  })
  await db.query(
    `
      DELETE FROM workbook_review_queue_item
      WHERE workbook_id = $1 AND thread_id = $2 AND actor_user_id = $3
    `,
    [record.documentId, record.threadId, record.actorUserId],
  )
  await runSequentially(record.reviewQueueItems, async (reviewItem) => {
    await db.query(
      `
          INSERT INTO workbook_review_queue_item (
            workbook_id,
            thread_id,
            actor_user_id,
            review_item_id,
            turn_id,
            goal_text,
            summary,
            scope,
            risk_class,
            review_mode,
            owner_user_id,
            status,
            decided_by_user_id,
            decided_at_unix_ms,
            base_revision,
            created_at_unix_ms,
            context_json,
            commands_json,
            affected_ranges_json,
            estimated_affected_cells,
            recommendations_json
          )
          VALUES (
            $1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11, $12, $13, $14, $15, $16, $17::jsonb, $18::jsonb, $19::jsonb, $20, $21::jsonb
          )
        `,
      [
        record.documentId,
        record.threadId,
        record.actorUserId,
        reviewItem.id,
        reviewItem.turnId,
        reviewItem.goalText,
        reviewItem.summary,
        reviewItem.scope,
        reviewItem.riskClass,
        reviewItem.reviewMode,
        reviewItem.ownerUserId,
        reviewItem.status,
        reviewItem.decidedByUserId,
        reviewItem.decidedAtUnixMs,
        reviewItem.baseRevision,
        reviewItem.createdAtUnixMs,
        JSON.stringify(reviewItem.context),
        JSON.stringify(reviewItem.commands),
        JSON.stringify(reviewItem.affectedRanges),
        reviewItem.estimatedAffectedCells,
        JSON.stringify(reviewItem.recommendations),
      ],
    )
  })
}

export async function loadWorkbookAgentThreadState(
  db: WorkbookChatThreadStoreConnection,
  input: {
    documentId: string
    threadId: string
    actorUserId: string
  },
): Promise<WorkbookAgentThreadStateRecord | null> {
  const thread = selectVisibleThread(await loadZeroWorkbookChatThreads(db, input), input)
  if (!thread) {
    return null
  }
  const executionPolicy = thread.executionPolicy ?? defaultExecutionPolicyForScope(thread.scope)
  const childRowInput = {
    documentId: thread.workbookId,
    threadId: thread.threadId,
    actorUserId: input.actorUserId,
  }
  const [itemRows, toolCallRows, reviewQueueRows] = await Promise.all([
    db.listWorkbookChatItemRows(childRowInput),
    db.listWorkbookChatToolCallRows(childRowInput),
    db.listWorkbookReviewQueueItemRows(childRowInput),
  ])
  const toolCallsByEntryId = new Map(
    toolCallRows.flatMap((row) => {
      const normalized = normalizeToolCallRow(toChatToolCallRow(row))
      return normalized ? [[normalized.entryId, normalized] as const] : []
    }),
  )
  const entries = itemRows
    .map((row) =>
      normalizeTimelineEntry({
        ...toChatItemRow(row),
        ...(typeof row['entryId'] === 'string' ? toolCallsByEntryId.get(row['entryId']) : undefined),
      }),
    )
    .filter((entry): entry is WorkbookAgentTimelineEntry => entry !== null)
  const reviewQueueItems = reviewQueueRows.flatMap((row) => {
    const normalized = normalizeReviewQueueItem(toReviewQueueItemRow(row))
    return normalized ? [normalized] : []
  })
  return {
    documentId: thread.workbookId,
    threadId: thread.threadId,
    actorUserId: thread.actorUserId,
    scope: thread.scope,
    executionPolicy,
    context: thread.context,
    entries,
    reviewQueueItems,
    updatedAtUnixMs: thread.updatedAtUnixMs,
  }
}

export async function listWorkbookAgentThreadSummaries(
  db: WorkbookChatThreadStoreConnection,
  input: {
    documentId: string
    actorUserId: string
  },
): Promise<WorkbookAgentThreadSummary[]> {
  const visibleThreadsByThreadId = new Map<string, NormalizedWorkbookChatThreadModel>()
  for (const thread of await loadZeroWorkbookChatThreads(db, input)) {
    if (!isVisibleThreadForActor(thread, input.actorUserId)) {
      continue
    }
    const existing = visibleThreadsByThreadId.get(thread.threadId)
    if (!existing || compareThreadVisibilityPreference(input.actorUserId)(thread, existing) < 0) {
      visibleThreadsByThreadId.set(thread.threadId, thread)
    }
  }
  const visibleThreads = [...visibleThreadsByThreadId.values()].toSorted((left, right) => right.updatedAtUnixMs - left.updatedAtUnixMs)
  return visibleThreads.map((thread) => normalizeThreadSummary(thread))
}
