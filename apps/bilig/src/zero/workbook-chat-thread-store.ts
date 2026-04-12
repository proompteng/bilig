import {
  isWorkbookAgentCommand,
  isWorkbookAgentCommandBundle,
  type WorkbookAgentExecutionPolicy,
  normalizeWorkbookAgentToolName,
  type WorkbookAgentCommandBundle,
  type WorkbookAgentSharedReviewState,
} from "@bilig/agent-api";
import type {
  WorkbookAgentTimelineCitation,
  WorkbookAgentThreadSummary,
  WorkbookAgentTimelineEntry,
  WorkbookAgentToolStatus,
  WorkbookAgentUiContext,
} from "@bilig/contracts";
import type { QueryResultRow, Queryable } from "./store.js";

export type WorkbookChatThreadScope = "private" | "shared";

export interface WorkbookAgentThreadStateRecord {
  readonly documentId: string;
  readonly threadId: string;
  readonly actorUserId: string;
  readonly scope: WorkbookChatThreadScope;
  readonly executionPolicy: WorkbookAgentExecutionPolicy;
  readonly context: WorkbookAgentUiContext | null;
  readonly entries: readonly WorkbookAgentTimelineEntry[];
  readonly pendingBundle: WorkbookAgentCommandBundle | null;
  readonly updatedAtUnixMs: number;
}

interface WorkbookChatThreadRow extends QueryResultRow {
  readonly workbookId?: unknown;
  readonly threadId?: unknown;
  readonly actorUserId?: unknown;
  readonly scope?: unknown;
  readonly executionPolicy?: unknown;
  readonly contextJson?: unknown;
  readonly updatedAtUnixMs?: unknown;
  readonly entryCount?: unknown;
  readonly hasPendingBundle?: unknown;
  readonly latestEntryText?: unknown;
}

interface WorkbookChatItemRow extends QueryResultRow {
  readonly entryId?: unknown;
  readonly turnId?: unknown;
  readonly kind?: unknown;
  readonly text?: unknown;
  readonly phase?: unknown;
  readonly toolName?: unknown;
  readonly toolStatus?: unknown;
  readonly argumentsText?: unknown;
  readonly outputText?: unknown;
  readonly success?: unknown;
  readonly citationsJson?: unknown;
  readonly sortOrder?: unknown;
}

interface WorkbookChatToolCallRow extends QueryResultRow {
  readonly entryId?: unknown;
  readonly turnId?: unknown;
  readonly toolName?: unknown;
  readonly toolStatus?: unknown;
  readonly argumentsText?: unknown;
  readonly outputText?: unknown;
  readonly success?: unknown;
  readonly sortOrder?: unknown;
}

interface WorkbookPendingBundleRow extends QueryResultRow {
  readonly bundleId?: unknown;
  readonly workbookId?: unknown;
  readonly threadId?: unknown;
  readonly actorUserId?: unknown;
  readonly turnId?: unknown;
  readonly goalText?: unknown;
  readonly summary?: unknown;
  readonly scope?: unknown;
  readonly riskClass?: unknown;
  readonly approvalMode?: unknown;
  readonly baseRevision?: unknown;
  readonly createdAtUnixMs?: unknown;
  readonly contextJson?: unknown;
  readonly commandsJson?: unknown;
  readonly affectedRangesJson?: unknown;
  readonly estimatedAffectedCells?: unknown;
  readonly sharedReviewJson?: unknown;
}

interface WorkbookChatThreadSummaryRow extends QueryResultRow {
  readonly threadId?: unknown;
  readonly scope?: unknown;
  readonly ownerUserId?: unknown;
  readonly updatedAtUnixMs?: unknown;
  readonly entryCount?: unknown;
  readonly hasPendingBundle?: unknown;
  readonly latestEntryText?: unknown;
}

function dedupeTimelineEntries(
  entries: readonly WorkbookAgentTimelineEntry[],
): WorkbookAgentTimelineEntry[] {
  const deduped: WorkbookAgentTimelineEntry[] = [];
  const indexById = new Map<string, number>();
  for (const entry of entries) {
    const existingIndex = indexById.get(entry.id);
    if (existingIndex === undefined) {
      indexById.set(entry.id, deduped.length);
      deduped.push(entry);
      continue;
    }
    deduped[existingIndex] = entry;
  }
  return deduped;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function parseNumericValue(value: unknown): number | null {
  if (typeof value === "number" && Number.isFinite(value)) {
    return Math.trunc(value);
  }
  if (typeof value === "string" && value.length > 0) {
    const parsed = Number.parseInt(value, 10);
    return Number.isFinite(parsed) ? parsed : null;
  }
  return null;
}

function normalizeTimelineToolName(toolName: string | null): string | null {
  if (typeof toolName !== "string") {
    return null;
  }
  return normalizeWorkbookAgentToolName(toolName);
}

function isWorkbookAgentUiContext(value: unknown): value is WorkbookAgentUiContext {
  return (
    isRecord(value) &&
    isRecord(value["selection"]) &&
    typeof value["selection"]["sheetName"] === "string" &&
    typeof value["selection"]["address"] === "string" &&
    (value["selection"]["range"] === undefined ||
      (isRecord(value["selection"]["range"]) &&
        typeof value["selection"]["range"]["startAddress"] === "string" &&
        typeof value["selection"]["range"]["endAddress"] === "string")) &&
    isRecord(value["viewport"]) &&
    typeof value["viewport"]["rowStart"] === "number" &&
    typeof value["viewport"]["rowEnd"] === "number" &&
    typeof value["viewport"]["colStart"] === "number" &&
    typeof value["viewport"]["colEnd"] === "number"
  );
}

function isExecutionPolicy(value: unknown): value is WorkbookAgentExecutionPolicy {
  return value === "autoApplySafe" || value === "autoApplyAll" || value === "ownerReview";
}

function isToolStatus(value: unknown): value is WorkbookAgentToolStatus {
  return value === "inProgress" || value === "completed" || value === "failed" || value === null;
}

function isTimelineKind(value: unknown): value is WorkbookAgentTimelineEntry["kind"] {
  return (
    value === "user" ||
    value === "assistant" ||
    value === "plan" ||
    value === "tool" ||
    value === "system"
  );
}

function isTimelineCitation(value: unknown): value is WorkbookAgentTimelineCitation {
  return (
    (isRecord(value) &&
      value["kind"] === "range" &&
      typeof value["sheetName"] === "string" &&
      typeof value["startAddress"] === "string" &&
      typeof value["endAddress"] === "string" &&
      (value["role"] === "target" || value["role"] === "source")) ||
    (isRecord(value) && value["kind"] === "revision" && typeof value["revision"] === "number")
  );
}

function normalizeTimelineEntry(row: WorkbookChatItemRow): WorkbookAgentTimelineEntry | null {
  if (
    typeof row.entryId !== "string" ||
    !isTimelineKind(row.kind) ||
    (row.turnId !== null && row.turnId !== undefined && typeof row.turnId !== "string") ||
    (row.text !== null && row.text !== undefined && typeof row.text !== "string") ||
    (row.phase !== null && row.phase !== undefined && typeof row.phase !== "string") ||
    (row.toolName !== null && row.toolName !== undefined && typeof row.toolName !== "string") ||
    !isToolStatus(row.toolStatus ?? null) ||
    (row.argumentsText !== null &&
      row.argumentsText !== undefined &&
      typeof row.argumentsText !== "string") ||
    (row.outputText !== null &&
      row.outputText !== undefined &&
      typeof row.outputText !== "string") ||
    (row.success !== null && row.success !== undefined && typeof row.success !== "boolean") ||
    (row.citationsJson !== null &&
      row.citationsJson !== undefined &&
      (!Array.isArray(row.citationsJson) ||
        !row.citationsJson.every((entry) => isTimelineCitation(entry))))
  ) {
    return null;
  }
  const toolStatus: WorkbookAgentToolStatus =
    row.toolStatus === "inProgress" || row.toolStatus === "completed" || row.toolStatus === "failed"
      ? row.toolStatus
      : null;
  return {
    id: row.entryId,
    kind: row.kind,
    turnId: typeof row.turnId === "string" ? row.turnId : null,
    text: typeof row.text === "string" ? row.text : null,
    phase: typeof row.phase === "string" ? row.phase : null,
    toolName: normalizeTimelineToolName(typeof row.toolName === "string" ? row.toolName : null),
    toolStatus,
    argumentsText: typeof row.argumentsText === "string" ? row.argumentsText : null,
    outputText: typeof row.outputText === "string" ? row.outputText : null,
    success: typeof row.success === "boolean" ? row.success : null,
    citations: Array.isArray(row.citationsJson) ? [...row.citationsJson] : [],
  };
}

function normalizeToolCallRow(row: WorkbookChatToolCallRow): {
  readonly entryId: string;
  readonly turnId: string | null;
  readonly toolName: string | null;
  readonly toolStatus: WorkbookAgentToolStatus;
  readonly argumentsText: string | null;
  readonly outputText: string | null;
  readonly success: boolean | null;
} | null {
  if (
    typeof row.entryId !== "string" ||
    (row.turnId !== null && row.turnId !== undefined && typeof row.turnId !== "string") ||
    (row.toolName !== null && row.toolName !== undefined && typeof row.toolName !== "string") ||
    !isToolStatus(row.toolStatus ?? null) ||
    (row.argumentsText !== null &&
      row.argumentsText !== undefined &&
      typeof row.argumentsText !== "string") ||
    (row.outputText !== null &&
      row.outputText !== undefined &&
      typeof row.outputText !== "string") ||
    (row.success !== null && row.success !== undefined && typeof row.success !== "boolean")
  ) {
    return null;
  }
  const toolStatus: WorkbookAgentToolStatus =
    row.toolStatus === "inProgress" || row.toolStatus === "completed" || row.toolStatus === "failed"
      ? row.toolStatus
      : null;
  return {
    entryId: row.entryId,
    turnId: typeof row.turnId === "string" ? row.turnId : null,
    toolName: normalizeTimelineToolName(typeof row.toolName === "string" ? row.toolName : null),
    toolStatus,
    argumentsText: typeof row.argumentsText === "string" ? row.argumentsText : null,
    outputText: typeof row.outputText === "string" ? row.outputText : null,
    success: typeof row.success === "boolean" ? row.success : null,
  };
}

function hasToolCallState(entry: WorkbookAgentTimelineEntry): boolean {
  return (
    entry.kind === "tool" ||
    entry.toolName !== null ||
    entry.toolStatus !== null ||
    entry.argumentsText !== null ||
    entry.outputText !== null ||
    entry.success !== null
  );
}

function normalizePendingBundle(row: WorkbookPendingBundleRow): WorkbookAgentCommandBundle | null {
  const baseRevision = parseNumericValue(row.baseRevision);
  const createdAtUnixMs = parseNumericValue(row.createdAtUnixMs);
  const estimatedAffectedCells =
    row.estimatedAffectedCells === null || row.estimatedAffectedCells === undefined
      ? null
      : parseNumericValue(row.estimatedAffectedCells);
  if (
    typeof row.bundleId !== "string" ||
    typeof row.workbookId !== "string" ||
    typeof row.threadId !== "string" ||
    typeof row.turnId !== "string" ||
    typeof row.goalText !== "string" ||
    typeof row.summary !== "string" ||
    (row.scope !== "selection" && row.scope !== "sheet" && row.scope !== "workbook") ||
    (row.riskClass !== "low" && row.riskClass !== "medium" && row.riskClass !== "high") ||
    (row.approvalMode !== "auto" &&
      row.approvalMode !== "preview" &&
      row.approvalMode !== "explicit") ||
    baseRevision === null ||
    createdAtUnixMs === null ||
    !Array.isArray(row.commandsJson) ||
    !row.commandsJson.every((entry) => isWorkbookAgentCommand(entry)) ||
    !Array.isArray(row.affectedRangesJson)
  ) {
    return null;
  }
  const bundle = {
    id: row.bundleId,
    documentId: row.workbookId,
    threadId: row.threadId,
    turnId: row.turnId,
    goalText: row.goalText,
    summary: row.summary,
    scope: row.scope,
    riskClass: row.riskClass,
    approvalMode: row.approvalMode,
    baseRevision,
    createdAtUnixMs,
    context: isWorkbookAgentUiContext(row.contextJson) ? row.contextJson : null,
    commands: [...row.commandsJson],
    affectedRanges: [...row.affectedRangesJson],
    estimatedAffectedCells,
    sharedReview: normalizeSharedReviewState(row.sharedReviewJson),
  } satisfies WorkbookAgentCommandBundle;
  return isWorkbookAgentCommandBundle(bundle) ? bundle : null;
}

function isSharedReviewState(value: unknown): value is WorkbookAgentSharedReviewState {
  return (
    typeof value === "object" &&
    value !== null &&
    "ownerUserId" in value &&
    typeof value.ownerUserId === "string" &&
    "status" in value &&
    (value.status === "pending" || value.status === "approved" || value.status === "rejected") &&
    "decidedByUserId" in value &&
    (value.decidedByUserId === null || typeof value.decidedByUserId === "string") &&
    "decidedAtUnixMs" in value &&
    (value.decidedAtUnixMs === null || typeof value.decidedAtUnixMs === "number") &&
    "recommendations" in value &&
    Array.isArray(value.recommendations) &&
    value.recommendations.every(
      (recommendation) =>
        typeof recommendation === "object" &&
        recommendation !== null &&
        "userId" in recommendation &&
        typeof recommendation.userId === "string" &&
        "decision" in recommendation &&
        (recommendation.decision === "approved" || recommendation.decision === "rejected") &&
        "decidedAtUnixMs" in recommendation &&
        typeof recommendation.decidedAtUnixMs === "number",
    )
  );
}

function normalizeSharedReviewState(value: unknown): WorkbookAgentSharedReviewState | null {
  if (isSharedReviewState(value)) {
    return value;
  }
  if (
    typeof value === "object" &&
    value !== null &&
    "ownerUserId" in value &&
    typeof value.ownerUserId === "string" &&
    "status" in value &&
    (value.status === "pending" || value.status === "approved" || value.status === "rejected") &&
    "decidedByUserId" in value &&
    (value.decidedByUserId === null || typeof value.decidedByUserId === "string") &&
    "decidedAtUnixMs" in value &&
    (value.decidedAtUnixMs === null || typeof value.decidedAtUnixMs === "number")
  ) {
    return {
      ownerUserId: value.ownerUserId,
      status: value.status,
      decidedByUserId: value.decidedByUserId,
      decidedAtUnixMs: value.decidedAtUnixMs,
      recommendations: [],
    };
  }
  return null;
}

export async function ensureWorkbookChatThreadSchema(db: Queryable): Promise<void> {
  await db.query(`
    CREATE TABLE IF NOT EXISTS workbook_chat_thread (
      workbook_id TEXT NOT NULL,
      thread_id TEXT NOT NULL,
      actor_user_id TEXT NOT NULL,
      scope TEXT NOT NULL DEFAULT 'private',
      execution_policy TEXT NOT NULL DEFAULT 'ownerReview',
      context_json JSONB,
      entry_count BIGINT NOT NULL DEFAULT 0,
      has_pending_bundle BOOLEAN NOT NULL DEFAULT FALSE,
      latest_entry_text TEXT,
      updated_at_unix_ms BIGINT NOT NULL,
      PRIMARY KEY (workbook_id, thread_id, actor_user_id)
    )
  `);
  await db.query(`
    ALTER TABLE workbook_chat_thread
      ADD COLUMN IF NOT EXISTS entry_count BIGINT NOT NULL DEFAULT 0;
  `);
  await db.query(`
    ALTER TABLE workbook_chat_thread
      ADD COLUMN IF NOT EXISTS has_pending_bundle BOOLEAN NOT NULL DEFAULT FALSE;
  `);
  await db.query(`
    ALTER TABLE workbook_chat_thread
      ADD COLUMN IF NOT EXISTS latest_entry_text TEXT;
  `);
  await db.query(`
    ALTER TABLE workbook_chat_thread
      ADD COLUMN IF NOT EXISTS execution_policy TEXT NOT NULL DEFAULT 'ownerReview';
  `);
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
  `);
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
  `);
  await db.query(`
    ALTER TABLE workbook_chat_item
      ADD COLUMN IF NOT EXISTS citations_json JSONB;
  `);
  await db.query(`
    CREATE TABLE IF NOT EXISTS workbook_pending_bundle (
      workbook_id TEXT NOT NULL,
      thread_id TEXT NOT NULL,
      actor_user_id TEXT NOT NULL,
      bundle_id TEXT NOT NULL,
      turn_id TEXT NOT NULL,
      goal_text TEXT NOT NULL,
      summary TEXT NOT NULL,
      scope TEXT NOT NULL,
      risk_class TEXT NOT NULL,
      approval_mode TEXT NOT NULL,
      base_revision BIGINT NOT NULL,
      created_at_unix_ms BIGINT NOT NULL,
      context_json JSONB,
      commands_json JSONB NOT NULL,
      affected_ranges_json JSONB NOT NULL,
      estimated_affected_cells BIGINT,
      shared_review_json JSONB,
      PRIMARY KEY (workbook_id, thread_id, actor_user_id)
    )
  `);
  await db.query(`
    ALTER TABLE workbook_pending_bundle
      ADD COLUMN IF NOT EXISTS shared_review_json JSONB;
  `);
  await db.query(`
    CREATE INDEX IF NOT EXISTS workbook_chat_thread_document_actor_updated_idx
      ON workbook_chat_thread (workbook_id, actor_user_id, updated_at_unix_ms DESC)
  `);
  await db.query(`
    CREATE INDEX IF NOT EXISTS workbook_chat_tool_call_thread_order_idx
      ON workbook_chat_tool_call (workbook_id, thread_id, actor_user_id, sort_order ASC)
  `);
}

function normalizeThreadSummary(
  row: WorkbookChatThreadSummaryRow,
): WorkbookAgentThreadSummary | null {
  const updatedAtUnixMs = parseNumericValue(row.updatedAtUnixMs);
  const entryCount = parseNumericValue(row.entryCount);
  if (
    typeof row.threadId !== "string" ||
    (row.scope !== "private" && row.scope !== "shared") ||
    typeof row.ownerUserId !== "string" ||
    updatedAtUnixMs === null ||
    entryCount === null ||
    typeof row.hasPendingBundle !== "boolean" ||
    (row.latestEntryText !== null &&
      row.latestEntryText !== undefined &&
      typeof row.latestEntryText !== "string")
  ) {
    return null;
  }
  return {
    threadId: row.threadId,
    scope: row.scope,
    ownerUserId: row.ownerUserId,
    updatedAtUnixMs,
    entryCount,
    hasPendingBundle: row.hasPendingBundle,
    latestEntryText: typeof row.latestEntryText === "string" ? row.latestEntryText : null,
  };
}

export async function saveWorkbookAgentThreadState(
  db: Queryable,
  record: WorkbookAgentThreadStateRecord,
): Promise<void> {
  const persistedEntries = dedupeTimelineEntries(record.entries);
  const latestEntryText =
    [...persistedEntries]
      .toReversed()
      .find((entry) => typeof entry.text === "string" && entry.text.trim().length > 0)?.text ??
    null;
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
        has_pending_bundle,
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
        has_pending_bundle = EXCLUDED.has_pending_bundle,
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
      record.pendingBundle !== null,
      latestEntryText,
      record.updatedAtUnixMs,
    ],
  );
  await db.query(
    `
      DELETE FROM workbook_chat_item
      WHERE workbook_id = $1 AND thread_id = $2 AND actor_user_id = $3
    `,
    [record.documentId, record.threadId, record.actorUserId],
  );
  await db.query(
    `
      DELETE FROM workbook_chat_tool_call
      WHERE workbook_id = $1 AND thread_id = $2 AND actor_user_id = $3
    `,
    [record.documentId, record.threadId, record.actorUserId],
  );
  await Promise.all(
    persistedEntries.map(async (entry, index) => {
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
      );
      if (!hasToolCallState(entry)) {
        return;
      }
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
      );
    }),
  );
  if (record.pendingBundle) {
    await db.query(
      `
        INSERT INTO workbook_pending_bundle (
          workbook_id,
          thread_id,
          actor_user_id,
          bundle_id,
          turn_id,
          goal_text,
          summary,
          scope,
          risk_class,
          approval_mode,
          base_revision,
          created_at_unix_ms,
          context_json,
          commands_json,
          affected_ranges_json,
          estimated_affected_cells,
          shared_review_json
        )
        VALUES (
          $1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11, $12, $13::jsonb, $14::jsonb, $15::jsonb, $16, $17::jsonb
        )
        ON CONFLICT (workbook_id, thread_id, actor_user_id)
        DO UPDATE SET
          bundle_id = EXCLUDED.bundle_id,
          turn_id = EXCLUDED.turn_id,
          goal_text = EXCLUDED.goal_text,
          summary = EXCLUDED.summary,
          scope = EXCLUDED.scope,
          risk_class = EXCLUDED.risk_class,
          approval_mode = EXCLUDED.approval_mode,
          base_revision = EXCLUDED.base_revision,
          created_at_unix_ms = EXCLUDED.created_at_unix_ms,
          context_json = EXCLUDED.context_json,
          commands_json = EXCLUDED.commands_json,
          affected_ranges_json = EXCLUDED.affected_ranges_json,
          estimated_affected_cells = EXCLUDED.estimated_affected_cells,
          shared_review_json = EXCLUDED.shared_review_json
      `,
      [
        record.documentId,
        record.threadId,
        record.actorUserId,
        record.pendingBundle.id,
        record.pendingBundle.turnId,
        record.pendingBundle.goalText,
        record.pendingBundle.summary,
        record.pendingBundle.scope,
        record.pendingBundle.riskClass,
        record.pendingBundle.approvalMode,
        record.pendingBundle.baseRevision,
        record.pendingBundle.createdAtUnixMs,
        JSON.stringify(record.pendingBundle.context),
        JSON.stringify(record.pendingBundle.commands),
        JSON.stringify(record.pendingBundle.affectedRanges),
        record.pendingBundle.estimatedAffectedCells,
        JSON.stringify(record.pendingBundle.sharedReview ?? null),
      ],
    );
  } else {
    await db.query(
      `
        DELETE FROM workbook_pending_bundle
        WHERE workbook_id = $1 AND thread_id = $2 AND actor_user_id = $3
      `,
      [record.documentId, record.threadId, record.actorUserId],
    );
  }
}

export async function loadWorkbookAgentThreadState(
  db: Queryable,
  input: {
    documentId: string;
    threadId: string;
    actorUserId: string;
  },
): Promise<WorkbookAgentThreadStateRecord | null> {
  const threadResult = await db.query<WorkbookChatThreadRow>(
    `
      SELECT
        workbook_id AS "workbookId",
        thread_id AS "threadId",
        actor_user_id AS "actorUserId",
        scope AS "scope",
        execution_policy AS "executionPolicy",
        context_json AS "contextJson",
        updated_at_unix_ms AS "updatedAtUnixMs"
      FROM workbook_chat_thread
      WHERE workbook_id = $1
        AND thread_id = $2
        AND (actor_user_id = $3 OR scope = 'shared')
      ORDER BY
        CASE WHEN actor_user_id = $3 THEN 0 ELSE 1 END ASC,
        updated_at_unix_ms DESC
      LIMIT 1
    `,
    [input.documentId, input.threadId, input.actorUserId],
  );
  const thread = threadResult.rows[0];
  const updatedAtUnixMs = parseNumericValue(thread?.updatedAtUnixMs);
  if (
    !thread ||
    typeof thread.workbookId !== "string" ||
    typeof thread.threadId !== "string" ||
    typeof thread.actorUserId !== "string" ||
    (thread.scope !== "private" && thread.scope !== "shared") ||
    !isExecutionPolicy(thread.executionPolicy) ||
    updatedAtUnixMs === null
  ) {
    return null;
  }
  const [itemResult, toolCallResult, pendingBundleResult] = await Promise.all([
    db.query<WorkbookChatItemRow>(
      `
        SELECT
          entry_id AS "entryId",
          turn_id AS "turnId",
          kind AS "kind",
          text AS "text",
          phase AS "phase",
          tool_name AS "toolName",
          tool_status AS "toolStatus",
          arguments_text AS "argumentsText",
          output_text AS "outputText",
          success AS "success",
          citations_json AS "citationsJson",
          sort_order AS "sortOrder"
        FROM workbook_chat_item
        WHERE workbook_id = $1 AND thread_id = $2 AND actor_user_id = $3
        ORDER BY sort_order ASC
      `,
      [thread.workbookId, thread.threadId, thread.actorUserId],
    ),
    db.query<WorkbookChatToolCallRow>(
      `
        SELECT
          entry_id AS "entryId",
          turn_id AS "turnId",
          tool_name AS "toolName",
          tool_status AS "toolStatus",
          arguments_text AS "argumentsText",
          output_text AS "outputText",
          success AS "success",
          sort_order AS "sortOrder"
        FROM workbook_chat_tool_call
        WHERE workbook_id = $1 AND thread_id = $2 AND actor_user_id = $3
        ORDER BY sort_order ASC
      `,
      [thread.workbookId, thread.threadId, thread.actorUserId],
    ),
    db.query<WorkbookPendingBundleRow>(
      `
        SELECT
          bundle_id AS "bundleId",
          workbook_id AS "workbookId",
          thread_id AS "threadId",
          actor_user_id AS "actorUserId",
          turn_id AS "turnId",
          goal_text AS "goalText",
          summary AS "summary",
          scope AS "scope",
          risk_class AS "riskClass",
          approval_mode AS "approvalMode",
          base_revision AS "baseRevision",
          created_at_unix_ms AS "createdAtUnixMs",
          context_json AS "contextJson",
          commands_json AS "commandsJson",
          affected_ranges_json AS "affectedRangesJson",
          estimated_affected_cells AS "estimatedAffectedCells",
          shared_review_json AS "sharedReviewJson"
        FROM workbook_pending_bundle
        WHERE workbook_id = $1 AND thread_id = $2 AND actor_user_id = $3
      `,
      [thread.workbookId, thread.threadId, thread.actorUserId],
    ),
  ]);
  const toolCallsByEntryId = new Map(
    toolCallResult.rows.flatMap((row) => {
      const normalized = normalizeToolCallRow(row);
      return normalized ? [[normalized.entryId, normalized] as const] : [];
    }),
  );
  const entries = itemResult.rows
    .map((row) =>
      normalizeTimelineEntry({
        ...row,
        ...(typeof row.entryId === "string" ? toolCallsByEntryId.get(row.entryId) : undefined),
      }),
    )
    .filter((entry): entry is WorkbookAgentTimelineEntry => entry !== null);
  const pendingBundle = normalizePendingBundle(pendingBundleResult.rows[0] ?? {});
  return {
    documentId: thread.workbookId,
    threadId: thread.threadId,
    actorUserId: thread.actorUserId,
    scope: thread.scope,
    executionPolicy: thread.executionPolicy,
    context: isWorkbookAgentUiContext(thread.contextJson) ? thread.contextJson : null,
    entries,
    pendingBundle,
    updatedAtUnixMs,
  };
}

export async function listWorkbookAgentThreadSummaries(
  db: Queryable,
  input: {
    documentId: string;
    actorUserId: string;
  },
): Promise<WorkbookAgentThreadSummary[]> {
  const result = await db.query<WorkbookChatThreadSummaryRow>(
    `
      SELECT
        thread.thread_id AS "threadId",
        thread.scope AS "scope",
        thread.actor_user_id AS "ownerUserId",
        thread.updated_at_unix_ms AS "updatedAtUnixMs",
        COALESCE(item_counts.entry_count, 0) AS "entryCount",
        pending.bundle_id IS NOT NULL AS "hasPendingBundle",
        latest_item.text AS "latestEntryText"
      FROM (
        SELECT ranked.workbook_id, ranked.thread_id, ranked.actor_user_id, ranked.scope, ranked.updated_at_unix_ms
        FROM (
          SELECT
            workbook_id,
            thread_id,
            actor_user_id,
            scope,
            updated_at_unix_ms,
            ROW_NUMBER() OVER (
              PARTITION BY thread_id
              ORDER BY
                CASE WHEN actor_user_id = $2 THEN 0 ELSE 1 END ASC,
                updated_at_unix_ms DESC
            ) AS row_rank
          FROM workbook_chat_thread
          WHERE workbook_id = $1
            AND (actor_user_id = $2 OR scope = 'shared')
        ) AS ranked
        WHERE ranked.row_rank = 1
      ) AS thread
      LEFT JOIN (
        SELECT workbook_id, thread_id, actor_user_id, COUNT(*)::integer AS entry_count
        FROM workbook_chat_item
        GROUP BY workbook_id, thread_id, actor_user_id
      ) AS item_counts
        ON item_counts.workbook_id = thread.workbook_id
       AND item_counts.thread_id = thread.thread_id
       AND item_counts.actor_user_id = thread.actor_user_id
      LEFT JOIN workbook_pending_bundle AS pending
        ON pending.workbook_id = thread.workbook_id
       AND pending.thread_id = thread.thread_id
       AND pending.actor_user_id = thread.actor_user_id
      LEFT JOIN LATERAL (
        SELECT text
        FROM workbook_chat_item
        WHERE workbook_id = thread.workbook_id
          AND thread_id = thread.thread_id
          AND actor_user_id = thread.actor_user_id
          AND text IS NOT NULL
        ORDER BY sort_order DESC
        LIMIT 1
      ) AS latest_item
        ON TRUE
      WHERE thread.workbook_id = $1
        AND (thread.actor_user_id = $2 OR thread.scope = 'shared')
      ORDER BY thread.updated_at_unix_ms DESC
    `,
    [input.documentId, input.actorUserId],
  );
  return result.rows
    .map((row) => normalizeThreadSummary(row))
    .filter((row): row is WorkbookAgentThreadSummary => row !== null);
}
