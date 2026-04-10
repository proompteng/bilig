import type {
  WorkbookAgentWorkflowArtifact,
  WorkbookAgentWorkflowRun,
  WorkbookAgentWorkflowStep,
} from "@bilig/contracts";
import type { QueryResultRow, Queryable } from "./store.js";

interface WorkbookWorkflowRunRow extends QueryResultRow {
  readonly runId?: unknown;
  readonly workbookId?: unknown;
  readonly threadId?: unknown;
  readonly actorUserId?: unknown;
  readonly workflowTemplate?: unknown;
  readonly title?: unknown;
  readonly summary?: unknown;
  readonly status?: unknown;
  readonly createdAtUnixMs?: unknown;
  readonly updatedAtUnixMs?: unknown;
  readonly completedAtUnixMs?: unknown;
  readonly errorMessage?: unknown;
  readonly stepsJson?: unknown;
  readonly artifactJson?: unknown;
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

function isMarkdownArtifact(value: unknown): value is WorkbookAgentWorkflowArtifact {
  return (
    typeof value === "object" &&
    value !== null &&
    "kind" in value &&
    value.kind === "markdown" &&
    "title" in value &&
    typeof value.title === "string" &&
    "text" in value &&
    typeof value.text === "string"
  );
}

function isWorkflowTemplate(value: unknown): value is WorkbookAgentWorkflowRun["workflowTemplate"] {
  return (
    value === "summarizeWorkbook" ||
    value === "describeRecentChanges" ||
    value === "findFormulaIssues"
  );
}

function isWorkflowStepStatus(value: unknown): value is WorkbookAgentWorkflowStep["status"] {
  return (
    value === "pending" || value === "running" || value === "completed" || value === "failed"
  );
}

function isWorkflowStep(value: unknown): value is WorkbookAgentWorkflowStep {
  return (
    typeof value === "object" &&
    value !== null &&
    "stepId" in value &&
    typeof value.stepId === "string" &&
    "label" in value &&
    typeof value.label === "string" &&
    "status" in value &&
    isWorkflowStepStatus(value.status) &&
    "summary" in value &&
    typeof value.summary === "string" &&
    "updatedAtUnixMs" in value &&
    parseNumericValue(value.updatedAtUnixMs) !== null
  );
}

function normalizeWorkflowSteps(value: unknown): WorkbookAgentWorkflowStep[] | null {
  if (!Array.isArray(value) || !value.every((entry) => isWorkflowStep(entry))) {
    return null;
  }
  return value.map((entry) => ({
    stepId: entry.stepId,
    label: entry.label,
    status: entry.status,
    summary: entry.summary,
    updatedAtUnixMs: parseNumericValue(entry.updatedAtUnixMs) ?? 0,
  }));
}

function normalizeWorkflowRun(row: WorkbookWorkflowRunRow): WorkbookAgentWorkflowRun | null {
  const createdAtUnixMs = parseNumericValue(row.createdAtUnixMs);
  const updatedAtUnixMs = parseNumericValue(row.updatedAtUnixMs);
  const completedAtUnixMs =
    row.completedAtUnixMs === null || row.completedAtUnixMs === undefined
      ? null
      : parseNumericValue(row.completedAtUnixMs);
  const steps = normalizeWorkflowSteps(row.stepsJson ?? []);
  if (
    typeof row.runId !== "string" ||
    typeof row.threadId !== "string" ||
    typeof row.actorUserId !== "string" ||
    !isWorkflowTemplate(row.workflowTemplate) ||
    typeof row.title !== "string" ||
    typeof row.summary !== "string" ||
    (row.status !== "running" && row.status !== "completed" && row.status !== "failed") ||
    createdAtUnixMs === null ||
    updatedAtUnixMs === null ||
    completedAtUnixMs === undefined ||
    steps === null ||
    (row.errorMessage !== null &&
      row.errorMessage !== undefined &&
      typeof row.errorMessage !== "string") ||
    (row.artifactJson !== null &&
      row.artifactJson !== undefined &&
      !isMarkdownArtifact(row.artifactJson))
  ) {
    return null;
  }
  return {
    runId: row.runId,
    threadId: row.threadId,
    startedByUserId: row.actorUserId,
    workflowTemplate: row.workflowTemplate,
    title: row.title,
    summary: row.summary,
    status: row.status,
    createdAtUnixMs,
    updatedAtUnixMs,
    completedAtUnixMs,
    errorMessage: typeof row.errorMessage === "string" ? row.errorMessage : null,
    steps,
    artifact: isMarkdownArtifact(row.artifactJson) ? row.artifactJson : null,
  };
}

export async function ensureWorkbookWorkflowRunSchema(db: Queryable): Promise<void> {
  await db.query(`
    CREATE TABLE IF NOT EXISTS workbook_workflow_run (
      run_id TEXT PRIMARY KEY,
      workbook_id TEXT NOT NULL,
      thread_id TEXT NOT NULL,
      actor_user_id TEXT NOT NULL,
      workflow_template TEXT NOT NULL,
      title TEXT NOT NULL,
      summary TEXT NOT NULL,
      status TEXT NOT NULL,
      created_at_unix_ms BIGINT NOT NULL,
      updated_at_unix_ms BIGINT NOT NULL,
      completed_at_unix_ms BIGINT,
      error_message TEXT,
      steps_json JSONB NOT NULL DEFAULT '[]'::jsonb,
      artifact_json JSONB
    )
  `);
  await db.query(`
    ALTER TABLE workbook_workflow_run
      ADD COLUMN IF NOT EXISTS steps_json JSONB NOT NULL DEFAULT '[]'::jsonb
  `);
  await db.query(`
    CREATE INDEX IF NOT EXISTS workbook_workflow_run_thread_updated_idx
      ON workbook_workflow_run (workbook_id, thread_id, updated_at_unix_ms DESC)
  `);
}

export async function upsertWorkbookWorkflowRun(
  db: Queryable,
  input: {
    documentId: string;
    run: WorkbookAgentWorkflowRun;
  },
): Promise<void> {
  await db.query(
    `
      INSERT INTO workbook_workflow_run (
        run_id,
        workbook_id,
        thread_id,
        actor_user_id,
        workflow_template,
        title,
        summary,
        status,
        created_at_unix_ms,
        updated_at_unix_ms,
        completed_at_unix_ms,
        error_message,
        steps_json,
        artifact_json
      )
      VALUES (
        $1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11, $12, $13::jsonb, $14::jsonb
      )
      ON CONFLICT (run_id)
      DO UPDATE SET
        workbook_id = EXCLUDED.workbook_id,
        thread_id = EXCLUDED.thread_id,
        actor_user_id = EXCLUDED.actor_user_id,
        workflow_template = EXCLUDED.workflow_template,
        title = EXCLUDED.title,
        summary = EXCLUDED.summary,
        status = EXCLUDED.status,
        created_at_unix_ms = EXCLUDED.created_at_unix_ms,
        updated_at_unix_ms = EXCLUDED.updated_at_unix_ms,
        completed_at_unix_ms = EXCLUDED.completed_at_unix_ms,
        error_message = EXCLUDED.error_message,
        steps_json = EXCLUDED.steps_json,
        artifact_json = EXCLUDED.artifact_json
    `,
    [
      input.run.runId,
      input.documentId,
      input.run.threadId,
      input.run.startedByUserId,
      input.run.workflowTemplate,
      input.run.title,
      input.run.summary,
      input.run.status,
      input.run.createdAtUnixMs,
      input.run.updatedAtUnixMs,
      input.run.completedAtUnixMs,
      input.run.errorMessage,
      JSON.stringify(input.run.steps),
      JSON.stringify(input.run.artifact),
    ],
  );
}

export async function listWorkbookThreadWorkflowRuns(
  db: Queryable,
  input: {
    documentId: string;
    actorUserId: string;
    threadId: string;
    limit?: number;
  },
): Promise<WorkbookAgentWorkflowRun[]> {
  const result = await db.query<WorkbookWorkflowRunRow>(
    `
      SELECT
        run.run_id AS "runId",
        run.workbook_id AS "workbookId",
        run.thread_id AS "threadId",
        run.actor_user_id AS "actorUserId",
        run.workflow_template AS "workflowTemplate",
        run.title AS "title",
        run.summary AS "summary",
        run.status AS "status",
        run.created_at_unix_ms AS "createdAtUnixMs",
        run.updated_at_unix_ms AS "updatedAtUnixMs",
        run.completed_at_unix_ms AS "completedAtUnixMs",
        run.error_message AS "errorMessage",
        run.steps_json AS "stepsJson",
        run.artifact_json AS "artifactJson"
      FROM workbook_workflow_run AS run
      LEFT JOIN workbook_chat_thread AS thread
        ON thread.workbook_id = run.workbook_id
       AND thread.thread_id = run.thread_id
       AND thread.actor_user_id = run.actor_user_id
      WHERE run.workbook_id = $1
        AND run.thread_id = $2
        AND (run.actor_user_id = $3 OR thread.scope = 'shared')
      ORDER BY run.updated_at_unix_ms DESC, run.run_id DESC
      LIMIT $4
    `,
    [input.documentId, input.threadId, input.actorUserId, input.limit ?? 20],
  );
  return result.rows.flatMap((row) => {
    const run = normalizeWorkflowRun(row);
    return run ? [run] : [];
  });
}
