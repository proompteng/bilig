import type { WorkbookAgentWorkflowArtifact, WorkbookAgentWorkflowRun, WorkbookAgentWorkflowStep } from '@bilig/contracts'
import { queries } from '@bilig/zero-sync'
import type { Row } from '@rocicorp/zero'
import { addColumnIfMissing, ensureDefaultedNotNullColumn } from './schema-upgrade.js'
import type { QueryResultRow, Queryable, ZeroQueryRunner } from './store.js'
import { parseNonNegativeInteger } from './store-support.js'
import { runQueryableTransaction, runSequentially } from './transaction-support.js'

type ZeroWorkbookWorkflowRunRow = Row['workbook_workflow_run']

export interface WorkbookWorkflowRunStoreConnection extends Queryable {
  listWorkbookWorkflowRunRows(input: {
    readonly documentId: string
    readonly actorUserId: string
    readonly threadId: string
  }): Promise<readonly ZeroWorkbookWorkflowRunRow[]>
  listWorkbookWorkflowStepRows(input: {
    readonly documentId: string
    readonly actorUserId: string
    readonly threadId: string
  }): Promise<readonly QueryResultRow[]>
  listWorkbookWorkflowArtifactRows(input: {
    readonly documentId: string
    readonly actorUserId: string
    readonly threadId: string
  }): Promise<readonly QueryResultRow[]>
}

interface WorkbookWorkflowRunRow extends QueryResultRow {
  readonly runId?: unknown
  readonly workbookId?: unknown
  readonly threadId?: unknown
  readonly actorUserId?: unknown
  readonly workflowTemplate?: unknown
  readonly title?: unknown
  readonly summary?: unknown
  readonly status?: unknown
  readonly createdAtUnixMs?: unknown
  readonly updatedAtUnixMs?: unknown
  readonly completedAtUnixMs?: unknown
  readonly errorMessage?: unknown
}

interface WorkbookWorkflowStepRow extends QueryResultRow {
  readonly runId?: unknown
  readonly stepId?: unknown
  readonly stepOrder?: unknown
  readonly label?: unknown
  readonly status?: unknown
  readonly summary?: unknown
  readonly updatedAtUnixMs?: unknown
}

interface WorkbookWorkflowArtifactRow extends QueryResultRow {
  readonly runId?: unknown
  readonly kind?: unknown
  readonly title?: unknown
  readonly text?: unknown
}

interface DurableWorkflowArtifact {
  readonly artifact: WorkbookAgentWorkflowArtifact | null
  readonly invalid: boolean
}

interface DurableWorkflowStepRows {
  readonly byRunId: ReadonlyMap<string, WorkbookAgentWorkflowStep[] | null>
}

interface DurableWorkflowArtifactRows {
  readonly byRunId: ReadonlyMap<string, DurableWorkflowArtifact>
}

export function createWorkbookWorkflowRunStoreConnection(db: Queryable & ZeroQueryRunner): WorkbookWorkflowRunStoreConnection {
  return {
    query: (text, values) => db.query(text, values),
    listWorkbookWorkflowRunRows: ({ actorUserId, documentId, threadId }) =>
      db.run(queries.workbookWorkflowRun.byThread.fn({ args: { documentId, threadId }, ctx: { userID: actorUserId } })),
    listWorkbookWorkflowStepRows: ({ actorUserId, documentId, threadId }) =>
      db.run(queries.workbookWorkflowStep.byThread.fn({ args: { documentId, threadId }, ctx: { userID: actorUserId } })),
    listWorkbookWorkflowArtifactRows: ({ actorUserId, documentId, threadId }) =>
      db.run(queries.workbookWorkflowArtifact.byThread.fn({ args: { documentId, threadId }, ctx: { userID: actorUserId } })),
  }
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function isMarkdownArtifact(value: unknown): value is WorkbookAgentWorkflowArtifact {
  return (
    isRecord(value) &&
    'kind' in value &&
    value['kind'] === 'markdown' &&
    'title' in value &&
    typeof value['title'] === 'string' &&
    'text' in value &&
    typeof value['text'] === 'string'
  )
}

function isWorkflowTemplate(value: unknown): value is WorkbookAgentWorkflowRun['workflowTemplate'] {
  return (
    value === 'summarizeWorkbook' ||
    value === 'summarizeCurrentSheet' ||
    value === 'describeRecentChanges' ||
    value === 'findFormulaIssues' ||
    value === 'highlightFormulaIssues' ||
    value === 'repairFormulaIssues' ||
    value === 'highlightCurrentSheetOutliers' ||
    value === 'styleCurrentSheetHeaders' ||
    value === 'normalizeCurrentSheetHeaders' ||
    value === 'normalizeCurrentSheetNumberFormats' ||
    value === 'normalizeCurrentSheetWhitespace' ||
    value === 'fillCurrentSheetFormulasDown' ||
    value === 'traceSelectionDependencies' ||
    value === 'explainSelectionCell' ||
    value === 'searchWorkbookQuery' ||
    value === 'createCurrentSheetRollup' ||
    value === 'createCurrentSheetReviewTab' ||
    value === 'createSheet' ||
    value === 'renameCurrentSheet' ||
    value === 'hideCurrentRow' ||
    value === 'hideCurrentColumn' ||
    value === 'unhideCurrentRow' ||
    value === 'unhideCurrentColumn'
  )
}

function isWorkflowStepStatus(value: unknown): value is WorkbookAgentWorkflowStep['status'] {
  return value === 'pending' || value === 'running' || value === 'completed' || value === 'failed' || value === 'cancelled'
}

function normalizeWorkflowStepRow(row: WorkbookWorkflowStepRow): {
  readonly runId: string
  readonly step: WorkbookAgentWorkflowStep
  readonly stepOrder: number
} | null {
  const updatedAtUnixMs = parseNonNegativeInteger(row.updatedAtUnixMs)
  const stepOrder = parseNonNegativeInteger(row.stepOrder)
  if (
    typeof row.runId !== 'string' ||
    typeof row.stepId !== 'string' ||
    typeof row.label !== 'string' ||
    !isWorkflowStepStatus(row.status) ||
    typeof row.summary !== 'string' ||
    updatedAtUnixMs === null ||
    stepOrder === null
  ) {
    return null
  }
  return {
    runId: row.runId,
    stepOrder,
    step: {
      stepId: row.stepId,
      label: row.label,
      status: row.status,
      summary: row.summary,
      updatedAtUnixMs,
    },
  }
}

function normalizeWorkflowArtifactRow(row: WorkbookWorkflowArtifactRow): {
  readonly runId: string
  readonly artifact: WorkbookAgentWorkflowArtifact
} | null {
  const artifact = {
    kind: row.kind,
    title: row.title,
    text: row.text,
  }
  if (typeof row.runId !== 'string' || !isMarkdownArtifact(artifact)) {
    return null
  }
  return {
    runId: row.runId,
    artifact,
  }
}

function normalizeWorkflowRun(
  row: WorkbookWorkflowRunRow,
  hydrated: {
    readonly steps: WorkbookAgentWorkflowStep[] | null
    readonly artifact: WorkbookAgentWorkflowArtifact | null
    readonly artifactInvalid?: boolean
  },
): WorkbookAgentWorkflowRun | null {
  const createdAtUnixMs = parseNonNegativeInteger(row.createdAtUnixMs)
  const updatedAtUnixMs = parseNonNegativeInteger(row.updatedAtUnixMs)
  const completedAtUnixMs =
    row.completedAtUnixMs === null || row.completedAtUnixMs === undefined ? null : parseNonNegativeInteger(row.completedAtUnixMs)
  const { artifact, steps } = hydrated
  if (
    typeof row.runId !== 'string' ||
    typeof row.threadId !== 'string' ||
    typeof row.actorUserId !== 'string' ||
    !isWorkflowTemplate(row.workflowTemplate) ||
    typeof row.title !== 'string' ||
    typeof row.summary !== 'string' ||
    (row.status !== 'running' && row.status !== 'completed' && row.status !== 'failed' && row.status !== 'cancelled') ||
    createdAtUnixMs === null ||
    updatedAtUnixMs === null ||
    completedAtUnixMs === undefined ||
    updatedAtUnixMs < createdAtUnixMs ||
    (completedAtUnixMs !== null && completedAtUnixMs < createdAtUnixMs) ||
    steps === null ||
    hydrated.artifactInvalid === true ||
    (row.errorMessage !== null && row.errorMessage !== undefined && typeof row.errorMessage !== 'string') ||
    (artifact !== null && !isMarkdownArtifact(artifact))
  ) {
    return null
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
    errorMessage: typeof row.errorMessage === 'string' ? row.errorMessage : null,
    steps,
    artifact,
  }
}

function toWorkflowRunRow(row: ZeroWorkbookWorkflowRunRow): WorkbookWorkflowRunRow {
  return {
    runId: row.runId,
    workbookId: row.workbookId,
    threadId: row.threadId,
    actorUserId: row.startedByUserId,
    workflowTemplate: row.workflowTemplate,
    title: row.title,
    summary: row.summary,
    status: row.status,
    createdAtUnixMs: row.createdAtUnixMs,
    updatedAtUnixMs: row.updatedAtUnixMs,
    completedAtUnixMs: row.completedAtUnixMs,
    errorMessage: row.errorMessage,
  }
}

function toWorkflowStepRow(row: QueryResultRow): WorkbookWorkflowStepRow {
  return {
    runId: row['runId'],
    stepId: row['stepId'],
    stepOrder: row['stepOrder'],
    label: row['label'],
    status: row['status'],
    summary: row['summary'],
    updatedAtUnixMs: row['updatedAtUnixMs'],
  }
}

function toWorkflowArtifactRow(row: QueryResultRow): WorkbookWorkflowArtifactRow {
  return {
    runId: row['runId'],
    kind: row['kind'],
    title: row['title'],
    text: row['text'],
  }
}

function loadDurableWorkflowSteps(rows: readonly QueryResultRow[], runIds: ReadonlySet<string>): DurableWorkflowStepRows {
  if (runIds.size === 0) {
    return {
      byRunId: new Map(),
    }
  }
  const stepsByRunId = new Map<string, WorkbookAgentWorkflowStep[] | null>()
  for (const row of rows) {
    const normalized = normalizeWorkflowStepRow(toWorkflowStepRow(row))
    if (!normalized) {
      if (typeof row['runId'] === 'string') {
        stepsByRunId.set(row['runId'], null)
      }
      continue
    }
    if (!runIds.has(normalized.runId)) {
      continue
    }
    if (stepsByRunId.get(normalized.runId) === null) {
      continue
    }
    const steps = stepsByRunId.get(normalized.runId)
    if (steps) {
      steps.push(normalized.step)
      continue
    }
    stepsByRunId.set(normalized.runId, [normalized.step])
  }
  return {
    byRunId: stepsByRunId,
  }
}

function loadDurableWorkflowArtifacts(rows: readonly QueryResultRow[], runIds: ReadonlySet<string>): DurableWorkflowArtifactRows {
  if (runIds.size === 0) {
    return {
      byRunId: new Map(),
    }
  }
  const artifactsByRunId = new Map<string, DurableWorkflowArtifact>()
  for (const row of rows) {
    const normalized = normalizeWorkflowArtifactRow(toWorkflowArtifactRow(row))
    if (!normalized) {
      if (typeof row['runId'] === 'string') {
        artifactsByRunId.set(row['runId'], {
          artifact: null,
          invalid: true,
        })
      }
      continue
    }
    if (!runIds.has(normalized.runId)) {
      continue
    }
    artifactsByRunId.set(normalized.runId, {
      artifact: normalized.artifact,
      invalid: false,
    })
  }
  return {
    byRunId: artifactsByRunId,
  }
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
  `)
  await db.query(`
    CREATE TABLE IF NOT EXISTS workbook_workflow_step (
      workbook_id TEXT NOT NULL,
      run_id TEXT NOT NULL,
      step_id TEXT NOT NULL,
      step_order INTEGER NOT NULL,
      label TEXT NOT NULL,
      status TEXT NOT NULL,
      summary TEXT NOT NULL,
      updated_at_unix_ms BIGINT NOT NULL,
      PRIMARY KEY (run_id, step_id)
    )
  `)
  await db.query(`
    CREATE TABLE IF NOT EXISTS workbook_workflow_artifact (
      run_id TEXT PRIMARY KEY,
      workbook_id TEXT NOT NULL,
      kind TEXT NOT NULL,
      title TEXT NOT NULL,
      text TEXT NOT NULL,
      updated_at_unix_ms BIGINT NOT NULL
    )
  `)
  await ensureDefaultedNotNullColumn(db, {
    tableName: 'workbook_workflow_run',
    columnName: 'steps_json',
    dataType: 'JSONB',
    defaultSql: "'[]'::jsonb",
  })
  await addColumnIfMissing(db, { tableName: 'workbook_workflow_run', columnName: 'completed_at_unix_ms', dataType: 'BIGINT' })
  await addColumnIfMissing(db, { tableName: 'workbook_workflow_run', columnName: 'error_message', dataType: 'TEXT' })
  await addColumnIfMissing(db, { tableName: 'workbook_workflow_run', columnName: 'artifact_json', dataType: 'JSONB' })
  await db.query(`
    ALTER TABLE workbook_workflow_artifact
      ADD COLUMN IF NOT EXISTS workbook_id TEXT
  `)
  await db.query(`
    UPDATE workbook_workflow_artifact AS artifact
    SET workbook_id = run.workbook_id
    FROM workbook_workflow_run AS run
    WHERE artifact.run_id = run.run_id
      AND (artifact.workbook_id IS NULL OR artifact.workbook_id = '')
  `)
  await db.query(`
    ALTER TABLE workbook_workflow_artifact
      ALTER COLUMN workbook_id SET NOT NULL
  `)
  await ensureDefaultedNotNullColumn(db, {
    tableName: 'workbook_workflow_artifact',
    columnName: 'updated_at_unix_ms',
    dataType: 'BIGINT',
    defaultSql: '0',
  })
  await db.query(`
    INSERT INTO workbook_workflow_step (
      workbook_id,
      run_id,
      step_id,
      step_order,
      label,
      status,
      summary,
      updated_at_unix_ms
    )
    SELECT
      run.workbook_id,
      run.run_id,
      step_item.step->>'stepId',
      (step_item.ordinality - 1)::integer,
      step_item.step->>'label',
      step_item.step->>'status',
      step_item.step->>'summary',
      (step_item.step->>'updatedAtUnixMs')::bigint
    FROM workbook_workflow_run AS run
    CROSS JOIN LATERAL jsonb_array_elements(
      CASE WHEN jsonb_typeof(run.steps_json) = 'array' THEN run.steps_json ELSE '[]'::jsonb END
    ) WITH ORDINALITY AS step_item(step, ordinality)
    WHERE jsonb_typeof(run.steps_json) = 'array'
      AND step_item.step->>'stepId' IS NOT NULL
      AND step_item.step->>'label' IS NOT NULL
      AND step_item.step->>'status' IN ('pending', 'running', 'completed', 'failed', 'cancelled')
      AND step_item.step->>'summary' IS NOT NULL
      AND step_item.step->>'updatedAtUnixMs' ~ '^[0-9]+$'
    ON CONFLICT (run_id, step_id) DO NOTHING
  `)
  await db.query(`
    INSERT INTO workbook_workflow_artifact (
      run_id,
      workbook_id,
      kind,
      title,
      text,
      updated_at_unix_ms
    )
    SELECT
      run.run_id,
      run.workbook_id,
      'markdown',
      run.artifact_json->>'title',
      run.artifact_json->>'text',
      run.updated_at_unix_ms
    FROM workbook_workflow_run AS run
    WHERE run.artifact_json->>'kind' = 'markdown'
      AND run.artifact_json->>'title' IS NOT NULL
      AND run.artifact_json->>'text' IS NOT NULL
    ON CONFLICT (run_id) DO NOTHING
  `)
  await db.query(`
    CREATE INDEX IF NOT EXISTS workbook_workflow_run_thread_updated_idx
      ON workbook_workflow_run (workbook_id, thread_id, updated_at_unix_ms DESC)
  `)
  await db.query(`
    CREATE INDEX IF NOT EXISTS workbook_workflow_step_run_order_idx
      ON workbook_workflow_step (workbook_id, run_id, step_order ASC)
  `)
  await db.query(`
    CREATE INDEX IF NOT EXISTS workbook_workflow_artifact_run_idx
      ON workbook_workflow_artifact (workbook_id, run_id)
  `)
}

export async function upsertWorkbookWorkflowRun(
  db: Queryable,
  input: {
    documentId: string
    run: WorkbookAgentWorkflowRun
  },
): Promise<void> {
  await runQueryableTransaction(db, async (transactionDb) => {
    await persistWorkbookWorkflowRun(transactionDb, input)
  })
}

async function persistWorkbookWorkflowRun(
  db: Queryable,
  input: {
    documentId: string
    run: WorkbookAgentWorkflowRun
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
        error_message
      )
      VALUES (
        $1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11, $12
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
        error_message = EXCLUDED.error_message
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
    ],
  )
  await db.query(
    `
      DELETE FROM workbook_workflow_step
      WHERE run_id = $1
    `,
    [input.run.runId],
  )
  await runSequentially(input.run.steps, async (step, stepOrder) => {
    await db.query(
      `
          INSERT INTO workbook_workflow_step (
            workbook_id,
            run_id,
            step_id,
            step_order,
            label,
            status,
            summary,
            updated_at_unix_ms
          )
          VALUES ($1, $2, $3, $4, $5, $6, $7, $8)
        `,
      [input.documentId, input.run.runId, step.stepId, stepOrder, step.label, step.status, step.summary, step.updatedAtUnixMs],
    )
  })
  if (input.run.artifact) {
    await db.query(
      `
        INSERT INTO workbook_workflow_artifact (
          run_id,
          workbook_id,
          kind,
          title,
          text,
          updated_at_unix_ms
        )
        VALUES ($1, $2, $3, $4, $5, $6)
        ON CONFLICT (run_id)
        DO UPDATE SET
          workbook_id = EXCLUDED.workbook_id,
          kind = EXCLUDED.kind,
          title = EXCLUDED.title,
          text = EXCLUDED.text,
          updated_at_unix_ms = EXCLUDED.updated_at_unix_ms
      `,
      [
        input.run.runId,
        input.documentId,
        input.run.artifact.kind,
        input.run.artifact.title,
        input.run.artifact.text,
        input.run.updatedAtUnixMs,
      ],
    )
    return
  }
  await db.query(
    `
      DELETE FROM workbook_workflow_artifact
      WHERE run_id = $1
    `,
    [input.run.runId],
  )
}

export async function listWorkbookThreadWorkflowRuns(
  db: WorkbookWorkflowRunStoreConnection,
  input: {
    documentId: string
    actorUserId: string
    threadId: string
    limit?: number
  },
): Promise<WorkbookAgentWorkflowRun[]> {
  const rows = (await db.listWorkbookWorkflowRunRows(input)).map(toWorkflowRunRow)
  const runIds = rows.flatMap((row) => (typeof row.runId === 'string' ? [row.runId] : []))
  const selectedRunIds = new Set(runIds)
  const [stepRows, artifactRows] = await Promise.all([db.listWorkbookWorkflowStepRows(input), db.listWorkbookWorkflowArtifactRows(input)])
  const durableSteps = loadDurableWorkflowSteps(stepRows, selectedRunIds)
  const durableArtifacts = loadDurableWorkflowArtifacts(artifactRows, selectedRunIds)
  const limit = input.limit ?? 20
  const runs: WorkbookAgentWorkflowRun[] = []
  for (const row of rows) {
    const runId = typeof row.runId === 'string' ? row.runId : null
    const hydrated: {
      steps: WorkbookAgentWorkflowStep[] | null
      artifact: WorkbookAgentWorkflowArtifact | null
      artifactInvalid?: boolean
    } = {
      steps: null,
      artifact: null,
    }
    if (runId) {
      hydrated.steps = durableSteps.byRunId.get(runId) ?? null
      const durableArtifact = durableArtifacts.byRunId.get(runId)
      if (durableArtifact) {
        hydrated.artifact = durableArtifact.artifact
        hydrated.artifactInvalid = durableArtifact.invalid
      }
    }
    const run = normalizeWorkflowRun(row, hydrated)
    if (run) {
      runs.push(run)
      if (runs.length >= limit) {
        break
      }
    }
  }
  return runs
}
