import { describe, expect, it } from 'vitest'
import {
  ensureWorkbookWorkflowRunSchema,
  listWorkbookThreadWorkflowRuns,
  upsertWorkbookWorkflowRun,
} from '../workbook-workflow-run-store.js'
import type { QueryResultRow, Queryable } from '../store.js'
import type { Row } from '@rocicorp/zero'

type ZeroWorkflowRunRow = Row['workbook_workflow_run']
type ZeroWorkflowStepRow = Row['workbook_workflow_step']
type ZeroWorkflowArtifactRow = Row['workbook_workflow_artifact']

const NON_REPLICATED_WORKFLOW_MUTATION_COLUMNS = [
  'mutation_executed',
  'verification_complete',
  'mutation_status',
  'mutation_receipt_json',
] as const

interface RecordedQuery {
  readonly text: string
  readonly values: readonly unknown[] | undefined
}

class FakeQueryable implements Queryable {
  readonly calls: RecordedQuery[] = []

  constructor(
    private readonly responders: readonly ((text: string, values: readonly unknown[] | undefined) => QueryResultRow[] | null)[] = [],
  ) {}

  async query<T extends QueryResultRow = QueryResultRow>(text: string, values?: unknown[]): Promise<{ rows: T[] }> {
    this.calls.push({ text, values })
    for (const responder of this.responders) {
      const rows = responder(text, values)
      if (rows) {
        return { rows: rows.filter((row): row is T => row !== null) }
      }
    }
    return { rows: [] }
  }
}

class FakeTransactionClient implements Queryable {
  readonly calls: RecordedQuery[] = []
  releaseCount = 0

  constructor(private readonly failOnText: string | null = null) {}

  async query<T extends QueryResultRow = QueryResultRow>(text: string, values?: unknown[]): Promise<{ rows: T[] }> {
    this.calls.push({ text, values })
    if (this.failOnText && text.includes(this.failOnText)) {
      throw new Error(`failed query: ${this.failOnText}`)
    }
    return { rows: [] }
  }

  release(): void {
    this.releaseCount += 1
  }
}

class FakeTransactionalQueryable implements Queryable {
  readonly calls: RecordedQuery[] = []
  readonly client: FakeTransactionClient
  connectCount = 0

  constructor(failOnText: string | null = null) {
    this.client = new FakeTransactionClient(failOnText)
  }

  async query<T extends QueryResultRow = QueryResultRow>(text: string, values?: unknown[]): Promise<{ rows: T[] }> {
    this.calls.push({ text, values })
    return { rows: [] as T[] }
  }

  async connect(): Promise<FakeTransactionClient> {
    this.connectCount += 1
    return this.client
  }
}

function createWorkflowRun() {
  return {
    runId: 'workflow-1',
    threadId: 'thr-1',
    startedByUserId: 'alex@example.com',
    workflowTemplate: 'summarizeWorkbook' as const,
    title: 'Summarize Workbook',
    summary: 'Summarized workbook structure across 2 sheets.',
    status: 'completed' as const,
    createdAtUnixMs: 100,
    updatedAtUnixMs: 120,
    completedAtUnixMs: 120,
    errorMessage: null,
    steps: [
      {
        stepId: 'inspect-workbook',
        label: 'Inspect workbook structure',
        status: 'completed' as const,
        summary: 'Read durable workbook structure across 2 sheets.',
        updatedAtUnixMs: 110,
      },
      {
        stepId: 'draft-summary',
        label: 'Draft summary artifact',
        status: 'completed' as const,
        summary: 'Prepared the durable workbook summary artifact for the thread.',
        updatedAtUnixMs: 120,
      },
    ],
    artifact: {
      kind: 'markdown' as const,
      title: 'Workbook Summary',
      text: '## Summary',
    },
  }
}

interface WorkflowRunFixture {
  readonly runId: string
  readonly threadId: string
  readonly startedByUserId: string
  readonly workflowTemplate: ZeroWorkflowRunRow['workflowTemplate']
  readonly title: string
  readonly summary: string
  readonly status: ZeroWorkflowRunRow['status']
  readonly createdAtUnixMs: number
  readonly updatedAtUnixMs: number
  readonly completedAtUnixMs: number | null
  readonly errorMessage: string | null
  readonly steps: ZeroWorkflowRunRow['steps']
  readonly artifact: ZeroWorkflowRunRow['artifact']
}

function createZeroWorkflowRunRow(
  run: WorkflowRunFixture,
  overrides: Partial<ZeroWorkflowRunRow> = {},
  documentId = 'doc-1',
): ZeroWorkflowRunRow {
  return {
    runId: run.runId,
    workbookId: documentId,
    threadId: run.threadId,
    startedByUserId: run.startedByUserId,
    workflowTemplate: run.workflowTemplate,
    title: run.title,
    summary: run.summary,
    status: run.status,
    createdAtUnixMs: run.createdAtUnixMs,
    updatedAtUnixMs: run.updatedAtUnixMs,
    completedAtUnixMs: run.completedAtUnixMs,
    errorMessage: run.errorMessage,
    steps: run.steps,
    artifact: run.artifact,
    ...overrides,
  }
}

function createWorkflowRunStoreConnection(
  workflowRows: readonly ZeroWorkflowRunRow[],
  responders: readonly ((text: string, values: readonly unknown[] | undefined) => QueryResultRow[] | null)[] = [],
  childRows: {
    readonly stepRows?: readonly QueryResultRow[]
    readonly artifactRows?: readonly QueryResultRow[]
    readonly mutationProofRows?: readonly QueryResultRow[]
  } = {},
) {
  const queryable = new FakeQueryable(responders)
  const workflowRunInputs: {
    readonly documentId: string
    readonly actorUserId: string
    readonly threadId: string
  }[] = []
  const workflowStepInputs: {
    readonly documentId: string
    readonly actorUserId: string
    readonly threadId: string
  }[] = []
  const workflowArtifactInputs: {
    readonly documentId: string
    readonly actorUserId: string
    readonly threadId: string
  }[] = []
  const workflowMutationProofInputs: {
    readonly documentId: string
    readonly actorUserId: string
    readonly threadId: string
  }[] = []
  return Object.assign(queryable, {
    workflowRunInputs,
    workflowStepInputs,
    workflowArtifactInputs,
    workflowMutationProofInputs,
    async listWorkbookWorkflowRunRows(input: { readonly documentId: string; readonly actorUserId: string; readonly threadId: string }) {
      workflowRunInputs.push(input)
      return workflowRows
    },
    async listWorkbookWorkflowMutationProofRows(input: {
      readonly documentId: string
      readonly actorUserId: string
      readonly threadId: string
    }) {
      workflowMutationProofInputs.push(input)
      return childRows.mutationProofRows ?? []
    },
    async listWorkbookWorkflowStepRows(input: { readonly documentId: string; readonly actorUserId: string; readonly threadId: string }) {
      workflowStepInputs.push(input)
      return childRows.stepRows ?? []
    },
    async listWorkbookWorkflowArtifactRows(input: {
      readonly documentId: string
      readonly actorUserId: string
      readonly threadId: string
    }) {
      workflowArtifactInputs.push(input)
      return childRows.artifactRows ?? []
    },
  })
}

function createZeroWorkflowStepRows(run: ReturnType<typeof createWorkflowRun>): ZeroWorkflowStepRow[] {
  return run.steps.map((step, index) => ({
    workbookId: 'doc-1',
    runId: run.runId,
    stepId: step.stepId,
    stepOrder: index,
    label: step.label,
    status: step.status,
    summary: step.summary,
    updatedAtUnixMs: step.updatedAtUnixMs,
  }))
}

function createZeroWorkflowArtifactRow(run: ReturnType<typeof createWorkflowRun>): ZeroWorkflowArtifactRow {
  if (!run.artifact) {
    throw new Error('Expected workflow run artifact')
  }
  return {
    runId: run.runId,
    workbookId: 'doc-1',
    kind: run.artifact.kind,
    title: run.artifact.title,
    text: run.artifact.text,
    updatedAtUnixMs: run.updatedAtUnixMs,
  }
}

function createWorkflowMutationProofRun(overrides: Partial<ReturnType<typeof createWorkflowRun>> = {}) {
  return {
    ...createWorkflowRun(),
    summary: 'Workflow mutation verification incomplete at revision r2: Rendered readback did not prove the mutation.',
    mutationExecuted: true,
    verificationComplete: false,
    mutationStatus: 'verification_incomplete' as const,
    mutationReceipt: {
      status: 'verification_incomplete',
      revision: { before: 1, after: 2 },
      renderedReadback: {
        matched: false,
        incompleteReason: 'Rendered readback did not prove the mutation.',
      },
    },
    ...overrides,
  }
}

function createZeroWorkflowMutationProofRow(run: ReturnType<typeof createWorkflowMutationProofRun>): QueryResultRow {
  return {
    runId: run.runId,
    workbookId: 'doc-1',
    mutationExecuted: run.mutationExecuted,
    verificationComplete: run.verificationComplete,
    mutationStatus: run.mutationStatus,
    mutationReceipt: run.mutationReceipt,
    updatedAtUnixMs: run.updatedAtUnixMs,
  }
}

describe('workbook-workflow-run-store', () => {
  it('does not create workflow mutation proof columns outside the Zero replication contract', async () => {
    const queryable = new FakeQueryable()

    await ensureWorkbookWorkflowRunSchema(queryable)

    for (const column of NON_REPLICATED_WORKFLOW_MUTATION_COLUMNS) {
      expect(
        queryable.calls.some(
          (call) => call.text.includes('ALTER TABLE workbook_workflow_run') && call.text.includes(`ADD COLUMN IF NOT EXISTS ${column}`),
        ),
      ).toBe(false)
    }
    expect(queryable.calls.some((call) => call.text.includes('CREATE TABLE IF NOT EXISTS workbook_workflow_mutation_proof'))).toBe(true)
  })

  it('persists workflow artifacts in durable rows', async () => {
    const queryable = new FakeQueryable()

    await upsertWorkbookWorkflowRun(queryable, {
      documentId: 'doc-1',
      run: createWorkflowRun(),
    })

    const insertQuery = queryable.calls.find((call) => call.text.includes('INSERT INTO workbook_workflow_run'))
    expect(insertQuery?.text).not.toContain('steps_json')
    expect(insertQuery?.text).not.toContain('artifact_json')
    for (const column of NON_REPLICATED_WORKFLOW_MUTATION_COLUMNS) {
      expect(insertQuery?.text).not.toContain(column)
    }
    expect(insertQuery?.values).not.toContain(JSON.stringify(createWorkflowRun().steps))
    expect(insertQuery?.values).not.toContain(JSON.stringify(createWorkflowRun().artifact))
    expect(queryable.calls.find((call) => call.text.includes('DELETE FROM workbook_workflow_step'))).toBeDefined()
    expect(queryable.calls.filter((call) => call.text.includes('INSERT INTO workbook_workflow_step'))).toHaveLength(
      createWorkflowRun().steps.length,
    )
    expect(queryable.calls.find((call) => call.text.includes('INSERT INTO workbook_workflow_artifact'))).toBeDefined()
  })

  it('persists workflow mutation proof in a durable child row instead of parent Zero columns', async () => {
    const queryable = new FakeQueryable()
    const run = createWorkflowMutationProofRun()

    await upsertWorkbookWorkflowRun(queryable, {
      documentId: 'doc-1',
      run,
    })

    const parentInsert = queryable.calls.find((call) => call.text.includes('INSERT INTO workbook_workflow_run'))
    for (const column of NON_REPLICATED_WORKFLOW_MUTATION_COLUMNS) {
      expect(parentInsert?.text).not.toContain(column)
    }
    const proofInsert = queryable.calls.find((call) => call.text.includes('INSERT INTO workbook_workflow_mutation_proof'))
    expect(proofInsert?.text).toContain('mutation_receipt_json')
    expect(proofInsert?.values).toEqual([
      run.runId,
      'doc-1',
      true,
      false,
      'verification_incomplete',
      JSON.stringify(run.mutationReceipt),
      run.updatedAtUnixMs,
    ])
  })

  it('deletes durable workflow mutation proof when the latest snapshot has no proof', async () => {
    const queryable = new FakeQueryable()

    await upsertWorkbookWorkflowRun(queryable, {
      documentId: 'doc-1',
      run: createWorkflowRun({ mutationExecuted: null, verificationComplete: null, mutationStatus: null, mutationReceipt: null }),
    })

    const deleteQuery = queryable.calls.find((call) => call.text.includes('DELETE FROM workbook_workflow_mutation_proof'))
    expect(deleteQuery?.text).toContain('WHERE run_id = $1')
    expect(deleteQuery?.values).toEqual(['workflow-1'])
  })

  it('replaces durable workflow steps by global run id before inserting the latest snapshot', async () => {
    const queryable = new FakeQueryable()
    const run = createWorkflowRun()

    await upsertWorkbookWorkflowRun(queryable, {
      documentId: 'doc-2',
      run,
    })

    const deleteQuery = queryable.calls.find((call) => call.text.includes('DELETE FROM workbook_workflow_step'))
    expect(deleteQuery?.text).toContain('WHERE run_id = $1')
    expect(deleteQuery?.text).not.toContain('workbook_id')
    expect(deleteQuery?.values).toEqual([run.runId])
  })

  it('deletes durable workflow artifacts by global run id when a snapshot no longer has an artifact', async () => {
    const queryable = new FakeQueryable()
    const run = {
      ...createWorkflowRun(),
      artifact: null,
    }

    await upsertWorkbookWorkflowRun(queryable, {
      documentId: 'doc-2',
      run,
    })

    const deleteQuery = queryable.calls.find((call) => call.text.includes('DELETE FROM workbook_workflow_artifact'))
    expect(deleteQuery?.text).toContain('WHERE run_id = $1')
    expect(deleteQuery?.text).not.toContain('workbook_id')
    expect(deleteQuery?.values).toEqual([run.runId])
  })

  it('saves workflow run snapshots atomically when the queryable supports transactions', async () => {
    const queryable = new FakeTransactionalQueryable()

    await upsertWorkbookWorkflowRun(queryable, {
      documentId: 'doc-1',
      run: createWorkflowRun(),
    })

    expect(queryable.connectCount).toBe(1)
    expect(queryable.calls).toEqual([])
    expect(queryable.client.releaseCount).toBe(1)
    expect(queryable.client.calls[0]?.text).toBe('BEGIN')
    expect(queryable.client.calls.at(-1)?.text).toBe('COMMIT')
    expect(queryable.client.calls.some((call) => call.text.includes('INSERT INTO workbook_workflow_run'))).toBe(true)
    expect(queryable.client.calls.some((call) => call.text.includes('DELETE FROM workbook_workflow_step'))).toBe(true)
    expect(queryable.client.calls.filter((call) => call.text.includes('INSERT INTO workbook_workflow_step'))).toHaveLength(
      createWorkflowRun().steps.length,
    )
    expect(queryable.client.calls.some((call) => call.text.includes('INSERT INTO workbook_workflow_artifact'))).toBe(true)
  })

  it('rolls back workflow run snapshots when a transactional child write fails', async () => {
    const queryable = new FakeTransactionalQueryable('INSERT INTO workbook_workflow_artifact')

    await expect(
      upsertWorkbookWorkflowRun(queryable, {
        documentId: 'doc-1',
        run: createWorkflowRun(),
      }),
    ).rejects.toThrow('failed query: INSERT INTO workbook_workflow_artifact')

    expect(queryable.connectCount).toBe(1)
    expect(queryable.client.releaseCount).toBe(1)
    expect(queryable.client.calls[0]?.text).toBe('BEGIN')
    expect(queryable.client.calls.at(-1)?.text).toBe('ROLLBACK')
    expect(queryable.client.calls.some((call) => call.text.includes('COMMIT'))).toBe(false)
  })

  it('loads visible workflow run rows through the shared Zero query model', async () => {
    const run = createWorkflowRun()
    const queryable = createWorkflowRunStoreConnection([createZeroWorkflowRunRow(run)], [], {
      stepRows: createZeroWorkflowStepRows(run),
      artifactRows: [createZeroWorkflowArtifactRow(run)],
    })

    const runs = await listWorkbookThreadWorkflowRuns(queryable, {
      documentId: 'doc-1',
      actorUserId: 'casey@example.com',
      threadId: 'thr-1',
    })

    expect(queryable.workflowRunInputs).toEqual([{ documentId: 'doc-1', actorUserId: 'casey@example.com', threadId: 'thr-1' }])
    expect(queryable.workflowStepInputs).toEqual([{ documentId: 'doc-1', actorUserId: 'casey@example.com', threadId: 'thr-1' }])
    expect(queryable.workflowArtifactInputs).toEqual([{ documentId: 'doc-1', actorUserId: 'casey@example.com', threadId: 'thr-1' }])
    expect(queryable.workflowMutationProofInputs).toEqual([{ documentId: 'doc-1', actorUserId: 'casey@example.com', threadId: 'thr-1' }])
    expect(queryable.calls.some((call) => call.text.includes('FROM workbook_workflow_run AS run'))).toBe(false)
    expect(queryable.calls.some((call) => call.text.includes('FROM workbook_workflow_step AS step'))).toBe(false)
    expect(queryable.calls.some((call) => call.text.includes('FROM workbook_workflow_artifact AS artifact'))).toBe(false)
    expect(runs).toEqual([
      expect.objectContaining({
        runId: 'workflow-1',
        startedByUserId: 'alex@example.com',
        artifact: expect.objectContaining({
          title: 'Workbook Summary',
        }),
      }),
    ])
  })

  it('hydrates durable workflow mutation proof after reload without parent Zero columns', async () => {
    const run = createWorkflowMutationProofRun()
    const queryable = createWorkflowRunStoreConnection([createZeroWorkflowRunRow(run)], [], {
      stepRows: createZeroWorkflowStepRows(run),
      artifactRows: [createZeroWorkflowArtifactRow(run)],
      mutationProofRows: [createZeroWorkflowMutationProofRow(run)],
    })

    const runs = await listWorkbookThreadWorkflowRuns(queryable, {
      documentId: 'doc-1',
      actorUserId: 'alex@example.com',
      threadId: 'thr-1',
    })

    expect(runs).toEqual([
      expect.objectContaining({
        runId: 'workflow-1',
        mutationExecuted: true,
        verificationComplete: false,
        mutationStatus: 'verification_incomplete',
        mutationReceipt: expect.objectContaining({
          status: 'verification_incomplete',
        }),
      }),
    ])
  })

  it('loads shared workflow runs for collaborator viewers', async () => {
    const run = createWorkflowRun()
    const queryable = createWorkflowRunStoreConnection([createZeroWorkflowRunRow(run)], [], {
      stepRows: createZeroWorkflowStepRows(run),
      artifactRows: [createZeroWorkflowArtifactRow(run)],
    })

    const runs = await listWorkbookThreadWorkflowRuns(queryable, {
      documentId: 'doc-1',
      actorUserId: 'casey@example.com',
      threadId: 'thr-1',
    })

    expect(runs).toEqual([
      expect.objectContaining({
        runId: 'workflow-1',
        startedByUserId: 'alex@example.com',
        workflowTemplate: 'summarizeWorkbook',
        steps: expect.arrayContaining([
          expect.objectContaining({
            stepId: 'inspect-workbook',
            status: 'completed',
          }),
        ]),
        artifact: expect.objectContaining({
          title: 'Workbook Summary',
        }),
      }),
    ])
    expect(queryable.workflowRunInputs).toEqual([{ documentId: 'doc-1', actorUserId: 'casey@example.com', threadId: 'thr-1' }])
    expect(queryable.calls.some((call) => call.text.includes('FROM workbook_workflow_run AS run'))).toBe(false)
  })

  it('hydrates structural workflow templates from durable rows after reload', async () => {
    const run = {
      ...createWorkflowRun(),
      runId: 'workflow-structural-1',
      workflowTemplate: 'hideCurrentRow' as const,
      title: 'Hide Current Row',
      summary: 'Staged a structural change set to hide row 7 on Sheet2.',
      steps: [
        {
          stepId: 'resolve-current-row',
          label: 'Resolve current row',
          status: 'completed' as const,
          summary: 'Resolved the selected row as row 7 on Sheet2.',
          updatedAtUnixMs: 110,
        },
        {
          stepId: 'stage-row-visibility-preview',
          label: 'Stage row visibility preview',
          status: 'completed' as const,
          summary: 'Staged the semantic preview that hides the current row.',
          updatedAtUnixMs: 120,
        },
      ],
      artifact: {
        kind: 'markdown' as const,
        title: 'Hide Row Preview',
        text: '## Hide Row Preview',
      },
    }
    const queryable = createWorkflowRunStoreConnection([createZeroWorkflowRunRow(run)], [], {
      stepRows: createZeroWorkflowStepRows(run),
      artifactRows: [createZeroWorkflowArtifactRow(run)],
    })

    const runs = await listWorkbookThreadWorkflowRuns(queryable, {
      documentId: 'doc-1',
      actorUserId: 'alex@example.com',
      threadId: 'thr-1',
    })

    expect(runs).toEqual([
      expect.objectContaining({
        runId: 'workflow-structural-1',
        workflowTemplate: 'hideCurrentRow',
        title: 'Hide Current Row',
        artifact: expect.objectContaining({
          title: 'Hide Row Preview',
        }),
        steps: expect.arrayContaining([
          expect.objectContaining({
            stepId: 'resolve-current-row',
          }),
          expect.objectContaining({
            stepId: 'stage-row-visibility-preview',
          }),
        ]),
      }),
    ])
  })

  it('hydrates newly added durable workflow templates from rows after reload', async () => {
    const run = {
      ...createWorkflowRun(),
      runId: 'workflow-import-1',
      workflowTemplate: 'normalizeCurrentSheetNumberFormats' as const,
      title: 'Normalize Current Sheet Number Formats',
      summary: 'Staged normalized number formats for 3 columns on Imports.',
      artifact: {
        kind: 'markdown' as const,
        title: 'Number Format Normalization Preview',
        text: '## Number Format Normalization Preview',
      },
    }
    const queryable = createWorkflowRunStoreConnection([createZeroWorkflowRunRow(run)], [], {
      stepRows: createZeroWorkflowStepRows(run),
      artifactRows: [createZeroWorkflowArtifactRow(run)],
    })

    const runs = await listWorkbookThreadWorkflowRuns(queryable, {
      documentId: 'doc-1',
      actorUserId: 'alex@example.com',
      threadId: 'thr-1',
    })

    expect(runs).toEqual([
      expect.objectContaining({
        runId: 'workflow-import-1',
        workflowTemplate: 'normalizeCurrentSheetNumberFormats',
      }),
    ])
  })

  it('hydrates formatting workflow templates from durable rows after reload', async () => {
    const run = {
      ...createWorkflowRun(),
      runId: 'workflow-formatting-1',
      workflowTemplate: 'highlightCurrentSheetOutliers' as const,
      title: 'Highlight Current Sheet Outliers',
      summary: 'Staged outlier highlights for 2 cells across 1 numeric column on Revenue.',
      artifact: {
        kind: 'markdown' as const,
        title: 'Current Sheet Outlier Highlights',
        text: '## Highlighted Numeric Outliers',
      },
    }
    const queryable = createWorkflowRunStoreConnection([createZeroWorkflowRunRow(run)], [], {
      stepRows: createZeroWorkflowStepRows(run),
      artifactRows: [createZeroWorkflowArtifactRow(run)],
    })

    const runs = await listWorkbookThreadWorkflowRuns(queryable, {
      documentId: 'doc-1',
      actorUserId: 'alex@example.com',
      threadId: 'thr-1',
    })

    expect(runs).toEqual([
      expect.objectContaining({
        runId: 'workflow-formatting-1',
        workflowTemplate: 'highlightCurrentSheetOutliers',
      }),
    ])
  })

  it('loads cancelled workflow runs with cancelled steps', async () => {
    const run = {
      ...createWorkflowRun(),
      summary: 'Cancelled workflow: Summarize Workbook',
      status: 'cancelled' as const,
      updatedAtUnixMs: 130,
      completedAtUnixMs: 130,
      errorMessage: 'Cancelled by alex@example.com.',
      steps: [
        {
          stepId: 'inspect-workbook',
          label: 'Inspect workbook structure',
          status: 'cancelled' as const,
          summary: 'Workflow cancelled before this step completed.',
          updatedAtUnixMs: 130,
        },
      ],
      artifact: null,
    }
    const queryable = createWorkflowRunStoreConnection([createZeroWorkflowRunRow(run)], [], {
      stepRows: createZeroWorkflowStepRows(run),
      artifactRows: [],
    })

    const runs = await listWorkbookThreadWorkflowRuns(queryable, {
      documentId: 'doc-1',
      actorUserId: 'alex@example.com',
      threadId: 'thr-1',
    })

    expect(runs).toEqual([
      expect.objectContaining({
        runId: 'workflow-1',
        status: 'cancelled',
        errorMessage: 'Cancelled by alex@example.com.',
        steps: [
          expect.objectContaining({
            stepId: 'inspect-workbook',
            status: 'cancelled',
          }),
        ],
        artifact: null,
      }),
    ])
  })

  it('drops workflow runs with impossible run timestamp ordering', async () => {
    const run = createWorkflowRun()
    const queryable = createWorkflowRunStoreConnection(
      [
        createZeroWorkflowRunRow(run, {
          runId: 'workflow-bad-time',
          createdAtUnixMs: 200,
          updatedAtUnixMs: 100,
        }),
      ],
      [
        (text, values) =>
          text.includes('FROM workbook_workflow_run AS run') && values?.[1] === 'thr-1'
            ? [
                {
                  runId: 'workflow-bad-time',
                  workbookId: 'doc-1',
                  threadId: run.threadId,
                  actorUserId: run.startedByUserId,
                  workflowTemplate: run.workflowTemplate,
                  title: run.title,
                  summary: run.summary,
                  status: run.status,
                  createdAtUnixMs: 200,
                  updatedAtUnixMs: 100,
                  completedAtUnixMs: run.completedAtUnixMs,
                  errorMessage: run.errorMessage,
                  stepsJson: run.steps,
                  artifactJson: run.artifact,
                } satisfies QueryResultRow,
              ]
            : null,
      ],
    )

    await expect(
      listWorkbookThreadWorkflowRuns(queryable, {
        documentId: 'doc-1',
        actorUserId: 'alex@example.com',
        threadId: 'thr-1',
      }),
    ).resolves.toEqual([])
  })

  it('drops workflow runs missing durable step rows once the reload batch has durable step coverage', async () => {
    const firstRun = createWorkflowRun()
    const secondRun = {
      ...createWorkflowRun(),
      runId: 'workflow-missing-steps',
      title: 'Missing Durable Steps',
    }
    const queryable = createWorkflowRunStoreConnection([createZeroWorkflowRunRow(firstRun), createZeroWorkflowRunRow(secondRun)], [], {
      stepRows: createZeroWorkflowStepRows(firstRun),
      artifactRows: [createZeroWorkflowArtifactRow(firstRun), createZeroWorkflowArtifactRow(secondRun)],
    })

    await expect(
      listWorkbookThreadWorkflowRuns(queryable, {
        documentId: 'doc-1',
        actorUserId: 'alex@example.com',
        threadId: 'thr-1',
      }),
    ).resolves.toEqual([
      expect.objectContaining({
        runId: firstRun.runId,
      }),
    ])
  })

  it('applies history limits after filtering malformed durable workflow rows', async () => {
    const invalidRun = createWorkflowRun()
    const validRun = {
      ...createWorkflowRun(),
      runId: 'workflow-valid-after-invalid',
      title: 'Valid Durable Workflow',
      updatedAtUnixMs: 130,
      completedAtUnixMs: 130,
    }
    const queryable = createWorkflowRunStoreConnection(
      [
        createZeroWorkflowRunRow(invalidRun, {
          runId: 'workflow-invalid-front',
          createdAtUnixMs: 200,
          updatedAtUnixMs: 100,
        }),
        createZeroWorkflowRunRow(validRun),
      ],
      [],
      {
        stepRows: createZeroWorkflowStepRows(validRun),
        artifactRows: [createZeroWorkflowArtifactRow(validRun)],
      },
    )

    await expect(
      listWorkbookThreadWorkflowRuns(queryable, {
        documentId: 'doc-1',
        actorUserId: 'alex@example.com',
        threadId: 'thr-1',
        limit: 1,
      }),
    ).resolves.toEqual([
      expect.objectContaining({
        runId: validRun.runId,
        title: 'Valid Durable Workflow',
      }),
    ])
  })

  it('loads workflow runs without durable artifact rows as artifact-free runs', async () => {
    const firstRun = createWorkflowRun()
    const secondRun = {
      ...createWorkflowRun(),
      runId: 'workflow-missing-artifact',
      title: 'Missing Durable Artifact',
    }
    const queryable = createWorkflowRunStoreConnection([createZeroWorkflowRunRow(firstRun), createZeroWorkflowRunRow(secondRun)], [], {
      stepRows: [...createZeroWorkflowStepRows(firstRun), ...createZeroWorkflowStepRows(secondRun)],
      artifactRows: [createZeroWorkflowArtifactRow(firstRun)],
    })

    await expect(
      listWorkbookThreadWorkflowRuns(queryable, {
        documentId: 'doc-1',
        actorUserId: 'alex@example.com',
        threadId: 'thr-1',
      }),
    ).resolves.toEqual([
      expect.objectContaining({
        runId: firstRun.runId,
      }),
      expect.objectContaining({
        artifact: null,
        runId: secondRun.runId,
      }),
    ])
  })
})
