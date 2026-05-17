import { describe, expect, it } from 'vitest'
import {
  appendWorkbookAgentRun,
  ensureWorkbookAgentRunSchema,
  listWorkbookAgentThreadRuns,
  listWorkbookAgentRuns,
} from '../workbook-agent-run-store.js'
import type { QueryResultRow, Queryable } from '../store.js'
import type { Row } from '@rocicorp/zero'

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
        return {
          rows: rows.filter((row): row is T => row !== null),
        }
      }
    }
    return { rows: [] }
  }
}

type ZeroAgentRunRow = Row['workbook_agent_run']

class FakeAgentRunConnection extends FakeQueryable {
  readonly zeroRunInputs: {
    readonly documentId: string
    readonly actorUserId: string
    readonly threadId?: string
  }[] = []

  constructor(
    responders: readonly ((text: string, values: readonly unknown[] | undefined) => QueryResultRow[] | null)[] = [],
    private readonly runRows: readonly ZeroAgentRunRow[] = [],
  ) {
    super(responders)
  }

  async listWorkbookAgentRunRows(input: {
    readonly documentId: string
    readonly actorUserId: string
    readonly threadId?: string
  }): Promise<readonly ZeroAgentRunRow[]> {
    this.zeroRunInputs.push(input)
    return this.runRows
  }
}

function createExecutionRecord() {
  return {
    id: 'run-1',
    bundleId: 'bundle-1',
    documentId: 'doc-1',
    threadId: 'thr-1',
    turnId: 'turn-1',
    actorUserId: 'alex@example.com',
    goalText: 'Apply only the selected command',
    planText: 'Apply the second command only',
    summary: 'Write cells in Sheet1!C3',
    scope: 'sheet' as const,
    riskClass: 'medium' as const,
    acceptedScope: 'partial' as const,
    appliedBy: 'user' as const,
    baseRevision: 3,
    appliedRevision: 4,
    createdAtUnixMs: 100,
    appliedAtUnixMs: 200,
    context: {
      selection: {
        sheetName: 'Sheet1',
        address: 'A1',
      },
      viewport: {
        rowStart: 0,
        rowEnd: 10,
        colStart: 0,
        colEnd: 5,
      },
    },
    commands: [
      {
        kind: 'writeRange' as const,
        sheetName: 'Sheet1',
        startAddress: 'C3',
        values: [[2]],
      },
    ],
    preview: null,
  }
}

function createZeroAgentRunRow(
  record: ReturnType<typeof createExecutionRecord>,
  overrides: Partial<ZeroAgentRunRow> = {},
): ZeroAgentRunRow {
  return {
    id: record.id,
    bundleId: record.bundleId,
    workbookId: record.documentId,
    threadId: record.threadId,
    turnId: record.turnId,
    actorUserId: record.actorUserId,
    goalText: record.goalText,
    planText: record.planText,
    summary: record.summary,
    scope: record.scope,
    riskClass: record.riskClass,
    acceptedScope: record.acceptedScope,
    appliedBy: record.appliedBy,
    baseRevision: record.baseRevision,
    appliedRevision: record.appliedRevision,
    createdAtUnixMs: record.createdAtUnixMs,
    appliedAtUnixMs: record.appliedAtUnixMs,
    context: record.context,
    commands: record.commands,
    preview: record.preview,
    ...overrides,
  }
}

describe('workbook-agent-run-store', () => {
  it('backfills legacy nullable bundle ids before enforcing the execution schema', async () => {
    const queryable = new FakeQueryable()

    await ensureWorkbookAgentRunSchema(queryable)

    const backfillIndex = queryable.calls.findIndex((call) => call.text.includes('SET bundle_id = id'))
    const notNullIndex = queryable.calls.findIndex((call) => call.text.includes('ALTER COLUMN bundle_id SET NOT NULL'))
    expect(backfillIndex).toBeGreaterThan(-1)
    expect(notNullIndex).toBeGreaterThan(backfillIndex)
  })

  it('backfills legacy acceptance metadata before enforcing execution schema defaults', async () => {
    const queryable = new FakeQueryable()

    await ensureWorkbookAgentRunSchema(queryable)

    const acceptedScopeBackfillIndex = queryable.calls.findIndex((call) => call.text.includes("SET accepted_scope = 'full'"))
    const acceptedScopeNotNullIndex = queryable.calls.findIndex((call) => call.text.includes('ALTER COLUMN accepted_scope SET NOT NULL'))
    const appliedByBackfillIndex = queryable.calls.findIndex((call) => call.text.includes("SET applied_by = 'user'"))
    const appliedByNotNullIndex = queryable.calls.findIndex((call) => call.text.includes('ALTER COLUMN applied_by SET NOT NULL'))
    expect(acceptedScopeBackfillIndex).toBeGreaterThan(-1)
    expect(acceptedScopeNotNullIndex).toBeGreaterThan(acceptedScopeBackfillIndex)
    expect(appliedByBackfillIndex).toBeGreaterThan(-1)
    expect(appliedByNotNullIndex).toBeGreaterThan(appliedByBackfillIndex)
  })

  it('persists partial accepted scope in execution rows', async () => {
    const queryable = new FakeQueryable()

    await appendWorkbookAgentRun(queryable, createExecutionRecord())

    const insertQuery = queryable.calls.find((call) => call.text.includes('INSERT INTO workbook_agent_run'))
    expect(insertQuery?.values?.[11]).toBe('partial')
  })

  it('loads partial execution records from stored rows', async () => {
    const record = createExecutionRecord()
    const queryable = new FakeAgentRunConnection(
      [],
      [
        {
          id: record.id,
          bundleId: record.bundleId,
          workbookId: record.documentId,
          threadId: record.threadId,
          turnId: record.turnId,
          actorUserId: record.actorUserId,
          goalText: record.goalText,
          planText: record.planText,
          summary: record.summary,
          scope: record.scope,
          riskClass: record.riskClass,
          acceptedScope: record.acceptedScope,
          appliedBy: record.appliedBy,
          baseRevision: record.baseRevision,
          appliedRevision: record.appliedRevision,
          createdAtUnixMs: record.createdAtUnixMs,
          appliedAtUnixMs: record.appliedAtUnixMs,
          context: record.context,
          commands: record.commands,
          preview: record.preview,
        },
      ],
    )

    const records = await listWorkbookAgentRuns(queryable, {
      documentId: 'doc-1',
      actorUserId: 'alex@example.com',
    })

    expect(queryable.zeroRunInputs).toEqual([{ documentId: 'doc-1', actorUserId: 'alex@example.com' }])
    expect(queryable.calls.some((call) => call.text.includes('FROM workbook_agent_run'))).toBe(false)
    expect(records).toEqual([
      expect.objectContaining({
        acceptedScope: 'partial',
        commands: [
          {
            kind: 'writeRange',
            sheetName: 'Sheet1',
            startAddress: 'C3',
            values: [[2]],
          },
        ],
      }),
    ])
  })

  it('hydrates execution rows from the shared Zero model', async () => {
    const record = createExecutionRecord()
    const queryable = new FakeAgentRunConnection(
      [],
      [
        {
          id: record.id,
          bundleId: record.bundleId,
          workbookId: record.documentId,
          threadId: record.threadId,
          turnId: record.turnId,
          actorUserId: record.actorUserId,
          goalText: record.goalText,
          planText: record.planText,
          summary: record.summary,
          scope: record.scope,
          riskClass: record.riskClass,
          acceptedScope: 'full',
          appliedBy: 'user',
          baseRevision: record.baseRevision,
          appliedRevision: record.appliedRevision,
          createdAtUnixMs: record.createdAtUnixMs,
          appliedAtUnixMs: record.appliedAtUnixMs,
          context: record.context,
          commands: record.commands,
          preview: record.preview,
        },
      ],
    )

    const records = await listWorkbookAgentRuns(queryable, {
      documentId: 'doc-1',
      actorUserId: 'alex@example.com',
    })

    expect(records).toEqual([
      expect.objectContaining({
        id: record.id,
        bundleId: record.bundleId,
      }),
    ])
  })

  it('drops execution rows with impossible revision or timestamp ordering', async () => {
    const valid = createExecutionRecord()
    const queryable = new FakeAgentRunConnection(
      [],
      [
        {
          id: 'run-bad-revision',
          bundleId: 'bundle-bad',
          workbookId: 'doc-1',
          threadId: 'thr-1',
          turnId: 'turn-1',
          actorUserId: 'alex@example.com',
          goalText: 'Bad revision row',
          planText: null,
          summary: 'Should not hydrate',
          scope: 'sheet',
          riskClass: 'medium',
          acceptedScope: 'partial',
          appliedBy: 'user',
          baseRevision: 9,
          appliedRevision: 8,
          createdAtUnixMs: 100,
          appliedAtUnixMs: 200,
          context: null,
          commands: valid.commands,
          preview: null,
        },
        {
          id: valid.id,
          bundleId: valid.bundleId,
          workbookId: valid.documentId,
          threadId: valid.threadId,
          turnId: valid.turnId,
          actorUserId: valid.actorUserId,
          goalText: valid.goalText,
          planText: valid.planText,
          summary: valid.summary,
          scope: valid.scope,
          riskClass: valid.riskClass,
          acceptedScope: valid.acceptedScope,
          appliedBy: valid.appliedBy,
          baseRevision: valid.baseRevision,
          appliedRevision: valid.appliedRevision,
          createdAtUnixMs: valid.createdAtUnixMs,
          appliedAtUnixMs: valid.appliedAtUnixMs,
          context: valid.context,
          commands: valid.commands,
          preview: valid.preview,
        },
        {
          id: 'run-bad-time',
          bundleId: 'bundle-bad-time',
          workbookId: 'doc-1',
          threadId: 'thr-1',
          turnId: 'turn-1',
          actorUserId: 'alex@example.com',
          goalText: 'Bad timestamp row',
          planText: null,
          summary: 'Should not hydrate',
          scope: 'sheet',
          riskClass: 'medium',
          acceptedScope: 'partial',
          appliedBy: 'user',
          baseRevision: 3,
          appliedRevision: 4,
          createdAtUnixMs: 300,
          appliedAtUnixMs: 200,
          context: null,
          commands: valid.commands,
          preview: null,
        },
      ],
    )

    const records = await listWorkbookAgentRuns(queryable, {
      documentId: 'doc-1',
      actorUserId: 'alex@example.com',
    })

    expect(records.map((record) => record.id)).toEqual(['run-1'])
  })

  it('applies workbook execution limits after filtering malformed Zero rows', async () => {
    const valid = {
      ...createExecutionRecord(),
      id: 'run-valid-after-invalid',
      bundleId: 'bundle-valid-after-invalid',
      summary: 'Valid run after a malformed front row',
    }
    const queryable = new FakeAgentRunConnection(
      [],
      [
        createZeroAgentRunRow(valid, {
          id: 'run-invalid-front',
          bundleId: 'bundle-invalid-front',
          baseRevision: 9,
          appliedRevision: 8,
        }),
        createZeroAgentRunRow(valid),
      ],
    )

    const records = await listWorkbookAgentRuns(queryable, {
      documentId: 'doc-1',
      actorUserId: 'alex@example.com',
      limit: 1,
    })

    expect(records.map((record) => record.id)).toEqual(['run-valid-after-invalid'])
  })

  it('applies thread execution limits after filtering malformed Zero rows', async () => {
    const valid = {
      ...createExecutionRecord(),
      id: 'thread-run-valid-after-invalid',
      bundleId: 'thread-bundle-valid-after-invalid',
      threadId: 'thr-shared',
    }
    const queryable = new FakeAgentRunConnection(
      [],
      [
        createZeroAgentRunRow(valid, {
          id: 'thread-run-invalid-front',
          bundleId: 'thread-bundle-invalid-front',
          createdAtUnixMs: 300,
          appliedAtUnixMs: 200,
        }),
        createZeroAgentRunRow(valid),
      ],
    )

    const records = await listWorkbookAgentThreadRuns(queryable, {
      documentId: 'doc-1',
      actorUserId: 'casey@example.com',
      threadId: 'thr-shared',
      limit: 1,
    })

    expect(records.map((record) => record.id)).toEqual(['thread-run-valid-after-invalid'])
  })

  it('loads shared thread execution records for collaborator viewers', async () => {
    const record = {
      ...createExecutionRecord(),
      threadId: 'thr-shared',
      actorUserId: 'alex@example.com',
    }
    const queryable = new FakeAgentRunConnection(
      [],
      [
        {
          id: record.id,
          bundleId: record.bundleId,
          workbookId: record.documentId,
          threadId: record.threadId,
          turnId: record.turnId,
          actorUserId: record.actorUserId,
          goalText: record.goalText,
          planText: record.planText,
          summary: record.summary,
          scope: record.scope,
          riskClass: record.riskClass,
          acceptedScope: record.acceptedScope,
          appliedBy: record.appliedBy,
          baseRevision: record.baseRevision,
          appliedRevision: record.appliedRevision,
          createdAtUnixMs: record.createdAtUnixMs,
          appliedAtUnixMs: record.appliedAtUnixMs,
          context: record.context,
          commands: record.commands,
          preview: record.preview,
        },
      ],
    )

    const records = await listWorkbookAgentThreadRuns(queryable, {
      documentId: 'doc-1',
      actorUserId: 'casey@example.com',
      threadId: 'thr-shared',
    })

    expect(queryable.zeroRunInputs).toEqual([{ documentId: 'doc-1', actorUserId: 'casey@example.com', threadId: 'thr-shared' }])
    expect(queryable.calls.some((call) => call.text.includes('FROM workbook_agent_run AS run'))).toBe(false)
    expect(records).toEqual([
      expect.objectContaining({
        threadId: 'thr-shared',
        actorUserId: 'alex@example.com',
      }),
    ])
  })

  it('delegates shared thread visibility to the owner-bound Zero query', async () => {
    const queryable = new FakeAgentRunConnection()

    await listWorkbookAgentThreadRuns(queryable, {
      documentId: 'doc-1',
      actorUserId: 'casey@example.com',
      threadId: 'thr-shared',
    })

    expect(queryable.zeroRunInputs).toEqual([{ documentId: 'doc-1', actorUserId: 'casey@example.com', threadId: 'thr-shared' }])
    expect(queryable.calls.some((call) => call.text.includes('workbook_chat_thread'))).toBe(false)
  })
})
