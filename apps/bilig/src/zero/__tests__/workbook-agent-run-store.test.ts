import { describe, expect, it } from 'vitest'
import {
  appendWorkbookAgentRun,
  ensureWorkbookAgentRunSchema,
  listWorkbookAgentThreadRuns,
  listWorkbookAgentRuns,
} from '../workbook-agent-run-store.js'
import type { QueryResultRow, Queryable } from '../store.js'

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
    const queryable = new FakeQueryable([
      (text) =>
        text.includes('FROM workbook_agent_run')
          ? [
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
                contextJson: record.context,
                commandsJson: record.commands,
                previewJson: record.preview,
              } satisfies QueryResultRow,
            ]
          : null,
    ])

    const records = await listWorkbookAgentRuns(queryable, {
      documentId: 'doc-1',
      actorUserId: 'alex@example.com',
    })

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

  it('hydrates legacy execution rows without bundle ids using the run id as the stable bundle id', async () => {
    const record = createExecutionRecord()
    const queryable = new FakeQueryable([
      (text) =>
        text.includes('FROM workbook_agent_run')
          ? [
              {
                id: record.id,
                bundleId: null,
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
                contextJson: record.context,
                commandsJson: record.commands,
                previewJson: record.preview,
              } satisfies QueryResultRow,
            ]
          : null,
    ])

    const records = await listWorkbookAgentRuns(queryable, {
      documentId: 'doc-1',
      actorUserId: 'alex@example.com',
    })

    expect(records).toEqual([
      expect.objectContaining({
        id: record.id,
        bundleId: record.id,
      }),
    ])
  })

  it('drops execution rows with impossible revision or timestamp ordering', async () => {
    const valid = createExecutionRecord()
    const queryable = new FakeQueryable([
      (text) =>
        text.includes('FROM workbook_agent_run')
          ? [
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
                contextJson: null,
                commandsJson: valid.commands,
                previewJson: null,
              } satisfies QueryResultRow,
              {
                ...valid,
                workbookId: valid.documentId,
                contextJson: valid.context,
                commandsJson: valid.commands,
                previewJson: valid.preview,
                baseRevision: valid.baseRevision,
                appliedRevision: valid.appliedRevision,
                createdAtUnixMs: valid.createdAtUnixMs,
                appliedAtUnixMs: valid.appliedAtUnixMs,
              } satisfies QueryResultRow,
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
                contextJson: null,
                commandsJson: valid.commands,
                previewJson: null,
              } satisfies QueryResultRow,
            ]
          : null,
    ])

    const records = await listWorkbookAgentRuns(queryable, {
      documentId: 'doc-1',
      actorUserId: 'alex@example.com',
    })

    expect(records.map((record) => record.id)).toEqual(['run-1'])
  })

  it('loads shared thread execution records for collaborator viewers', async () => {
    const record = {
      ...createExecutionRecord(),
      threadId: 'thr-shared',
      actorUserId: 'alex@example.com',
    }
    const queryable = new FakeQueryable([
      (text, values) =>
        text.includes('FROM workbook_agent_run AS run') && values?.[1] === 'thr-shared' && values?.[2] === 'casey@example.com'
          ? [
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
                contextJson: record.context,
                commandsJson: record.commands,
                previewJson: record.preview,
              } satisfies QueryResultRow,
            ]
          : null,
    ])

    const records = await listWorkbookAgentThreadRuns(queryable, {
      documentId: 'doc-1',
      actorUserId: 'casey@example.com',
      threadId: 'thr-shared',
    })

    expect(records).toEqual([
      expect.objectContaining({
        threadId: 'thr-shared',
        actorUserId: 'alex@example.com',
      }),
    ])
  })
})
