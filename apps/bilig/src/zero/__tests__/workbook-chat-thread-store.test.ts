import { describe, expect, it } from 'vitest'
import {
  ensureWorkbookChatThreadSchema,
  listWorkbookAgentThreadSummaries,
  loadWorkbookAgentThreadState,
  saveWorkbookAgentThreadState,
} from '../workbook-chat-thread-store.js'
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

function createThreadState() {
  return {
    documentId: 'doc-1',
    threadId: 'thr-1',
    actorUserId: 'alex@example.com',
    scope: 'private' as const,
    executionPolicy: 'autoApplyAll' as const,
    context: {
      selection: {
        sheetName: 'Sheet1',
        address: 'A1',
      },
      viewport: {
        rowStart: 0,
        rowEnd: 20,
        colStart: 0,
        colEnd: 8,
      },
    },
    entries: [
      {
        id: 'entry-user-1',
        kind: 'user' as const,
        turnId: 'turn-1',
        text: 'Summarize Sheet1',
        phase: null,
        toolName: null,
        toolStatus: null,
        argumentsText: null,
        outputText: null,
        success: null,
        citations: [],
      },
      {
        id: 'tool-call-1',
        kind: 'tool' as const,
        turnId: 'turn-1',
        text: null,
        phase: null,
        toolName: 'read_workbook',
        toolStatus: 'completed' as const,
        argumentsText: '{"sheetName":"Sheet1"}',
        outputText: '{"summary":"Loaded workbook"}',
        success: true,
        citations: [],
      },
      {
        id: 'system-review-item:review-1',
        kind: 'system' as const,
        turnId: 'turn-1',
        text: 'Review item queued',
        phase: null,
        toolName: null,
        toolStatus: null,
        argumentsText: null,
        outputText: null,
        success: null,
        citations: [
          {
            kind: 'range' as const,
            sheetName: 'Sheet1',
            startAddress: 'B2',
            endAddress: 'B2',
            role: 'target' as const,
          },
        ],
      },
    ],
    reviewQueueItems: [
      {
        id: 'review-1',
        documentId: 'doc-1',
        threadId: 'thr-1',
        turnId: 'turn-1',
        goalText: 'Normalize selection',
        summary: 'Write cells in Sheet1!B2',
        scope: 'selection' as const,
        riskClass: 'low' as const,
        reviewMode: 'manual' as const,
        ownerUserId: null,
        status: 'pending' as const,
        decidedByUserId: null,
        decidedAtUnixMs: null,
        recommendations: [],
        baseRevision: 12,
        createdAtUnixMs: 100,
        context: {
          selection: {
            sheetName: 'Sheet1',
            address: 'A1',
          },
          viewport: {
            rowStart: 0,
            rowEnd: 20,
            colStart: 0,
            colEnd: 8,
          },
        },
        commands: [
          {
            kind: 'writeRange' as const,
            sheetName: 'Sheet1',
            startAddress: 'B2',
            values: [[42]],
          },
        ],
        affectedRanges: [
          {
            sheetName: 'Sheet1',
            startAddress: 'B2',
            endAddress: 'B2',
            role: 'target' as const,
          },
        ],
        estimatedAffectedCells: 1,
      },
    ],
    updatedAtUnixMs: 1234,
  }
}

function createReviewQueueItemRow(state: ReturnType<typeof createThreadState>) {
  return state.reviewQueueItems.map((reviewItem) => ({
    reviewItemId: reviewItem.id,
    workbookId: reviewItem.documentId,
    threadId: reviewItem.threadId,
    actorUserId: state.actorUserId,
    turnId: reviewItem.turnId,
    goalText: reviewItem.goalText,
    summary: reviewItem.summary,
    scope: reviewItem.scope,
    riskClass: reviewItem.riskClass,
    reviewMode: reviewItem.reviewMode,
    ownerUserId: reviewItem.ownerUserId,
    status: reviewItem.status,
    decidedByUserId: reviewItem.decidedByUserId,
    decidedAtUnixMs: reviewItem.decidedAtUnixMs,
    baseRevision: reviewItem.baseRevision,
    createdAtUnixMs: reviewItem.createdAtUnixMs,
    contextJson: reviewItem.context,
    commandsJson: reviewItem.commands,
    affectedRangesJson: reviewItem.affectedRanges,
    estimatedAffectedCells: reviewItem.estimatedAffectedCells,
    recommendationsJson: reviewItem.recommendations,
  })) satisfies QueryResultRow[]
}

describe('workbook-chat-thread-store', () => {
  it('creates the review-queue schema without legacy pending-bundle compatibility paths', async () => {
    const queryable = new FakeQueryable()

    await ensureWorkbookChatThreadSchema(queryable)

    expect(queryable.calls.some((call) => call.text.includes('CREATE TABLE IF NOT EXISTS workbook_review_queue_item'))).toBe(true)
    expect(
      queryable.calls.some(
        (call) =>
          call.text.includes('information_schema.tables') ||
          call.text.includes('DROP TABLE') ||
          call.text.includes('DROP COLUMN IF EXISTS'),
      ),
    ).toBe(false)
  })

  it('persists thread metadata, timeline items, and review queue rows', async () => {
    const queryable = new FakeQueryable()
    const state = createThreadState()

    await saveWorkbookAgentThreadState(queryable, state)

    expect(queryable.calls.some((call) => call.text.includes('INSERT INTO workbook_chat_thread'))).toBe(true)
    const threadInsert = queryable.calls.find((call) => call.text.includes('INSERT INTO workbook_chat_thread'))
    expect(threadInsert?.values?.[4]).toBe(state.executionPolicy)
    expect(threadInsert?.values?.[5]).toBe(JSON.stringify(state.context))
    expect(threadInsert?.values?.[6]).toBe(3)
    expect(threadInsert?.values?.[7]).toBe('Review item queued')
    expect(threadInsert?.values?.[8]).toBe(state.updatedAtUnixMs)
    expect(queryable.calls.some((call) => call.text.includes('DELETE FROM workbook_chat_item'))).toBe(true)
    expect(queryable.calls.some((call) => call.text.includes('DELETE FROM workbook_review_queue_item'))).toBe(true)
    expect(queryable.calls.filter((call) => call.text.includes('INSERT INTO workbook_chat_item')).length).toBe(3)
    expect(queryable.calls.filter((call) => call.text.includes('INSERT INTO workbook_chat_tool_call')).length).toBe(1)
    const reviewInsert = queryable.calls.find((call) => call.text.includes('INSERT INTO workbook_review_queue_item'))
    expect(reviewInsert?.values?.[3]).toBe('review-1')
    expect(reviewInsert?.values?.[9]).toBe('manual')
    expect(reviewInsert?.values?.[20]).toBe(JSON.stringify([]))
    const itemInsert = queryable.calls.find(
      (call) => call.text.includes('INSERT INTO workbook_chat_item') && call.values?.[3] === 'system-review-item:review-1',
    )
    expect(itemInsert?.values?.[14]).toBe(
      JSON.stringify([
        {
          kind: 'range',
          sheetName: 'Sheet1',
          startAddress: 'B2',
          endAddress: 'B2',
          role: 'target',
        },
      ]),
    )
    const toolInsert = queryable.calls.find(
      (call) => call.text.includes('INSERT INTO workbook_chat_tool_call') && call.values?.[3] === 'tool-call-1',
    )
    expect(toolInsert?.values?.[6]).toBe('read_workbook')
    expect(toolInsert?.values?.[8]).toBe('{"sheetName":"Sheet1"}')
    expect(toolInsert?.values?.[9]).toBe('{"summary":"Loaded workbook"}')
  })

  it('dedupes duplicate entry ids before inserting durable chat items', async () => {
    const queryable = new FakeQueryable()
    const state = createThreadState()

    await saveWorkbookAgentThreadState(queryable, {
      ...state,
      entries: [
        state.entries[0],
        {
          ...state.entries[0],
          text: 'Updated prompt',
        },
        state.entries[1],
      ],
    })

    const itemInserts = queryable.calls.filter((call) => call.text.includes('INSERT INTO workbook_chat_item'))
    expect(itemInserts).toHaveLength(2)
    expect(itemInserts[0]?.values?.[3]).toBe('entry-user-1')
    expect(itemInserts[0]?.values?.[7]).toBe('Updated prompt')
    expect(itemInserts.every((call) => call.text.includes('ON CONFLICT'))).toBe(true)
  })

  it('loads a durable thread snapshot with entries and review queue items', async () => {
    const state = createThreadState()
    const queryable = new FakeQueryable([
      (text) =>
        text.includes('FROM workbook_chat_thread')
          ? [
              {
                workbookId: state.documentId,
                threadId: state.threadId,
                actorUserId: state.actorUserId,
                scope: state.scope,
                executionPolicy: state.executionPolicy,
                contextJson: state.context,
                updatedAtUnixMs: state.updatedAtUnixMs,
              } satisfies QueryResultRow,
            ]
          : null,
      (text) =>
        text.includes('FROM workbook_chat_item')
          ? state.entries.map((entry, index) => ({
              entryId: entry.id,
              turnId: entry.turnId,
              kind: entry.kind,
              text: entry.text,
              phase: entry.phase,
              toolName: entry.id === 'tool-call-1' ? 'bilig_read_workbook' : entry.toolName,
              toolStatus: entry.toolStatus,
              argumentsText: entry.argumentsText,
              outputText: entry.outputText,
              success: entry.success,
              citationsJson: entry.citations,
              sortOrder: index,
            }))
          : null,
      (text) =>
        text.includes('FROM workbook_chat_tool_call')
          ? state.entries
              .filter((entry) => entry.kind === 'tool')
              .map((entry, index) => ({
                entryId: entry.id,
                turnId: entry.turnId,
                toolName: entry.id === 'tool-call-1' ? 'bilig_read_workbook' : entry.toolName,
                toolStatus: entry.toolStatus,
                argumentsText: entry.argumentsText,
                outputText: entry.outputText,
                success: entry.success,
                sortOrder: index,
              }))
          : null,
      (text) => (text.includes('FROM workbook_review_queue_item') ? createReviewQueueItemRow(state) : null),
    ])

    const loaded = await loadWorkbookAgentThreadState(queryable, {
      documentId: 'doc-1',
      threadId: 'thr-1',
      actorUserId: 'alex@example.com',
    })

    expect(loaded).toEqual(state)
  })

  it('falls back to a collaborator-owned shared thread when the current user has no local row', async () => {
    const state = {
      ...createThreadState(),
      threadId: 'thr-shared',
      actorUserId: 'alex@example.com',
      scope: 'shared' as const,
      executionPolicy: 'ownerReview' as const,
    }
    const queryable = new FakeQueryable([
      (text, values) =>
        text.includes('FROM workbook_chat_thread') && values?.[2] === 'casey@example.com'
          ? [
              {
                workbookId: state.documentId,
                threadId: state.threadId,
                actorUserId: state.actorUserId,
                scope: state.scope,
                executionPolicy: state.executionPolicy,
                contextJson: state.context,
                updatedAtUnixMs: state.updatedAtUnixMs,
              } satisfies QueryResultRow,
            ]
          : null,
      (text, values) =>
        text.includes('FROM workbook_chat_item') && values?.[2] === 'alex@example.com'
          ? state.entries.map((entry, index) => ({
              entryId: entry.id,
              turnId: entry.turnId,
              kind: entry.kind,
              text: entry.text,
              phase: entry.phase,
              toolName: entry.toolName,
              toolStatus: entry.toolStatus,
              argumentsText: entry.argumentsText,
              outputText: entry.outputText,
              success: entry.success,
              citationsJson: entry.citations,
              sortOrder: index,
            }))
          : null,
      (text, values) =>
        text.includes('FROM workbook_chat_tool_call') && values?.[2] === 'alex@example.com'
          ? state.entries
              .filter((entry) => entry.kind === 'tool')
              .map((entry, index) => ({
                entryId: entry.id,
                turnId: entry.turnId,
                toolName: entry.toolName,
                toolStatus: entry.toolStatus,
                argumentsText: entry.argumentsText,
                outputText: entry.outputText,
                success: entry.success,
                sortOrder: index,
              }))
          : null,
      (text, values) =>
        text.includes('FROM workbook_review_queue_item') && values?.[2] === 'alex@example.com' ? createReviewQueueItemRow(state) : null,
    ])

    const loaded = await loadWorkbookAgentThreadState(queryable, {
      documentId: 'doc-1',
      threadId: 'thr-shared',
      actorUserId: 'casey@example.com',
    })

    expect(loaded).toEqual(state)
  })

  it('hydrates tool call state from dedicated durable tool call rows', async () => {
    const state = createThreadState()
    const toolEntry = state.entries.find((entry) => entry.id === 'tool-call-1')
    const queryable = new FakeQueryable([
      (text) =>
        text.includes('FROM workbook_chat_thread')
          ? [
              {
                workbookId: state.documentId,
                threadId: state.threadId,
                actorUserId: state.actorUserId,
                scope: state.scope,
                executionPolicy: state.executionPolicy,
                contextJson: state.context,
                updatedAtUnixMs: state.updatedAtUnixMs,
              } satisfies QueryResultRow,
            ]
          : null,
      (text) =>
        text.includes('FROM workbook_chat_item')
          ? state.entries.map((entry, index) => ({
              entryId: entry.id,
              turnId: entry.turnId,
              kind: entry.kind,
              text: entry.text,
              phase: entry.phase,
              toolName: entry.id === 'tool-call-1' ? null : entry.toolName,
              toolStatus: entry.id === 'tool-call-1' ? null : entry.toolStatus,
              argumentsText: entry.id === 'tool-call-1' ? null : entry.argumentsText,
              outputText: entry.id === 'tool-call-1' ? null : entry.outputText,
              success: entry.id === 'tool-call-1' ? null : entry.success,
              citationsJson: entry.citations,
              sortOrder: index,
            }))
          : null,
      (text) =>
        text.includes('FROM workbook_chat_tool_call') && toolEntry
          ? [
              {
                entryId: toolEntry.id,
                turnId: toolEntry.turnId,
                toolName: toolEntry.toolName,
                toolStatus: toolEntry.toolStatus,
                argumentsText: toolEntry.argumentsText,
                outputText: toolEntry.outputText,
                success: toolEntry.success,
                sortOrder: 1,
              } satisfies QueryResultRow,
            ]
          : null,
      (text) => (text.includes('FROM workbook_review_queue_item') ? createReviewQueueItemRow(state) : null),
    ])

    const loaded = await loadWorkbookAgentThreadState(queryable, {
      documentId: 'doc-1',
      threadId: 'thr-1',
      actorUserId: 'alex@example.com',
    })

    expect(loaded?.entries.find((entry) => entry.id === 'tool-call-1')).toEqual(toolEntry)
  })

  it('lists durable thread summaries ordered by most recent activity', async () => {
    const queryable = new FakeQueryable([
      (text) =>
        text.includes('ROW_NUMBER() OVER')
          ? [
              {
                threadId: 'thr-2',
                scope: 'shared',
                ownerUserId: 'alex@example.com',
                updatedAtUnixMs: 200,
                entryCount: 3,
                reviewQueueItemCount: 0,
                latestEntryText: 'Applied shared cleanup at revision r7',
              } satisfies QueryResultRow,
              {
                threadId: 'thr-1',
                scope: 'private',
                ownerUserId: 'alex@example.com',
                updatedAtUnixMs: 100,
                entryCount: 1,
                reviewQueueItemCount: 1,
                latestEntryText: 'Review item queued',
              } satisfies QueryResultRow,
            ]
          : null,
    ])

    const summaries = await listWorkbookAgentThreadSummaries(queryable, {
      documentId: 'doc-1',
      actorUserId: 'alex@example.com',
    })

    expect(summaries).toEqual([
      {
        threadId: 'thr-2',
        scope: 'shared',
        ownerUserId: 'alex@example.com',
        updatedAtUnixMs: 200,
        entryCount: 3,
        reviewQueueItemCount: 0,
        latestEntryText: 'Applied shared cleanup at revision r7',
      },
      {
        threadId: 'thr-1',
        scope: 'private',
        ownerUserId: 'alex@example.com',
        updatedAtUnixMs: 100,
        entryCount: 1,
        reviewQueueItemCount: 1,
        latestEntryText: 'Review item queued',
      },
    ])
  })

  it('includes collaborator-owned shared threads in the summary query', async () => {
    const queryable = new FakeQueryable([
      (text) =>
        text.includes("thread.scope = 'shared'")
          ? [
              {
                threadId: 'thr-shared',
                scope: 'shared',
                ownerUserId: 'alex@example.com',
                updatedAtUnixMs: 300,
                entryCount: 2,
                reviewQueueItemCount: 0,
                latestEntryText: 'Applied shared cleanup at revision r9',
              } satisfies QueryResultRow,
            ]
          : null,
    ])

    const summaries = await listWorkbookAgentThreadSummaries(queryable, {
      documentId: 'doc-1',
      actorUserId: 'casey@example.com',
    })

    expect(summaries).toEqual([
      {
        threadId: 'thr-shared',
        scope: 'shared',
        ownerUserId: 'alex@example.com',
        updatedAtUnixMs: 300,
        entryCount: 2,
        reviewQueueItemCount: 0,
        latestEntryText: 'Applied shared cleanup at revision r9',
      },
    ])
    expect(queryable.calls.at(-1)?.text).toContain("AND (thread.actor_user_id = $2 OR thread.scope = 'shared')")
  })
})
