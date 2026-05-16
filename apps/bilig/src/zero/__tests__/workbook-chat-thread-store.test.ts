import { describe, expect, it } from 'vitest'
import {
  ensureWorkbookChatThreadSchema,
  listWorkbookAgentThreadSummaries,
  loadWorkbookAgentThreadState,
  saveWorkbookAgentThreadState,
  type WorkbookChatThreadStoreConnection,
} from '../workbook-chat-thread-store.js'
import type { QueryResultRow, Queryable } from '../store.js'
import type { ZeroWorkbookChatThreadRow } from '../workbook-chat-thread-normalizers.js'

interface RecordedQuery {
  readonly text: string
  readonly values: readonly unknown[] | undefined
}

class FakeQueryable implements Queryable, WorkbookChatThreadStoreConnection {
  readonly calls: RecordedQuery[] = []
  readonly zeroThreadInputs: { readonly documentId: string; readonly actorUserId: string }[] = []

  constructor(
    private readonly responders: readonly ((text: string, values: readonly unknown[] | undefined) => QueryResultRow[] | null)[] = [],
    private readonly runRows: readonly ZeroWorkbookChatThreadRow[] = [],
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

  async listWorkbookChatThreadRows(input: {
    readonly documentId: string
    readonly actorUserId: string
  }): Promise<readonly ZeroWorkbookChatThreadRow[]> {
    this.zeroThreadInputs.push(input)
    return this.runRows
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

function createZeroThreadRow(state: ReturnType<typeof createThreadState>): ZeroWorkbookChatThreadRow {
  return {
    workbookId: state.documentId,
    threadId: state.threadId,
    ownerUserId: state.actorUserId,
    scope: state.scope,
    updatedAtUnixMs: state.updatedAtUnixMs,
    entryCount: state.entries.length,
    reviewQueueItemCount: state.reviewQueueItems.length,
    latestEntryText: 'Review item queued',
  }
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

  it('reconciles denormalized thread summaries from durable child tables during schema bootstrap', async () => {
    const queryable = new FakeQueryable()

    await ensureWorkbookChatThreadSchema(queryable)

    const reviewTableIndex = queryable.calls.findIndex((call) =>
      call.text.includes('CREATE TABLE IF NOT EXISTS workbook_review_queue_item'),
    )
    const reconcileIndex = queryable.calls.findIndex((call) => call.text.includes('WITH entry_stats AS'))
    const indexIndex = queryable.calls.findIndex((call) =>
      call.text.includes('CREATE INDEX IF NOT EXISTS workbook_chat_thread_document_actor_updated_idx'),
    )
    expect(reviewTableIndex).toBeGreaterThanOrEqual(0)
    expect(reconcileIndex).toBeGreaterThan(reviewTableIndex)
    expect(indexIndex).toBeGreaterThan(reconcileIndex)

    const reconcileQuery = queryable.calls[reconcileIndex]?.text ?? ''
    expect(reconcileQuery).toContain('FROM workbook_chat_item')
    expect(reconcileQuery).toContain('FROM workbook_review_queue_item')
    expect(reconcileQuery).toContain('COUNT(*)::bigint AS entry_count')
    expect(reconcileQuery).toContain('COUNT(*)::bigint AS review_queue_item_count')
    expect(reconcileQuery).toContain('ARRAY_AGG(text ORDER BY sort_order DESC, entry_id DESC)')
    expect(reconcileQuery).toContain('btrim(text) <>')
    expect(reconcileQuery).toContain('UPDATE workbook_chat_thread AS thread')
    expect(reconcileQuery).toContain('IS DISTINCT FROM thread_stats.entry_count')
    expect(reconcileQuery).toContain('IS DISTINCT FROM thread_stats.review_queue_item_count')
    expect(reconcileQuery).toContain('IS DISTINCT FROM thread_stats.latest_entry_text')
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
    expect(threadInsert?.values?.[7]).toBe(1)
    expect(threadInsert?.values?.[8]).toBe('Review item queued')
    expect(threadInsert?.values?.[9]).toBe(state.updatedAtUnixMs)
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

  it('saves thread snapshots atomically when the queryable supports transactions', async () => {
    const queryable = new FakeTransactionalQueryable()
    const state = createThreadState()

    await saveWorkbookAgentThreadState(queryable, state)

    expect(queryable.connectCount).toBe(1)
    expect(queryable.calls).toEqual([])
    expect(queryable.client.releaseCount).toBe(1)
    expect(queryable.client.calls[0]?.text).toBe('BEGIN')
    expect(queryable.client.calls.at(-1)?.text).toBe('COMMIT')
    expect(queryable.client.calls.some((call) => call.text.includes('INSERT INTO workbook_chat_thread'))).toBe(true)
    expect(queryable.client.calls.some((call) => call.text.includes('DELETE FROM workbook_chat_item'))).toBe(true)
    expect(queryable.client.calls.some((call) => call.text.includes('INSERT INTO workbook_review_queue_item'))).toBe(true)
  })

  it('rolls back the whole snapshot save when a transactional child write fails', async () => {
    const queryable = new FakeTransactionalQueryable('INSERT INTO workbook_review_queue_item')

    await expect(saveWorkbookAgentThreadState(queryable, createThreadState())).rejects.toThrow(
      'failed query: INSERT INTO workbook_review_queue_item',
    )

    expect(queryable.connectCount).toBe(1)
    expect(queryable.client.releaseCount).toBe(1)
    expect(queryable.client.calls[0]?.text).toBe('BEGIN')
    expect(queryable.client.calls.at(-1)?.text).toBe('ROLLBACK')
    expect(queryable.client.calls.some((call) => call.text.includes('COMMIT'))).toBe(false)
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
    const queryable = new FakeQueryable(
      [
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
      ],
      [createZeroThreadRow(state)],
    )

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
    const queryable = new FakeQueryable(
      [
        (text, values) =>
          text.includes('FROM workbook_chat_thread') && values?.[2] === 'alex@example.com'
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
      ],
      [createZeroThreadRow(state)],
    )

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
    const queryable = new FakeQueryable(
      [
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
      ],
      [createZeroThreadRow(state)],
    )

    const loaded = await loadWorkbookAgentThreadState(queryable, {
      documentId: 'doc-1',
      threadId: 'thr-1',
      actorUserId: 'alex@example.com',
    })

    expect(loaded?.entries.find((entry) => entry.id === 'tool-call-1')).toEqual(toolEntry)
  })

  it('lists durable thread summaries ordered by most recent activity', async () => {
    const queryable = new FakeQueryable(
      [
        (text) =>
          text.includes('FROM workbook_review_queue_item')
            ? [
                {
                  threadId: 'thr-1',
                  actorUserId: 'alex@example.com',
                  reviewQueueItemCount: 1,
                } satisfies QueryResultRow,
              ]
            : null,
      ],
      [
        {
          workbookId: 'doc-1',
          threadId: 'thr-2',
          ownerUserId: 'alex@example.com',
          scope: 'shared',
          updatedAtUnixMs: 200,
          entryCount: 3,
          reviewQueueItemCount: 0,
          latestEntryText: 'Applied shared cleanup at revision r7',
        },
        {
          workbookId: 'doc-1',
          threadId: 'thr-1',
          ownerUserId: 'alex@example.com',
          scope: 'private',
          updatedAtUnixMs: 100,
          entryCount: 1,
          reviewQueueItemCount: 1,
          latestEntryText: 'Review item queued',
        },
      ],
    )

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
    const queryable = new FakeQueryable(
      [],
      [
        {
          workbookId: 'doc-1',
          threadId: 'thr-shared',
          ownerUserId: 'alex@example.com',
          scope: 'shared',
          updatedAtUnixMs: 300,
          entryCount: 2,
          reviewQueueItemCount: 0,
          latestEntryText: 'Applied shared cleanup at revision r9',
        },
      ],
    )

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
    expect(queryable.zeroThreadInputs).toEqual([{ documentId: 'doc-1', actorUserId: 'casey@example.com' }])
  })
})
