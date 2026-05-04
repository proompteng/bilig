import { SpreadsheetEngine } from '@bilig/core'
import {
  createWorkbookAgentCommandBundle,
  WORKBOOK_AGENT_TOOL_NAMES,
  type CodexDynamicToolCallResult,
  type WorkbookAgentCommand,
} from '@bilig/agent-api'
import { describe, expect, it, vi } from 'vitest'
import type { AuthoritativeWorkbookEventBatch } from '@bilig/zero-sync'
import { z } from 'zod'
import { buildWorkbookSourceProjectionFromEngine } from '../zero/projection.js'
import type { ZeroSyncService } from '../zero/service.js'
import type { WorkbookRuntime } from '../workbook-runtime/runtime-manager.js'
import { handleWorkbookAgentAnnotationToolCall } from './workbook-agent-annotation-tools.js'

async function createEngine(): Promise<SpreadsheetEngine> {
  const engine = new SpreadsheetEngine({
    workbookName: 'doc-1',
    replicaId: 'server:annotation-test',
  })
  await engine.ready()
  engine.createSheet('Sheet1')
  engine.setCellValue('Sheet1', 'B2', 'Draft')
  engine.setCellValue('Sheet1', 'C3', 'Alex')
  engine.setCellValue('Sheet1', 'E5', 'Outside')
  engine.setCommentThread({
    threadId: 'thread-1',
    sheetName: 'Sheet1',
    address: 'B2',
    comments: [{ id: 'comment-1', body: 'Review this total.' }],
  })
  engine.setCommentThread({
    threadId: 'thread-2',
    sheetName: 'Sheet1',
    address: 'E5',
    comments: [{ id: 'comment-2', body: 'Outside scope.' }],
  })
  engine.setNote({
    sheetName: 'Sheet1',
    address: 'C3',
    text: 'Manual override',
  })
  engine.setNote({
    sheetName: 'Sheet1',
    address: 'E5',
    text: 'Ignore me',
  })
  return engine
}

function createZeroSyncHarness(engine: SpreadsheetEngine) {
  const zeroSyncService: ZeroSyncService = {
    enabled: true,
    async initialize() {},
    async close() {},
    async handleQuery() {
      throw new Error('not used')
    },
    async handleMutate() {
      throw new Error('not used')
    },
    async inspectWorkbook(_documentId, task) {
      const runtime: WorkbookRuntime = {
        documentId: 'doc-1',
        engine,
        projection: buildWorkbookSourceProjectionFromEngine('doc-1', engine, {
          revision: 1,
          calculatedRevision: 1,
          ownerUserId: 'alex@example.com',
          updatedBy: 'alex@example.com',
          updatedAt: '2026-04-12T12:00:00.000Z',
        }),
        headRevision: 1,
        calculatedRevision: 1,
        ownerUserId: 'alex@example.com',
      }
      return await task(runtime)
    },
    async applyServerMutator() {
      throw new Error('not used')
    },
    async applyAgentCommandBundle() {
      throw new Error('not used')
    },
    async listWorkbookChanges() {
      return []
    },
    async listWorkbookAgentRuns() {
      return []
    },
    async appendWorkbookAgentRun() {
      throw new Error('not used')
    },
    async listWorkbookAgentThreadRuns() {
      return []
    },
    async listWorkbookAgentThreadSummaries() {
      return []
    },
    async loadWorkbookAgentThreadState() {
      return null
    },
    async saveWorkbookAgentThreadState() {
      throw new Error('not used')
    },
    async listWorkbookThreadWorkflowRuns() {
      return []
    },
    async upsertWorkbookWorkflowRun() {
      throw new Error('not used')
    },
    async getWorkbookHeadRevision() {
      return 1
    },
    async loadAuthoritativeEvents() {
      return {
        afterRevision: 1,
        headRevision: 1,
        calculatedRevision: 1,
        events: [],
      } satisfies AuthoritativeWorkbookEventBatch
    },
  }
  return { zeroSyncService }
}

function createBundle(command: WorkbookAgentCommand) {
  return createWorkbookAgentCommandBundle({
    documentId: 'doc-1',
    threadId: 'thr-1',
    turnId: 'turn-1',
    goalText: 'annotation test',
    baseRevision: 1,
    now: 1,
    context: null,
    commands: [command],
  })
}

function parsePayload(result: CodexDynamicToolCallResult): unknown {
  expect(result.success).toBe(true)
  const item = result.contentItems[0]
  expect(item?.type).toBe('inputText')
  return JSON.parse(item && 'text' in item ? item.text : '')
}

const annotationListPayloadSchema = z.object({
  documentId: z.string(),
  commentThreadCount: z.number(),
  noteCount: z.number(),
  commentThreads: z.array(
    z.object({
      threadId: z.string(),
      sheetName: z.string(),
      address: z.string(),
    }),
  ),
  notes: z.array(
    z.object({
      sheetName: z.string(),
      address: z.string(),
      text: z.string(),
    }),
  ),
})

const stagedPayloadSchema = z.object({
  staged: z.boolean(),
  reviewQueued: z.boolean(),
  bundleId: z.string(),
  mutationReceipt: z.object({
    status: z.string(),
  }),
})

describe('workbook agent annotation tools', () => {
  it('filters comments and notes by intersecting normalized range targets', async () => {
    const engine = await createEngine()
    const { zeroSyncService } = createZeroSyncHarness(engine)

    const result = await handleWorkbookAgentAnnotationToolCall(
      {
        documentId: 'doc-1',
        session: { userID: 'alex@example.com', roles: ['editor'] },
        uiContext: null,
        zeroSyncService,
        stageCommand: vi.fn(async () => createBundle({ kind: 'deleteCommentThread', sheetName: 'Sheet1', address: 'B2' })),
      },
      {
        threadId: 'thr-1',
        turnId: 'turn-1',
        callId: 'call-get-comments',
        tool: WORKBOOK_AGENT_TOOL_NAMES.getComments,
        arguments: {
          range: {
            sheetName: 'Sheet1',
            startAddress: 'D4',
            endAddress: 'B2',
          },
        },
      },
    )

    const payload = annotationListPayloadSchema.parse(parsePayload(result ?? { success: false, contentItems: [] }))
    expect(payload.commentThreadCount).toBe(1)
    expect(payload.noteCount).toBe(1)
    expect(payload.commentThreads).toEqual([
      expect.objectContaining({
        threadId: 'thread-1',
        address: 'B2',
      }),
    ])
    expect(payload.notes).toEqual([
      expect.objectContaining({
        address: 'C3',
        text: 'Manual override',
      }),
    ])
  })

  it('stages resolve, delete, update, and delete annotation mutations for single-cell targets', async () => {
    const engine = await createEngine()
    const { zeroSyncService } = createZeroSyncHarness(engine)
    const stageCommand = vi.fn(async (command: WorkbookAgentCommand) => createBundle(command))

    const resolveResult = await handleWorkbookAgentAnnotationToolCall(
      {
        documentId: 'doc-1',
        session: { userID: 'alex@example.com', roles: ['editor'] },
        uiContext: {
          selection: {
            sheetName: 'Sheet1',
            address: 'B2',
          },
          viewport: {
            rowStart: 0,
            rowEnd: 10,
            colStart: 0,
            colEnd: 5,
          },
        },
        zeroSyncService,
        stageCommand,
      },
      {
        threadId: 'thr-1',
        turnId: 'turn-1',
        callId: 'call-resolve-comment',
        tool: WORKBOOK_AGENT_TOOL_NAMES.resolveComment,
        arguments: {
          selector: {
            kind: 'currentSelection',
          },
        },
      },
    )

    const resolvePayload = stagedPayloadSchema.parse(parsePayload(resolveResult ?? { success: false, contentItems: [] }))
    expect(resolvePayload.staged).toBe(true)
    expect(resolvePayload.reviewQueued).toBe(true)
    expect(stageCommand).toHaveBeenCalledWith({
      kind: 'upsertCommentThread',
      thread: {
        threadId: 'thread-1',
        sheetName: 'Sheet1',
        address: 'B2',
        comments: [{ id: 'comment-1', body: 'Review this total.' }],
        resolved: true,
      },
    })

    await handleWorkbookAgentAnnotationToolCall(
      {
        documentId: 'doc-1',
        session: { userID: 'alex@example.com', roles: ['editor'] },
        uiContext: null,
        zeroSyncService,
        stageCommand,
      },
      {
        threadId: 'thr-1',
        turnId: 'turn-1',
        callId: 'call-delete-comment',
        tool: WORKBOOK_AGENT_TOOL_NAMES.deleteComment,
        arguments: {
          range: {
            sheetName: 'Sheet1',
            startAddress: 'B2',
            endAddress: 'B2',
          },
        },
      },
    )

    expect(stageCommand).toHaveBeenLastCalledWith({
      kind: 'deleteCommentThread',
      sheetName: 'Sheet1',
      address: 'B2',
    })

    await handleWorkbookAgentAnnotationToolCall(
      {
        documentId: 'doc-1',
        session: { userID: 'alex@example.com', roles: ['editor'] },
        uiContext: null,
        zeroSyncService,
        stageCommand,
      },
      {
        threadId: 'thr-1',
        turnId: 'turn-1',
        callId: 'call-update-note',
        tool: WORKBOOK_AGENT_TOOL_NAMES.updateNote,
        arguments: {
          range: {
            sheetName: 'Sheet1',
            startAddress: 'C3',
            endAddress: 'C3',
          },
          text: 'Pinned owner',
        },
      },
    )

    expect(stageCommand).toHaveBeenLastCalledWith({
      kind: 'upsertNote',
      note: {
        sheetName: 'Sheet1',
        address: 'C3',
        text: 'Pinned owner',
      },
    })

    await handleWorkbookAgentAnnotationToolCall(
      {
        documentId: 'doc-1',
        session: { userID: 'alex@example.com', roles: ['editor'] },
        uiContext: {
          selection: {
            sheetName: 'Sheet1',
            address: 'C3',
          },
          viewport: {
            rowStart: 0,
            rowEnd: 10,
            colStart: 0,
            colEnd: 5,
          },
        },
        zeroSyncService,
        stageCommand,
      },
      {
        threadId: 'thr-1',
        turnId: 'turn-1',
        callId: 'call-delete-note',
        tool: WORKBOOK_AGENT_TOOL_NAMES.deleteNote,
        arguments: {
          selector: {
            kind: 'currentSelection',
          },
        },
      },
    )

    expect(stageCommand).toHaveBeenLastCalledWith({
      kind: 'deleteNote',
      sheetName: 'Sheet1',
      address: 'C3',
    })
  })
})
