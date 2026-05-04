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
import { handleWorkbookAgentProtectionToolCall } from './workbook-agent-protection-tools.js'

async function createEngine(): Promise<SpreadsheetEngine> {
  const engine = new SpreadsheetEngine({
    workbookName: 'doc-1',
    replicaId: 'server:protection-test',
  })
  await engine.ready()
  engine.createSheet('Sheet1')
  engine.setCellFormula('Sheet1', 'B2', 'A1+1')
  engine.setRangeProtection({
    id: 'protect-main',
    range: {
      sheetName: 'Sheet1',
      startAddress: 'B2',
      endAddress: 'C4',
    },
    hideFormulas: true,
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
    goalText: 'protection test',
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

const protectionStatusPayloadSchema = z.object({
  sheetName: z.string(),
  sheetProtection: z.unknown().nullable(),
  rangeProtections: z.array(
    z.object({
      id: z.string(),
      range: z.object({
        sheetName: z.string(),
        startAddress: z.string(),
        endAddress: z.string(),
      }),
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

describe('workbook agent protection tools', () => {
  it('filters protection status by intersecting normalized range targets', async () => {
    const engine = await createEngine()
    const { zeroSyncService } = createZeroSyncHarness(engine)

    const result = await handleWorkbookAgentProtectionToolCall(
      {
        documentId: 'doc-1',
        session: { userID: 'alex@example.com', roles: ['editor'] },
        uiContext: null,
        zeroSyncService,
        stageCommand: vi.fn(async () => createBundle({ kind: 'clearSheetProtection', sheetName: 'Sheet1' })),
      },
      {
        threadId: 'thr-1',
        turnId: 'turn-1',
        callId: 'call-get-protection-status',
        tool: WORKBOOK_AGENT_TOOL_NAMES.getProtectionStatus,
        arguments: {
          range: {
            sheetName: 'Sheet1',
            startAddress: 'D4',
            endAddress: 'B2',
          },
        },
      },
    )

    const payload = protectionStatusPayloadSchema.parse(parsePayload(result ?? { success: false, contentItems: [] }))
    expect(payload.sheetName).toBe('Sheet1')
    expect(payload.rangeProtections).toEqual([
      expect.objectContaining({
        id: 'protect-main',
        range: expect.objectContaining({
          startAddress: 'B2',
          endAddress: 'C4',
        }),
      }),
    ])
  })

  it('resolves exact-match range removals through normalized range targets', async () => {
    const engine = await createEngine()
    const { zeroSyncService } = createZeroSyncHarness(engine)
    const stageCommand = vi.fn(async (command: WorkbookAgentCommand) => createBundle(command))

    const result = await handleWorkbookAgentProtectionToolCall(
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
        callId: 'call-unprotect-range',
        tool: WORKBOOK_AGENT_TOOL_NAMES.unprotectRange,
        arguments: {
          range: {
            sheetName: 'Sheet1',
            startAddress: 'C4',
            endAddress: 'B2',
          },
        },
      },
    )

    expect(stageCommand).toHaveBeenCalledWith({
      kind: 'deleteRangeProtection',
      id: 'protect-main',
      range: {
        sheetName: 'Sheet1',
        startAddress: 'B2',
        endAddress: 'C4',
      },
    })
    const payload = stagedPayloadSchema.parse(parsePayload(result ?? { success: false, contentItems: [] }))
    expect(payload.staged).toBe(true)
    expect(payload.reviewQueued).toBe(true)
  })
})
