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
import { handleWorkbookAgentConditionalFormatToolCall } from './workbook-agent-conditional-format-tools.js'

async function createEngine(): Promise<SpreadsheetEngine> {
  const engine = new SpreadsheetEngine({
    workbookName: 'doc-1',
    replicaId: 'server:conditional-format-test',
  })
  await engine.ready()
  engine.createSheet('Sheet1')
  engine.setCellValue('Sheet1', 'B2', 12)
  engine.setCellValue('Sheet1', 'B3', 8)
  engine.setCellValue('Sheet1', 'C4', 21)
  engine.setCellValue('Sheet1', 'E2', 'watch')
  engine.setConditionalFormat({
    id: 'cf-main',
    range: {
      sheetName: 'Sheet1',
      startAddress: 'B2',
      endAddress: 'C4',
    },
    rule: {
      kind: 'cellIs',
      operator: 'greaterThan',
      values: [10],
    },
    style: {
      fill: { backgroundColor: '#ff0000' },
    },
    stopIfTrue: true,
    priority: 7,
  })
  engine.setConditionalFormat({
    id: 'cf-outside',
    range: {
      sheetName: 'Sheet1',
      startAddress: 'E2',
      endAddress: 'E4',
    },
    rule: {
      kind: 'textContains',
      text: 'watch',
    },
    style: {
      font: { bold: true },
    },
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
    goalText: 'conditional format test',
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

const conditionalFormatListPayloadSchema = z.object({
  documentId: z.string(),
  conditionalFormatCount: z.number(),
  conditionalFormats: z.array(
    z.object({
      id: z.string(),
      range: z.object({
        sheetName: z.string(),
        startAddress: z.string(),
        endAddress: z.string(),
      }),
      rule: z.object({
        kind: z.string(),
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

describe('workbook agent conditional format tools', () => {
  it('filters conditional formats by intersecting normalized range targets', async () => {
    const engine = await createEngine()
    const { zeroSyncService } = createZeroSyncHarness(engine)

    const result = await handleWorkbookAgentConditionalFormatToolCall(
      {
        documentId: 'doc-1',
        session: { userID: 'alex@example.com', roles: ['editor'] },
        uiContext: null,
        zeroSyncService,
        stageCommand: vi.fn(async () =>
          createBundle({
            kind: 'deleteConditionalFormat',
            id: 'cf-main',
            range: { sheetName: 'Sheet1', startAddress: 'B2', endAddress: 'C4' },
          }),
        ),
      },
      {
        threadId: 'thr-1',
        turnId: 'turn-1',
        callId: 'call-get-conditional-formats',
        tool: WORKBOOK_AGENT_TOOL_NAMES.getConditionalFormats,
        arguments: {
          range: {
            sheetName: 'Sheet1',
            startAddress: 'C4',
            endAddress: 'B2',
          },
        },
      },
    )

    const payload = conditionalFormatListPayloadSchema.parse(parsePayload(result ?? { success: false, contentItems: [] }))
    expect(payload.conditionalFormatCount).toBe(1)
    expect(payload.conditionalFormats).toEqual([
      expect.objectContaining({
        id: 'cf-main',
        range: expect.objectContaining({
          startAddress: 'B2',
          endAddress: 'C4',
        }),
      }),
    ])
  })

  it('normalizes selector ranges and clears priority when staging conditional format mutations', async () => {
    const engine = await createEngine()
    const { zeroSyncService } = createZeroSyncHarness(engine)
    const stageCommand = vi.fn(async (command: WorkbookAgentCommand) => createBundle(command))

    const addResult = await handleWorkbookAgentConditionalFormatToolCall(
      {
        documentId: 'doc-1',
        session: { userID: 'alex@example.com', roles: ['editor'] },
        uiContext: {
          selection: {
            sheetName: 'Sheet1',
            address: 'B2',
            range: {
              startAddress: 'C4',
              endAddress: 'B2',
            },
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
        callId: 'call-add-conditional-format',
        tool: WORKBOOK_AGENT_TOOL_NAMES.addConditionalFormat,
        arguments: {
          selector: {
            kind: 'currentSelection',
          },
          rule: {
            kind: 'textContains',
            text: 'watch',
          },
          style: {
            fill: {
              backgroundColor: '#93c47d',
            },
            font: {
              bold: true,
              color: '#111827',
            },
            alignment: {
              horizontal: 'right',
              wrap: true,
            },
            borders: {
              top: {
                style: 'solid',
                weight: 'thin',
                color: '#111111',
              },
            },
          },
          stopIfTrue: false,
          priority: 3,
        },
      },
    )

    const addPayload = stagedPayloadSchema.parse(parsePayload(addResult ?? { success: false, contentItems: [] }))
    expect(addPayload.staged).toBe(true)
    expect(addPayload.reviewQueued).toBe(true)
    expect(stageCommand).toHaveBeenCalledWith({
      kind: 'upsertConditionalFormat',
      format: {
        id: expect.any(String),
        range: {
          sheetName: 'Sheet1',
          startAddress: 'B2',
          endAddress: 'C4',
        },
        rule: {
          kind: 'textContains',
          text: 'watch',
        },
        style: {
          fill: {
            backgroundColor: '#93c47d',
          },
          font: {
            bold: true,
            color: '#111827',
          },
          alignment: {
            horizontal: 'right',
            wrap: true,
          },
          borders: {
            top: {
              style: 'solid',
              weight: 'thin',
              color: '#111111',
            },
          },
        },
        stopIfTrue: false,
        priority: 3,
      },
    })

    await handleWorkbookAgentConditionalFormatToolCall(
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
        callId: 'call-update-conditional-format',
        tool: WORKBOOK_AGENT_TOOL_NAMES.updateConditionalFormat,
        arguments: {
          id: 'cf-main',
          range: {
            sheetName: 'Sheet1',
            startAddress: 'D5',
            endAddress: 'C3',
          },
          rule: {
            kind: 'formula',
            formula: 'C3>0',
          },
          style: {
            fill: null,
            borders: {
              bottom: {
                style: 'double',
                weight: 'medium',
                color: '#222222',
              },
            },
          },
          priority: null,
        },
      },
    )

    const updateCommand = stageCommand.mock.lastCall?.[0]
    expect(updateCommand).toMatchObject({
      kind: 'upsertConditionalFormat',
      format: {
        id: 'cf-main',
        range: {
          sheetName: 'Sheet1',
          startAddress: 'C3',
          endAddress: 'D5',
        },
        rule: {
          kind: 'formula',
          formula: 'C3>0',
        },
        style: {
          fill: null,
          borders: {
            bottom: {
              style: 'double',
              weight: 'medium',
              color: '#222222',
            },
          },
        },
        stopIfTrue: true,
      },
    })
    expect(updateCommand && updateCommand.kind === 'upsertConditionalFormat' ? updateCommand.format.priority : undefined).toBeUndefined()
  })

  it('stages conditional format removals with the stored target range', async () => {
    const engine = await createEngine()
    const { zeroSyncService } = createZeroSyncHarness(engine)
    const stageCommand = vi.fn(async (command: WorkbookAgentCommand) => createBundle(command))

    const result = await handleWorkbookAgentConditionalFormatToolCall(
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
        callId: 'call-remove-conditional-format',
        tool: WORKBOOK_AGENT_TOOL_NAMES.removeConditionalFormat,
        arguments: {
          id: 'cf-main',
        },
      },
    )

    const payload = stagedPayloadSchema.parse(parsePayload(result ?? { success: false, contentItems: [] }))
    expect(payload.staged).toBe(true)
    expect(stageCommand).toHaveBeenCalledWith({
      kind: 'deleteConditionalFormat',
      id: 'cf-main',
      range: {
        sheetName: 'Sheet1',
        startAddress: 'B2',
        endAddress: 'C4',
      },
    })
  })
})
