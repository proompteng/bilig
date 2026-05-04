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
import { handleWorkbookAgentValidationToolCall } from './workbook-agent-validation-tools.js'

async function createEngine(): Promise<SpreadsheetEngine> {
  const engine = new SpreadsheetEngine({
    workbookName: 'doc-1',
    replicaId: 'server:validation-test',
  })
  await engine.ready()
  engine.createSheet('Sheet1')
  engine.setCellValue('Sheet1', 'A1', 'Status')
  engine.setCellValue('Sheet1', 'B1', 'Owner')
  engine.setCellValue('Sheet1', 'A2', 'Draft')
  engine.setCellValue('Sheet1', 'B2', 'Alex')
  engine.setCellValue('Sheet1', 'A3', 'Final')
  engine.setCellValue('Sheet1', 'B3', 'Pat')
  engine.setDefinedName('StatusCells', {
    kind: 'range-ref',
    sheetName: 'Sheet1',
    startAddress: 'B3',
    endAddress: 'A2',
  })
  engine.setDataValidation({
    range: {
      sheetName: 'Sheet1',
      startAddress: 'A2',
      endAddress: 'A3',
    },
    rule: {
      kind: 'list',
      values: ['Draft', 'Final'],
    },
    allowBlank: false,
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
    goalText: 'validation test',
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

const validationListPayloadSchema = z.object({
  documentId: z.string(),
  validationCount: z.number(),
  validations: z.array(
    z.object({
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

describe('workbook agent validation tools', () => {
  it('filters validation rules by intersecting normalized range targets', async () => {
    const engine = await createEngine()
    const { zeroSyncService } = createZeroSyncHarness(engine)

    const result = await handleWorkbookAgentValidationToolCall(
      {
        documentId: 'doc-1',
        session: { userID: 'alex@example.com', roles: ['editor'] },
        uiContext: null,
        zeroSyncService,
        stageCommand: vi.fn(async () =>
          createBundle({ kind: 'clearDataValidation', range: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' } }),
        ),
      },
      {
        threadId: 'thr-1',
        turnId: 'turn-1',
        callId: 'call-list-validations',
        tool: WORKBOOK_AGENT_TOOL_NAMES.listDataValidationRules,
        arguments: {
          range: {
            sheetName: 'Sheet1',
            startAddress: 'A3',
            endAddress: 'A2',
          },
        },
      },
    )

    const payload = validationListPayloadSchema.parse(parsePayload(result ?? { success: false, contentItems: [] }))
    expect(payload.validationCount).toBe(1)
    expect(payload.validations).toEqual([
      expect.objectContaining({
        range: expect.objectContaining({
          startAddress: 'A2',
          endAddress: 'A3',
        }),
        rule: expect.objectContaining({
          kind: 'list',
        }),
      }),
    ])
  })

  it('normalizes selector ranges before staging create and remove validation commands', async () => {
    const engine = await createEngine()
    const { zeroSyncService } = createZeroSyncHarness(engine)
    const stageCommand = vi.fn(async (command: WorkbookAgentCommand) => createBundle(command))

    await handleWorkbookAgentValidationToolCall(
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
        callId: 'call-create-validation',
        tool: WORKBOOK_AGENT_TOOL_NAMES.createDataValidation,
        arguments: {
          selector: {
            kind: 'namedRange',
            name: 'StatusCells',
          },
          rule: {
            kind: 'checkbox',
            checkedValue: true,
            uncheckedValue: false,
          },
        },
      },
    )

    expect(stageCommand).toHaveBeenCalledWith({
      kind: 'setDataValidation',
      validation: {
        range: {
          sheetName: 'Sheet1',
          startAddress: 'A2',
          endAddress: 'B3',
        },
        rule: {
          kind: 'checkbox',
          checkedValue: true,
          uncheckedValue: false,
        },
      },
    })

    await handleWorkbookAgentValidationToolCall(
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
        callId: 'call-remove-validation',
        tool: WORKBOOK_AGENT_TOOL_NAMES.removeDataValidation,
        arguments: {
          range: {
            sheetName: 'Sheet1',
            startAddress: 'A3',
            endAddress: 'A2',
          },
        },
      },
    )

    expect(stageCommand).toHaveBeenLastCalledWith({
      kind: 'clearDataValidation',
      range: {
        sheetName: 'Sheet1',
        startAddress: 'A2',
        endAddress: 'A3',
      },
    })
  })
})
