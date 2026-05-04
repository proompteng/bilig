import { SpreadsheetEngine } from '@bilig/core'
import { WORKBOOK_AGENT_TOOL_NAMES, type CodexDynamicToolCallResult } from '@bilig/agent-api'
import { describe, expect, it } from 'vitest'
import type { AuthoritativeWorkbookEventBatch } from '@bilig/zero-sync'
import { z } from 'zod'
import type { WorkbookAgentUiContext } from '@bilig/contracts'
import { buildWorkbookSourceProjectionFromEngine } from '../zero/projection.js'
import type { ZeroSyncService } from '../zero/service.js'
import type { WorkbookRuntime } from '../workbook-runtime/runtime-manager.js'
import { handleWorkbookAgentSheetReadToolCall } from './workbook-agent-sheet-read-tools.js'

async function createEngine(): Promise<SpreadsheetEngine> {
  const engine = new SpreadsheetEngine({
    workbookName: 'doc-1',
    replicaId: 'server:sheet-read-test',
  })
  await engine.ready()
  engine.createSheet('Sheet1')
  engine.setCellValue('Sheet1', 'A1', 'Revenue')
  engine.setCellValue('Sheet1', 'B1', 'Margin')
  engine.setCellValue('Sheet1', 'A2', 10)
  engine.setCellValue('Sheet1', 'B2', 2)
  engine.setCellValue('Sheet1', 'A3', 12)
  engine.setCellValue('Sheet1', 'B3', 3)
  engine.setCellValue('Sheet1', 'D5', 'island')
  engine.updateRowMetadata('Sheet1', 1, 2, 28, true)
  engine.updateRowMetadata('Sheet1', 5, 1, 40, false)
  engine.updateColumnMetadata('Sheet1', 0, 1, 120, true)
  engine.updateColumnMetadata('Sheet1', 3, 2, 160, false)
  engine.setFreezePane('Sheet1', 1, 0)
  engine.setFilter('Sheet1', { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B3' })
  engine.setSort('Sheet1', { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B3' }, [{ keyAddress: 'B1', direction: 'desc' }])
  engine.createSheet('Ops Search')
  engine.setCellValue('Ops Search', 'C2', 'Northwind Import')
  engine.createSheet('Blank')
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

function parsePayload(result: CodexDynamicToolCallResult): unknown {
  expect(result.success).toBe(true)
  const item = result.contentItems[0]
  expect(item?.type).toBe('inputText')
  return JSON.parse(item && 'text' in item ? item.text : '')
}

const usedRangeSchema = z
  .object({
    startAddress: z.string(),
    endAddress: z.string(),
    rowCount: z.number(),
    columnCount: z.number(),
    cellCount: z.number(),
  })
  .nullable()

const listSheetsPayloadSchema = z.object({
  documentId: z.string(),
  sheetCount: z.number(),
  sheets: z.array(
    z.object({
      name: z.string(),
      order: z.number(),
      usedRange: usedRangeSchema,
    }),
  ),
})

const sheetViewPayloadSchema = z.object({
  documentId: z.string(),
  sheetName: z.string(),
  usedRange: usedRangeSchema,
  freezePane: z
    .object({
      sheetName: z.string(),
      rows: z.number(),
      cols: z.number(),
    })
    .nullable(),
  filters: z.array(
    z.object({
      sheetName: z.string(),
      range: z.object({
        sheetName: z.string(),
        startAddress: z.string(),
        endAddress: z.string(),
      }),
    }),
  ),
  sorts: z.array(
    z.object({
      sheetName: z.string(),
      range: z.object({
        sheetName: z.string(),
        startAddress: z.string(),
        endAddress: z.string(),
      }),
      keys: z.array(
        z.object({
          keyAddress: z.string(),
          direction: z.string(),
        }),
      ),
    }),
  ),
})

const usedRangePayloadSchema = z.object({
  documentId: z.string(),
  sheetName: z.string(),
  usedRange: usedRangeSchema,
})

const currentRegionPayloadSchema = z.object({
  documentId: z.string(),
  sheetName: z.string().nullable(),
  resolvedSelector: z.object({
    displayLabel: z.string(),
    derivedA1Ranges: z.array(
      z.object({
        sheetName: z.string(),
        startAddress: z.string(),
        endAddress: z.string(),
      }),
    ),
  }),
})

const axisMetadataPayloadSchema = z.object({
  documentId: z.string(),
  sheetName: z.string(),
  entryCount: z.number(),
  entries: z.array(
    z.object({
      start: z.number(),
      count: z.number(),
      size: z.number().nullable().optional(),
      hidden: z.boolean().nullable().optional(),
    }),
  ),
})

describe('workbook agent sheet read tools', () => {
  it('lists sheets in visible order and includes used ranges for populated sheets', async () => {
    const engine = await createEngine()
    const { zeroSyncService } = createZeroSyncHarness(engine)

    const result = await handleWorkbookAgentSheetReadToolCall(
      {
        documentId: 'doc-1',
        session: { userID: 'alex@example.com', roles: ['editor'] },
        uiContext: null,
        zeroSyncService,
      },
      {
        threadId: 'thr-1',
        turnId: 'turn-1',
        callId: 'call-list-sheets',
        tool: WORKBOOK_AGENT_TOOL_NAMES.listSheets,
        arguments: {},
      },
    )

    const payload = listSheetsPayloadSchema.parse(parsePayload(result ?? { success: false, contentItems: [] }))
    expect(payload.sheetCount).toBe(3)
    expect(payload.sheets.map((sheet) => sheet.name)).toEqual(['Sheet1', 'Ops Search', 'Blank'])
    expect(payload.sheets[0]?.usedRange).toEqual(
      expect.objectContaining({
        startAddress: 'A1',
        endAddress: 'D5',
      }),
    )
    expect(payload.sheets[1]?.usedRange).toEqual(
      expect.objectContaining({
        startAddress: 'C2',
        endAddress: 'C2',
      }),
    )
    expect(payload.sheets[2]?.usedRange).toBeNull()
  })

  it('uses browser context when sheet-level reads omit sheetName', async () => {
    const engine = await createEngine()
    const { zeroSyncService } = createZeroSyncHarness(engine)
    const uiContext: WorkbookAgentUiContext = {
      selection: {
        sheetName: 'Sheet1',
        address: 'B2',
      },
      viewport: {
        rowStart: 0,
        rowEnd: 6,
        colStart: 0,
        colEnd: 4,
      },
    }

    const sheetViewResult = await handleWorkbookAgentSheetReadToolCall(
      {
        documentId: 'doc-1',
        session: { userID: 'alex@example.com', roles: ['editor'] },
        uiContext,
        zeroSyncService,
      },
      {
        threadId: 'thr-1',
        turnId: 'turn-1',
        callId: 'call-sheet-view',
        tool: WORKBOOK_AGENT_TOOL_NAMES.getSheetView,
        arguments: {},
      },
    )
    const usedRangeResult = await handleWorkbookAgentSheetReadToolCall(
      {
        documentId: 'doc-1',
        session: { userID: 'alex@example.com', roles: ['editor'] },
        uiContext,
        zeroSyncService,
      },
      {
        threadId: 'thr-1',
        turnId: 'turn-1',
        callId: 'call-used-range',
        tool: WORKBOOK_AGENT_TOOL_NAMES.getUsedRange,
        arguments: {},
      },
    )

    const sheetViewPayload = sheetViewPayloadSchema.parse(parsePayload(sheetViewResult ?? { success: false, contentItems: [] }))
    expect(sheetViewPayload.sheetName).toBe('Sheet1')
    expect(sheetViewPayload.freezePane).toEqual({ sheetName: 'Sheet1', rows: 1, cols: 0 })
    expect(sheetViewPayload.filters).toHaveLength(1)
    expect(sheetViewPayload.sorts).toHaveLength(1)

    const usedRangePayload = usedRangePayloadSchema.parse(parsePayload(usedRangeResult ?? { success: false, contentItems: [] }))
    expect(usedRangePayload.sheetName).toBe('Sheet1')
    expect(usedRangePayload.usedRange).toEqual({
      startAddress: 'A1',
      endAddress: 'D5',
      rowCount: 5,
      columnCount: 4,
      cellCount: 20,
    })
  })

  it('reads current region from browser context and filters axis metadata windows', async () => {
    const engine = await createEngine()
    const { zeroSyncService } = createZeroSyncHarness(engine)
    const uiContext: WorkbookAgentUiContext = {
      selection: {
        sheetName: 'Sheet1',
        address: 'A2',
      },
      viewport: {
        rowStart: 0,
        rowEnd: 6,
        colStart: 0,
        colEnd: 4,
      },
    }

    const currentRegionResult = await handleWorkbookAgentSheetReadToolCall(
      {
        documentId: 'doc-1',
        session: { userID: 'alex@example.com', roles: ['editor'] },
        uiContext,
        zeroSyncService,
      },
      {
        threadId: 'thr-1',
        turnId: 'turn-1',
        callId: 'call-current-region',
        tool: WORKBOOK_AGENT_TOOL_NAMES.getCurrentRegion,
        arguments: {},
      },
    )
    const rowMetadataResult = await handleWorkbookAgentSheetReadToolCall(
      {
        documentId: 'doc-1',
        session: { userID: 'alex@example.com', roles: ['editor'] },
        uiContext,
        zeroSyncService,
      },
      {
        threadId: 'thr-1',
        turnId: 'turn-1',
        callId: 'call-row-metadata',
        tool: WORKBOOK_AGENT_TOOL_NAMES.getRowMetadata,
        arguments: {
          sheetName: 'Sheet1',
          startIndex: 2,
          count: 2,
        },
      },
    )
    const columnMetadataResult = await handleWorkbookAgentSheetReadToolCall(
      {
        documentId: 'doc-1',
        session: { userID: 'alex@example.com', roles: ['editor'] },
        uiContext,
        zeroSyncService,
      },
      {
        threadId: 'thr-1',
        turnId: 'turn-1',
        callId: 'call-column-metadata',
        tool: WORKBOOK_AGENT_TOOL_NAMES.getColumnMetadata,
        arguments: {
          sheetName: 'Sheet1',
          startIndex: 3,
          count: 2,
        },
      },
    )

    const currentRegionPayload = currentRegionPayloadSchema.parse(parsePayload(currentRegionResult ?? { success: false, contentItems: [] }))
    expect(currentRegionPayload.sheetName).toBe('Sheet1')
    expect(currentRegionPayload.resolvedSelector.derivedA1Ranges).toEqual([
      {
        sheetName: 'Sheet1',
        startAddress: 'A1',
        endAddress: 'B3',
      },
    ])

    const rowMetadataPayload = axisMetadataPayloadSchema.parse(parsePayload(rowMetadataResult ?? { success: false, contentItems: [] }))
    expect(rowMetadataPayload.entryCount).toBe(2)
    expect(rowMetadataPayload.entries).toEqual([
      expect.objectContaining({
        start: 1,
        count: 2,
        size: 28,
        hidden: true,
      }),
    ])

    const columnMetadataPayload = axisMetadataPayloadSchema.parse(
      parsePayload(columnMetadataResult ?? { success: false, contentItems: [] }),
    )
    expect(columnMetadataPayload.entryCount).toBe(2)
    expect(columnMetadataPayload.entries).toEqual([
      expect.objectContaining({
        start: 3,
        count: 2,
        size: 160,
        hidden: false,
      }),
    ])
  })
})
