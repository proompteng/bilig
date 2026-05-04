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
import { handleWorkbookAgentObjectToolCall } from './workbook-agent-object-tools.js'

async function createEngine(): Promise<SpreadsheetEngine> {
  const engine = new SpreadsheetEngine({
    workbookName: 'doc-1',
    replicaId: 'server:object-tools-test',
  })
  await engine.ready()
  engine.createSheet('Sheet1')
  engine.createSheet('Dashboard')
  engine.setRangeValues({ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B4' }, [
    ['Revenue', 'Margin'],
    [10, 2],
    [12, 3],
    [9, 1],
  ])
  engine.setTable({
    name: 'RevenueTable',
    sheetName: 'Sheet1',
    startAddress: 'A1',
    endAddress: 'B4',
    columnNames: ['Revenue', 'Margin'],
    headerRow: true,
    totalsRow: false,
  })
  engine.setDefinedName('ExistingFormula', {
    kind: 'formula',
    formula: '=SUM(Sheet1!A2:A4)',
  })
  engine.setPivotTable('Dashboard', 'E2', {
    name: 'RevenuePivot',
    source: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B4' },
    groupBy: ['Revenue'],
    values: [{ sourceColumn: 'Margin', summarizeBy: 'sum' }],
  })
  engine.setChart({
    id: 'Revenue Chart',
    sheetName: 'Dashboard',
    address: 'B2',
    source: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B4' },
    chartType: 'column',
    rows: 12,
    cols: 8,
    title: 'Revenue',
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
    goalText: 'object tools test',
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

const listPivotsPayloadSchema = z.object({
  documentId: z.string(),
  pivotCount: z.number(),
  pivots: z.array(
    z.object({
      name: z.string(),
      sheetName: z.string(),
      address: z.string(),
    }),
  ),
})

const listChartsPayloadSchema = z.object({
  documentId: z.string(),
  chartCount: z.number(),
  charts: z.array(
    z.object({
      id: z.string(),
      sheetName: z.string(),
      address: z.string(),
      chartType: z.string(),
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

describe('workbook agent object tools', () => {
  it('lists pivots and charts from the authoritative runtime', async () => {
    const engine = await createEngine()
    const { zeroSyncService } = createZeroSyncHarness(engine)

    const pivotsResult = await handleWorkbookAgentObjectToolCall(
      {
        documentId: 'doc-1',
        session: { userID: 'alex@example.com', roles: ['editor'] },
        uiContext: null,
        zeroSyncService,
        stageCommand: vi.fn(async () => createBundle({ kind: 'deletePivotTable', sheetName: 'Dashboard', address: 'E2' })),
      },
      {
        threadId: 'thr-1',
        turnId: 'turn-1',
        callId: 'call-list-pivots',
        tool: WORKBOOK_AGENT_TOOL_NAMES.listPivots,
        arguments: {},
      },
    )
    const chartsResult = await handleWorkbookAgentObjectToolCall(
      {
        documentId: 'doc-1',
        session: { userID: 'alex@example.com', roles: ['editor'] },
        uiContext: null,
        zeroSyncService,
        stageCommand: vi.fn(async () => createBundle({ kind: 'deleteChart', id: 'Revenue Chart' })),
      },
      {
        threadId: 'thr-1',
        turnId: 'turn-1',
        callId: 'call-list-charts',
        tool: WORKBOOK_AGENT_TOOL_NAMES.listCharts,
        arguments: {},
      },
    )

    const pivotsPayload = listPivotsPayloadSchema.parse(parsePayload(pivotsResult ?? { success: false, contentItems: [] }))
    expect(pivotsPayload.pivotCount).toBe(1)
    expect(pivotsPayload.pivots).toContainEqual(
      expect.objectContaining({
        name: 'RevenuePivot',
        sheetName: 'Dashboard',
        address: 'E2',
      }),
    )

    const chartsPayload = listChartsPayloadSchema.parse(parsePayload(chartsResult ?? { success: false, contentItems: [] }))
    expect(chartsPayload.chartCount).toBe(1)
    expect(chartsPayload.charts).toContainEqual(
      expect.objectContaining({
        id: 'Revenue Chart',
        sheetName: 'Dashboard',
        address: 'B2',
        chartType: 'column',
      }),
    )
  })

  it('stages selector-aware named range, table, pivot, and chart updates', async () => {
    const engine = await createEngine()
    const { zeroSyncService } = createZeroSyncHarness(engine)
    const stageCommand = vi.fn(async (command: WorkbookAgentCommand) => createBundle(command))

    const updateNamedRangeResult = await handleWorkbookAgentObjectToolCall(
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
        callId: 'call-update-named-range',
        tool: WORKBOOK_AGENT_TOOL_NAMES.updateNamedRange,
        arguments: {
          name: 'RevenueFormula',
          value: {
            kind: 'formula',
            formula: 'SUM(Sheet1!A2:A4)',
          },
        },
      },
    )
    const resizeTableResult = await handleWorkbookAgentObjectToolCall(
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
        callId: 'call-resize-table',
        tool: WORKBOOK_AGENT_TOOL_NAMES.resizeTable,
        arguments: {
          name: 'RevenueTable',
          selector: {
            kind: 'table',
            table: 'RevenueTable',
          },
          totalsRow: true,
        },
      },
    )
    const updatePivotResult = await handleWorkbookAgentObjectToolCall(
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
        callId: 'call-update-pivot',
        tool: WORKBOOK_AGENT_TOOL_NAMES.updatePivotTable,
        arguments: {
          name: 'RevenuePivot',
          sheetName: 'Dashboard',
          address: 'G4',
          selector: {
            kind: 'table',
            table: 'RevenueTable',
          },
          groupBy: ['Revenue'],
          values: [{ sourceColumn: 'Margin', summarizeBy: 'count', outputLabel: 'Margin Count' }],
        },
      },
    )
    const updateChartResult = await handleWorkbookAgentObjectToolCall(
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
        callId: 'call-update-chart',
        tool: WORKBOOK_AGENT_TOOL_NAMES.updateChart,
        arguments: {
          id: 'Revenue Chart',
          sheetName: 'Dashboard',
          address: 'J3',
          selector: {
            kind: 'tableColumn',
            table: 'RevenueTable',
            column: 'Margin',
          },
          chartType: 'line',
          rows: 10,
          cols: 6,
          title: 'Margin Trend',
          seriesOrientation: 'columns',
          firstRowAsHeaders: true,
          firstColumnAsLabels: false,
          legendPosition: 'bottom',
        },
      },
    )

    expect(stagedPayloadSchema.parse(parsePayload(updateNamedRangeResult ?? { success: false, contentItems: [] }))).toEqual(
      expect.objectContaining({ staged: true, reviewQueued: true }),
    )
    expect(stageCommand).toHaveBeenNthCalledWith(1, {
      kind: 'upsertDefinedName',
      name: 'RevenueFormula',
      value: {
        kind: 'formula',
        formula: '=SUM(Sheet1!A2:A4)',
      },
    })

    expect(stagedPayloadSchema.parse(parsePayload(resizeTableResult ?? { success: false, contentItems: [] }))).toEqual(
      expect.objectContaining({ staged: true, reviewQueued: true }),
    )
    expect(stageCommand).toHaveBeenNthCalledWith(2, {
      kind: 'upsertTable',
      table: {
        name: 'RevenueTable',
        sheetName: 'Sheet1',
        startAddress: 'A1',
        endAddress: 'B4',
        columnNames: ['Revenue', 'Margin'],
        headerRow: true,
        totalsRow: true,
      },
    })

    expect(stagedPayloadSchema.parse(parsePayload(updatePivotResult ?? { success: false, contentItems: [] }))).toEqual(
      expect.objectContaining({ staged: true, reviewQueued: true }),
    )
    expect(stageCommand).toHaveBeenNthCalledWith(3, {
      kind: 'upsertPivotTable',
      pivot: {
        name: 'RevenuePivot',
        sheetName: 'Dashboard',
        address: 'G4',
        source: {
          sheetName: 'Sheet1',
          startAddress: 'A1',
          endAddress: 'B4',
        },
        groupBy: ['Revenue'],
        values: [{ sourceColumn: 'Margin', summarizeBy: 'count', outputLabel: 'Margin Count' }],
        rows: 1,
        cols: 2,
      },
    })

    expect(stagedPayloadSchema.parse(parsePayload(updateChartResult ?? { success: false, contentItems: [] }))).toEqual(
      expect.objectContaining({ staged: true, reviewQueued: true }),
    )
    expect(stageCommand).toHaveBeenNthCalledWith(4, {
      kind: 'upsertChart',
      chart: {
        id: 'Revenue Chart',
        sheetName: 'Dashboard',
        address: 'J3',
        source: {
          sheetName: 'Sheet1',
          startAddress: 'B2',
          endAddress: 'B4',
        },
        chartType: 'line',
        rows: 10,
        cols: 6,
        title: 'Margin Trend',
        seriesOrientation: 'columns',
        firstRowAsHeaders: true,
        firstColumnAsLabels: false,
        legendPosition: 'bottom',
      },
    })
  })

  it('stages object deletions by name, id, and resolved pivot target', async () => {
    const engine = await createEngine()
    const { zeroSyncService } = createZeroSyncHarness(engine)
    const stageCommand = vi.fn(async (command: WorkbookAgentCommand) => createBundle(command))

    const deleteNamedRangeResult = await handleWorkbookAgentObjectToolCall(
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
        callId: 'call-delete-named-range',
        tool: WORKBOOK_AGENT_TOOL_NAMES.deleteNamedRange,
        arguments: {
          name: 'ExistingFormula',
        },
      },
    )
    const deleteTableResult = await handleWorkbookAgentObjectToolCall(
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
        callId: 'call-delete-table',
        tool: WORKBOOK_AGENT_TOOL_NAMES.deleteTable,
        arguments: {
          name: 'RevenueTable',
        },
      },
    )
    const deletePivotResult = await handleWorkbookAgentObjectToolCall(
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
        callId: 'call-delete-pivot',
        tool: WORKBOOK_AGENT_TOOL_NAMES.deletePivotTable,
        arguments: {
          name: 'RevenuePivot',
        },
      },
    )
    const deleteChartResult = await handleWorkbookAgentObjectToolCall(
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
        callId: 'call-delete-chart',
        tool: WORKBOOK_AGENT_TOOL_NAMES.deleteChart,
        arguments: {
          id: 'Revenue Chart',
        },
      },
    )

    expect(stagedPayloadSchema.parse(parsePayload(deleteNamedRangeResult ?? { success: false, contentItems: [] }))).toEqual(
      expect.objectContaining({ staged: true }),
    )
    expect(stagedPayloadSchema.parse(parsePayload(deleteTableResult ?? { success: false, contentItems: [] }))).toEqual(
      expect.objectContaining({ staged: true }),
    )
    expect(stagedPayloadSchema.parse(parsePayload(deletePivotResult ?? { success: false, contentItems: [] }))).toEqual(
      expect.objectContaining({ staged: true }),
    )
    expect(stagedPayloadSchema.parse(parsePayload(deleteChartResult ?? { success: false, contentItems: [] }))).toEqual(
      expect.objectContaining({ staged: true }),
    )

    expect(stageCommand).toHaveBeenNthCalledWith(1, {
      kind: 'deleteDefinedName',
      name: 'ExistingFormula',
    })
    expect(stageCommand).toHaveBeenNthCalledWith(2, {
      kind: 'deleteTable',
      name: 'RevenueTable',
    })
    expect(stageCommand).toHaveBeenNthCalledWith(3, {
      kind: 'deletePivotTable',
      sheetName: 'Dashboard',
      address: 'E2',
    })
    expect(stageCommand).toHaveBeenNthCalledWith(4, {
      kind: 'deleteChart',
      id: 'Revenue Chart',
    })
  })
})
