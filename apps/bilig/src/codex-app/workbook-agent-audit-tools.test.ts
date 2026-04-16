import { SpreadsheetEngine } from '@bilig/core'
import { WORKBOOK_AGENT_TOOL_NAMES, type CodexDynamicToolCallResult } from '@bilig/agent-api'
import { describe, expect, it, vi } from 'vitest'
import type { AuthoritativeWorkbookEventBatch } from '@bilig/zero-sync'
import { z } from 'zod'
import { buildWorkbookSourceProjectionFromEngine } from '../zero/projection.js'
import type { ZeroSyncService } from '../zero/service.js'
import type { WorkbookRuntime } from '../workbook-runtime/runtime-manager.js'
import { handleWorkbookAgentToolCall } from './workbook-agent-tools.js'

async function createAuditEngine(): Promise<SpreadsheetEngine> {
  const engine = new SpreadsheetEngine({
    workbookName: 'doc-1',
    replicaId: 'server:audit-test',
  })
  await engine.ready()
  engine.createSheet('Sheet1')
  engine.createSheet('Imports')
  engine.createSheet('Bloat')

  engine.setCellValue('Imports', 'A1', 9)
  engine.setCellFormula('Sheet1', 'E1', 'Imports!A1')
  engine.deleteSheet('Imports')

  engine.setCellValue('Sheet1', 'B3', 3)
  engine.setCellValue('Sheet1', 'B4', 4)
  engine.setCellValue('Sheet1', 'B5', 5)
  engine.setCellFormula('Sheet1', 'C3', 'B3*2')
  engine.setCellFormula('Sheet1', 'C4', 'B4*2')
  engine.setCellFormula('Sheet1', 'C5', 'B4*2')
  engine.setCellFormula('Sheet1', 'D6', 'SUM(B3:B5)')
  engine.setCellFormula('Sheet1', 'F1', 'LEN(B3:B5)')
  engine.updateRowMetadata('Sheet1', 3, 1, null, true)

  engine.setCellValue('Bloat', 'A1', 'anchor')
  engine.setRangeStyle(
    {
      sheetName: 'Bloat',
      startAddress: 'Z200',
      endAddress: 'Z200',
    },
    {
      fill: { backgroundColor: '#fee2e2' },
    },
  )
  engine.setImage({
    id: 'bloat-image',
    sheetName: 'Bloat',
    address: 'Y190',
    sourceUrl: 'https://example.com/bloat.png',
    rows: 3,
    cols: 2,
  })
  engine.setShape({
    id: 'bloat-shape',
    sheetName: 'Bloat',
    address: 'X180',
    shapeType: 'rectangle',
    rows: 2,
    cols: 4,
    fillColor: '#dbeafe',
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

async function callAuditTool(
  engine: SpreadsheetEngine,
  tool: string,
  args: Record<string, string | number | boolean | null> = {},
): Promise<unknown> {
  const { zeroSyncService } = createZeroSyncHarness(engine)
  const result = await handleWorkbookAgentToolCall(
    {
      documentId: 'doc-1',
      session: {
        userID: 'alex@example.com',
        roles: ['editor'],
      },
      uiContext: null,
      zeroSyncService,
      stageCommand: vi.fn(async () => {
        throw new Error('stageCommand should not be used by audit tools')
      }),
    },
    {
      threadId: 'thr-1',
      turnId: 'turn-1',
      callId: `call-${tool}`,
      tool,
      arguments: args,
    },
  )
  return parseTextToolPayload(result)
}

function parseTextToolPayload(result: CodexDynamicToolCallResult): unknown {
  expect(result.success).toBe(true)
  const item = result.contentItems[0]
  expect(item?.type).toBe('inputText')
  return JSON.parse(item && 'text' in item ? item.text : '')
}

const brokenReferencePayloadSchema = z.object({
  summary: z.object({
    brokenReferenceCount: z.number(),
  }),
  issues: z.array(
    z.object({
      sheetName: z.string(),
      address: z.string(),
      errorText: z.string().nullable(),
    }),
  ),
})

const hiddenRowPayloadSchema = z.object({
  summary: z.object({
    affectedFormulaCount: z.number(),
  }),
  issues: z.array(
    z.object({
      address: z.string(),
      hiddenPrecedents: z.array(
        z.object({
          sheetName: z.string(),
          address: z.string(),
          rowNumber: z.number(),
        }),
      ),
    }),
  ),
})

const inconsistentFormulaPayloadSchema = z.object({
  summary: z.object({
    inconsistentGroupCount: z.number(),
  }),
  groups: z.array(
    z.object({
      axis: z.string(),
      groupRange: z.object({
        startAddress: z.string(),
        endAddress: z.string(),
      }),
      outliers: z.array(
        z.object({
          address: z.string(),
          actualFormula: z.string(),
          expectedFormula: z.string(),
        }),
      ),
    }),
  ),
})

const usedRangeBloatPayloadSchema = z.object({
  sheets: z.array(
    z.object({
      sheetName: z.string(),
      compositeRange: z.object({
        endAddress: z.string(),
      }),
      drivers: z.array(
        z.object({
          source: z.string(),
        }),
      ),
    }),
  ),
})

const performanceHotspotPayloadSchema = z.object({
  hotspots: z.array(
    z.object({
      sheetName: z.string(),
      reasons: z.array(z.string()),
    }),
  ),
})

const invariantPayloadSchema = z.object({
  summary: z.object({
    ok: z.boolean(),
    roundTripStable: z.boolean(),
  }),
  problems: z.array(z.unknown()),
})

describe('workbook agent audit tools', () => {
  it('scans workbook formulas for broken references', async () => {
    const engine = await createAuditEngine()

    const payload = brokenReferencePayloadSchema.parse(await callAuditTool(engine, WORKBOOK_AGENT_TOOL_NAMES.scanBrokenReferences))

    expect(payload.summary.brokenReferenceCount).toBe(1)
    expect(payload.issues).toContainEqual(
      expect.objectContaining({
        sheetName: 'Sheet1',
        address: 'E1',
        errorText: '#REF!',
      }),
    )
  })

  it('finds formulas whose results depend on hidden rows', async () => {
    const engine = await createAuditEngine()

    const payload = hiddenRowPayloadSchema.parse(
      await callAuditTool(engine, WORKBOOK_AGENT_TOOL_NAMES.scanHiddenRowsAffectingResults, {
        sheetName: 'Sheet1',
      }),
    )

    expect(payload.summary.affectedFormulaCount).toBeGreaterThan(0)
    expect(payload.issues).toContainEqual(
      expect.objectContaining({
        address: 'D6',
        hiddenPrecedents: expect.arrayContaining([
          expect.objectContaining({
            sheetName: 'Sheet1',
            address: 'B4',
            rowNumber: 4,
          }),
        ]),
      }),
    )
  })

  it('detects inconsistent copied formulas in contiguous runs', async () => {
    const engine = await createAuditEngine()

    const payload = inconsistentFormulaPayloadSchema.parse(
      await callAuditTool(engine, WORKBOOK_AGENT_TOOL_NAMES.scanInconsistentFormulas, {
        sheetName: 'Sheet1',
      }),
    )

    expect(payload.summary.inconsistentGroupCount).toBeGreaterThan(0)
    expect(payload.groups).toContainEqual(
      expect.objectContaining({
        axis: 'column',
        groupRange: expect.objectContaining({
          startAddress: 'C3',
          endAddress: 'C5',
        }),
        outliers: expect.arrayContaining([
          expect.objectContaining({
            address: 'C5',
            actualFormula: '=B4*2',
            expectedFormula: '=B5*2',
          }),
        ]),
      }),
    )
  })

  it('reports used-range bloat from far metadata extents', async () => {
    const engine = await createAuditEngine()

    const payload = usedRangeBloatPayloadSchema.parse(await callAuditTool(engine, WORKBOOK_AGENT_TOOL_NAMES.scanUsedRangeBloat))

    expect(payload.sheets).toContainEqual(
      expect.objectContaining({
        sheetName: 'Bloat',
        compositeRange: expect.objectContaining({
          endAddress: 'AA200',
        }),
        drivers: expect.arrayContaining([
          expect.objectContaining({
            source: 'styleRange',
          }),
          expect.objectContaining({
            source: 'image',
          }),
          expect.objectContaining({
            source: 'shape',
          }),
        ]),
      }),
    )
  })

  it('ranks workbook performance hotspots with recalc metrics', async () => {
    const engine = await createAuditEngine()

    const payload = performanceHotspotPayloadSchema.parse(
      await callAuditTool(engine, WORKBOOK_AGENT_TOOL_NAMES.scanPerformanceHotspots, {
        limit: 5,
      }),
    )

    expect(payload.hotspots).toContainEqual(
      expect.objectContaining({
        sheetName: 'Sheet1',
        reasons: expect.arrayContaining([expect.stringContaining('JS-only formula')]),
      }),
    )
  })

  it('verifies workbook invariants and round-trip stability', async () => {
    const engine = await createAuditEngine()

    const payload = invariantPayloadSchema.parse(await callAuditTool(engine, WORKBOOK_AGENT_TOOL_NAMES.verifyInvariants))

    expect(payload.summary.ok).toBe(true)
    expect(payload.summary.roundTripStable).toBe(true)
    expect(payload.problems).toEqual([])
  })
})
