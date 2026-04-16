import { SpreadsheetEngine } from '@bilig/core'
import { describe, expect, it } from 'vitest'
import { buildWorkbookSourceProjectionFromEngine } from '../zero/projection.js'
import type { ZeroSyncService } from '../zero/service.js'
import type { WorkbookRuntime } from '../workbook-runtime/runtime-manager.js'
import { describeWorkbookAgentWorkflowTemplate, executeWorkbookAgentWorkflow } from './workbook-agent-workflows.js'

async function createWorkbookRuntime(): Promise<WorkbookRuntime> {
  const engine = new SpreadsheetEngine({
    workbookName: 'doc-1',
    replicaId: 'server:test',
  })
  await engine.ready()
  engine.createSheet('Sheet1')
  engine.setCellValue('Sheet1', 'A1', 42)
  engine.setCellValue('Sheet1', 'A2', 'Gross Margin')
  engine.setCellFormula('Sheet1', 'B2', 'SUM(A1:A1)')
  return {
    documentId: 'doc-1',
    engine,
    projection: buildWorkbookSourceProjectionFromEngine('doc-1', engine, {
      revision: 1,
      calculatedRevision: 1,
      ownerUserId: 'alex@example.com',
      updatedBy: 'alex@example.com',
      updatedAt: '2026-04-10T00:00:00.000Z',
    }),
    headRevision: 1,
    calculatedRevision: 1,
    ownerUserId: 'alex@example.com',
  }
}

function createZeroSyncStub(input?: { onInspectWorkbook?: () => void; createRuntime?: () => Promise<WorkbookRuntime> }): ZeroSyncService {
  return {
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
      input?.onInspectWorkbook?.()
      if (input?.createRuntime) {
        return await task(await input.createRuntime())
      }
      throw new Error('inspectWorkbook should not be called')
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
    async listWorkbookAgentThreadRuns() {
      return []
    },
    async appendWorkbookAgentRun() {
      throw new Error('not used')
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
      throw new Error('not used')
    },
  }
}

describe('workbook agent workflows', () => {
  it('describes structural workflow templates through the structural metadata path', () => {
    expect(
      describeWorkbookAgentWorkflowTemplate('createSheet', {
        name: 'Ops Review',
      }),
    ).toEqual({
      title: 'Create Sheet',
      runningSummary: 'Preparing a structural workbook change set to create Ops Review.',
    })
  })

  it('describes number-format normalization through the import metadata path', () => {
    expect(
      describeWorkbookAgentWorkflowTemplate('normalizeCurrentSheetNumberFormats', {
        sheetName: 'Imports',
      }),
    ).toEqual({
      title: 'Normalize Current Sheet Number Formats',
      runningSummary: 'Running number-format normalization workflow for Imports.',
    })
  })

  it('describes outlier highlighting through the formatting metadata path', () => {
    expect(
      describeWorkbookAgentWorkflowTemplate('highlightCurrentSheetOutliers', {
        sheetName: 'Revenue',
      }),
    ).toEqual({
      title: 'Highlight Current Sheet Outliers',
      runningSummary: 'Running outlier highlight workflow for Revenue.',
    })
  })

  it('executes structural workflow templates without durable workbook inspection', async () => {
    let inspectedWorkbook = false
    const result = await executeWorkbookAgentWorkflow({
      documentId: 'doc-1',
      zeroSyncService: createZeroSyncStub({
        onInspectWorkbook: () => {
          inspectedWorkbook = true
        },
      }),
      workflowTemplate: 'hideCurrentRow',
      context: {
        selection: {
          sheetName: 'Sheet1',
          address: 'B3',
        },
        viewport: {
          rowStart: 0,
          rowEnd: 20,
          colStart: 0,
          colEnd: 10,
        },
      },
    })

    expect(inspectedWorkbook).toBe(false)
    expect(result.title).toBe('Hide Current Row')
    expect(result.commands).toEqual([
      {
        kind: 'updateRowMetadata',
        sheetName: 'Sheet1',
        startRow: 2,
        count: 1,
        hidden: true,
      },
    ])
  })

  it('executes unhide structural workflow templates without durable workbook inspection', async () => {
    let inspectedWorkbook = false
    const result = await executeWorkbookAgentWorkflow({
      documentId: 'doc-1',
      zeroSyncService: createZeroSyncStub({
        onInspectWorkbook: () => {
          inspectedWorkbook = true
        },
      }),
      workflowTemplate: 'unhideCurrentColumn',
      context: {
        selection: {
          sheetName: 'Sheet1',
          address: 'C3',
        },
        viewport: {
          rowStart: 0,
          rowEnd: 20,
          colStart: 0,
          colEnd: 10,
        },
      },
    })

    expect(inspectedWorkbook).toBe(false)
    expect(result.title).toBe('Unhide Current Column')
    expect(result.commands).toEqual([
      {
        kind: 'updateColumnMetadata',
        sheetName: 'Sheet1',
        startCol: 2,
        count: 1,
        hidden: false,
      },
    ])
  })

  it('executes summarize workbook through the durable inspection path', async () => {
    let inspectedWorkbook = false
    const result = await executeWorkbookAgentWorkflow({
      documentId: 'doc-1',
      zeroSyncService: createZeroSyncStub({
        onInspectWorkbook: () => {
          inspectedWorkbook = true
        },
        createRuntime: createWorkbookRuntime,
      }),
      workflowTemplate: 'summarizeWorkbook',
    })

    expect(inspectedWorkbook).toBe(true)
    expect(result.title).toBe('Summarize Workbook')
    expect(result.summary).toContain('Summarized workbook structure across 1 sheet')
  })

  it('executes current-sheet number-format normalization through the durable inspection path', async () => {
    const result = await executeWorkbookAgentWorkflow({
      documentId: 'doc-1',
      zeroSyncService: createZeroSyncStub({
        createRuntime: async () => {
          const engine = new SpreadsheetEngine({
            workbookName: 'doc-1',
            replicaId: 'server:test',
          })
          await engine.ready()
          engine.createSheet('Imports')
          engine.setCellValue('Imports', 'A1', 'order_date')
          engine.setCellValue('Imports', 'B1', 'gross_margin_pct')
          engine.setCellValue('Imports', 'C1', 'revenue_usd')
          engine.setCellValue('Imports', 'A2', 45292)
          engine.setCellValue('Imports', 'A3', 45293)
          engine.setCellValue('Imports', 'B2', 0.42)
          engine.setCellValue('Imports', 'B3', 0.31)
          engine.setCellValue('Imports', 'C2', 1200)
          engine.setCellValue('Imports', 'C3', 950.5)
          return {
            documentId: 'doc-1',
            engine,
            projection: buildWorkbookSourceProjectionFromEngine('doc-1', engine, {
              revision: 1,
              calculatedRevision: 1,
              ownerUserId: 'alex@example.com',
              updatedBy: 'alex@example.com',
              updatedAt: '2026-04-10T00:00:00.000Z',
            }),
            headRevision: 1,
            calculatedRevision: 1,
            ownerUserId: 'alex@example.com',
          }
        },
      }),
      workflowTemplate: 'normalizeCurrentSheetNumberFormats',
      context: {
        selection: {
          sheetName: 'Imports',
          address: 'A1',
        },
        viewport: {
          rowStart: 0,
          rowEnd: 20,
          colStart: 0,
          colEnd: 10,
        },
      },
      workflowInput: {
        sheetName: 'Imports',
      },
    })

    expect(result.title).toBe('Normalize Current Sheet Number Formats')
    expect(result.summary).toContain('Staged normalized number formats')
    expect(result.commands).toEqual([
      expect.objectContaining({
        kind: 'formatRange',
        range: {
          sheetName: 'Imports',
          startAddress: 'A2',
          endAddress: 'A3',
        },
        numberFormat: expect.objectContaining({
          kind: 'date',
        }),
      }),
      expect.objectContaining({
        kind: 'formatRange',
        range: {
          sheetName: 'Imports',
          startAddress: 'B2',
          endAddress: 'B3',
        },
        numberFormat: expect.objectContaining({
          kind: 'percent',
        }),
      }),
      expect.objectContaining({
        kind: 'formatRange',
        range: {
          sheetName: 'Imports',
          startAddress: 'C2',
          endAddress: 'C3',
        },
        numberFormat: expect.objectContaining({
          kind: 'currency',
        }),
      }),
    ])
  })

  it('executes current-sheet rollup workflows through the durable inspection path', async () => {
    const result = await executeWorkbookAgentWorkflow({
      documentId: 'doc-1',
      zeroSyncService: createZeroSyncStub({
        createRuntime: async () => {
          const engine = new SpreadsheetEngine({
            workbookName: 'doc-1',
            replicaId: 'server:test',
          })
          await engine.ready()
          engine.createSheet('Revenue')
          engine.setCellValue('Revenue', 'A1', 'Region')
          engine.setCellValue('Revenue', 'B1', 'January')
          engine.setCellValue('Revenue', 'C1', 'February')
          engine.setCellValue('Revenue', 'D1', 'Gross Margin')
          engine.setCellValue('Revenue', 'A2', 'West')
          engine.setCellValue('Revenue', 'B2', 100)
          engine.setCellValue('Revenue', 'C2', 120)
          engine.setCellValue('Revenue', 'D2', 0.4)
          engine.setCellValue('Revenue', 'A3', 'East')
          engine.setCellValue('Revenue', 'B3', 200)
          engine.setCellValue('Revenue', 'C3', 240)
          engine.setCellValue('Revenue', 'D3', 0.35)
          return {
            documentId: 'doc-1',
            engine,
            projection: buildWorkbookSourceProjectionFromEngine('doc-1', engine, {
              revision: 1,
              calculatedRevision: 1,
              ownerUserId: 'alex@example.com',
              updatedBy: 'alex@example.com',
              updatedAt: '2026-04-10T00:00:00.000Z',
            }),
            headRevision: 1,
            calculatedRevision: 1,
            ownerUserId: 'alex@example.com',
          }
        },
      }),
      workflowTemplate: 'createCurrentSheetRollup',
      context: {
        selection: {
          sheetName: 'Revenue',
          address: 'A1',
        },
        viewport: {
          rowStart: 0,
          rowEnd: 20,
          colStart: 0,
          colEnd: 10,
        },
      },
      workflowInput: {
        sheetName: 'Revenue',
      },
    })

    expect(result.title).toBe('Create Current Sheet Rollup')
    expect(result.summary).toContain('Revenue Rollup')
    expect(result.commands).toEqual([
      {
        kind: 'createSheet',
        name: 'Revenue Rollup',
      },
      expect.objectContaining({
        kind: 'writeRange',
        sheetName: 'Revenue Rollup',
        startAddress: 'A1',
        values: expect.arrayContaining([expect.arrayContaining(['Metric', 'Value']), expect.arrayContaining(['Source Sheet', 'Revenue'])]),
      }),
    ])
  })

  it('executes whitespace-normalization workflows through the durable inspection path', async () => {
    const result = await executeWorkbookAgentWorkflow({
      documentId: 'doc-1',
      zeroSyncService: createZeroSyncStub({
        createRuntime: async () => {
          const engine = new SpreadsheetEngine({
            workbookName: 'doc-1',
            replicaId: 'server:test',
          })
          await engine.ready()
          engine.createSheet('Imports')
          engine.setCellValue('Imports', 'A1', ' Customer   Name ')
          engine.setCellValue('Imports', 'B1', 'Notes')
          engine.setCellValue('Imports', 'A2', '  Ada   Lovelace  ')
          engine.setCellValue('Imports', 'B2', '  First\tentry  ')
          engine.setCellValue('Imports', 'A3', 'Grace Hopper')
          engine.setCellValue('Imports', 'B3', 'Already clean')
          return {
            documentId: 'doc-1',
            engine,
            projection: buildWorkbookSourceProjectionFromEngine('doc-1', engine, {
              revision: 1,
              calculatedRevision: 1,
              ownerUserId: 'alex@example.com',
              updatedBy: 'alex@example.com',
              updatedAt: '2026-04-10T00:00:00.000Z',
            }),
            headRevision: 1,
            calculatedRevision: 1,
            ownerUserId: 'alex@example.com',
          }
        },
      }),
      workflowTemplate: 'normalizeCurrentSheetWhitespace',
      context: {
        selection: {
          sheetName: 'Imports',
          address: 'A1',
        },
        viewport: {
          rowStart: 0,
          rowEnd: 20,
          colStart: 0,
          colEnd: 10,
        },
      },
      workflowInput: {
        sheetName: 'Imports',
      },
    })

    expect(result.title).toBe('Normalize Current Sheet Whitespace')
    expect(result.summary).toContain('Staged normalized whitespace')
    expect(result.commands).toEqual([
      expect.objectContaining({
        kind: 'writeRange',
        sheetName: 'Imports',
        startAddress: 'A1',
        values: [
          ['Customer Name', 'Notes'],
          ['Ada Lovelace', 'First entry'],
          ['Grace Hopper', 'Already clean'],
        ],
      }),
    ])
  })

  it('executes formula fill-down workflows through the durable inspection path', async () => {
    const result = await executeWorkbookAgentWorkflow({
      documentId: 'doc-1',
      zeroSyncService: createZeroSyncStub({
        createRuntime: async () => {
          const engine = new SpreadsheetEngine({
            workbookName: 'doc-1',
            replicaId: 'server:test',
          })
          await engine.ready()
          engine.createSheet('Imports')
          engine.setCellValue('Imports', 'A1', 'Revenue')
          engine.setCellValue('Imports', 'B1', 'Tax')
          engine.setCellValue('Imports', 'A2', 100)
          engine.setCellValue('Imports', 'A3', 120)
          engine.setCellValue('Imports', 'A4', 140)
          engine.setCellFormula('Imports', 'B2', 'A2*0.1')
          return {
            documentId: 'doc-1',
            engine,
            projection: buildWorkbookSourceProjectionFromEngine('doc-1', engine, {
              revision: 1,
              calculatedRevision: 1,
              ownerUserId: 'alex@example.com',
              updatedBy: 'alex@example.com',
              updatedAt: '2026-04-10T00:00:00.000Z',
            }),
            headRevision: 1,
            calculatedRevision: 1,
            ownerUserId: 'alex@example.com',
          }
        },
      }),
      workflowTemplate: 'fillCurrentSheetFormulasDown',
      context: {
        selection: {
          sheetName: 'Imports',
          address: 'A1',
        },
        viewport: {
          rowStart: 0,
          rowEnd: 20,
          colStart: 0,
          colEnd: 10,
        },
      },
      workflowInput: {
        sheetName: 'Imports',
      },
    })

    expect(result.title).toBe('Fill Current Sheet Formulas Down')
    expect(result.summary).toContain('Staged formula fill-down')
    expect(result.commands).toEqual([
      expect.objectContaining({
        kind: 'fillRange',
        source: {
          sheetName: 'Imports',
          startAddress: 'B2',
          endAddress: 'B2',
        },
        target: {
          sheetName: 'Imports',
          startAddress: 'B3',
          endAddress: 'B4',
        },
      }),
    ])
  })

  it('executes header-style workflows through the durable inspection path', async () => {
    const result = await executeWorkbookAgentWorkflow({
      documentId: 'doc-1',
      zeroSyncService: createZeroSyncStub({
        createRuntime: async () => {
          const engine = new SpreadsheetEngine({
            workbookName: 'doc-1',
            replicaId: 'server:test',
          })
          await engine.ready()
          engine.createSheet('Imports')
          engine.setCellValue('Imports', 'A1', 'Customer')
          engine.setCellValue('Imports', 'B1', 'Revenue')
          engine.setCellValue('Imports', 'A2', 'Ada')
          engine.setCellValue('Imports', 'B2', 100)
          return {
            documentId: 'doc-1',
            engine,
            projection: buildWorkbookSourceProjectionFromEngine('doc-1', engine, {
              revision: 1,
              calculatedRevision: 1,
              ownerUserId: 'alex@example.com',
              updatedBy: 'alex@example.com',
              updatedAt: '2026-04-10T00:00:00.000Z',
            }),
            headRevision: 1,
            calculatedRevision: 1,
            ownerUserId: 'alex@example.com',
          }
        },
      }),
      workflowTemplate: 'styleCurrentSheetHeaders',
      context: {
        selection: {
          sheetName: 'Imports',
          address: 'A1',
        },
        viewport: {
          rowStart: 0,
          rowEnd: 20,
          colStart: 0,
          colEnd: 10,
        },
      },
      workflowInput: {
        sheetName: 'Imports',
      },
    })

    expect(result.title).toBe('Style Current Sheet Headers')
    expect(result.summary).toContain('Prepared a consistent header style change set')
    expect(result.commands).toEqual([
      expect.objectContaining({
        kind: 'formatRange',
        range: {
          sheetName: 'Imports',
          startAddress: 'A1',
          endAddress: 'B1',
        },
        patch: expect.objectContaining({
          fill: expect.objectContaining({
            backgroundColor: '#E2E8F0',
          }),
          font: expect.objectContaining({
            bold: true,
          }),
        }),
      }),
    ])
  })

  it('executes current-sheet review-tab workflows through the durable inspection path', async () => {
    const result = await executeWorkbookAgentWorkflow({
      documentId: 'doc-1',
      zeroSyncService: createZeroSyncStub({
        createRuntime: async () => {
          const engine = new SpreadsheetEngine({
            workbookName: 'doc-1',
            replicaId: 'server:test',
          })
          await engine.ready()
          engine.createSheet('Revenue')
          engine.setCellValue('Revenue', 'A1', 'Region')
          engine.setCellValue('Revenue', 'B1', 'January')
          engine.setCellValue('Revenue', 'A2', 'West')
          engine.setCellValue('Revenue', 'B2', 100)
          engine.setCellValue('Revenue', 'A3', 'East')
          engine.setCellValue('Revenue', 'B3', 200)
          return {
            documentId: 'doc-1',
            engine,
            projection: buildWorkbookSourceProjectionFromEngine('doc-1', engine, {
              revision: 1,
              calculatedRevision: 1,
              ownerUserId: 'alex@example.com',
              updatedBy: 'alex@example.com',
              updatedAt: '2026-04-10T00:00:00.000Z',
            }),
            headRevision: 1,
            calculatedRevision: 1,
            ownerUserId: 'alex@example.com',
          }
        },
      }),
      workflowTemplate: 'createCurrentSheetReviewTab',
      context: {
        selection: {
          sheetName: 'Revenue',
          address: 'A1',
        },
        viewport: {
          rowStart: 0,
          rowEnd: 20,
          colStart: 0,
          colEnd: 10,
        },
      },
      workflowInput: {
        sheetName: 'Revenue',
      },
    })

    expect(result.title).toBe('Create Current Sheet Review Tab')
    expect(result.summary).toContain('Staged a review-tab change set')
    expect(result.commands).toEqual([
      expect.objectContaining({
        kind: 'createSheet',
        name: 'Revenue Review',
      }),
      expect.objectContaining({
        kind: 'copyRange',
        source: {
          sheetName: 'Revenue',
          startAddress: 'A1',
          endAddress: 'B3',
        },
        target: {
          sheetName: 'Revenue Review',
          startAddress: 'A1',
          endAddress: 'B3',
        },
      }),
    ])
  })

  it('executes current-sheet outlier-highlighting workflows through the durable inspection path', async () => {
    const result = await executeWorkbookAgentWorkflow({
      documentId: 'doc-1',
      zeroSyncService: createZeroSyncStub({
        createRuntime: async () => {
          const engine = new SpreadsheetEngine({
            workbookName: 'doc-1',
            replicaId: 'server:test',
          })
          await engine.ready()
          engine.createSheet('Revenue')
          engine.setCellValue('Revenue', 'A1', 'Region')
          engine.setCellValue('Revenue', 'B1', 'Revenue')
          engine.setCellValue('Revenue', 'A2', 'West')
          engine.setCellValue('Revenue', 'B2', 100)
          engine.setCellValue('Revenue', 'A3', 'East')
          engine.setCellValue('Revenue', 'B3', 105)
          engine.setCellValue('Revenue', 'A4', 'North')
          engine.setCellValue('Revenue', 'B4', 98)
          engine.setCellValue('Revenue', 'A5', 'South')
          engine.setCellValue('Revenue', 'B5', 102)
          engine.setCellValue('Revenue', 'A6', 'Enterprise')
          engine.setCellValue('Revenue', 'B6', 450)
          return {
            documentId: 'doc-1',
            engine,
            projection: buildWorkbookSourceProjectionFromEngine('doc-1', engine, {
              revision: 1,
              calculatedRevision: 1,
              ownerUserId: 'alex@example.com',
              updatedBy: 'alex@example.com',
              updatedAt: '2026-04-10T00:00:00.000Z',
            }),
            headRevision: 1,
            calculatedRevision: 1,
            ownerUserId: 'alex@example.com',
          }
        },
      }),
      workflowTemplate: 'highlightCurrentSheetOutliers',
      context: {
        selection: {
          sheetName: 'Revenue',
          address: 'A1',
        },
        viewport: {
          rowStart: 0,
          rowEnd: 20,
          colStart: 0,
          colEnd: 10,
        },
      },
      workflowInput: {
        sheetName: 'Revenue',
      },
    })

    expect(result.title).toBe('Highlight Current Sheet Outliers')
    expect(result.summary).toContain('Revenue')
    expect(result.commands).toEqual([
      expect.objectContaining({
        kind: 'formatRange',
        range: {
          sheetName: 'Revenue',
          startAddress: 'B6',
          endAddress: 'B6',
        },
        patch: expect.objectContaining({
          fill: expect.objectContaining({
            backgroundColor: '#FEF3C7',
          }),
        }),
      }),
    ])
    expect(result.artifact).toEqual(
      expect.objectContaining({
        title: 'Current Sheet Outlier Highlights',
        text: expect.stringContaining('## Highlighted Numeric Outliers'),
      }),
    )
  })

  it('executes formula-repair workflows through the durable inspection path', async () => {
    const result = await executeWorkbookAgentWorkflow({
      documentId: 'doc-1',
      zeroSyncService: createZeroSyncStub({
        createRuntime: async () => {
          const engine = new SpreadsheetEngine({
            workbookName: 'doc-1',
            replicaId: 'server:test',
          })
          await engine.ready()
          engine.createSheet('Sheet1')
          engine.setCellValue('Sheet1', 'A1', 10)
          engine.setCellValue('Sheet1', 'A2', 12)
          engine.setCellValue('Sheet1', 'A3', 14)
          engine.setCellFormula('Sheet1', 'B1', 'A1*2')
          engine.setCellFormula('Sheet1', 'B2', '1/0')
          engine.setCellFormula('Sheet1', 'B3', '1/0')
          return {
            documentId: 'doc-1',
            engine,
            projection: buildWorkbookSourceProjectionFromEngine('doc-1', engine, {
              revision: 1,
              calculatedRevision: 1,
              ownerUserId: 'alex@example.com',
              updatedBy: 'alex@example.com',
              updatedAt: '2026-04-10T00:00:00.000Z',
            }),
            headRevision: 1,
            calculatedRevision: 1,
            ownerUserId: 'alex@example.com',
          }
        },
      }),
      workflowTemplate: 'repairFormulaIssues',
      context: {
        selection: {
          sheetName: 'Sheet1',
          address: 'A1',
        },
        viewport: {
          rowStart: 0,
          rowEnd: 20,
          colStart: 0,
          colEnd: 10,
        },
      },
      workflowInput: {
        sheetName: 'Sheet1',
      },
    })

    expect(result.title).toBe('Repair Formula Issues')
    expect(result.summary).toContain('Staged 2 formula repairs')
    expect(result.commands).toEqual([
      expect.objectContaining({
        kind: 'writeRange',
        sheetName: 'Sheet1',
        startAddress: 'B2',
        values: [[{ formula: 'A2*2' }]],
      }),
      expect.objectContaining({
        kind: 'writeRange',
        sheetName: 'Sheet1',
        startAddress: 'B3',
        values: [[{ formula: 'A3*2' }]],
      }),
    ])
    expect(result.artifact).toEqual(
      expect.objectContaining({
        title: 'Formula Repair Preview',
        text: expect.stringContaining('## Formula Repair Preview'),
      }),
    )
  })
})
