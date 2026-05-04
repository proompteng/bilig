import { SpreadsheetEngine } from '@bilig/core'
import { describe, expect, it } from 'vitest'
import type { WorkbookAgentUiContext } from '@bilig/contracts'
import { buildWorkbookSourceProjectionFromEngine } from '../zero/projection.js'
import type { WorkbookRuntime } from '../workbook-runtime/runtime-manager.js'
import {
  inspectWorkbookCell,
  inspectWorkbookContext,
  inspectWorkbookRange,
  normalizeWorkbookAgentUiContext,
} from './workbook-agent-inspection.js'

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

async function createRuntime(): Promise<WorkbookRuntime> {
  const engine = new SpreadsheetEngine({
    workbookName: 'doc-1',
    replicaId: 'server:inspection-test',
  })
  await engine.ready()
  engine.createSheet('Sheet1')
  engine.createSheet('Sheet2')
  engine.setCellValue('Sheet1', 'B2', 'Revenue')
  engine.setCellValue('Sheet1', 'B3', 'Draft')
  engine.setCellValue('Sheet1', 'C3', 12)
  engine.setCellFormula('Sheet1', 'D4', 'C3*2')
  engine.setDataValidation({
    range: {
      sheetName: 'Sheet1',
      startAddress: 'B3',
      endAddress: 'B4',
    },
    rule: {
      kind: 'list',
      values: ['Draft', 'Final'],
    },
  })
  engine.setConditionalFormat({
    id: 'cf-1',
    range: {
      sheetName: 'Sheet1',
      startAddress: 'C3',
      endAddress: 'D4',
    },
    rule: {
      kind: 'cellIs',
      operator: 'greaterThan',
      values: [10],
    },
    style: {
      fill: { backgroundColor: '#ff0000' },
    },
  })
  engine.setCommentThread({
    threadId: 'thread-1',
    sheetName: 'Sheet1',
    address: 'B3',
    comments: [{ id: 'comment-1', body: 'Review this cell.' }],
  })
  engine.setNote({
    sheetName: 'Sheet1',
    address: 'C3',
    text: 'Manual override',
  })
  engine.setFreezePane('Sheet1', 1, 0)
  engine.updateRowMetadata('Sheet1', 2, 1, 28, true)
  engine.updateColumnMetadata('Sheet1', 2, 1, 120, true)
  engine.setRangeProtection({
    id: 'protect-d4',
    range: {
      sheetName: 'Sheet1',
      startAddress: 'D4',
      endAddress: 'D4',
    },
    hideFormulas: true,
  })

  return {
    documentId: 'doc-1',
    engine,
    projection: buildWorkbookSourceProjectionFromEngine('doc-1', engine, {
      revision: 7,
      calculatedRevision: 7,
      ownerUserId: 'alex@example.com',
      updatedBy: 'alex@example.com',
      updatedAt: '2026-04-12T12:00:00.000Z',
    }),
    headRevision: 7,
    calculatedRevision: 7,
    ownerUserId: 'alex@example.com',
  }
}

describe('workbook agent inspection helpers', () => {
  it('normalizes ui context to existing sheets and safe addresses', async () => {
    const runtime = await createRuntime()

    expect(
      normalizeWorkbookAgentUiContext(runtime, {
        selection: {
          sheetName: 'Missing',
          address: 'not-an-address',
          range: {
            startAddress: 'bad',
            endAddress: 'still-bad',
          },
        },
        viewport: {
          rowStart: -4.8,
          rowEnd: 8.2,
          colStart: -3.1,
          colEnd: 5.9,
        },
        rendered: {
          capturedAtUnixMs: 10,
          capturedRevision: 7,
          batchId: 1,
          selection: null,
          visibleRange: null,
        },
      }),
    ).toEqual({
      selection: {
        sheetName: 'Sheet1',
        address: 'A1',
        range: {
          startAddress: 'A1',
          endAddress: 'A1',
        },
      },
      viewport: {
        rowStart: 0,
        rowEnd: 8,
        colStart: 0,
        colEnd: 5,
      },
    })
  })

  it('summarizes workbook context from normalized selection and viewport state', async () => {
    const runtime = await createRuntime()
    const context: WorkbookAgentUiContext = {
      selection: {
        sheetName: 'Sheet1',
        address: 'B3',
        range: {
          startAddress: 'B3',
          endAddress: 'D4',
        },
      },
      viewport: {
        rowStart: 1,
        rowEnd: 4,
        colStart: 1,
        colEnd: 4,
      },
    }

    const parsedSummary: unknown = JSON.parse(inspectWorkbookContext(runtime, context))
    expect(isRecord(parsedSummary)).toBe(true)
    if (!isRecord(parsedSummary)) {
      throw new Error('Expected workbook context summary object')
    }
    const summary = parsedSummary

    expect(summary['selection']).toEqual(
      expect.objectContaining({
        kind: 'range',
        sheetName: 'Sheet1',
        startAddress: 'B3',
        endAddress: 'D4',
        rowCount: 2,
        columnCount: 3,
      }),
    )
    expect(summary['visibleRange']).toEqual(
      expect.objectContaining({
        sheetName: 'Sheet1',
        startAddress: 'B2',
        endAddress: 'E5',
      }),
    )
    expect(summary['sheetState']).toEqual(
      expect.objectContaining({
        freezePane: { sheetName: 'Sheet1', rows: 1, cols: 0 },
      }),
    )
  })

  it('includes intersecting workbook metadata and hides protected formulas in range inspection', async () => {
    const runtime = await createRuntime()

    const inspection = inspectWorkbookRange(runtime, {
      sheetName: 'Sheet1',
      startAddress: 'B3',
      endAddress: 'D4',
    })

    expect(inspection.range).toEqual({
      sheetName: 'Sheet1',
      startAddress: 'B3',
      endAddress: 'D4',
    })
    expect(inspection.dataValidations).toEqual([
      expect.objectContaining({
        range: expect.objectContaining({ startAddress: 'B3', endAddress: 'B4' }),
      }),
    ])
    expect(inspection.conditionalFormats).toEqual([
      expect.objectContaining({
        id: 'cf-1',
      }),
    ])
    expect(inspection.commentThreads).toEqual([
      expect.objectContaining({
        threadId: 'thread-1',
      }),
    ])
    expect(inspection.notes).toEqual([
      expect.objectContaining({
        text: 'Manual override',
      }),
    ])
    expect(inspection.rangeProtections).toEqual([
      expect.objectContaining({
        id: 'protect-d4',
      }),
    ])
    expect(inspection.sheetState).toEqual(
      expect.objectContaining({
        hiddenRows: [{ rowNumber: 3 }],
        hiddenColumns: [{ columnIndex: 2, columnLabel: 'C' }],
      }),
    )
    expect(inspection.rows[1]).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          address: 'D4',
          formula: null,
        }),
      ]),
    )
  })

  it('includes per-cell metadata and hides protected formulas in cell inspection', async () => {
    const runtime = await createRuntime()

    const inspection = inspectWorkbookCell(runtime, {
      sheetName: 'Sheet1',
      address: 'D4',
    })

    expect(inspection.sheetName).toBe('Sheet1')
    expect(inspection.address).toBe('D4')
    expect(inspection.formula).toBeNull()
    expect(inspection.rangeProtections).toEqual([
      expect.objectContaining({
        id: 'protect-d4',
      }),
    ])
    expect(inspection.conditionalFormats).toEqual([
      expect.objectContaining({
        id: 'cf-1',
      }),
    ])
  })
})
