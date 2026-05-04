import { SpreadsheetEngine } from '@bilig/core'
import { describe, expect, it } from 'vitest'
import type { WorkbookAgentUiContext } from '@bilig/contracts'
import { buildWorkbookSourceProjectionFromEngine } from '../zero/projection.js'
import type { WorkbookRuntime } from '../workbook-runtime/runtime-manager.js'
import {
  resolveFormulaRangeRequest,
  resolveRangeOrSelectorRequest,
  resolveReadRangeRequest,
  resolveTransferRangeRequest,
  resolveWriteRangeRequest,
} from './workbook-agent-selector-tooling.js'

async function createRuntime(): Promise<WorkbookRuntime> {
  const engine = new SpreadsheetEngine({
    workbookName: 'doc-1',
    replicaId: 'server:selector-tooling-test',
  })
  await engine.ready()
  engine.createSheet('Sheet1')
  engine.setCellValue('Sheet1', 'A1', 'Revenue')
  engine.setCellValue('Sheet1', 'B1', 'Margin')
  engine.setCellValue('Sheet1', 'A2', 10)
  engine.setCellValue('Sheet1', 'B2', 2)
  engine.setCellValue('Sheet1', 'A3', 12)
  engine.setCellValue('Sheet1', 'B3', 3)
  engine.setDefinedName('Inputs', {
    kind: 'range-ref',
    sheetName: 'Sheet1',
    startAddress: 'B3',
    endAddress: 'A2',
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

function createUiContext(): WorkbookAgentUiContext {
  return {
    selection: {
      sheetName: 'Sheet1',
      address: 'B2',
      range: {
        startAddress: 'B2',
        endAddress: 'C3',
      },
    },
    viewport: {
      rowStart: 1,
      rowEnd: 3,
      colStart: 0,
      colEnd: 2,
    },
  }
}

describe('workbook agent selector tooling', () => {
  it('normalizes direct ranges and selector-derived ranges for reads and single-range requests', async () => {
    const runtime = await createRuntime()

    expect(
      resolveReadRangeRequest({
        runtime,
        args: {
          sheetName: 'Sheet1',
          startAddress: 'B3',
          endAddress: 'A2',
        },
        uiContext: null,
      }),
    ).toEqual({
      ranges: [
        {
          sheetName: 'Sheet1',
          startAddress: 'A2',
          endAddress: 'B3',
        },
      ],
      resolution: null,
    })

    expect(
      resolveRangeOrSelectorRequest({
        runtime,
        args: {
          range: {
            sheetName: 'Sheet1',
            startAddress: 'C4',
            endAddress: 'B2',
          },
        },
        uiContext: null,
      }),
    ).toEqual({
      range: {
        sheetName: 'Sheet1',
        startAddress: 'B2',
        endAddress: 'C4',
      },
      resolution: null,
    })

    expect(
      resolveReadRangeRequest({
        runtime,
        args: {
          selector: {
            kind: 'namedRange',
            name: 'Inputs',
          },
        },
        uiContext: null,
      }),
    ).toEqual({
      ranges: [
        {
          sheetName: 'Sheet1',
          startAddress: 'A2',
          endAddress: 'B3',
        },
      ],
      resolution: expect.objectContaining({
        displayLabel: 'Inputs',
      }),
    })
  })

  it('resolves transfer targets and expands single-cell formula targets', async () => {
    const runtime = await createRuntime()

    expect(
      resolveTransferRangeRequest({
        runtime,
        args: {
          source: {
            sheetName: 'Sheet1',
            startAddress: 'B3',
            endAddress: 'A2',
          },
          target: {
            sheetName: 'Sheet1',
            startAddress: 'D4',
            endAddress: 'C3',
          },
        },
        uiContext: null,
      }),
    ).toEqual({
      source: {
        sheetName: 'Sheet1',
        startAddress: 'A2',
        endAddress: 'B3',
      },
      target: {
        sheetName: 'Sheet1',
        startAddress: 'C3',
        endAddress: 'D4',
      },
      sourceResolution: null,
      targetResolution: null,
    })

    expect(
      resolveFormulaRangeRequest({
        runtime,
        args: {
          range: {
            sheetName: 'Sheet1',
            startAddress: 'D4',
            endAddress: 'D4',
          },
          formulas: [
            ['SUM(A1:A1)', 'SUM(B1:B1)'],
            ['SUM(A2:A2)', 'SUM(B2:B2)'],
          ],
        },
        uiContext: null,
      }),
    ).toEqual({
      range: {
        sheetName: 'Sheet1',
        startAddress: 'D4',
        endAddress: 'E5',
      },
      resolution: null,
    })
  })

  it('validates selector-driven write and formula dimensions against resolved ranges', async () => {
    const runtime = await createRuntime()

    expect(
      resolveWriteRangeRequest({
        runtime,
        args: {
          selector: {
            kind: 'namedRange',
            name: 'Inputs',
          },
          values: [
            ['North', '10'],
            ['South', '12'],
          ],
        },
        uiContext: createUiContext(),
      }),
    ).toEqual({
      sheetName: 'Sheet1',
      startAddress: 'A2',
      resolution: expect.objectContaining({
        displayLabel: 'Inputs',
      }),
    })

    expect(() =>
      resolveWriteRangeRequest({
        runtime,
        args: {
          selector: {
            kind: 'namedRange',
            name: 'Inputs',
          },
          values: [['Only one row']],
        },
        uiContext: createUiContext(),
      }),
    ).toThrow('Selector Inputs resolves to 2x2, but write_range received 1x1 values')

    expect(() =>
      resolveFormulaRangeRequest({
        runtime,
        args: {
          selector: {
            kind: 'namedRange',
            name: 'Inputs',
          },
          formulas: [['SUM(A1:A1)']],
        },
        uiContext: createUiContext(),
      }),
    ).toThrow('Selector Inputs resolves to 2x2, but set_formula received 1x1 formulas')
  })
})
