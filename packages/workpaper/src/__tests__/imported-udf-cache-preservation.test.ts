import { ErrorCode, ValueTag, type CellValue, type WorkbookSnapshot } from '@bilig/protocol'
import { describe, expect, it } from 'vitest'

import { WorkPaper } from '../index.js'

function cellValue(workbook: WorkPaper, sheetName: string, row: number, col: number): CellValue {
  return workbook.getCellValue({ sheet: workbook.getSheetId(sheetName)!, row, col })
}

function issueSnapshot(options: { readonly manualNativeOpen?: boolean } = {}): WorkbookSnapshot {
  return {
    version: 1,
    workbook: {
      name: 'UDF cached workbook',
      ...(options.manualNativeOpen
        ? {
            metadata: {
              calculationSettings: {
                mode: 'manual',
                compatibilityMode: 'excel-modern',
                fullCalcOnLoad: false,
              },
            },
          }
        : {}),
    },
    sheets: [
      {
        id: 1,
        name: 'Model',
        order: 0,
        cells: [
          {
            address: 'A1',
            formula: '_xldudf_WISEPRICE(B1,"Shares Outstanding")',
            value: 14935800000,
          },
          { address: 'B1', value: 'AAPL' },
          { address: 'C1', formula: 'A1/1000000' },
          { address: 'D1', formula: '_FV(B1,"Ticker symbol",TRUE)', value: 'AAPL' },
          { address: 'E1', formula: 'D1&" ok"' },
        ],
      },
    ],
  }
}

describe('imported UDF cached formula values', () => {
  it.each([false, true])(
    'hydrates cached unsupported UDF formula values for manual native-open imports with useColumnIndex=%s',
    (useColumnIndex) => {
      const workbook = WorkPaper.buildFromSnapshot(issueSnapshot({ manualNativeOpen: true }), { maxRows: 8, maxColumns: 8, useColumnIndex })

      expect(cellValue(workbook, 'Model', 0, 0)).toEqual({
        tag: ValueTag.Number,
        value: 14935800000,
      })
      expect(cellValue(workbook, 'Model', 0, 2)).toEqual({
        tag: ValueTag.Number,
        value: 14935.8,
      })
      expect(cellValue(workbook, 'Model', 0, 3)).toEqual({
        tag: ValueTag.String,
        value: 'AAPL',
        stringId: expect.any(Number),
      })
      expect(cellValue(workbook, 'Model', 0, 4)).toEqual({
        tag: ValueTag.String,
        value: 'AAPL ok',
        stringId: expect.any(Number),
      })

      workbook.dispose()
    },
  )

  it('does not preserve imported UDF caches during explicit full recalculation', () => {
    const workbook = WorkPaper.buildFromSnapshot(issueSnapshot(), { maxRows: 8, maxColumns: 8, useColumnIndex: true })

    workbook.rebuildAndRecalculate()

    expect(cellValue(workbook, 'Model', 0, 0)).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Name,
    })
    expect(cellValue(workbook, 'Model', 0, 2)).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Name,
    })
    expect(cellValue(workbook, 'Model', 0, 3)).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Field,
    })
    expect(cellValue(workbook, 'Model', 0, 4)).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Field,
    })

    workbook.dispose()
  })

  it('does not hydrate stale cached values for formulas the engine can evaluate', () => {
    const snapshot = issueSnapshot()
    snapshot.sheets[0].cells.push({ address: 'A2', formula: 'SUM(1,2)', value: 999 })
    const workbook = WorkPaper.buildFromSnapshot(snapshot, { maxRows: 8, maxColumns: 8, useColumnIndex: true })

    expect(cellValue(workbook, 'Model', 1, 0)).toEqual({
      tag: ValueTag.Number,
      value: 3,
    })

    workbook.dispose()
  })

  it('does not hydrate or preserve arbitrary unknown formula caches during full recalculation', () => {
    const snapshot = issueSnapshot()
    snapshot.sheets[0].cells = [{ address: 'A1', formula: 'UNKNOWNFUNC(42)', value: 99 }]
    const workbook = WorkPaper.buildFromSnapshot(snapshot, { maxRows: 8, maxColumns: 8, useColumnIndex: true })

    expect(cellValue(workbook, 'Model', 0, 0)).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Name,
    })

    workbook.rebuildAndRecalculate()

    expect(cellValue(workbook, 'Model', 0, 0)).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Name,
    })

    workbook.dispose()
  })
})
