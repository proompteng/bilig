import { describe, expect, it } from 'vitest'
import { ValueTag, type CellValue, type WorkbookSnapshot } from '@bilig/protocol'
import { WorkPaper } from '../index.js'

function expectNumber(value: CellValue, expected: number): void {
  expect(value.tag).toBe(ValueTag.Number)
  if (value.tag !== ValueTag.Number) {
    throw new Error(`Expected number ${String(expected)}, received ${JSON.stringify(value)}`)
  }
  expect(value.value).toBeCloseTo(expected, 12)
}

describe('workbook defined-name semantics', () => {
  it('resolves names like change1 through INDIRECT instead of invalid far-right cells', () => {
    const snapshot: WorkbookSnapshot = {
      version: 1,
      workbook: {
        name: 'issue-120-indirect-change-name',
        metadata: {
          definedNames: [
            { name: 'change1', value: { kind: 'range-ref', sheetName: 'YieldChanges', startAddress: 'A1', endAddress: 'A4' } },
            { name: 'change2', value: { kind: 'range-ref', sheetName: 'YieldChanges', startAddress: 'B1', endAddress: 'B4' } },
          ],
        },
      },
      sheets: [
        {
          id: 1,
          name: 'YieldChanges',
          order: 0,
          cells: [
            { address: 'A1', value: 1 },
            { address: 'A2', value: 2 },
            { address: 'A3', value: 3 },
            { address: 'A4', value: 4 },
            { address: 'B1', value: 2 },
            { address: 'B2', value: 4 },
            { address: 'B3', value: 6 },
            { address: 'B4', value: 8 },
          ],
        },
        {
          id: 2,
          name: 'Main',
          order: 1,
          cells: [
            { address: 'I5', value: 'change1' },
            { address: 'J4', value: 'change2' },
            { address: 'J5', formula: 'COVARIANCE.P(INDIRECT($I5),INDIRECT(J$4))' },
            { address: 'T5', formula: 'CORREL(INDIRECT($I5),INDIRECT(J$4))' },
          ],
        },
      ],
    }

    const workbook = WorkPaper.buildFromSnapshot(snapshot, { maxRows: 32, maxColumns: 32, useColumnIndex: true })
    const mainSheetId = workbook.getSheetId('Main')!

    try {
      expectNumber(workbook.getCellValue({ sheet: mainSheetId, row: 4, col: 9 }), 2.5)
      expectNumber(workbook.getCellValue({ sheet: mainSheetId, row: 4, col: 19 }), 1)
    } finally {
      workbook.dispose()
    }
  })

  it('intersects row-vector defined names in scalar logical formulas', () => {
    const snapshot: WorkbookSnapshot = {
      version: 1,
      workbook: {
        name: 'issue-120-root-row-vector-intersection',
        metadata: {
          definedNames: [
            { name: 'root1', value: { kind: 'range-ref', sheetName: 'Main', startAddress: 'B51', endAddress: 'D51' } },
            { name: 'root2', value: { kind: 'range-ref', sheetName: 'Main', startAddress: 'B52', endAddress: 'D52' } },
          ],
        },
      },
      sheets: [
        {
          id: 1,
          name: 'Main',
          order: 0,
          cells: [
            { address: 'B51', value: 1 },
            { address: 'C51', value: 2 },
            { address: 'D51', value: 0.5 },
            { address: 'B52', value: 10 },
            { address: 'C52', value: 20 },
            { address: 'D52', value: 30 },
            { address: 'B54', formula: 'IF(AND(root1<=1,root1>=0),root1,root2)' },
            { address: 'C54', formula: 'IF(AND(root1<=1,root1>=0),root1,root2)' },
            { address: 'D54', formula: 'IF(AND(root1<=1,root1>=0),root1,root2)' },
          ],
        },
      ],
    }

    const workbook = WorkPaper.buildFromSnapshot(snapshot, { maxRows: 64, maxColumns: 8, useColumnIndex: true })
    const mainSheetId = workbook.getSheetId('Main')!

    try {
      expectNumber(workbook.getCellValue({ sheet: mainSheetId, row: 53, col: 1 }), 1)
      expectNumber(workbook.getCellValue({ sheet: mainSheetId, row: 53, col: 2 }), 20)
      expectNumber(workbook.getCellValue({ sheet: mainSheetId, row: 53, col: 3 }), 0.5)
    } finally {
      workbook.dispose()
    }
  })
})
