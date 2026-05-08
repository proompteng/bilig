import { ValueTag, type CellValue, type WorkbookSnapshot } from '@bilig/protocol'
import { describe, expect, it } from 'vitest'

import { WorkPaper } from '../index.js'

function cellValue(workbook: WorkPaper, ref: string): CellValue {
  const address = workbook.simpleCellAddressFromString(ref)
  if (!address) {
    throw new Error(`Expected ${ref} to resolve`)
  }
  return workbook.getCellValue(address)
}

function expectNumberClose(value: CellValue, expected: number): void {
  expect(value.tag).toBe(ValueTag.Number)
  if (value.tag !== ValueTag.Number) {
    throw new Error(`Expected number ${expected}, got ${JSON.stringify(value)}`)
  }
  expect(value.value).toBeCloseTo(expected, 12)
}

describe('GitHub issue #120 legacy array-formula import context', () => {
  it.each([false, true])('keeps range-valued names in array context for imported spill owners with useColumnIndex=%s', (useColumnIndex) => {
    const snapshot: WorkbookSnapshot = {
      version: 1,
      workbook: {
        name: 'legacy-array-name-context',
        metadata: {
          definedNames: [
            { name: 'corrmat', value: { kind: 'range-ref', sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B2' } },
            { name: 'volmat', value: { kind: 'range-ref', sheetName: 'Sheet1', startAddress: 'D1', endAddress: 'E2' } },
            { name: 'covmat', value: { kind: 'range-ref', sheetName: 'Sheet1', startAddress: 'A4', endAddress: 'B5' } },
            { name: 'posvec', value: { kind: 'range-ref', sheetName: 'Sheet1', startAddress: 'G1', endAddress: 'G2' } },
            { name: 'w', value: { kind: 'range-ref', sheetName: 'Sheet1', startAddress: 'D4', endAddress: 'D5' } },
          ],
          spills: [
            { sheetName: 'Sheet1', address: 'A4', rows: 2, cols: 2 },
            { sheetName: 'Sheet1', address: 'D4', rows: 2, cols: 1 },
          ],
        },
      },
      sheets: [
        {
          id: 1,
          name: 'Sheet1',
          order: 0,
          cells: [
            { address: 'A1', value: 1 },
            { address: 'B1', value: 2 },
            { address: 'A2', value: 3 },
            { address: 'B2', value: 4 },
            { address: 'D1', value: 10 },
            { address: 'E1', value: 20 },
            { address: 'D2', value: 30 },
            { address: 'E2', value: 40 },
            { address: 'G1', value: 2 },
            { address: 'G2', value: 3 },
            { address: 'A4', formula: 'corrmat*volmat' },
            { address: 'B4', value: 0 },
            { address: 'A5', value: 0 },
            { address: 'B5', value: 0 },
            { address: 'D4', formula: 'posvec/SUM(posvec)' },
            { address: 'D5', value: 0 },
            { address: 'F4', formula: 'MMULT(MMULT(TRANSPOSE(w),covmat),w)' },
          ],
        },
      ],
    }

    const workbook = WorkPaper.buildFromSnapshot(snapshot, { maxRows: 10, maxColumns: 10, useColumnIndex })

    expect(cellValue(workbook, 'Sheet1!A4')).toEqual({ tag: ValueTag.Number, value: 10 })
    expect(cellValue(workbook, 'Sheet1!B4')).toEqual({ tag: ValueTag.Number, value: 40 })
    expect(cellValue(workbook, 'Sheet1!A5')).toEqual({ tag: ValueTag.Number, value: 90 })
    expect(cellValue(workbook, 'Sheet1!B5')).toEqual({ tag: ValueTag.Number, value: 160 })
    expectNumberClose(cellValue(workbook, 'Sheet1!D4'), 0.4)
    expectNumberClose(cellValue(workbook, 'Sheet1!D5'), 0.6)
    expectNumberClose(cellValue(workbook, 'Sheet1!F4'), 90.4)

    workbook.dispose()
  })
})
