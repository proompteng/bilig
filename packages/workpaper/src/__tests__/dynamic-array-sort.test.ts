import { ValueTag, type CellValue } from '@bilig/protocol'
import { describe, expect, it } from 'vitest'

import { WorkPaper } from '../index.js'

type TestCell = string | number | null

function cellValue(workbook: WorkPaper, ref: string): CellValue {
  const address = workbook.simpleCellAddressFromString(ref)
  if (!address) {
    throw new Error(`Expected ${ref} to resolve`)
  }
  return workbook.getCellValue(address)
}

function expectString(value: CellValue, expected: string): void {
  expect(value).toMatchObject({ tag: ValueTag.String, value: expected })
}

function expectNumber(value: CellValue, expected: number): void {
  expect(value).toEqual({ tag: ValueTag.Number, value: expected })
}

function buildLedgerWorkbook(useColumnIndex: boolean): WorkPaper {
  return WorkPaper.buildFromSheets(
    {
      Ledger: [
        ['Account', 'Amount', null, 'Unique accounts', 'Positive account', 'Positive amount', 'Sorted accounts'],
        ['4000', 1000, null, '=UNIQUE(A2:A5)', '=FILTER(A2:B5,B2:B5>0)', null, '=SORT(A2:A5)'],
        ['5000', -50, null, null, null, null, null],
        ['6100', 0, null, null, null, null, null],
        ['4000', 200, null, null, null, null, null],
      ] satisfies TestCell[][],
    },
    { maxRows: 12, maxColumns: 10, useColumnIndex },
  )
}

describe('dynamic array SORT', () => {
  it.each([false, true])('spills sorted text ledger account codes with useColumnIndex=%s', (useColumnIndex) => {
    const workbook = buildLedgerWorkbook(useColumnIndex)

    expectString(cellValue(workbook, 'Ledger!D2'), '4000')
    expectString(cellValue(workbook, 'Ledger!D3'), '5000')
    expectString(cellValue(workbook, 'Ledger!D4'), '6100')

    expectString(cellValue(workbook, 'Ledger!E2'), '4000')
    expectNumber(cellValue(workbook, 'Ledger!F2'), 1000)
    expectString(cellValue(workbook, 'Ledger!E3'), '4000')
    expectNumber(cellValue(workbook, 'Ledger!F3'), 200)

    expectString(cellValue(workbook, 'Ledger!G2'), '4000')
    expectString(cellValue(workbook, 'Ledger!G3'), '4000')
    expectString(cellValue(workbook, 'Ledger!G4'), '5000')
    expectString(cellValue(workbook, 'Ledger!G5'), '6100')

    workbook.dispose()
  })

  it.each([false, true])('counts blank values returned by UNIQUE with useColumnIndex=%s', (useColumnIndex) => {
    const workbook = WorkPaper.buildFromSheets(
      {
        Data: [
          ['Type', 'Direct count', 'Unique count'],
          [null, '=COUNTA(A2:A5)', '=COUNTA(UNIQUE(A2:A5))'],
          [null, null, null],
          ['Water', null, null],
          ['Water', null, null],
        ] satisfies TestCell[][],
      },
      { maxRows: 8, maxColumns: 4, useColumnIndex },
    )

    expectNumber(cellValue(workbook, 'Data!B2'), 2)
    expectNumber(cellValue(workbook, 'Data!C2'), 2)

    workbook.dispose()
  })
})
