import { describe, expect, it } from 'vitest'
import { ValueTag, type CellValue } from '@bilig/protocol'

import { WorkPaper } from '../index.js'

type TestCell = string | number | null

function cellValue(workbook: WorkPaper, sheetName: string, row: number, col: number): CellValue {
  return workbook.getCellValue({ sheet: workbook.getSheetId(sheetName), row, col })
}

function expectNumber(value: CellValue, expected: number): void {
  expect(value).toEqual({ tag: ValueTag.Number, value: expected })
}

function expectString(value: CellValue, expected: string): void {
  expect(value).toMatchObject({ tag: ValueTag.String, value: expected })
}

function buildReconciliationSheets(rowCount: number): Record<string, TestCell[][]> {
  const donations = Array.from({ length: rowCount }, () => Array<TestCell>(14).fill(null))
  const deposits = Array.from({ length: rowCount }, () => Array<TestCell>(12).fill(null))

  donations[0] = [
    'date',
    'donor',
    'desc',
    'method',
    'ref',
    'key',
    'gross',
    'fee',
    'net',
    'net_num',
    'donations_by_key',
    'deposits_by_key',
    'difference',
    'status',
  ]
  deposits[0] = [
    'date',
    'desc',
    'account',
    'key',
    'ref',
    'memo',
    'amount',
    'amount_num',
    'donations_by_key',
    'deposits_by_key',
    'difference',
    'status',
  ]

  for (let row = 2; row <= rowCount; row += 1) {
    const index = row - 1
    const key = `batch-${index % 250}`

    donations[index][0] = 45_000 + (index % 30)
    donations[index][5] = key
    donations[index][6] = (index % 100) + 0.25
    donations[index][7] = index % 3 === 0 ? 0.25 : 0
    donations[index][8] = `=IF(AND(G${row}<>"",H${row}<>""),G${row}-H${row},IF(G${row}<>"",G${row},""))`
    donations[index][9] = `=IF(I${row}="","",IFERROR(1*I${row},0))`
    donations[index][10] = `=IF(F${row}="","",SUMIFS($J$2:$J$${rowCount},$F$2:$F$${rowCount},F${row}))`
    donations[index][11] = `=IF(F${row}="","",SUMIFS('Bank Deposits'!$H$2:$H$${rowCount},'Bank Deposits'!$D$2:$D$${rowCount},F${row}))`
    donations[index][12] = `=IF(K${row}="","",K${row}-L${row})`
    donations[index][13] = `=IF(COUNTA(A${row}:H${row})=0,"",IF(F${row}="","Missing Key",IF(ABS(M${row})<0.01,"Matched","Investigate")))`

    deposits[index][0] = 45_000 + (index % 30)
    deposits[index][3] = key
    deposits[index][6] = (index % 100) + (index % 3 === 0 ? 0 : 0.25)
    deposits[index][7] = `=IF(G${row}="","",IFERROR(1*G${row},0))`
    deposits[index][8] = `=IF(D${row}="","",SUMIFS(Donations!$J$2:$J$${rowCount},Donations!$F$2:$F$${rowCount},D${row}))`
    deposits[index][9] = `=IF(D${row}="","",SUMIFS($H$2:$H$${rowCount},$D$2:$D$${rowCount},D${row}))`
    deposits[index][10] = `=IF(I${row}="","",I${row}-J${row})`
    deposits[index][11] = `=IF(COUNTA(A${row}:G${row})=0,"",IF(D${row}="","Missing Key",IF(ABS(K${row})<0.01,"Matched","Investigate")))`
  }

  return {
    Donations: donations,
    'Bank Deposits': deposits,
    Summary: [
      ['check', 'value'],
      ['diff', `=SUM(Donations!$M$2:$M$${rowCount})`],
      ['exceptions', `=COUNTIF(Donations!$N$2:$N$${rowCount},"Investigate")`],
    ],
  }
}

describe('guarded SUMIFS workpaper performance', () => {
  const reconciliationEvaluationTimeoutMs = 15_000

  it('evaluates blank-key guards without leaving the direct criteria path', () => {
    const workbook = WorkPaper.buildFromSheets(
      {
        Data: [
          ['key', 'amount'],
          ['batch-a', 10],
          ['batch-a', 5],
          ['', 99],
        ],
        Summary: [
          ['key', 'total'],
          ['batch-a', '=IF(A2="","",SUMIFS(Data!$B$2:$B$4,Data!$A$2:$A$4,A2))'],
          ['', '=IF(A3="","",SUMIFS(Data!$B$2:$B$4,Data!$A$2:$A$4,A3))'],
          ['batch-a', '=IF(A4<>"",SUMIFS(Data!$B$2:$B$4,Data!$A$2:$A$4,A4),"")'],
        ],
      },
      { maxRows: 12, maxColumns: 6, useColumnIndex: true },
    )

    expectNumber(cellValue(workbook, 'Summary', 1, 1), 15)
    expectString(cellValue(workbook, 'Summary', 2, 1), '')
    expectNumber(cellValue(workbook, 'Summary', 3, 1), 15)
    expect(workbook.getPerformanceCounters().directFormulaInitialEvaluations).toBe(3)
  })

  it(
    'builds repeated reconciliation SUMIFS formulas within a bounded budget',
    () => {
      const rowCount = 1_500
      const workbook = WorkPaper.buildFromSheets(buildReconciliationSheets(rowCount), {
        evaluationTimeoutMs: reconciliationEvaluationTimeoutMs,
        useWildcards: true,
        useRegularExpressions: false,
      })

      expectNumber(cellValue(workbook, 'Summary', 1, 1), 0)
      expectNumber(cellValue(workbook, 'Summary', 2, 1), 0)
      expect(workbook.getConfig().useColumnIndex).toBe(true)
    },
    reconciliationEvaluationTimeoutMs + 5_000,
  )
})
