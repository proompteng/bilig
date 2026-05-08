import { ValueTag, type CellValue } from '@bilig/protocol'
import { describe, expect, it } from 'vitest'

import { WorkPaper } from '../index.js'

type TestCell = string | number | null

const TARGET_ACCOUNT = '401000 Sales:SumIt Solution'

function columnLabel(index: number): string {
  let value = ''
  let current = index + 1
  while (current > 0) {
    const remainder = (current - 1) % 26
    value = String.fromCharCode(65 + remainder) + value
    current = Math.floor((current - 1) / 26)
  }
  return value
}

function cellValue(workbook: WorkPaper, ref: string): CellValue {
  const address = workbook.simpleCellAddressFromString(ref)
  if (!address) {
    throw new Error(`Expected ${ref} to resolve`)
  }
  return workbook.getCellValue(address)
}

function buildDeferredRevenueSheet(): { readonly rows: TestCell[][]; readonly expectedMonthlySums: readonly number[] } {
  const rowCount = 293
  const colCount = 57
  const firstDataRow = 11
  const lastDataRow = 293
  const firstMonthCol = 13
  const lastMonthCol = 56
  const summaryRows = 9
  const expectedMonthlySums = Array.from({ length: lastMonthCol - firstMonthCol + 1 }, () => 0)
  const rows = Array.from({ length: rowCount }, () => Array<TestCell>(colCount).fill(null))

  for (let col = firstMonthCol; col <= lastMonthCol; col += 1) {
    const monthIndex = col - firstMonthCol
    const month = (monthIndex % 12) + 1
    const colName = columnLabel(col)
    rows[9][col] = `=DATE(2024,${month},15)`
    for (let summaryRow = 0; summaryRow < summaryRows; summaryRow += 1) {
      rows[summaryRow][col] =
        `=SUMIFS($D$11:$D$293,$G$11:$G$293,"${TARGET_ACCOUNT}",$E$11:$E$293,">="&DATE(YEAR(${colName}$10),MONTH(${colName}$10),1),$E$11:$E$293,"<"&EDATE(DATE(YEAR(${colName}$10),MONTH(${colName}$10),1),1))`
    }
  }

  for (let rowNumber = firstDataRow; rowNumber <= lastDataRow; rowNumber += 1) {
    const row = rows[rowNumber - 1]
    const amount = rowNumber * 10
    const month = ((rowNumber - firstDataRow) % 12) + 1
    const isTargetAccount = rowNumber % 3 !== 0
    row[3] = amount
    row[4] = `=DATE(2024,${month},${(rowNumber % 28) + 1})`
    row[6] = isTargetAccount ? TARGET_ACCOUNT : 'Other'
    row[7] = '=DATE(2024,1,1)'
    row[8] = '=DATE(2025,12,31)'
    row[9] = amount / 24
    row[10] = 730
    if (isTargetAccount) {
      for (let monthIndex = month - 1; monthIndex < expectedMonthlySums.length; monthIndex += 12) {
        expectedMonthlySums[monthIndex] = (expectedMonthlySums[monthIndex] ?? 0) + amount
      }
    }
    for (let col = firstMonthCol; col <= lastMonthCol; col += 1) {
      const colName = columnLabel(col)
      row[col] = `=IFERROR(IF(AND(${colName}$10>=$H${rowNumber},${colName}$10<=$I${rowNumber}),$J${rowNumber},0),0)`
    }
  }

  return { rows, expectedMonthlySums }
}

describe('deferred-revenue timeout coverage', () => {
  it('builds deferred-revenue style workpapers with month-boundary SUMIFS criteria', () => {
    const { rows, expectedMonthlySums } = buildDeferredRevenueSheet()
    const workbook = WorkPaper.buildFromSheets(
      { 'Deferred Revenue': rows },
      {
        evaluationTimeoutMs: 10_000,
        maxColumns: 64,
        maxRows: 320,
        useColumnIndex: true,
      },
    )

    expect(cellValue(workbook, 'Deferred Revenue!N1')).toEqual({
      tag: ValueTag.Number,
      value: expectedMonthlySums[0],
    })
    expect(cellValue(workbook, 'Deferred Revenue!P1')).toEqual({
      tag: ValueTag.Number,
      value: expectedMonthlySums[2],
    })
    expect(cellValue(workbook, 'Deferred Revenue!N11')).toEqual({
      tag: ValueTag.Number,
      value: 110 / 24,
    })

    workbook.dispose()
  })
})
