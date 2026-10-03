import { describe, expect, it } from 'vitest'
import { ValueTag, type CellValue } from '@bilig/protocol'
import { indexToColumn } from '@bilig/formula'

import { WorkPaper, type WorkPaperCellAddress } from '../index.js'

type TestCell = string | number | null

function cellValue(workbook: WorkPaper, sheetName: string, row: number, col: number): CellValue {
  return workbook.getCellValue({ sheet: workbook.getSheetId(sheetName), row, col })
}

function cell(sheet: number, row: number, col: number): WorkPaperCellAddress {
  return { sheet, row, col }
}

function cellChanges(changes: Array<{ kind: string; a1?: string }>): Set<string> {
  return new Set(changes.flatMap((change) => (change.kind === 'cell' && change.a1 ? [change.a1] : [])))
}

function expectNumberClose(value: CellValue, expected: number): void {
  expect(value.tag).toBe(ValueTag.Number)
  if (value.tag !== ValueTag.Number) {
    throw new Error(`Expected number ${String(expected)}, received ${JSON.stringify(value)}`)
  }
  expect(value.value).toBeCloseTo(expected, 12)
}

function buildIssue54Workbook(useColumnIndex: boolean): WorkPaper {
  const debtGrid = Array.from({ length: 5 }, () => Array.from<TestCell>({ length: 11 }).fill(null))
  for (let col = 0; col < debtGrid[0].length; col += 1) {
    debtGrid[0][col] = col / 10
  }
  debtGrid[2][2] = 45616.01199262912
  debtGrid[3][2] = 1.0233773685605514
  debtGrid[4][2] = 0.10664524406185244

  return WorkPaper.buildFromSheets(
    {
      DebtGrid: debtGrid,
      Summary: [['=HLOOKUP(0.2,DebtGrid!A1:K5,3,FALSE)', '=HLOOKUP(0.2,DebtGrid!A1:K5,4,FALSE)', '=HLOOKUP(0.2,DebtGrid!A1:K5,5,FALSE)']],
    },
    { maxRows: 100, maxColumns: 20, useColumnIndex },
  )
}

describe('HLOOKUP exact horizontal matches', () => {
  it.each([false, true])('returns the requested row from exact numeric horizontal matches with useColumnIndex=%s', (useColumnIndex) => {
    const workbook = buildIssue54Workbook(useColumnIndex)

    expectNumberClose(cellValue(workbook, 'Summary', 0, 0), 45616.01199262912)
    expectNumberClose(cellValue(workbook, 'Summary', 0, 1), 1.0233773685605514)
    expectNumberClose(cellValue(workbook, 'Summary', 0, 2), 0.10664524406185244)
  })

  it('updates exact HLOOKUP operands through the leaf table-lookup fast path', () => {
    const colCount = 128
    const endColumn = indexToColumn(colCount - 1)
    const workbook = WorkPaper.buildFromSheets({
      Bench: [
        Array.from({ length: colCount }, (_, index) => index + 1),
        Array.from({ length: colCount }, (_, index) => (index + 1) * 5),
        ['', '', '', 20, `=HLOOKUP(D3,A1:${endColumn}2,2,FALSE)`],
      ],
    })
    const sheetId = workbook.getSheetId('Bench')!

    expectNumberClose(cellValue(workbook, 'Bench', 2, 4), 100)

    workbook.resetPerformanceCounters()
    const changes = workbook.setCellContents(cell(sheetId, 2, 3), 99)

    expect(cellChanges(changes)).toEqual(new Set(['D3', 'E3']))
    expectNumberClose(cellValue(workbook, 'Bench', 2, 4), 495)
    expect(workbook.getStats().lastMetrics).toMatchObject({
      dirtyFormulaCount: 0,
      jsFormulaCount: 0,
      wasmFormulaCount: 0,
    })
    expect(workbook.getPerformanceCounters()).toMatchObject({
      directFormulaKernelSyncOnlyRecalcSkips: 1,
      formulasBound: 0,
      topoRepairs: 0,
    })
  })
})
