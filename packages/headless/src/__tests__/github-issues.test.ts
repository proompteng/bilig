import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'

import { WorkPaper } from '../index.js'

type TestCell = string | number | null

function cellValue(workbook: WorkPaper, sheetName: string, row: number, col: number): CellValue {
  return workbook.getCellValue({ sheet: workbook.getSheetId(sheetName), row, col })
}

function expectNumber(value: CellValue, expected: number): void {
  expect(value).toEqual({ tag: ValueTag.Number, value: expected })
}

function expectNumberClose(value: CellValue, expected: number): void {
  expect(value.tag).toBe(ValueTag.Number)
  if (value.tag !== ValueTag.Number) {
    throw new Error(`Expected number ${String(expected)}, received ${JSON.stringify(value)}`)
  }
  expect(value.value).toBeCloseTo(expected, 12)
}

function expectString(value: CellValue, expected: string): void {
  expect(value).toMatchObject({ tag: ValueTag.String, value: expected })
}

function readHeadlessPackageVersion(): string {
  const packageJson: unknown = JSON.parse(readFileSync(new URL('../../package.json', import.meta.url), 'utf8'))
  if (typeof packageJson !== 'object' || packageJson === null) {
    throw new Error('Expected headless package.json to parse as an object')
  }
  const version = Reflect.get(packageJson, 'version')
  if (typeof version !== 'string') {
    throw new Error('Expected headless package.json to contain a string version')
  }
  return version
}

describe('GitHub issue reductions', () => {
  it('evaluates initial cross-sheet references after all sheets are materialized', () => {
    const activity = Array.from({ length: 20 }, () => Array.from<TestCell>({ length: 8 }).fill(null))
    activity[0] = ['A', 'Type', 'C', 'D', 'E', 'F', 'G', 'Amount']
    activity[14][1] = 'Deposit'
    activity[14][7] = 3500
    activity[15][7] = -18_269.09

    const summary = Array.from({ length: 14 }, () => Array.from<TestCell>({ length: 5 }).fill(null))
    summary[8][2] = "=SUMIFS('Activity Detail'!$H$2:$H$20,'Activity Detail'!$B$2:$B$20,\"Deposit\")"
    summary[8][3] = '=C9-3500'
    summary[8][4] = '=IF(ABS(D9)<0.01,"PASS","FAIL")'
    summary[11][1] = '=Activity!H16'

    const payroll = Array.from({ length: 114 }, () => [null, null] as TestCell[])
    payroll[113][0] = 'AUSTIN, ZACHARY'
    payroll[113][1] = '=IFERROR(VLOOKUP(A114,Ref!$A$2:$B$20,2,FALSE),"Unmapped")'

    const ref = Array.from({ length: 20 }, () => [null, null] as TestCell[])
    ref[1] = ['ADRIAN, JOSEPH', 'Officers']
    ref[2] = ['ADRIAN, NICOLE', 'Officers']
    ref[3] = ['AUSTIN, ZACHARY', 'Captain Operations']
    ref[4] = ['BATSON, ADAM', 'Charter Manager']

    const comparison = Array.from({ length: 14 }, () => [null, null, null, null] as TestCell[])
    comparison[13][2] = 'txn-123'
    comparison[13][3] = '=IFERROR(XLOOKUP(C14,Bank!$D$2:$D$31,Bank!$B$2:$B$31,"",0),"")'

    const bank = Array.from({ length: 31 }, () => [null, null, null, null] as TestCell[])
    bank[1][1] = '2026-04-01'
    bank[1][3] = 'txn-123'

    const workbook = WorkPaper.buildFromSheets(
      {
        'Account Summary': summary,
        'Activity Detail': activity,
        Activity: activity,
        'Payroll Report_per cut off': payroll,
        Ref: ref,
        'Bank & Amex Comparison': comparison,
        Bank: bank,
      },
      { maxRows: 1000, maxColumns: 50, useColumnIndex: true },
    )

    expectNumber(cellValue(workbook, 'Account Summary', 8, 2), 3500)
    expectString(cellValue(workbook, 'Account Summary', 8, 4), 'PASS')
    expectNumber(cellValue(workbook, 'Account Summary', 11, 1), -18_269.09)
    expectString(cellValue(workbook, 'Payroll Report_per cut off', 113, 1), 'Captain Operations')
    expectString(cellValue(workbook, 'Bank & Amex Comparison', 13, 3), '2026-04-01')
  })

  it('counts populated cells in an initial cross-sheet COUNTA range', () => {
    const rows = Array.from({ length: 249 }, () => Array.from<TestCell>({ length: 10 }).fill(null))
    rows[248][5] = '=COUNTA(JE!$A$2:$A$71)'

    const je = Array.from({ length: 71 }, () => [null] as TestCell[])
    for (let row = 1; row <= 70; row += 1) {
      je[row][0] = `line-${row}`
    }

    const workbook = WorkPaper.buildFromSheets(
      {
        'PDF Extract': rows,
        JE: je,
      },
      { maxRows: 400, maxColumns: 20, useColumnIndex: true },
    )

    expectNumber(cellValue(workbook, 'PDF Extract', 248, 5), 70)
  })

  it('resolves row-offset INDEX and MATCH lookups during initial load', () => {
    const comparison = Array.from({ length: 14 }, () => [null, null, null, null, null] as TestCell[])
    comparison[13][2] = 'txn-123'
    comparison[13][3] = '=IFERROR(INDEX(Bank!$B$2:$B$31,MATCH(C14,Bank!$D$2:$D$31,0)),"")'
    comparison[13][4] = '=IFERROR(INDEX(Bank!$D$2:$D$31,MATCH(C14,Bank!$D$2:$D$31,0)),0)'

    const bank = Array.from({ length: 31 }, () => [null, null, null, null] as TestCell[])
    bank[1][1] = '2026-04-01'
    bank[1][3] = 'txn-123'

    const workbook = WorkPaper.buildFromSheets(
      {
        'Bank & Amex Comparison': comparison,
        Bank: bank,
      },
      { maxRows: 1000, maxColumns: 50, useColumnIndex: true },
    )

    expectString(cellValue(workbook, 'Bank & Amex Comparison', 13, 3), '2026-04-01')
    expectString(cellValue(workbook, 'Bank & Amex Comparison', 13, 4), 'txn-123')
  })

  it('uses HYPERLINK friendly names as calculated display values', () => {
    const workbook = WorkPaper.buildFromSheets(
      {
        Sheet1: [
          ['internal link', '=HYPERLINK("#\'Summary\'!A1","Go to Summary")'],
          ['external link', '=HYPERLINK("https://example.com/workpaper","Link")'],
        ],
        Summary: [['ok']],
      },
      { maxRows: 20, maxColumns: 10, useColumnIndex: true },
    )

    expectString(cellValue(workbook, 'Sheet1', 0, 1), 'Go to Summary')
    expectString(cellValue(workbook, 'Sheet1', 1, 1), 'Link')
  })

  it('formats Excel date serials with TEXT formulas below the input row', () => {
    const workbook = WorkPaper.buildFromSheets(
      {
        Sheet1: [[46_127], ['formatted', '=TEXT(A1,"mm.dd.yy")']],
      },
      { maxRows: 20, maxColumns: 10, useColumnIndex: true },
    )

    expectString(cellValue(workbook, 'Sheet1', 1, 1), '04.15.26')
  })

  it('evaluates date functions below the input row when they reference A1', () => {
    const workbook = WorkPaper.buildFromSheets(
      {
        Dates: [
          [46_127],
          ['workday', '=WORKDAY(A1,2)'],
          ['edate', '=EDATE(A1,1)'],
          ['eomonth', '=EOMONTH(A1,1)'],
          ['day', '=DAY(A1)'],
          ['month', '=MONTH(A1)'],
        ],
      },
      { maxRows: 20, maxColumns: 10, useColumnIndex: true },
    )

    expectNumber(cellValue(workbook, 'Dates', 1, 1), 46_129)
    expectNumber(cellValue(workbook, 'Dates', 2, 1), 46_157)
    expectNumber(cellValue(workbook, 'Dates', 3, 1), 46_173)
    expectNumber(cellValue(workbook, 'Dates', 4, 1), 15)
    expectNumber(cellValue(workbook, 'Dates', 5, 1), 4)
  })

  it('throws a public timeout error when initial evaluation exceeds the configured budget', () => {
    const config = {
      maxRows: 10,
      maxColumns: 4,
      useColumnIndex: true,
      evaluationTimeoutMs: 0,
    } as Parameters<typeof WorkPaper.buildFromSheets>[1]

    let thrown: unknown
    try {
      WorkPaper.buildFromSheets(
        {
          Sheet1: [
            [1, '=A1+1'],
            [null, '=B1+1'],
          ],
        },
        config,
      )
    } catch (error) {
      thrown = error
    }

    expect(thrown).toBeInstanceOf(Error)
    if (!(thrown instanceof Error)) {
      throw new Error('Expected timeout to throw an Error instance')
    }
    expect(thrown.name).toBe('WorkPaperEvaluationTimeoutError')
    expect(thrown.message).toContain('timed out')
  })

  it('keeps issue #7 wrapped criteria aggregates on the direct evaluator path', () => {
    const rowCount = 5_000
    const formulaCount = 20
    const data = Array.from({ length: rowCount }, (_, row) => [row % 3 === 0 ? 'Deposit' : 'Withdrawal', row + 1])
    const formulas = Array.from({ length: formulaCount }, (_, row) => [
      row,
      `=IFERROR(ROUND(SUMIFS(Data!$B$1:$B$${rowCount},Data!$A$1:$A$${rowCount},"Deposit"),2),0)`,
    ])

    const workbook = WorkPaper.buildFromSheets(
      {
        Data: data,
        Formulas: formulas,
      },
      { maxRows: rowCount + formulaCount + 10, maxColumns: 8, useColumnIndex: true },
    )

    expectNumber(cellValue(workbook, 'Formulas', formulaCount - 1, 1), 4_167_500)
    expect(workbook.getPerformanceCounters().directFormulaInitialEvaluations).toBe(formulaCount)
  })

  it('recalculates rounded criteria aggregates after edits instead of applying raw deltas', () => {
    const workbook = WorkPaper.buildFromSheets(
      {
        Data: [
          ['Deposit', 1.235],
          ['Deposit', 0],
        ],
        Summary: [['=ROUND(SUMIFS(Data!$B$1:$B$2,Data!$A$1:$A$2,"Deposit"),2)']],
      },
      { maxRows: 10, maxColumns: 4, useColumnIndex: true },
    )
    const dataSheetId = workbook.getSheetId('Data')

    expectNumber(cellValue(workbook, 'Summary', 0, 0), 1.24)
    workbook.setCellContents({ sheet: dataSheetId, row: 0, col: 1 }, 1.236)

    expectNumber(cellValue(workbook, 'Summary', 0, 0), 1.24)
  })

  it('matches Excel cached results for rounded AVERAGE ranges with blanks and text', () => {
    const rows = Array.from({ length: 316 }, () => Array.from<TestCell>({ length: 13 }).fill(null))
    const numericValues = [12.5, 24, 18.75, 20.25, 19.04]
    for (let index = 0; index < numericValues.length; index += 1) {
      rows[1 + index][4] = numericValues[index]!
    }
    rows[6][4] = 'Department'
    rows[7][4] = ''
    rows[8][4] = null
    rows[314][4] = null
    rows[315][4] = '=ROUND(AVERAGE(E2:E315),2)'

    const workbook = WorkPaper.buildFromSheets(
      {
        'Action on requests': rows,
      },
      { maxRows: 400, maxColumns: 20, useColumnIndex: true },
    )

    expectNumber(cellValue(workbook, 'Action on requests', 315, 4), 18.91)
  })

  it('exposes issue #24 XIRR invalid-date diagnostics without breaking numeric-date XIRR', () => {
    const workbook = WorkPaper.buildFromSheets(
      {
        Tax: [
          ['Metric', 'Value', 'Date serial', 'All negative', 'Text cash flow'],
          ['Cash flows', -100_000, 45_292, -100_000, -100_000],
          ['Year 1', 25_000, 45_658, -25_000, 25_000],
          ['Year 2', 35_000, 46_023, -35_000, 'bad'],
          ['Year 3', 45_000, 46_388, -45_000, 45_000],
          ['IRR', '=IRR(B2:B5)', null],
          ['XIRR', '=XIRR(B2:B5,C2:C5)', null],
          ['Invalid XIRR', '=XIRR(B2:B5,A2:A5)', null],
          ['Mismatched XIRR', '=XIRR(B2:B5,C2:C4)', null],
          ['Missing positive XIRR', '=XIRR(D2:D5,C2:C5)', null],
          ['Invalid cash XIRR', '=XIRR(E2:E5,C2:C5)', null],
        ],
      },
      { maxRows: 100_000, maxColumns: 512, useColumnIndex: true },
    )

    expectNumberClose(cellValue(workbook, 'Tax', 5, 1), 0.02259730507537016)
    expectNumberClose(cellValue(workbook, 'Tax', 6, 1), 0.02256857579463996)
    expect(cellValue(workbook, 'Tax', 7, 1)).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    const tax = workbook.getSheetId('Tax')!
    const invalidXirr = { sheet: tax, row: 7, col: 1 }
    expect(workbook.getCellDisplayValue(invalidXirr)).toBe('#VALUE!')
    expect(workbook.getCellFormulaDiagnostics(invalidXirr)).toMatchObject([
      {
        severity: 'error',
        code: 'financial-unsupported-date-coercion',
        functionName: 'XIRR',
        errorText: '#VALUE!',
        references: ['Tax!A2:A5', 'Tax!A2'],
      },
    ])
    expect(workbook.getCellFormulaDiagnostics(invalidXirr)[0]?.message).toContain('Use numeric Excel serial dates')
    expect(workbook.getCellFormulaDiagnostics({ sheet: tax, row: 8, col: 1 })[0]?.code).toBe('financial-mismatched-dimensions')
    expect(workbook.getCellFormulaDiagnostics({ sheet: tax, row: 9, col: 1 })[0]?.code).toBe('financial-missing-positive-cash-flow')
    expect(workbook.getCellFormulaDiagnostics({ sheet: tax, row: 10, col: 1 })[0]?.code).toBe('financial-invalid-cash-flow')
  })

  it('evaluates issue #24 XIRR over formula-derived numeric cash-flow cells', () => {
    const workbook = WorkPaper.buildFromSheets(
      {
        Project: [
          ['Metric', 'Amount'],
          ['Equity required', 37_237_200],
          ['Sale proceeds', 101_705_094.7368421],
          ['Project IRR', '=XIRR(B6:B7,A6:A7)'],
          ['Date', 'Cash flow'],
          [44_927, '=-B2'],
          [46_388, '=B3'],
          [null, null],
          ['Literal project IRR', '=XIRR(B10:B11,A10:A11)'],
          [44_927, -37_237_200],
          [46_388, 101_705_094.7368421],
        ],
      },
      { maxRows: 100_000, maxColumns: 512, useColumnIndex: true },
    )

    expectNumberClose(cellValue(workbook, 'Project', 5, 1), -37_237_200)
    expectNumberClose(cellValue(workbook, 'Project', 6, 1), 101_705_094.7368421)
    expectNumberClose(cellValue(workbook, 'Project', 3, 1), 0.28533624352898757)
    expect(cellValue(workbook, 'Project', 3, 1)).toEqual(cellValue(workbook, 'Project', 8, 1))
  })

  it('reports the published package version through WorkPaper.version', () => {
    expect(WorkPaper.version).toBe(readHeadlessPackageVersion())
  })
})
