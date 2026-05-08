import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { ErrorCode, ValueTag, type CellValue, type WorkbookSnapshot } from '@bilig/protocol'

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

function slope(ys: readonly number[], xs: readonly number[]): number {
  const xMean = xs.reduce((sum, value) => sum + value, 0) / xs.length
  const yMean = ys.reduce((sum, value) => sum + value, 0) / ys.length
  let numerator = 0
  let denominator = 0
  for (let index = 0; index < xs.length; index += 1) {
    numerator += (xs[index] - xMean) * (ys[index] - yMean)
    denominator += (xs[index] - xMean) ** 2
  }
  return numerator / denominator
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

  it('resolves issue #93 blank-reference formulas as numeric zero', () => {
    const workbook = WorkPaper.buildFromSheets(
      {
        Inputs: [[null, null]],
        Summary: [
          ['=Inputs!A1', '=Inputs!$A$1', '=IF(Inputs!A1="",Inputs!B1,Inputs!A1)', '=Inputs!A1+1', '=SUM(Inputs!A1)', '="x"&Inputs!A1'],
          [null, '=A2', '=IF(A2="",D2,A2)', null, '=T(1)', '=""'],
        ],
      },
      { maxRows: 8, maxColumns: 8, useColumnIndex: true },
    )
    const inputs = workbook.getSheetId('Inputs')!

    expectNumber(cellValue(workbook, 'Summary', 0, 0), 0)
    expectNumber(cellValue(workbook, 'Summary', 0, 1), 0)
    expectNumber(cellValue(workbook, 'Summary', 0, 2), 0)
    expectNumber(cellValue(workbook, 'Summary', 1, 1), 0)
    expectNumber(cellValue(workbook, 'Summary', 1, 2), 0)

    expectNumber(cellValue(workbook, 'Summary', 0, 3), 1)
    expectNumber(cellValue(workbook, 'Summary', 0, 4), 0)
    expectString(cellValue(workbook, 'Summary', 0, 5), 'x')
    expect(cellValue(workbook, 'Summary', 1, 4)).toEqual({ tag: ValueTag.Empty })
    expectString(cellValue(workbook, 'Summary', 1, 5), '')

    workbook.setCellContents({ sheet: inputs, row: 0, col: 0 }, 7)

    expectNumber(cellValue(workbook, 'Summary', 0, 0), 7)
    expectNumber(cellValue(workbook, 'Summary', 0, 1), 7)
    expectNumber(cellValue(workbook, 'Summary', 0, 2), 7)
    expectNumber(cellValue(workbook, 'Summary', 0, 3), 8)
    expectNumber(cellValue(workbook, 'Summary', 0, 4), 7)
    expectString(cellValue(workbook, 'Summary', 0, 5), 'x7')

    workbook.setCellContents({ sheet: inputs, row: 0, col: 0 }, null)

    expectNumber(cellValue(workbook, 'Summary', 0, 0), 0)
    expectNumber(cellValue(workbook, 'Summary', 0, 1), 0)
    expectNumber(cellValue(workbook, 'Summary', 0, 2), 0)
    expectNumber(cellValue(workbook, 'Summary', 0, 3), 1)
    expectNumber(cellValue(workbook, 'Summary', 0, 4), 0)
    expectString(cellValue(workbook, 'Summary', 0, 5), 'x')
  })

  it('resolves issue #92 whole-column criteria ranges in conditional aggregates', () => {
    const workbook = WorkPaper.buildFromSheets(
      {
        Data: [
          ['Year', 'Line', 'Amount'],
          [2024, 'Revenue', 1000],
          [2024, 'Revenue', 1922],
          [2024, 'COGS', -700],
          [2024, 'COGS', -701],
          [2025, 'Revenue', 3000],
          [2025, 'Revenue', 3977],
        ],
        Summary: [
          ['Line', 2024, 2025, 'Whole-column', 'Bounded'],
          [
            'Revenue',
            null,
            null,
            '=SUMIFS(Data!$C:$C,Data!$B:$B,$A2,Data!$A:$A,B$1)',
            '=SUMIFS(Data!$C$2:$C$7,Data!$B$2:$B$7,$A2,Data!$A$2:$A$7,B$1)',
          ],
          [
            'COGS',
            null,
            null,
            '=SUMIFS(Data!$C:$C,Data!$B:$B,$A3,Data!$A:$A,B$1)',
            '=SUMIFS(Data!$C$2:$C$7,Data!$B$2:$B$7,$A3,Data!$A$2:$A$7,B$1)',
          ],
          [
            'Other',
            null,
            null,
            '=SUMIFS(Data!$C:$C,Data!$B:$B,$A4,Data!$A:$A,B$1)',
            '=SUMIFS(Data!$C$2:$C$7,Data!$B$2:$B$7,$A4,Data!$A$2:$A$7,B$1)',
          ],
          [
            'Revenue count',
            null,
            null,
            '=COUNTIFS(Data!$B:$B,"Revenue",Data!$A:$A,C$1)',
            '=COUNTIFS(Data!$B$2:$B$7,"Revenue",Data!$A$2:$A$7,C$1)',
          ],
          [
            'Revenue average',
            null,
            null,
            '=AVERAGEIFS(Data!$C:$C,Data!$B:$B,"Revenue")',
            '=AVERAGEIFS(Data!$C$2:$C$7,Data!$B$2:$B$7,"Revenue")',
          ],
        ],
      },
      { maxRows: 64, maxColumns: 8, useColumnIndex: true },
    )
    const data = workbook.getSheetId('Data')!

    expectNumber(cellValue(workbook, 'Summary', 1, 4), 2922)
    expectNumber(cellValue(workbook, 'Summary', 2, 4), -1401)
    expectNumber(cellValue(workbook, 'Summary', 3, 4), 0)
    expectNumber(cellValue(workbook, 'Summary', 4, 4), 2)
    expectNumberClose(cellValue(workbook, 'Summary', 5, 4), 2474.75)

    expectNumber(cellValue(workbook, 'Summary', 1, 3), 2922)
    expectNumber(cellValue(workbook, 'Summary', 2, 3), -1401)
    expectNumber(cellValue(workbook, 'Summary', 3, 3), 0)
    expectNumber(cellValue(workbook, 'Summary', 4, 3), 2)
    expectNumberClose(cellValue(workbook, 'Summary', 5, 3), 2474.75)

    workbook.setCellContents({ sheet: data, row: 7, col: 0 }, 2024)
    workbook.setCellContents({ sheet: data, row: 7, col: 1 }, 'Revenue')
    workbook.setCellContents({ sheet: data, row: 7, col: 2 }, 78)

    expectNumber(cellValue(workbook, 'Summary', 1, 4), 2922)
    expectNumber(cellValue(workbook, 'Summary', 1, 3), 3000)
    expectNumberClose(cellValue(workbook, 'Summary', 5, 4), 2474.75)
    expectNumberClose(cellValue(workbook, 'Summary', 5, 3), 1995.4)
  })

  it('resolves sheet-scoped defined names from imported snapshots before broken globals', () => {
    const snapshot: WorkbookSnapshot = {
      version: 1,
      workbook: {
        name: 'sheet-scoped-defined-names',
        metadata: {
          definedNames: [
            { name: 'LocalBonus', scopeSheetName: 'Local', value: { kind: 'cell-ref', sheetName: 'Local', address: 'A1' } },
            { name: 'LocalRevenue', scopeSheetName: 'Local', value: { kind: 'cell-ref', sheetName: 'Local', address: 'B1' } },
            { name: 'LocalBonus', value: { kind: 'formula', formula: '=#REF!' } },
          ],
        },
      },
      sheets: [
        {
          id: 1,
          name: 'Global',
          order: 0,
          cells: [{ address: 'A1', value: 100 }],
        },
        {
          id: 2,
          name: 'Local',
          order: 1,
          cells: [
            { address: 'A1', value: 7 },
            { address: 'B1', value: 10 },
            { address: 'C1', formula: 'LocalBonus*LocalRevenue' },
          ],
        },
      ],
    }

    const workbook = WorkPaper.buildFromSnapshot(snapshot, { maxRows: 20, maxColumns: 8, useColumnIndex: true })
    const localId = workbook.getSheetId('Local')!

    expectNumber(cellValue(workbook, 'Local', 0, 2), 70)
    expect(workbook.getCellFormula({ sheet: localId, row: 0, col: 2 })).toBe('=LocalBonus*LocalRevenue')
  })

  it('coerces double-unary range comparisons inside SUMPRODUCT', () => {
    const workbook = WorkPaper.buildFromSheets(
      {
        Sheet1: [
          [10, 20, 30, 40, 50, 60],
          ['Series A', 'Series B', 'Series A', 'Series C', 'Series A', 'Series B'],
          [1, 0, 1, 0, 1, 0],
          [
            '=SUMPRODUCT($A$1:$F$1,--($A$2:$F$2="Series A"))',
            '=SUMPRODUCT(--($A$2:$F$2="Series A"))',
            '=SUMPRODUCT($A$1:$F$1,($A$2:$F$2="Series A")*1)',
            '=SUMPRODUCT($A$1:$F$1,$A$3:$F$3)',
          ],
        ],
      },
      { maxRows: 20, maxColumns: 8, useColumnIndex: true },
    )

    expectNumber(cellValue(workbook, 'Sheet1', 3, 0), 90)
    expectNumber(cellValue(workbook, 'Sheet1', 3, 1), 3)
    expectNumber(cellValue(workbook, 'Sheet1', 3, 2), 90)
    expectNumber(cellValue(workbook, 'Sheet1', 3, 3), 90)
  })

  it('passes INDEX-wrapped boolean arrays into MATCH', () => {
    const workbook = WorkPaper.buildFromSheets(
      {
        Sheet1: [
          [0, 0, 25, 40, 0, 60],
          [
            '=MATCH(TRUE,A1:F1<>0,0)',
            '=INDEX(A1:F1,MATCH(TRUE,A1:F1<>0,0))',
            '=MATCH(TRUE,INDEX(A1:F1<>0,),0)',
            '=INDEX(A1:F1,MATCH(TRUE,INDEX(A1:F1<>0,),0))',
            '=MATCH(INDEX(A1:F1,MATCH(TRUE,INDEX(A1:F1<>0,),0)),A1:F1,0)',
          ],
        ],
      },
      { maxRows: 10, maxColumns: 8, useColumnIndex: true },
    )

    expectNumber(cellValue(workbook, 'Sheet1', 1, 0), 3)
    expectNumber(cellValue(workbook, 'Sheet1', 1, 1), 25)
    expectNumber(cellValue(workbook, 'Sheet1', 1, 2), 3)
    expectNumber(cellValue(workbook, 'Sheet1', 1, 3), 25)
    expectNumber(cellValue(workbook, 'Sheet1', 1, 4), 3)
  })

  it('resolves INDEX zero row and column vector selections to scalar cells', () => {
    const workbook = WorkPaper.buildFromSheets(
      {
        Sheet1: [
          ['na', 'Non Participating Preferred', 'Shadow', 'Participating Preferred', 'na', 'Exit', 'Seed', 'Common'],
          ['Seed', 'Series A', 'Series B', 'Series C', 'Series D', 'Exit', 'Series A', 'Non Participating Preferred'],
          [100, 150, 200, 300, 400, 500, 'Series B', 'Shadow'],
          [null, null, null, null, null, null, 'Series C', 'Participating Preferred'],
          [null, null, null, null, null, null, 'Series D', 400],
          [null, null, null, null, null, null, 'Exit', 500],
          [
            '=INDEX($A$1:$F$1,0,MATCH("Series C",$A$2:$F$2,0))',
            '=INDEX($A$3:$F$3,0,MATCH("Series C",$A$2:$F$2,0))',
            '=INDEX($H$1:$H$6,MATCH("Series C",$G$1:$G$6,0),0)',
          ],
        ],
      },
      { maxRows: 20, maxColumns: 10, useColumnIndex: true },
    )

    expectString(cellValue(workbook, 'Sheet1', 6, 0), 'Participating Preferred')
    expectNumber(cellValue(workbook, 'Sheet1', 6, 1), 300)
    expectString(cellValue(workbook, 'Sheet1', 6, 2), 'Participating Preferred')
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

  it('honors holiday ranges for WORKDAY and NETWORKDAYS formulas', () => {
    const workbook = WorkPaper.buildFromSheets(
      {
        Dates: [
          ['start', '=DATE(2026,1,1)'],
          ['end', '=DATE(2026,1,10)'],
          ['first holiday', '=DATE(2026,1,1)'],
          ['second holiday', '=DATE(2026,1,2)'],
          ['workday with holidays', '=WORKDAY(B1,3,B3:B4)'],
          ['networkdays with holidays', '=NETWORKDAYS(B1,B2,B3:B4)'],
          ['workday no holidays', '=WORKDAY(B1,3)'],
          ['networkdays no holidays', '=NETWORKDAYS(B1,B2)'],
          ['workday negative holidays', '=WORKDAY(B2,-3,B3:B4)'],
        ],
      },
      { maxRows: 20, maxColumns: 10, useColumnIndex: true },
    )

    expectNumber(cellValue(workbook, 'Dates', 4, 1), 46_029)
    expectNumber(cellValue(workbook, 'Dates', 5, 1), 5)
    expectNumber(cellValue(workbook, 'Dates', 6, 1), 46_028)
    expectNumber(cellValue(workbook, 'Dates', 7, 1), 7)
    expectNumber(cellValue(workbook, 'Dates', 8, 1), 46_029)
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
  }, 15_000)

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

  it('treats issue #95 undefined and sparse buildFromSheets cells as blanks', () => {
    const leadingSparseRow: unknown[] = []
    leadingSparseRow[1] = 'Revenue'
    const middleSparseRow: unknown[] = Array(3)
    middleSparseRow[0] = 1
    middleSparseRow[2] = 3
    const allSparseRow: unknown[] = Array(3)
    const sheets = {
      ExplicitUndefined: [[undefined]],
      LeadingSparse: [leadingSparseRow],
      MiddleSparse: [middleSparseRow],
      AllSparse: [allSparseRow],
      FormulaSparse: [[undefined, '=A1+1']],
    } satisfies Record<string, readonly (readonly unknown[])[]>

    const buildFromSheets: unknown = Reflect.get(WorkPaper, 'buildFromSheets')
    if (typeof buildFromSheets !== 'function') {
      throw new Error('Expected WorkPaper.buildFromSheets to be callable')
    }
    const buildResult: unknown = Reflect.apply(buildFromSheets, WorkPaper, [sheets, { maxRows: 8, maxColumns: 8, useColumnIndex: true }])
    expect(buildResult).toBeInstanceOf(WorkPaper)
    if (!(buildResult instanceof WorkPaper)) {
      throw new Error('Expected buildFromSheets to return a WorkPaper')
    }
    const workbook = buildResult

    expect(workbook.getSheetDimensions(workbook.getSheetId('ExplicitUndefined'))).toEqual({ width: 0, height: 0 })
    expect(workbook.getSheetDimensions(workbook.getSheetId('AllSparse'))).toEqual({ width: 0, height: 0 })
    expect(workbook.getSheetDimensions(workbook.getSheetId('LeadingSparse'))).toEqual({ width: 2, height: 1 })
    expect(workbook.getSheetDimensions(workbook.getSheetId('MiddleSparse'))).toEqual({ width: 3, height: 1 })
    expect(cellValue(workbook, 'LeadingSparse', 0, 0).tag).toBe(ValueTag.Empty)
    expectString(cellValue(workbook, 'LeadingSparse', 0, 1), 'Revenue')
    expectNumber(cellValue(workbook, 'MiddleSparse', 0, 0), 1)
    expect(cellValue(workbook, 'MiddleSparse', 0, 1).tag).toBe(ValueTag.Empty)
    expectNumber(cellValue(workbook, 'MiddleSparse', 0, 2), 3)
    expectNumber(cellValue(workbook, 'FormulaSparse', 0, 1), 1)
  })

  it('resolves issue #113 LOOKUP latest-value array idioms used by financial schedules', () => {
    const rows: TestCell[][] = [
      [10, null, null, 1, 10],
      ['x', 'Seed', null, 1, 20],
      [20, '', null, 1, 30],
      [null, 'Series A', null, null, null],
      [30, 'Series C', null, null, null],
      ['skip', '', null, null, null],
      [
        '=LOOKUP(2,1/(ISNUMBER(A1:A6)),A1:A6)',
        '=LOOKUP(2,1/(B1:B6<>""),B1:B6)',
        '=INDEX(E1:E3,MATCH(2,D1:D3,1))',
        '=LOOKUP(2,D1:D3,E1:E3)',
        '=IFERROR(LOOKUP(2,1/(ISNUMBER(A1:A6)),A1:A6),"na")',
      ],
    ]

    const workbook = WorkPaper.buildFromSheets({ Sheet1: rows }, { maxRows: 20, maxColumns: 10, useColumnIndex: true })

    expectNumber(cellValue(workbook, 'Sheet1', 6, 0), 30)
    expectString(cellValue(workbook, 'Sheet1', 6, 1), 'Series C')
    expectNumber(cellValue(workbook, 'Sheet1', 6, 2), 30)
    expectNumber(cellValue(workbook, 'Sheet1', 6, 3), 30)
    expectNumber(cellValue(workbook, 'Sheet1', 6, 4), 30)
  })

  it('resolves issue #105 VLOOKUP cell-reference keys across translated formula templates', () => {
    const workbook = WorkPaper.buildFromSheets(
      {
        Sheet1: [
          [null, null, null, 'Key', 'Day'],
          [1, null, null, 1, 'Sunday'],
          [2, '=VLOOKUP(A3,$D$2:$E$8,2,FALSE)', null, 2, 'Monday'],
          [3, '=VLOOKUP(A4,$D$2:$E$8,2,FALSE)', null, 3, 'Tuesday'],
          [4, '=VLOOKUP(A5,$D$2:$E$8,2,FALSE)', null, 4, 'Wednesday'],
          [5, null, null, 5, 'Thursday'],
          [6, null, null, 6, 'Friday'],
          [7, null, null, 7, 'Saturday'],
        ],
        'Step 1': [
          [null, null, null, null],
          [null, null, 'Key', 'Value'],
          [null, null, 3, 'Three'],
          [null, null, 4, 'Four'],
          [null, null, 5, 'Five'],
        ],
        Quoted: [
          [null, null],
          [null, null],
          [4, "=VLOOKUP(A3,'Step 1'!$C$3:$D$5,2,FALSE)"],
        ],
      },
      { maxRows: 40, maxColumns: 10, useColumnIndex: true },
    )

    expectString(cellValue(workbook, 'Sheet1', 2, 1), 'Monday')
    expectString(cellValue(workbook, 'Sheet1', 3, 1), 'Tuesday')
    expectString(cellValue(workbook, 'Sheet1', 4, 1), 'Wednesday')
    expectString(cellValue(workbook, 'Quoted', 2, 1), 'Four')
  })

  it('resolves issues #101, #107, and #109 range-valued defined names with scalar implicit intersection', () => {
    const rows = Array.from({ length: 43 }, () => Array.from<TestCell>({ length: 8 }).fill(null))
    rows[2][2] = 1
    rows[2][3] = 2
    rows[2][4] = 3
    rows[3][2] = 0.05
    rows[3][3] = 0.052
    rows[3][4] = 0.055
    rows[4][5] = 2
    rows[4][3] = '=TimeValues^3'
    rows[5][5] = 3
    rows[6][0] = 1
    rows[6][2] = '=1/(1+rate+(Year_=Year)*0.0001)^Year'
    rows[6][5] = 5
    rows[6][6] = '=2*3*TimeValues^1'
    rows[7][0] = 2
    rows[7][3] = '=1/(1+rate+(Year_=Year)*0.0001)^Year'
    rows[7][5] = 7
    rows[7][6] = '="a_"&TimeValues'
    rows[8][0] = 3
    rows[8][4] = '=1/(1+rate+(Year_=Year)*0.0001)^Year'
    rows[9][6] = '=SUM(TimeValues)'
    rows[40][2] = 10
    rows[40][3] = 20
    rows[41][2] = '=HorizontalValues*C43'
    rows[41][3] = '=HorizontalValues*D43'
    rows[41][5] = '=INDEX(HorizontalValues,1,2)'
    rows[41][6] = '=SUM(HorizontalValues)'
    rows[42][2] = 2
    rows[42][3] = 3

    const snapshot: WorkbookSnapshot = {
      version: 1,
      workbook: {
        name: 'range-defined-name-implicit-intersection',
        metadata: {
          definedNames: [
            { name: 'rate', value: { kind: 'range-ref', sheetName: 'Model', startAddress: 'C4', endAddress: 'E4' } },
            { name: 'Year', value: { kind: 'range-ref', sheetName: 'Model', startAddress: 'C3', endAddress: 'E3' } },
            { name: 'Year_', value: { kind: 'range-ref', sheetName: 'Model', startAddress: 'A7', endAddress: 'A9' } },
            { name: 'TimeValues', value: { kind: 'range-ref', sheetName: 'Model', startAddress: 'F5', endAddress: 'F8' } },
            {
              name: 'HorizontalValues',
              value: { kind: 'range-ref', sheetName: 'Model', startAddress: 'C41', endAddress: 'D41' },
            },
          ],
        },
      },
      sheets: [
        {
          id: 1,
          name: 'Model',
          order: 0,
          cells: rows.flatMap((row, rowIndex) =>
            row.flatMap((content, colIndex) => {
              if (content === null) {
                return []
              }
              const address = `${String.fromCharCode(65 + colIndex)}${rowIndex + 1}`
              return typeof content === 'string' && content.startsWith('=')
                ? [{ address, formula: content.slice(1) }]
                : [{ address, value: content }]
            }),
          ),
        },
      ],
    }

    const workbook = WorkPaper.buildFromSnapshot(snapshot, { maxRows: 80, maxColumns: 16, useColumnIndex: true })

    expectNumberClose(cellValue(workbook, 'Model', 6, 2), 0.9522902580706599)
    expectNumberClose(cellValue(workbook, 'Model', 7, 3), 0.9034122159454043)
    expectNumberClose(cellValue(workbook, 'Model', 8, 4), 0.8513715450616418)
    expectNumber(cellValue(workbook, 'Model', 4, 3), 8)
    expectNumber(cellValue(workbook, 'Model', 6, 6), 30)
    expectString(cellValue(workbook, 'Model', 7, 6), 'a_7')
    expectNumber(cellValue(workbook, 'Model', 9, 6), 17)
    expectNumber(cellValue(workbook, 'Model', 41, 2), 20)
    expectNumber(cellValue(workbook, 'Model', 41, 3), 60)
    expectNumber(cellValue(workbook, 'Model', 41, 5), 20)
    expectNumber(cellValue(workbook, 'Model', 41, 6), 30)
  })

  it('resolves issue #110 adjacent SLOPE formulas that share an absolute x-range', () => {
    const buildRows = (withFirstSlope: boolean): TestCell[][] => {
      const rows = Array.from({ length: 28 }, () => Array.from<TestCell>({ length: 16 }).fill(null))
      const firstYValues = [1 / 100, 1 / 101, 1 / 102, 1 / 103]
      const secondYValues = [1 / 200, 1 / 201, 1 / 202, 1 / 203]
      const xValues = [1 / 300, 1 / 301, 1 / 302, 1 / 303]
      for (let index = 0; index < xValues.length; index += 1) {
        const row = 19 + index
        rows[row][11] = firstYValues[index]
        rows[row][12] = secondYValues[index]
        rows[row][13] = xValues[index]
      }
      if (withFirstSlope) {
        rows[14][11] = '=SLOPE(L20:L23,$N$20:$N$23)'
      }
      rows[14][12] = '=SLOPE(M20:M23,$N$20:$N$23)'
      return rows
    }
    const firstYValues = [1 / 100, 1 / 101, 1 / 102, 1 / 103]
    const secondYValues = [1 / 200, 1 / 201, 1 / 202, 1 / 203]
    const xValues = [1 / 300, 1 / 301, 1 / 302, 1 / 303]
    const workbook = WorkPaper.buildFromSheets(
      {
        SharedX: buildRows(true),
        SingleSlope: buildRows(false),
      },
      { maxRows: 100, maxColumns: 20, useColumnIndex: true },
    )
    const sharedX = workbook.getSheetId('SharedX')!
    const singleSlope = workbook.getSheetId('SingleSlope')!
    const expectSlopeNumber = (value: CellValue, expected: number): void => {
      expect(value.tag).toBe(ValueTag.Number)
      if (value.tag !== ValueTag.Number) {
        throw new Error(`Expected number ${String(expected)}, received ${JSON.stringify(value)}`)
      }
      expect(value.value).toBeCloseTo(expected, 9)
    }

    expectSlopeNumber(cellValue(workbook, 'SharedX', 14, 11), slope(firstYValues, xValues))
    expectSlopeNumber(cellValue(workbook, 'SharedX', 14, 12), slope(secondYValues, xValues))
    expectSlopeNumber(cellValue(workbook, 'SingleSlope', 14, 12), slope(secondYValues, xValues))
    expect(workbook.getCellFormulaDiagnostics({ sheet: sharedX, row: 14, col: 12 })).toEqual([])
    expect(workbook.getCellFormulaDiagnostics({ sheet: singleSlope, row: 14, col: 12 })).toEqual([])
  })

  it('reports the published package version through WorkPaper.version', () => {
    expect(WorkPaper.version).toBe(readHeadlessPackageVersion())
  })

  it('resolves issue #106 normal CDF drift in option-pricing formulas', () => {
    const workbook = WorkPaper.buildFromSheets(
      {
        Sheet1: [
          [
            '=NORMSDIST(-0.8281017980432489)',
            '=NORMSDIST(-0.9281017980432489)',
            '=NORM.DIST(-0.8281017980432489,0,1,TRUE)',
            '=NORM.DIST(-0.8281017980432489,0,1,FALSE)',
          ],
        ],
      },
      { maxRows: 10, maxColumns: 8, useColumnIndex: true },
    )

    expectNumberClose(cellValue(workbook, 'Sheet1', 0, 0), 0.203806425664055)
    expectNumberClose(cellValue(workbook, 'Sheet1', 0, 1), 0.17667738351319964)
    expectNumberClose(cellValue(workbook, 'Sheet1', 0, 2), 0.203806425664055)
    expectNumberClose(cellValue(workbook, 'Sheet1', 0, 3), 0.2831397103054239)
  })

  it('resolves issue #104 whole-column AVERAGE references', () => {
    const rows = Array.from({ length: 25 }, () => Array.from<TestCell>({ length: 30 }).fill(null))
    const values = [5, 10, 20, 25, 40]

    rows[13][21] = 'ignored'
    values.forEach((value, index) => {
      rows[index + 14][21] = value
    })
    rows[0][0] = '=AVERAGE(V:V)'
    rows[0][1] = '=AVERAGE(V15:V19)'
    rows[0][2] = '=SUM(V:V)'
    rows[0][3] = '=MAX(V:V)'
    rows[0][4] = '=MIN(V:V)'
    rows[0][5] = '=COUNT(V:V)'
    rows[0][6] = '=AVG(V:V)'

    const data = Array.from({ length: 8 }, () => [null] as TestCell[])
    values.forEach((value, index) => {
      data[index + 1][0] = value
    })
    rows[1][0] = '=AVERAGE(Data!$A:$A)'
    rows[1][1] = '=AVERAGE(Data!$A$2:$A$6)'
    rows[1][2] = '=SUM(Data!$A:$A)'
    rows[1][3] = '=MAX(Data!$A:$A)'
    rows[1][4] = '=MIN(Data!$A:$A)'
    rows[1][5] = '=COUNT(Data!$A:$A)'
    rows[1][6] = '=AVG(Data!$A:$A)'

    const workbook = WorkPaper.buildFromSheets(
      {
        Sheet1: rows,
        Data: data,
      },
      { maxRows: 100, maxColumns: 30, useColumnIndex: true },
    )

    expectNumber(cellValue(workbook, 'Sheet1', 0, 0), 20)
    expectNumber(cellValue(workbook, 'Sheet1', 0, 1), 20)
    expectNumber(cellValue(workbook, 'Sheet1', 0, 2), 100)
    expectNumber(cellValue(workbook, 'Sheet1', 0, 3), 40)
    expectNumber(cellValue(workbook, 'Sheet1', 0, 4), 5)
    expectNumber(cellValue(workbook, 'Sheet1', 0, 5), 5)
    expectNumber(cellValue(workbook, 'Sheet1', 0, 6), 20)
    expectNumber(cellValue(workbook, 'Sheet1', 1, 0), 20)
    expectNumber(cellValue(workbook, 'Sheet1', 1, 1), 20)
    expectNumber(cellValue(workbook, 'Sheet1', 1, 2), 100)
    expectNumber(cellValue(workbook, 'Sheet1', 1, 3), 40)
    expectNumber(cellValue(workbook, 'Sheet1', 1, 4), 5)
    expectNumber(cellValue(workbook, 'Sheet1', 1, 5), 5)
    expectNumber(cellValue(workbook, 'Sheet1', 1, 6), 20)
  })

  it('resolves issue #103 worksheet-reference OFFSET ranges', () => {
    const rows = Array.from({ length: 20 }, () => Array.from<TestCell>({ length: 12 }).fill(null))

    rows[1][1] = 100
    rows[1][2] = 200
    rows[5][1] = 155
    rows[5][2] = 1030
    rows[0][4] = '=SUM(OFFSET(B2:C2,4,0))'
    rows[1][4] = '=OFFSET(B2:C2,4,0)'

    rows[2][10] = 1
    rows[10][3] = 2
    rows[2][3] = '=CORREL(OFFSET($C$15:$C$18,0,$K3),OFFSET($C$15:$C$18,0,D$11))'
    rows[2][5] = '=SUM(OFFSET($C$15:$C$18,0,$K3))'
    ;[1, 2, 3, 4].forEach((value, index) => {
      rows[index + 14][3] = value
      rows[index + 14][4] = value * 2
    })

    const workbook = WorkPaper.buildFromSheets(
      {
        Sheet1: rows,
      },
      { maxRows: 40, maxColumns: 12, useColumnIndex: true },
    )

    expectNumber(cellValue(workbook, 'Sheet1', 0, 4), 1185)
    expectNumber(cellValue(workbook, 'Sheet1', 1, 4), 155)
    expectNumberClose(cellValue(workbook, 'Sheet1', 2, 3), 1)
    expectNumber(cellValue(workbook, 'Sheet1', 2, 5), 10)

    const sheetId = workbook.getSheetId('Sheet1')!
    workbook.setCellContents({ sheet: sheetId, row: 14, col: 3 }, 10)

    expectNumber(cellValue(workbook, 'Sheet1', 2, 5), 19)
  })

  it('resolves issue #116 advanced XLOOKUP modes and spill returns', () => {
    const rows = Array.from({ length: 20 }, () => Array.from<TestCell>({ length: 22 }).fill(null))

    rows[0][0] = '=XLOOKUP(72,G1:G5,H1:H5,,-1)'
    ;[
      [50, 'D'],
      [60, 'C'],
      [70, 'B'],
      [80, 'A'],
      [90, 'S'],
    ].forEach(([score, grade], index) => {
      rows[index][6] = score
      rows[index][7] = grade
    })

    rows[2][0] = '=XLOOKUP("ID2",O1:O3,P1:R3)'
    rows[0][14] = 'ID1'
    rows[1][14] = 'ID2'
    rows[2][14] = 'ID3'
    rows[0][15] = 'Alex'
    rows[0][16] = 'North'
    rows[0][17] = 10
    rows[1][15] = 'James'
    rows[1][16] = 'South'
    rows[1][17] = 20
    rows[2][15] = 'Mina'
    rows[2][16] = 'West'
    rows[2][17] = 30

    rows[4][0] = '=XLOOKUP(T1:T3,J1:M1,J2:M2)'
    rows[0][19] = 'Q2'
    rows[1][19] = 'Q4'
    rows[2][19] = 'Q1'
    rows[0][9] = 'Q1'
    rows[0][10] = 'Q2'
    rows[0][11] = 'Q3'
    rows[0][12] = 'Q4'
    rows[1][9] = 'Keyboard'
    rows[1][10] = 'Printer'
    rows[1][11] = 'Monitor'
    rows[1][12] = 'Dock'

    const workbook = WorkPaper.buildFromSheets(
      {
        Sheet1: rows,
      },
      { maxRows: 40, maxColumns: 24, useColumnIndex: true },
    )

    expectString(cellValue(workbook, 'Sheet1', 0, 0), 'B')
    expectString(cellValue(workbook, 'Sheet1', 2, 0), 'James')
    expectString(cellValue(workbook, 'Sheet1', 2, 1), 'South')
    expectNumber(cellValue(workbook, 'Sheet1', 2, 2), 20)
    expectString(cellValue(workbook, 'Sheet1', 4, 0), 'Printer')
    expectString(cellValue(workbook, 'Sheet1', 5, 0), 'Dock')
    expectString(cellValue(workbook, 'Sheet1', 6, 0), 'Keyboard')
  })

  it('resolves issue #102 formula number text coercion during concatenation', () => {
    const rows = Array.from({ length: 45 }, () => Array.from<TestCell>({ length: 10 }).fill(null))

    rows[5][4] = 1989
    rows[5][5] = 1
    rows[5][6] = '=E6&"|"&IF(F6<10,"0"&F6,F6)'
    rows[15][2] = '$'
    rows[15][3] = 1_000_000
    rows[27][4] = 1989
    rows[27][5] = 2
    rows[27][6] = '=E28&"|"&IF(F28<10,"0"&F28,F28)'
    rows[44][3] = '=IF(D16>=1000000,$C16&ROUNDUP(D16/1000000,2)&"m",$C16&ROUNDUP(D16/1000,0)&"k")&" at "'
    rows[44][4] = '=ROUND(0.5*100,0)&"%"'
    rows[44][5] = '=(0.1+0.2)&" cash"'

    const workbook = WorkPaper.buildFromSheets(
      {
        'Step 1': rows,
      },
      { maxRows: 80, maxColumns: 12, useColumnIndex: true },
    )

    expectString(cellValue(workbook, 'Step 1', 5, 6), '1989|01')
    expectString(cellValue(workbook, 'Step 1', 27, 6), '1989|02')
    expectString(cellValue(workbook, 'Step 1', 44, 3), '$1m at ')
    expectString(cellValue(workbook, 'Step 1', 44, 4), '50%')
    expectString(cellValue(workbook, 'Step 1', 44, 5), '0.3 cash')
  })

  it('resolves issue #117 imported hidden-row SUBTOTAL values', () => {
    const tableRows: Array<{ readonly amount: number; readonly quarter: string }> = [
      { amount: 3255, quarter: 'Qtr 2' },
      { amount: 4865, quarter: 'Qtr 4' },
      { amount: 9339, quarter: 'Qtr 2' },
      { amount: 14808, quarter: 'Qtr 4' },
      { amount: 1390, quarter: 'Qtr 3' },
      { amount: 7433, quarter: 'Qtr 1' },
      { amount: 9213, quarter: 'Qtr 4' },
      { amount: 9698, quarter: 'Qtr 1' },
      { amount: 16753, quarter: 'Qtr 3' },
      { amount: 18919, quarter: 'Qtr 3' },
      { amount: 10644, quarter: 'Qtr 2' },
      { amount: 12438, quarter: 'Qtr 1' },
      { amount: 14867, quarter: 'Qtr 3' },
      { amount: 19302, quarter: 'Qtr 4' },
    ]
    const hiddenRows = [3, 6, 9, 11, 12, 14]
    const dataCells = tableRows.flatMap((row, index) => {
      const excelRow = index + 2
      return [
        { address: `B${excelRow}`, value: row.amount },
        { address: `D${excelRow}`, value: row.quarter },
      ]
    })
    const snapshot: WorkbookSnapshot = {
      version: 1,
      workbook: { name: 'issue-117-subtotal-filtered-table' },
      sheets: [
        {
          id: 1,
          name: 'Table',
          order: 0,
          metadata: {
            rows: hiddenRows.map((index) => ({ id: `row:${index}`, index, hidden: true })),
          },
          cells: [...dataCells, { address: 'B16', formula: 'SUBTOTAL(109,B2:B15)' }, { address: 'D16', formula: 'SUBTOTAL(103,D2:D15)' }],
        },
      ],
    }

    const workbook = WorkPaper.buildFromSnapshot(snapshot, { maxRows: 50, maxColumns: 8, useColumnIndex: true })

    expectNumber(cellValue(workbook, 'Table', 15, 1), 77_015)
    expectNumber(cellValue(workbook, 'Table', 15, 3), 8)
  })
})
