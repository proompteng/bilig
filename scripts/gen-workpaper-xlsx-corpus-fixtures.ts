#!/usr/bin/env bun

import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { join, resolve } from 'node:path'
import { pathToFileURL } from 'node:url'
import {
  decodeCellAddress,
  decodeCellRange,
  encodeCellAddress,
  writeSimpleXlsxWorkbook,
  type SimpleXlsxCell,
  type SimpleXlsxSheet,
  type SimpleXlsxWorkbook,
} from '@bilig/xlsx'
import { bytesEqual, xlsxZipEntryContentsEqual } from './xlsx-fixture-comparison.ts'

const fixtureDirectory = resolve(process.env.BILIG_XLSX_CORPUS_FIXTURE_DIR ?? 'packages/workpaper/fixtures/xlsx-corpus')

interface XlsxCorpusFixture {
  readonly fileName: string
  readonly sourcePath?: string
  readonly workbook?: SimpleXlsxWorkbook
}

function buildFixtures(): readonly XlsxCorpusFixture[] {
  return [
    {
      fileName: 'issue-8-production-regressions.xlsx',
      workbook: buildIssue8ProductionRegressionWorkbook(),
    },
    {
      fileName: 'macos-excel-threaded-comments-source.xlsx',
      sourcePath: resolve('packages/workpaper/fixtures/excel-oracle/macos-excel-threaded-comments-source.xlsx'),
    },
  ]
}

type FixtureCellValue = boolean | number | string | null
type FixtureCellPatch = Omit<SimpleXlsxCell, 'address' | 'row' | 'col'>

function literalCells(rows: readonly (readonly FixtureCellValue[])[]): readonly SimpleXlsxCell[] {
  const cells: SimpleXlsxCell[] = []
  for (let rowIndex = 0; rowIndex < rows.length; rowIndex += 1) {
    const row = rows[rowIndex] ?? []
    for (let colIndex = 0; colIndex < row.length; colIndex += 1) {
      const value = row[colIndex]
      if (value === null) {
        continue
      }
      const address = encodeCellAddress({ r: rowIndex, c: colIndex })
      cells.push({ address, row: rowIndex, col: colIndex, value })
    }
  }
  return cells
}

function formulaCell(address: string, patch: FixtureCellPatch): SimpleXlsxCell {
  const decoded = decodeCellAddress(address)
  return {
    address: encodeCellAddress(decoded),
    row: decoded.r,
    col: decoded.c,
    ...patch,
  }
}

function fixtureSheet(
  name: string,
  rows: readonly (readonly FixtureCellValue[])[],
  formulaCells: readonly SimpleXlsxCell[] = [],
  dimension?: string,
): SimpleXlsxSheet {
  return {
    name,
    cells: [...literalCells(rows), ...formulaCells],
    ...(dimension ? { dimension: decodeCellRange(dimension) } : {}),
  }
}

function buildIssue8ProductionRegressionWorkbook(): SimpleXlsxWorkbook {
  const summary = fixtureSheet(
    'Summary',
    [
      ['Metric', 'Value', 'Lookup key', 'Lookup value'],
      ['Deposits', null, null, null],
      ['Deposit check', null, null, null],
      ['Activity rows', null, null, null],
      ['Internal link', null, null, null],
      ['Formatted date', null, null, null],
      ['Day', null, null, null],
      ['Workday', null, null, null],
      ['Deposit count', null, null, null],
      ['Non-fee sum', null, null, null],
      ['Wrapped deposits', null, null, null],
      ['XLOOKUP bank date', null, null, null],
      ['Average ignores blanks', null, null, null],
      ['Bank lookup', null, 'txn-123', null],
    ],
    [
      formulaCell('B2', { formula: 'SUMIFS(Activity!$B$2:$B$4,Activity!$A$2:$A$4,"Deposit")', value: 3500 }),
      formulaCell('B3', { formula: 'IF(ABS(B2-3500)<0.01,"PASS","FAIL")', value: 'PASS' }),
      formulaCell('C3', { formula: '1/0', error: '#DIV/0!' }),
      formulaCell('B4', { formula: 'COUNTA(Activity!$A$2:$A$4)', value: 3 }),
      formulaCell('B5', { formula: 'HYPERLINK("#\'Summary\'!A1","Go to Summary")', value: 'Go to Summary' }),
      formulaCell('B6', { formula: 'TEXT(46127,"mm.dd.yy")', value: '04.15.26' }),
      formulaCell('B7', { formula: 'DAY(46127)', value: 15 }),
      formulaCell('B8', { formula: 'WORKDAY(46127,2)', value: 46_129 }),
      formulaCell('B9', { formula: 'COUNTIF(Activity!$A$2:$A$4,"Deposit")', value: 1 }),
      formulaCell('B10', { formula: 'SUMIF(Activity!$A$2:$A$4,"<>Fee",Activity!$B$2:$B$4)', value: 3250 }),
      formulaCell('B11', {
        formula: 'IFERROR(ROUND(SUMIFS(Activity!$B$2:$B$4,Activity!$A$2:$A$4,"Deposit"),2),0)',
        value: 3500,
      }),
      formulaCell('B12', { formula: 'IFERROR(XLOOKUP(C14,Bank!$D$2:$D$31,Bank!$B$2:$B$31,"",0),"")', value: '2026-04-01' }),
      formulaCell('B13', { formula: 'ROUND(AVERAGE(AverageInputs!$A$2:$A$20),2)', value: 18.91 }),
      formulaCell('D14', { formula: 'IFERROR(INDEX(Bank!$B$2:$B$31,MATCH(C14,Bank!$D$2:$D$31,0)),"")', value: '2026-04-01' }),
    ],
    'A1:D14',
  )

  return {
    sheets: [
      summary,
      fixtureSheet('Activity', [
        ['Type', 'Amount'],
        ['Deposit', 3500],
        ['Fee', -18.5],
        ['Withdrawal', -250],
      ]),
      fixtureSheet(
        'Bank',
        [
          ['Date label', 'Date', 'Description', 'Transaction ID'],
          ['Posted', '2026-04-01', 'Deposit', 'txn-123'],
        ],
        [],
        'A1:D31',
      ),
      fixtureSheet('AverageInputs', [['Value'], [12.5], [24], [18.75], [20.25], [19.04], ['Department'], ['']], [], 'A1:A20'),
    ],
  }
}

function fixtureBytes(workbook: SimpleXlsxWorkbook): Buffer {
  return Buffer.from(writeSimpleXlsxWorkbook(workbook))
}

function fixtureMatches(fixture: XlsxCorpusFixture, actual: Buffer, expected: Buffer): boolean {
  return fixture.workbook ? xlsxZipEntryContentsEqual(actual, expected) : bytesEqual(actual, expected)
}

function run(check: boolean): void {
  if (!check) {
    mkdirSync(fixtureDirectory, { recursive: true })
  }

  for (const fixture of buildFixtures()) {
    const path = join(fixtureDirectory, fixture.fileName)
    const expected = fixture.workbook ? fixtureBytes(fixture.workbook) : readFileSync(fixture.sourcePath!)
    if (check) {
      if (!existsSync(path)) {
        throw new Error(`Missing XLSX corpus fixture: ${path}`)
      }
      const actual = readFileSync(path)
      if (!fixtureMatches(fixture, actual, expected)) {
        throw new Error(`XLSX corpus fixture is stale: ${path}`)
      }
      continue
    }
    writeFileSync(path, expected)
  }
}

function isCliEntrypoint(): boolean {
  const entrypoint = process.argv[1]
  return entrypoint !== undefined && import.meta.url === pathToFileURL(resolve(entrypoint)).href
}

if (isCliEntrypoint()) {
  run(process.argv.includes('--check'))
}
