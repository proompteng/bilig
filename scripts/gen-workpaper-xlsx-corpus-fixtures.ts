#!/usr/bin/env bun

import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { join, resolve } from 'node:path'
import * as XLSX from 'xlsx'

const fixtureDirectory = resolve('packages/headless/fixtures/xlsx-corpus')

interface XlsxCorpusFixture {
  readonly fileName: string
  readonly workbook: XLSX.WorkBook
}

function buildFixtures(): readonly XlsxCorpusFixture[] {
  return [
    {
      fileName: 'issue-8-production-regressions.xlsx',
      workbook: buildIssue8ProductionRegressionWorkbook(),
    },
  ]
}

function buildIssue8ProductionRegressionWorkbook(): XLSX.WorkBook {
  const workbook = XLSX.utils.book_new()

  const summary = XLSX.utils.aoa_to_sheet([
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
    [],
    ['Bank lookup', null, 'txn-123', null],
  ])
  summary.B2 = {
    t: 'n',
    f: 'SUMIFS(Activity!$B$2:$B$4,Activity!$A$2:$A$4,"Deposit")',
    v: 3500,
  }
  summary.B3 = { t: 's', f: 'IF(ABS(B2-3500)<0.01,"PASS","FAIL")', v: 'PASS' }
  summary.C3 = { t: 'e', f: '1/0', v: 7, w: '#DIV/0!' }
  summary.B4 = { t: 'n', f: 'COUNTA(Activity!$A$2:$A$4)', v: 3 }
  summary.B5 = { t: 's', f: 'HYPERLINK("#\'Summary\'!A1","Go to Summary")', v: 'Go to Summary' }
  summary.B6 = { t: 's', f: 'TEXT(46127,"mm.dd.yy")', v: '04.15.26' }
  summary.B7 = { t: 'n', f: 'DAY(46127)', v: 15 }
  summary.B8 = { t: 'n', f: 'WORKDAY(46127,2)', v: 46_129 }
  summary.B9 = { t: 'n', f: 'COUNTIF(Activity!$A$2:$A$4,"Deposit")', v: 1 }
  summary.B10 = {
    t: 'n',
    f: 'SUMIF(Activity!$A$2:$A$4,"<>Fee",Activity!$B$2:$B$4)',
    v: 3250,
  }
  summary.B11 = {
    t: 'n',
    f: 'IFERROR(ROUND(SUMIFS(Activity!$B$2:$B$4,Activity!$A$2:$A$4,"Deposit"),2),0)',
    v: 3500,
  }
  summary.B12 = {
    t: 's',
    f: 'IFERROR(XLOOKUP(C14,Bank!$D$2:$D$31,Bank!$B$2:$B$31,"",0),"")',
    v: '2026-04-01',
  }
  summary.D14 = {
    t: 's',
    f: 'IFERROR(INDEX(Bank!$B$2:$B$31,MATCH(C14,Bank!$D$2:$D$31,0)),"")',
    v: '2026-04-01',
  }
  summary['!ref'] = 'A1:D14'

  const activity = XLSX.utils.aoa_to_sheet([
    ['Type', 'Amount'],
    ['Deposit', 3500],
    ['Fee', -18.5],
    ['Withdrawal', -250],
  ])

  const bank = XLSX.utils.aoa_to_sheet([
    ['Date label', 'Date', 'Description', 'Transaction ID'],
    ['Posted', '2026-04-01', 'Deposit', 'txn-123'],
  ])
  bank['!ref'] = 'A1:D31'

  XLSX.utils.book_append_sheet(workbook, summary, 'Summary')
  XLSX.utils.book_append_sheet(workbook, activity, 'Activity')
  XLSX.utils.book_append_sheet(workbook, bank, 'Bank')
  return workbook
}

function fixtureBytes(workbook: XLSX.WorkBook): Buffer {
  const encoded: unknown = XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' })
  if (Buffer.isBuffer(encoded)) {
    return encoded
  }
  if (encoded instanceof Uint8Array) {
    return Buffer.from(encoded)
  }
  throw new Error('Expected XLSX writer to return a buffer')
}

function run(check: boolean): void {
  if (!check) {
    mkdirSync(fixtureDirectory, { recursive: true })
  }

  for (const fixture of buildFixtures()) {
    const path = join(fixtureDirectory, fixture.fileName)
    const expected = fixtureBytes(fixture.workbook)
    if (check) {
      if (!existsSync(path)) {
        throw new Error(`Missing XLSX corpus fixture: ${path}`)
      }
      const actual = readFileSync(path)
      if (!actual.equals(expected)) {
        throw new Error(`XLSX corpus fixture is stale: ${path}`)
      }
      continue
    }
    writeFileSync(path, expected)
  }
}

run(process.argv.includes('--check'))
