import { spawnSync } from 'node:child_process'
import { mkdtempSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'
import * as XLSX from 'xlsx'

import { runWorkPaperXlsxCorpus, runWorkPaperXlsxCorpusInChildProcesses } from '../check-workpaper-xlsx-corpus.ts'

describe('WorkPaper XLSX corpus verifier', () => {
  it('keeps a checked-in issue #8 XLSX compatibility corpus green', () => {
    const result = runWorkPaperXlsxCorpus([checkedInCorpusDir()])

    expect(result.summary).toMatchObject({
      totalFiles: 1,
      filesProcessed: 1,
      ok: 1,
      failedErrors: 0,
      failedTimeouts: 0,
      formulaCells: 14,
      comparableFormulaCells: 14,
      matchingFormulaCells: 14,
      mismatchedFormulaCells: 0,
      skippedFormulaCells: 0,
      matchRate: 1,
    })
    expect(result.files[0]).toMatchObject({
      fileName: 'issue-8-production-regressions.xlsx',
      status: 'ok',
      formulaCells: 14,
    })
    expect(result.mismatches).toEqual([])
  })

  it('can isolate each workbook check in a child process with the same parity result', () => {
    const direct = runWorkPaperXlsxCorpus([checkedInCorpusDir()])
    const isolated = runWorkPaperXlsxCorpusInChildProcesses([checkedInCorpusDir()], {
      childProcessTimeoutMs: 10_000,
    })

    expect(isolated.summary).toMatchObject({
      totalFiles: direct.summary.totalFiles,
      filesProcessed: direct.summary.filesProcessed,
      ok: direct.summary.ok,
      failedErrors: 0,
      failedTimeouts: 0,
      formulaCells: direct.summary.formulaCells,
      comparableFormulaCells: direct.summary.comparableFormulaCells,
      matchingFormulaCells: direct.summary.matchingFormulaCells,
      mismatchedFormulaCells: 0,
      skippedFormulaCells: direct.summary.skippedFormulaCells,
      matchRate: 1,
    })
    expect(isolated.mismatches).toEqual([])
  })

  it('matches cached formula results from a production-style XLSX reduction corpus', () => {
    withTempCorpus((corpusDir) => {
      writeWorkbook(join(corpusDir, 'issue-regressions.xlsx'), buildIssueRegressionWorkbook())

      const result = runWorkPaperXlsxCorpus([corpusDir])

      expect(result.summary).toMatchObject({
        totalFiles: 1,
        filesProcessed: 1,
        ok: 1,
        failedTimeouts: 0,
        formulaCells: 9,
        comparableFormulaCells: 9,
        matchingFormulaCells: 9,
        mismatchedFormulaCells: 0,
        skippedFormulaCells: 0,
        matchRate: 1,
      })
      expect(result.files[0]?.status).toBe('ok')
      expect(result.mismatches).toEqual([])
    })
  })

  it('fails oversized workbook files before loading them', () => {
    withTempCorpus((corpusDir) => {
      const workbookPath = join(corpusDir, 'oversized.xlsx')
      writeWorkbook(workbookPath, buildIssueRegressionWorkbook())

      const result = runWorkPaperXlsxCorpus([workbookPath], { maxFileBytes: 1 })

      expect(result.summary).toMatchObject({
        totalFiles: 1,
        filesProcessed: 1,
        ok: 0,
        failedErrors: 1,
        failedTimeouts: 0,
      })
      expect(result.files[0]).toMatchObject({
        fileName: 'oversized.xlsx',
        status: 'error',
        formulaCells: 0,
      })
      expect(result.files[0]?.error).toContain('XLSX file exceeds max file size')
    })
  })

  it('refuses unisolated CLI corpus runs unless explicitly enabled for debugging', () => {
    const env = { ...process.env }
    delete env.BILIG_ALLOW_UNISOLATED_XLSX_CORPUS

    const result = spawnSync('bun', [checkerScriptPath(), '--no-isolate', checkedInCorpusDir()], {
      encoding: 'utf8',
      env,
    })

    expect(result.status).toBe(2)
    expect(result.stderr).toContain('--no-isolate is disabled for corpus CLI runs')
  })

  it('refuses unisolated CLI directory sweeps even with the debug escape hatch', () => {
    const result = spawnSync('bun', [checkerScriptPath(), '--no-isolate', checkedInCorpusDir()], {
      encoding: 'utf8',
      env: { ...process.env, BILIG_ALLOW_UNISOLATED_XLSX_CORPUS: '1' },
    })

    expect(result.status).toBe(2)
    expect(result.stderr).toContain('--no-isolate only supports one explicit XLSX file')
  })

  it('refuses broad CLI directory sweeps while the corpus stop marker is active', () => {
    withTempCorpus((tempDir) => {
      const stopMarkerPath = join(tempDir, 'stop.md')
      writeFileSync(stopMarkerPath, 'stop')

      const result = spawnSync('bun', [checkerScriptPath(), '--corpus-run-stop-marker', stopMarkerPath, checkedInCorpusDir()], {
        encoding: 'utf8',
      })

      expect(result.status).toBe(2)
      expect(result.stderr).toContain('workpaper:xlsx-corpus directory sweep is disabled while the public corpus stop marker is active')
      expect(result.stderr).toContain('--allow-active-stop-marker')
    })
  })

  it('allows a single-file CLI debugger check while the corpus stop marker is active', () => {
    withTempCorpus((tempDir) => {
      const stopMarkerPath = join(tempDir, 'stop.md')
      writeFileSync(stopMarkerPath, 'stop')

      const result = spawnSync('bun', [checkerScriptPath(), '--corpus-run-stop-marker', stopMarkerPath, checkedInCorpusFile()], {
        encoding: 'utf8',
      })

      expect(result.status).toBe(0)
      expect(JSON.parse(result.stdout)).toMatchObject({
        summary: {
          totalFiles: 1,
          filesProcessed: 1,
          ok: 1,
        },
      })
    })
  })

  it('reports actionable mismatch samples with workbook, sheet, address, formula, expected, and actual values', () => {
    withTempCorpus((corpusDir) => {
      const workbook = XLSX.utils.book_new()
      const sheet = XLSX.utils.aoa_to_sheet([[1, null]])
      sheet.B1 = { t: 'n', f: 'A1+1', v: 99 }
      sheet['!ref'] = 'A1:B1'
      XLSX.utils.book_append_sheet(workbook, sheet, 'Sheet1')
      writeWorkbook(join(corpusDir, 'mismatch.xlsx'), workbook)

      const result = runWorkPaperXlsxCorpus([corpusDir])

      expect(result.summary).toMatchObject({
        totalFiles: 1,
        filesProcessed: 1,
        ok: 0,
        formulaCells: 1,
        comparableFormulaCells: 1,
        matchingFormulaCells: 0,
        mismatchedFormulaCells: 1,
        matchRate: 0,
      })
      expect(result.files[0]).toMatchObject({
        fileName: 'mismatch.xlsx',
        status: 'mismatched',
        formulaCells: 1,
        mismatchedFormulaCells: 1,
      })
      expect(result.mismatches[0]).toMatchObject({
        fileName: 'mismatch.xlsx',
        sheetName: 'Sheet1',
        address: 'B1',
        formula: 'A1+1',
        expected: { kind: 'number', value: 99 },
        actual: { kind: 'number', value: 2 },
      })
    })
  })

  it('counts cached-less and volatile formulas as skipped instead of comparable parity failures', () => {
    withTempCorpus((corpusDir) => {
      const workbook = XLSX.utils.book_new()
      const sheet = XLSX.utils.aoa_to_sheet([[null, null]])
      sheet.A1 = { t: 'n', f: 'NOW()', v: 46_127 }
      sheet.B1 = { t: 'n', f: 'A1+1' }
      sheet['!ref'] = 'A1:B1'
      XLSX.utils.book_append_sheet(workbook, sheet, 'Sheet1')
      writeWorkbook(join(corpusDir, 'skipped.xlsx'), workbook)

      const result = runWorkPaperXlsxCorpus([corpusDir])

      expect(result.summary).toMatchObject({
        totalFiles: 1,
        filesProcessed: 1,
        ok: 1,
        formulaCells: 2,
        comparableFormulaCells: 0,
        matchingFormulaCells: 0,
        mismatchedFormulaCells: 0,
        skippedFormulaCells: 2,
        matchRate: 1,
      })
      expect(result.files[0]?.skippedFormulaCells).toBe(2)
      expect(result.skippedByReason).toEqual({
        'missing-cached-result': 1,
        'unsupported-cached-result-type': 0,
        'volatile-or-environment-dependent-formula': 1,
      })
    })
  })
})

function checkedInCorpusDir(): string {
  return join(dirname(fileURLToPath(import.meta.url)), '../../packages/headless/fixtures/xlsx-corpus')
}

function checkedInCorpusFile(): string {
  return join(checkedInCorpusDir(), 'issue-8-production-regressions.xlsx')
}

function checkerScriptPath(): string {
  return join(dirname(fileURLToPath(import.meta.url)), '../check-workpaper-xlsx-corpus.ts')
}

function withTempCorpus(run: (corpusDir: string) => void): void {
  const corpusDir = mkdtempSync(join(tmpdir(), 'bilig-workpaper-xlsx-corpus-'))
  try {
    run(corpusDir)
  } finally {
    rmSync(corpusDir, { recursive: true, force: true })
  }
}

function writeWorkbook(path: string, workbook: XLSX.WorkBook): void {
  writeFileSync(path, XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' }))
}

function buildIssueRegressionWorkbook(): XLSX.WorkBook {
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
    [],
    [],
    [],
    [],
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
